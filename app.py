# -*- coding: utf-8 -*-
import io
import os
import re
import tempfile
import subprocess
from datetime import datetime, date
from decimal import Decimal
from zipfile import ZipFile, ZIP_DEFLATED

import streamlit as st
from openpyxl import load_workbook
from docx import Document
from docx.table import _Cell
from docx.text.paragraph import Paragraph

# 선택: docx2pdf가 있으면 활용
try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

# =========================
#           CONST
# =========================
TOKEN_RE = re.compile(r"\{\{([A-Z]+[0-9]+)(?:\|([^}]+))?\}\}")  # {{A1}} or {{A1|FORMAT}}
LEFTOVER_RE = re.compile(r"\{\{[^}]+\}\}")
DEFAULT_OUT = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행.docx"
TARGET_SHEET = "2.  배정후 청약시"

# =========================
#         UTILITIES
# =========================
def ensure_docx(name: str) -> str:
    name = (name or "").strip()
    return name if name.lower().endswith(".docx") else (name + ".docx")

def ensure_pdf(name: str) -> str:
    base = (name or "output").strip()
    if base.lower().endswith(".docx"):
        base = base[:-5]
    return base + ".pdf"

def has_soffice() -> bool:
    return any(
        os.path.isfile(os.path.join(p, "soffice")) or os.path.isfile(os.path.join(p, "soffice.bin"))
        for p in os.environ.get("PATH", "").split(os.pathsep)
    )

def try_format_as_date(v) -> str:
    try:
        if v is None:
            return ""
        if isinstance(v, (datetime, date)):
            return f"{v.year}. {v.month}. {v.day}."
        s = str(v).strip()
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
            dt = datetime.strptime(s, "%Y-%m-%d").date()
            return f"{dt.year}. {dt.month}. {dt.day}."
    except Exception:
        pass
    return ""

def fmt_number(v) -> str:
    try:
        if isinstance(v, (int, float, Decimal)):
            return f"{float(v):,.0f}"
        if isinstance(v, str):
            raw = v.replace(",", "")
            if re.fullmatch(r"-?\d+(\.\d+)?", raw):
                return f"{float(raw):,.0f}"
    except Exception:
        pass
    return ""

def value_to_text(v) -> str:
    s = try_format_as_date(v)
    if s:
        return s
    s = fmt_number(v)
    if s:
        return s
    return "" if v is None else str(v)

# =========================
#     INLINE FORMATTER
# =========================
def apply_inline_format(value, fmt: str | None) -> str:
    """
    {{A1|#,###}}, {{B7|YYYY.MM.DD}} 형태의 포맷을 지원.
    - 날짜: YYYY/MM/DD → %Y/%m/%d 매핑
    - 숫자: '#,###' / '#,###.00' 등 소수 자릿수 인식
    """
    if fmt is None or fmt.strip() == "":
        return value_to_text(value)

    # 날짜 포맷
    if any(tok in fmt for tok in ("YYYY", "MM", "DD")):
        if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value.strip()):
            value = datetime.strptime(value.strip(), "%Y-%m-%d").date()
        if isinstance(value, (datetime, date)):
            f = fmt.replace("YYYY", "%Y").replace("MM", "%m").replace("DD", "%d")
            return value.strftime(f)
        return value_to_text(value)

    # 숫자 포맷
    if re.fullmatch(r"[#,0]+(?:\.[0#]+)?", fmt.replace(",", "")):
        try:
            num = float(str(value).replace(",", ""))
            decimals = 0
            if "." in fmt:
                decimals = len(fmt.split(".")[1])
            return f"{num:,.{decimals}f}"
        except Exception:
            return value_to_text(value)

    return value_to_text(value)

# =========================
#      DOCX TRAVERSAL
# =========================
def iter_block_items(parent):
    """문서의 문단/표 셀 모두 순회 (본문, 헤더/푸터 공통 사용). duck-typing으로 안전 처리."""
    if hasattr(parent, "paragraphs") and hasattr(parent, "tables"):
        for p in parent.paragraphs:
            yield p
        for t in parent.tables:
            for row in t.rows:
                for cell in row.cells:
                    for item in iter_block_items(cell):
                        yield item
    elif isinstance(parent, _Cell):
        for p in parent.paragraphs:
            yield p
        for t in parent.tables:
            for row in t.rows:
                for cell in row.cells:
                    for item in iter_block_items(cell):
                        yield item

def replace_in_paragraph(par: Paragraph, repl_func):
    changed = False
    for run in par.runs:
        new_text = repl_func(run.text)
        if new_text != run.text:
            run.text = new_text
            changed = True
    if changed:
        return
    full_text = "".join(r.text for r in par.runs)
    new_text = repl_func(full_text)
    if new_text == full_text:
        return
    if par.runs:
        par.runs[0].text = new_text
        for r in par.runs[1:]:
            r.text = ""

def replace_everywhere(doc: Document, repl_func):
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            replace_in_paragraph(item, repl_func)
    for section in doc.sections:
        for container in (section.header, section.footer):
            for item in iter_block_items(container):
                if isinstance(item, Paragraph):
                    replace_in_paragraph(item, repl_func)

# =========================
#    EXCEL → REPLACER
# =========================
def make_replacer(ws):
    def _repl(text: str) -> str:
        def sub(m):
            addr, fmt = m.group(1), m.group(2)
            try:
                v = ws[addr].value
            except Exception:
                v = None
            return apply_inline_format(v, fmt)

        replaced = TOKEN_RE.sub(sub, text)

        # YYYY/MM/DD 같은 더미 템플릿 치환(간단)
        sp = "    "
        today = datetime.today()
        today_str = f"{today.year}년{sp}{today.month}월{sp}{today.day}일"
        for token in ["YYYY년 MM월 DD일", "YYYY년    MM월    DD일", "YYYY 년 MM 월 DD 일"]:
            replaced = replaced.replace(token, today_str)
        return replaced
    return _repl

# =========================
#       DOCX → PDF
# =========================
def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> bytes | None:
    try:
        with tempfile.TemporaryDirectory() as td:
            in_path = os.path.join(td, "doc.docx")
            out_path = os.path.join(td, "doc.pdf")
            with open(in_path, "wb") as f:
                f.write(docx_bytes)

            # Word (Windows) 경로
            if docx2pdf_convert is not None:
                try:
                    docx2pdf_convert(in_path, out_path)
                    if os.path.exists(out_path):
                        with open(out_path, "rb") as f:
                            return f.read()
                except Exception:
                    pass

            # LibreOffice headless
            if has_soffice():
                try:
                    subprocess.run(
                        ["soffice", "--headless", "--convert-to", "pdf", in_path, "--outdir", td],
                        check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
                    )
                    if os.path.exists(out_path):
                        with open(out_path, "rb") as f:
                            return f.read()
                except Exception:
                    pass
    except Exception:
        pass
    return None

# =========================
#   LEFTOVER TOKEN SCAN
# =========================
def collect_leftover_tokens(doc: Document) -> set[str]:
    leftovers = set()
    def _scan(container):
        for item in iter_block_items(container):
            if isinstance(item, Paragraph):
                text = "".join(r.text for r in item.runs) if item.runs else item.text
                for m in LEFTOVER_RE.findall(text or ""):
                    leftovers.add(m)

    _scan(doc)
    for section in doc.sections:
        for container in (section.header, section.footer):
            _scan(container)
    return leftovers

# =========================
#           UI
# =========================
st.set_page_config(page_title="Document Generator", page_icon="🧩", layout="wide")

# --- Custom CSS (dark, neo-style) ---
st.markdown("""
<style>
:root{
  --bg:#0f172a;         /* slate-900 */
  --panel:#0b1224;      /* deep navy */
  --card:#111a2e;       /* card */
  --muted:#8ea1c0;
  --acc-1:#22d3ee;      /* cyan */
  --acc-2:#60a5fa;      /* blue */
  --ok:#34d399;         /* green */
  --warn:#fbbf24;       /* amber */
  --err:#f87171;        /* red */
  --radius:18px;
}

html, body, [data-testid="stAppViewContainer"]{
  background: radial-gradient(1200px 600px at 20% -10%, rgba(96,165,250,.12), transparent 40%),
              radial-gradient(1000px 600px at 80% 0%, rgba(34,211,238,.08), transparent 35%),
              var(--bg);
}

.sidebar .sidebar-content, [data-testid="stSidebar"]{
  background: linear-gradient(180deg, #0b1224, #0b1224);
  border-right: 1px solid rgba(255,255,255,.06);
}

.block-container{
  padding-top: 2rem;
}

h1,h2,h3,h4,h5,p,span,div,label{
  color: #e2e8f0;
}

/* Header */
.app-head{
  display:flex; align-items:center; justify-content:space-between;
  padding: 12px 18px; border-radius: var(--radius);
  background: linear-gradient(180deg, rgba(255,255,255,.03), rgba(255,255,255,.01));
  border: 1px solid rgba(255,255,255,.06);
  box-shadow: 0 10px 30px rgba(0,0,0,.25);
  margin-bottom: 18px;
}
.app-title{font-size: 26px; font-weight: 800; letter-spacing:.2px}
.app-sub{color: var(--muted); font-size: 13px}

/* Grid cards */
.card{
  background: var(--card);
  border: 1px solid rgba(255,255,255,.06);
  border-radius: var(--radius);
  padding: 22px;
  box-shadow: 0 14px 35px rgba(0,0,0,.35), inset 0 1px 0 rgba(255,255,255,.04);
}
.card h3{ margin: 0 0 8px 0; font-weight: 700; }

/* Upload card icons */
.icon{
  width:40px;height:40px;border-radius:10px; display:flex;align-items:center;justify-content:center;
  background: rgba(96,165,250,.15); color:#93c5fd; font-size:20px; margin-right:10px;
  border:1px solid rgba(96,165,250,.25);
}

/* Status pills */
.pill{display:inline-flex; align-items:center; gap:6px;
  padding:6px 10px; border-radius:999px; font-size:12px; border:1px solid rgba(255,255,255,.08);
  background: rgba(255,255,255,.04); color:#cbd5e1; margin-right:8px;
}
.pill.ok{ border-color: rgba(52,211,153,.25); color:#a7f3d0; }
.pill.warn{ border-color: rgba(251,191,36,.25); color:#fde68a; }
.pill.err{ border-color: rgba(248,113,113,.25); color:#fecaca; }

/* Progress bar wrapper to look sleeker */
.progress-wrap{
   background: rgba(255,255,255,.06); border-radius: 999px; padding: 6px;
   border: 1px solid rgba(255,255,255,.08);
}
.small{ color: var(--muted); font-size: 12px; }

/* Download button */
button[kind="secondary"]{
  border-radius: 12px !important;
}

/* Hide default top padding around widgets inside cards */
.card [data-testid="stMarkdownContainer"] > p { margin: 0; }

</style>
""", unsafe_allow_html=True)

# --- Sidebar (simple nav look) ---
with st.sidebar:
    st.markdown("### 🧭 Navigation")
    st.markdown("- **Dashboard**")
    st.markdown("- Templates")
    st.markdown("- Documents")
    st.markdown("- Settings")
    st.divider()
    st.markdown("**Quick Guides**")
    st.caption("• 템플릿에 {{A1|#,###}} 처럼 포맷을 적으면 그대로 적용됩니다.")
    st.caption("• 남은 {{...}} 토큰은 자동 스캔되어 리포트됩니다.")

# --- Header ---
st.markdown(
    '<div class="app-head">'
    '<div><div class="app-title">DOCUMENT GENERATOR</div>'
    '<div class="app-sub">Automate your documents</div></div>'
    '<div class="pill ok">● Ready</div>'
    '</div>',
    unsafe_allow_html=True
)

# =========================
#        MAIN GRID
# =========================
left, right = st.columns([1.15, 0.85], gap="large")

with left:
    # Upload Cards
    c1, c2 = st.columns(2, gap="large")

    with c1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown(
            '<div style="display:flex;align-items:center;margin-bottom:10px">'
            '<div class="icon">X</div><div><h3>UPLOAD EXCEL TEMPLATE</h3>'
            '<div class="small">엑셀 파일(.xlsx, .xlsm)을 업로드하세요</div></div></div>',
            unsafe_allow_html=True,
        )
        xlsx_file = st.file_uploader("", type=["xlsx", "xlsm"], key="excel")
        st.markdown('</div>', unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown(
            '<div style="display:flex;align-items:center;margin-bottom:10px">'
            '<div class="icon">W</div><div><h3>UPLOAD WORD TEMPLATE</h3>'
            '<div class="small">워드 템플릿(.docx)을 업로드하세요</div></div></div>',
            unsafe_allow_html=True,
        )
        docx_tpl = st.file_uploader("", type=["docx"], key="word")
        st.markdown('</div>', unsafe_allow_html=True)

    # Options + Action
    st.markdown('<div class="card">', unsafe_allow_html=True)
    out_name = st.text_input("출력 파일명", value=DEFAULT_OUT)

    sheet_choice = None
    if xlsx_file:
        try:
            wb_tmp = load_workbook(filename=io.BytesIO(xlsx_file.getvalue()), data_only=True)
            sheet_choice = st.selectbox(
                "시트 선택",
                wb_tmp.sheetnames,
                index=wb_tmp.sheetnames.index(TARGET_SHEET) if TARGET_SHEET in wb_tmp.sheetnames else 0
            )
        except Exception as e:
            st.warning(f"시트 미리보기 실패: {e}")

    col_run = st.columns([1, 1, 3])
    with col_run[0]:
        run = st.button("문서 생성", type="primary", use_container_width=True)
    with col_run[1]:
        clear = st.button("초기화", use_container_width=True)

    if clear:
        st.session_state.pop("recent_jobs", None)
        st.experimental_rerun()

    # Progress placeholder
    prog_box = st.empty()
    msg_box = st.empty()

    # Output placeholders
    out_box = st.empty()
    leftover_box = st.empty()
    st.markdown('</div>', unsafe_allow_html=True)

with right:
    # Status / Capability card
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("#### Status")
    pills = []
    if docx2pdf_convert is not None:
        pills.append('<span class="pill ok">MS Word (docx2pdf)</span>')
    if has_soffice():
        pills.append('<span class="pill ok">LibreOffice (soffice)</span>')
    if not pills:
        pills.append('<span class="pill warn">PDF 변환기 없음 · ZIP에 DOCX만 포함</span>')
    st.markdown(" ".join(pills), unsafe_allow_html=True)
    st.markdown('<div class="small">두 경로 모두 가능하면 Word 우선 → 실패 시 LibreOffice 시도</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Recent jobs
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("#### Recent Generations")
    recent = st.session_state.get("recent_jobs", [])
    if recent:
        for r in recent[-5:][::-1]:
            st.markdown(
                f"- **{r['when']}** · {r['name']} · "
                + (":green[PDF]" if r.get("pdf") else ":orange[DOCX]") 
            )
    else:
        st.caption("생성 기록이 없습니다.")
    st.markdown('</div>', unsafe_allow_html=True)

# =========================
#        GENERATE
# =========================
if run:
    if not xlsx_file or not docx_tpl:
        st.error("엑셀 파일과 워드 템플릿을 모두 업로드하세요.")
        st.stop()

    try:
        # Progress feel
        prog_box.markdown('<div class="progress-wrap">', unsafe_allow_html=True)
        prog = st.progress(0, text="엑셀 로드 중...")

        # Excel
        wb = load_workbook(filename=io.BytesIO(xlsx_file.read()), data_only=True)
        prog.progress(10, text="시트 선택...")
        ws = wb[sheet_choice] if sheet_choice else (wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb[wb.sheetnames[0]])

        # Word template
        prog.progress(25, text="워드 템플릿 로드...")
        tpl_bytes = docx_tpl.read()
        doc = Document(io.BytesIO(tpl_bytes))

        # Replace
        prog.progress(55, text="치환 처리 중...")
        replacer = make_replacer(ws)
        replace_everywhere(doc, replacer)

        # Save DOCX
        prog.progress(70, text="DOCX 저장...")
        docx_buf = io.BytesIO()
        doc.save(docx_buf)
        docx_buf.seek(0)
        docx_bytes = docx_buf.getvalue()

        # Try PDF
        prog.progress(82, text="PDF 변환 시도...")
        pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)

        # Zip
        prog.progress(90, text="ZIP 패키징...")
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            zf.writestr(ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT, docx_bytes)
            if pdf_bytes:
                zf.writestr(ensure_pdf(out_name), pdf_bytes)
        zip_buf.seek(0)

        # Leftover tokens
        prog.progress(96, text="남은 토큰 검사...")
        doc_after = Document(io.BytesIO(docx_bytes))
        leftovers = sorted(list(collect_leftover_tokens(doc_after)))

        # Outputs
        prog.progress(100, text="완료!")
        msg_box.success("완료되었습니다.")

        # Leftovers report (collapsible)
        if leftovers:
            with leftover_box.expander("템플릿에 남은 치환 토큰(참고용)"):
                st.write(", ".join(leftovers))
        else:
            leftover_box.empty()

        # Downloads
        with out_box:
            cdl, cdoc, cpdf = st.columns([2,1,1])
            with cdl:
                st.download_button(
                    "WORD + PDF 한번에 다운로드 (ZIP)",
                    data=zip_buf,
                    file_name=(ensure_pdf(out_name).replace(".pdf", "") + "_both.zip"),
                    mime="application/zip",
                    use_container_width=True,
                )
            with cdoc:
                st.download_button(
                    "DOCX만",
                    data=docx_bytes,
                    file_name=ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            with cpdf:
                if pdf_bytes:
                    st.download_button(
                        "PDF만",
                        data=pdf_bytes,
                        file_name=ensure_pdf(out_name),
                        mime="application/pdf",
                        use_container_width=True,
                    )
                else:
                    st.button("PDF 불가", disabled=True, use_container_width=True)

        # Recent list
        rec = st.session_state.get("recent_jobs", [])
        rec.append({"when": datetime.now().strftime("%H:%M:%S"), "name": ensure_docx(out_name), "pdf": bool(pdf_bytes)})
        st.session_state["recent_jobs"] = rec

    except Exception as e:
        msg_box.error("생성 중 오류가 발생했습니다.")
        st.exception(e)
