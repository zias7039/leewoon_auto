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

# ----------------- 상수 -----------------
TOKEN_RE = re.compile(r"\{\{([A-Z]+[0-9]+)(?:\|([^}]+))?\}\}")  # {{A1}} or {{A1|FORMAT}}
LEFTOVER_RE = re.compile(r"\{\{[^}]+\}\}")
DEFAULT_OUT = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행.docx"
TARGET_SHEET = "2.  배정후 청약시"

# ----------------- 유틸 -----------------
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

# ----------------- 포맷 적용 -----------------
def apply_inline_format(value, fmt: str | None) -> str:
    # {{A1|#,###}}, {{B7|YYYY.MM.DD}}
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

    # 숫자 포맷(간이)
    if re.fullmatch(r"[#,0]+(?:\.[0#]+)?", fmt.replace(",", "")):
        try:
            num = float(str(value).replace(",", ""))
            decimals = len(fmt.split(".")[1]) if "." in fmt else 0
            return f"{num:,.{decimals}f}"
        except Exception:
            return value_to_text(value)

    return value_to_text(value)

# ----------------- 문서 순회/치환 -----------------
def iter_block_items(parent):
    # 문단/표 셀 모두 순회 (duck-typing)
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

# ----------------- Excel → 치환 콜백 -----------------
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

        # 더미 'YYYY년 MM월 DD일' 치환
        sp = "    "
        today = datetime.today()
        today_str = f"{today.year}년{sp}{today.month}월{sp}{today.day}일"
        for token in ["YYYY년 MM월 DD일", "YYYY년    MM월    DD일", "YYYY 년 MM 월 DD 일"]:
            replaced = replaced.replace(token, today_str)
        return replaced
    return _repl

# ----------------- DOCX → PDF -----------------
def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> bytes | None:
    try:
        with tempfile.TemporaryDirectory() as td:
            in_path = os.path.join(td, "doc.docx")
            out_path = os.path.join(td, "doc.pdf")
            with open(in_path, "wb") as f:
                f.write(docx_bytes)

            # 1) MS Word 경로
            if docx2pdf_convert is not None:
                try:
                    docx2pdf_convert(in_path, out_path)
                    if os.path.exists(out_path):
                        with open(out_path, "rb") as f:
                            return f.read()
                except Exception:
                    pass

            # 2) LibreOffice headless
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

# ----------------- 누락 토큰 수집 -----------------
def collect_leftover_tokens(doc: Document) -> set[str]:
    leftovers = set()
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            text = "".join(r.text for r in item.runs) if item.runs else item.text
            for m in LEFTOVER_RE.findall(text or ""):
                leftovers.add(m)
    for section in doc.sections:
        for container in (section.header, section.footer):
            for item in iter_block_items(container):
                if isinstance(item, Paragraph):
                    text = "".join(r.text for r in item.runs) if item.runs else item.text
                    for m in LEFTOVER_RE.findall(text or ""):
                        leftovers.add(m)
    return leftovers

# ----------------- UI 스타일 -----------------
def inject_css():
    st.markdown(
        """
<style>
/* 전체 배경과 기본 글꼴 */
html, body, [data-testid="stAppViewContainer"] {
  background: radial-gradient(1200px 700px at 20% -10%, #111827 10%, #0b1220 60%, #0b0f19 100%) fixed;
  color: #e5e7eb;
  font-family: ui-sans-serif, system-ui, -apple-system, "Segoe UI", Roboto, "Apple SD Gothic Neo", "Noto Sans KR", "Malgun Gothic", "Apple Color Emoji", "Segoe UI Emoji";
}

/* 사이드바 */
[data-testid="stSidebar"] {
  background: rgba(15,23,42,.65);
  backdrop-filter: blur(10px);
  border-right: 1px solid rgba(148,163,184,.25);
}
.sidebar-chip{
  display:flex;align-items:center;gap:.6rem;
  padding:.6rem .9rem;border-radius:.7rem;
  color:#94a3b8;border-left:4px solid transparent;
}
.sidebar-chip.active{background:rgba(6,182,212,.12); color:#67e8f9; border-left-color:#06b6d4;}
.sidebar-chip:hover{background:rgba(51,65,85,.35); color:#cbd5e1}

/* 헤더 */
.header{
  background: rgba(15,23,42,.4);
  border-bottom:1px solid rgba(148,163,184,.25);
  backdrop-filter: blur(10px);
  padding: 14px 24px;
}

/* 카드 */
.card{
  position:relative;
  background: rgba(30,41,59,.55);
  border:1px solid rgba(148,163,184,.25);
  border-radius: 16px;
  padding: 24px;
  transition: border-color .25s ease, transform .25s ease, box-shadow .25s ease;
}
.card:hover{
  border-color: rgba(34,211,238,.45);
  box-shadow: 0 10px 40px rgba(34,211,238,.15);
  transform: translateY(-2px);
}

/* 버튼 */
.btn{
  display:inline-flex; align-items:center; justify-content:center;
  padding: .8rem 1.2rem; gap:.5rem;
  font-weight:700; border-radius: 12px;
  border: 2px solid rgba(6,182,212,.9);
  color:#67e8f9; background: transparent;
}
.btn:hover{ background: rgba(6,182,212,.12); }

/* 큰 버튼 */
.btn-primary{
  width:100%; padding: 1.1rem 1.4rem; font-size:1.05rem;
  border: none; color:white;
  background: linear-gradient(90deg, #06b6d4, #3b82f6);
  box-shadow: 0 10px 40px rgba(34,211,238,.25);
}
.btn-primary:disabled{ background:#374151; color:#6b7280; box-shadow:none; }

/* 진행바 박스 */
.progress{
  background: rgba(31,41,55,.7);
  border:1px solid rgba(148,163,184,.25);
  border-radius: 12px; padding: 12px 14px;
}
.progress-track{ width:100%; height:10px; background:#374151; border-radius: 999px; overflow:hidden; }
.progress-bar{ height:100%; background: linear-gradient(90deg,#06b6d4,#3b82f6); transition: width .3s ease; }

/* 상태칩 */
.kb{ display:flex; align-items:center; gap:.5rem; color:#9ca3af; }
.badge{ display:inline-flex; align-items:center; gap:.35rem; padding:.25rem .5rem; border-radius:999px; font-size:.8rem; }
.badge.ok{ background:rgba(34,197,94,.15); color:#4ade80; }
.badge.wait{ background:rgba(234,179,8,.15); color:#facc15; }
.badge.err{ background:rgba(239,68,68,.15); color:#f87171; }

/* 업로더 오버레이 클릭 영역 감추기 */
.block-container { padding-top: 0rem; }
</style>
        """,
        unsafe_allow_html=True,
    )

def sidebar(active="documents"):
    st.sidebar.markdown("### ")
    def nav_chip(text, id_):
        cls = "sidebar-chip active" if id_ == active else "sidebar-chip"
        st.sidebar.markdown(f'<div class="{cls}">• {text}</div>', unsafe_allow_html=True)
    nav_chip("Dashboard", "dashboard")
    nav_chip("Templates", "templates")
    nav_chip("Documents", "documents")
    st.sidebar.markdown("---")
    nav_chip("Settings", "settings")
    nav_chip("Help", "help")

def header():
    st.markdown(
        """
<div class="header" style="display:flex;justify-content:flex-end;">
  <div style="display:flex;align-items:center;gap:.6rem;padding:.35rem .8rem;border:1px solid rgba(148,163,184,.25);border-radius:999px;background:rgba(31,41,55,.6);">
    <div style="width:10px;height:10px;border-radius:999px;background:#22c55e;"></div>
    <span style="color:#cbd5e1;">Jin-Young</span>
  </div>
</div>
        """,
        unsafe_allow_html=True,
    )

def upload_card(title, hint, accept, key):
    c = st.container()
    with c:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        file = st.file_uploader("", type=accept, key=key, label_visibility="collapsed")
        icon = "✅" if file else "☁️"
        st.markdown(
            f"""
<div style="display:flex;flex-direction:column;align-items:center;gap:12px;text-align:center;">
  <div style="width:84px;height:84px;border-radius:16px;display:flex;align-items:center;justify-content:center;background:{'rgba(34,197,94,.18)' if file else 'rgba(55,65,81,.4)'};font-size:38px;">{icon}</div>
  <div>
    <div style="font-weight:800;font-size:20px;color:white;letter-spacing:.3px;">{title}</div>
    <div style="color:#94a3b8;font-size:13px;margin-top:4px;">{file.name if file else hint}</div>
  </div>
  <div><span class="btn">Browse Files</span></div>
</div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown('</div>', unsafe_allow_html=True)
    return file

# ----------------- APP -----------------
def main():
    st.set_page_config(page_title="Document Generator", page_icon="📄", layout="wide")
    inject_css()
    sidebar("documents")
    header()

    st.markdown("<div style='padding:28px;'></div>", unsafe_allow_html=True)
    st.markdown(
        "<div style='max-width:1080px;margin:0 auto;'>"
        "<div style='margin-bottom:28px;'>"
        "<div style='font-size:34px;font-weight:900;color:white;margin-bottom:6px;'>DOCUMENT GENERATOR</div>"
        "<div style='color:#94a3b8;font-size:16px;'>Automate Your Documents</div>"
        "</div>",
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2, gap="large")
    with col1:
        xlsx_file = upload_card(
            "UPLOAD EXCEL TEMPLATE",
            "엑셀 템플릿(.xlsx / .xlsm)",
            ["xlsx", "xlsm"],
            "excel",
        )
    with col2:
        docx_tpl = upload_card(
            "UPLOAD WORD TEMPLATE",
            "워드 템플릿(.docx)",
            ["docx"],
            "docx",
        )

    # 파일명
    st.markdown('<div class="card" style="margin-top:16px;">', unsafe_allow_html=True)
    out_name = st.text_input("출력 파일명", value=DEFAULT_OUT, label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

    # 시트 선택(있으면)
    sheet_choice = None
    if xlsx_file:
        wb_tmp = load_workbook(filename=io.BytesIO(xlsx_file.getvalue()), data_only=True)
        sheet_choice = st.selectbox(
            "Excel 시트",
            wb_tmp.sheetnames,
            index=wb_tmp.sheetnames.index(TARGET_SHEET) if TARGET_SHEET in wb_tmp.sheetnames else 0,
            label_visibility="collapsed",
        )

    # 생성 버튼
    disabled = not (xlsx_file and docx_tpl)
    gen = st.button(
        "문서 생성하기",
        type="primary",
        disabled=disabled,
        use_container_width=True,
        help=None,
    )
    st.markdown(
        f'<style>.stButton button{{}} .stButton button {{}} .stButton button {{}} </style>',
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<div><button class="btn-primary" {"disabled" if disabled else ""} style="display:none;"></button></div>',
        unsafe_allow_html=True,
    )

    # 진행 + 처리
    if gen:
        prog_box = st.container()
        with prog_box:
            st.markdown('<div class="progress">', unsafe_allow_html=True)
            top = st.empty()
            bar = st.empty()
            st.markdown('</div>', unsafe_allow_html=True)

        def draw_progress(pct):
            top.markdown(
                f'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">'
                f'<span style="color:#e5e7eb;">Generating Documents...</span>'
                f'<span style="color:#67e8f9;font-weight:800;">{pct}%</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
            bar.markdown(
                f'<div class="progress-track"><div class="progress-bar" style="width:{pct}%"></div></div>',
                unsafe_allow_html=True,
            )

        draw_progress(12)

        try:
            # Excel 로드
            wb = load_workbook(filename=io.BytesIO(xlsx_file.read()), data_only=True)
            ws = wb[sheet_choice] if sheet_choice else (
                wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb[wb.sheetnames[0]]
            )

            # Word 템플릿 로드
            tpl_bytes = docx_tpl.read()
            doc = Document(io.BytesIO(tpl_bytes))

            # 치환
            replacer = make_replacer(ws)
            replace_everywhere(doc, replacer)
            draw_progress(48)

            # DOCX 저장
            docx_buf = io.BytesIO()
            doc.save(docx_buf)
            docx_buf.seek(0)
            docx_bytes = docx_buf.getvalue()

            # PDF 변환
            pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)
            draw_progress(86)

            # ZIP
            zip_buf = io.BytesIO()
            with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
                zf.writestr(ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT, docx_bytes)
                if pdf_bytes:
                    zf.writestr(ensure_pdf(out_name), pdf_bytes)
            zip_buf.seek(0)

            # 누락 토큰
            doc_after = Document(io.BytesIO(docx_bytes))
            leftovers = sorted(list(collect_leftover_tokens(doc_after)))

            draw_progress(100)

            # 결과 블록
            st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
            st.markdown(
                '<div class="card" style="display:flex;flex-direction:column;gap:12px;">',
                unsafe_allow_html=True,
            )
            st.markdown(
                '<div class="kb">'
                '<span class="badge ok">● COMPLETED</span>'
                f'<span style="margin-left:.5rem;color:#9ca3af;">{ensure_docx(out_name)}</span>'
                '</div>',
                unsafe_allow_html=True,
            )

            c1, c2, c3 = st.columns([1,1,1], gap="large")
            with c1:
                st.download_button(
                    "⬇️ Download DOCX",
                    data=docx_bytes,
                    file_name=ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            with c2:
                st.download_button(
                    "⬇️ Download PDF",
                    data=pdf_bytes if pdf_bytes else b"",
                    file_name=ensure_pdf(out_name),
                    mime="application/pdf",
                    use_container_width=True,
                    disabled=pdf_bytes is None,
                    help=None if pdf_bytes else "PDF 변환 환경(Word/LibreOffice)이 없어 비활성화됨",
                )
            with c3:
                st.download_button(
                    "⬇️ Download ZIP (Both)",
                    data=zip_buf,
                    file_name=(ensure_pdf(out_name).replace(".pdf", "") + "_both.zip"),
                    mime="application/zip",
                    use_container_width=True,
                )

            if leftovers:
                with st.expander("템플릿에 남은 치환 토큰"):
                    st.write(", ".join(leftovers))

            st.markdown('</div>', unsafe_allow_html=True)

            # Recent-like 리스트(예시)
            st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
            st.markdown(
                '<div style="display:flex;justify-content:space-between;align-items:center;">'
                '<div style="font-weight:800;color:white;">RECENT GENERATIONS</div>'
                '<div class="kb" style="gap:16px;">'
                '<span class="badge ok">✔ COMPLETED</span>'
                '<span class="badge wait">⏳ PENDING</span>'
                '<span class="badge err">✖ ERROR</span>'
                '</div></div>',
                unsafe_allow_html=True,
            )
            for txt, badge in [
                (f"{datetime.today():%Y-%m-%d}_납입요청서_완료.docx", "ok"),
                (f"{datetime.today():%Y-%m-%d}_납입요청서_대기.docx", "wait"),
                (f"{datetime.today():%Y-%m-%d}_납입요청서_오류.docx", "err"),
            ]:
                st.markdown(
                    f'<div class="card" style="padding:14px;display:flex;align-items:center;justify-content:space-between;">'
                    f'<div style="display:flex;align-items:center;gap:.6rem;">'
                    f'<span class="badge {badge}">●</span><span style="color:#cbd5e1;">{txt}</span></div>'
                    f'<span style="color:#94a3b8;font-size:13px;">just now</span></div>',
                    unsafe_allow_html=True,
                )

        except Exception as e:
            draw_progress(100)
            st.error("문서 생성 중 오류가 발생했습니다.")
            st.exception(e)

    # 푸터 여백
    st.markdown("</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
