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

# 선택: docx2pdf가 있으면 활용(Windows+Word 환경에서만 동작)
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
    """
    {{A1|#,###}}, {{B7|YYYY.MM.DD}} 지원
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

    # 숫자 포맷(간이)
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

# ----------------- 문서 순회/치환 -----------------
def iter_block_items(parent):
    """문서의 문단/표 셀 모두 순회 (본문, 헤더/푸터 공통 사용)."""
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

        # YYYY/MM/DD 같은 더미 템플릿 치환(간단)
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

            # 1) Word(docx2pdf)
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
st.set_page_config(page_title="Document Generator", page_icon="🧩", layout="wide")

st.markdown("""
<style>
/* 배경 그라데이션 */
.stApp {
  background: radial-gradient(1200px 600px at 10% 0%, rgba(34,211,238,.06), rgba(0,0,0,0)) ,
              radial-gradient(1200px 600px at 90% 20%, rgba(59,130,246,.06), rgba(0,0,0,0)) ,
              #0b1220;
}

/* 타이틀 */
.h1-title {
  font-size: 36px; font-weight: 800; color: #fff; letter-spacing: .02em;
  text-shadow: 0 1px 0 rgba(255,255,255,.05);
}
.h1-sub { color:#94a3b8; margin-top:6px }

/* 카드 */
.dg-card {
  position: relative; border: 1px solid rgba(148,163,184,.25);
  background: rgba(15,23,42,.55); border-radius: 16px; padding: 24px;
  backdrop-filter: blur(6px);
  transition: border-color .2s ease, box-shadow .2s ease, transform .2s ease;
}
.dg-card:hover { border-color: rgba(34,211,238,.45); box-shadow: 0 8px 30px rgba(34,211,238,.08); }

/* 카드 아이콘 원형 */
.icon-bubble {
  width: 80px; height: 80px; border-radius: 16px;
  display: flex; align-items: center; justify-content: center;
  background: rgba(148,163,184,.15);
  margin: 8px auto 18px auto;
}

/* 버튼 */
.dg-btn-primary {
  background: linear-gradient(90deg, #06b6d4, #3b82f6);
  color:#fff; border:0; padding: 12px 18px; border-radius: 12px;
  font-weight: 700; letter-spacing:.02em;
}
.dg-btn-primary:hover { filter: brightness(1.06); box-shadow: 0 6px 22px rgba(56,189,248,.35); }

.dg-btn-outline {
  background: transparent; color:#22d3ee;
  border: 2px solid rgba(34,211,238,.6);
  padding: 10px 16px; border-radius: 10px; font-weight:600;
}
.dg-btn-outline:hover{ background: rgba(34,211,238,.08); }

/* 입력창 */
.dg-input input {
  background: rgba(2,6,23,.5); border:1px solid rgba(148,163,184,.35);
  color:#e5e7eb; border-radius: 12px; padding: 12px 14px;
}
.dg-input input:focus{ border-color:#22d3ee; box-shadow:none; }

/* 진행바 */
.progress-wrap{ background: rgba(100,116,139,.25); height:10px; border-radius: 999px; overflow:hidden;}
.progress-bar{ height:100%; background: linear-gradient(90deg,#06b6d4,#3b82f6); }
.badge{ font-size:12px; color:#a3aed0; }
</style>
""", unsafe_allow_html=True)

# ----------------- Streamlit UI -----------------
st.markdown('<div class="h1-title">DOCUMENT GENERATOR</div>', unsafe_allow_html=True)
st.markdown('<div class="h1-sub">Automate Your Documents</div>', unsafe_allow_html=True)
st.write("")

left, right = st.columns(2, gap="large")

with left:
    st.markdown('<div class="dg-card">', unsafe_allow_html=True)
    st.markdown('<div class="icon-bubble">📊</div>', unsafe_allow_html=True)
    st.subheader("UPLOAD EXCEL TEMPLATE")
    st.caption("엑셀 템플릿(.xlsx / .xlsm)")
    xlsx_file = st.file_uploader("Drag&Drop or Browse", type=["xlsx", "xlsm"], label_visibility="collapsed")
    st.markdown('<div style="text-align:center;">', unsafe_allow_html=True)
    st.button("Browse Files", key="btn_xlsx_dummy", help="위 업로더와 동일")
    st.markdown('</div></div>', unsafe_allow_html=True)

with right:
    st.markdown('<div class="dg-card">', unsafe_allow_html=True)
    st.markdown('<div class="icon-bubble">📝</div>', unsafe_allow_html=True)
    st.subheader("UPLOAD WORD TEMPLATE")
    st.caption("워드 템플릿(.docx)")
    docx_tpl = st.file_uploader("Drag&Drop or Browse ", type=["docx"], label_visibility="collapsed")
    st.markdown('<div style="text-align:center;">', unsafe_allow_html=True)
    st.button("Browse Files ", key="btn_docx_dummy", help="위 업로더와 동일")
    st.markdown('</div></div>', unsafe_allow_html=True)

st.write("")
st.markdown('<div class="dg-card">', unsafe_allow_html=True)
out_name = st.text_input("출력 파일명", value=DEFAULT_OUT, key="outname", label_visibility="visible")
st.markdown('</div>', unsafe_allow_html=True)

# 시트 선택(간단 드롭다운)
sheet_choice = None
if xlsx_file:
    wb_tmp = load_workbook(filename=io.BytesIO(xlsx_file.getvalue()), data_only=True)
    sheet_choice = st.selectbox("Excel 시트 선택", wb_tmp.sheetnames,
                                index=wb_tmp.sheetnames.index(TARGET_SHEET) if TARGET_SHEET in wb_tmp.sheetnames else 0)

gen_col, _ = st.columns([1,3])
with gen_col:
    run = st.button("문서 생성하기", type="primary", use_container_width=True)

# ----------------- 실행 -----------------
if run:
    if not xlsx_file or not docx_tpl:
        st.error("엑셀 파일과 워드 템플릿을 모두 업로드하세요.")
        st.stop()

    # 진행바 느낌(UX)
    prog = st.empty()
    with st.spinner("Generating Documents…"):
        prog.markdown('<div class="dg-card"><div class="badge">Generating… 0%</div>'
                      '<div class="progress-wrap"><div class="progress-bar" style="width:0%"></div></div></div>',
                      unsafe_allow_html=True)

        # Excel 로드
        wb = load_workbook(filename=io.BytesIO(xlsx_file.read()), data_only=True)
        ws = wb[sheet_choice] if sheet_choice else (wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb[wb.sheetnames[0]])
        prog.markdown('<div class="dg-card"><div class="badge">Generating… 25%</div>'
                      '<div class="progress-wrap"><div class="progress-bar" style="width:25%"></div></div></div>',
                      unsafe_allow_html=True)

        # Word 템플릿 로드
        tpl_bytes = docx_tpl.read()
        doc = Document(io.BytesIO(tpl_bytes))

        # 치환
        replacer = make_replacer(ws)
        replace_everywhere(doc, replacer)
        prog.markdown('<div class="dg-card"><div class="badge">Generating… 60%</div>'
                      '<div class="progress-wrap"><div class="progress-bar" style="width:60%"></div></div></div>',
                      unsafe_allow_html=True)

        # DOCX 메모리 저장
        docx_buf = io.BytesIO()
        doc.save(docx_buf)
        docx_buf.seek(0)
        docx_bytes = docx_buf.getvalue()

        # PDF 변환 시도
        pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)
        prog.markdown('<div class="dg-card"><div class="badge">Generating… 85%</div>'
                      '<div class="progress-wrap"><div class="progress-bar" style="width:85%"></div></div></div>',
                      unsafe_allow_html=True)

        # ZIP 묶기
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            zf.writestr(ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT, docx_bytes)
            if pdf_bytes:
                zf.writestr(ensure_pdf(out_name), pdf_bytes)
        zip_buf.seek(0)

        prog.markdown('<div class="dg-card"><div class="badge">Completed 100%</div>'
                      '<div class="progress-wrap"><div class="progress-bar" style="width:100%"></div></div></div>',
                      unsafe_allow_html=True)

    # 누락 토큰 리포트(정보용)
    doc_after = Document(io.BytesIO(docx_bytes))
    leftovers = sorted(list(collect_leftover_tokens(doc_after)))
    if leftovers:
        with st.expander("템플릿에 남은 치환 토큰(참고용)"):
            st.write(", ".join(leftovers))

    st.success("완료되었습니다.")
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "WORD+PDF 한번에 다운로드 (ZIP)",
            data=zip_buf,
            file_name=(ensure_pdf(out_name).replace(".pdf", "") + "_both.zip"),
            mime="application/zip",
            use_container_width=True
        )
    with c2:
        st.download_button(
            "DOCX만 다운로드",
            data=docx_bytes,
            file_name=ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    if pdf_bytes:
        st.download_button(
            "PDF만 다운로드",
            data=pdf_bytes,
            file_name=ensure_pdf(out_name),
            mime="application/pdf",
            use_container_width=True
        )
