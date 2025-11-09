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

# app.py

import streamlit as st
from openpyxl import load_workbook
from docx import Document
# ... 기존 import 그대로 ...

st.set_page_config(page_title="납입요청서 자동 생성", layout="wide")  # ← 있으면 유지, 없으면 넣어줘

# ✅ 여기 아래에 바로 CSS + 업로드 카드 UI 넣으면 됨
st.markdown("""
<style>
.upload-card {
  background: rgba(2,6,23,.65);
  border: 1px solid rgba(148,163,184,.25);
  border-radius: 16px;
  padding: 24px;
  box-shadow: 0 8px 30px rgba(0,0,0,.25);
}
.upload-card [data-testid="stFileUploaderDropzone"]{
  background: rgba(17,24,39,.55);
  border: 1px solid rgba(148,163,184,.25);
  border-radius: 12px;
}
.upload-card [data-testid="stFileUploader"] section { gap: 6px; }
.upload-card [data-testid="stFileUploader"] button { border-radius: 10px; }
.upload-title{ font-weight: 800; font-size: 20px; color: #e5e7eb; margin-bottom: 6px; }
.upload-sub{ color:#94a3b8; font-size:13px; margin-bottom: 14px; }
</style>
""", unsafe_allow_html=True)

st.title("DOCUMENT GENERATOR")
st.write("Automate Your Documents")

col1, col2 = st.columns(2, gap="large")

with col1:
    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.markdown('<div class="upload-title">UPLOAD EXCEL TEMPLATE</div>', unsafe_allow_html=True)
    st.markdown('<div class="upload-sub">엑셀 템플릿(.xlsx / .xlsm)</div>', unsafe_allow_html=True)
    excel_file = st.file_uploader("", type=["xlsx","xlsm"], label_visibility="collapsed", key="excel")
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.markdown('<div class="upload-title">UPLOAD WORD TEMPLATE</div>', unsafe_allow_html=True)
    st.markdown('<div class="upload-sub">워드 템플릿(.docx)</div>', unsafe_allow_html=True)
    docx_file = st.file_uploader("", type=["docx"], label_visibility="collapsed", key="docx")
    st.markdown('</div>', unsafe_allow_html=True)

# ✅ 그리고 나서 “출력 파일명 + 문서 생성” 버튼 + 로직 이어 붙이면 됨


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
