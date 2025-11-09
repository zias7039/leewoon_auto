# -*- coding: utf-8 -*-

import io, os, re, tempfile, subprocess
from datetime import datetime, date
from decimal import Decimal
from zipfile import ZipFile, ZIP_DEFLATED

import streamlit as st
from openpyxl import load_workbook
from docx import Document
from docx.table import _Cell
from docx.text.paragraph import Paragraph

# 스타일 모듈
from ui_style import inject_style

# 선택: docx2pdf
try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

TOKEN_RE = re.compile(r"\{\{([A-Z]+[0-9]+)(?:\|([^}]+))?\}\}")
LEFTOVER_RE = re.compile(r"\{\{[^}]+\}\}")
DEFAULT_OUT = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행.docx"
TARGET_SHEET = "2.  배정후 청약시"

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
    if s: return s
    s = fmt_number(v)
    if s: return s
    return "" if v is None else str(v)

def apply_inline_format(value, fmt: str | None) -> str:
    if fmt is None or fmt.strip() == "":
        return value_to_text(value)
    if any(tok in fmt for tok in ("YYYY", "MM", "DD")):
        if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value.strip()):
            value = datetime.strptime(value.strip(), "%Y-%m-%d").date()
        if isinstance(value, (datetime, date)):
            f = fmt.replace("YYYY", "%Y").replace("MM","%m").replace("DD","%d")
            return value.strftime(f)
        return value_to_text(value)
    if re.fullmatch(r"[#,0]+(?:\.[0#]+)?", fmt.replace(",", "")):
        try:
            num = float(str(value).replace(",", ""))
            decimals = len(fmt.split(".")[1]) if "." in fmt else 0
            return f"{num:,.{decimals}f}"
        except Exception:
            return value_to_text(value)
    return value_to_text(value)

def iter_block_items(parent):
    if hasattr(parent, "paragraphs") and hasattr(parent, "tables"):
        for p in parent.paragraphs: yield p
        for t in parent.tables:
            for row in t.rows:
                for cell in row.cells:
                    for item in iter_block_items(cell): yield item
    elif isinstance(parent, _Cell):
        for p in parent.paragraphs: yield p
        for t in parent.tables:
            for row in t.rows:
                for cell in row.cells:
                    for item in iter_block_items(cell): yield item

def replace_in_paragraph(par: Paragraph, repl_func):
    changed = False
    for run in par.runs:
        new_text = repl_func(run.text)
        if new_text != run.text:
            run.text = new_text
            changed = True
    if changed: return
    full_text = "".join(r.text for r in par.runs)
    new_text = repl_func(full_text)
    if new_text == full_text: return
    if par.runs:
        par.runs[0].text = new_text
        for r in par.runs[1:]: r.text = ""

def replace_everywhere(doc: Document, repl_func):
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            replace_in_paragraph(item, repl_func)
    for section in doc.sections:
        for container in (section.header, section.footer):
            for item in iter_block_items(container):
                if isinstance(item, Paragraph):
                    replace_in_paragraph(item, repl_func)

def make_replacer(ws):
    def _repl(text: str) -> str:
        def sub(m):
            addr, fmt = m.group(1), m.group(2)
            try: v = ws[addr].value
            except Exception: v = None
            return apply_inline_format(v, fmt)
        replaced = TOKEN_RE.sub(sub, text)
        sp = "    "
        today = datetime.today()
        today_str = f"{today.year}년{sp}{today.month}월{sp}{today.day}일"
        for token in ["YYYY년 MM월 DD일", "YYYY년    MM월    DD일", "YYYY 년 MM 월 DD 일"]:
            replaced = replaced.replace(token, today_str)
        return replaced
    return _repl

def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> bytes | None:
    try:
        with tempfile.TemporaryDirectory() as td:
            in_path = os.path.join(td, "doc.docx")
            out_path = os.path.join(td, "doc.pdf")
            with open(in_path, "wb") as f: f.write(docx_bytes)
            if docx2pdf_convert is not None:
                try:
                    docx2pdf_convert(in_path, out_path)
                    if os.path.exists(out_path):
                        with open(out_path, "rb") as f: return f.read()
                except Exception: pass
            if has_soffice():
                try:
                    subprocess.run(
                        ["soffice", "--headless", "--convert-to", "pdf", in_path, "--outdir", td],
                        check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
                    )
                    if os.path.exists(out_path):
                        with open(out_path, "rb") as f: return f.read()
                except Exception: pass
    except Exception:
        pass
    return None

def collect_leftover_tokens(doc: Document) -> set[str]:
    leftovers = set()
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            text = "".join(r.text for r in item.runs) if item.runs else item.text
            for m in LEFTOVER_RE.findall(text or ""): leftovers.add(m)
    for section in doc.sections:
        for container in (section.header, section.footer):
            for item in iter_block_items(container):
                if isinstance(item, Paragraph):
                    text = "".join(r.text for r in item.runs) if item.runs else item.text
                    for m in LEFTOVER_RE.findall(text or ""): leftovers.add(m)
    return leftovers

# ===================== UI =====================
st.set_page_config(
    page_title="Document Generator", 
    page_icon="📄", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

inject_style()

# 초기 세션 상태
if 'generation_status' not in st.session_state:
    st.session_state.generation_status = None
if 'xlsx_uploaded' not in st.session_state:
    st.session_state.xlsx_uploaded = False
if 'docx_uploaded' not in st.session_state:
    st.session_state.docx_uploaded = False

# 헤더
st.markdown("""
<div class="header-container">
    <div class="user-profile">
        <div class="avatar">JD</div>
        <span class="username">John Doe</span>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown('<h1 class="main-title">DOCUMENT GENERATOR</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Automate Your Documents</p>', unsafe_allow_html=True)

# 메인 업로드 섹션
col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="upload-card excel-card">
        <div class="upload-icon excel-icon">
            <svg width="48" height="48" viewBox="0 0 24 24" fill="none">
                <path d="M13 2H6C5.46957 2 4.96086 2.21071 4.58579 2.58579C4.21071 2.96086 4 3.46957 4 4V20C4 20.5304 4.21071 21.0391 4.58579 21.4142C4.96086 21.7893 5.46957 22 6 22H18C18.5304 22 19.0391 21.7893 19.4142 21.4142C19.7893 21.0391 20 20.5304 20 20V9L13 2Z" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                <path d="M13 2V9H20" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
            </svg>
        </div>
        <h3 class="upload-title">UPLOAD EXCEL TEMPLATE</h3>
        <p class="upload-subtitle">then Populate Your Demo nite</p>
    </div>
    """, unsafe_allow_html=True)
    
    xlsx_file = st.file_uploader(
        "Excel", 
        type=["xlsx", "xlsm"], 
        key="xlsx_upl",
        label_visibility="collapsed"
    )
    
    if xlsx_file:
        st.session_state.xlsx_uploaded = True

with col2:
    st.markdown("""
    <div class="upload-card word-card">
        <div class="upload-icon word-icon">
            <svg width="48" height="48" viewBox="0 0 24 24" fill="none">
                <path d="M13 2H6C5.46957 2 4.96086 2.21071 4.58579 2.58579C4.21071 2.96086 4 3.46957 4 4V20C4 20.5304 4.21071 21.0391 4.58579 21.4142C4.96086 21.7893 5.46957 22 6 22H18C18.5304 22 19.0391 21.7893 19.4142 21.4142C19.7893 21.0391 20 20.5304 20 20V9L13 2Z" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                <path d="M13 2V9H20" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
            </svg>
        </div>
        <h3 class="upload-title">UPLOAD WORD TEMPLATE</h3>
        <p class="upload-subtitle">then transfer Your Demo nite</p>
    </div>
    """, unsafe_allow_html=True)
    
    docx_tpl = st.file_uploader(
        "Word", 
        type=["docx"], 
        key="docx_upl",
        label_visibility="collapsed"
    )
    
    if docx_tpl:
        st.session_state.docx_uploaded = True

# 설정 및 생성
st.markdown('<div class="settings-section">', unsafe_allow_html=True)

col_settings1, col_settings2, col_settings3 = st.columns([2, 2, 1])

with col_settings1:
    out_name = st.text_input("출력 파일명", value=DEFAULT_OUT, label_visibility="collapsed", placeholder="출력 파일명")

with col_settings2:
    sheet_choice = None
    if xlsx_file is not None:
        try:
            wb_tmp = load_workbook(filename=io.BytesIO(xlsx_file.getvalue()), data_only=True)
            sheet_choice = st.selectbox(
                "시트",
                wb_tmp.sheetnames,
                index=wb_tmp.sheetnames.index(TARGET_SHEET) if TARGET_SHEET in wb_tmp.sheetnames else 0,
                label_visibility="collapsed"
            )
        except Exception:
            pass

with col_settings3:
    generate_btn = st.button("생성", use_container_width=True, type="primary")

st.markdown('</div>', unsafe_allow_html=True)

# RECENT GENERATIONS 섹션
st.markdown('<div class="recent-section">', unsafe_allow_html=True)
st.markdown('<h2 class="section-title">RECENT GENERATIONS</h2>', unsafe_allow_html=True)

if generate_btn:
    if xlsx_file and docx_tpl:
        st.session_state.generation_status = "generating"
        
# 상태 표시
status_col1, status_col2, status_col3 = st.columns(3)

with status_col1:
    if st.session_state.generation_status == "complete":
        st.markdown("""
        <div class="status-item status-complete">
            <svg width="20" height="20" viewBox="0 0 20 20" fill="currentColor">
                <path d="M10 0C4.48 0 0 4.48 0 10C0 15.52 4.48 20 10 20C15.52 20 20 15.52 20 10C20 4.48 15.52 0 10 0ZM8 15L3 10L4.41 8.59L8 12.17L15.59 4.58L17 6L8 15Z"/>
            </svg>
            <span>COMPLETED</span>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="status-item status-inactive">
            <svg width="20" height="20" viewBox="0 0 20 20" fill="currentColor">
                <path d="M10 0C4.48 0 0 4.48 0 10C0 15.52 4.48 20 10 20C15.52 20 20 15.52 20 10C20 4.48 15.52 0 10 0ZM8 15L3 10L4.41 8.59L8 12.17L15.59 4.58L17 6L8 15Z"/>
            </svg>
            <span>COMPLETED</span>
        </div>
        """, unsafe_allow_html=True)

with status_col2:
    st.markdown("""
    <div class="status-item status-pending">
        <svg width="20" height="20" viewBox="0 0 20 20" fill="currentColor">
            <circle cx="10" cy="10" r="8" stroke="currentColor" stroke-width="2" fill="none"/>
            <path d="M10 6V10L13 13"/>
        </svg>
        <span>PENDING APPROVAL</span>
    </div>
    """, unsafe_allow_html=True)

with status_col3:
    st.markdown("""
    <div class="status-item status-error">
        <svg width="20" height="20" viewBox="0 0 20 20" fill="currentColor">
            <circle cx="10" cy="10" r="9"/>
            <line x1="6" y1="6" x2="14" y2="14" stroke="white" stroke-width="2"/>
            <line x1="14" y1="6" x2="6" y2="14" stroke="white" stroke-width="2"/>
        </svg>
        <span>ERROR: Data Mismatch</span>
    </div>
    """, unsafe_allow_html=True)

# 진행률 표시
if st.session_state.generation_status == "generating":
    st.markdown("""
    <div class="progress-container">
        <div class="progress-label">Generating Documents...</div>
        <div class="progress-bar">
            <div class="progress-fill" style="width: 75%"></div>
        </div>
        <div class="progress-percentage">75%</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# QUICK START GUIDES
st.markdown('<h2 class="section-title">QUICK START GUIDES</h2>', unsafe_allow_html=True)

# ================== 생성 실행 ==================
if generate_btn:
    if not xlsx_file or not docx_tpl:
        st.error("엑셀과 템플릿을 모두 업로드하세요.")
        st.stop()

    with st.status("문서 생성 중...", expanded=True) as status:
        try:
            st.write("1) 엑셀 로드")
            wb = load_workbook(filename=io.BytesIO(xlsx_file.read()), data_only=True)
            ws = wb[sheet_choice] if sheet_choice else (
                wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb[wb.sheetnames[0]]
            )

            st.write("2) 템플릿 로드")
            tpl_bytes = docx_tpl.read()
            doc = Document(io.BytesIO(tpl_bytes))

            st.write("3) 치환 실행")
            replacer = make_replacer(ws)
            replace_everywhere(doc, replacer)

            st.write("4) WORD 저장")
            docx_buf = io.BytesIO()
            doc.save(docx_buf); docx_buf.seek(0)
            docx_bytes = docx_buf.getvalue()

            st.write("5) PDF 변환 시도")
            pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)
            pdf_ok = pdf_bytes is not None

            st.write("6) 남은 토큰 확인")
            doc_after = Document(io.BytesIO(docx_bytes))
            leftovers = sorted(list(collect_leftover_tokens(doc_after)))

            status.update(label="완료", state="complete", expanded=False)
            st.session_state.generation_status = "complete"
        except Exception as e:
            status.update(label="오류", state="error", expanded=True)
            st.exception(e)
            st.stop()

    st.success("문서가 준비되었습니다.")
    dl_cols = st.columns(3)
    with dl_cols[0]:
        st.download_button("📄 WORD 다운로드", data=docx_bytes,
            file_name=ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True)
    with dl_cols[1]:
        st.download_button("🖨 PDF 다운로드", data=(pdf_bytes or b""),
            file_name=ensure_pdf(out_name), mime="application/pdf",
            disabled=not pdf_ok,
            use_container_width=True)
    with dl_cols[2]:
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            zf.writestr(ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT, docx_bytes)
            if pdf_ok: zf.writestr(ensure_pdf(out_name), pdf_bytes)
        zip_buf.seek(0)
        st.download_button("📦 ZIP", data=zip_buf,
            file_name=(ensure_pdf(out_name).replace(".pdf","") + "_both.zip"),
            mime="application/zip", use_container_width=True)

    if leftovers:
        with st.expander("템플릿에 남아있는 토큰"):
            st.write(", ".join(leftovers))
