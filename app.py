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

try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

# ---------- 상수 ----------
TOKEN_RE = re.compile(r"\{\{([A-Z]+[0-9]+)(?:\|([^}]+))?\}\}")
LEFTOVER_RE = re.compile(r"\{\{[^}]+\}\}")
DEFAULT_OUT = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행.docx"
TARGET_SHEET = "2.  배정후 청약시"

# ---------- 유틸 ----------
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

def apply_inline_format(value, fmt: str | None) -> str:
    if fmt is None or fmt.strip() == "":
        return value_to_text(value)

    if any(tok in fmt for tok in ("YYYY", "MM", "DD")):
        if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value.strip()):
            value = datetime.strptime(value.strip(), "%Y-%m-%d").date()
        if isinstance(value, (datetime, date)):
            f = fmt.replace("YYYY", "%Y").replace("MM", "%m").replace("DD", "%d")
            return value.strftime(f)
        return value_to_text(value)

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

def iter_block_items(parent):
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
            with open(in_path, "wb") as f:
                f.write(docx_bytes)

            if docx2pdf_convert is not None:
                try:
                    docx2pdf_convert(in_path, out_path)
                    if os.path.exists(out_path):
                        with open(out_path, "rb") as f:
                            return f.read()
                except Exception:
                    pass

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

# ---------- UI ----------
st.set_page_config(page_title="납입요청서 자동 생성", page_icon="🧾", layout="wide")

# Glassmorphism + 브랜드 컬러
st.markdown("""
<style>
/* 공통 Glassmorphism 토큰 */
:root{
  --glass-bg: rgba(15, 23, 42, .35);         /* 유리 배경 */
  --glass-bd: rgba(148, 163, 184, .35);      /* 테두리 */
  --glass-shadow: 0 8px 32px rgba(0,0,0,.35);
}

/* 래퍼 공통 카드 느낌 */
.upload-wrap{
  border-radius: 16px;
  padding: 12px;
  margin: 8px 0 18px 0;
  position: relative;
  background: linear-gradient(180deg, rgba(255,255,255,.06), rgba(255,255,255,.02));
  border: 1px solid var(--glass-bd);
  box-shadow: var(--glass-shadow);
  backdrop-filter: blur(10px);
}

/* 시그니처 컬러 변수 */
.excel-upload{ --brand:#107C41; }   /* MS Excel green */
.word-upload { --brand:#185ABD; }   /* MS Word blue  */

/* 업로더 드롭존 자체를 정확히 타겟팅 */
.upload-wrap [data-testid="stFileUploaderDropzone"]{
  background: var(--glass-bg) !important;
  border: 1px solid color-mix(in srgb, var(--brand) 45%, #ffffff 0%) !important;
  border-radius: 12px !important;
  transition: border-color .2s ease, box-shadow .2s ease, background .2s ease;
  box-shadow: inset 0 0 0 1px rgba(255,255,255,.06);
}

/* 호버/포커스 */
.upload-wrap [data-testid="stFileUploaderDropzone"]:hover{
  border-color: color-mix(in srgb, var(--brand) 70%, #ffffff 0%) !important;
  background: rgba(15,23,42,.42) !important;
}

/* 내부 텍스트/아이콘 컬러 */
.upload-wrap [data-testid="stFileUploader"] *{
  color: color-mix(in srgb, var(--brand) 80%, #e5e7eb 20%) !important;
}

/* Browse 버튼 */
.upload-wrap [data-testid="stFileUploader"] button{
  border-radius: 10px !important;
  background: linear-gradient(180deg, color-mix(in srgb, var(--brand) 85%, #ffffff 0%), color-mix(in srgb, var(--brand) 65%, #000000 0%)) !important;
  border: 1px solid color-mix(in srgb, var(--brand) 90%, #000 10%) !important;
}
.upload-wrap [data-testid="stFileUploader"] button:hover{
  filter: brightness(1.05);
}

/* 파일 확장자·용량 캡션 가독성 */
.upload-wrap [data-testid="stFileUploader"] small, 
.upload-wrap [data-testid="stFileUploader"] p, 
.upload-wrap [data-testid="stFileUploader"] span{
  color: rgba(226,232,240,.9) !important;
}

/* (스트림릿 버전 호환용) 베이스웹 드롭존에도 적용 */
.upload-wrap [data-baseweb="dropzone"]{
  background: var(--glass-bg) !important;
  border: 1px solid color-mix(in srgb, var(--brand) 45%, #ffffff 0%) !important;
  border-radius: 12px !important;
}
</style>
""", unsafe_allow_html=True)

st.title("🧾 납입요청서 자동 생성 (DOCX + PDF)")

col_left, col_right = st.columns([1.2, 1])
with col_left:
    with st.form("input_form", clear_on_submit=False):
        st.markdown('<div class="excel-upload glass-uploader">', unsafe_allow_html=True)
        xlsx_file = st.file_uploader("엑셀 파일", type=["xlsx", "xlsm"], accept_multiple_files=False)
        st.markdown('</div>', unsafe_allow_html=True)

    # 워드 업로더 (워드 블루 / Glass UI)
        st.markdown('<div class="word-upload glass-uploader">', unsafe_allow_html=True)
        docx_tpl = st.file_uploader("워드 템플릿(.docx)", type=["docx"], accept_multiple_files=False)
        st.markdown('</div>', unsafe_allow_html=True)

        out_name = st.text_input("출력 파일명", value=DEFAULT_OUT)

        sheet_choice = None
        if xlsx_file is not None:
            try:
                wb_tmp = load_workbook(filename=io.BytesIO(xlsx_file.getvalue()), data_only=True)
                sheet_choice = st.selectbox(
                    "사용할 시트",
                    wb_tmp.sheetnames,
                    index=wb_tmp.sheetnames.index(TARGET_SHEET) if TARGET_SHEET in wb_tmp.sheetnames else 0
                )
            except Exception:
                st.warning("엑셀 미리보기 중 문제가 발생했습니다. 생성 시도는 가능합니다.")

        submitted = st.form_submit_button("문서 생성", use_container_width=True)

with col_right:
    st.markdown("#### 안내")
    st.markdown(
        "- `{{A1}}`, `{{B7|YYYY.MM.DD}}`, `{{C3|#,###.00}}` 포맷 지원\n"
        "- 생성 시 WORD와 PDF 각각 다운로드 + ZIP 제공\n"
        "- PDF 변환은 MS Word(docx2pdf) 또는 LibreOffice(soffice) 필요"
    )
    if 'word_up' in st.session_state and st.session_state['word_up'] is not None:
        try:
            doc_preview = Document(io.BytesIO(st.session_state['word_up'].getvalue()))
            sample_tokens = set()
            for p in doc_preview.paragraphs[:80]:
                for m in re.findall(r"\{\{[^}]+\}\}", p.text or ""):
                    if len(sample_tokens) < 12:
                        sample_tokens.add(m)
            st.markdown("**템플릿 토큰 샘플**" if sample_tokens else "템플릿에서 토큰을 찾지 못했습니다.")
            if sample_tokens:
                st.code(", ".join(list(sample_tokens)))
        except Exception:
            st.caption("템플릿 미리보기를 불러오지 못했습니다.")

# ---------- 생성 실행 ----------
if submitted:
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
            doc.save(docx_buf)
            docx_buf.seek(0)
            docx_bytes = docx_buf.getvalue()

            st.write("5) PDF 변환 시도")
            pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)
            pdf_ok = pdf_bytes is not None

            st.write("6) 남은 토큰 확인")
            doc_after = Document(io.BytesIO(docx_bytes))
            leftovers = sorted(list(collect_leftover_tokens(doc_after)))

            status.update(label="완료", state="complete", expanded=False)
        except Exception as e:
            status.update(label="오류", state="error", expanded=True)
            st.exception(e)
            st.stop()

    st.success("문서가 준비되었습니다.")
    dl_cols = st.columns(3)
    with dl_cols[0]:
        st.download_button(
            "📄 WORD 다운로드",
            data=docx_bytes,
            file_name=ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with dl_cols[1]:
        st.download_button(
            "🖨 PDF 다운로드",
            data=pdf_bytes if pdf_ok else b"",
            file_name=ensure_pdf(out_name),
            mime="application/pdf",
            disabled=not pdf_ok,
            help=None if pdf_ok else "PDF 변환 엔진(Word 또는 LibreOffice)이 없는 환경입니다.",
            use_container_width=True,
        )
    with dl_cols[2]:
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            zf.writestr(ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT, docx_bytes)
            if pdf_ok:
                zf.writestr(ensure_pdf(out_name), pdf_bytes)
        zip_buf.seek(0)
        st.download_button(
            "📦 ZIP (WORD+PDF)",
            data=zip_buf,
            file_name=(ensure_pdf(out_name).replace(".pdf", "") + "_both.zip"),
            mime="application/zip",
            use_container_width=True,
        )

    if leftovers:
        with st.expander("템플릿에 남아있는 토큰"):
            st.write(", ".join(leftovers))
    else:
        st.caption("모든 토큰이 치환되었습니다.")
