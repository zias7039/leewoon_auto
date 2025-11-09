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

# 선택: docx2pdf(윈도우/오피스) 있으면 먼저 사용
try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

# ----------------- 기본 설정 -----------------
st.set_page_config(page_title="Document Generator", layout="wide")

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

# ----------------- 인라인 포맷 -----------------
def apply_inline_format(value, fmt: str | None) -> str:
    if not fmt:
        return value_to_text(value)

    # 날짜 포맷 (YYYY/MM/DD 등)
    if any(tok in fmt for tok in ("YYYY", "MM", "DD")):
        if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value.strip()):
            value = datetime.strptime(value.strip(), "%Y-%m-%d").date()
        if isinstance(value, (datetime, date)):
            f = fmt.replace("YYYY", "%Y").replace("MM", "%m").replace("DD", "%d")
            return value.strftime(f)
        return value_to_text(value)

    # 숫자 포맷 (#,###.00 등)
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
    # python-docx 타입 이름에 의존하지 않고 duck-typing
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

        # 단순 템플릿(YYYY년 MM월 DD일) 오늘 날짜로 치환
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
            with open(in_path, "wb"):
                _.write(docx_bytes)  # intentional NameError? No. We'll write properly below.

    except Exception:
        pass
    # (위에서 변수 오타 방지 재작성)
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

# ----------------- 남은 토큰 수집(정보용) -----------------
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

# ----------------- 스타일 (카드와 업로더를 한 덩어리로) -----------------
st.markdown("""
<style>
:root{
  --card-bg: rgba(2,6,23,.65);
  --card-bd: rgba(148,163,184,.25);
  --dz-bg: rgba(17,24,39,.55);
}
h1,h2,h3{ letter-spacing:.2px }
.page-title{ font-size:40px; font-weight:800; color:#fff; margin: 8px 0 2px }
.page-sub{ color:#9ca3af; margin-bottom: 24px }

.upload-grid{ display:grid; grid-template-columns: 1fr 1fr; gap: 28px }
.upload-card{
  background: var(--card-bg);
  border: 1px solid var(--card-bd);
  border-radius: 18px;
  padding: 22px;
  box-shadow: 0 10px 34px rgba(0,0,0,.28);
}

/* Streamlit 업로더 요소가 카드 안에서 꽉 차도록 */
.upload-card [data-testid="stFileUploader"]{ width:100%; }
.upload-card [data-testid="stFileUploader"] > div{ width:100%; }
.upload-card [data-testid="stFileUploaderDropzone"]{
  background: var(--dz-bg);
  border: 1px solid var(--card-bd);
  border-radius: 12px;
}
.upload-title{ font-weight:800; font-size:20px; color:#e5e7eb; margin-bottom:6px }
.upload-sub{ color:#94a3b8; font-size:13px; margin-bottom:14px }
</style>
""", unsafe_allow_html=True)

# ----------------- UI -----------------
st.markdown('<div class="page-title">DOCUMENT GENERATOR</div>', unsafe_allow_html=True)
st.markdown('<div class="page-sub">Automate Your Documents</div>', unsafe_allow_html=True)

# 업로더 카드 2개
st.markdown('<div class="upload-grid">', unsafe_allow_html=True)

# Excel 카드
with st.container():
    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.markdown('<div class="upload-title">UPLOAD EXCEL TEMPLATE</div>', unsafe_allow_html=True)
    st.markdown('<div class="upload-sub">엑셀 템플릿(.xlsx / .xlsm)</div>', unsafe_allow_html=True)
    excel_file = st.file_uploader("", type=["xlsx","xlsm"], label_visibility="collapsed", key="excel")
    st.markdown('</div>', unsafe_allow_html=True)

# Word 카드
with st.container():
    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.markdown('<div class="upload-title">UPLOAD WORD TEMPLATE</div>', unsafe_allow_html=True)
    st.markdown('<div class="upload-sub">워드 템플릿(.docx)</div>', unsafe_allow_html=True)
    docx_tpl = st.file_uploader("", type=["docx"], label_visibility="collapsed", key="docx")
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)  # .upload-grid end

out_name = st.text_input("출력 파일명", value=DEFAULT_OUT, label_visibility="visible")

run = st.button("문서 생성하기", use_container_width=True)

# ----------------- 동작 -----------------
if run:
    if not excel_file or not docx_tpl:
        st.error("엑셀 파일과 워드 템플릿을 모두 업로드하세요.")
        st.stop()

    try:
        # Excel 로드 (+ 시트 선택 자동화: 기본 TARGET 있으면 그걸로)
        wb_tmp = load_workbook(filename=io.BytesIO(excel_file.getvalue()), data_only=True)
        ws = wb_tmp[TARGET_SHEET] if TARGET_SHEET in wb_tmp.sheetnames else wb_tmp[wb_tmp.sheetnames[0]]

        # Word 템플릿 로드
        tpl_bytes = docx_tpl.getvalue()
        doc = Document(io.BytesIO(tpl_bytes))

        # 치환
        replacer = make_replacer(ws)
        replace_everywhere(doc, replacer)

        # DOCX 메모리 저장
        docx_buf = io.BytesIO()
        doc.save(docx_buf)
        docx_buf.seek(0)
        docx_bytes = docx_buf.getvalue()

        # PDF 변환 (가능 시)
        pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)

        # ZIP 묶어 단일 다운로드
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            zf.writestr(ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT, docx_bytes)
            if pdf_bytes:
                zf.writestr(ensure_pdf(out_name), pdf_bytes)
        zip_buf.seek(0)

        # 남은 토큰(정보용)
        leftovers = sorted(list(collect_leftover_tokens(Document(io.BytesIO(docx_bytes)))))
        if leftovers:
            with st.expander("템플릿에 남은 치환 토큰(참고)"):
                st.write(", ".join(leftovers))

        st.success("완료되었습니다.")
        st.download_button(
            "WORD + PDF 한번에 다운로드 (ZIP)",
            data=zip_buf,
            file_name=(ensure_pdf(out_name).replace(".pdf", "") + "_both.zip"),
            mime="application/zip",
            use_container_width=True,
        )
    except Exception as e:
        st.exception(e)
