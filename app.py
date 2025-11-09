# -*- coding: utf-8 -*-
import io, os, re, tempfile, subprocess, shutil
from datetime import datetime, date
from decimal import Decimal
from zipfile import ZipFile, ZIP_DEFLATED

import streamlit as st
from openpyxl import load_workbook
from docx import Document
from docx.table import _Cell
from docx.text.paragraph import Paragraph

# docx→pdf (선택) : Windows에서 Word가 있으면 사용
try:
    from docx2pdf import convert as docx2pdf_convert  # noqa
except Exception:
    docx2pdf_convert = None

# -------------------- 상수/정규식 --------------------
TARGET_SHEET = "2.  배정후 청약시"
PLACEHOLDER_RE = re.compile(r"\{\{([A-Z]+[0-9]+)\}\}")     # {{A1}}, {{B7}} ...
DATE_TOKEN_RE  = re.compile(r"Y{4}\s*년\s*M{2}\s*월\s*D{2}\s*일")
TODAY          = datetime.today()
DEFAULT_BASENAME = f"{TODAY:%Y%m%d}_#_납입요청서_DB저축은행"

# -------------------- 유틸 --------------------
def which(cmd: str) -> str | None:
    return shutil.which(cmd)

def out_docx(name: str) -> str:
    n = (name or DEFAULT_BASENAME).strip()
    return n if n.lower().endswith(".docx") else n + ".docx"

def out_pdf(name: str) -> str:
    base = (name or DEFAULT_BASENAME).strip()
    return (base[:-5] if base.lower().endswith(".docx") else base) + ".pdf"

def fmt_date_like(v) -> str:
    if v is None: return ""
    if isinstance(v, (datetime, date)):
        return f"{v.year}. {v.month}. {v.day}."
    s = str(v).strip()
    try:
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
            d = datetime.strptime(s, "%Y-%m-%d")
            return f"{d.year}. {d.month}. {d.day}."
    except Exception:
        pass
    return ""

def fmt_number_like(v) -> str:
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

def to_text(v) -> str:
    return fmt_date_like(v) or fmt_number_like(v) or ("" if v is None else str(v))

# -------------------- 문서 치환 --------------------
def iter_blocks(parent):
    # 문단/표셀 재귀 순회 (본문+헤더/푸터 공통)
    if hasattr(parent, "paragraphs"):
        for p in parent.paragraphs: yield p
        for t in parent.tables:
            for r in t.rows:
                for c in r.cells:
                    yield from iter_blocks(c)
    elif isinstance(parent, _Cell):
        for p in parent.paragraphs: yield p
        for t in parent.tables:
            for r in t.rows:
                for c in r.cells:
                    yield from iter_blocks(c)

def replace_in_paragraph(par: Paragraph, repl):
    # run 단위 치환 → 실패 시 문단 전체 치환
    changed = False
    for run in par.runs:
        new = repl(run.text)
        if new != run.text:
            run.text = new
            changed = True
    if changed: return
    full = "".join(r.text for r in par.runs)
    new = repl(full)
    if new != full and par.runs:
        par.runs[0].text = new
        for r in par.runs[1:]: r.text = ""

def replace_everywhere(doc: Document, repl):
    for item in iter_blocks(doc):
        if isinstance(item, Paragraph):
            replace_in_paragraph(item, repl)
    for sec in doc.sections:
        for box in (sec.header, sec.footer):
            for item in iter_blocks(box):
                if isinstance(item, Paragraph):
                    replace_in_paragraph(item, repl)

def make_replacer(ws):
    today_str = f"{TODAY.year}년    {TODAY.month}월    {TODAY.day}일"  # 사이 공백 4칸
    def _cell_sub(m):
        try:
            return to_text(ws[m.group(1)].value)
        except Exception:
            return ""
    def _repl(text: str) -> str:
        text = PLACEHOLDER_RE.sub(_cell_sub, text)
        return DATE_TOKEN_RE.sub(today_str, text)
    return _repl

# -------------------- DOCX → PDF --------------------
def docx_bytes_to_pdf_bytes(docx_bytes: bytes) -> bytes | None:
    try:
        with tempfile.TemporaryDirectory() as td:
            in_path  = os.path.join(td, "doc.docx")
            out_path = os.path.join(td, "doc.pdf")
            with open(in_path, "wb") as f: f.write(docx_bytes)

            # 1) Word(docx2pdf) 우선
            if docx2pdf_convert:
                try:
                    docx2pdf_convert(in_path, out_path)
                    if os.path.exists(out_path):
                        return open(out_path, "rb").read()
                except Exception:
                    pass

            # 2) LibreOffice(soffice)
            if which("soffice"):
                try:
                    subprocess.run(
                        ["soffice", "--headless", "--convert-to", "pdf", in_path, "--outdir", td],
                        check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
                    )
                    if os.path.exists(out_path):
                        return open(out_path, "rb").read()
                except Exception:
                    pass
    except Exception:
        pass
    return None

# -------------------- UI --------------------
st.title("납입요청서 자동 생성 (DOCX + PDF)")

xlsx_file = st.file_uploader("엑셀 파일(.xlsx, .xlsm)", type=["xlsx", "xlsm"])
docx_tpl  = st.file_uploader("워드 템플릿(.docx)", type=["docx"])
out_name  = st.text_input("출력 파일명", value=DEFAULT_BASENAME + ".docx")
run       = st.button("문서 생성")

if run:
    if not xlsx_file or not docx_tpl:
        st.error("엑셀과 워드 템플릿을 모두 업로드하세요.")
        st.stop()

    try:
        # Excel 시트 선택
        wb = load_workbook(filename=io.BytesIO(xlsx_file.read()), data_only=True)
        ws = wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb[wb.sheetnames[0]]

        # 템플릿 로드 & 치환
        doc = Document(io.BytesIO(docx_tpl.read()))
        replace_everywhere(doc, make_replacer(ws))

        # DOCX bytes
        buf = io.BytesIO()
        doc.save(buf)
        docx_bytes = buf.getvalue()

        # PDF bytes (가능하면)
        pdf_bytes = docx_bytes_to_pdf_bytes(docx_bytes)

        # ZIP 묶어서 한 번에 내려주기
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            zf.writestr(out_docx(out_name), docx_bytes)
            if pdf_bytes:
                zf.writestr(out_pdf(out_name), pdf_bytes)
        zip_buf.seek(0)

        st.success("완료되었습니다.")
        st.download_button(
            "WORD+PDF 한번에 다운로드 (ZIP)",
            data=zip_buf,
            file_name=(out_pdf(out_name).removesuffix(".pdf") + "_both.zip"),
            mime="application/zip",
            use_container_width=True,
        )

    except Exception as e:
        st.exception(e)
