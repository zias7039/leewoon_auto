# -*- coding: utf-8 -*-
import io
import os
import re
import shutil
import tempfile
import subprocess
from datetime import datetime, date
from decimal import Decimal

import streamlit as st
from openpyxl import load_workbook
from docx import Document
from docx.table import _Cell
from docx.text.paragraph import Paragraph

# 선택적 의존성: 있으면 사용
try:
    import aspose.words as aw
    HAS_ASPOSE = True
except Exception:
    HAS_ASPOSE = False

try:
    from docx2pdf import convert as docx2pdf_convert
    HAS_DOCX2PDF = True
except Exception:
    HAS_DOCX2PDF = False

PLACEHOLDER_RE = re.compile(r"\{\{([A-Z]+[0-9]+)\}\}")
DEFAULT_OUT = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행.docx"
TARGET_SHEET = "2.  배정후 청약시"

# ---------- 유틸 ----------
def ensure_docx(name: str) -> str:
    name = (name or "").strip()
    return name if name.lower().endswith(".docx") else (name + ".docx")

def ensure_pdf(name: str) -> str:
    name = (name or "").strip()
    base = name[:-5] if name.lower().endswith(".docx") else name
    return f"{base}.pdf"

# ---------- 값 포맷 ----------
def force_font(doc, font_name="한컴바탕"):
    for p in doc.paragraphs:
        for r in p.runs:
            r.font.name = font_name
    for section in doc.sections:
        for hdrftr in (section.header, section.footer):
            for p in hdrftr.paragraphs:
                for r in p.runs:
                    r.font.name = font_name

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

# ---------- 문서 치환 ----------
def iter_block_items(parent):
    if hasattr(parent, "element") and hasattr(parent, "paragraphs"):
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

# ---------- Excel → 치환 콜백 ----------
def make_replacer(ws):
    def _repl(text: str) -> str:
        def cell_sub(m):
            addr = m.group(1)
            try:
                v = ws[addr].value
            except Exception:
                v = None
            return value_to_text(v)

        replaced = PLACEHOLDER_RE.sub(cell_sub, text)

        sp = "    "
        today = datetime.today()
        today_str = f"{today.year}년{sp}{today.month}월{sp}{today.day}일"
        for token in ["YYYY년 MM월 DD일", "YYYY년    MM월    DD일", "YYYY 년 MM 월 DD 일"]:
            replaced = replaced.replace(token, today_str)
        return replaced
    return _repl

# ---------- DOCX → PDF (바이트 변환) ----------
def docx_bytes_to_pdf_bytes(docx_bytes: bytes) -> bytes | None:
    """
    가능한 방법 순서:
    1) aspose-words
    2) Windows + docx2pdf (MS Word)
    3) soffice(headless)
    실패 시 None
    """
    # 1) Aspose
    if HAS_ASPOSE:
        try:
            in_stream = io.BytesIO(docx_bytes)
            adoc = aw.Document(in_stream)
            out_stream = io.BytesIO()
            adoc.save(out_stream, aw.SaveFormat.PDF)
            return out_stream.getvalue()
        except Exception:
            pass

    # 2) Windows + docx2pdf
    if os.name == "nt" and HAS_DOCX2PDF:
        try:
            with tempfile.TemporaryDirectory() as td:
                in_path = os.path.join(td, "temp.docx")
                out_path = os.path.join(td, "temp.pdf")
                with open(in_path, "wb") as f:
                    f.write(docx_bytes)
                docx2pdf_convert(in_path, out_path)
                with open(out_path, "rb") as f:
                    return f.read()
        except Exception:
            pass

    # 3) soffice (LibreOffice)
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            with tempfile.TemporaryDirectory() as td:
                in_path = os.path.join(td, "temp.docx")
                with open(in_path, "wb") as f:
                    f.write(docx_bytes)
                cmd = [
                    soffice, "--headless", "--nologo", "--nofirststartwizard",
                    "--convert-to", "pdf", "--outdir", td, in_path
                ]
                subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                out_path = os.path.join(td, "temp.pdf")
                if os.path.exists(out_path):
                    with open(out_path, "rb") as f:
                        return f.read()
        except Exception:
            pass

    return None

# ---------- Streamlit UI ----------
xlsx_file = st.file_uploader("엑셀 파일(.xlsx, .xlsm)", type=["xlsx", "xlsm"])
docx_tpl = st.file_uploader("워드 템플릿(.docx)", type=["docx"])

col1, = st.columns(1)
with col1:
    out_name = st.text_input("출력 파일명", value=DEFAULT_OUT)

run = st.button("문서 생성")

if run:
    if not xlsx_file or not docx_tpl:
        st.error("엑셀 파일과 워드 템플릿을 모두 업로드하세요.")
        st.stop()

    try:
        # Excel
        wb = load_workbook(filename=io.BytesIO(xlsx_file.read()), data_only=True)
        sheet_names = wb.sheetnames
        ws = wb[TARGET_SHEET] if TARGET_SHEET in sheet_names else wb[sheet_names[0]]

        # Word 템플릿 로드 & 치환
        tpl_bytes = docx_tpl.read()
        doc = Document(io.BytesIO(tpl_bytes))
        replacer = make_replacer(ws)
        replace_everywhere(doc, replacer)
        # 필요시 폰트 강제
        # force_font(doc, "한컴바탕")

        # DOCX 결과 저장(메모리)
        docx_buf = io.BytesIO()
        doc.save(docx_buf)
        docx_buf.seek(0)
        docx_filename = ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT
        pdf_filename = ensure_pdf(docx_filename)

        st.success("DOCX 생성 완료")
        st.download_button(
            label="DOCX 다운로드",
            data=docx_buf.getvalue(),
            file_name=docx_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        # PDF 변환 시도
        pdf_bytes = docx_bytes_to_pdf_bytes(docx_buf.getvalue())
        if pdf_bytes:
            st.success("PDF 변환 완료")
            st.download_button(
                label="PDF 다운로드",
                data=pdf_bytes,
                file_name=pdf_filename,
                mime="application/pdf",
            )
        else:
            st.info("PDF 변환 환경이 없어 DOCX만 제공합니다. (aspose-words 설치 또는 로컬/서버에 MS Word 또는 LibreOffice 필요)")

    except Exception as e:
        st.exception(e)
