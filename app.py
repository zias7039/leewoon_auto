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

# 스타일 모듈 (중요)
from ui_style import inject as inject_style, open_div, close_div, h4

try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

TOKEN_RE = re.compile(r"\{\{([A-Z]+[0-9]+)(?:\|([^}]+))?\}\}")
LEFTOVER_RE = re.compile(r"\{\{[^}]+\}\}")
DEFAULT_OUT = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행.docx"
TARGET_SHEET = "2.  배정후 청약시"

def ensure_docx(name): return name if name.lower().endswith(".docx") else (name + ".docx")
def ensure_pdf(name): return (name[:-5] if name.lower().endswith(".docx") else name) + ".pdf"

def has_soffice():
    return any(os.path.isfile(os.path.join(p, "soffice")) for p in os.environ.get("PATH","").split(os.pathsep))

def try_format_as_date(v):
    try:
        if v is None: return ""
        if isinstance(v, (datetime, date)): return f"{v.year}. {v.month}. {v.day}."
        s = str(v).strip()
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
            dt = datetime.strptime(s, "%Y-%m-%d").date()
            return f"{dt.year}. {dt.month}. {dt.day}."
    except: pass
    return ""

def fmt_number(v):
    try:
        if isinstance(v,(int,float,Decimal)): return f"{float(v):,.0f}"
        if isinstance(v,str):
            raw=v.replace(",","")
            if re.fullmatch(r"-?\d+(\.\d+)?", raw):
                return f"{float(raw):,.0f}"
    except: pass
    return ""

def value_to_text(v):
    return try_format_as_date(v) or fmt_number(v) or ("" if v is None else str(v))

def apply_inline_format(value, fmt):
    if not fmt: return value_to_text(value)
    if any(tok in fmt for tok in ("YYYY","MM","DD")):
        if isinstance(value,str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}",value.strip()):
            value=datetime.strptime(value.strip(),"%Y-%m-%d").date()
        if isinstance(value,(datetime,date)):
            return value.strftime(fmt.replace("YYYY","%Y").replace("MM","%m").replace("DD","%d"))
        return value_to_text(value)
    if re.fullmatch(r"[#,0]+(?:\.[0#]+)?", fmt.replace(",","")):
        try:
            num=float(str(value).replace(",",""))
            decimals = len(fmt.split(".")[1]) if "." in fmt else 0
            return f"{num:,.{decimals}f}"
        except: return value_to_text(value)
    return value_to_text(value)

def iter_block_items(parent):
    if hasattr(parent,"paragraphs"):
        for p in parent.paragraphs: yield p
        for t in parent.tables:
            for row in t.rows:
                for cell in row.cells:
                    yield from iter_block_items(cell)
    elif isinstance(parent,_Cell):
        for p in parent.paragraphs: yield p
        for t in parent.tables:
            for row in t.rows:
                for cell in row.cells:
                    yield from iter_block_items(cell)

def replace_in_paragraph(par,repl):
    changed=False
    for run in par.runs:
        new=repl(run.text)
        if new!=run.text: run.text=new; changed=True
    if changed: return
    full="".join(r.text for r in par.runs)
    new=repl(full)
    if new==full: return
    par.runs[0].text=new
    for r in par.runs[1:]: r.text=""

def replace_everywhere(doc,repl):
    for item in iter_block_items(doc):
        if isinstance(item,Paragraph):
            replace_in_paragraph(item,repl)

def make_replacer(ws):
    def repl(text):
        def sub(m):
            addr,fmt=m.group(1),m.group(2)
            try: v=ws[addr].value
            except: v=None
            return apply_inline_format(v,fmt)
        return TOKEN_RE.sub(sub,text)
    return repl

def convert_docx_to_pdf_bytes(docx_bytes):
    try:
        with tempfile.TemporaryDirectory() as td:
            in_path=os.path.join(td,"doc.docx")
            out_path=os.path.join(td,"doc.pdf")
            open(in_path,"wb").write(docx_bytes)
            if docx2pdf_convert:
                try:
                    docx2pdf_convert(in_path,out_path)
                    if os.path.exists(out_path): return open(out_path,"rb").read()
                except: pass
            if has_soffice():
                try:
                    subprocess.run(["soffice","--headless","--convert-to","pdf",in_path,"--outdir",td],check=True)
                    if os.path.exists(out_path): return open(out_path,"rb").read()
                except: pass
    except: pass
    return None

def collect_leftover_tokens(doc):
    leftovers=set()
    for item in iter_block_items(doc):
        if isinstance(item,Paragraph):
            text="".join(r.text for r in item.runs)
            leftovers |= set(LEFTOVER_RE.findall(text))
    return leftovers

# ================================= UI ================================= #
st.set_page_config(page_title="납입요청서 자동 생성", page_icon="🧾", layout="wide")
inject_style()

st.title("🧾 납입요청서 자동 생성 (DOCX + PDF)")

col_left, col_right = st.columns([1.2,1])

with col_left:
    open_div("upload-card")
    with st.form("input_form"):
        xlsx_file = st.file_uploader("엑셀 파일", type=["xlsx","xlsm"])
        docx_tpl  = st.file_uploader("워드 템플릿(.docx)", type=["docx"])
        out_name = st.text_input("출력 파일명", value=DEFAULT_OUT)
        if xlsx_file:
            try:
                wb_tmp=load_workbook(io.BytesIO(xlsx_file.getvalue()),data_only=True)
                sheet_choice = st.selectbox("사용할 시트", wb_tmp.sheetnames)
            except:
                sheet_choice=None
        else: sheet_choice=None
        submitted = st.form_submit_button("문서 생성", use_container_width=True)
    close_div()

with col_right:
    h4("안내")
    st.markdown("- **{{A1}}**, **{{B7|YYYY.MM.DD}}**, **{{C3|#,###.00}}** 포맷 지원")
    st.markdown("- PDF 변환은 Word 또는 LibreOffice 필요")

if submitted:
    if not xlsx_file or not docx_tpl:
        st.error("엑셀과 템플릿을 모두 업로드하세요.")
        st.stop()

    wb = load_workbook(io.BytesIO(xlsx_file.read()),data_only=True)
    ws = wb[sheet_choice]
    doc = Document(io.BytesIO(docx_tpl.read()))
    replace_everywhere(doc, make_replacer(ws))

    buf=io.BytesIO(); doc.save(buf); buf.seek(0); docx_bytes = buf.getvalue()
    pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)
    leftovers = collect_leftover_tokens(Document(io.BytesIO(docx_bytes)))

    st.success("문서 생성 완료")

    c1,c2,c3 = st.columns(3)
    c1.download_button("📄 WORD", docx_bytes, file_name=ensure_docx(out_name))
    c2.download_button("🖨 PDF", pdf_bytes or b"", file_name=ensure_pdf(out_name), disabled=(pdf_bytes is None))
    z=io.BytesIO()
    with ZipFile(z,"w",ZIP_DEFLATED) as f:
        f.writestr(ensure_docx(out_name),docx_bytes)
        if pdf_bytes: f.writestr(ensure_pdf(out_name),pdf_bytes)
    z.seek(0)
    c3.download_button("📦 ZIP", z, file_name="export.zip")

    if leftovers:
        st.warning("템플릿에 남아있는 토큰: " + ", ".join(leftovers))
