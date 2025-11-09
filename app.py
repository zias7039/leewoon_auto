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

# -------------------- 상수/정규식 --------------------
PLACEHOLDER_RE = re.compile(r"\{\{([A-Z]+[0-9]+)\}\}")  # {{A1}}, {{B7}}...
DATE_TOKEN_RE  = re.compile(r"Y{4}\s*년\s*M{2}\s*월\s*D{2}\s*일")  # 공백 가변 허용
TARGET_SHEET   = "2.  배정후 청약시"
DEFAULT_BASENAME = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행"

# -------------------- 유틸 --------------------
def basename(name: str | None) -> str:
    """확장자 제거된 베이스 이름 반환"""
    n = (name or "").strip() or DEFAULT_BASENAME
    return n[:-5] if n.lower().endswith(".docx") else n

def to_docx(name: str) -> str:
    b = basename(name)
    return b + ".docx"

def to_pdf(name: str) -> str:
    b = basename(name)
    return b + ".pdf"

def has_soffice() -> bool:
    paths = os.environ.get("PATH", "").split(os.pathsep)
    cand = ("soffice", "soffice.bin", "soffice.exe")
    return any(os.path.isfile(os.path.join(p, c)) for p in paths for c in cand)

def value_to_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return f"{v.year}. {v.month}. {v.day}."
    # 숫자 혹은 숫자문자열 -> 천단위
    if isinstance(v, (int, float, Decimal)):
        return f"{float(v):,.0f}"
    if isinstance(v, str):
        s = v.strip()
        # YYYY-MM-DD -> 한국식
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
            dt = datetime.strptime(s, "%Y-%m-%d").date()
            return f"{dt.year}. {dt.month}. {dt.day}."
        raw = s.replace(",", "")
        if re.fullmatch(r"-?\d+(\.\d+)?", raw):
            return f"{float(raw):,.0f}"
        return s
    return str(v)

# -------------------- 문서 치환 --------------------
def iter_paragraphs(container):
    """본문/표 셀/헤더/푸터를 모두 Paragraph로 평탄화"""
    if isinstance(container, Document):
        # 본문
        for p in container.paragraphs:
            yield p
        for t in container.tables:
            for r in t.rows:
                for c in r.cells:
                    yield from iter_paragraphs(c)
        # 헤더/푸터
        for sec in container.sections:
            for hf in (sec.header, sec.footer):
                for p in hf.paragraphs:
                    yield p
                for t in hf.tables:
                    for r in t.rows:
                        for c in r.cells:
                            yield from iter_paragraphs(c)
    elif isinstance(container, _Cell):
        for p in container.paragraphs:
            yield p
        for t in container.tables:
            for r in t.rows:
                for c in r.cells:
                    yield from iter_paragraphs(c)

def replace_paragraph_text(par: Paragraph, repl_func):
    # run 단위 치환(서식 보존) → 실패 시 문단 전체 치환
    changed = False
    for run in par.runs:
        new = repl_func(run.text)
        if new != run.text:
            run.text = new
            changed = True
    if changed:
        return
    full = "".join(r.text for r in par.runs)
    new = repl_func(full)
    if new != full and par.runs:
        par.runs[0].text = new
        for r in par.runs[1:]:
            r.text = ""

def make_replacer(ws):
    today = datetime.today()
    today_str = f"{today.year}년    {today.month}월    {today.day}일"
    def _repl(text: str) -> str:
        # 1) {{A1}} 치환
        def cell_sub(m):
            addr = m.group(1)
            try:
                v = ws[addr].value
            except Exception:
                v = None
            return value_to_text(v)
        s = PLACEHOLDER_RE.sub(cell_sub, text)
        # 2) YYYY/MM/DD 한글 템플릿(공백 가변) 치환
        s = DATE_TOKEN_RE.sub(today_str, s)
        # 3) {{TODAY}} 같은 심플 토큰도 지원(선택)
        s = s.replace("{{TODAY}}", today_str)
        return s
    return _repl

# -------------------- DOCX → PDF --------------------
def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> bytes | None:
    try:
        with tempfile.TemporaryDirectory() as td:
            docx_path = os.path.join(td, "doc.docx")
            pdf_path  = os.path.join(td, "doc.pdf")
            with open(docx_path, "wb") as f:
                f.write(docx_bytes)

            # 1) MS Word (Windows)
            if docx2pdf_convert is not None:
                try:
                    docx2pdf_convert(docx_path, pdf_path)
                    if os.path.exists(pdf_path):
                        return open(pdf_path, "rb").read()
                except Exception:
                    pass

            # 2) LibreOffice
            if has_soffice():
                try:
                    cmd = ["soffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", td]
                    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    if os.path.exists(pdf_path):
                        return open(pdf_path, "rb").read()
                except Exception:
                    pass
    except Exception:
        pass
    return None

# -------------------- Streamlit UI --------------------
st.title("납입요청서 자동 생성 (DOCX + PDF)")

xlsx_file = st.file_uploader("엑셀 (.xlsx, .xlsm)", type=["xlsx", "xlsm"])
docx_tpl  = st.file_uploader("워드 템플릿 (.docx)", type=["docx"])
out_name  = st.text_input("출력 파일명", value=DEFAULT_BASENAME)

if st.button("문서 생성"):
    if not xlsx_file or not docx_tpl:
        st.error("엑셀/워드 템플릿을 모두 업로드하세요.")
        st.stop()
    try:
        xlsx_bytes = xlsx_file.read()
        tpl_bytes  = docx_tpl.read()

        wb = load_workbook(filename=io.BytesIO(xlsx_bytes), data_only=True)
        ws = wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb[wb.sheetnames[0]]

        doc = Document(io.BytesIO(tpl_bytes))
        replacer = make_replacer(ws)

        for p in iter_paragraphs(doc):
            replace_paragraph_text(p, replacer)

        # DOCX 메모리 저장
        mem_docx = io.BytesIO()
        doc.save(mem_docx)
        doc_bytes = mem_docx.getvalue()

        # PDF 변환 시도
        pdf_bytes = convert_docx_to_pdf_bytes(doc_bytes)

        # ZIP 구성
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            zf.writestr(to_docx(out_name), doc_bytes)
            if pdf_bytes:
                zf.writestr(to_pdf(out_name), pdf_bytes)
        zip_buf.seek(0)

        st.success("완료되었습니다.")
        st.download_button(
            "WORD+PDF 한번에 다운로드 (ZIP)",
            data=zip_buf,
            file_name=basename(out_name) + "_both.zip",
            mime="application/zip",
            use_container_width=True,
        )

    except Exception as e:
        st.exception(e)
