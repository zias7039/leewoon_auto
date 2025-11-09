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
from docx import Document           # <- 팩토리 함수(클래스 아님)
from docx.text.paragraph import Paragraph
from docx.table import _Cell

# 선택: docx2pdf 사용 가능 시 활용(주로 Windows+Word)
try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

# -------------------- 상수/정규식 --------------------
PLACEHOLDER_RE = re.compile(r"\{\{([A-Z]+[0-9]+)\}\}")   # {{A1}}, {{B7}} ...
TARGET_SHEET = "2.  배정후 청약시"
DEFAULT_BASENAME = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행"

# -------------------- 유틸 --------------------
def ensure_ext(name: str, ext: str) -> str:
    base = (name or "").strip() or DEFAULT_BASENAME
    if not ext.startswith("."):
        ext = "." + ext
    if base.lower().endswith(ext.lower()):
        return base
    # .docx -> .pdf처럼 바꿔 달라는 경우를 고려
    root, old = os.path.splitext(base)
    if ext.lower() == ".pdf" and old.lower() == ".docx":
        return root + ".pdf"
    return base + ext

def has_soffice() -> bool:
    paths = os.environ.get("PATH", "").split(os.pathsep)
    for p in paths:
        for binname in ("soffice", "soffice.bin", "libreoffice"):
            if os.path.isfile(os.path.join(p, binname)):
                return True
    return False

# -------------------- 포맷터 --------------------
def _try_format_as_date(v) -> str:
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

def _fmt_number(v) -> str:
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
    return _try_format_as_date(v) or _fmt_number(v) or ("" if v is None else str(v))

# -------------------- 문서 순회/치환 --------------------
def iter_block_items(container):
    """
    Document/Section Header/Footer/Table Cell 등 다양한 컨테이너에서
    문단과 셀 내부를 재귀적으로 순회. 타입체크에 의존하지 않고 hasattr 기반.
    """
    # 문단
    if hasattr(container, "paragraphs"):
        for p in container.paragraphs:
            yield p
    # 표
    if hasattr(container, "tables"):
        for t in container.tables:
            for row in t.rows:
                for cell in row.cells:
                    # 셀은 _Cell이지만 타입체크 없이 재귀
                    yield from iter_block_items(cell)

def replace_in_paragraph(par: Paragraph, repl_func):
    # 1) run 단위 치환 시도(서식 보존)
    changed = False
    for run in par.runs:
        new_text = repl_func(run.text)
        if new_text != run.text:
            run.text = new_text
            changed = True
    if changed:
        return
    # 2) 여러 run에 걸친 토큰은 문단 전체 텍스트로 재치환
    full = "".join(r.text for r in par.runs)
    new_full = repl_func(full)
    if new_full != full and par.runs:
        par.runs[0].text = new_full
        for r in par.runs[1:]:
            r.text = ""

def replace_everywhere(doc):
    # 본문
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            replace_in_paragraph(item, _REPLACER)

    # 머리글/바닥글
    for section in doc.sections:
        for container in (section.header, section.footer):
            for item in iter_block_items(container):
                if isinstance(item, Paragraph):
                    replace_in_paragraph(item, _REPLACER)

# -------------------- 치환 콜백 (Excel 의존) --------------------
def make_replacer(ws):
    # 날짜 템플릿(YYYY년 MM월 DD일) → 오늘 날짜(사이 공백 4칸)
    sp = "    "
    today = datetime.today()
    today_str = f"{today.year}년{sp}{today.month}월{sp}{today.day}일"
    tokens = {
        "YYYY년 MM월 DD일",
        "YYYY년    MM월    DD일",
        "YYYY 년 MM 월 DD 일",
    }

    def _repl(text: str) -> str:
        def cell_sub(m):
            addr = m.group(1)
            try:
                v = ws[addr].value
            except Exception:
                v = None
            return value_to_text(v)

        replaced = PLACEHOLDER_RE.sub(cell_sub, text)
        # 날짜 템플릿 치환
        for t in tokens:
            if t in replaced:
                replaced = replaced.replace(t, today_str)
        return replaced

    return _repl

# -------------------- DOCX → PDF --------------------
def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> bytes | None:
    """
    가능한 경우 PDF bytes 반환.
    1) Windows+Word: docx2pdf
    2) 리눅스/서버: LibreOffice(soffice) headless
    실패 시 None
    """
    try:
        with tempfile.TemporaryDirectory() as td:
            in_path = os.path.join(td, "doc.docx")
            out_path = os.path.join(td, "doc.pdf")
            with open(in_path, "wb") as f:
                f.write(docx_bytes)

            # 1) Word 기반
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
                    cmd = ["soffice", "--headless", "--convert-to", "pdf", in_path, "--outdir", td]
                    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    if os.path.exists(out_path):
                        with open(out_path, "rb") as f:
                            return f.read()
                except Exception:
                    pass

    except Exception:
        pass

    return None

# -------------------- Streamlit UI --------------------
st.title("납입요청서 자동 생성 (DOCX + PDF)")

xlsx_file = st.file_uploader("엑셀 파일(.xlsx, .xlsm)", type=["xlsx", "xlsm"])
docx_tpl = st.file_uploader("워드 템플릿(.docx)", type=["docx"])

out_base = st.text_input("출력 파일명(확장자 없이 입력 권장)", value=DEFAULT_BASENAME)
run = st.button("문서 생성")

if run:
    if not xlsx_file or not docx_tpl:
        st.error("엑셀 파일과 워드 템플릿을 모두 업로드하세요.")
        st.stop()

    try:
        # Excel 로드
        wb = load_workbook(filename=io.BytesIO(xlsx_file.read()), data_only=True)
        ws = wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb[wb.sheetnames[0]]

        # Word 템플릿 로드 + 치환
        tpl_bytes = docx_tpl.read()
        doc = Document(io.BytesIO(tpl_bytes))

        # 치환기 준비 (전역 접근 위해 바인딩)
        _REPLACER = make_replacer(ws)
        replace_everywhere(doc)

        # DOCX bytes
        docx_buf = io.BytesIO()
        doc.save(docx_buf)
        docx_bytes = docx_buf.getvalue()

        # PDF 변환 시도
        pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)

        # ZIP 묶기(한 번 클릭으로 두 파일 내려받기)
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            zf.writestr(ensure_ext(out_base, ".docx"), docx_bytes)
            if pdf_bytes is not None:
                zf.writestr(ensure_ext(out_base, ".pdf"), pdf_bytes)
        zip_buf.seek(0)

        st.success("완료되었습니다.")
        st.download_button(
            "WORD+PDF 한번에 다운로드 (ZIP)",
            data=zip_buf,
            file_name=ensure_ext(out_base, ".zip"),
            mime="application/zip",
        )

        # 옵션: 각각도 제공하고 싶다면 주석 해제
        # st.download_button("DOCX만 다운로드", data=docx_bytes,
        #                   file_name=ensure_ext(out_base, ".docx"),
        #                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        # if pdf_bytes is not None:
        #     st.download_button("PDF만 다운로드", data=pdf_bytes,
        #                        file_name=ensure_ext(out_base, ".pdf"),
        #                        mime="application/pdf")

    except Exception as e:
        st.exception(e)
