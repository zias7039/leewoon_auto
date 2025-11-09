# -*- coding: utf-8 -*-
import io
import os
import re
import tempfile
import subprocess
from datetime import datetime, date
from decimal import Decimal
from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook
from docx import Document
from docx.table import _Cell
from docx.text.paragraph import Paragraph

# 선택: docx2pdf가 있으면 활용(주로 Windows/Word)
try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

# ===== 상수 =====
PLACEHOLDER_RE = re.compile(r"\{\{([A-Z]+[0-9]+)\}\}")   # {{A1}}, {{B7}} ...
TARGET_SHEET = "2.  배정후 청약시"
DEFAULT_BASENAME = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행"

# ===== 유틸 =====
def ensure_docx(name: str) -> str:
    name = (name or "").strip() or DEFAULT_BASENAME
    return name if name.lower().endswith(".docx") else (name + ".docx")

def ensure_pdf(name: str) -> str:
    base = (name or "").strip() or DEFAULT_BASENAME
    base = base[:-5] if base.lower().endswith(".docx") else base
    return base + ".pdf"

def has_soffice() -> bool:
    for p in os.environ.get("PATH", "").split(os.pathsep):
        if (Path(p) / "soffice").is_file() or (Path(p) / "soffice.bin").is_file():
            return True
    return False

# ===== 값 포맷 =====
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
    return try_format_as_date(v) or fmt_number(v) or ("" if v is None else str(v))

# ===== 문서 치환(본문/표/머리글/바닥글 전체 순회) =====
def iter_block_items(container):
    """덕타이핑 기반 순회: paragraphs / tables 속성이 있으면 재귀 탐색."""
    if hasattr(container, "paragraphs"):
        for p in container.paragraphs:
            yield p
    if hasattr(container, "tables"):
        for t in container.tables:
            for row in t.rows:
                for cell in row.cells:
                    # _Cell 또는 임베디드 테이블까지 재귀
                    yield from iter_block_items(cell)

def replace_in_paragraph(par: Paragraph, repl_func):
    # 1) run 단위 치환(서식 보존)
    changed = False
    for run in par.runs:
        new_text = repl_func(run.text)
        if new_text != run.text:
            run.text = new_text
            changed = True
    if changed:
        return
    # 2) 여러 run 걸친 토큰은 문단 전체 텍스트 기준 치환
    full_text = "".join(r.text for r in par.runs)
    new_text = repl_func(full_text)
    if new_text != full_text and par.runs:
        par.runs[0].text = new_text
        for r in par.runs[1:]:
            r.text = ""

def replace_everywhere(doc: Document, repl_func):
    # 본문
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            replace_in_paragraph(item, repl_func)
    # 머리글/바닥글
    for section in doc.sections:
        for container in (section.header, section.footer):
            for item in iter_block_items(container):
                if isinstance(item, Paragraph):
                    replace_in_paragraph(item, repl_func)

# ===== Excel → 치환 콜백 =====
def make_replacer(ws):
    def _repl(text: str) -> str:
        # {{A1}} 같은 셀 참조
        def cell_sub(m):
            addr = m.group(1)
            try:
                v = ws[addr].value
            except Exception:
                v = None
            return value_to_text(v)

        replaced = PLACEHOLDER_RE.sub(cell_sub, text)

        # 날짜 템플릿(YYYY/MM/DD 변형) 치환
        sp = "    "  # 4칸
        today = datetime.today()
        today_str = f"{today.year}년{sp}{today.month}월{sp}{today.day}일"
        for token in ("YYYY년 MM월 DD일", "YYYY년    MM월    DD일", "YYYY 년 MM 월 DD 일"):
            replaced = replaced.replace(token, today_str)
        return replaced
    return _repl

# ===== DOCX → PDF (docx2pdf 또는 soffice) =====
def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> bytes | None:
    try:
        with tempfile.TemporaryDirectory() as td:
            in_path = Path(td) / "doc.docx"
            out_path = Path(td) / "doc.pdf"
            in_path.write_bytes(docx_bytes)

            # 1) Word(Windows) 환경: docx2pdf
            if docx2pdf_convert is not None:
                try:
                    docx2pdf_convert(str(in_path), str(out_path))
                    if out_path.exists():
                        return out_path.read_bytes()
                except Exception:
                    pass

            # 2) LibreOffice headless
            if has_soffice():
                try:
                    cmd = [
                        "soffice",
                        "--headless",
                        "--convert-to",
                        "pdf",
                        str(in_path),
                        "--outdir",
                        str(Path(td)),
                    ]
                    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    if out_path.exists():
                        return out_path.read_bytes()
                except Exception:
                    pass
    except Exception:
        pass
    return None

# ===== Streamlit UI =====
st.title("납입요청서 자동 생성 (DOCX + PDF)")

xlsx_file = st.file_uploader("엑셀 파일(.xlsx, .xlsm)", type=["xlsx", "xlsm"])
docx_tpl = st.file_uploader("워드 템플릿(.docx)", type=["docx"])
out_name = st.text_input("출력 파일명", value=DEFAULT_BASENAME)

if st.button("문서 생성"):
    if not xlsx_file or not docx_tpl:
        st.error("엑셀 파일과 워드 템플릿을 모두 업로드하세요.")
        st.stop()

    try:
        # Excel 로드
        wb = load_workbook(filename=io.BytesIO(xlsx_file.read()), data_only=True)
        ws = wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb[wb.sheetnames[0]]

        # Word 템플릿 로드
        doc = Document(io.BytesIO(docx_tpl.read()))

        # 치환
        replacer = make_replacer(ws)
        replace_everywhere(doc, replacer)

        # DOCX 결과 메모리 저장
        docx_buf = io.BytesIO()
        doc.save(docx_buf)
        docx_bytes = docx_buf.getvalue()

        # PDF 변환 시도
        pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)

        # ZIP(WORD + PDF(있으면))
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            zf.writestr(ensure_docx(out_name), docx_bytes)
            if pdf_bytes:
                zf.writestr(ensure_pdf(out_name), pdf_bytes)
        zip_buf.seek(0)

        st.success("완료되었습니다.")
        st.download_button(
            "WORD+PDF 한번에 다운로드 (ZIP)",
            data=zip_buf,
            file_name=f"{Path(ensure_pdf(out_name)).stem}_both.zip",
            mime="application/zip",
        )

        # 원하면 개별 파일도 제공(선택)
        with st.expander("개별 파일로 받기 (선택)"):
            st.download_button(
                "DOCX만 다운로드",
                data=docx_bytes,
                file_name=ensure_docx(out_name),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            if pdf_bytes:
                st.download_button(
                    "PDF만 다운로드",
                    data=pdf_bytes,
                    file_name=ensure_pdf(out_name),
                    mime="application/pdf",
                )
            else:
                st.info("PDF 변환 불가: 서버에 MS Word(docx2pdf)나 LibreOffice(soffice)가 없어 보입니다.")

    except Exception as e:
        st.exception(e)
