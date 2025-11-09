# -*- coding: utf-8 -*-
import io
import os
import re
import tempfile
import subprocess
from datetime import datetime, date
from decimal import Decimal

import streamlit as st
from openpyxl import load_workbook
from docx import Document
from docx.table import _Cell
from docx.text.paragraph import Paragraph

# ===== 설정/상수 =====
PLACEHOLDER_RE = re.compile(r"\{\{([A-Z]+[0-9]+)\}\}")   # {{A1}}, {{B7}} ...
DEFAULT_OUT = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행.docx"
TARGET_SHEET = "2.  배정후 청약시"
do_strict = False  # 시트 미존재 시 에러 여부

# ===== 유틸 =====
def ensure_docx(name: str) -> str:
    name = (name or "").strip()
    return name if name.lower().endswith(".docx") else (name + ".docx")

def ensure_pdf(name: str) -> str:
    name = (name or "").strip()
    return re.sub(r"\.docx?$", "", name, flags=re.I) + ".pdf"

# ===== 값 포맷 함수 =====
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

# ===== 문서 치환 유틸 =====
def iter_block_items(parent):
    """문서의 문단/표 셀 모두 순회 (본문, 헤더/푸터 공통 사용)."""
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
    # 1) run 단위 치환
    changed = False
    for run in par.runs:
        new_text = repl_func(run.text)
        if new_text != run.text:
            run.text = new_text
            changed = True
    if changed:
        return
    # 2) 문단 전체 텍스트 기준 치환
    full_text = "".join(r.text for r in par.runs)
    new_text = repl_func(full_text)
    if new_text == full_text:
        return
    if par.runs:
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
        # 1) {{A1}} 같은 토큰
        def cell_sub(m):
            addr = m.group(1)
            try:
                v = ws[addr].value
            except Exception:
                v = None
            return value_to_text(v)

        replaced = PLACEHOLDER_RE.sub(cell_sub, text)

        # 2) 날짜 템플릿 치환 (년/월/일 사이 공백 4칸)
        sp = "    "
        today = datetime.today()
        today_str = f"{today.year}년{sp}{today.month}월{sp}{today.day}일"
        for token in [
            "YYYY년 MM월 DD일",
            "YYYY년    MM월    DD일",
            "YYYY 년 MM 월 DD 일",
        ]:
            replaced = replaced.replace(token, today_str)
        return replaced
    return _repl

# ===== DOCX → PDF 변환 =====
def convert_docx_to_pdf(input_path: str, output_path: str) -> None:
    """
    1순위: docx2pdf (Windows/Mac, MS Word 필요)
    2순위: LibreOffice 'soffice' CLI
    둘 다 실패하면 예외 발생
    """
    # 1) docx2pdf 시도
    try:
        import docx2pdf  # type: ignore
        # docx2pdf.convert(in, out)는 out이 파일 경로가 아니라 디렉토리인 경우가 많아
        # 직접 파일명 지정이 필요하면 임시 폴더에 변환 후 rename
        tmp_dir = tempfile.mkdtemp()
        docx2pdf.convert(input_path, tmp_dir)  # 같은 파일명.pdf로 생성
        base = os.path.splitext(os.path.basename(input_path))[0] + ".pdf"
        gen_pdf = os.path.join(tmp_dir, base)
        if not os.path.exists(gen_pdf):
            raise RuntimeError("docx2pdf: 출력 파일 생성 실패")
        os.replace(gen_pdf, output_path)
        return
    except Exception:
        pass  # 다음 방법으로 이어감

    # 2) LibreOffice 시도
    try:
        # soffice --headless --convert-to pdf --outdir <dir> <file>
        outdir = os.path.dirname(output_path) or "."
        result = subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, input_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            check=False,
            text=True,
        )
        # 변환 결과 확인
        if not os.path.exists(output_path):
            # LibreOffice는 보통 같은 파일명.pdf로 outdir에 생성
            base_pdf = os.path.splitext(os.path.basename(input_path))[0] + ".pdf"
            candidate = os.path.join(outdir, base_pdf)
            if os.path.exists(candidate):
                os.replace(candidate, output_path)
            else:
                raise RuntimeError(f"LibreOffice 변환 실패\n{result.stdout}")
        return
    except Exception as e:
        raise RuntimeError(f"PDF 변환 불가: docx2pdf/LibreOffice 미설치 또는 실행 실패\n{e}")

# ===== Streamlit UI =====
st.title("납입요청서 생성기 (DOCX + PDF)")

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
        # Excel 로드
        wb = load_workbook(filename=io.BytesIO(xlsx_file.read()), data_only=True)
        sheet_names = wb.sheetnames
        if TARGET_SHEET in sheet_names:
            ws = wb[TARGET_SHEET]
        else:
            if do_strict:
                st.error(f"시트 '{TARGET_SHEET}' 를 찾지 못했습니다.")
                st.stop()
            ws = wb[sheet_names[0]]

        # Word 템플릿 로드
        tpl_bytes = docx_tpl.read()
        doc = Document(io.BytesIO(tpl_bytes))

        # 치환 실행
        replacer = make_replacer(ws)
        replace_everywhere(doc, replacer)
        # (선택) 폰트 강제 적용이 필요하면 주석 해제
        # force_font(doc, "한컴바탕")

        # 결과 저장 (DOCX 바이트 + 임시 파일)
        buf_docx = io.BytesIO()
        doc.save(buf_docx)
        buf_docx.seek(0)

        # 파일명 정리
        docx_name = ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT
        pdf_name = ensure_pdf(docx_name)

        # 임시 경로에 저장 후 PDF 변환
        with tempfile.TemporaryDirectory() as td:
            tmp_docx_path = os.path.join(td, docx_name)
            tmp_pdf_path = os.path.join(td, pdf_name)

            # DOCX 저장
            with open(tmp_docx_path, "wb") as f:
                f.write(buf_docx.getbuffer())

            # PDF 변환 시도
            pdf_ok = True
            pdf_err = ""
            try:
                convert_docx_to_pdf(tmp_docx_path, tmp_pdf_path)
            except Exception as e:
                pdf_ok = False
                pdf_err = str(e)

            # 최종 다운로드용 바이트 준비
            with open(tmp_docx_path, "rb") as f:
                final_docx = f.read()
            final_pdf = None
            if pdf_ok and os.path.exists(tmp_pdf_path):
                with open(tmp_pdf_path, "rb") as f:
                    final_pdf = f.read()

        st.success("문서 생성 완료.")
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                label="DOCX 다운로드",
                data=final_docx,
                file_name=docx_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        with c2:
            if final_pdf:
                st.download_button(
                    label="PDF 다운로드",
                    data=final_pdf,
                    file_name=pdf_name,
                    mime="application/pdf",
                )
            else:
                st.warning("PDF 변환 실패: 로컬에 Microsoft Word(docx2pdf) 또는 LibreOffice(soffice)가 설치되어 있어야 합니다.\n"
                           "Windows/Mac: docx2pdf 설치 후 사용 권장\nLinux/서버: LibreOffice 설치 후 사용 권장")

    except Exception as e:
        st.exception(e)
