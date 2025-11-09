# -*- coding: utf-8 -*-
import io
import os
import re
import sys
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

PLACEHOLDER_RE = re.compile(r"\{\{([A-Z]+[0-9]+)\}\}")   # {{A1}}, {{B7}} ...
DEFAULT_OUT = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행.docx"
TARGET_SHEET = "2.  배정후 청약시"

# ============== 유틸 ==============
def ensure_docx(name: str) -> str:
    name = (name or "").strip()
    return name if name.lower().endswith(".docx") else (name + ".docx")

def has_cmd(cmd: str) -> bool:
    return shutil.which(cmd) is not None

def is_windows() -> bool:
    return sys.platform.startswith("win")

# ============== 값 포맷 함수 ==============
def force_font(doc, font_name="한컴바탕"):
    # 본문
    for p in doc.paragraphs:
        for r in p.runs:
            r.font.name = font_name
    # 머리글/바닥글
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

# ============== 문서 치환 유틸 ==============
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
    # 1) run 단위로 치환 시도(서식 보존)
    changed = False
    for run in par.runs:
        new_text = repl_func(run.text)
        if new_text != run.text:
            run.text = new_text
            changed = True
    if changed:
        return
    # 2) 여러 run에 걸친 토큰은 문단 전체 텍스트 기준으로 치환(최소 파괴)
    full_text = "".join(r.text for r in par.runs)
    new_text = repl_func(full_text)
    if new_text == full_text:
        return
    if par.runs:
        par.runs[0].text = new_text
        for r in par.runs[1:]:
            r.text = ""  # run 객체는 남겨 레이아웃 변화 최소화

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

# ============== Excel → 치환 콜백 ==============
def make_replacer(ws):
    def _repl(text: str) -> str:
        # 1) {{A1}} 같은 토큰 치환
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

# ============== PDF 내보내기 백엔드 ==============
def docx_bytes_to_pdf_bytes_via_soffice(docx_bytes: bytes) -> bytes:
    """LibreOffice가 있을 때 headless 변환."""
    if not has_cmd("soffice"):
        raise RuntimeError("LibreOffice(soffice) 미탐지")
    with tempfile.TemporaryDirectory() as td:
        docx_path = os.path.join(td, "tmp.docx")
        pdf_path = os.path.join(td, "tmp.pdf")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)
        cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", td, docx_path]
        run = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if run.returncode != 0 or not os.path.exists(pdf_path):
            raise RuntimeError(f"LibreOffice 변환 실패: rc={run.returncode}, err={run.stderr.decode(errors='ignore')}")
        with open(pdf_path, "rb") as f:
            return f.read()

def docx_bytes_to_pdf_bytes_via_docx2pdf(docx_bytes: bytes) -> bytes:
    """윈도우 + MS Word 환경에서 docx2pdf 사용."""
    from docx2pdf import convert  # optional import
    if not is_windows():
        raise RuntimeError("docx2pdf는 Windows+Word에서만 안정 동작")
    with tempfile.TemporaryDirectory() as td:
        docx_path = os.path.join(td, "tmp.docx")
        out_dir = td
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)
        convert(docx_path, out_dir)  # creates tmp.pdf
        pdf_path = os.path.join(td, "tmp.pdf")
        if not os.path.exists(pdf_path):
            # 일부 버전에선 파일명이 바뀌기도 함 → 확장자만 교체 탐색
            cand = [p for p in os.listdir(out_dir) if p.lower().endswith(".pdf")]
            if not cand:
                raise RuntimeError("docx2pdf 변환 실패")
            pdf_path = os.path.join(out_dir, cand[0])
        with open(pdf_path, "rb") as f:
            return f.read()

def docx_bytes_to_pdf_bytes_fallback_simple(docx_bytes: bytes) -> bytes:
    """ReportLab로 간단 렌더링(서식 단순화)."""
    # docx 파싱 → 텍스트만 꺼내서 단순 PDF
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import mm
    from docx import Document as DocxReader

    doc = DocxReader(io.BytesIO(docx_bytes))
    lines = []
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            txt = item.text.strip()
            if txt:
                lines.append(txt)
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    x = 20 * mm
    y = h - 20 * mm
    lh = 6 * mm  # line height
    for line in lines:
        # 너무 긴 줄 wrap 간단 처리
        while line:
            if len(line) > 90:
                seg = line[:90]
                line = line[90:]
            else:
                seg = line
                line = ""
            c.drawString(x, y, seg)
            y -= lh
            if y < 20 * mm:
                c.showPage()
                y = h - 20 * mm
    c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()

def export_pdf(docx_bytes: bytes) -> (bytes, str):
    """최적→대체→폴백 순서로 PDF 변환을 시도하고, 사용된 백엔드 라벨도 반환."""
    # 1) LibreOffice
    try:
        pdf_bytes = docx_bytes_to_pdf_bytes_via_soffice(docx_bytes)
        return pdf_bytes, "LibreOffice(soffice)"
    except Exception as e:
        dbg1 = str(e)

    # 2) Windows + docx2pdf
    try:
        pdf_bytes = docx_bytes_to_pdf_bytes_via_docx2pdf(docx_bytes)
        return pdf_bytes, "docx2pdf(MS Word)"
    except Exception as e:
        dbg2 = str(e)

    # 3) ReportLab fallback
    try:
        pdf_bytes = docx_bytes_to_pdf_bytes_fallback_simple(docx_bytes)
        return pdf_bytes, "ReportLab(간단 렌더링)"
    except Exception as e:
        raise RuntimeError(
            "PDF 변환 실패\n"
            f"- soffice: {dbg1}\n"
            f"- docx2pdf: {dbg2}\n"
            f"- reportlab: {e}"
        )

# ============== Streamlit UI ==============
st.title("엑셀→Word 치환 & PDF 동시 생성")

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
            # 엄격 모드 제거: 없으면 첫 시트 사용
            ws = wb[sheet_names[0]]

        # Word 템플릿 로드
        tpl_bytes = docx_tpl.read()
        doc = Document(io.BytesIO(tpl_bytes))

        # 치환 실행
        replacer = make_replacer(ws)
        replace_everywhere(doc, replacer)

        # 폰트 정리(선택)
        force_font(doc, "한컴바탕")

        # 결과 저장(DOCX)
        docx_buf = io.BytesIO()
        doc.save(docx_buf)
        docx_buf.seek(0)
        final_docx_name = ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT

        # PDF 동시 생성
        pdf_bytes, backend = export_pdf(docx_buf.getvalue())
        final_pdf_name = final_docx_name.rsplit(".", 1)[0] + ".pdf"

        # 결과 표시
        st.success(f"완료되었습니다. (PDF 변환 백엔드: {backend})")

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                label="DOCX 다운로드",
                data=docx_buf.getvalue(),
                file_name=final_docx_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        with c2:
            st.download_button(
                label="PDF 다운로드",
                data=pdf_bytes,
                file_name=final_pdf_name,
                mime="application/pdf",
            )

        # 디버그 정보
        st.caption(f"환경: platform={sys.platform}, soffice={'O' if has_cmd('soffice') else 'X'}")

    except Exception as e:
        st.exception(e)
