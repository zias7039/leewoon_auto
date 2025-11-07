# -*- coding: utf-8 -*-
import io
import re
from datetime import datetime, date
from decimal import Decimal

import streamlit as st
from openpyxl import load_workbook
from docx import Document
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

PLACEHOLDER_RE = re.compile(r"\{\{([A-Z]+[0-9]+)\}\}")   # {{A1}}, {{B7}} ...

TARGET_SHEET = "2.  배정후 청약시"

# ---------- 값 포맷 함수 ----------
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
    # 1) 날짜 우선
    s = try_format_as_date(v)
    if s:
        return s
    # 2) 숫자 포맷
    s = fmt_number(v)
    if s:
        return s
    # 3) 일반 문자열
    return "" if v is None else str(v)

# ---------- 문서 치환 유틸 ----------
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
    """
    문단의 run 분할로 인해 토큰이 쪼개지는 문제를 피하기 위해
    모든 run을 합쳐 치환 후 run을 재구성한다.
    주의: 문단 내 세부 서식(run별 굵게/기울임 등)은 단일 run로 축약될 수 있음.
    """
    full_text = "".join(run.text for run in par.runs)
    new_text = repl_func(full_text)
    if new_text == full_text:
        return

    # 기존 run 제거
    for _ in range(len(par.runs)):
        par.runs[0].clear()
        par.runs[0].text = ""
        par.runs[0].element.getparent().remove(par.runs[0].element)

    # 새 run 하나로 삽입 (문단 스타일은 유지, run 스타일은 단일화)
    run = par.add_run(new_text)

def replace_everywhere(doc: Document, repl_func):
    # 본문
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            replace_in_paragraph(item, repl_func)

    # 머리글/바닥글
    for section in doc.sections:
        header = section.header
        footer = section.footer
        for container in (header, footer):
            for item in iter_block_items(container):
                if isinstance(item, Paragraph):
                    replace_in_paragraph(item, repl_func)

# ---------- Excel → 치환 콜백 ----------
def make_replacer(ws):
    # ws[cell]을 읽어 포맷해서 돌려주는 콜백
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

        # 2) 날짜 템플릿 치환
        today = datetime.today()
        today_str = f"{today.year}년 {today.month}월 {today.day}일"
        for token in [
            "YYYY년 MM월 DD일",
            "YYYY년    MM월    DD일",
            "YYYY 년 MM 월 DD 일",
        ]:
            replaced = replaced.replace(token, today_str)

        return replaced

    return _repl

# ---------- Streamlit UI ----------
st.title("Word 템플릿 치환 (엑셀 셀 참조)")
st.caption("· 템플릿의 {{A1}}, {{B7}} 토큰을 엑셀 값으로 치환합니다.  · 'YYYY년 MM월 DD일'은 오늘 날짜로 바뀝니다.  · 시트명은 자동으로 '2.  배정후 청약시'를 사용합니다.")

xlsx_file = st.file_uploader("엑셀 파일(.xlsx, .xlsm)", type=["xlsx", "xlsm"])
docx_tpl = st.file_uploader("워드 템플릿(.docx)", type=["docx"])

col1, col2 = st.columns(2)
with col1:
    do_strict = st.checkbox("시트명이 없으면 중단(기본은 첫 시트로 대체)", value=False)
with col2:
    out_name = st.text_input("출력 파일명", value="출력.docx")

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
        doc = Document(io.BytesIO(docx_tpl.read()))

        # 치환 실행
        replacer = make_replacer(ws)
        replace_everywhere(doc, replacer)

        # 결과 저장 → 다운로드 버튼
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)

        st.success("완료되었습니다.")
        st.download_button(
            label="결과 문서 다운로드",
            data=buf,
            file_name=out_name if out_name.strip() else "출력.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as e:
        st.exception(e)
