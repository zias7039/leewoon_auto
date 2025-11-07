# -*- coding: utf-8 -*-
import io
import re
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

# ---------- 유틸 ----------
def ensure_docx(name: str) -> str:
    name = (name or "").strip()
    return name if name.lower().endswith(".docx") else (name + ".docx")

# ---------- 값 포맷 함수 ----------
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
    # 1) run 단위로 치환 시도(서식 보존)
    changed = False
    for run in par.runs:
        new_text = repl_func(run.text)
        if new_text != run.text:
            run.text = new_text
            changed = True
    if changed:
        return
    # 2) 여러 run에 걸친 토큰만 문단 전체 텍스트 기준으로 치환(최소 파괴)
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

# ---------- Excel → 치환 콜백 ----------
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

# ---------- Streamlit UI ----------
st.title("Word 템플릿 치환 (python-docx)")
st.caption("· 템플릿의 {{A1}}, {{B7}} 토큰을 엑셀 값으로 치환 · 'YYYY년 MM월 DD일'은 공백 4칸으로 오늘 날짜로 치환")

xlsx_file = st.file_uploader("엑셀 파일(.xlsx, .xlsm)", type=["xlsx", "xlsm"])
docx_tpl = st.file_uploader("워드 템플릿(.docx)", type=["docx"])

c1, c2 = st.columns(2)
with c1:
    do_strict = st.checkbox("시트명이 없으면 중단", value=False)  # ← 정의 추가
with c2:
    unify_font = st.checkbox("문서 전체 글꼴을 한컴바탕으로 통일", value=False)

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

        # 글꼴 통일 옵션
        if unify_font:
            force_font(doc, "한컴바탕")

        # 결과 저장 → 다운로드 버튼
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)

        st.success("완료되었습니다.")
        st.download_button(
            label="결과 문서 다운로드",
            data=buf,
            file_name=ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as e:
        st.exception(e)
