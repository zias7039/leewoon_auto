# -*- coding: utf-8 -*-
import io
import re
from datetime import datetime, date
from decimal import Decimal

import streamlit as st
from openpyxl import load_workbook
from docxtpl import DocxTemplate

# =========================
# 설정값
# =========================
TARGET_SHEET = "2.  배정후 청약시"
CELL_TOKEN_RE = re.compile(r"\{\{([A-Z]+[0-9]+)\}\}")  # {{A1}}, {{B7}} ...
SPACER = "    "  # 년/월/일 사이 공백 4칸
DEFAULT_OUT = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행.docx"


# =========================
# 값 포맷 함수
# =========================
def try_format_as_date(v):
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


def fmt_number(v):
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


def value_to_text(v):
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


def ensure_docx(name: str) -> str:
    name = (name or "").strip()
    return name if name.lower().endswith(".docx") else (name + ".docx")


# =========================
# 전체 폰트 강제 통일(선택)
# =========================
def force_font(doc, font_name="한컴바탕"):
    """
    경고: 문서 전체 run의 글꼴을 지정 폰트로 통일한다.
    굵기/기울임/색/크기는 유지되지만, '다른 폰트' 의도는 사라진다.
    """
    # 본문
    for p in doc.paragraphs:
        for r in p.runs:
            r.font.name = font_name
    # 머리글/바닥글
    for section in doc.sections:
        for container in (section.header, section.footer):
            for p in container.paragraphs:
                for r in p.runs:
                    r.font.name = font_name


# =========================
# Streamlit UI
# =========================
st.title("Word 템플릿 치환 (엑셀 셀 → docx)")
st.caption(
    "· 템플릿의 {{A1}}, {{B7}} 토큰을 엑셀 값으로 치환합니다. "
    "· 날짜는 {{TODAY}} 토큰을 사용하세요(년/월/일 사이 공백 4칸). "
    "· 시트는 자동으로 '2.  배정후 청약시'를 선택합니다(없으면 첫 시트)."
)

xlsx_file = st.file_uploader("엑셀 파일(.xlsx, .xlsm)", type=["xlsx", "xlsm"])
docx_tpl = st.file_uploader("워드 템플릿(.docx)", type=["docx"])

c1, c2 = st.columns(2)
with c1:
    strict = st.checkbox("시트명이 없으면 중단", value=False)
with c2:
    unify_font = st.checkbox("문서 전체 폰트를 한컴바탕으로 통일", value=False)

col1, = st.columns(1)
with col1:
    out_name = st.text_input("출력 파일명", value=DEFAULT_OUT)

if st.button("문서 생성"):
    if not xlsx_file or not docx_tpl:
        st.error("엑셀 파일과 워드 템플릿을 모두 업로드하세요.")
        st.stop()

    try:
        # -------- Excel 로드 --------
        wb = load_workbook(filename=io.BytesIO(xlsx_file.read()), data_only=True)
        sheet_names = wb.sheetnames
        if TARGET_SHEET in sheet_names:
            ws = wb[TARGET_SHEET]
        else:
            if strict:
                st.error(f"시트 '{TARGET_SHEET}' 를 찾지 못했습니다.")
                st.stop()
            ws = wb[sheet_names[0]]

        # -------- 템플릿에서 필요한 셀 토큰 스캔 --------
        tpl_bytes = docx_tpl.read()
        # docxtpl은 컨텍스트 키만 치환하므로, 템플릿에 실제 쓰인 키만 수집
        text_for_scan = tpl_bytes.decode("utf-8", errors="ignore")
        needed_cells = set(m.group(1) for m in CELL_TOKEN_RE.finditer(text_for_scan))

        # -------- 컨텍스트 구성 --------
        ctx = {}
        for addr in needed_cells:
            try:
                v = ws[addr].value
            except Exception:
                v = None
            ctx[addr] = value_to_text(v)

        # 날짜(YYYY    MM    DD 간격 4칸) -> {{TODAY}} 로 넣기
        today = datetime.today()
        ctx["TODAY"] = f"{today.year}년{SPACER}{today.month}월{SPACER}{today.day}일"

        # -------- 템플릿 렌더 --------
        doc = DocxTemplate(io.BytesIO(tpl_bytes))
        doc.render(ctx)

        # 필요시 전체 폰트 통일
        if unify_font:
            force_font(doc, "한컴바탕")

        # -------- 결과 저장/다운로드 --------
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)

        st.success("완료되었습니다.")
        st.download_button(
            "결과 문서 다운로드",
            data=buf,
            file_name=ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        st.info("팁: 템플릿에서 날짜 위치는 'YYYY년 MM월 DD일'처럼 고정 텍스트 대신 {{TODAY}} 토큰을 사용하세요. 토큰이 들어간 그 자리에 준 서식(폰트/크기)이 유지됩니다.")

    except Exception as e:
        st.exception(e)
