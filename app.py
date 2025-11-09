# -*- coding: utf-8 -*-
import io
import os
import re
import sys
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

# 선택: docx2pdf가 있으면 활용 (Windows+Word)
try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

# ===================== 설정/상수 =====================
PLACEHOLDER_RE = re.compile(r"\{\{([A-Z]+[0-9]+)(?:\|([^}]+))?\}\}")  # {{A1}} or {{A1|#,###}} or {{B7|YYYY.MM.DD}}
RAW_TOKEN_RE   = re.compile(r"\{\{[^}]+\}\}")
TARGET_SHEET_DEFAULT = "2.  배정후 청약시"
DEFAULT_OUT = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행.docx"

# ===================== 유틸 =====================
def ensure_docx(name: str) -> str:
    name = (name or "").strip()
    return name if name.lower().endswith(".docx") else (name + ".docx")

def ensure_pdf(name: str) -> str:
    base = (name or "output").strip()
    if base.lower().endswith(".docx"):
        base = base[:-5]
    return base + ".pdf"

def has_soffice() -> bool:
    paths = os.environ.get("PATH", "").split(os.pathsep)
    return any(
        os.path.isfile(os.path.join(p, "soffice")) or os.path.isfile(os.path.join(p, "soffice.bin"))
        for p in paths
    )

def coerce_date(val):
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val
    s = str(val).strip()
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    return None

def format_by_pattern(value, pattern, thousands_sep=",", decimal_sep="."):
    """
    pattern:
      - 날짜: 'YYYY.MM.DD', 'YYYY-MM-DD', 'YYYY년 MM월 DD일' 등
      - 숫자: '#,###', '#,###.0', '#,###.00' 등 (자리수 0 개수로 소수 자릿수 결정)
    """
    # 날짜 패턴 감지
    if any(k in pattern for k in ("YYYY", "YY")):
        dt = coerce_date(value)
        if not dt:
            return ""
        # 토큰 치환용
        y = f"{dt.year:04d}"
        yy = f"{dt.year % 100:02d}"
        m = f"{dt.month:02d}"
        d = f"{dt.day:02d}"
        out = (pattern
               .replace("YYYY", y)
               .replace("YY", yy)
               .replace("MM", m)
               .replace("DD", d))
        return out

    # 숫자 패턴
    # 기본 소수 자릿수 추출
    decimals = 0
    if "." in pattern:
        decimals = pattern.split(".")[1].count("0")  # '#,###.00' -> 2
    # 값 파싱
    try:
        if isinstance(value, (int, float, Decimal)):
            num = float(value)
        else:
            raw = str(value).replace(",", "").replace(" ", "")
            num = float(raw)
    except Exception:
        return ""
    # 기본 포맷은 python 포맷 사용(, . 고정) → 후처리로 로케일 구분자 적용
    py_fmt = f"{{:,.{decimals}f}}"
    s = py_fmt.format(num)
    if thousands_sep != "," or decimal_sep != ".":
        s = s.replace(",", "§").replace(".", "¤")
        s = s.replace("§", thousands_sep).replace("¤", decimal_sep)
    # 소수 0 제거 옵션이 필요하면 여기서 trim 가능(요청 시)
    return s

def value_to_text(v, thousands_sep=",", decimal_sep="."):
    # 날짜 우선
    dt = coerce_date(v)
    if dt:
        return f"{dt.year}. {dt.month}. {dt.day}."
    # 숫자
    try:
        if isinstance(v, (int, float, Decimal)):
            s = f"{v:,.0f}"
        else:
            raw = str(v).replace(",", "")
            if re.fullmatch(r"-?\d+(\.\d+)?", raw):
                s = f"{float(raw):,.0f}"
            else:
                return "" if v is None else str(v)
        if thousands_sep != "," or decimal_sep != ".":
            s = s.replace(",", thousands_sep).replace(".", decimal_sep)
        return s
    except Exception:
        return "" if v is None else str(v)

# ===================== 문서 순회/치환 =====================
def iter_block_items(container):
    """
    python-docx 객체를 안전하게 순회 (본문/헤더/푸터/표 셀 재귀)
    - Document 클래스를 isinstance로 검사하지 않고 duck-typing으로 처리
    """
    # Paragraph 리스트
    if hasattr(container, "paragraphs"):
        for p in container.paragraphs:
            yield p
    # Tables 재귀
    if hasattr(container, "tables"):
        for t in container.tables:
            for row in t.rows:
                for cell in row.cells:
                    # _Cell인 경우에도 동일 로직 재귀
                    for item in iter_block_items(cell):
                        yield item
    # _Cell 방어: (위에서 이미 처리되지만 안전망)
    if isinstance(container, _Cell):
        for p in container.paragraphs:
            yield p
        for t in container.tables:
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
    # run 쪼개짐 케이스: 문단 단위 재치환
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
    # 헤더/푸터
    for section in doc.sections:
        for container in (section.header, section.footer):
            for item in iter_block_items(container):
                if isinstance(item, Paragraph):
                    replace_in_paragraph(item, repl_func)

def collect_raw_tokens(doc: Document):
    """치환 후 템플릿에 남아있는 {{...}} 토큰 수집"""
    found = set()
    def scan(container):
        if hasattr(container, "paragraphs"):
            for p in container.paragraphs:
                for m in RAW_TOKEN_RE.findall(p.text or ""):
                    found.add(m)
        if hasattr(container, "tables"):
            for t in container.tables:
                for row in t.rows:
                    for cell in row.cells:
                        scan(cell)
    scan(doc)
    for section in doc.sections:
        for container in (section.header, section.footer):
            scan(container)
    return sorted(found)

# ===================== Excel → 치환 콜백 =====================
def make_replacer(ws, thousands_sep=",", decimal_sep=".", strict_missing_cell=False):
    def _repl(text: str) -> str:
        def cell_sub(m):
            addr, fmt = m.group(1), m.group(2)
            try:
                v = ws[addr].value
            except Exception:
                v = None
                if strict_missing_cell:
                    # 엄격 모드: 없는 셀 발견 시 빈문자 대신 명확히 표시
                    return f"<!MISS:{addr}!>"
            if fmt:  # 인라인 포맷 파이프
                return format_by_pattern(v, fmt, thousands_sep, decimal_sep)
            return value_to_text(v, thousands_sep, decimal_sep)

        replaced = PLACEHOLDER_RE.sub(cell_sub, text)

        # 보너스: 흔한 'YYYY년    MM월    DD일' 같은 고정 텍스트 치환
        sp = "    "
        today = datetime.today()
        today_str = f"{today.year}년{sp}{today.month}월{sp}{today.day}일"
        for token in ("YYYY년 MM월 DD일", "YYYY년    MM월    DD일", "YYYY 년 MM 월 DD 일"):
            replaced = replaced.replace(token, today_str)
        return replaced
    return _repl

# ===================== DOCX → PDF =====================
def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> bytes | None:
    """
    가능한 경우 PDF로 변환해 bytes 반환.
    1) Windows + MS Word: docx2pdf
    2) soffice(libreooffice) 있으면 headless 변환
    실패 시 None
    """
    try:
        with tempfile.TemporaryDirectory() as td:
            in_path = os.path.join(td, "doc.docx")
            out_path = os.path.join(td, "doc.pdf")
            with open(in_path, "wb") as f:
                f.write(docx_bytes)

            # 1) docx2pdf
            if docx2pdf_convert is not None:
                try:
                    docx2pdf_convert(in_path, out_path)
                    if os.path.exists(out_path):
                        return open(out_path, "rb").read()
                except Exception:
                    pass

            # 2) LibreOffice
            if has_soffice():
                try:
                    cmd = ["soffice", "--headless", "--convert-to", "pdf", in_path, "--outdir", td]
                    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    if os.path.exists(out_path):
                        return open(out_path, "rb").read()
                except Exception:
                    pass
    except Exception:
        pass
    return None

# ===================== Streamlit UI =====================
st.title("납입요청서 자동 생성 (DOCX + PDF)")

xlsx_file = st.file_uploader("엑셀 파일(.xlsx, .xlsm)", type=["xlsx", "xlsm"])
docx_tpl  = st.file_uploader("워드 템플릿(.docx)", type=["docx"])

# 사이드바: 옵션
with st.sidebar:
    st.subheader("옵션")
    thousands_sep = st.selectbox("천단위 구분자", [",", ".", " "], index=0)
    decimal_sep   = st.selectbox("소수점 구분자", [".", ","], index=0)
    strict_mode   = st.checkbox("엄격 모드 (시트/토큰 오류 시 중단)", value=False)
    strict_missing_cell = st.checkbox("엄격: 누락 셀 마킹(<!MISS:AXX!>)", value=False)
    st.caption("인라인 포맷 예: {{A1|#,###}}, {{B7|#,###.00}}, {{C3|YYYY.MM.DD}}")

out_name = st.text_input("출력 파일명", value=DEFAULT_OUT)

# Excel 시트 선택
selected_sheet = None
if xlsx_file:
    try:
        wb_probe = load_workbook(filename=io.BytesIO(xlsx_file.getvalue()), data_only=True, read_only=True)
        sheets = wb_probe.sheetnames
        default_idx = sheets.index(TARGET_SHEET_DEFAULT) if TARGET_SHEET_DEFAULT in sheets else 0
        selected_sheet = st.selectbox("데이터 시트 선택", sheets, index=default_idx)
    except Exception as e:
        st.warning(f"시트 정보를 읽는 중 문제가 발생했습니다: {e}")

# 액션 버튼
if st.button("문서 생성"):
    if not xlsx_file or not docx_tpl:
        st.error("엑셀 파일과 워드 템플릿을 모두 업로드하세요.")
        st.stop()

    try:
        # ---- Excel 로드
        wb = load_workbook(filename=io.BytesIO(xlsx_file.getvalue()), data_only=True)
        sheet_names = wb.sheetnames
        if selected_sheet and selected_sheet in sheet_names:
            ws = wb[selected_sheet]
        else:
            if TARGET_SHEET_DEFAULT in sheet_names:
                ws = wb[TARGET_SHEET_DEFAULT]
            else:
                if strict_mode:
                    st.error(f"시트 '{TARGET_SHEET_DEFAULT}'를 찾을 수 없습니다.")
                    st.stop()
                ws = wb[sheet_names[0]]

        # ---- DOCX 템플릿 로드
        tpl_bytes = docx_tpl.getvalue()
        doc = Document(io.BytesIO(tpl_bytes))

        # ---- 치환
        replacer = make_replacer(ws, thousands_sep, decimal_sep, strict_missing_cell)
        replace_everywhere(doc, replacer)

        # ---- 남은 토큰 리포트
        leftover = collect_raw_tokens(doc)
        with st.sidebar:
            st.subheader("남은 토큰")
            if leftover:
                st.code("\n".join(leftover))
                if strict_mode:
                    st.error("엄격 모드: 남은 토큰이 있어 중단합니다.")
                    st.stop()
            else:
                st.caption("모든 토큰이 치환되었습니다.")

        # ---- DOCX 메모리 저장
        docx_buf = io.BytesIO()
        doc.save(docx_buf)
        docx_buf.seek(0)
        docx_bytes = docx_buf.getvalue()

        # ---- PDF 변환 시도
        pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)

        # ---- ZIP 묶기 (DOCX + 가능하면 PDF)
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            zf.writestr(ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT, docx_bytes)
            if pdf_bytes:
                zf.writestr(ensure_pdf(out_name), pdf_bytes)
        zip_buf.seek(0)

        # ---- 다운로드
        st.success("완료되었습니다.")
        st.download_button(
            "WORD+PDF 한번에 다운로드 (ZIP)",
            data=zip_buf,
            file_name=(ensure_pdf(out_name).replace(".pdf", "") + "_both.zip"),
            mime="application/zip",
        )

    except Exception as e:
        st.exception(e)
