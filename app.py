# -*- coding: utf-8 -*-
import io
import os
import re
import tempfile
import subprocess
from datetime import datetime, date
from decimal import Decimal
from zipfile import ZipFile, ZIP_DEFLATED
from typing import Iterable, Tuple, List, Set

import streamlit as st
from openpyxl import load_workbook
from docx import Document
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

# 선택: docx2pdf가 있으면 활용 (Windows/Word)
try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

# =========================
# 전역 설정
# =========================
DEFAULT_OUT_BASENAME = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행"
DEFAULT_OUT_DOCX = DEFAULT_OUT_BASENAME + ".docx"
TARGET_SHEET_CANDIDATE = "2.  배정후 청약시"

TOKEN_RE = re.compile(r"\{\{([A-Z]+[0-9]+)(\|[^}]+)?\}\}")  # {{A1}} 또는 {{A1|...}}
TRAILED_TOKEN_RE = re.compile(r"\{\{[^}]+\}\}")             # 치환 후 남은 모든 {{...}}

# =========================
# 유틸
# =========================
def ensure_ext(name: str, ext: str) -> str:
    name = (name or "").strip()
    if not name:
        return DEFAULT_OUT_BASENAME + ext
    low = name.lower()
    if low.endswith(ext):
        return name
    # 다른 확장자가 붙어있으면 제거
    for e in (".docx", ".pdf", ".zip"):
        if low.endswith(e):
            name = name[: -len(e)]
            break
    return name + ext

def has_soffice() -> bool:
    paths = os.environ.get("PATH", "").split(os.pathsep)
    for p in paths:
        for cand in ("soffice", "soffice.bin"):
            if os.path.isfile(os.path.join(p, cand)):
                return True
    return False

def iter_block_items(container):
    """Document/Cell 공통으로 문단을 재귀 순회."""
    if isinstance(container, Document):
        parents: Iterable = [container]
    elif isinstance(container, _Cell):
        parents = [container]
    else:
        # header/footer/section/table 등 다양한 컨테이너도 paragraphs/tables 속성이 있으면 처리
        parents = [container]

    for parent in parents:
        # 문단
        if hasattr(parent, "paragraphs"):
            for p in parent.paragraphs:
                yield p
        # 표
        if hasattr(parent, "tables"):
            for t in parent.tables:
                for row in t.rows:
                    for cell in row.cells:
                        yield from iter_block_items(cell)

def replace_in_paragraph(par: Paragraph, repl_func):
    """run 단위 보존 → 안 되면 문단 전체 치환."""
    changed = False
    for run in par.runs:
        new_text = repl_func(run.text)
        if new_text != run.text:
            run.text = new_text
            changed = True
    if changed:
        return

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

def collect_remaining_tokens(doc: Document) -> List[str]:
    """치환 이후 문서 내 남은 {{...}} 토큰 수집."""
    found: Set[str] = set()

    def scan_text(s: str):
        for m in TRAILED_TOKEN_RE.finditer(s or ""):
            found.add(m.group(0))

    # 본문
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            scan_text("".join(run.text for run in item.runs))

    # 머리글/바닥글
    for section in doc.sections:
        for container in (section.header, section.footer):
            for item in iter_block_items(container):
                if isinstance(item, Paragraph):
                    scan_text("".join(run.text for run in item.runs))

    # 표 속성에 토큰이 들어갈 가능성은 낮지만 안전차원에서 표 텍스트도 훑기
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    scan_text("".join(run.text for run in p.runs))

    return sorted(found)

# =========================
# 포맷 도우미
# =========================
def try_format_as_date(v, date_fmt_py: str) -> str:
    """값을 날짜로 해석 가능하면 strftime(date_fmt_py) 적용."""
    try:
        if v is None:
            return ""
        if isinstance(v, (datetime, date)):
            dt = v if isinstance(v, datetime) else datetime(v.year, v.month, v.day)
            return dt.strftime(date_fmt_py)
        s = str(v).strip()
        # ISO yyyy-mm-dd or yyyy.mm.dd 형태도 허용
        for pat in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
            try:
                dt = datetime.strptime(s, pat)
                return dt.strftime(date_fmt_py)
            except Exception:
                pass
    except Exception:
        pass
    return ""

def normalize_date_mask(mask: str) -> str:
    """
    인라인: YYYY.MM.DD, YYYY년 MM월 DD일 등 → Python strftime.
    YYYY/YY/MM/DD 만 변환, 나머지는 리터럴 그대로 유지.
    """
    m = mask
    # 긴 것부터 치환해 겹침 방지
    m = m.replace("YYYY", "%Y").replace("yyyy", "%Y")
    m = m.replace("YY", "%y").replace("yy", "%y")
    m = m.replace("MM", "%m")
    m = m.replace("DD", "%d").replace("dd", "%d")
    return m

def format_number_with_locale(val: float, decimals: int, thousand_sep: str, decimal_sep: str) -> str:
    fmt = f"{{:,.{decimals}f}}".format(val)
    # 기본 , . → 로케일 적용
    if thousand_sep != "," or decimal_sep != ".":
        fmt = fmt.replace(",", "§").replace(".", "¤")
        fmt = fmt.replace("§", thousand_sep).replace("¤", decimal_sep)
    # 소수 0 자리면 정수 표기
    if decimals == 0 and decimal_sep in fmt:
        fmt = fmt.split(decimal_sep)[0]
    return fmt

def parse_number_mask(mask: str) -> int:
    """
    '#,###', '#,##0.00' 등에서 소수 자리 추출. 없으면 0.
    아주 단순 규칙: 마지막 '.' 뒤 자릿수 카운트.
    """
    if "." in mask:
        return max(0, len(mask.split(".")[-1].strip()))
    return 0

def value_to_text_auto(v, default_date_pyfmt: str, thousand_sep: str, decimal_sep: str) -> str:
    # 날짜 시도
    s = try_format_as_date(v, default_date_pyfmt)
    if s:
        return s
    # 숫자 시도
    try:
        if isinstance(v, (int, float, Decimal)):
            return format_number_with_locale(float(v), 0, thousand_sep, decimal_sep)
        if isinstance(v, str):
            raw = v.replace(",", "").replace(" ", "")
            if re.fullmatch(r"-?\d+(\.\d+)?", raw):
                return format_number_with_locale(float(raw), 0, thousand_sep, decimal_sep)
    except Exception:
        pass
    return "" if v is None else str(v)

# =========================
# 치환 콜백
# =========================
def make_replacer(ws, default_date_pyfmt: str, thousand_sep: str, decimal_sep: str):
    """
    {{A1|mask}} 지원. mask 예시:
      - "#,###" / "#,##0.00"  → 숫자 형식(소수 자리 자동 감지)
      - "YYYY.MM.DD" / "YYYY년 MM월 DD일" → 날짜 형식
    미지정시: 날짜/숫자 자동 판별
    """
    def _repl(text: str) -> str:
        def sub(m):
            addr = m.group(1)
            mask = (m.group(2) or "").lstrip("|").strip()

            # 엑셀 값 읽기
            try:
                v = ws[addr].value
            except Exception:
                v = None

            if mask:
                # 날짜 마스크로 보이면 변환
                if re.search(r"[Yy]|M|D|d", mask):
                    pyfmt = normalize_date_mask(mask)
                    s = try_format_as_date(v, pyfmt)
                    if s:
                        return s
                    return ""  # 날짜 마스크인데 날짜 아님 → 빈칸
                # 숫자 마스크
                decs = parse_number_mask(mask)
                try:
                    num = None
                    if isinstance(v, (int, float, Decimal)):
                        num = float(v)
                    elif isinstance(v, str):
                        raw = v.replace(",", "").replace(" ", "")
                        if re.fullmatch(r"-?\d+(\.\d+)?", raw):
                            num = float(raw)
                    if num is not None:
                        return format_number_with_locale(num, decs, thousand_sep, decimal_sep)
                except Exception:
                    pass
                # 마스크 주었지만 숫자 아님 → 원문 str
                return "" if v is None else str(v)

            # 기본(자동 판별)
            return value_to_text_auto(v, default_date_pyfmt, thousand_sep, decimal_sep)

        replaced = TOKEN_RE.sub(sub, text)

        # 추가: '오늘 날짜' 토큰 패치 (레이아웃 문서에서 가끔 사용하는 고정 문자열)
        sp4 = "    "
        today = datetime.today()
        today_str = today.strftime(default_date_pyfmt)
        for token in ("YYYY년 MM월 DD일", "YYYY년    MM월    DD일", "YYYY 년 MM 월 DD 일"):
            replaced = replaced.replace(token, today_str.replace(".", "년").replace("-", "년").replace("/", "년"))

        return replaced
    return _repl

# =========================
# DOCX → PDF
# =========================
def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> bytes | None:
    """
    1) Windows+Word: docx2pdf
    2) LibreOffice(soffice) headless
    실패 시 None (PDF 미동봉)
    """
    try:
        with tempfile.TemporaryDirectory() as td:
            in_path = os.path.join(td, "doc.docx")
            out_path = os.path.join(td, "doc.pdf")
            with open(in_path, "wb") as f:
                f.write(docx_bytes)

            # 1) Word
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

# =========================
# Streamlit UI
# =========================
st.title("납입요청서 자동 생성 (DOCX + PDF, 인라인 포맷/리포트/엄격 모드)")

with st.sidebar:
    st.subheader("포맷/옵션")
    # 로케일/포맷 설정
    thousand_sep = st.text_input("천단위 구분자", value=",", max_chars=2)
    decimal_sep = st.text_input("소수점 구분자", value=".", max_chars=2)
    default_date_mask = st.text_input("기본 날짜 포맷(예: YYYY.MM.DD)", value="YYYY.MM.DD")
    default_date_pyfmt = normalize_date_mask(default_date_mask)

    strict_mode = st.checkbox("엄격 모드(누락 토큰·시트 오류 시 중단)", value=False)
    show_remain = st.checkbox("치환 후 남은 토큰 리포트 표시", value=True)

xlsx_file = st.file_uploader("엑셀 파일(.xlsx, .xlsm)", type=["xlsx", "xlsm"])
docx_tpl = st.file_uploader("워드 템플릿(.docx)", type=["docx"])

# 시트 선택(기본은 후보 있으면 그걸로, 없으면 첫 시트)
selected_sheet = None
if xlsx_file:
    wb_probe = load_workbook(filename=io.BytesIO(xlsx_file.read()), data_only=True)
    sheet_names = wb_probe.sheetnames
    default_idx = 0
    if TARGET_SHEET_CANDIDATE in sheet_names:
        default_idx = sheet_names.index(TARGET_SHEET_CANDIDATE)
    selected_sheet = st.selectbox("시트 선택", sheet_names, index=default_idx)
    # 다시 사용하려면 파일 포인터 초기화 필요
    xlsx_file.seek(0)

out_base = st.text_input("출력 파일명(확장자 제외 가능)", value=DEFAULT_OUT_BASENAME)

run = st.button("문서 생성 (DOCX+PDF ZIP)")

if run:
    if not xlsx_file or not docx_tpl:
        st.error("엑셀/워드 템플릿을 모두 업로드하세요.")
        st.stop()

    try:
        # Excel
        wb = load_workbook(filename=io.BytesIO(xlsx_file.read()), data_only=True)
        xlsx_file.seek(0)
        if selected_sheet and selected_sheet in wb.sheetnames:
            ws = wb[selected_sheet]
        else:
            if strict_mode:
                st.error(f"엄격 모드: 선택된 시트를 찾을 수 없습니다.")
                st.stop()
            ws = wb[wb.sheetnames[0]]

        # Word
        tpl_bytes = docx_tpl.read()
        doc = Document(io.BytesIO(tpl_bytes))

        # 치환
        replacer = make_replacer(ws, default_date_pyfmt, thousand_sep or ",", decimal_sep or ".")
        replace_everywhere(doc, replacer)

        # 누락 토큰 리포트
        remaining = collect_remaining_tokens(doc)
        if show_remain:
            with st.sidebar:
                st.subheader("남은 토큰")
                if remaining:
                    st.warning(f"{len(remaining)}개 남음")
                    st.code("\n".join(remaining), language="text")
                else:
                    st.success("모든 토큰이 치환되었습니다.")

        if strict_mode and remaining:
            st.error("엄격 모드: 템플릿 내 미치환 토큰이 남아 종료합니다.")
            st.stop()

        # DOCX bytes
        docx_buf = io.BytesIO()
        doc.save(docx_buf)
        docx_buf.seek(0)
        docx_bytes = docx_buf.getvalue()

        # PDF bytes
        pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)
        pdf_ready = pdf_bytes is not None

        # ZIP (둘 다 한 번에 다운로드)
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            zf.writestr(ensure_ext(out_base, ".docx"), docx_bytes)
            if pdf_ready:
                zf.writestr(ensure_ext(out_base, ".pdf"), pdf_bytes)
        zip_buf.seek(0)

        st.success("생성이 완료되었습니다.")
        st.download_button(
            "WORD+PDF 한번에 다운로드 (ZIP)",
            data=zip_buf,
            file_name=ensure_ext(out_base, ".zip"),
            mime="application/zip",
        )

        # 선택: 개별로도 제공
        colA, colB = st.columns(2)
        with colA:
            st.download_button(
                "DOCX만 다운로드",
                data=docx_bytes,
                file_name=ensure_ext(out_base, ".docx"),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        with colB:
            st.download_button(
                "PDF만 다운로드",
                data=pdf_bytes if pdf_ready else b"",
                file_name=ensure_ext(out_base, ".pdf"),
                mime="application/pdf",
                disabled=not pdf_ready,
                help=None if pdf_ready else "PDF 변환 불가(Word 또는 LibreOffice 필요)",
            )

    except Exception as e:
        st.exception(e)
