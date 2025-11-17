import base64
import io
import os
import re
import subprocess
import tempfile
from datetime import date, datetime
from decimal import Decimal
from typing import Optional, Set
from zipfile import BadZipFile, ZipFile, ZIP_DEFLATED

import streamlit as st
from docx import Document
from docx.table import _Cell
from docx.text.paragraph import Paragraph
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from ui_style import inject as inject_style, h4, small_note

# 선택: docx2pdf
try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

# ================== 상수 & 정규식 ================== #

TOKEN_RE = re.compile(r"\{\{([A-Z]+[0-9]+)(?:\|([^}]+))?\}\}")
LEFTOVER_RE = re.compile(r"\{\{[^}]+\}\}")

DEFAULT_OUT = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행.docx"
TARGET_SHEET = "2. 배정후 청약시"


# ================== 파일명 유틸 ================== #

def ensure_docx(name: str) -> str:
    """파일명에 .docx 확장자를 보장."""
    name = (name or "").strip()
    return name if name.lower().endswith(".docx") else (name + ".docx")


def ensure_pdf(name: str) -> str:
    """파일명에 .pdf 확장자를 보장."""
    base = (name or "output").strip()
    return base if base.lower().endswith(".pdf") else (base + ".pdf")


def has_soffice() -> bool:
    """LibreOffice(soffice) 사용 가능 여부 확인."""
    try:
        subprocess.run(
            ["soffice", "--version"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        return True
    except FileNotFoundError:
        return False


# ================== 값 포맷팅 유틸 ================== #

def try_format_as_date(v) -> str:
    """value를 'YYYY. M. D.' 형식의 문자열로 포맷 (가능한 경우만)."""
    try:
        if isinstance(v, (datetime, date)):
            return f"{v.year}. {v.month}. {v.day}."
        if isinstance(v, str):
            s = v.strip()
            # 2024-01-01 형식만 간단 처리
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
                dt = datetime.strptime(s, "%Y-%m-%d").date()
                return f"{dt.year}. {dt.month}. {dt.day}."
    except Exception:
        pass
    return ""


def fmt_number(v) -> str:
    """숫자형 값을 천단위 콤마 문자열로 포맷."""
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
    """셀 값을 날짜/숫자 우선 포맷 후 문자열로 변환."""
    s = try_format_as_date(v)
    if s:
        return s
    s = fmt_number(v)
    if s:
        return s
    return "" if v is None else str(v)


def apply_inline_format(value, fmt: Optional[str]) -> str:
    """
    {{A1|FORMAT}} 에서 FORMAT에 따라 value 포맷팅.
    - 날짜 포맷: YYYY/MM/DD 등
    - 숫자 포맷: #,###.00 등
    """
    if fmt is None or fmt.strip() == "":
        return value_to_text(value)

    # 날짜 포맷 처리
    if any(tok in fmt for tok in ("YYYY", "MM", "DD")):
        if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value.strip()):
            value = datetime.strptime(value.strip(), "%Y-%m-%d").date()
        if isinstance(value, (datetime, date)):
            f = (
                fmt.replace("YYYY", "%Y")
                .replace("MM", "%m")
                .replace("DD", "%d")
            )
            return value.strftime(f)
        return value_to_text(value)

    # 숫자 포맷 처리 (#,###.00 등)
    if re.fullmatch(r"[#,0]+(?:\.[0#]+)?", fmt.replace(",", "")):
        try:
            num = float(str(value).replace(",", ""))
            decimals = len(fmt.split(".")[1]) if "." in fmt else 0
            return f"{num:,.{decimals}f}"
        except Exception:
            return value_to_text(value)

    return value_to_text(value)


# ================== DOCX 치환 유틸 ================== #

def replace_in_paragraph(paragraph: Paragraph, repl_func):
    """문단 텍스트의 {{A1}} 토큰 치환."""
    if not paragraph.text:
        return

    new_text = repl_func(paragraph.text)
    if new_text == paragraph.text:
        return

    # run 구조는 무시하고 전체 텍스트 교체
    for run in paragraph.runs:
        run.text = ""
    if paragraph.runs:
        paragraph.runs[0].text = new_text
    else:
        paragraph.add_run(new_text)


def replace_in_table(cell: _Cell, repl_func):
    """테이블 셀 내부 문단/중첩 테이블 치환."""
    for p in cell.paragraphs:
        replace_in_paragraph(p, repl_func)
    for t in cell.tables:
        for row in t.rows:
            for c in row.cells:
                replace_in_table(c, repl_func)


def iter_block_items(parent):
    """문서/헤더/푸터/셀 안의 단락과 셀을 순회."""
    if hasattr(parent, "paragraphs") and hasattr(parent, "tables"):
        for p in parent.paragraphs:
            yield p
        for t in parent.tables:
            for row in t.rows:
                for cell in row.cells:
                    for item in iter_block_items(cell):
                        yield item


def replace_everywhere(doc: Document, repl_func):
    """본문 + 헤더/푸터 전체에 대해 토큰 치환."""
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


def make_replacer(ws):
    """엑셀 워크시트 기반 치환 함수 생성."""

    def _repl(text: str) -> str:
        def sub(m):
            addr, fmt = m.group(1), m.group(2)
            try:
                v = ws[addr].value
            except Exception:
                v = None
            return apply_inline_format(v, fmt)

        replaced = TOKEN_RE.sub(sub, text)

        # 간이 날짜 더미 치환 (YYYY년 MM월 DD일 → 오늘 날짜)
        today = datetime.today()
        today_str = f"{today.year}년 {today.month}월 {today.day}일"
        for token in [
            "YYYY년 MM월 DD일",
            "YYYY 년 MM 월 DD 일",
        ]:
            replaced = replaced.replace(token, today_str)

        return replaced

    return _repl


def collect_leftover_tokens(doc: Document) -> Set[str]:
    """치환 후에도 남아 있는 {{...}} 토큰 수집."""
    leftovers: Set[str] = set()

    def _scan(parent):
        for item in iter_block_items(parent):
            if isinstance(item, Paragraph) and item.text:
                for m in LEFTOVER_RE.findall(item.text):
                    leftovers.add(m)

    _scan(doc)

    for section in doc.sections:
        for container in (section.header, section.footer):
            _scan(container)

    return leftovers


# ================== 엑셀/워드 로드 & 변환 ================== #

def load_workbook_from_bytes(data: bytes, filename: str = "file.xlsx") -> Workbook:
    """바이트 데이터에서 워크북 로드."""
    if not data or len(data) == 0:
        raise InvalidFileException(
            f"파일이 비어있습니다 (0 bytes)\n"
            f"파일명: {filename}\n\n"
            f"해결 방법:\n"
            f"1. 파일이 실제로 손상되었을 수 있습니다\n"
            f"2. 엑셀에서 파일을 열어 '다른 이름으로 저장'하세요\n"
            f"3. 파일명을 영문으로 변경해보세요 (예: data.xlsx)"
        )

    try:
        return load_workbook(filename=io.BytesIO(data), data_only=True)
    except BadZipFile:
        raise InvalidFileException(
            "엑셀 파일이 손상되었거나 실제로는 XLS 형식일 수 있습니다.\n"
            "엑셀에서 '다른 이름으로 저장 > Excel 통합 문서 (*.xlsx)'로 저장하세요."
        )
    except Exception as e:
        raise InvalidFileException(f"엑셀 파일 로드 오류: {e}")


def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> Optional[bytes]:
    """DOCX 바이트를 PDF 바이트로 변환(MS Word 또는 LibreOffice 필요)."""
    try:
        with tempfile.TemporaryDirectory() as td:
            in_path = os.path.join(td, "doc.docx")
            out_path = os.path.join(td, "doc.pdf")

            with open(in_path, "wb") as f:
                f.write(docx_bytes)

            # 1) docx2pdf (Windows/Office 환경)
            if docx2pdf_convert is not None:
                try:
                    docx2pdf_convert(in_path, out_path)
                    if os.path.exists(out_path):
                        with open(out_path, "rb") as f:
                            return f.read()
                except Exception:
                    pass

            # 2) LibreOffice(soffice) 사용
            if has_soffice():
                try:
                    subprocess.run(
                        [
                            "soffice",
                            "--headless",
                            "--convert-to",
                            "pdf",
                            in_path,
                            "--outdir",
                            td,
                        ],
                        check=True,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                    )
                    if os.path.exists(out_path):
                        with open(out_path, "rb") as f:
                            return f.read()
                except Exception:
                    pass
    except Exception:
        pass

    return None


# ================== Streamlit UI ================== #

def init_session_state():
    """세션 상태 초기화."""
    if "xlsx_data" not in st.session_state:
        st.session_state.xlsx_data = None
    if "xlsx_name" not in st.session_state:
        st.session_state.xlsx_name = None
    if "docx_data" not in st.session_state:
        st.session_state.docx_data = None
    if "docx_name" not in st.session_state:
        st.session_state.docx_name = None

        # Base64 업로드
        with st.expander("📋 또는 Base64로 붙여넣기 (방법 2)", expanded=False):
            st.markdown(
                """
                **파일 업로드가 안될 때 사용하세요:**
                1. 터미널/명령 프롬프트에서 실행:
                ```bash
                # Windows (PowerShell)
                [Convert]::ToBase64String([IO.File]::ReadAllBytes("파일경로.xlsx"))
               
                # Mac/Linux
                base64 파일경로.xlsx
                ```
                2. 출력된 텍스트를 복사해서 아래 박스에 붙여넣기
                """
            )
            xlsx_base64 = st.text_area(
                "Base64 텍스트",
                height=100,
                placeholder="여기에 Base64 인코딩된 엑셀 파일을 붙여넣으세요...",
                key="xlsx_base64",
            )
            xlsx_fname = st.text_input("파일명", value="data.xlsx", key="xlsx_fname")

            if st.button("Base64에서 로드", key="load_xlsx_base64"):
                try:
                    xlsx_bytes = base64.b64decode(xlsx_base64.strip())
                    st.session_state.xlsx_data = xlsx_bytes
                    st.session_state.xlsx_name = xlsx_fname
                    st.success(f"✅ 엑셀 파일 로드 완료: {len(xlsx_bytes):,} bytes")
                except Exception as e:
                    st.error(f"Base64 디코딩 실패: {e}")

        # 일반 업로드 처리
        if xlsx_file is not None:
            try:
                xlsx_bytes = xlsx_file.getvalue()
                if len(xlsx_bytes) > 0:
                    st.session_state.xlsx_data = xlsx_bytes
                    st.session_state.xlsx_name = xlsx_file.name
                    st.success(f"✅ {xlsx_file.name}: {len(xlsx_bytes):,} bytes")
                else:
                    st.error("⚠️ 업로드된 파일이 0 bytes입니다. 방법 2를 사용해보세요.")
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")

        st.markdown("---")

        # ===== 시트 선택 =====
        sheet_choice = None
        if st.session_state.xlsx_data:
            try:
                wb_tmp = load_workbook_from_bytes(
                    st.session_state.xlsx_data, st.session_state.xlsx_name
                )
                default_idx = (
                    wb_tmp.sheetnames.index(TARGET_SHEET)
                    if TARGET_SHEET in wb_tmp.sheetnames
                    else 0
                )
                sheet_choice = st.selectbox(
                    "사용할 시트",
                    wb_tmp.sheetnames,
                    index=default_idx,
                    key="sheet_choice",
                )
            except Exception as e:
                st.error(f"엑셀 미리보기 오류: {e}")

        out_name = st.text_input("출력 파일명", value=DEFAULT_OUT)
        gen = st.button("문서 생성", use_container_width=True, type="primary")

    # 오른쪽 컬럼은 따로 렌더링
    render_right_column()

    return sheet_choice, out_name, gen

def handle_generate(sheet_choice: Optional[str], out_name: str):
    """문서 생성 버튼 클릭 시 실행 로직."""
    if not st.session_state.xlsx_data or not st.session_state.docx_data:
        st.error("엑셀과 템플릿을 모두 로드하세요.")
        st.stop()

    with st.status("문서 생성 중...", expanded=True) as status:
        try:
            st.write("1) 엑셀 로드")
            wb = load_workbook_from_bytes(
                st.session_state.xlsx_data, st.session_state.xlsx_name
            )
            ws = (
                wb[sheet_choice]
                if sheet_choice
                else (
                    wb[TARGET_SHEET]
                    if TARGET_SHEET in wb.sheetnames
                    else wb[wb.sheetnames[0]]
                )
            )

            st.write("2) 템플릿 로드")
            doc = Document(io.BytesIO(st.session_state.docx_data))

            st.write("3) 치환 실행")
            replacer = make_replacer(ws)
            replace_everywhere(doc, replacer)

            st.write("4) WORD 저장")
            docx_buf = io.BytesIO()
            doc.save(docx_buf)
            docx_buf.seek(0)
            docx_bytes = docx_buf.getvalue()

            st.write("5) PDF 변환 시도")
            pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)
            pdf_ok = pdf_bytes is not None

            st.write("6) 남은 토큰 확인")
            doc_after = Document(io.BytesIO(docx_bytes))
            leftovers = sorted(list(collect_leftover_tokens(doc_after)))
            if leftovers:
                with st.expander("남아 있는 토큰 목록"):
                    st.code("\n".join(leftovers))
            else:
                small_note("모든 토큰이 정상적으로 치환되었습니다.")

            status.update(label="완료", state="complete", expanded=False)

        except InvalidFileException as e:
            status.update(label="엑셀 형식 오류", state="error", expanded=True)
            st.error(str(e))
            st.stop()
        except Exception as e:
            status.update(label="오류", state="error", expanded=True)
            st.exception(e)
            st.stop()

    st.success("문서가 준비되었습니다.")
    render_download_buttons(docx_bytes, pdf_bytes, pdf_ok, out_name)


def render_download_buttons(docx_bytes: bytes, pdf_bytes: Optional[bytes],
                            pdf_ok: bool, out_name: str):
    """WORD / PDF / ZIP 다운로드 버튼 렌더링."""
    dl_cols = st.columns(3)

    # WORD
    with dl_cols[0]:
        st.download_button(
            "📄 WORD 다운로드",
            data=docx_bytes,
            file_name=ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    # PDF
    with dl_cols[1]:
        st.download_button(
            "🖨 PDF 다운로드",
            data=(pdf_bytes or b""),
            file_name=ensure_pdf(out_name),
            mime="application/pdf",
            disabled=not pdf_ok,
            help=None
            if pdf_ok
            else "PDF 변환 엔진(Word 또는 LibreOffice)이 없는 환경입니다.",
            use_container_width=True,
        )

    # ZIP (WORD + PDF)
    with dl_cols[2]:
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            # WORD
            zf.writestr(
                ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT,
                docx_bytes,
            )
            # PDF (가능한 경우에만)
            if pdf_ok and pdf_bytes:
                zf.writestr(ensure_pdf(out_name), pdf_bytes)

        zip_buf.seek(0)
        st.download_button(
            "📦 ZIP (WORD+PDF)",
            data=zip_buf,
            file_name=ensure_pdf(out_name).replace(".pdf", "") + "_both.zip",
            use_container_width=True,
        )


# ================== 엔트리 포인트 ================== #

def main():
    inject_style()
    init_session_state()

    st.title("🧾 납입요청서 자동 생성 (DOCX + PDF)")

    sheet_choice, out_name, gen = render_left_column()

    if gen:
        handle_generate(sheet_choice, out_name)


if __name__ == "__main__":
    main()
