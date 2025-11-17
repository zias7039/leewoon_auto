import io
import os
import re
import subprocess
import tempfile
from datetime import date, datetime
from decimal import Decimal
from typing import Optional
from zipfile import BadZipFile, ZipFile, ZIP_DEFLATED

import streamlit as st
from docx import Document
from docx.table import _Cell
from docx.text.paragraph import Paragraph
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from ui_style import inject as inject_style, h4, section_caption, small_note

# docx → pdf (환경에 없으면 PDF는 ZIP에 안 넣음)
try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

TOKEN_RE = re.compile(r"\{\{([A-Z]+[0-9]+)(?:\|([^}]+))?\}\}")
DEFAULT_OUT = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행.docx"
TARGET_SHEET = "2. 배정후 청약시"


# ---------- 유틸 ----------

def ensure_docx(name: str) -> str:
    name = (name or "").strip()
    return name if name.lower().endswith(".docx") else name + ".docx"


def ensure_pdf(name: str) -> str:
    base = (name or "output").strip()
    return base if base.lower().endswith(".pdf") else base + ".pdf"


def has_soffice() -> bool:
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


def try_format_as_date(v) -> str:
    try:
        if isinstance(v, (datetime, date)):
            return f"{v.year}. {v.month}. {v.day}."
        if isinstance(v, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", v.strip()):
            dt = datetime.strptime(v.strip(), "%Y-%m-%d").date()
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


def apply_inline_format(value, fmt: Optional[str]) -> str:
    if not fmt or not fmt.strip():
        return value_to_text(value)

    # 날짜 포맷
    if any(tok in fmt for tok in ("YYYY", "MM", "DD")):
        if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value.strip()):
            value = datetime.strptime(value.strip(), "%Y-%m-%d").date()
        if isinstance(value, (datetime, date)):
            f = fmt.replace("YYYY", "%Y").replace("MM", "%m").replace("DD", "%d")
            return value.strftime(f)
        return value_to_text(value)

    # 숫자 포맷 (#,###.00 등)
    if re.fullmatch(r"[#,0]+(?:\.[0#]+)?", fmt.replace(",", "")):
        try:
            num = float(str(value).replace(",", ""))
            decimals = len(fmt.split(".")[1]) if "." in fmt else 0
            return f"{num:,.{decimals}f}"
        except Exception:
            return value_to_text(value)

    return value_to_text(value)


# ---------- DOCX 치환 ----------

def replace_in_paragraph(paragraph: Paragraph, repl_func):
    if not paragraph.text:
        return
    new_text = repl_func(paragraph.text)
    if new_text == paragraph.text:
        return
    for run in paragraph.runs:
        run.text = ""
    if paragraph.runs:
        paragraph.runs[0].text = new_text
    else:
        paragraph.add_run(new_text)


def replace_in_table(cell: _Cell, repl_func):
    for p in cell.paragraphs:
        replace_in_paragraph(p, repl_func)
    for t in cell.tables:
        for row in t.rows:
            for c in row.cells:
                replace_in_table(c, repl_func)


def iter_block_items(parent):
    if hasattr(parent, "paragraphs") and hasattr(parent, "tables"):
        for p in parent.paragraphs:
            yield p
        for t in parent.tables:
            for row in t.rows:
                for cell in row.cells:
                    for item in iter_block_items(cell):
                        yield item


def replace_everywhere(doc: Document, repl_func):
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            replace_in_paragraph(item, repl_func)
    for section in doc.sections:
        for container in (section.header, section.footer):
            for item in iter_block_items(container):
                if isinstance(item, Paragraph):
                    replace_in_paragraph(item, repl_func)


def make_replacer(ws):
    def _repl(text: str) -> str:
        def sub(m):
            addr, fmt = m.group(1), m.group(2)
            try:
                v = ws[addr].value
            except Exception:
                v = None
            return apply_inline_format(v, fmt)

        replaced = TOKEN_RE.sub(sub, text)

        today = datetime.today()
        today_str = f"{today.year}년 {today.month}월 {today.day}일"
        for token in ("YYYY년 MM월 DD일", "YYYY 년 MM 월 DD 일"):
            replaced = replaced.replace(token, today_str)

        return replaced

    return _repl


# ---------- 파일 로드 & 변환 ----------

def load_workbook_from_bytes(data: bytes, filename: str = "file.xlsx") -> Workbook:
    if not data:
        raise InvalidFileException("엑셀 파일이 비어 있습니다 (0 bytes).")
    try:
        return load_workbook(filename=io.BytesIO(data), data_only=True)
    except BadZipFile:
        raise InvalidFileException("엑셀 파일이 손상되었거나 XLS 형식일 수 있습니다.")
    except Exception as e:
        raise InvalidFileException(f"엑셀 파일 로드 오류: {e}")


def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> Optional[bytes]:
    try:
        with tempfile.TemporaryDirectory() as td:
            in_path = os.path.join(td, "doc.docx")
            out_path = os.path.join(td, "doc.pdf")
            with open(in_path, "wb") as f:
                f.write(docx_bytes)

            # 1) MS Word (docx2pdf)
            if docx2pdf_convert is not None:
                try:
                    docx2pdf_convert(in_path, out_path)
                    if os.path.exists(out_path):
                        with open(out_path, "rb") as f:
                            return f.read()
                except Exception:
                    pass

            # 2) LibreOffice
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


# ---------- Streamlit UI ----------

def init_session_state():
    for key in ("xlsx_data", "xlsx_name", "docx_data", "docx_name"):
        if key not in st.session_state:
            st.session_state[key] = None


def render_top_bar() -> bool:
    """상단 고정 ZIP 생성 버튼."""
    st.markdown('<div class="top-bar"><div class="top-bar-inner">', unsafe_allow_html=True)
    col1, col2 = st.columns([3, 1])
    with col1:
        st.markdown(
            '<div class="top-bar-title">납입요청서 일괄 생성 · ZIP 다운로드</div>',
            unsafe_allow_html=True,
        )
    with col2:
        gen_top = st.button("ZIP 생성", key="btn_top", use_container_width=True)
    st.markdown("</div></div>", unsafe_allow_html=True)
    return gen_top


def render_inputs():
    """2열 레이아웃 입력 카드."""
    st.markdown('<div class="app-card">', unsafe_allow_html=True)

    col_left, col_right = st.columns(2)

    # 왼쪽: 엑셀 업로드
    with col_left:
        h4("엑셀 파일")
        section_caption("청약/납입 데이터가 들어있는 엑셀 파일")
        xlsx_file = st.file_uploader("엑셀 업로드", type=["xlsx", "xlsm"], key="xlsx")
        if xlsx_file is not None:
            try:
                data = xlsx_file.getvalue()
                if data:
                    st.session_state.xlsx_data = data
                    st.session_state.xlsx_name = xlsx_file.name
                    st.success(f"{xlsx_file.name}: {len(data):,} bytes")
                else:
                    st.error("엑셀 파일이 0 bytes입니다.")
            except Exception as e:
                st.error(f"엑셀 파일 읽기 오류: {e}")

    # 오른쪽: 워드 업로드
    with col_right:
        h4("워드 템플릿 (.docx)")
        section_caption("{{A1}}, {{B5|#,###}}, {{C3|YYYY.MM.DD}} 태그가 포함된 템플릿")
        docx_file = st.file_uploader("워드 템플릿 업로드", type=["docx"], key="docx")
        if docx_file is not None:
            try:
                data = docx_file.getvalue()
                if data:
                    st.session_state.docx_data = data
                    st.session_state.docx_name = docx_file.name
                    st.success(f"{docx_file.name}: {len(data):,} bytes")
                else:
                    st.error("워드 템플릿이 0 bytes입니다.")
            except Exception as e:
                st.error(f"워드 파일 읽기 오류: {e}")

    st.markdown("---")

    # 시트 선택 + 출력 파일명 + 하단 ZIP 버튼
    sheet_choice = None
    if st.session_state.xlsx_data:
        try:
            wb = load_workbook_from_bytes(
                st.session_state.xlsx_data, st.session_state.xlsx_name
            )
            sheets = wb.sheetnames
            index = sheets.index(TARGET_SHEET) if TARGET_SHEET in sheets else 0
            h4("사용할 시트")
            sheet_choice = st.selectbox("시트 선택", sheets, index=index)
        except Exception as e:
            st.error(f"엑셀 시트 읽기 오류: {e}")

    h4("출력 파일명")
    out_name = st.text_input("파일명", value=DEFAULT_OUT)

    gen_bottom = st.button("ZIP 생성", key="btn_bottom", use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

    return sheet_choice, out_name, gen_bottom


def handle_generate(sheet_choice: Optional[str], out_name: str):
    if not st.session_state.xlsx_data or not st.session_state.docx_data:
        st.error("엑셀과 워드 템플릿을 모두 업로드하세요.")
        return

    progress = st.progress(0)
    try:
        with st.spinner("ZIP 생성 중입니다..."):
            # 1) 엑셀 로드
            progress.progress(10)
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

            # 2) 워드 템플릿 로드
            progress.progress(35)
            doc = Document(io.BytesIO(st.session_state.docx_data))

            # 3) 치환
            replacer = make_replacer(ws)
            replace_everywhere(doc, replacer)
            progress.progress(60)

            # 4) DOCX 저장
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            docx_bytes = buf.getvalue()
            progress.progress(75)

            # 5) PDF 변환 (가능한 경우)
            pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)
            pdf_ok = pdf_bytes is not None
            progress.progress(90)

        progress.progress(100)
    except InvalidFileException as e:
        st.error(str(e))
        return
    except Exception as e:
        st.exception(e)
        return

    st.success("ZIP 파일이 준비되었습니다.")
    render_zip_download(docx_bytes, pdf_bytes, pdf_ok, out_name)


def render_zip_download(
    docx_bytes: bytes,
    pdf_bytes: Optional[bytes],
    pdf_ok: bool,
    out_name: str,
):
    zip_buf = io.BytesIO()
    with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
        docx_name = ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT
        zf.writestr(docx_name, docx_bytes)

        if pdf_ok and pdf_bytes:
            pdf_name = ensure_pdf(out_name)
            zf.writestr(pdf_name, pdf_bytes)

    zip_buf.seek(0)

    base_zip_name = (ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT)
    base_zip_name = base_zip_name.replace(".docx", "")
    zip_name = f"{base_zip_name}_both.zip"

    st.download_button(
        "ZIP 다운로드 (WORD + PDF)",
        data=zip_buf,
        file_name=zip_name,
        use_container_width=True,
    )


def main():
    inject_style()
    init_session_state()

    st.title("납입요청서 자동 생성")
    st.markdown(
        '<div class="app-subtitle">엑셀 데이터와 워드 템플릿을 결합해 납입요청서 DOCX/PDF를 만들고, ZIP으로 일괄 내려받는 도구입니다.</div>',
        unsafe_allow_html=True,
    )

    gen_top = render_top_bar()  # 상단 고정 ZIP 버튼
    sheet_choice, out_name, gen_bottom = render_inputs()  # 2열 레이아웃 입력

    generate = gen_top or gen_bottom
    if generate:
        handle_generate(sheet_choice, out_name)


if __name__ == "__main__":
    main()
