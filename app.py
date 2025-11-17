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

# 선택: docx2pdf (없으면 WORD만 생성)
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
    name = (name or "").strip()
    return name if name.lower().endswith(".docx") else (name + ".docx")


def ensure_pdf(name: str) -> str:
    base = (name or "output").strip()
    return base if base.lower().endswith(".pdf") else (base + ".pdf")


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


# ================== 값 포맷팅 ================== #

def try_format_as_date(v) -> str:
    try:
        if isinstance(v, (datetime, date)):
            return f"{v.year}. {v.month}. {v.day}."
        if isinstance(v, str):
            s = v.strip()
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


def apply_inline_format(value, fmt: Optional[str]) -> str:
    """{{A1|FORMAT}} 에서 FORMAT에 따라 포맷."""
    if fmt is None or fmt.strip() == "":
        return value_to_text(value)

    # 날짜 포맷 (YYYY/MM/DD 등)
    if any(tok in fmt for tok in ("YYYY", "MM", "DD")):
        if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value.strip()):
            value = datetime.strptime(value.strip(), "%Y-%m-%d").date()
        if isinstance(value, (datetime, date)):
            f = (fmt
                 .replace("YYYY", "%Y")
                 .replace("MM", "%m")
                 .replace("DD", "%d"))
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


# ================== DOCX 치환 ================== #

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
    def _repl(text: str) -> str:
        def sub(m):
            addr, fmt = m.group(1), m.group(2)
            try:
                v = ws[addr].value
            except Exception:
                v = None
            return apply_inline_format(v, fmt)

        replaced = TOKEN_RE.sub(sub, text)

        # YYYY년 MM월 DD일 → 오늘 날짜
        today = datetime.today()
        today_str = f"{today.year}년 {today.month}월 {today.day}일"
        for token in ["YYYY년 MM월 DD일", "YYYY 년 MM 월 DD 일"]:
            replaced = replaced.replace(token, today_str)

        return replaced

    return _repl


def collect_leftover_tokens(doc: Document) -> Set[str]:
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
    if not data or len(data) == 0:
        raise InvalidFileException("엑셀 파일이 비어 있습니다 (0 bytes).")

    try:
        return load_workbook(filename=io.BytesIO(data), data_only=True)
    except BadZipFile:
        raise InvalidFileException("엑셀 파일이 손상되었거나 XLS 형식일 수 있습니다.")
    except Exception as e:
        raise InvalidFileException(f"엑셀 파일 로드 오류: {e}")


def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> Optional[bytes]:
    """WORD 환경(docx2pdf) 또는 LibreOffice가 있을 때만 PDF 생성."""
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


# ================== Streamlit UI ================== #

def init_session_state():
    if "xlsx_data" not in st.session_state:
        st.session_state.xlsx_data = None
    if "xlsx_name" not in st.session_state:
        st.session_state.xlsx_name = None
    if "docx_data" not in st.session_state:
        st.session_state.docx_data = None
    if "docx_name" not in st.session_state:
        st.session_state.docx_name = None


def render_inputs():
    """엑셀/워드 업로드 + 시트 선택 + 버튼."""
    h4("엑셀 파일")

    xlsx_file = st.file_uploader(
        "엑셀 업로드",
        type=["xlsx", "xlsm"],
        key="xlsx_normal",
    )

    if xlsx_file is not None:
        try:
            xlsx_bytes = xlsx_file.getvalue()
            if len(xlsx_bytes) > 0:
                st.session_state.xlsx_data = xlsx_bytes
                st.session_state.xlsx_name = xlsx_file.name
                st.success(f"{xlsx_file.name}: {len(xlsx_bytes):,} bytes")
            else:
                st.error("업로드된 엑셀 파일이 0 bytes입니다.")
        except Exception as e:
            st.error(f"엑셀 파일 읽기 오류: {e}")

    st.markdown("---")

    h4("워드 템플릿(.docx)")

    docx_tpl = st.file_uploader(
        "템플릿 업로드",
        type=["docx"],
        key="docx_normal",
    )

    if docx_tpl is not None:
        try:
            docx_bytes = docx_tpl.getvalue()
            if len(docx_bytes) > 0:
                st.session_state.docx_data = docx_bytes
                st.session_state.docx_name = docx_tpl.name
                st.success(f"{docx_tpl.name}: {len(docx_bytes):,} bytes")
            else:
                st.error("업로드된 워드 템플릿이 0 bytes입니다.")
        except Exception as e:
            st.error(f"워드 파일 읽기 오류: {e}")

    st.markdown("---")

    # 시트 선택
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

    return sheet_choice, out_name, gen


def handle_generate(sheet_choice: Optional[str], out_name: str):
    if not st.session_state.xlsx_data or not st.session_state.docx_data:
        st.error("엑셀과 템플릿을 모두 로드하세요.")
        return

    try:
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

        doc = Document(io.BytesIO(st.session_state.docx_data))

        replacer = make_replacer(ws)
        replace_everywhere(doc, replacer)

        docx_buf = io.BytesIO()
        doc.save(docx_buf)
        docx_buf.seek(0)
        docx_bytes = docx_buf.getvalue()

        pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)
        pdf_ok = pdf_bytes is not None

        # 남은 토큰 안내 (디버그용)
        doc_after = Document(io.BytesIO(docx_bytes))
        leftovers = sorted(list(collect_leftover_tokens(doc_after)))
        if leftovers:
            with st.expander("남아 있는 토큰 목록"):
                st.code("\n".join(leftovers))
        else:
            small_note("모든 토큰이 정상적으로 치환되었습니다.")

    except InvalidFileException as e:
        st.error(str(e))
        return
    except Exception as e:
        st.exception(e)
        return

    st.success("문서가 준비되었습니다.")
    render_download_buttons(docx_bytes, pdf_bytes, pdf_ok, out_name)


def render_download_buttons(
    docx_bytes: bytes,
    pdf_bytes: Optional[bytes],
    pdf_ok: bool,
    out_name: str,
):
    col1, col2, col3 = st.columns(3)

    with col1:
        st.download_button(
            "WORD 다운로드",
            data=docx_bytes,
            file_name=ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    with col2:
        st.download_button(
            "PDF 다운로드",
            data=(pdf_bytes or b""),
            file_name=ensure_pdf(out_name),
            mime="application/pdf",
            disabled=not pdf_ok,
            help=None
            if pdf_ok
            else "PDF 변환 엔진(Word 또는 LibreOffice)이 없는 환경입니다.",
            use_container_width=True,
        )

    with col3:
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            zf.writestr(
                ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT,
                docx_bytes,
            )
            if pdf_ok and pdf_bytes:
                zf.writestr(ensure_pdf(out_name), pdf_bytes)

        zip_buf.seek(0)
        st.download_button(
            "ZIP (WORD+PDF)",
            data=zip_buf,
            file_name=ensure_pdf(out_name).replace(".pdf", "") + "_both.zip",
            use_container_width=True,
        )


# ================== 엔트리 포인트 ================== #

def main():
    inject_style()
    init_session_state()

    st.title("🧾 납입요청서 자동 생성 (DOCX + PDF)")

    sheet_choice, out_name, gen = render_inputs()

    if gen:
        handle_generate(sheet_choice, out_name)


if __name__ == "__main__":
    main()
