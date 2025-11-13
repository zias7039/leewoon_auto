import io, os, re, tempfile, subprocess
from datetime import datetime, date
from decimal import Decimal
from zipfile import ZipFile, ZIP_DEFLATED, BadZipFile

import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.utils.exceptions import InvalidFileException
from docx import Document
from docx.table import _Cell
from docx.text.paragraph import Paragraph

# 스타일
from ui_style import inject as inject_style, h4, small_note

# 선택: docx2pdf
try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

# -------- 치환 유틸 --------
TOKEN_RE = re.compile(r"\{\{([A-Z]+[0-9]+)(?:\|([^}]+))?\}\}")
LEFTOVER_RE = re.compile(r"\{\{[^}]+\}\}")
DEFAULT_OUT = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행.docx"
TARGET_SHEET = "2.  배정후 청약시"

def ensure_docx(name: str) -> str:
    name = (name or "").strip()
    return name if name.lower().endswith(".docx") else (name + ".docx")

def ensure_pdf(name: str) -> str:
    base = (name or "output").strip()
@@ -127,50 +128,65 @@ def replace_everywhere(doc: Document, repl_func):
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
            try: v = ws[addr].value
            except Exception: v = None
            return apply_inline_format(v, fmt)
        replaced = TOKEN_RE.sub(sub, text)
        # 간이 날짜 더미 치환
        sp = "    "
        today = datetime.today()
        today_str = f"{today.year}년{sp}{today.month}월{sp}{today.day}일"
        for token in ["YYYY년 MM월 DD일", "YYYY년    MM월    DD일", "YYYY 년 MM 월 DD 일"]:
            replaced = replaced.replace(token, today_str)
        return replaced
    return _repl


def load_uploaded_workbook(uploaded_file) -> Workbook:
    """Load an uploaded workbook while providing user-friendly errors."""
    data = uploaded_file.getvalue() if uploaded_file is not None else None
    if not data:
        raise InvalidFileException("엑셀 파일이 비어 있습니다.")
    # XLSX/XLTM/XLAM files are ZIP archives. Guard against classic XLS uploads.
    if not data.startswith(b"PK"):
        raise InvalidFileException("XLSX 형식의 파일만 지원합니다. 다른 형식(xls 등)은 변환 후 업로드하세요.")
    try:
        return load_workbook(filename=io.BytesIO(data), data_only=True)
    except BadZipFile as exc:
        raise InvalidFileException("엑셀 파일이 손상되었거나 XLSX 형식이 아닙니다.") from exc


def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> bytes | None:
    try:
        with tempfile.TemporaryDirectory() as td:
            in_path = os.path.join(td, "doc.docx")
            out_path = os.path.join(td, "doc.pdf")
            with open(in_path, "wb") as f: f.write(docx_bytes)
            if docx2pdf_convert is not None:
                try:
                    docx2pdf_convert(in_path, out_path)
                    if os.path.exists(out_path):
                        with open(out_path, "rb") as f: return f.read()
                except Exception: pass
            if has_soffice():
                try:
                    subprocess.run(
                        ["soffice", "--headless", "--convert-to", "pdf", in_path, "--outdir", td],
                        check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
                    )
                    if os.path.exists(out_path):
                        with open(out_path, "rb") as f: return f.read()
                except Exception: pass
    except Exception:
        pass
    return None

@@ -196,105 +212,113 @@ st.title("🧾 납입요청서 자동 생성 (DOCX + PDF)")

col_left, col_right = st.columns([1.25, 1])

with col_left:
    # 업로더는 form 바깥: 업로드 즉시 rerun → 시트 목록 바로 표시
    h4("엑셀 파일")
    st.markdown('<div class="excel-uploader">', unsafe_allow_html=True)
    xlsx_file = st.file_uploader(
        "엑셀 업로드", type=["xlsx", "xlsm"], key="xlsx_upl",
        help="엑셀 파일을 업로드하세요", label_visibility="collapsed"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    h4("워드 템플릿(.docx)")
    st.markdown('<div class="word-uploader">', unsafe_allow_html=True)
    docx_tpl = st.file_uploader(
        "워드 템플릿 업로드", type=["docx"], key="docx_upl",
        help="Word 템플릿 파일을 업로드하세요", label_visibility="collapsed"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # 시트 선택은 업로드 직후 표시
    sheet_choice = None
    if xlsx_file is not None:
        try:
            wb_tmp = load_uploaded_workbook(xlsx_file)
            default_idx = wb_tmp.sheetnames.index(TARGET_SHEET) if TARGET_SHEET in wb_tmp.sheetnames else 0
            sheet_choice = st.selectbox("사용할 시트", wb_tmp.sheetnames, index=default_idx, key="sheet_choice")
        except InvalidFileException as e:
            st.error("지원하지 않는 엑셀 형식입니다. XLSX 파일을 업로드하세요.")
            small_note(str(e))
            xlsx_file = None
        except Exception as e:
            st.warning("엑셀 미리보기 중 문제가 발생했습니다. 생성은 가능할 수 있습니다.")
            small_note(str(e))

    out_name = st.text_input("출력 파일명", value=DEFAULT_OUT)

    gen = st.button("문서 생성", use_container_width=True)

with col_right:
    st.markdown("#### 안내")
    st.markdown(
        "- **{{A1}} / {{B7|YYYY.MM.DD}} / {{C3|#,###.00}}** 형식의 인라인 포맷 지원\n"
        "- 생성 시 WORD와 PDF 제공, **개별 다운로드** 및 **ZIP 묶음** 제공\n"
        "- PDF 변환은 **MS Word(docx2pdf)** 또는 **LibreOffice(soffice)** 필요"
    )

# ================== 생성 실행 ==================
if gen:
    if not xlsx_file or not docx_tpl:
        st.error("엑셀과 템플릿을 모두 업로드하세요.")
        st.stop()

    with st.status("문서 생성 중...", expanded=True) as status:
        try:
            st.write("1) 엑셀 로드")
            wb = load_uploaded_workbook(xlsx_file)
            ws = wb[sheet_choice] if sheet_choice else (
                wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb[wb.sheetnames[0]]
            )

            st.write("2) 템플릿 로드")
            tpl_bytes = docx_tpl.getvalue()
            doc = Document(io.BytesIO(tpl_bytes))

            st.write("3) 치환 실행")
            replacer = make_replacer(ws)
            replace_everywhere(doc, replacer)

            st.write("4) WORD 저장")
            docx_buf = io.BytesIO()
            doc.save(docx_buf); docx_buf.seek(0)
            docx_bytes = docx_buf.getvalue()

            st.write("5) PDF 변환 시도")
            pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)
            pdf_ok = pdf_bytes is not None

            st.write("6) 남은 토큰 확인")
            doc_after = Document(io.BytesIO(docx_bytes))
            leftovers = sorted(list(collect_leftover_tokens(doc_after)))

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
    dl_cols = st.columns(3)
    with dl_cols[0]:
        st.download_button("📄 WORD 다운로드", data=docx_bytes,
            file_name=ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True)
    with dl_cols[1]:
        st.download_button("🖨 PDF 다운로드", data=(pdf_bytes or b""),
            file_name=ensure_pdf(out_name), mime="application/pdf",
            disabled=not pdf_ok, help=None if pdf_ok else "PDF 변환 엔진(Word 또는 LibreOffice)이 없는 환경입니다.",
            use_container_width=True)
    with dl_cols[2]:
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            zf.writestr(ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT, docx_bytes)
            if pdf_ok: zf.writestr(ensure_pdf(out_name), pdf_bytes)
        zip_buf.seek(0)
        st.download_button("📦 ZIP (WORD+PDF)", data=zip_buf,
            file_name=(ensure_pdf(out_name).replace(".pdf","") + "_both.zip"),
