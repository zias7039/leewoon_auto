# invoicegen/app.py
# -*- coding: utf-8 -*-
import io
from datetime import datetime
import streamlit as st
from docx import Document

from constants import DEFAULT_OUT, TARGET_SHEET
from utils.excel_tools import load_wb_and_guess_sheet
from utils.paths import ensure_docx, ensure_pdf
from services.generator import generate_documents

st.set_page_config(page_title="납입요청서 자동 생성", page_icon="🧾", layout="wide")

# --- 최소 CSS ---
st.markdown("""
<style>
#MainMenu {visibility: hidden;} footer {visibility: hidden;}
.block-container {padding-top: 1.2rem;}
div[data-testid="stForm"] {border: 1px solid rgba(0,0,0,.08); padding: 1rem 1rem .5rem 1rem; border-radius: 12px;}
.stButton>button {height: 44px; border-radius: 10px;}
[data-testid="stDownloadButton"] > button {min-width: 220px;}
.small-note {font-size:.85rem; color: rgba(0,0,0,.6);}
</style>
""", unsafe_allow_html=True)

st.title("🧾 납입요청서 자동 생성 (DOCX + PDF)")

col_left, col_right = st.columns([1.2, 1])
with col_left:
    with st.form("input_form", clear_on_submit=False):
        xlsx_file = st.file_uploader("엑셀 파일", type=["xlsx", "xlsm"], accept_multiple_files=False)
        docx_tpl = st.file_uploader("워드 템플릿(.docx)", type=["docx"], accept_multiple_files=False)

        out_name = st.text_input("출력 파일명", value=DEFAULT_OUT)

        sheet_choice = None
        if xlsx_file is not None:
            sheet_choice = load_wb_and_guess_sheet(xlsx_file, TARGET_SHEET, show_warning=True)

        submitted = st.form_submit_button("문서 생성", use_container_width=True)

with col_right:
    st.markdown("#### 안내")
    st.markdown(
        "- **{{A1}} / {{B7|YYYY.MM.DD}} / {{C3|#,###.00}}** 형식의 인라인 포맷을 지원합니다.\n"
        "- **문서 생성**을 누르면 WORD와 PDF를 만들어 **개별 다운로드**와 **ZIP 묶음**을 제공합니다.\n"
        "- PDF 변환은 **MS Word(docx2pdf)** 또는 **LibreOffice(soffice)** 가 설치된 환경에서 동작합니다.",
    )
    if docx_tpl is not None:
        try:
            doc_preview = Document(io.BytesIO(docx_tpl.getvalue()))
            sample_tokens = set()
            for p in doc_preview.paragraphs[:80]:
                for m in __import__("re").findall(r"\{\{[^}]+\}\}", p.text or ""):
                    if len(sample_tokens) < 12:
                        sample_tokens.add(m)
            if sample_tokens:
                st.markdown("**템플릿 토큰 샘플**")
                st.code(", ".join(list(sample_tokens)))
            else:
                st.caption("템플릿에서 토큰을 찾지 못했습니다.")
        except Exception:
            st.caption("템플릿 미리보기를 불러오지 못했습니다.")

if submitted:
    if not xlsx_file or not docx_tpl:
        st.error("엑셀과 템플릿을 모두 업로드하세요.")
        st.stop()

    with st.status("문서 생성 중...", expanded=True) as status:
        try:
            status.write("1) 생성 실행")
            result = generate_documents(
                xlsx_bytes=xlsx_file.read(),
                docx_tpl_bytes=docx_tpl.read(),
                sheet_name=sheet_choice or TARGET_SHEET,
                out_name=out_name,
            )
            status.update(label="완료", state="complete", expanded=False)
        except Exception as e:
            status.update(label="오류", state="error", expanded=True)
            st.exception(e)
            st.stop()

    st.success("문서가 준비되었습니다.")
    dl_cols = st.columns(3)
    with dl_cols[0]:
        st.download_button(
            "📄 WORD 다운로드",
            data=result.docx_bytes,
            file_name=ensure_docx(result.out_name),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with dl_cols[1]:
        st.download_button(
            "🖨 PDF 다운로드",
            data=result.pdf_bytes if result.pdf_ok else b"",
            file_name=ensure_pdf(result.out_name),
            mime="application/pdf",
            disabled=not result.pdf_ok,
            help=None if result.pdf_ok else "PDF 변환 엔진(Word 또는 LibreOffice)이 없는 환경입니다.",
            use_container_width=True,
        )
    with dl_cols[2]:
        st.download_button(
            "📦 ZIP (WORD+PDF)",
            data=result.zip_bytes,
            file_name=(ensure_pdf(result.out_name).replace(".pdf", "") + "_both.zip"),
            mime="application/zip",
            use_container_width=True,
        )

    if result.leftovers:
        with st.expander("템플릿에 남아있는 토큰"):
            st.write(", ".join(result.leftovers))
    else:
        st.caption("모든 토큰이 치환되었습니다.")
