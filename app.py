# -*- coding: utf-8 -*-
"""Streamlit entry point for 자동 납입요청서 생성기."""

import io
from zipfile import ZIP_DEFLATED, ZipFile

import streamlit as st

from document_processing import (
    DEFAULT_OUT,
    TARGET_SHEET,
    DocumentResult,
    ensure_docx,
    ensure_pdf,
    extract_template_tokens,
    generate_documents,
    get_sheet_names,
)

# ---------- UI ----------
st.set_page_config(page_title="납입요청서 자동 생성", page_icon="🧾", layout="wide")

# Glassmorphism + 브랜드 컬러
st.markdown(
    """
<style>
/* 공통 Glassmorphism 토큰 */
:root{
  --glass-bg: rgba(15, 23, 42, 0.35);         /* 유리 배경 */
  --glass-bd: rgba(148, 163, 184, 0.35);      /* 테두리 */
  --glass-shadow: 0 8px 32px rgba(0,0,0,0.35);
}

/* 래퍼 공통 카드 느낌 */
.upload-wrap{
  border-radius: 16px;
  padding: 12px;
  margin: 8px 0 18px 0;
  position: relative;
  background: linear-gradient(180deg, rgba(255,255,255,0.06), rgba(255,255,255,0.02));
  border: 1px solid var(--glass-bd);
  box-shadow: var(--glass-shadow);
  backdrop-filter: blur(10px);
}

/* 시그니처 컬러 변수 */
.excel-upload{ --brand:#107C41; }   /* MS Excel green */
.word-upload { --brand:#185ABD; }   /* MS Word blue  */

/* 업로더 드롭존 자체를 정확히 타겟팅 */
.upload-wrap [data-testid="stFileUploaderDropzone"]{
  background: var(--glass-bg) !important;
  border: 1px solid color-mix(in srgb, var(--brand) 45%, #ffffff 0%) !important;
  border-radius: 12px !important;
  transition: border-color 0.2s ease, box-shadow 0.2s ease, background 0.2s ease;
  box-shadow: inset 0 0 0 1px rgba(255,255,255,0.06);
}

/* 호버/포커스 */
.upload-wrap [data-testid="stFileUploaderDropzone"]:hover{
  border-color: color-mix(in srgb, var(--brand) 70%, #ffffff 0%) !important;
  background: rgba(15,23,42,0.42) !important;
}

/* 내부 텍스트/아이콘 컬러 */
.upload-wrap [data-testid="stFileUploader"] *{
  color: color-mix(in srgb, var(--brand) 80%, #e5e7eb 20%) !important;
}

/* Browse 버튼 */
.upload-wrap [data-testid="stFileUploader"] button{
  border-radius: 10px !important;
  background: linear-gradient(180deg, color-mix(in srgb, var(--brand) 85%, #ffffff 0%), color-mix(in srgb, var(--brand) 65%, #000000 0%)) !important;
  border: 1px solid color-mix(in srgb, var(--brand) 90%, #000 10%) !important;
}
.upload-wrap [data-testid="stFileUploader"] button:hover{
  filter: brightness(1.05);
}

/* 파일 확장자·용량 캡션 가독성 */
.upload-wrap [data-testid="stFileUploader"] small,
.upload-wrap [data-testid="stFileUploader"] p,
.upload-wrap [data-testid="stFileUploader"] span{
  color: rgba(226,232,240,0.9) !important;
}

/* (스트림릿 버전 호환용) 베이스웹 드롭존에도 적용 */
.upload-wrap [data-testid="stFileUploader"] [data-baseweb="dropzone"]{
  background: var(--glass-bg) !important;
  border: 1px solid color-mix(in srgb, var(--brand) 45%, #ffffff 0%) !important;
  border-radius: 12px !important;
}
</style>
""",
    unsafe_allow_html=True,
)

st.title("🧾 납입요청서 자동 생성 (DOCX + PDF)")

col_left, col_right = st.columns([1.2, 1])

with col_left:
    with st.form("input_form", clear_on_submit=False):
        # 엑셀 업로더 (초록)
        st.markdown('<div class="upload-wrap excel-upload">', unsafe_allow_html=True)
        xlsx_file = st.file_uploader(
            "엑셀 파일",
            type=["xlsx", "xlsm"],
            accept_multiple_files=False,
            key="xlsx_up",
        )
        st.markdown('</div>', unsafe_allow_html=True)

        # 워드 업로더 (파랑)
        st.markdown('<div class="upload-wrap word-upload">', unsafe_allow_html=True)
        docx_tpl = st.file_uploader(
            "워드 템플릿(.docx)",
            type=["docx"],
            accept_multiple_files=False,
            key="docx_up",
        )
        st.markdown('</div>', unsafe_allow_html=True)

        out_name = st.text_input("출력 파일명", value=DEFAULT_OUT)

        sheet_choice = None
        if xlsx_file is not None:
            try:
                sheet_names = get_sheet_names(xlsx_file.getvalue())
                default_index = (
                    sheet_names.index(TARGET_SHEET)
                    if TARGET_SHEET in sheet_names
                    else 0
                )
                sheet_choice = st.selectbox(
                    "사용할 시트",
                    sheet_names,
                    index=default_index,
                )
            except Exception:
                st.warning("엑셀 미리보기 중 문제가 발생했습니다. 생성 시도는 가능합니다.")

        submitted = st.form_submit_button("문서 생성", use_container_width=True)

with col_right:
    st.markdown("#### 안내")
    st.markdown(
        "- `{{A1}}`, `{{B7|YYYY.MM.DD}}`, `{{C3|#,###.00}}` 포맷 지원\n"
        "- 생성 시 WORD와 PDF 각각 다운로드 + ZIP 제공\n"
        "- PDF 변환은 MS Word(docx2pdf) 또는 LibreOffice(soffice) 필요"
    )

    if docx_tpl is not None:
        try:
            sample_tokens = extract_template_tokens(docx_tpl.getvalue())
            st.markdown("**템플릿 토큰 샘플**" if sample_tokens else "템플릿에서 토큰을 찾지 못했습니다.")
            if sample_tokens:
                st.code(", ".join(sample_tokens))
        except Exception:
            st.caption("템플릿 미리보기를 불러오지 못했습니다.")


# ---------- 생성 실행 ----------
if submitted:
    if not xlsx_file or not docx_tpl:
        st.error("엑셀과 템플릿을 모두 업로드하세요.")
        st.stop()

    with st.status("문서 생성 중...", expanded=True) as status:
        try:
            st.write("1) 엑셀 로드")
            xlsx_bytes = xlsx_file.getvalue()

            st.write("2) 템플릿 로드")
            tpl_bytes = docx_tpl.getvalue()

            st.write("3) 문서 생성")
            result: DocumentResult = generate_documents(
                xlsx_bytes,
                tpl_bytes,
                sheet_choice,
                target_sheet=TARGET_SHEET,
            )

            docx_bytes = result.docx_bytes
            pdf_bytes = result.pdf_bytes
            leftovers = result.leftovers

            st.write("4) WORD/PDF 준비 완료")
            pdf_ok = pdf_bytes is not None

            status.update(label="완료", state="complete", expanded=False)
        except Exception as exc:  # pragma: no cover - UI feedback
            status.update(label="오류", state="error", expanded=True)
            st.exception(exc)
            st.stop()

    st.success("문서가 준비되었습니다.")
    dl_cols = st.columns(3)
    with dl_cols[0]:
        st.download_button(
            "📄 WORD 다운로드",
            data=docx_bytes,
            file_name=ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with dl_cols[1]:
        st.download_button(
            "🖨 PDF 다운로드",
            data=pdf_bytes if pdf_ok else b"",
            file_name=ensure_pdf(out_name),
            mime="application/pdf",
            disabled=not pdf_ok,
            help=None if pdf_ok else "PDF 변환 엔진(Word 또는 LibreOffice)이 없는 환경입니다.",
            use_container_width=True,
        )
    with dl_cols[2]:
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
