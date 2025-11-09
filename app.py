# leewoon_auto/app.py
# Streamlit 앱 엔트리. 패키지로 실행하든 파일로 실행하든 임포트가 깨지지 않도록 폴백 처리.

from __future__ import annotations
import io
from pathlib import Path
from typing import Optional

import streamlit as st

# -------- 안전한 임포트(절대 -> 상대 -> 더미) --------
def _import_constants():
    try:
        from leewoon_auto.constants import DEFAULT_OUT, TARGET_SHEET  # 절대
        return DEFAULT_OUT, TARGET_SHEET
    except Exception:
        try:
            from .constants import DEFAULT_OUT, TARGET_SHEET           # 상대
            return DEFAULT_OUT, TARGET_SHEET
        except Exception:
            # 최후 폴백: 최소한 앱은 뜨게
            return ("output.docx", "2.  배정후 청약시")

DEFAULT_OUT, TARGET_SHEET = _import_constants()

def _import_generator():
    try:
        from leewoon_auto.services.generator import generate_documents
        return generate_documents
    except Exception:
        try:
            from .services.generator import generate_documents
            return generate_documents
        except Exception as e:
            def _missing(*args, **kwargs):
                raise ImportError(
                    "generate_documents 함수를 찾을 수 없습니다. "
                    "services/generator.py에 `def generate_documents(...)`를 정의해 주세요.\n"
                    f"원인: {e}"
                )
            return _missing

generate_documents = _import_generator()

def _import_utils_paths():
    # 선택 모듈: 없어도 앱이 동작하도록 더미 제공
    try:
        from leewoon_auto.services.utils.paths import ensure_docx, ensure_pdf
        return ensure_docx, ensure_pdf
    except Exception:
        try:
            from .services.utils.paths import ensure_docx, ensure_pdf
            return ensure_docx, ensure_pdf
        except Exception:
            return (lambda p: Path(p), lambda p: Path(p))

ensure_docx, ensure_pdf = _import_utils_paths()

# -------- Streamlit UI --------
st.set_page_config(page_title="Leewoon Auto Generator", layout="centered")

st.title("문서 생성기")
st.caption(f"기본 시트: **{TARGET_SHEET}** · 기본 출력: **{DEFAULT_OUT}**")

with st.form("gen-form", clear_on_submit=False):
    xlsx = st.file_uploader("엑셀 파일(.xlsx)", type=["xlsx"])
    docx = st.file_uploader("워드 템플릿(.docx)", type=["docx"])
    out_name = st.text_input("출력 파일명", value=DEFAULT_OUT)
    run = st.form_submit_button("생성하기")

if run:
    if not xlsx or not docx:
        st.error("엑셀과 워드 템플릿을 모두 업로드하세요.")
        st.stop()

    try:
        # 스트림으로 받은 파일을 BytesIO로 전달
        xlsx_bytes = io.BytesIO(xlsx.getvalue())
        docx_bytes = io.BytesIO(docx.getvalue())

        # 실제 생성 함수 호출 (구현은 services/generator.py)
        # 기대 시그니처 예시:
        # generate_documents(xlsx_file: BytesIO|str, docx_template: BytesIO|str, out_name: str) -> bytes|Path
        result = generate_documents(xlsx_bytes, docx_bytes, out_name)

        # 결과 처리: bytes면 다운로드 제공, Path면 파일 읽어 제공
        if isinstance(result, (bytes, bytearray)):
            st.success("생성 완료!")
            st.download_button("다운로드", data=result, file_name=out_name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        else:
            path = Path(result)
            st.success(f"생성 완료: {path.name}")
            st.download_button("다운로드", data=path.read_bytes(), file_name=path.name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    except ImportError as e:
        # generate_documents 미구현 등
        st.error(str(e))
    except Exception as e:
        st.exception(e)

# -------- 도움말(접기) --------
with st.expander("도움말"):
    st.markdown(
        """
- 실행: **터미널에서** `streamlit run leewoon_auto/app.py`
- `python -m leewoon_auto.app` 는 일반 파이썬 모듈 실행 방식이고, Streamlit 앱을 띄우려면 위 명령을 쓰세요.
- VS Code 터미널 열기: **Ctrl + `** (백틱).
- `DEFAULT_OUT` : 기본 출력 파일명.
- `TARGET_SHEET` : 엑셀에서 읽을 기본 시트명.
- 임포트 에러가 나면 패키지 루트에 `leewoon_auto/__init__.py` 가 있는지와
  `services/`, `services/utils/`에도 `__init__.py`가 있는지 확인하세요(패키지 인식용).
        """
    )
