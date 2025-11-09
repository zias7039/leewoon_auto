# leewoon_auto/app.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import sys
from pathlib import Path
import traceback

import streamlit as st

# -----------------------------------------------------------------------------
# 패키지 경로 보정: 로컬 실행 시 프로젝트 루트를 sys.path에 추가
# (…/project_root/leewoon_auto/app.py 라고 가정)
# -----------------------------------------------------------------------------
THIS_FILE = Path(__file__).resolve()
PKG_DIR = THIS_FILE.parent             # leewoon_auto/
PROJ_ROOT = PKG_DIR.parent             # project root
if str(PROJ_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJ_ROOT))

# -----------------------------------------------------------------------------
# 필수 모듈 임포트 (절대경로)
# -----------------------------------------------------------------------------
try:
    from leewoon_auto.constants import DEFAULT_OUT, TARGET_SHEET
    from leewoon_auto.services.generator import generate_documents
    # (옵션) 유틸이 있다면 사용
    try:
        from leewoon_auto.utils.paths import ensure_docx, ensure_pdf  # noqa: F401
    except Exception:
        ensure_docx = ensure_pdf = None
except Exception as e:
    # 임포트 실패 시 UI로 친절하게 원인 노출
    st.set_page_config(page_title="문서 생성기 - Import 오류")
    st.error(
        "패키지 임포트에 실패했습니다. 아래 사항을 확인하세요.\n\n"
        "1) 폴더 구조가 다음과 같은지:\n"
        "   project_root/\n"
        "     └─ leewoon_auto/\n"
        "         ├─ __init__.py\n"
        "         ├─ app.py\n"
        "         ├─ constants.py\n"
        "         ├─ services/\n"
        "         │   ├─ __init__.py\n"
        "         │   └─ generator.py\n"
        "         └─ utils/\n"
        "             ├─ __init__.py\n"
        "             └─ (docx_tools.py 등)\n\n"
        "2) 각 폴더에 __init__.py 가 있는지 (루트/ services/ utils/ 총 3개)\n"
        "3) 지금 파일(app.py)이 leewoon_auto/ 바로 아래에 위치하는지\n"
        "4) PROJ_ROOT(한 단계 위 경로)가 sys.path에 추가되는지\n"
    )
    with st.expander("Python 에러 트레이스 보기"):
        st.code("".join(traceback.format_exception(e)), language="python")
    st.stop()

# -----------------------------------------------------------------------------
# Streamlit 페이지 설정
# -----------------------------------------------------------------------------
st.set_page_config(page_title="엑셀→워드 자동 문서 생성기", page_icon="🧩", layout="centered")

st.title("🧩 엑셀→워드 자동 문서 생성기")
st.caption(
    f"기본 시트: **{TARGET_SHEET}**, 기본 출력 파일명: **{DEFAULT_OUT}**"
)

# -----------------------------------------------------------------------------
# 업로드 위젯
# -----------------------------------------------------------------------------
xlsx_file = st.file_uploader("엑셀 파일 (.xlsx, .xlsm)", type=["xlsx", "xlsm"], key="xlsx_upl")
docx_file = st.file_uploader("워드 템플릿 (.docx)", type=["docx"], key="docx_upl")

out_name = st.text_input("출력 파일명", value=DEFAULT_OUT, help="예: 20251109_#_납입요청서_DB저축은행.docx")

left, right = st.columns([1, 1])
with left:
    run_btn = st.button("생성하기", type="primary")
with right:
    st.write("")  # spacing

# -----------------------------------------------------------------------------
# 헬퍼
# -----------------------------------------------------------------------------
def _to_bytes(uploaded) -> bytes:
    buf = io.BytesIO(uploaded.read())
    return buf.getvalue()

def _offer_download(name_hint: str, data_or_path):
    """
    data_or_path 가 (bytes | str[경로]) 모두 가능하도록 처리.
    """
    if data_or_path is None:
        return
    if isinstance(data_or_path, (bytes, bytearray)):
        st.download_button(
            label=f"📥 {name_hint} 다운로드",
            data=data_or_path,
            file_name=name_hint,
            mime="application/octet-stream",
        )
    else:
        p = Path(str(data_or_path))
        if p.exists():
            st.download_button(
                label=f"📥 {p.name} 다운로드",
                data=p.read_bytes(),
                file_name=p.name,
                mime="application/octet-stream",
            )

# -----------------------------------------------------------------------------
# 실행
# -----------------------------------------------------------------------------
if run_btn:
    if not xlsx_file or not docx_file:
        st.warning("엑셀과 워드 템플릿을 모두 업로드하세요.")
        st.stop()

    xlsx_bytes = _to_bytes(xlsx_file)
    docx_bytes = _to_bytes(docx_file)

    try:
        with st.spinner("문서 생성 중..."):
            # generate_documents 인터페이스 호환 처리
            # 기대 인자: (xlsx_bytes, docx_bytes, out_name)
            result = generate_documents(xlsx_bytes, docx_bytes, out_name)

        st.success("생성이 완료되었습니다.")

        # 반환 타입에 따라 유연 처리
        # 1) dict: {'docx': bytes|path, 'pdf': bytes|path, 'logs': str, ...}
        # 2) tuple/list: (docx, pdf?) 혹은 (docx,)
        # 3) 단일 bytes/path
        if isinstance(result, dict):
            docx_out = result.get("docx") or result.get("docx_path")
            pdf_out  = result.get("pdf") or result.get("pdf_path")
            logs     = result.get("logs")
            if docx_out:
                _offer_download(out_name if isinstance(docx_out, (bytes, bytearray)) else docx_out, docx_out)
            if pdf_out:
                pdf_name = Path(out_name).with_suffix(".pdf").name
                _offer_download(pdf_name if isinstance(pdf_out, (bytes, bytearray)) else pdf_out, pdf_out)
            if logs:
                with st.expander("로그 보기"):
                    st.code(str(logs))
        elif isinstance(result, (tuple, list)):
            if len(result) >= 1:
                docx_out = result[0]
                _offer_download(out_name if isinstance(docx_out, (bytes, bytearray)) else docx_out, docx_out)
            if len(result) >= 2 and result[1] is not None:
                pdf_out = result[1]
                pdf_name = Path(out_name).with_suffix(".pdf").name
                _offer_download(pdf_name if isinstance(pdf_out, (bytes, bytearray)) else pdf_out, pdf_out)
        else:
            # 단일 결과
            _offer_download(out_name if isinstance(result, (bytes, bytearray)) else result, result)

    except Exception as e:
        st.error("문서 생성 중 오류가 발생했습니다. 아래 내용을 확인하세요.")
        with st.expander("에러 세부정보"):
            st.code("".join(traceback.format_exception(e)), language="python")

# -----------------------------------------------------------------------------
# 디버그/도움말
# -----------------------------------------------------------------------------
with st.expander("도움말 / 환경 진단"):
    st.markdown(
        "- **DEFAULT_OUT**: 기본 출력 파일명 템플릿 (예: 오늘 날짜 기반)\n"
        "- **TARGET_SHEET**: 엑셀에서 기본으로 참조할 시트 이름\n"
        "- 임포트 오류 시 `__init__.py`가 **leewoon_auto/**, **leewoon_auto/services/**, **leewoon_auto/utils/**에 각각 존재해야 합니다."
    )
    st.write("프로젝트 루트:", str(PROJ_ROOT))
    st.write("sys.path[0:3]:", sys.path[:3])
