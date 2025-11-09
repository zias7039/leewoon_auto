# app.py — Streamlit 엔트리포인트 (패키지/스크립트 모두 호환)
# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import sys
import importlib
from pathlib import Path
from typing import Optional

import streamlit as st

# --------- 안전한 모듈 임포트 유틸 ---------
HERE = Path(__file__).resolve()
PKG_DIR = HERE.parent  # .../leewoon_auto
ROOT_DIR = PKG_DIR.parent

def _import(module_names: list[str]):
    """
    주어진 모듈 후보들을 순서대로 import 시도.
    실패 시 sys.path에 PKG_DIR/ROOT_DIR를 추가하고 재시도.
    """
    last_err = None
    for name in module_names:
        try:
            return importlib.import_module(name)
        except Exception as e:
            last_err = e
    # 경로 보정 후 재시도
    for p in (str(PKG_DIR), str(ROOT_DIR)):
        if p not in sys.path:
            sys.path.insert(0, p)
    for name in module_names:
        try:
            return importlib.import_module(name)
        except Exception as e:
            last_err = e
    raise last_err if last_err else ImportError(f"Cannot import any of {module_names}")

# --------- 의존 모듈 로딩 ---------
# constants
try:
    constants = _import(["leewoon_auto.constants", "constants", ".constants"])
    DEFAULT_OUT = getattr(constants, "DEFAULT_OUT", "output.docx")
    TARGET_SHEET = getattr(constants, "TARGET_SHEET", None)
except Exception:
    DEFAULT_OUT = "output.docx"
    TARGET_SHEET = None
    constants = None

# services.generator
try:
    generator = _import(
        ["leewoon_auto.services.generator", "services.generator", ".services.generator"]
    )
    generate_documents = getattr(generator, "generate_documents", None)
except Exception:
    generator = None
    generate_documents = None

# utils.paths (선택)
ensure_docx = ensure_pdf = None
try:
    paths_mod = _import(["leewoon_auto.utils.paths", "utils.paths", ".utils.paths"])
    ensure_docx = getattr(paths_mod, "ensure_docx", None)
    ensure_pdf = getattr(paths_mod, "ensure_pdf", None)
except Exception:
    pass

# --------- Streamlit UI ---------
st.set_page_config(page_title="Leewoon Auto", page_icon="🗂️", layout="wide")
st.title("Leewoon Auto – 문서 생성")

with st.sidebar:
    st.subheader("기본값")
    out_name = st.text_input("출력 파일명", value=DEFAULT_OUT, help="예: 20251109_납입요청서.docx")
    target_sheet = st.text_input(
        "엑셀 시트명 (선택)", value=TARGET_SHEET or "", placeholder="미지정 시 자동 추정"
    )

col1, col2 = st.columns(2)
with col1:
    xlsx_file = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])
with col2:
    docx_tmpl = st.file_uploader("워드 템플릿 업로드 (.docx)", type=["docx"])

run = st.button("문서 생성 실행", use_container_width=True)

def _save_to_tmp(uploaded) -> Path:
    data = uploaded.read()
    p = (Path(st.session_state.get("_tmp_dir", str(PKG_DIR))) / uploaded.name).resolve()
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_bytes(data)
    return p

def _lazy_default(name: str) -> Path:
    p = (ROOT_DIR / "outputs" / name).resolve()
    p.parent.mkdir(parents=True, exist_ok=True)
    return p

if run:
    if not generate_documents:
        st.error("generate_documents 함수를 찾을 수 없습니다. services/generator.py를 확인하세요.")
        st.stop()

    if not xlsx_file or not docx_tmpl:
        st.warning("엑셀과 템플릿을 모두 업로드하세요.")
        st.stop()

    # 업로드 저장
    xlsx_path = _save_to_tmp(xlsx_file)
    docx_path = _save_to_tmp(docx_tmpl)
    out_path = _lazy_default(out_name or DEFAULT_OUT)

    # 선택적 보정 훅
    if ensure_docx:
        docx_path = ensure_docx(str(docx_path))
    if out_path.suffix.lower() != ".docx":
        out_path = out_path.with_suffix(".docx")

    # 실행
    with st.status("생성 중…", expanded=True) as s:
        try:
            # generate_documents의 시그니처가 다를 수 있어 유연하게 호출
            # 우선순위: (excel, docx_tmpl, out, sheet) → (excel, docx_tmpl, out) → (excel, docx_tmpl)
            called = False
            for args in [
                (str(xlsx_path), str(docx_path), str(out_path), (target_sheet or None)),
                (str(xlsx_path), str(docx_path), str(out_path)),
                (str(xlsx_path), str(docx_path)),
            ]:
                try:
                    res = generate_documents(*args)
                    called = True
                    st.write(f"호출 인자: {args}")
                    break
                except TypeError:
                    continue
            if not called:
                raise RuntimeError("generate_documents 시그니처가 맞지 않습니다.")

            # 결과 안내
            if Path(out_path).exists():
                st.success(f"완료: {out_path.name}")
                st.download_button(
                    "다운로드", data=out_path.read_bytes(), file_name=out_path.name
                )
            else:
                st.info("생성 함수는 정상 호출되었으나, 출력 파일을 찾지 못했습니다. generator 내부 로직을 확인하세요.")
        except Exception as e:
            st.exception(e)
