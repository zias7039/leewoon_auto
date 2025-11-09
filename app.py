# -*- coding: utf-8 -*-
import io
import os
import re
import tempfile
import subprocess
from datetime import datetime, date
from decimal import Decimal
from zipfile import ZipFile, ZIP_DEFLATED

import streamlit as st
from openpyxl import load_workbook
from docx import Document
from docx.table import _Cell
from docx.text.paragraph import Paragraph

# 선택: docx2pdf가 있으면 활용
try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

# ----------------- 상수 -----------------
TOKEN_RE = re.compile(r"\{\{([A-Z]+[0-9]+)(?:\|([^}]+))?\}\}")  # {{A1}} or {{A1|FORMAT}}
LEFTOVER_RE = re.compile(r"\{\{[^}]+\}\}")
DEFAULT_OUT = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행.docx"
TARGET_SHEET = "2.  배정후 청약시"

# ----------------- 유틸 -----------------
def ensure_docx(name: str) -> str:
    name = (name or "").strip()
    return name if name.lower().endswith(".docx") else (name + ".docx")

def ensure_pdf(name: str) -> str:
    base = (name or "output").strip()
    if base.lower().endswith(".docx"):
        base = base[:-5]
    return base + ".pdf"

def has_soffice() -> bool:
    return any(
        os.path.isfile(os.path.join(p, "soffice")) or os.path.isfile(os.path.join(p, "soffice.bin"))
        for p in os.environ.get("PATH", "").split(os.pathsep)
    )

def try_format_as_date(v) -> str:
    try:
        if v is None:
            return ""
        if isinstance(v, (datetime, date)):
            return f"{v.year}. {v.month}. {v.day}."
        s = str(v).strip()
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

# ----------------- 포맷 적용 -----------------
def apply_inline_format(value, fmt: str | None) -> str:
    """
    {{A1|#,###}}, {{B7|YYYY.MM.DD}} 형태의 포맷을 간단 지원.
    - 날짜 포맷: YYYY -> %Y, MM -> %m, DD -> %d
    - 숫자 포맷: '#,###' / '#,###.00' 식 → 그룹핑 + 소수 자릿수
    """
    if fmt is None or fmt.strip() == "":
        return value_to_text(value)

    # 날짜 포맷 감지
    if any(tok in fmt for tok in ("YYYY", "MM", "DD")):
        # 값이 문자열이어도 'YYYY-MM-DD'면 날짜로 파싱
        if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value.strip()):
            value = datetime.strptime(value.strip(), "%Y-%m-%d").date()
        if isinstance(value, (datetime, date)):
            f = fmt
            f = f.replace("YYYY", "%Y").replace("MM", "%m").replace("DD", "%d")
            return value.strftime(f)
        return value_to_text(value)

    # 숫자 포맷 간이 처리
    if re.fullmatch(r"[#,0]+(?:\.[0#]+)?", fmt.replace(",", "")):
        try:
            num = float(str(value).replace(",", ""))
            # 소수점 자릿수 계산
            decimals = 0
            if "." in fmt:
                decimals = len(fmt.split(".")[1])
            return f"{num:,.{decimals}f}"
        except Exception:
            return value_to_text(value)

    # 그 외는 기본 변환
    return value_to_text(value)

# ----------------- 문서 순회/치환 -----------------
def iter_block_items(parent):
    """문서의 문단/표 셀 모두 순회 (본문, 헤더/푸터 공통 사용)."""
    # python-docx 타입 체크 대신 duck-typing으로 안전 처리
    if hasattr(parent, "paragraphs") and hasattr(parent, "tables"):
        for p in parent.paragraphs:
            yield p
        for t in parent.tables:
            for row in t.rows:
                for cell in row.cells:
                    for item in iter_block_items(cell):
                        yield item
    elif isinstance(parent, _Cell):
        for p in parent.paragraphs:
            yield p
        for t in parent.tables:
            for row in t.rows:
                for cell in row.cells:
                    for item in iter_block_items(cell):
                        yield item

def replace_in_paragraph(par: Paragraph, repl_func):
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
    if new_text == full_text:
        return
    if par.runs:
        par.runs[0].text = new_text
        for r in par.runs[1:]:
            r.text = ""

def replace_everywhere(doc: Document, repl_func):
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            replace_in_paragraph(item, repl_func)
    for section in doc.sections:
        for container in (section.header, section.footer):
            for item in iter_block_items(container):
                if isinstance(item, Paragraph):
                    replace_in_paragraph(item, repl_func)

# ----------------- Excel → 치환 콜백 -----------------
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

        # YYYY/MM/DD 같은 더미 템플릿 치환(간단)
        sp = "    "
        today = datetime.today()
        today_str = f"{today.year}년{sp}{today.month}월{sp}{today.day}일"
        for token in ["YYYY년 MM월 DD일", "YYYY년    MM월    DD일", "YYYY 년 MM 월 DD 일"]:
            replaced = replaced.replace(token, today_str)
        return replaced
    return _repl

# ----------------- DOCX → PDF -----------------
def convert_docx_to_pdf_bytes(docx_bytes: bytes) -> bytes | None:
    try:
        with tempfile.TemporaryDirectory() as td:
            in_path = os.path.join(td, "doc.docx")
            out_path = os.path.join(td, "doc.pdf")
            with open(in_path, "wb") as f:
                f.write(docx_bytes)

            # 1) Word (Windows) 경로
            if docx2pdf_convert is not None:
                try:
                    docx2pdf_convert(in_path, out_path)
                    if os.path.exists(out_path):
                        with open(out_path, "rb") as f:
                            return f.read()
                except Exception:
                    pass

            # 2) LibreOffice headless
            if has_soffice():
                try:
                    subprocess.run(
                        ["soffice", "--headless", "--convert-to", "pdf", in_path, "--outdir", td],
                        check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
                    )
                    if os.path.exists(out_path):
                        with open(out_path, "rb") as f:
                            return f.read()
                except Exception:
                    pass
    except Exception:
        pass
    return None

# ----------------- 누락 토큰 수집 -----------------
def collect_leftover_tokens(doc: Document) -> set[str]:
    leftovers = set()
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            text = "".join(r.text for r in item.runs) if item.runs else item.text
            for m in LEFTOVER_RE.findall(text or ""):
                leftovers.add(m)
    for section in doc.sections:
        for container in (section.header, section.footer):
            for item in iter_block_items(container):
                if isinstance(item, Paragraph):
                    text = "".join(r.text for r in item.runs) if item.runs else item.text
                    for m in LEFTOVER_RE.findall(text or ""):
                        leftovers.add(m)
    return leftovers

# ----------------- Streamlit UI -----------------
st.set_page_config(
    page_title="납입요청서 자동 생성",
    page_icon="🧾",
    layout="wide",
)

# 최소 CSS 다듬기
st.markdown("""
<style>
/* 공통 glass layer */
.excel-upload [data-testid="stFileUploaderDropzone"],
.word-upload [data-testid="stFileUploaderDropzone"] {
  backdrop-filter: blur(18px);
  -webkit-backdrop-filter: blur(18px);
  border-radius: 14px !important;
  border: 1px solid rgba(255,255,255,0.22) !important;
  box-shadow: 0 8px 24px rgba(0,0,0,0.28);
  transition: 0.25s ease;
  padding: 6px !important;
}

/* 엑셀 업로드 : glass green */
.excel-upload [data-testid="stFileUploaderDropzone"] {
  background: linear-gradient(
      135deg,
      rgba(24, 92, 55, 0.55),
      rgba(24, 92, 55, 0.28)
  );
}
.excel-upload [data-testid="stFileUploaderDropzone"]:hover {
  background: linear-gradient(
      135deg,
      rgba(24, 92, 55, 0.68),
      rgba(24, 92, 55, 0.38)
  );
}

/* 워드 업로드 : glass blue */
.word-upload [data-testid="stFileUploaderDropzone"] {
  background: linear-gradient(
      135deg,
      rgba(24, 90, 189, 0.55),
      rgba(24, 90, 189, 0.28)
  );
}
.word-upload [data-testid="stFileUploaderDropzone"]:hover {
  background: linear-gradient(
      135deg,
      rgba(24, 90, 189, 0.68),
      rgba(24, 90, 189, 0.38)
  );
}

/* 내부 텍스트 색 */
.excel-upload [data-testid="stFileUploaderDropzone"] div,
.word-upload [data-testid="stFileUploaderDropzone"] div {
  color: rgba(255,255,255,0.92) !important;
  font-weight: 500;
}

/* Browse 버튼 */
.excel-upload [data-testid="stFileUploaderBrowseButton"],
.word-upload [data-testid="stFileUploaderBrowseButton"] {
  backdrop-filter: blur(12px);
  -webkit-backdrop-filter: blur(12px);
  background: rgba(0,0,0,0.35) !important;
  border: 1px solid rgba(255,255,255,0.35) !important;
  color: white !important;
  border-radius: 10px !important;
  padding: 6px 16px !important;
  transition: 0.25s ease;
}

.excel-upload [data-testid="stFileUploaderBrowseButton"]:hover,
.word-upload [data-testid="stFileUploaderBrowseButton"]:hover {
  background: rgba(0,0,0,0.55) !important;
  border-color: rgba(255,255,255,0.55) !important;
}
</style>
""", unsafe_allow_html=True)

st.title("🧾 납입요청서 자동 생성 (DOCX + PDF)")

col_left, col_right = st.columns([1.2, 1])
with col_left:
    with st.form("input_form", clear_on_submit=False):
        st.markdown('<div class="excel-upload">', unsafe_allow_html=True)
        xlsx_file = st.file_uploader("엑셀 파일", type=["xlsx", "xlsm"], accept_multiple_files=False)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="word-upload">', unsafe_allow_html=True)
        docx_tpl = st.file_uploader("워드 템플릿(.docx)", type=["docx"], accept_multiple_files=False)
        st.markdown('</div>', unsafe_allow_html=True)

    out_name = st.text_input("출력 파일명", value=DEFAULT_OUT)

        # 업로드되면 시트 이름 미리 읽어 선택
        sheet_choice = None
        if xlsx_file is not None:
            try:
                wb_tmp = load_workbook(filename=io.BytesIO(xlsx_file.getvalue()), data_only=True)
                sheet_choice = st.selectbox(
                    "사용할 시트",
                    wb_tmp.sheetnames,
                    index=wb_tmp.sheetnames.index(TARGET_SHEET) if TARGET_SHEET in wb_tmp.sheetnames else 0
                )
            except Exception as e:
                st.warning("엑셀 미리보기 중 문제가 발생했습니다. 생성 시도는 가능합니다.")

        submitted = st.form_submit_button("문서 생성", use_container_width=True)

with col_right:
    st.markdown("#### 안내")
    st.markdown(
        "- **{{A1}} / {{B7|YYYY.MM.DD}} / {{C3|#,###.00}}** 형식의 인라인 포맷을 지원합니다.\n"
        "- **문서 생성**을 누르면 WORD와 PDF를 만들어 **개별 다운로드**와 **ZIP 묶음**을 제공합니다.\n"
        "- PDF 변환은 **MS Word(docx2pdf)** 또는 **LibreOffice(soffice)** 가 설치된 환경에서 동작합니다.",
    )
    # 템플릿 토큰 간단 미리보기(있을 때만)
    if docx_tpl is not None:
        try:
            doc_preview = Document(io.BytesIO(docx_tpl.getvalue()))
            sample_tokens = set()
            for i, p in enumerate(doc_preview.paragraphs[:80]):  # 처음 80문단만 가볍게 스캔
                for m in re.findall(r"\{\{[^}]+\}\}", p.text or ""):
                    if len(sample_tokens) < 12:
                        sample_tokens.add(m)
            if sample_tokens:
                st.markdown("**템플릿 토큰 샘플**")
                st.code(", ".join(list(sample_tokens)))
            else:
                st.caption("템플릿에서 토큰을 찾지 못했습니다.")
        except Exception:
            st.caption("템플릿 미리보기를 불러오지 못했습니다.")

# ============ 생성 실행 ============
if submitted:
    if not xlsx_file or not docx_tpl:
        st.error("엑셀과 템플릿을 모두 업로드하세요.")
        st.stop()

    # 진행 상태 카드
    with st.status("문서 생성 중...", expanded=True) as status:
        try:
            st.write("1) 엑셀 로드")
            wb = load_workbook(filename=io.BytesIO(xlsx_file.read()), data_only=True)
            ws = wb[sheet_choice] if sheet_choice else (
                wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb[wb.sheetnames[0]]
            )

            st.write("2) 템플릿 로드")
            tpl_bytes = docx_tpl.read()
            doc = Document(io.BytesIO(tpl_bytes))

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

            # 남은 토큰 조사
            st.write("6) 남은 토큰 확인")
            doc_after = Document(io.BytesIO(docx_bytes))
            leftovers = sorted(list(collect_leftover_tokens(doc_after)))

            status.update(label="완료", state="complete", expanded=False)
        except Exception as e:
            status.update(label="오류", state="error", expanded=True)
            st.exception(e)
            st.stop()

    # ===== 결과 영역 =====
    st.success("문서가 준비되었습니다.")

    # 개별 다운로드 버튼 (Word / PDF)
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

    # ZIP 묶음
    with dl_cols[2]:
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            zf.writestr(ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT, docx_bytes)
            if pdf_ok:
                zf.writestr(ensure_pdf(out_name), pdf_bytes)
        zip_buf.seek(0)
        st.download_button(
            "📦 ZIP (WORD+PDF)",
            data=zip_buf,
            file_name=(ensure_pdf(out_name).replace(".pdf", "") + "_both.zip"),
            mime="application/zip",
            use_container_width=True,
        )

    # 남은 토큰 보고(있을 때만)
    if leftovers:
        with st.expander("템플릿에 남아있는 토큰"):
            st.write(", ".join(leftovers))
    else:
        st.caption("모든 토큰이 치환되었습니다.")
