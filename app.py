# app.py
# -*- coding: utf-8 -*-
import io, os, re, tempfile, subprocess
from datetime import datetime, date
from decimal import Decimal
from zipfile import ZipFile, ZIP_DEFLATED

import streamlit as st
from openpyxl import load_workbook
from docx import Document
from docx.table import _Cell
from docx.text.paragraph import Paragraph

from ui_style import inject, page_header, legend  # 스타일

# 선택: docx2pdf
try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

# --------- Tokens & Defaults ----------
TOKEN_RE = re.compile(r"\{\{([A-Z]+[0-9]+)(?:\|([^}]+))?\}\}")
LEFTOVER_RE = re.compile(r"\{\{[^}]+\}\}")
DEFAULT_OUT = f"{datetime.today():%Y%m%d}_#_납입요청서_DB저축은행.docx"
TARGET_SHEET = "2.  배정후 청약시"

# --------- Utils ----------
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
    if s: return s
    s = fmt_number(v)
    if s: return s
    return "" if v is None else str(v)

def apply_inline_format(value, fmt: str | None) -> str:
    if fmt is None or fmt.strip() == "":
        return value_to_text(value)
    # 날짜 포맷
    if any(tok in fmt for tok in ("YYYY", "MM", "DD")):
        if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value.strip()):
            value = datetime.strptime(value.strip(), "%Y-%m-%d").date()
        if isinstance(value, (datetime, date)):
            f = fmt.replace("YYYY", "%Y").replace("MM","%m").replace("DD","%d")
            return value.strftime(f)
        return value_to_text(value)
    # 숫자 포맷
    if re.fullmatch(r"[#,0]+(?:\.[0#]+)?", fmt.replace(",", "")):
        try:
            num = float(str(value).replace(",", ""))
            decimals = len(fmt.split(".")[1]) if "." in fmt else 0
            return f"{num:,.{decimals}f}"
        except Exception:
            return value_to_text(value)
    return value_to_text(value)

def iter_block_items(parent):
    if hasattr(parent, "paragraphs") and hasattr(parent, "tables"):
        for p in parent.paragraphs: yield p
        for t in parent.tables:
            for row in t.rows:
                for cell in row.cells:
                    for item in iter_block_items(cell): yield item
    elif isinstance(parent, _Cell):
        for p in parent.paragraphs: yield p
        for t in parent.tables:
            for row in t.rows:
                for cell in row.cells:
                    for item in iter_block_items(cell): yield item

def replace_in_paragraph(par: Paragraph, repl_func):
    changed = False
    for run in par.runs:
        new_text = repl_func(run.text)
        if new_text != run.text:
            run.text = new_text
            changed = True
    if changed: return
    full_text = "".join(r.text for r in par.runs)
    new_text = repl_func(full_text)
    if new_text == full_text: return
    if par.runs:
        par.runs[0].text = new_text
        for r in par.runs[1:]: r.text = ""

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
            try: v = ws[addr].value
            except Exception: v = None
            return apply_inline_format(v, fmt)
        replaced = TOKEN_RE.sub(sub, text)
        # YYYY년 MM월 DD일 같은 플레이스홀더(간이)
        sp = "    "
        today = datetime.today()
        today_str = f"{today.year}년{sp}{today.month}월{sp}{today.day}일"
        for token in ["YYYY년 MM월 DD일", "YYYY년    MM월    DD일", "YYYY 년 MM 월 DD 일"]:
            replaced = replaced.replace(token, today_str)
        return replaced
    return _repl

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

def collect_leftover_tokens(doc: Document) -> set[str]:
    leftovers = set()
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            text = "".join(r.text for r in item.runs) if item.runs else item.text
            for m in LEFTOVER_RE.findall(text or ""): leftovers.add(m)
    for section in doc.sections:
        for container in (section.header, section.footer):
            for item in iter_block_items(container):
                if isinstance(item, Paragraph):
                    text = "".join(r.text for r in item.runs) if item.runs else item.text
                    for m in LEFTOVER_RE.findall(text or ""): leftovers.add(m)
    return leftovers

# ----------------------- UI -----------------------
st.set_page_config(page_title="Document Generator", page_icon="🧾", layout="wide")
inject()

with st.sidebar:
    st.markdown("### ")
    st.markdown(
        """
- <span class="sidebar-item"><span class="sidebar-icon">🏠</span>Dashboard</span>
- <span class="sidebar-item"><span class="sidebar-icon">🧩</span>Templates</span>
- <span class="sidebar-item"><span class="sidebar-icon">📄</span>Documents</span>
- <span class="sidebar-item"><span class="sidebar-icon">⚙️</span>Settings</span>
- <span class="sidebar-item"><span class="sidebar-icon">❓</span>Help</span>
        """,
        unsafe_allow_html=True,
    )

page_header("DOCUMENT GENERATOR", "Automate your documents")

left, right = st.columns([1.25, 1])

with left:
    # 두 업로더 카드 (엑셀 / 워드)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="card excel">', unsafe_allow_html=True)
        st.markdown('<div class="card-header">☁️⬆️  Upload Excel Template <span class="badge">.xlsx / .xlsm</span></div>', unsafe_allow_html=True)
        xlsx_file = st.file_uploader(" ", type=["xlsx", "xlsm"], key="upl_xlsx", label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="card word">', unsafe_allow_html=True)
        st.markdown('<div class="card-header">☁️⬆️  Upload Word Template <span class="badge">.docx</span></div>', unsafe_allow_html=True)
        docx_tpl = st.file_uploader("  ", type=["docx"], key="upl_docx", label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)

    # 폼 없이 바로 이름/시트 선택
    out_name = st.text_input("Output file name", value=DEFAULT_OUT, help="확장자는 자동으로 맞춰집니다(.docx/.pdf)")
    sheet_choice = None
    if xlsx_file is not None:
        try:
            wb_tmp = load_workbook(filename=io.BytesIO(xlsx_file.getvalue()), data_only=True)
            sheet_choice = st.selectbox("Select worksheet", wb_tmp.sheetnames,
                                        index=wb_tmp.sheetnames.index(TARGET_SHEET) if TARGET_SHEET in wb_tmp.sheetnames else 0)
        except Exception:
            st.warning("엑셀 미리보기 중 문제가 발생했습니다. 생성은 계속할 수 있습니다.")

    go = st.button("Generate Documents", use_container_width=True)

with right:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-header">Recent Generations</div>', unsafe_allow_html=True)
    # 진행바 (시각만)
    prog = st.empty()
    prog_bar_html = """
      <div class="progress-wrap"><div class="progress-bar" style="width:{w}%"></div></div>
    """
    prog.markdown(prog_bar_html.format(w=0), unsafe_allow_html=True)
    st.markdown('<div style="height:.6rem"></div>', unsafe_allow_html=True)
    st.markdown('<div class="card-header" style="margin-top:.2rem">Status</div>', unsafe_allow_html=True)
    legend()
    st.markdown('</div>', unsafe_allow_html=True)

# -------------------- Action --------------------
if go:
    if not xlsx_file or not docx_tpl:
        st.error("엑셀과 템플릿을 모두 업로드하세요.")
        st.stop()

    # 진행바 연출
    right.container().markdown(prog_bar_html.format(w=15), unsafe_allow_html=True)

    try:
        # 1) 엑셀 로드
        wb = load_workbook(filename=io.BytesIO(xlsx_file.read()), data_only=True)
        ws = wb[sheet_choice] if sheet_choice else (wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb[wb.sheetnames[0]])
        right.container().markdown(prog_bar_html.format(w=35), unsafe_allow_html=True)

        # 2) 템플릿 로드
        tpl_bytes = docx_tpl.read()
        doc = Document(io.BytesIO(tpl_bytes))
        right.container().markdown(prog_bar_html.format(w=55), unsafe_allow_html=True)

        # 3) 치환
        replacer = make_replacer(ws)
        replace_everywhere(doc, replacer)
        right.container().markdown(prog_bar_html.format(w=75), unsafe_allow_html=True)

        # 4) DOCX 저장
        docx_buf = io.BytesIO()
        doc.save(docx_buf); docx_buf.seek(0)
        docx_bytes = docx_buf.getvalue()

        # 5) PDF 변환 (가능 시)
        pdf_bytes = convert_docx_to_pdf_bytes(docx_bytes)
        pdf_ok = pdf_bytes is not None
        right.container().markdown(prog_bar_html.format(w=92), unsafe_allow_html=True)

        # 6) 잔여 토큰 검사
        doc_after = Document(io.BytesIO(docx_bytes))
        leftovers = sorted(list(collect_leftover_tokens(doc_after)))
        right.container().markdown(prog_bar_html.format(w=100), unsafe_allow_html=True)

    except Exception as e:
        st.error("생성 중 오류가 발생했습니다.")
        st.exception(e)
        st.stop()

    st.success("문서가 준비되었습니다.")

    dl_cols = st.columns(3)
    with dl_cols[0]:
        st.download_button("📄 WORD Download", data=docx_bytes,
            file_name=ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True)
    with dl_cols[1]:
        st.download_button("🖨 PDF Download", data=(pdf_bytes or b""),
            file_name=ensure_pdf(out_name), mime="application/pdf",
            disabled=not pdf_ok,
            help=None if pdf_ok else "PDF 변환 엔진(Word 또는 LibreOffice)이 없는 환경입니다.",
            use_container_width=True)
    with dl_cols[2]:
        zip_buf = io.BytesIO()
        with ZipFile(zip_buf, "w", ZIP_DEFLATED) as zf:
            zf.writestr(ensure_docx(out_name) if out_name.strip() else DEFAULT_OUT, docx_bytes)
            if pdf_ok: zf.writestr(ensure_pdf(out_name), pdf_bytes)
        zip_buf.seek(0)
        st.download_button("📦 ZIP (WORD+PDF)", data=zip_buf,
            file_name=(ensure_pdf(out_name).replace(".pdf","") + "_both.zip"),
            mime="application/zip", use_container_width=True)

    if leftovers:
        with st.expander("템플릿에 남아있는 토큰"):
            st.write(", ".join(leftovers))
    else:
        st.caption("모든 토큰이 치환되었습니다.")
