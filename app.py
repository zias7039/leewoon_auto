# ui_style.py
import streamlit as st

EXCEL_GREEN = "#217346"   # Excel signature green
WORD_BLUE   = "#185ABD"   # Word signature blue

BASE_CSS = f"""
/* ---------- 공통 레이아웃 ---------- */
#MainMenu, footer {{visibility: hidden;}}
.block-container {{ padding-top: 1.2rem; max-width: 1200px; }}

/* 버튼 */
.stButton>button {{
  height: 44px; border-radius: 10px; font-weight: 500; transition: .2s;
}}
.stButton>button:hover {{
  transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,.15);
}}
[data-testid="stDownloadButton"] > button {{ min-width: 220px; }}

/* 폼 카드 */
[data-testid="stForm"] {{
  background: rgba(248,250,252,.5);
  border: 1px solid rgba(226,232,240,.8);
  border-radius: 16px; padding: 24px;
}}

/* 입력 */
input[type="text"] {{
  border-radius: 8px!important; border: 1px solid rgba(203,213,225,.8)!important;
  padding: 10px 12px!important;
}}
input[type="text"]:focus {{
  border-color: rgba(59,130,246,.5)!important;
  box-shadow: 0 0 0 3px rgba(59,130,246,.1)!important;
}}

/* 소제목 */
.h4 {{ font-weight:700; font-size:1.05rem; margin:.25rem 0 .75rem; }}

/* ---------- 업로더 공통 ---------- */
[data-testid="stFileUploader"] {{ animation: fadeIn .3s ease-out; }}
[data-testid="stFileUploaderDropzone"] {{
  background: rgba(248,250,252,.6)!important;
  border: 2px dashed rgba(203,213,225,.6)!important;
  border-radius: 12px!important; padding: 32px 24px!important; min-height: 140px!important;
  transition: .3s cubic-bezier(.4,0,.2,1)!important;
}}
[data-testid="stFileUploaderDropzone"]:hover {{
  border-color: rgba(148,163,184,.8)!important;
  background: rgba(241,245,249,.8)!important; transform: translateY(-2px)!important;
}}

/* ---------- Excel 업로더 (초록) ---------- */
.excel-uploader [data-testid="stFileUploaderDropzone"] {{
  border-color: rgba(33,115,70,.6)!important;
  background: linear-gradient(135deg, rgba(33,115,70,.08), rgba(33,115,70,.15))!important;
}}
.excel-uploader [data-testid="stFileUploaderDropzone"]:hover {{
  border-color: rgba(33,115,70,.9)!important;
  background: linear-gradient(135deg, rgba(33,115,70,.15), rgba(33,115,70,.25))!important;
  box-shadow: 0 6px 24px rgba(33,115,70,.25)!important;
}}
.excel-uploader button {{
  background: linear-gradient(135deg, {EXCEL_GREEN}, #1a5c38)!important;
  border:1px solid rgba(33,115,70,.8)!important; color:#fff!important; font-weight:600!important;
}}

/* ---------- Word 업로더 (파랑) ---------- */
.word-uploader [data-testid="stFileUploaderDropzone"] {{
  border-color: rgba(24,90,189,.6)!important;
  background: linear-gradient(135deg, rgba(24,90,189,.08), rgba(24,90,189,.15))!important;
}}
.word-uploader [data-testid="stFileUploaderDropzone"]:hover {{
  border-color: rgba(24,90,189,.9)!important;
  background: linear-gradient(135deg, rgba(24,90,189,.15), rgba(24,90,189,.25))!important;
  box-shadow: 0 6px 24px rgba(24,90,189,.25)!important;
}}
.word-uploader button {{
  background: linear-gradient(135deg, {WORD_BLUE}, #1349a0)!important;
  border:1px solid rgba(24,90,189,.8)!important; color:#fff!important; font-weight:600!important;
}}

/* ---------- 다크 모드에서 색 유지 ---------- */
@media (prefers-color-scheme: dark) {{
  [data-testid="stForm"] {{
    background: rgba(30,41,59,.4); border-color: rgba(51,65,85,.6);
  }}
  [data-testid="stFileUploaderDropzone"] {{
    background: rgba(30,41,59,.35)!important; border-color: rgba(71,85,105,.5)!important;
  }}
  [data-testid="stFileUploaderDropzone"]:hover {{
    background: rgba(30,41,59,.55)!important; border-color: rgba(100,116,139,.7)!important;
  }}
  .h4 {{ color: rgba(248,250,252,.9); }}
}}

/* 애니메이션 & 스크롤바 */
@keyframes fadeIn {{ from{{opacity:0;transform:translateY(10px)}} to{{opacity:1;transform:translateY(0)}} }}
::-webkit-scrollbar {{ width:8px; height:8px; }}
::-webkit-scrollbar-thumb {{ background: rgba(148,163,184,.5); border-radius:4px; }}
::-webkit-scrollbar-thumb:hover {{ background: rgba(100,116,139,.7); }}
"""

def inject():
    # CSS
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)
    # 업로더에 식별 클래스 부착 (DOM 변동에도 유지)
    st.markdown(
        """
        <script>
        (function attach() {
          const apply = () => {
            const u = document.querySelectorAll('[data-testid="stFileUploader"]');
            if (u.length >= 2) {
              u[0].classList.add('excel-uploader');
              u[1].classList.add('word-uploader');
            }
          };
          apply();
          const mo = new MutationObserver(apply);
          mo.observe(document.body, {subtree:true, childList:true});
        })();
        </script>
        """,
        unsafe_allow_html=True,
    )

def h4(text: str):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)
