# ui_style.py
import streamlit as st

EXCEL_GREEN = "#22c55e"   # 밝은 엑셀 포인트 (스크린샷 느낌)
WORD_BLUE   = "#3b82f6"   # 밝은 워드 포인트

BASE_CSS = f"""
/* ---------- Global Reset / Layout ---------- */
#MainMenu, footer {{ display:none; }}
.block-container {{
  max-width: 1180px;
  padding-top: 1.2rem;
}}
html, body, [data-testid="stAppViewContainer"] {{
  background: radial-gradient(1000px 600px at 10% 0%, rgba(59,130,246,.06), transparent 40%),
              radial-gradient(900px 600px at 95% 10%, rgba(34,197,94,.06), transparent 35%),
              #0b1220;
}}
/* Side bar */
[data-testid="stSidebar"] > div:first-child {{
  background: rgba(10,15,25,.7);
}}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] ul {{
  padding-left: 0.6rem;
}}
.sidebar-item {{
  display:flex; align-items:center; gap:.5rem;
  padding:.45rem .6rem; border-radius:.6rem;
  color: #a9b5cc;
}}
.sidebar-item:hover {{ background: rgba(148,163,184,.08); color:#e5eaf3; }}
.sidebar-icon {{
  width: 20px; height: 20px; display:inline-grid; place-items:center;
  border-radius:.5rem; background: rgba(148,163,184,.12);
}}

/* ---------- Headings / Subtexts ---------- */
.app-title {{
  font-weight: 800; letter-spacing:.3px;
  font-size: 1.55rem; color:#e6edf7; margin: .2rem 0 .15rem;
}}
.app-subtitle {{
  color:#9fb1cc; margin-bottom:1.0rem;
}}

/* ---------- Glass Cards ---------- */
.card {{
  background: rgba(17,25,40,.55);
  border: 1px solid rgba(148,163,184,.16);
  border-radius: 16px;
  padding: 22px;
  box-shadow: 0 12px 40px rgba(0,0,0,.28);
  backdrop-filter: blur(8px);
}}
.card-header {{
  display:flex; align-items:center; gap:.6rem; margin-bottom:.6rem;
  color:#dfe7f5; font-weight:700;
}}
.badge {{
  font-size:.75rem; padding:.15rem .48rem; border-radius:999px;
  border:1px solid rgba(148,163,184,.25); color:#aab8d0;
}}

/* ---------- File Uploader ---------- */
[data-testid="stFileUploader"] section {{ gap: 8px !important; }}
[data-testid="stFileUploader"] button {{
  border-radius: 10px !important; padding: 8px 18px !important;
  font-weight: 600 !important;
}}
/* dropzone base */
.card [data-testid="stFileUploaderDropzone"] {{
  min-height: 140px !important;
  border-radius: 14px !important;
  border: 2px dashed rgba(148,163,184,.28) !important;
  background: rgba(15,23,42,.55) !important;
  transition: .2s ease all !important;
}}
.card [data-testid="stFileUploaderDropzone"]:hover {{
  transform: translateY(-1px);
  box-shadow: 0 8px 28px rgba(0,0,0,.25) !important;
}}
.card [data-testid="stFileUploaderDropzone"] p,
.card [data-testid="stFileUploaderDropzone"] small {{
  color:#b9c8e5 !important;
}}
/* Excel (left) */
.excel [data-testid="stFileUploaderDropzone"] {{
  border-color: {EXCEL_GREEN}33 !important;
  background: linear-gradient(135deg, rgba(34,197,94,.10), rgba(34,197,94,.05)) !important;
}}
.excel [data-testid="stFileUploaderDropzone"]:hover {{
  border-color: {EXCEL_GREEN}88 !important;
  box-shadow: 0 10px 32px {EXCEL_GREEN}33 !important;
}}
.excel [data-testid="stFileUploader"] button {{
  background: linear-gradient(135deg, #22c55e, #16a34a) !important;
  color: #0c141f !important;
}}
/* Word (right) */
.word [data-testid="stFileUploaderDropzone"] {{
  border-color: {WORD_BLUE}33 !important;
  background: linear-gradient(135deg, rgba(59,130,246,.10), rgba(59,130,246,.05)) !important;
}}
.word [data-testid="stFileUploaderDropzone"]:hover {{
  border-color: {WORD_BLUE}88 !important;
  box-shadow: 0 10px 32px {WORD_BLUE}33 !important;
}}
.word [data-testid="stFileUploader"] button {{
  background: linear-gradient(135deg, #3b82f6, #1d4ed8) !important;
  color: #eef4ff !important;
}}

/* ---------- Progress / Status ---------- */
.progress-wrap {{
  margin-top:.6rem; background: rgba(148,163,184,.15);
  height: 10px; border-radius: 20px; overflow:hidden;
}}
.progress-bar {{
  height:100%; width:0%; border-radius:20px;
  background: linear-gradient(90deg, #22c55e, #3b82f6);
  transition: width .35s ease;
}}
.legend {{
  display:grid; gap:.45rem; margin-top:.4rem;
}}
.legend-item {{ display:flex; align-items:center; gap:.5rem; color:#b6c5df; }}
.dot {{ width:10px; height:10px; border-radius:999px; display:inline-block; }}
.dot.green {{ background:#22c55e; }}
.dot.yellow{{ background:#fbbf24; }}
.dot.red   {{ background:#ef4444; }}

/* ---------- Small helpers ---------- */
.hint {{ color:#9fb1cc; font-size:.9rem; }}
.kbd {{
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  padding:.08rem .35rem; border-radius:.35rem;
  border:1px solid rgba(148,163,184,.25); color:#cfe0ff;
  background: rgba(17,24,39,.6);
}}
"""

def inject():
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

def page_header(title: str, subtitle: str = ""):
    st.markdown(
        f"""
        <div class="app-title">{title}</div>
        <div class="app-subtitle">{subtitle}</div>
        """,
        unsafe_allow_html=True,
    )

def legend():
    st.markdown(
        """
        <div class="legend">
          <div class="legend-item"><span class="dot green"></span>Completed</div>
          <div class="legend-item"><span class="dot yellow"></span>Pending Approval</div>
          <div class="legend-item"><span class="dot red"></span>Error: Data Mismatch</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
