# ui_style.py
import streamlit as st

EXCEL_GREEN = "#217346"   # Microsoft Excel Signature Green
BASE_CSS = """
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* Excel 테마 드롭존 */
.excel-upload [data-testid="stFileUploaderDropzone"] {
  border: 1px solid color-mix(in srgb, var(--excel-green) 70%, white);
  background: color-mix(in srgb, var(--excel-green) 12%, transparent);
  border-radius: 12px;
  backdrop-filter: blur(8px);
}

.excel-upload [data-testid="stFileUploaderDropzone"] p,
.excel-upload [data-testid="stFileUploaderDropzone"] span {
  color: color-mix(in srgb, var(--excel-green) 85%, white);
}

.excel-upload [data-testid="stFileUploader"] button {
  background: var(--excel-green);
  border-radius: 10px;
  border: 1px solid color-mix(in srgb, var(--excel-green) 60%, white);
  color: white;
}
.excel-upload [data-testid="stFileUploader"] button:hover {
  filter: brightness(1.08);
}


.block-container {padding-top: 1.2rem;}

.stButton>button {height: 44px; border-radius: 10px;}
[data-testid="stDownloadButton"] > button {min-width: 220px;}

.small-note {font-size:.85rem; color: rgba(0,0,0,.6);}

.upload-card {
  background: rgba(2,6,23,.65);
  border: 1px solid rgba(148,163,184,.25);
  border-radius: 16px;
  padding: 24px;
  box-shadow: 0 8px 30px rgba(0,0,0,.25);
}

/* 업로더 박스 */
.upload-card [data-testid="stFileUploaderDropzone"]{
  background: rgba(17,24,39,.55);
  border: 1px solid rgba(148,163,184,.25);
  border-radius: 12px;
  backdrop-filter: blur(8px);
  -webkit-backdrop-filter: blur(8px);
}
.upload-card [data-testid="stFileUploader"] section {
  gap: 6px;
}
.upload-card [data-testid="stFileUploader"] button {
  border-radius: 10px;
}

/* 소제목 */
.h4 { font-weight: 700; font-size: 1.05rem; margin:.25rem 0 .75rem; }
"""

def inject():
    st.markdown("""
    <style>
      :root {
        --excel-green: #217346;
      }
    </style>
    """, unsafe_allow_html=True)
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

def open_div(cls=""):
    st.markdown(f'<div class="{cls}">', unsafe_allow_html=True)

def close_div():
    st.markdown("</div>", unsafe_allow_html=True)

def h4(text):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)
