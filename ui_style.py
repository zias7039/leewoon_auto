# ui_style.py
import streamlit as st

BASE_CSS = """
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

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
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

def open_div(cls=""):
    st.markdown(f'<div class="{cls}">', unsafe_allow_html=True)

def close_div():
    st.markdown("</div>", unsafe_allow_html=True)

def h4(text):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)
