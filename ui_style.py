# ui_style.py
import streamlit as st

EXCEL_GREEN = "#217346"   # Excel signature green
WORD_BLUE   = "#185ABD"   # (필요하면) Word blue

BASE_CSS = """
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container {padding-top: 1.2rem;}
.stButton>button {height: 44px; border-radius: 10px;}
[data-testid="stDownloadButton"] > button {min-width: 220px;}
.small-note {font-size:.85rem; color: rgba(0,0,0,.6);}

/* 카드 스타일 공통 */
.upload-card {
  background: rgba(2,6,23,.65);
  border: 1px solid rgba(148,163,184,.25);
  border-radius: 16px;
  padding: 24px;
  box-shadow: 0 8px 30px rgba(0,0,0,.25);
  margin-bottom: 14px;
}

/* 업로더 박스 공통 */
.upload-card [data-testid="stFileUploaderDropzone"]{
  background: rgba(17,24,39,.55);
  border: 1px solid rgba(148,163,184,.25);
  border-radius: 12px;
  backdrop-filter: blur(8px);
  -webkit-backdrop-filter: blur(8px);
}
.upload-card [data-testid="stFileUploader"] section { gap: 6px; }
.upload-card [data-testid="stFileUploader"] button { border-radius: 10px; }

/* 소제목 */
.h4 { font-weight: 700; font-size: 1.05rem; margin:.25rem 0 .75rem; }

/* ===== Excel 전용 테마 (우선 적용 위해 파일 끝쪽에 둠) ===== */
.upload-card.excel-upload [data-testid="stFileUploaderDropzone"] {
  border: 1px solid rgba(33, 115, 70, 0.55);
  background: rgba(33, 115, 70, 0.12);
}
.upload-card.excel-upload [data-testid="stFileUploaderDropzone"] p,
.upload-card.excel-upload [data-testid="stFileUploaderDropzone"] span {
  color: rgba(33, 115, 70, 0.85);
}
.upload-card.excel-upload [data-testid="stFileUploader"] button {
  background: var(--excel-green);
  border: 1px solid rgba(33, 115, 70, 0.6);
  color: white;
}
.upload-card.excel-upload [data-testid="stFileUploader"] button:hover { filter: brightness(1.08); }

@supports (background: color-mix(in srgb, white 10%, black)) {
  .upload-card.excel-upload [data-testid="stFileUploaderDropzone"] {
    border: 1px solid color-mix(in srgb, var(--excel-green) 70%, white);
    background: color-mix(in srgb, var(--excel-green) 12%, transparent);
  }
  .upload-card.excel-upload [data-testid="stFileUploaderDropzone"] p,
  .upload-card.excel-upload [data-testid="stFileUploaderDropzone"] span {
    color: color-mix(in srgb, var(--excel-green) 85%, white);
  }
  .upload-card.excel-upload [data-testid="stFileUploader"] button {
    border: 1px solid color-mix(in srgb, var(--excel-green) 60%, white);
  }
}
"""

def inject():
    st.markdown("""
    <style>
      :root { --excel-green: #217346; --word-blue: #185ABD; }
    </style>
    """, unsafe_allow_html=True)
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

def open_div(cls=""):
    st.markdown(f'<div class="{cls}">', unsafe_allow_html=True)

def close_div():
    st.markdown("</div>", unsafe_allow_html=True)

def h4(text):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)
