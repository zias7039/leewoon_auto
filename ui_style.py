# ui_style.py
import streamlit as st

EXCEL_GREEN = "#217346"   # Excel signature
WORD_BLUE   = "#185ABD"   # Word signature

BASE_CSS = f"""
#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}

.block-container {{ padding-top: 1.2rem; }}
.stButton>button {{ height: 44px; border-radius: 10px; }}
[data-testid="stDownloadButton"] > button {{ min-width: 220px; }}
.small-note {{ font-size:.85rem; color: rgba(0,0,0,.6); }}

/* 공통 업로드 카드 */
.upload-card {{
  background: rgba(2,6,23,.65);
  border: 1px solid rgba(148,163,184,.25);
  border-radius: 16px;
  padding: 24px;
  box-shadow: 0 8px 30px rgba(0,0,0,.25);
}}

/* 업로더 박스 공통 */
.upload-card [data-testid="stFileUploaderDropzone"]{{
  background: rgba(17,24,39,.55);
  border: 1px solid rgba(148,163,184,.25);
  border-radius: 12px;
  backdrop-filter: blur(8px);
  -webkit-backdrop-filter: blur(8px);
}}
.upload-card [data-testid="stFileUploader"] section {{ gap: 6px; }}
.upload-card [data-testid="stFileUploader"] button {{ border-radius: 10px; }}

/* 소제목 */
.h4 {{ font-weight: 700; font-size: 1.05rem; margin:.25rem 0 .75rem; }}

/* --- Excel 테마 --- */
.excel-upload [data-testid="stFileUploaderDropzone"] {{
  border: 1px solid {EXCEL_GREEN}33;
  background: {EXCEL_GREEN}1F;  /* 약한 틴트 */
}}
.excel-upload [data-testid="stFileUploaderDropzone"] p,
.excel-upload [data-testid="stFileUploaderDropzone"] span {{ color: {EXCEL_GREEN}; }}
.excel-upload [data-testid="stFileUploader"] button {{
  background: {EXCEL_GREEN}; border: 1px solid {EXCEL_GREEN}99; color: white;
}}
.excel-upload [data-testid="stFileUploader"] button:hover {{ filter: brightness(1.08); }}

/* --- Word 테마*
