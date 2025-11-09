# ui_style.py
import streamlit as st

EXCEL_GREEN = "#217346"   # Excel signature green
WORD_BLUE   = "#185ABD"   # Word signature blue

BASE_CSS = """
/* ========== 기본 레이아웃 ========== */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container { padding-top: 1.2rem; max-width: 1200px; }

/* 버튼 */
.stButton>button{
  height:44px; border-radius:10px; font-weight:500; transition:all .2s ease;
}
.stButton>button:hover{ transform:translateY(-2px); box-shadow:0 4px 12px rgba(0,0,0,.15); }

[data-testid="stDownloadButton"]>button{ min-width:220px; }

/* Form 카드 */
[data-testid="stForm"]{
  background:rgba(248,250,252,.5);
  border:1px solid rgba(226,232,240,.8);
  border-radius:16px; padding:24px;
}

/* 텍스트 입력 */
input[type="text"]{
  border-radius:8px!important; border:1px solid rgba(203,213,225,.8)!important; padding:10px 12px!important;
}
input[type="text"]:focus{
  border-color:rgba(59,130,246,.5)!important; box-shadow:0 0 0 3px rgba(59,130,246,.1)!important;
}

/* 업로더 공통 */
[data-testid="stFileUploaderDropzone"]{
  background:rgba(248,250,252,.6)!important;
  border:2px dashed rgba(203,213,225,.6)!important;
  border-radius:12px!important; padding:32px 24px!important;
  transition:all .3s cubic-bezier(.4,0,.2,1)!important; min-height:140px!important;
}
[data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(148,163,184,.8)!important;
  background:rgba(241,245,249,.8)!important; transform:translateY(-2px)!important;
}
[data-testid="stFileUploader"] section{ gap:10px!important; }
[data-testid="stFileUploader"] button{
  border-radius:8px!important; padding:8px 20px!important; font-weight:500!important; transition:all .2s ease!important;
}
[data-testid="stFileUploader"] button:hover{ transform:translateY(-1px)!important; box-shadow:0 2px 8px rgba(0,0,0,.15)!important; }

/* ========== 라이트모드: 업로더 테마 ========== */
/* Excel = 첫 번째 업로더 */
[data-testid="stForm"] [data-testid="stFileUploader"]:first-of-type [data-testid="stFileUploaderDropzone"]{
  border:2px dashed rgba(33,115,70,.6)!important;
  background:linear-gradient(135deg, rgba(33,115,70,.08) 0%, rgba(33,115,70,.15) 100%)!important;
}
[data-testid="stForm"] [data-testid="stFileUploader"]:first-of-type [data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(33,115,70,.9)!important;
  background:linear-gradient(135deg, rgba(33,115,70,.15), rgba(33,115,70,.25))!important;
  box-shadow:0 6px 24px rgba(33,115,70,.25)!important;
}
[data-testid="stForm"] [data-testid="stFileUploader"]:first-of-type button{
  background:linear-gradient(135deg, #217346 0%, #1a5c38 100%)!important;
  border:1px solid rgba(33,115,70,.8)!important; color:#fff!important; font-weight:600!important;
}

/* Word = 두 번째 업로더 */
[data-testid="stForm"] [data-testid="stFileUploader"]:nth-of-type(2) [data-testid="stFileUploaderDropzone"]{
  border:2px dashed rgba(24,90,189,.6)!important;
  background:linear-gradient(135deg, rgba(24,90,189,.08) 0%, rgba(24,90,189,.15) 100%)!important;
}
[data-testid="stForm"] [data-testid="stFileUploader"]:nth-of-type(2) [data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(24,90,189,.9)!important;
  background:linear-gradient(135deg, rgba(24,90,189,.15), rgba(24,90,189,.25))!important;
  box-shadow:0 6px 24px rgba(24,90,189,.25)!important;
}
[data-testid="stForm"] [data-testid="stFileUploader"]:nth-of-type(2) button{
  background:linear-gradient(135deg, #185ABD 0%, #1349a0 100%)!important;
  border:1px solid rgba(24,90,189,.8)!important; color:#fff!important; font-weight:600!important;
}

/* ========== 다크모드 공통 오버라이드 ========== */
@media (prefers-color-scheme: dark){
  [data-testid="stForm"]{ background:rgba(30,41,59,.4); border-color:rgba(51,65,85,.6); }

  /* 다크 기본값(회색톤) — 아래 '테마 고정'이 이걸 이김 */
  [data-testid="stFileUploaderDropzone"]{
    background:rgba(30,41,59,.40)!important; border-color:rgba(71,85,105,.50)!important;
  }
  [data-testid="stFileUploaderDropzone"]:hover{
    background:rgba(30,41,59,.60)!important; border-color:rgba(100,116,139,.70)!important;
  }
  .h4{ color:rgba(248,250,252,.9); }
}

/* ========== 다크모드 테마 고정(최후 우선순위) ========== */
/* 이 블록이 반드시 가장 아래 있어야 함 */
@media (prefers-color-scheme: dark){
  /* Excel */
  [data-testid="stForm"] [data-testid="stFileUploader"]:first-of-type [data-testid="stFileUploaderDropzone"]{
    background:linear-gradient(135deg, rgba(33,115,70,.18), rgba(33,115,70,.28))!important;
    border-color:rgba(33,115,70,.90)!important;
  }
  [data-testid="stForm"] [data-testid="stFileUploader"]:first-of-type button{
    background:linear-gradient(135deg, #217346, #1a5c38)!important;
    border-color:rgba(33,115,70,.95)!important; color:#fff!important;
  }

  /* Word */
  [data-testid="stForm"] [data-testid="stFileUploader"]:nth-of-type(2) [data-testid="stFileUploaderDropzone"]{
    background:linear-gradient(135deg, rgba(24,90,189,.18), rgba(24,90,189,.28))!important;
    border-color:rgba(24,90,189,.90)!important;
  }
  [data-testid="stForm"] [data-testid="stFileUploader"]:nth-of-type(2) button{
    background:linear-gradient(135deg, #185ABD, #1349a0)!important;
    border-color:rgba(24,90,189,.95)!important; color:#fff!important;
  }
}
"""

def inject():
    """CSS만 주입 (JS 불필요)"""
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

def h4(text):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)
