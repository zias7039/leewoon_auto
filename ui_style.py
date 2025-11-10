# ui_style.py
import streamlit as st

EXCEL_GREEN = "#217346"   # Excel signature green
WORD_BLUE   = "#185ABD"   # Word signature blue

BASE_CSS = """
/* 기본 레이아웃 */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
.block-container { padding-top: 1.2rem; max-width: 1200px; }

/* 버튼 */
.stButton>button {
  height: 44px; border-radius: 10px; font-weight: 500; transition: all .2s ease;
}
.stButton>button:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,.15); }

/* 다운로드 버튼 폭 */
[data-testid="stDownloadButton"] > button { min-width: 220px; }

/* 텍스트 입력 */
input[type="text"]{ border-radius:8px !important; border:1px solid rgba(203,213,225,.8)!important; padding:10px 12px!important;}
input[type="text"]:focus{ border-color:rgba(59,130,246,.5)!important; box-shadow:0 0 0 3px rgba(59,130,246,.1)!important;}

/* 소제목 */
.h4{ font-weight:700; font-size:1.05rem; margin:.25rem 0 .75rem; color:rgba(15,23,42,.9); }

/* 업로더 공통 */
[data-testid="stFileUploaderDropzone"]{
  border:2px dashed rgba(203,213,225,.6) !important;
  border-radius:12px !important; padding:32px 24px !important;
  transition: all .3s ease !important; min-height:140px !important;
}
[data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(148,163,184,.8)!important; background:rgba(241,245,249,.8)!important;
  transform:translateY(-2px)!important;
}
[data-testid="stFileUploader"] section{ gap:10px!important; }
[data-testid="stFileUploader"] button{
  border-radius:8px!important; padding:8px 20px!important; font-weight:500!important; transition: all .2s ease!important;
}
[data-testid="stFileUploader"] button:hover{ transform: translateY(-1px)!important; box-shadow:0 2px 8px rgba(0,0,0,.15)!important; }

/* Excel 테마 (래퍼 클래스) */
.excel-uploader [data-testid="stFileUploaderDropzone"]{
  border:2px dashed rgba(33,115,70,.6)!important;
  background:linear-gradient(135deg, rgba(33,115,70,.08) 0%, rgba(33,115,70,.15) 100%)!important;
}
.excel-uploader [data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(33,115,70,.9)!important;
  background:linear-gradient(135deg, rgba(33,115,70,.15) 0%, rgba(33,115,70,.25) 100%)!important;
  box-shadow:0 6px 24px rgba(33,115,70,.25)!important;
}
.excel-uploader p, .excel-uploader span{ color:rgba(33,115,70,1)!important; font-weight:600!important; }
.excel-uploader small{ color:rgba(33,115,70,.75)!important; }
.excel-uploader button{
  background:linear-gradient(135deg, #217346 0%, #1a5c38 100%)!important;
  border:1px solid rgba(33,115,70,.8)!important; color:#fff!important; font-weight:600!important;
}

/* Word 테마 (래퍼 클래스) */
.word-uploader [data-testid="stFileUploaderDropzone"]{
  border:2px dashed rgba(24,90,189,.6)!important;
  background:linear-gradient(135deg, rgba(24,90,189,.08) 0%, rgba(24,90,189,.15) 100%)!important;
}
.word-uploader [data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(24,90,189,.9)!important;
  background:linear-gradient(135deg, rgba(24,90,189,.15) 0%, rgba(24,90,189,.25) 100%)!important;
  box-shadow:0 6px 24px rgba(24,90,189,.25)!important;
}
.word-uploader p, .word-uploader span{ color:rgba(24,90,189,1)!important; font-weight:600!important; }
.word-uploader small{ color:rgba(24,90,189,.75)!important; }
.word-uploader button{
  background:linear-gradient(135deg, #185ABD 0%, #1349a0 100%)!important;
  border:1px solid rgba(24,90,189,.8)!important; color:#fff!important; font-weight:600!important;
}

/* 다크모드: 색 보존 */
@media (prefers-color-scheme: dark){
  .h4{ color:rgba(248,250,252,.9); }
  .excel-uploader [data-testid="stFileUploaderDropzone"]{
    background:linear-gradient(135deg, rgba(33,115,70,.18), rgba(33,115,70,.28))!important;
    border-color:rgba(33,115,70,.85)!important;
  }
  .excel-uploader [data-testid="stFileUploaderDropzone"]:hover{
    background:linear-gradient(135deg, rgba(33,115,70,.25), rgba(33,115,70,.35))!important;
  }
  .word-uploader [data-testid="stFileUploaderDropzone"]{
    background:linear-gradient(135deg, rgba(24,90,189,.18), rgba(24,90,189,.28))!important;
    border-color:rgba(24,90,189,.85)!important;
  }
  .word-uploader [data-testid="stFileUploaderDropzone"]:hover{
    background:linear-gradient(135deg, rgba(24,90,189,.25), rgba(24,90,189,.35))!important;
  }
}

/* 애니메이션 */
@keyframes fadeIn{ from{opacity:0; transform:translateY(10px);} to{opacity:1; transform:translateY(0);} }
[data-testid="stFileUploader"]{ animation: fadeIn .3s ease-out; }

/* 스크롤바 */
::-webkit-scrollbar{ width:8px; height:8px; }
::-webkit-scrollbar-track{ background:rgba(241,245,249,.5); border-radius:4px; }
::-webkit-scrollbar-thumb{ background:rgba(148,163,184,.5); border-radius:4px; }
::-webkit-scrollbar-thumb:hover{ background:rgba(100,116,139,.7); }
"""

def inject():
    st.markdown("""
    <style>
      :root {
        --excel-green: #217346;
        --word-blue: #185ABD;
        --border-radius: 12px;
        --transition: all 0.2s ease;
      }
    </style>
    """, unsafe_allow_html=True)
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

def h4(text): st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)
def small_note(text): st.markdown(f'<div class="small-note">{text}</div>', unsafe_allow_html=True)
