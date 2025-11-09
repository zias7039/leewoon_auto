# ui_style.py
import streamlit as st

EXCEL_GREEN = "#217346"   # Excel signature green
WORD_BLUE   = "#185ABD"   # Word signature blue

BASE_CSS = r"""
/* ===== 기본 레이아웃/컴포넌트 ===== */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
.block-container { padding-top: 1.2rem; max-width: 1200px; }

.stButton>button{
  height:44px;border-radius:10px;font-weight:500;transition:.2s ease;
}
.stButton>button:hover{ transform:translateY(-2px); box-shadow:0 4px 12px rgba(0,0,0,.15); }
[data-testid="stDownloadButton"]>button{ min-width:220px; }

[data-testid="stForm"]{
  background:rgba(248,250,252,.5); border:1px solid rgba(226,232,240,.8);
  border-radius:16px; padding:24px;
}

input[type="text"]{
  border-radius:8px!important; border:1px solid rgba(203,213,225,.8)!important; padding:10px 12px!important;
}
input[type="text"]:focus{
  border-color:rgba(59,130,246,.5)!important; box-shadow:0 0 0 3px rgba(59,130,246,.1)!important;
}

.h4{ font-weight:700; font-size:1.05rem; margin:.25rem 0 .75rem; color:rgba(15,23,42,.9); }
.small-note{ font-size:.85rem; color:rgba(100,116,139,.8); }

/* 공통 업로더 베이스 */
[data-testid="stFileUploaderDropzone"]{
  background:rgba(248,250,252,.6)!important;
  border:2px dashed rgba(203,213,225,.6)!important;
  border-radius:12px!important; padding:32px 24px!important;
  transition:all .3s cubic-bezier(.4,0,.2,1)!important; min-height:140px!important;
}
[data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(148,163,184,.8)!important; background:rgba(241,245,249,.8)!important;
  transform:translateY(-2px)!important; box-shadow:0 4px 12px rgba(0,0,0,.08)!important;
}
[data-testid="stFileUploader"] section{ gap:10px!important; }
[data-testid="stFileUploader"] button{
  border-radius:8px!important; padding:8px 20px!important; font-weight:500!important; transition:.2s ease!important;
}
[data-testid="stFileUploader"] button:hover{ transform:translateY(-1px)!important; box-shadow:0 2px 8px rgba(0,0,0,.15)!important; }

/* ===== 테마: 클래스 기반 (정확, 견고) ===== */
/* Excel */
.excel-uploader [data-testid="stFileUploaderDropzone"]{
  border:2px dashed rgba(33,115,70,.6)!important;
  background:linear-gradient(135deg, rgba(33,115,70,.08) 0%, rgba(33,115,70,.15) 100%)!important;
}
.excel-uploader [data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(33,115,70,.9)!important;
  background:linear-gradient(135deg, rgba(33,115,70,.15), rgba(33,115,70,.25))!important;
  box-shadow:0 6px 24px rgba(33,115,70,.25)!important;
}
.excel-uploader button{
  background:linear-gradient(135deg, #217346 0%, #1a5c38 100%)!important;
  border:1px solid rgba(33,115,70,.8)!important; color:#fff!important; font-weight:600!important;
}

/* Word */
.word-uploader [data-testid="stFileUploaderDropzone"]{
  border:2px dashed rgba(24,90,189,.6)!important;
  background:linear-gradient(135deg, rgba(24,90,189,.08) 0%, rgba(24,90,189,.15) 100%)!important;
}
.word-uploader [data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(24,90,189,.9)!important;
  background:linear-gradient(135deg, rgba(24,90,189,.15), rgba(24,90,189,.25))!important;
  box-shadow:0 6px 24px rgba(24,90,189,.25)!important;
}
.word-uploader button{
  background:linear-gradient(135deg, #185ABD 0%, #1349a0 100%)!important;
  border:1px solid rgba(24,90,189,.8)!important; color:#fff!important; font-weight:600!important;
}

/* 상태/익스팬더 기타 */
.stAlert{ border-radius:12px; border-left-width:4px; }
[data-testid="stStatusWidget"]{ border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,.1); }
[data-testid="stExpander"]{ border-radius:10px; border:1px solid rgba(226,232,240,.8); }
[data-testid="stExpander"] summary{ border-radius:10px; }
[data-testid="column"]{ padding:0 8px; }

/* 다크모드 — 색상 보존 오버라이드 */
@media (prefers-color-scheme: dark){
  [data-testid="stForm"]{ background:rgba(30,41,59,.4); border-color:rgba(51,65,85,.6); }
  [data-testid="stFileUploaderDropzone"]{
    background:rgba(30,41,59,.4)!important; border-color:rgba(71,85,105,.5)!important;
  }
  [data-testid="stFileUploaderDropzone"]:hover{
    background:rgba(30,41,59,.6)!important; border-color:rgba(100,116,139,.7)!important;
  }
  .h4{ color:rgba(248,250,252,.9); }

  /* Excel 다크 */
  .excel-uploader [data-testid="stFileUploaderDropzone"]{
    background:linear-gradient(135deg, rgba(33,115,70,.18), rgba(33,115,70,.28))!important;
    border-color:rgba(33,115,70,.85)!important;
  }
  .excel-uploader [data-testid="stFileUploaderDropzone"]:hover{
    background:linear-gradient(135deg, rgba(33,115,70,.25), rgba(33,115,70,.35))!important;
  }
  .excel-uploader button{
    background:linear-gradient(135deg, #217346, #1a5c38)!important; border-color:rgba(33,115,70,.9)!important;
  }

  /* Word 다크 */
  .word-uploader [data-testid="stFileUploaderDropzone"]{
    background:linear-gradient(135deg, rgba(24,90,189,.18), rgba(24,90,189,.28))!important;
    border-color:rgba(24,90,189,.85)!important;
  }
  .word-uploader [data-testid="stFileUploaderDropzone"]:hover{
    background:linear-gradient(135deg, rgba(24,90,189,.25), rgba(24,90,189,.35))!important;
  }
  .word-uploader button{
    background:linear-gradient(135deg, #185ABD, #1349a0)!important; border-color:rgba(24,90,189,.9)!important;
  }
}

/* 애니메이션 & 스크롤바 */
@keyframes fadeIn{ from{opacity:0;transform:translateY(10px);} to{opacity:1;transform:translateY(0);} }
[data-testid="stFileUploader"]{ animation:fadeIn .3s ease-out; }
::-webkit-scrollbar{ width:8px; height:8px; }
::-webkit-scrollbar-track{ background:rgba(241,245,249,.5); border-radius:4px; }
::-webkit-scrollbar-thumb{ background:rgba(148,163,184,.5); border-radius:4px; }
::-webkit-scrollbar-thumb:hover{ background:rgba(100,116,139,.7); }
"""

def inject():
    # CSS 먼저
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

    # 업로더를 '키'로 정확히 찾아 클래스 부착
    st.markdown("""
    <script>
    (function attach(){
      function mark(){
        // key 로 부착: xlsx_upl -> excel, docx_upl -> word
        const x = document.querySelector('input#xlsx_upl');
        const d = document.querySelector('input#docx_upl');
        if (x) { const u = x.closest('[data-testid="stFileUploader"]'); if (u) u.classList.add('excel-uploader'); }
        if (d) { const u = d.closest('[data-testid="stFileUploader"]'); if (u) u.classList.add('word-uploader'); }
      }
      // 초기 시도 + 동적 렌더 대응
      setTimeout(mark, 100);
      const obs = new MutationObserver(mark);
      obs.observe(document.body, {childList:true, subtree:true});
    })();
    </script>
    """, unsafe_allow_html=True)

def h4(text):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)

def small_note(text):
    st.markdown(f'<div class="small-note">{text}</div>', unsafe_allow_html=True)
