# ui_style.py
import streamlit as st

EXCEL_GREEN = "#217346"   # Excel signature green
WORD_BLUE   = "#185ABD"   # Word signature blue

BASE_CSS = """
/* ----------------- 공통 레이아웃 ----------------- */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container { padding-top: 1.2rem; max-width: 1200px; }

/* 버튼 */
.stButton>button {
  height: 44px; border-radius: 10px; font-weight: 500; transition: all .2s ease;
}
.stButton>button:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,.15); }

/* 다운로드 버튼 */
[data-testid="stDownloadButton"] > button { min-width: 220px; }

/* Form 카드 */
[data-testid="stForm"] {
  background: rgba(248,250,252,.5);
  border: 1px solid rgba(226,232,240,.8);
  border-radius: 16px; padding: 24px;
}

/* 텍스트 입력 */
input[type="text"] {
  border-radius: 8px !important;
  border: 1px solid rgba(203,213,225,.8) !important;
  padding: 10px 12px !important;
}
input[type="text"]:focus {
  border-color: rgba(59,130,246,.5) !important;
  box-shadow: 0 0 0 3px rgba(59,130,246,.1) !important;
}

/* 업로더(기본) */
[data-testid="stFileUploaderDropzone"]{
  background: rgba(248,250,252,.6) !important;
  border: 2px dashed rgba(203,213,225,.6) !important;
  border-radius: 12px !important; padding: 32px 24px !important;
  transition: all .3s cubic-bezier(.4,0,.2,1) !important; min-height: 140px !important;
}
[data-testid="stFileUploaderDropzone"]:hover{
  border-color: rgba(148,163,184,.8) !important;
  background: rgba(241,245,249,.8) !important;
  transform: translateY(-2px) !important; box-shadow: 0 4px 12px rgba(0,0,0,.08) !important;
}
[data-testid="stFileUploader"] section { gap: 10px !important; }
[data-testid="stFileUploader"] button {
  border-radius: 8px !important; padding: 8px 20px !important; font-weight: 500 !important;
  transition: all .2s ease !important;
}
[data-testid="stFileUploader"] button:hover { transform: translateY(-1px) !important; box-shadow: 0 2px 8px rgba(0,0,0,.15) !important; }

/* 소제목 / 노트 */
.h4 { font-weight: 700; font-size: 1.05rem; margin: .25rem 0 .75rem; color: rgba(15,23,42,.9); }
.small-note { font-size: .85rem; color: rgba(100,116,139,.8); }

/* 다크모드(베이스) — 테마가 이 위를 덮어씌움 */
@media (prefers-color-scheme: dark){
  [data-testid="stForm"]{ background: rgba(30,41,59,.4); border-color: rgba(51,65,85,.6); }
  [data-testid="stFileUploaderDropzone"]{
    background: rgba(30,41,59,.4) !important; border-color: rgba(71,85,105,.5) !important;
  }
  [data-testid="stFileUploaderDropzone"]:hover{
    background: rgba(30,41,59,.6) !important; border-color: rgba(100,116,139,.7) !important;
  }
  .h4 { color: rgba(248,250,252,.9); }
}

/* 애니메이션/스크롤바 */
@keyframes fadeIn{ from{opacity:0;transform:translateY(10px)} to{opacity:1;transform:translateY(0)} }
[data-testid="stFileUploader"]{ animation: fadeIn .3s ease-out; }
::-webkit-scrollbar{ width:8px; height:8px; } ::-webkit-scrollbar-track{ background: rgba(241,245,249,.5); border-radius:4px; }
::-webkit-scrollbar-thumb{ background: rgba(148,163,184,.5); border-radius:4px; }
::-webkit-scrollbar-thumb:hover{ background: rgba(100,116,139,.7); }
"""

/* ---- 테마(Excel/Word) : 명시적 클래스 기반, 항상 맨 마지막에 와서 덮어쓰기 ---- */
THEME_CSS = """
/* Excel 업로더 */
.excel-uploader [data-testid="stFileUploaderDropzone"]{
  border: 2px dashed rgba(33,115,70,.75) !important;
  background: linear-gradient(135deg, rgba(33,115,70,.10) 0%, rgba(33,115,70,.20) 100%) !important;
}
.excel-uploader [data-testid="stFileUploaderDropzone"]:hover{
  border-color: rgba(33,115,70,.95) !important;
  background: linear-gradient(135deg, rgba(33,115,70,.18), rgba(33,115,70,.28)) !important;
  box-shadow: 0 6px 24px rgba(33,115,70,.25) !important;
}
.excel-uploader button{
  background: linear-gradient(135deg, #217346 0%, #1a5c38 100%) !important;
  border: 1px solid rgba(33,115,70,.9) !important; color:#fff !important; font-weight:600 !important;
}

/* Word 업로더 */
.word-uploader [data-testid="stFileUploaderDropzone"]{
  border: 2px dashed rgba(24,90,189,.75) !important;
  background: linear-gradient(135deg, rgba(24,90,189,.10) 0%, rgba(24,90,189,.20) 100%) !important;
}
.word-uploader [data-testid="stFileUploaderDropzone"]:hover{
  border-color: rgba(24,90,189,.95) !important;
  background: linear-gradient(135deg, rgba(24,90,189,.18), rgba(24,90,189,.28)) !important;
  box-shadow: 0 6px 24px rgba(24,90,189,.25) !important;
}
.word-uploader button{
  background: linear-gradient(135deg, #185ABD 0%, #1349a0 100%) !important;
  border: 1px solid rgba(24,90,189,.9) !important; color:#fff !important; font-weight:600 !important;
}
"""

def inject():
    # CSS 순서: 베이스 → 테마 (테마가 항상 덮어씀)
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)
    st.markdown(f"<style>{THEME_CSS}</style>", unsafe_allow_html=True)

    # 업로더에 클래스 부착 (Form 내부에서 1,2번째 업로더만 지정)
    st.markdown("""
    <script>
    (function(){
      function applyClasses(){
        const forms = document.querySelectorAll('[data-testid="stForm"]');
        forms.forEach(function(form){
          const ups = form.querySelectorAll('[data-testid="stFileUploader"]');
          if(ups.length >= 1) ups[0].classList.add('excel-uploader');
          if(ups.length >= 2) ups[1].classList.add('word-uploader');
        });
      }
      // 최초 적용
      window.requestAnimationFrame(applyClasses);
      // 동적 렌더링 대응
      const obs = new MutationObserver(function(){ applyClasses(); });
      obs.observe(document.body, { childList:true, subtree:true });
    })();
    </script>
    """, unsafe_allow_html=True)

def h4(text: str):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)

def small_note(text: str):
    st.markdown(f'<div class="small-note">{text}</div>', unsafe_allow_html=True)
