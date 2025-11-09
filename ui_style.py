# ui_style.py
import streamlit as st

EXCEL_GREEN = "#217346"   # Excel signature green
WORD_BLUE   = "#185ABD"   # Word signature blue

BASE_CSS = """
/* -------------------- 공통 레이아웃 -------------------- */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
.block-container { padding-top: 1.2rem; max-width: 1200px; }

/* 버튼 */
.stButton>button{
  height:44px;border-radius:10px;font-weight:500;transition:all .2s ease;
}
.stButton>button:hover{ transform:translateY(-2px); box-shadow:0 4px 12px rgba(0,0,0,.15); }
[data-testid="stDownloadButton"]>button{ min-width:220px; }

/* Form 카드 */
[data-testid="stForm"]{
  background:rgba(248,250,252,.5);
  border:1px solid rgba(226,232,240,.8);
  border-radius:16px; padding:24px;
}

/* 입력 */
input[type="text"]{
  border-radius:8px !important; border:1px solid rgba(203,213,225,.8) !important; padding:10px 12px !important;
}
input[type="text"]:focus{
  border-color:rgba(59,130,246,.5) !important; box-shadow:0 0 0 3px rgba(59,130,246,.1) !important;
}

/* 업로더 기본(테마 적용 전의 기본값) */
[data-testid="stFileUploaderDropzone"]{
  background:rgba(248,250,252,.6) !important;
  border:2px dashed rgba(203,213,225,.6) !important;
  border-radius:12px !important; padding:32px 24px !important;
  transition:all .3s cubic-bezier(.4,0,.2,1) !important; min-height:140px !important;
}
[data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(148,163,184,.8) !important; background:rgba(241,245,249,.8) !important;
  transform:translateY(-2px) !important; box-shadow:0 4px 12px rgba(0,0,0,.08) !important;
}
[data-testid="stFileUploader"] section{ gap:10px !important; }
[data-testid="stFileUploader"] button{
  border-radius:8px !important; padding:8px 20px !important; font-weight:500 !important; transition:all .2s ease !important;
}
[data-testid="stFileUploader"] button:hover{ transform:translateY(-1px) !important; box-shadow:0 2px 8px rgba(0,0,0,.15) !important; }

/* -------------------- 다크모드 공통 오버라이드 -------------------- */
@media (prefers-color-scheme: dark){
  [data-testid="stForm"]{
    background:rgba(30,41,59,.40); border-color:rgba(51,65,85,.60);
  }
  /* 기본 드롭존(테마가 덮어씀) */
  [data-testid="stFileUploaderDropzone"]{
    background:rgba(30,41,59,.40) !important; border-color:rgba(71,85,105,.50) !important;
  }
  [data-testid="stFileUploaderDropzone"]:hover{
    background:rgba(30,41,59,.60) !important; border-color:rgba(100,116,139,.70) !important;
  }
}

/* =========================================================
   테마 색 (가장 마지막에 둬서 무엇이 됐든 이 규칙이 이기도록)
   클래스 기반(.excel-uploader / .word-uploader) + !important
   ========================================================= */

/* Excel 테마 (라이트/다크 공통) */
[data-testid="stForm"] .excel-uploader [data-testid="stFileUploaderDropzone"]{
  border:2px dashed rgba(33,115,70,.60) !important;
  background:linear-gradient(135deg, rgba(33,115,70,.08) 0%, rgba(33,115,70,.18) 100%) !important;
}
[data-testid="stForm"] .excel-uploader [data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(33,115,70,.95) !important;
  background:linear-gradient(135deg, rgba(33,115,70,.18) 0%, rgba(33,115,70,.28) 100%) !important;
  box-shadow:0 6px 24px rgba(33,115,70,.25) !important;
}
[data-testid="stForm"] .excel-uploader button{
  background:linear-gradient(135deg, #217346 0%, #1a5c38 100%) !important;
  border:1px solid rgba(33,115,70,.85) !important; color:#fff !important; font-weight:600 !important;
}

/* Word 테마 (라이트/다크 공통) */
[data-testid="stForm"] .word-uploader [data-testid="stFileUploaderDropzone"]{
  border:2px dashed rgba(24,90,189,.60) !important;
  background:linear-gradient(135deg, rgba(24,90,189,.08) 0%, rgba(24,90,189,.18) 100%) !important;
}
[data-testid="stForm"] .word-uploader [data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(24,90,189,.95) !important;
  background:linear-gradient(135deg, rgba(24,90,189,.18) 0%, rgba(24,90,189,.28) 100%) !important;
  box-shadow:0 6px 24px rgba(24,90,189,.25) !important;
}
[data-testid="stForm"] .word-uploader button{
  background:linear-gradient(135deg, #185ABD 0%, #1349a0 100%) !important;
  border:1px solid rgba(24,90,189,.85) !important; color:#fff !important; font-weight:600 !important;
}

/* (안전장치) 포지션 기반 백업: 폼 안 첫 번째/두 번째 업로더 */
[data-testid="stForm"] [data-testid="stFileUploader"]:first-of-type [data-testid="stFileUploaderDropzone"]{
  border-color:rgba(33,115,70,.60) !important;
  background:linear-gradient(135deg, rgba(33,115,70,.08), rgba(33,115,70,.18)) !important;
}
[data-testid="stForm"] [data-testid="stFileUploader"]:nth-of-type(2) [data-testid="stFileUploaderDropzone"]{
  border-color:rgba(24,90,189,.60) !important;
  background:linear-gradient(135deg, rgba(24,90,189,.08), rgba(24,90,189,.18)) !important;
}
"""

def inject():
    # CSS 먼저
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

    # 업로더마다 클래스를 안정적으로 부착
    st.markdown(
        """
        <script>
        (function(){
          function tagUploaders(){
            var forms = document.querySelectorAll('[data-testid="stForm"]');
            forms.forEach(function(f){
              var ups = f.querySelectorAll('[data-testid="stFileUploader"]');
              if(ups.length){
                // 순서대로 지정
                if(ups[0]) ups[0].classList.add('excel-uploader');
                if(ups[1]) ups[1].classList.add('word-uploader');
              }
            });
          }
          // 최초 시도
          setTimeout(tagUploaders, 120);
          // 동적 변경 감지
          var obs = new MutationObserver(function(){ setTimeout(tagUploaders, 50); });
          obs.observe(document.body, {childList:true, subtree:true});
        })();
        </script>
        """,
        unsafe_allow_html=True,
    )

def h4(text: str):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)
