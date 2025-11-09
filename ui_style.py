# ui_style.py
import streamlit as st

EXCEL_GREEN = "#217346"   # Excel signature green
WORD_BLUE   = "#185ABD"   # Word signature blue

BASE_CSS = """
/* ===== 기본 레이아웃 ===== */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
.block-container { padding-top: 1.2rem; max-width: 1200px; }

/* 버튼 */
.stButton>button {
  height: 44px; border-radius: 10px; font-weight: 500; transition: all .2s ease;
}
.stButton>button:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,.15); }

/* 다운로드 */
[data-testid="stDownloadButton"] > button { min-width: 220px; }

/* 폼 카드 */
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

/* 섹션 제목/메모 */
.h4{ font-weight:700; font-size:1.05rem; margin:.25rem 0 .75rem; color:rgba(15,23,42,.9); }
.small-note{ font-size:.85rem; color:rgba(100,116,139,.8); }

/* ===== 업로더 공통 ===== */
[data-testid="stFileUploader"] { animation: fadeIn .25s ease-out; }
[data-testid="stFileUploader"] section { gap: 10px !important; }
[data-testid="stFileUploader"] button {
  border-radius: 8px !important; padding: 8px 20px !important; font-weight: 600 !important;
  transition: all .2s ease !important;
}
[data-testid="stFileUploaderDropzone"]{
  background: rgba(248,250,252,.6) !important;
  border: 2px dashed rgba(203,213,225,.6) !important;
  border-radius: 12px !important; padding: 32px 24px !important; min-height: 140px !important;
  transition: all .3s cubic-bezier(.4,0,.2,1) !important;
}
[data-testid="stFileUploaderDropzone"]:hover{
  border-color: rgba(148,163,184,.8) !important;
  background: rgba(241,245,249,.8) !important;
  transform: translateY(-2px) !important; box-shadow: 0 4px 12px rgba(0,0,0,.08) !important;
}
[data-testid="stFileUploaderDropzone"] p{ color: rgba(71,85,105,.9) !important; font-size:.95rem !important; }
[data-testid="stFileUploaderDropzone"] small{ color: rgba(100,116,139,.7) !important; font-size:.85rem !important; }

/* ===== Excel 테마: .excel-uploader ===== */
.excel-uploader [data-testid="stFileUploaderDropzone"]{
  border-color: rgba(33,115,70,.6) !important;
  background: linear-gradient(135deg, rgba(33,115,70,.08), rgba(33,115,70,.15)) !important;
}
.excel-uploader [data-testid="stFileUploaderDropzone"]:hover{
  border-color: rgba(33,115,70,.9) !important;
  background: linear-gradient(135deg, rgba(33,115,70,.15), rgba(33,115,70,.25)) !important;
  box-shadow: 0 6px 24px rgba(33,115,70,.25) !important;
}
.excel-uploader [data-testid="stFileUploaderDropzone"] p,
.excel-uploader [data-testid="stFileUploaderDropzone"] span{ color: rgba(33,115,70,1) !important; font-weight:600 !important; }
.excel-uploader [data-testid="stFileUploaderDropzone"] small{ color: rgba(33,115,70,.75) !important; }
.excel-uploader button{
  background: linear-gradient(135deg, #217346, #1a5c38) !important;
  border: 1px solid rgba(33,115,70,.8) !important; color: #fff !important;
}
.excel-uploader button:hover{
  background: linear-gradient(135deg, #25824f, #1e6841) !important;
  box-shadow: 0 4px 16px rgba(33,115,70,.35) !important;
}

/* ===== Word 테마: .word-uploader ===== */
.word-uploader [data-testid="stFileUploaderDropzone"]{
  border-color: rgba(24,90,189,.6) !important;
  background: linear-gradient(135deg, rgba(24,90,189,.08), rgba(24,90,189,.15)) !important;
}
.word-uploader [data-testid="stFileUploaderDropzone"]:hover{
  border-color: rgba(24,90,189,.9) !important;
  background: linear-gradient(135deg, rgba(24,90,189,.15), rgba(24,90,189,.25)) !important;
  box-shadow: 0 6px 24px rgba(24,90,189,.25) !important;
}
.word-uploader [data-testid="stFileUploaderDropzone"] p,
.word-uploader [data-testid="stFileUploaderDropzone"] span{ color: rgba(24,90,189,1) !important; font-weight:600 !important; }
.word-uploader [data-testid="stFileUploaderDropzone"] small{ color: rgba(24,90,189,.75) !important; }
.word-uploader button{
  background: linear-gradient(135deg, #185ABD, #1349a0) !important;
  border: 1px solid rgba(24,90,189,.8) !important; color: #fff !important;
}
.word-uploader button:hover{
  background: linear-gradient(135deg, #1c66d1, #1552b3) !important;
  box-shadow: 0 4px 16px rgba(24,90,189,.35) !important;
}

/* ===== 다크모드(색 보존) ===== */
@media (prefers-color-scheme: dark){
  [data-testid="stForm"]{ background: rgba(30,41,59,.4); border-color: rgba(51,65,85,.6); }
  /* 기본값은 어둡게, 개별 테마는 위 규칙이 그대로 덮어씀 */
  [data-testid="stFileUploaderDropzone"]{
    background: rgba(30,41,59,.35) !important; border-color: rgba(71,85,105,.45) !important;
  }
  [data-testid="stFileUploaderDropzone"]:hover{
    background: rgba(30,41,59,.55) !important; border-color: rgba(100,116,139,.65) !important;
  }
  .h4{ color: rgba(248,250,252,.9); }
}

/* 애니메이션 & 스크롤바 */
@keyframes fadeIn{ from{opacity:0; transform: translateY(10px);} to{opacity:1; transform: translateY(0);} }
::-webkit-scrollbar{ width:8px; height:8px; }
::-webkit-scrollbar-track{ background: rgba(241,245,249,.5); border-radius:4px; }
::-webkit-scrollbar-thumb{ background: rgba(148,163,184,.5); border-radius:4px; }
::-webkit-scrollbar-thumb:hover{ background: rgba(100,116,139,.7); }
"""

def inject():
    # CSS 먼저 주입
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

    # 업로더 DOM이 렌더링될 때까지 기다린 뒤 클래스 부착
    st.markdown("""
    <script>
    (function() {
      function tagUploaders() {
        const list = document.querySelectorAll('[data-testid="stFileUploader"]');
        if (list && list.length >= 2) {
          list.forEach(el => { el.classList.remove('excel-uploader','word-uploader'); });
          list[0].classList.add('excel-uploader');
          list[1].classList.add('word-uploader');
        }
      }
      // 최초 시도 + 약간의 지연
      setTimeout(tagUploaders, 120);

      // DOM 변경 시 재태깅
      const obs = new MutationObserver(() => tagUploaders());
      obs.observe(document.body, {childList:true, subtree:true});
    })();
    </script>
    """, unsafe_allow_html=True)

def h4(text: str):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)

def small_note(text: str):
    st.markdown(f'<div class="small-note">{text}</div>', unsafe_allow_html=True)
