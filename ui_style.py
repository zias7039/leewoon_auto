# ui_style.py
import streamlit as st

EXCEL_GREEN = "#217346"   # Excel signature green
WORD_BLUE   = "#185ABD"   # Word signature blue

BASE_CSS = r"""
/* ---- 기본 레이아웃/컨트롤 ---- */
#MainMenu, footer { visibility: hidden; }
.block-container { padding-top: 1.2rem; max-width: 1200px; }

.stButton>button{
  height:44px;border-radius:10px;font-weight:500;transition:.2s;
}
.stButton>button:hover{ transform:translateY(-2px); box-shadow:0 4px 12px rgba(0,0,0,.15); }
[data-testid="stDownloadButton"]>button{ min-width:220px; }

[data-testid="stForm"]{
  background:rgba(248,250,252,.5);
  border:1px solid rgba(226,232,240,.8);
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

/* ---- 업로더 공통 기본 ---- */
[data-testid="stFileUploaderDropzone"]{
  background:rgba(248,250,252,.6)!important;
  border:2px dashed rgba(203,213,225,.6)!important;
  border-radius:12px!important; padding:32px 24px!important;
  min-height:140px!important; transition:all .3s cubic-bezier(.4,0,.2,1)!important;
}
[data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(148,163,184,.8)!important;
  background:rgba(241,245,249,.8)!important;
  transform:translateY(-2px)!important; box-shadow:0 4px 12px rgba(0,0,0,.08)!important;
}
[data-testid="stFileUploader"] section{ gap:10px!important; }
[data-testid="stFileUploader"] button{
  border-radius:8px!important; padding:8px 20px!important; font-weight:500!important; transition:.2s!important;
}
[data-testid="stFileUploader"] button:hover{ transform:translateY(-1px)!important; box-shadow:0 2px 8px rgba(0,0,0,.15)!important; }

/* =================================================================== */
/*  테마 지정 — 클래스 기반 (nth-of-type 사용 안 함)                   */
/* =================================================================== */

/* Excel (초록) */
.excel-uploader [data-testid="stFileUploaderDropzone"]{
  border-color:rgba(33,115,70,.6)!important;
  background:linear-gradient(135deg, rgba(33,115,70,.08) 0%, rgba(33,115,70,.15) 100%)!important;
}
.excel-uploader [data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(33,115,70,.9)!important;
  background:linear-gradient(135deg, rgba(33,115,70,.15) 0%, rgba(33,115,70,.25) 100%)!important;
  box-shadow:0 6px 24px rgba(33,115,70,.25)!important;
}
.excel-uploader [data-testid="stFileUploaderDropzone"] p,
.excel-uploader [data-testid="stFileUploaderDropzone"] span{ color:rgba(33,115,70,1)!important; font-weight:600!important; }
.excel-uploader [data-testid="stFileUploaderDropzone"] small{ color:rgba(33,115,70,.75)!important; }
.excel-uploader button{
  background:linear-gradient(135deg, #217346 0%, #1a5c38 100%)!important;
  border:1px solid rgba(33,115,70,.8)!important; color:#fff!important; font-weight:600!important;
}
.excel-uploader button:hover{
  background:linear-gradient(135deg, #25824f 0%, #1e6841 100%)!important;
  box-shadow:0 4px 16px rgba(33,115,70,.35)!important;
}

/* Word (파랑) */
.word-uploader [data-testid="stFileUploaderDropzone"]{
  border-color:rgba(24,90,189,.6)!important;
  background:linear-gradient(135deg, rgba(24,90,189,.08) 0%, rgba(24,90,189,.15) 100%)!important;
}
.word-uploader [data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(24,90,189,.9)!important;
  background:linear-gradient(135deg, rgba(24,90,189,.15) 0%, rgba(24,90,189,.25) 100%)!important;
  box-shadow:0 6px 24px rgba(24,90,189,.25)!important;
}
.word-uploader [data-testid="stFileUploaderDropzone"] p,
.word-uploader [data-testid="stFileUploaderDropzone"] span{ color:rgba(24,90,189,1)!important; font-weight:600!important; }
.word-uploader [data-testid="stFileUploaderDropzone"] small{ color:rgba(24,90,189,.75)!important; }
.word-uploader button{
  background:linear-gradient(135deg, #185ABD 0%, #1349a0 100%)!important;
  border:1px solid rgba(24,90,189,.8)!important; color:#fff!important; font-weight:600!important;
}
.word-uploader button:hover{
  background:linear-gradient(135deg, #1c66d1 0%, #1552b3 100%)!important;
  box-shadow:0 4px 16px rgba(24,90,189,.35)!important;
}

/* ---- 상태/익스팬더/스크롤바 ---- */
.stAlert{ border-radius:12px; border-left-width:4px; }
[data-testid="stStatusWidget"]{ border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,.1); }
[data-testid="stExpander"]{ border-radius:10px; border:1px solid rgba(226,232,240,.8); }
[data-testid="stExpander"] summary{ border-radius:10px; }
[data-testid="column"]{ padding:0 8px; }

/* ---- 다크모드 오버라이드 (색감 유지) ---- */
@media (prefers-color-scheme: dark){
  [data-testid="stForm"]{ background:rgba(30,41,59,.4); border-color:rgba(51,65,85,.6); }
  [data-testid="stFileUploaderDropzone"]{
    background:rgba(30,41,59,.4)!important; border-color:rgba(71,85,105,.5)!important;
  }
  [data-testid="stFileUploaderDropzone"]:hover{
    background:rgba(30,41,59,.6)!important; border-color:rgba(100,116,139,.7)!important;
  }
  .h4{ color:rgba(248,250,252,.9); }

  /* 테마 색상 유지 (다크에서도) */
  .excel-uploader [data-testid="stFileUploaderDropzone"]{
    background:linear-gradient(135deg, rgba(33,115,70,.18), rgba(33,115,70,.28))!important;
    border-color:rgba(33,115,70,.85)!important;
  }
  .excel-uploader [data-testid="stFileUploaderDropzone"]:hover{
    background:linear-gradient(135deg, rgba(33,115,70,.25), rgba(33,115,70,.35))!important;
  }
  .excel-uploader button{
    background:linear-gradient(135deg, #217346, #1a5c38)!important; border-color:rgba(33,115,70,.9)!important; color:#fff!important;
  }

  .word-uploader [data-testid="stFileUploaderDropzone"]{
    background:linear-gradient(135deg, rgba(24,90,189,.18), rgba(24,90,189,.28))!important;
    border-color:rgba(24,90,189,.85)!important;
  }
  .word-uploader [data-testid="stFileUploaderDropzone"]:hover{
    background:linear-gradient(135deg, rgba(24,90,189,.25), rgba(24,90,189,.35))!important;
  }
  .word-uploader button{
    background:linear-gradient(135deg, #185ABD, #1349a0)!important; border-color:rgba(24,90,189,.9)!important; color:#fff!important;
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

JS_ATTACH_CLASSES = """
<script>
(function(){
  function tagUploaders(){
    const nodes = document.querySelectorAll('[data-testid="stFileUploader"]');
    if(!nodes || nodes.length < 2) return;
    // 첫 번째(엑셀), 두 번째(워드)에 명시적으로 클래스 부착
    nodes[0].classList.add('excel-uploader');
    nodes[1].classList.add('word-uploader');
  }
  // 초기 시도 + 재렌더 감지
  setTimeout(tagUploaders, 100);
  const obs = new MutationObserver(() => setTimeout(tagUploaders, 50));
  obs.observe(document.body, {childList:true, subtree:true});
})();
</script>
"""

def inject():
    st.markdown("<style>:root{--excel-green:%s;--word-blue:%s;}</style>" % (EXCEL_GREEN, WORD_BLUE),
                unsafe_allow_html=True)
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)
    st.markdown(JS_ATTACH_CLASSES, unsafe_allow_html=True)

def h4(text: str):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)
