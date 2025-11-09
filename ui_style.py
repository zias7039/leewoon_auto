# ui_style.py
import streamlit as st

EXCEL_GREEN = "#217346"   # Excel signature green
WORD_BLUE   = "#185ABD"   # Word signature blue

BASE_CSS = r"""
/* ========== 공통 레이아웃 ========== */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
.block-container { padding-top: 1.2rem; max-width: 1200px; }

/* 버튼 */
.stButton>button{
  height:44px;border-radius:10px;font-weight:500;transition:all .2s ease;
}
.stButton>button:hover{ transform:translateY(-2px); box-shadow:0 4px 12px rgba(0,0,0,.15); }

/* 다운로드 버튼 */
[data-testid="stDownloadButton"]>button{ min-width:220px; }

/* Form 카드 */
[data-testid="stForm"]{
  background:rgba(248,250,252,.5);
  border:1px solid rgba(226,232,240,.8);
  border-radius:16px;padding:24px;
}

/* 텍스트 입력 */
input[type="text"]{
  border-radius:8px!important;border:1px solid rgba(203,213,225,.8)!important;padding:10px 12px!important;
}
input[type="text"]:focus{
  border-color:rgba(59,130,246,.5)!important; box-shadow:0 0 0 3px rgba(59,130,246,.1)!important;
}

/* 소제목 */
.h4{ font-weight:700;font-size:1.05rem;margin:.25rem 0 .75rem;color:rgba(15,23,42,.9); }

/* 업로더 기본형(테마 적용 전) */
[data-testid="stFileUploaderDropzone"]{
  background:transparent!important;
  border:2px dashed rgba(148,163,184,.35)!important;
  border-radius:12px!important; padding:32px 24px!important; min-height:140px!important;
  transition:all .25s ease!important;
}
[data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(148,163,184,.8)!important; transform:translateY(-1px)!important;
}
[data-testid="stFileUploader"] button{
  border-radius:10px!important; padding:8px 20px!important; font-weight:600!important;
  transition:all .2s ease!important;
}

/* ========== 라이트 모드 테마 ========== */
.excel-uploader [data-testid="stFileUploaderDropzone"]{
  border-color:rgba(33,115,70,.65)!important;
  background:linear-gradient(135deg, rgba(33,115,70,.08), rgba(33,115,70,.15))!important;
}
.excel-uploader [data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(33,115,70,.9)!important;
  background:linear-gradient(135deg, rgba(33,115,70,.15), rgba(33,115,70,.25))!important;
  box-shadow:0 6px 24px rgba(33,115,70,.22)!important;
}
.excel-uploader button{
  background:linear-gradient(135deg, #217346, #1a5c38)!important;
  border:1px solid rgba(33,115,70,.85)!important; color:#fff!important;
}

.word-uploader [data-testid="stFileUploaderDropzone"]{
  border-color:rgba(24,90,189,.65)!important;
  background:linear-gradient(135deg, rgba(24,90,189,.08), rgba(24,90,189,.15))!important;
}
.word-uploader [data-testid="stFileUploaderDropzone"]:hover{
  border-color:rgba(24,90,189,.9)!important;
  background:linear-gradient(135deg, rgba(24,90,189,.15), rgba(24,90,189,.25))!important;
  box-shadow:0 6px 24px rgba(24,90,189,.22)!important;
}
.word-uploader button{
  background:linear-gradient(135deg, #185ABD, #1349a0)!important;
  border:1px solid rgba(24,90,189,.85)!important; color:#fff!important;
}

/* ========== 다크 모드 오버라이드 ========== */
@media (prefers-color-scheme: dark){
  [data-testid="stForm"]{
    background:rgba(30,41,59,.40); border-color:rgba(51,65,85,.60);
  }
  .h4{ color:rgba(248,250,252,.9); }

  .excel-uploader [data-testid="stFileUploaderDropzone"]{
    border-color:rgba(33,115,70,.85)!important;
    background:linear-gradient(135deg, rgba(33,115,70,.16), rgba(33,115,70,.26))!important;
  }
  .excel-uploader [data-testid="stFileUploaderDropzone"]:hover{
    background:linear-gradient(135deg, rgba(33,115,70,.22), rgba(33,115,70,.32))!important;
  }

  .word-uploader [data-testid="stFileUploaderDropzone"]{
    border-color:rgba(24,90,189,.85)!important;
    background:linear-gradient(135deg, rgba(24,90,189,.16), rgba(24,90,189,.26))!important;
  }
  .word-uploader [data-testid="stFileUploaderDropzone"]:hover{
    background:linear-gradient(135deg, rgba(24,90,189,.22), rgba(24,90,189,.32))!important;
  }
}

/* 작은 메모 */
.small-note{ font-size:.85rem;color:rgba(100,116,139,.8); }

/* 스크롤바 */
::-webkit-scrollbar{ width:8px;height:8px; }
::-webkit-scrollbar-track{ background:rgba(241,245,249,.5); border-radius:4px; }
::-webkit-scrollbar-thumb{ background:rgba(148,163,184,.5); border-radius:4px; }
::-webkit-scrollbar-thumb:hover{ background:rgba(100,116,139,.7); }
"""

JS_ATTACH_CLASSES = r"""
<script>
(function(){
  function tagUploaders(){
    // 모든 업로더를 수집
    const all = Array.from(document.querySelectorAll('[data-testid="stFileUploader"]'));
    if (!all.length) return;

    // 같은 Form 안에서 첫 번째/두 번째만 테마 적용
    // (진영님 코드처럼 form("input_form") 안에 업로더가 2개라는 전제)
    const forms = Array.from(document.querySelectorAll('form'));
    forms.forEach(form => {
      const ups = Array.from(form.querySelectorAll('[data-testid="stFileUploader"]'));
      if (ups.length >= 1) {
        ups[0].classList.add('excel-uploader');
        ups[0].classList.remove('word-uploader');
      }
      if (ups.length >= 2) {
        ups[1].classList.add('word-uploader');
        ups[1].classList.remove('excel-uploader');
      }
    });
  }

  // 최초 실행
  tagUploaders();

  // DOM 변경 감지(스트림릿 재실행/리렌더 대응)
  const mo = new MutationObserver(() => tagUploaders());
  mo.observe(document.body, {subtree:true, childList:true});

  // 간헐적 렌더 타이밍 보정용 안전망
  let ticks = 0;
  const iv = setInterval(() => {
    tagUploaders();
    if (++ticks > 20) clearInterval(iv); // 약 2초 동안만
  }, 100);
})();
</script>
"""

def inject():
    """CSS + JS 주입"""
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)
    st.markdown(JS_ATTACH_CLASSES, unsafe_allow_html=True)

def h4(text: str):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)

def small_note(text: str):
    st.markdown(f'<div class="small-note">{text}</div>', unsafe_allow_html=True)
