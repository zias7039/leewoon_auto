# ui_style.py
import streamlit as st

EXCEL_GREEN = "#217346"   # Excel signature green
WORD_BLUE   = "#185ABD"   # Word signature blue

BASE_CSS = f"""
/* --- 공통 레이아웃 --- */
#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}
.block-container {{ padding-top: 1.2rem; max-width: 1200px; }}

/* 버튼 */
.stButton>button {{
  height: 44px; border-radius: 10px; font-weight: 500; transition: .2s;
}}
.stButton>button:hover {{
  transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,.15);
}}
[data-testid="stDownloadButton"]>button {{ min-width: 220px; }}

/* Form */
[data-testid="stForm"] {{
  background: rgba(248,250,252,.5);
  border: 1px solid rgba(226,232,240,.8);
  border-radius: 16px; padding: 24px;
}}

/* 입력 */
input[type="text"] {{
  border-radius: 8px !important;
  border: 1px solid rgba(203,213,225,.8) !important;
  padding: 10px 12px !important;
}}
input[type="text"]:focus {{
  border-color: rgba(59,130,246,.5) !important;
  box-shadow: 0 0 0 3px rgba(59,130,246,.1) !important;
}}

/* 업로더 공통 */
[data-testid="stFileUploaderDropzone"] {{
  border: 2px dashed rgba(203,213,225,.6) !important;
  border-radius: 12px !important;
  padding: 32px 24px !important;
  min-height: 140px !important;
  transition: .3s cubic-bezier(.4,0,.2,1) !important;
}}
[data-testid="stFileUploaderDropzone"]:hover {{
  border-color: rgba(148,163,184,.8) !important;
  background: rgba(241,245,249,.8) !important;
  transform: translateY(-2px) !important;
}}
[data-testid="stFileUploader"] section {{ gap: 10px !important; }}
[data-testid="stFileUploader"] button {{
  border-radius: 8px !important; padding: 8px 20px !important; font-weight: 600 !important;
}}

/* ───────────────────────────────────────────────────────────── */
/*  Excel / Word 테마 (업로더 노드에 붙는 클래스 기준)           */
/*  - JS가 [data-testid=stFileUploader]에 excel-uploader /       */
/*    word-uploader 클래스를 강제로 부여합니다.                  */
/*  - 다크모드까지 확실히 덮어쓰도록 강한 선택자 + !important    */
/* ───────────────────────────────────────────────────────────── */

/* Excel (녹색) */
[data-testid="stForm"] .excel-uploader [data-testid="stFileUploaderDropzone"],
.excel-uploader [data-testid="stFileUploaderDropzone"] {{
  border-color: rgba(33,115,70,.75) !important;
  background: linear-gradient(135deg, rgba(33,115,70,.10) 0%, rgba(33,115,70,.18) 100%) !important;
}}
[data-testid="stForm"] .excel-uploader [data-testid="stFileUploaderDropzone"]:hover,
.excel-uploader [data-testid="stFileUploaderDropzone"]:hover {{
  border-color: rgba(33,115,70,.95) !important;
  background: linear-gradient(135deg, rgba(33,115,70,.18) 0%, rgba(33,115,70,.28) 100%) !important;
  box-shadow: 0 6px 24px rgba(33,115,70,.28) !important;
}}
[data-testid="stForm"] .excel-uploader button,
.excel-uploader button {{
  background: linear-gradient(135deg, {EXCEL_GREEN} 0%, #1a5c38 100%) !important;
  border: 1px solid rgba(33,115,70,.9) !important;
  color: #fff !important;
}}

/* Word (파랑) */
[data-testid="stForm"] .word-uploader [data-testid="stFileUploaderDropzone"],
.word-uploader [data-testid="stFileUploaderDropzone"] {{
  border-color: rgba(24,90,189,.75) !important;
  background: linear-gradient(135deg, rgba(24,90,189,.10) 0%, rgba(24,90,189,.18) 100%) !important;
}}
[data-testid="stForm"] .word-uploader [data-testid="stFileUploaderDropzone"]:hover,
.word-uploader [data-testid="stFileUploaderDropzone"]:hover {{
  border-color: rgba(24,90,189,.95) !important;
  background: linear-gradient(135deg, rgba(24,90,189,.18) 0%, rgba(24,90,189,.28) 100%) !important;
  box-shadow: 0 6px 24px rgba(24,90,189,.28) !important;
}}
[data-testid="stForm"] .word-uploader button,
.word-uploader button {{
  background: linear-gradient(135deg, {WORD_BLUE} 0%, #1349a0 100%) !important;
  border: 1px solid rgba(24,90,189,.9) !important;
  color: #fff !important;
}}

/* 다크모드에서도 색 보존(배경이 씌워도 우리가 이김) */
@media (prefers-color-scheme: dark) {{
  [data-testid="stForm"] {{
    background: rgba(30,41,59,.4);
    border-color: rgba(51,65,85,.6);
  }}
  /* Excel 다크 */
  [data-testid="stForm"] .excel-uploader [data-testid="stFileUploaderDropzone"],
  .excel-uploader [data-testid="stFileUploaderDropzone"] {{
    background: linear-gradient(135deg, rgba(33,115,70,.18), rgba(33,115,70,.28)) !important;
    border-color: rgba(33,115,70,.9) !important;
  }}
  /* Word 다크 */
  [data-testid="stForm"] .word-uploader [data-testid="stFileUploaderDropzone"],
  .word-uploader [data-testid="stFileUploaderDropzone"] {{
    background: linear-gradient(135deg, rgba(24,90,189,.18), rgba(24,90,189,.28)) !important;
    border-color: rgba(24,90,189,.9) !important;
  }}
}}
"""

JS_ATTACH_CLASSES = """
<script>
(function(){
  function tagUploaders(){
    const ups = document.querySelectorAll('[data-testid="stFileUploader"]');
    if(!ups || ups.length === 0) return;
    // 첫 업로더 = Excel, 두 번째 = Word
    ups.forEach((el, idx) => {
      el.classList.remove('excel-uploader','word-uploader');
      if(idx === 0) el.classList.add('excel-uploader');
      if(idx === 1) el.classList.add('word-uploader');
    });
  }
  // 최초 시도 + 약간의 지연
  tagUploaders();
  setTimeout(tagUploaders, 150);

  // DOM이 바뀔 때마다 재부착
  const obs = new MutationObserver(() => tagUploaders());
  obs.observe(document.body, {childList: true, subtree: true});
})();
</script>
"""

def inject():
    # CSS 먼저
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)
    # 업로더에 클래스 부착(JS)
    st.markdown(JS_ATTACH_CLASSES, unsafe_allow_html=True)

def h4(text):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)
