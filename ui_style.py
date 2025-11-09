import streamlit as st

EXCEL_GREEN = "#217346"   # Excel signature green
WORD_BLUE   = "#185ABD"   # Word signature blue

BASE_CSS = """
/* 기본 레이아웃 */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container {
    padding-top: 1.2rem;
    max-width: 1200px;
}

/* 버튼 스타일 */
.stButton>button {
    height: 44px;
    border-radius: 10px;
    font-weight: 500;
    transition: all 0.2s ease;
}

.stButton>button:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
}

[data-testid="stDownloadButton"] > button {
    min-width: 220px;
}

/* 폼 스타일 */
[data-testid="stForm"] {
    background: rgba(248, 250, 252, 0.5);
    border: 1px solid rgba(226, 232, 240, 0.8);
    border-radius: 16px;
    padding: 24px;
}

/* 텍스트 입력 */
input[type="text"] {
    border-radius: 8px !important;
    border: 1px solid rgba(203, 213, 225, 0.8) !important;
    padding: 10px 12px !important;
}

input[type="text"]:focus {
    border-color: rgba(59, 130, 246, 0.5) !important;
    box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1) !important;
}

/* 파일 업로더 공통 스타일 */
[data-testid="stFileUploaderDropzone"] {
    border-radius: 12px !important;
    padding: 32px 24px !important;
    transition: all 0.25s ease !important;
}

/* ========================= */
/*  Excel 업로더 (초록 테마) */
/* ========================= */
/* ===== 업로더를 클래스 기반으로 정확히 타겟팅 ===== */
.excel-uploader [data-testid="stFileUploaderDropzone"] {
    border: 2px dashed rgba(33, 115, 70, 0.6) !important;
    background: linear-gradient(135deg, rgba(33,115,70,0.08) 0%, rgba(33,115,70,0.15) 100%) !important;
}
.excel-uploader [data-testid="stFileUploaderDropzone"]:hover {
    border-color: rgba(33,115,70,0.9) !important;
    background: linear-gradient(135deg, rgba(33,115,70,0.15) 0%, rgba(33,115,70,0.25) 100%) !important;
    box-shadow: 0 6px 24px rgba(33,115,70,0.25) !important;
}
.excel-uploader [data-testid="stFileUploaderDropzone"] p,
.excel-uploader [data-testid="stFileUploaderDropzone"] span { 
    color: rgba(33,115,70,1) !important; font-weight: 600 !important; 
}
.excel-uploader [data-testid="stFileUploaderDropzone"] small { 
    color: rgba(33,115,70,0.75) !important; 
}
.excel-uploader button {
    background: linear-gradient(135deg, #217346 0%, #1a5c38 100%) !important;
    border: 1px solid rgba(33,115,70,0.8) !important;
    color: #fff !important; font-weight: 600 !important;
}
.excel-uploader button:hover {
    background: linear-gradient(135deg, #25824f 0%, #1e6841 100%) !important;
    box-shadow: 0 4px 16px rgba(33,115,70,0.35) !important;
}

/* ===== Word 전용 테마 (두 번째 업로더) - 클래스 기반 ===== */
.word-uploader [data-testid="stFileUploaderDropzone"] {
    border: 2px dashed rgba(24, 90, 189, 0.6) !important;
    background: linear-gradient(135deg, rgba(24,90,189,0.08) 0%, rgba(24,90,189,0.15) 100%) !important;
}
.word-uploader [data-testid="stFileUploaderDropzone"]:hover {
    border-color: rgba(24,90,189,0.9) !important;
    background: linear-gradient(135deg, rgba(24,90,189,0.15) 0%, rgba(24,90,189,0.25) 100%) !important;
    box-shadow: 0 6px 24px rgba(24,90,189,0.25) !important;
}
.word-uploader [data-testid="stFileUploaderDropzone"] p,
.word-uploader [data-testid="stFileUploaderDropzone"] span { 
    color: rgba(24,90,189,1) !important; font-weight: 600 !important; 
}
.word-uploader [data-testid="stFileUploaderDropzone"] small { 
    color: rgba(24,90,189,0.75) !important; 
}
.word-uploader button {
    background: linear-gradient(135deg, #185ABD 0%, #1349a0 100%) !important;
    border: 1px solid rgba(24,90,189,0.8) !important;
    color: #fff !important; font-weight: 600 !important;
}
.word-uploader button:hover {
    background: linear-gradient(135deg, #1c66d1 0%, #1552b3 100%) !important;
    box-shadow: 0 4px 16px rgba(24,90,189,0.35) !important;
}


def inject():
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

    st.markdown("""
    <script>
    setTimeout(function() {
        const uploaders = document.querySelectorAll('[data-testid="stFileUploader"]');
        if (uploaders.length >= 2) {
            uploaders[0].classList.add('excel-uploader');
            uploaders[1].classList.add('word-uploader');
        }
    }, 150);
    </script>
    """, unsafe_allow_html=True)

def h4(text):
    st.markdown(f'<div style="font-weight:600; margin-top:8px;">{text}</div>', unsafe_allow_html=True)
