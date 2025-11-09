# ui_style.py
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

/* 다운로드 버튼 */
[data-testid="stDownloadButton"] > button {
    min-width: 220px;
}

/* Form 스타일 */
[data-testid="stForm"] {
    background: rgba(248,250,252,0.5);
    border: 1px solid rgba(226,232,240,0.8);
    border-radius: 16px;
    padding: 24px;
}

/* 텍스트 입력 */
input[type="text"] {
    border-radius: 8px !important;
    border: 1px solid rgba(203,213,225,0.8) !important;
    padding: 10px 12px !important;
}
input[type="text"]:focus {
    border-color: rgba(59,130,246,0.5) !important;
    box-shadow: 0 0 0 3px rgba(59,130,246,0.1) !important;
}

/* 기본 업로더 공통 스타일 */
[data-testid="stFileUploaderDropzone"] {
    border: 2px dashed rgba(203,213,225,0.6) !important;
    border-radius: 12px !important;
    padding: 32px 24px !important;
    transition: all 0.3s ease !important;
    min-height: 140px !important;
}
[data-testid="stFileUploaderDropzone"]:hover {
    border-color: rgba(148,163,184,0.8) !important;
    background: rgba(241,245,249,0.8) !important;
    transform: translateY(-2px) !important;
}

/* ==================================================================== */
/* 다크모드 전용 오버라이드 — 색 보존                                   */
/* ==================================================================== */
@media (prefers-color-scheme: dark) {

    /* Excel */
    .excel-uploader [data-testid="stFileUploaderDropzone"] {
        background: linear-gradient(135deg, rgba(33,115,70,0.18), rgba(33,115,70,0.28)) !important;
        border-color: rgba(33,115,70,0.85) !important;
    }
    .excel-uploader [data-testid="stFileUploaderDropzone"]:hover {
        background: linear-gradient(135deg, rgba(33,115,70,0.25), rgba(33,115,70,0.35)) !important;
    }
    .excel-uploader button {
        background: linear-gradient(135deg, #217346, #1a5c38) !important;
        border-color: rgba(33,115,70,0.9) !important;
        color: #fff !important;
    }

    /* Word */
    .word-uploader [data-testid="stFileUploaderDropzone"] {
        background: linear-gradient(135deg, rgba(24,90,189,0.18), rgba(24,90,189,0.28)) !important;
        border-color: rgba(24,90,189,0.85) !important;
    }
    .word-uploader [data-testid="stFileUploaderDropzone"]:hover {
        background: linear-gradient(135deg, rgba(24,90,189,0.25), rgba(24,90,189,0.35)) !important;
    }
    .word-uploader button {
        background: linear-gradient(135deg, #185ABD, #1349a0) !important;
        border-color: rgba(24,90,189,0.9) !important;
        color: #fff !important;
    }
}
"""

def inject():
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

    # 업로더에 클래스 부착 (가장 중요)
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
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)
