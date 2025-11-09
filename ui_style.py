# ui_style.py
import streamlit as st

EXCEL_GREEN = "#217346"   # Excel signature green
WORD_BLUE   = "#185ABD"   # Word blue

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

/* 셀렉트박스 */
[data-baseweb="select"] {
    border-radius: 8px;
}

/* 소제목 */
.h4 {
    font-weight: 700;
    font-size: 1.05rem;
    margin: 0.25rem 0 0.75rem;
    color: rgba(15, 23, 42, 0.9);
}

/* 작은 노트 */
.small-note {
    font-size: 0.85rem;
    color: rgba(100, 116, 139, 0.8);
}

/* ===== 파일 업로더 기본 스타일 ===== */
[data-testid="stFileUploaderDropzone"] {
    background: rgba(248, 250, 252, 0.6) !important;
    border: 2px dashed rgba(203, 213, 225, 0.6) !important;
    border-radius: 12px !important;
    padding: 32px 24px !important;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    min-height: 140px !important;
}

[data-testid="stFileUploaderDropzone"]:hover {
    border-color: rgba(148, 163, 184, 0.8) !important;
    background: rgba(241, 245, 249, 0.8) !important;
    transform: translateY(-2px) !important;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08) !important;
}

[data-testid="stFileUploaderDropzone"] p {
    color: rgba(71, 85, 105, 0.9) !important;
    font-size: 0.95rem !important;
}

[data-testid="stFileUploaderDropzone"] small {
    color: rgba(100, 116, 139, 0.7) !important;
    font-size: 0.85rem !important;
}

[data-testid="stFileUploader"] section {
    gap: 10px !important;
}

[data-testid="stFileUploader"] button {
    border-radius: 8px !important;
    padding: 8px 20px !important;
    font-weight: 500 !important;
    transition: all 0.2s ease !important;
}

[data-testid="stFileUploader"] button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15) !important;
}

/* ===== Excel 전용 테마 (첫 번째 업로더) ===== */
[data-testid="stForm"] [data-testid="stFileUploader"]:first-of-type [data-testid="stFileUploaderDropzone"] {
    border: 2px dashed rgba(33, 115, 70, 0.6) !important;
    background: linear-gradient(135deg, rgba(33, 115, 70, 0.08) 0%, rgba(33, 115, 70, 0.15) 100%) !important;
}

[data-testid="stForm"] [data-testid="stFileUploader"]:first-of-type [data-testid="stFileUploaderDropzone"]:hover {
    border-color: rgba(33, 115, 70, 0.9) !important;
    background: linear-gradient(135deg, rgba(33, 115, 70, 0.15) 0%, rgba(33, 115, 70, 0.25) 100%) !important;
    box-shadow: 0 6px 24px rgba(33, 115, 70, 0.25) !important;
}

[data-testid="stForm"] [data-testid="stFileUploader"]:first-of-type [data-testid="stFileUploaderDropzone"] p,
[data-testid="stForm"] [data-testid="stFileUploader"]:first-of-type [data-testid="stFileUploaderDropzone"] span {
    color: rgba(33, 115, 70, 1) !important;
    font-weight: 600 !important;
}

[data-testid="stForm"] [data-testid="stFileUploader"]:first-of-type [data-testid="stFileUploaderDropzone"] small {
    color: rgba(33, 115, 70, 0.75) !important;
}

[data-testid="stForm"] [data-testid="stFileUploader"]:first-of-type button {
    background: linear-gradient(135deg, #217346 0%, #1a5c38 100%) !important;
    border: 1px solid rgba(33, 115, 70, 0.8) !important;
    color: white !important;
    font-weight: 600 !important;
}

[data-testid="stForm"] [data-testid="stFileUploader"]:first-of-type button:hover {
    background: linear-gradient(135deg, #25824f 0%, #1e6841 100%) !important;
    box-shadow: 0 4px 16px rgba(33, 115, 70, 0.35) !important;
}

/* ===== Word 전용 테마 (두 번째 업로더) ===== */
[data-testid="stForm"] [data-testid="stFileUploader"]:nth-of-type(2) [data-testid="stFileUploaderDropzone"] {
    border: 2px dashed rgba(24, 90, 189, 0.6) !important;
    background: linear-gradient(135deg, rgba(24, 90, 189, 0.08) 0%, rgba(24, 90, 189, 0.15) 100%) !important;
}

[data-testid="stForm"] [data-testid="stFileUploader"]:nth-of-type(2) [data-testid="stFileUploaderDropzone"]:hover {
    border-color: rgba(24, 90, 189, 0.9) !important;
    background: linear-gradient(135deg, rgba(24, 90, 189, 0.15) 0%, rgba(24, 90, 189, 0.25) 100%) !important;
    box-shadow: 0 6px 24px rgba(24, 90, 189, 0.25) !important;
}

[data-testid="stForm"] [data-testid="stFileUploader"]:nth-of-type(2) [data-testid="stFileUploaderDropzone"] p,
[data-testid="stForm"] [data-testid="stFileUploader"]:nth-of-type(2) [data-testid="stFileUploaderDropzone"] span {
    color: rgba(24, 90, 189, 1) !important;
    font-weight: 600 !important;
}

[data-testid="stForm"] [data-testid="stFileUploader"]:nth-of-type(2) [data-testid="stFileUploaderDropzone"] small {
    color: rgba(24, 90, 189, 0.75) !important;
}

[data-testid="stForm"] [data-testid="stFileUploader"]:nth-of-type(2) button {
    background: linear-gradient(135deg, #185ABD 0%, #1349a0 100%) !important;
    border: 1px solid rgba(24, 90, 189, 0.8) !important;
    color: white !important;
    font-weight: 600 !important;
}

[data-testid="stForm"] [data-testid="stFileUploader"]:nth-of-type(2) button:hover {
    background: linear-gradient(135deg, #1c66d1 0%, #1552b3 100%) !important;
    box-shadow: 0 4px 16px rgba(24, 90, 189, 0.35) !important;
}

/* ===== 상태 메시지 스타일 ===== */
.stAlert {
    border-radius: 12px;
    border-left-width: 4px;
}

[data-testid="stStatusWidget"] {
    border-radius: 12px;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

/* ===== Expander 스타일 ===== */
[data-testid="stExpander"] {
    border-radius: 10px;
    border: 1px solid rgba(226, 232, 240, 0.8);
}

[data-testid="stExpander"] summary {
    border-radius: 10px;
}

/* ===== 컬럼 간격 조정 ===== */
[data-testid="column"] {
    padding: 0 8px;
}

/* ===== 다크모드 지원 ===== */
@media (prefers-color-scheme: dark) {
    [data-testid="stForm"] {
        background: rgba(30, 41, 59, 0.4);
        border-color: rgba(51, 65, 85, 0.6);
    }
    
    [data-testid="stFileUploaderDropzone"] {
        background: rgba(30, 41, 59, 0.4) !important;
        border-color: rgba(71, 85, 105, 0.5) !important;
    }
    
    [data-testid="stFileUploaderDropzone"]:hover {
        background: rgba(30, 41, 59, 0.6) !important;
        border-color: rgba(100, 116, 139, 0.7) !important;
    }
    
    .h4 {
        color: rgba(248, 250, 252, 0.9);
    }
}

/* ===== 애니메이션 ===== */
@keyframes fadeIn {
    from {
        opacity: 0;
        transform: translateY(10px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

[data-testid="stFileUploader"] {
    animation: fadeIn 0.3s ease-out;
}

/* ===== 스크롤바 스타일 ===== */
::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}

::-webkit-scrollbar-track {
    background: rgba(241, 245, 249, 0.5);
    border-radius: 4px;
}

::-webkit-scrollbar-thumb {
    background: rgba(148, 163, 184, 0.5);
    border-radius: 4px;
}

::-webkit-scrollbar-thumb:hover {
    background: rgba(100, 116, 139, 0.7);
}
"""

def inject():
    """CSS 스타일을 페이지에 주입합니다."""
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
    
    # JavaScript로 클래스 추가
    st.markdown("""
    <script>
    // 페이지 로드 후 클래스 추가
    setTimeout(function() {
        const fileUploaders = document.querySelectorAll('[data-testid="stFileUploader"]');
        
        if (fileUploaders.length >= 2) {
            // 첫 번째는 Excel
            fileUploaders[0].classList.add('excel-uploader');
            // 두 번째는 Word
            fileUploaders[1].classList.add('word-uploader');
        }
        
        // MutationObserver로 동적 변경 감지
        const observer = new MutationObserver(function(mutations) {
            const uploaders = document.querySelectorAll('[data-testid="stFileUploader"]');
            if (uploaders.length >= 2) {
                uploaders[0].classList.add('excel-uploader');
                uploaders[1].classList.add('word-uploader');
            }
        });
        
        observer.observe(document.body, {
            childList: true,
            subtree: true
        });
    }, 100);
    </script>
    """, unsafe_allow_html=True)

def h4(text):
    """커스텀 h4 제목을 렌더링합니다."""
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)

def small_note(text):
    """작은 노트 텍스트를 렌더링합니다."""
    st.markdown(f'<div class="small-note">{text}</div>', unsafe_allow_html=True)
