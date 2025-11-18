import streamlit as st

BASE_CSS = """
/* 전체 폰트 및 배경 */
@import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@400;500;600;700;800&display=swap');

* {
    font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, sans-serif;
}

html, body, [data-testid="stAppViewContainer"] {
    background: #ffffff;
}

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container {
    padding: 1.5rem 1.5rem 3rem;
    max-width: 1200px !important;
    margin: auto;
}

/* 2열 그리드 레이아웃 */
.upload-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 1.5rem;
    margin-bottom: 2rem;
}

/* 엑셀 카드 - 초록색 */
.excel-card {
    background: #ffffff;
    border-radius: 20px;
    padding: 1.5rem;
    border: 2px solid #d1fae5;
    transition: all 0.3s ease;
    box-shadow: 0 4px 12px rgba(16, 185, 129, 0.08);
    display: flex;
    flex-direction: column;
}

.excel-card:hover {
    border-color: #10b981;
    box-shadow: 0 8px 24px rgba(16, 185, 129, 0.15);
    transform: translateY(-2px);
}

.excel-card .card-icon {
    width: 48px;
    height: 48px;
    background: linear-gradient(135deg, #10b981, #059669);
    border-radius: 14px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.5rem;
    margin-bottom: 0.8rem;
}

.excel-card .card-title {
    font-size: 1.25rem;
    font-weight: 700;
    color: #065f46;
    margin-bottom: 1rem;
}

.excel-card .card-description {
    display: none;
}

/* 워드 카드 - 파란색 */
.word-card {
    background: #ffffff;
    border-radius: 20px;
    padding: 1.5rem;
    border: 2px solid #dbeafe;
    transition: all 0.3s ease;
    box-shadow: 0 4px 12px rgba(59, 130, 246, 0.08);
    display: flex;
    flex-direction: column;
}

.word-card:hover {
    border-color: #3b82f6;
    box-shadow: 0 8px 24px rgba(59, 130, 246, 0.15);
    transform: translateY(-2px);
}

.word-card .card-icon {
    width: 48px;
    height: 48px;
    background: linear-gradient(135deg, #3b82f6, #2563eb);
    border-radius: 14px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.5rem;
    margin-bottom: 0.8rem;
}

.word-card .card-title {
    font-size: 1.25rem;
    font-weight: 700;
    color: #1e40af;
    margin-bottom: 1rem;
}

.word-card .card-description {
    display: none;
}

/* 파일 업로더 - 엑셀용 (초록색) */
.excel-card [data-testid="stFileUploader"] {
    border: 2px dashed #86efac;
    border-radius: 16px;
    padding: 3rem 1.5rem;
    background: #f0fdf4;
    transition: all 0.3s ease;
    text-align: center;
    margin-top: 0;
}

.excel-card [data-testid="stFileUploader"]:hover {
    border-color: #10b981;
    background: #dcfce7;
}

.excel-card [data-testid="stFileUploader"] button {
    background: linear-gradient(135deg, #10b981, #059669) !important;
    color: white !important;
    border: none !important;
    padding: 0.75rem 2rem !important;
    border-radius: 12px !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    transition: all 0.2s ease !important;
}

.excel-card [data-testid="stFileUploader"] button:hover {
    background: linear-gradient(135deg, #059669, #047857) !important;
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3) !important;
}

/* 파일 업로더 - 워드용 (파란색) */
.word-card [data-testid="stFileUploader"] {
    border: 2px dashed #93c5fd;
    border-radius: 16px;
    padding: 3rem 1.5rem;
    background: #eff6ff;
    transition: all 0.3s ease;
    text-align: center;
    margin-top: 0;
}

.word-card [data-testid="stFileUploader"]:hover {
    border-color: #3b82f6;
    background: #dbeafe;
}

.word-card [data-testid="stFileUploader"] button {
    background: linear-gradient(135deg, #3b82f6, #2563eb) !important;
    color: white !important;
    border: none !important;
    padding: 0.75rem 2rem !important;
    border-radius: 12px !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    transition: all 0.2s ease !important;
}

.word-card [data-testid="stFileUploader"] button:hover {
    background: linear-gradient(135deg, #2563eb, #1d4ed8) !important;
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3) !important;
}

/* 공통 파일 업로더 스타일 */
[data-testid="stFileUploader"] section {
    border: none !important;
    padding: 0 !important;
    background: transparent !important;
}

[data-testid="stFileUploader"] label {
    display: none !important;
}

/* 파일 업로더 내 모든 텍스트 숨김 */
[data-testid="stFileUploader"] small,
[data-testid="stFileUploader"] [data-testid="stMarkdownContainer"] {
    display: none !important;
}

/* 업로드 완료 표시 완전히 숨김 */
[data-testid="stUploadedFile"],
[data-testid="stUploadedFileName"],
[data-testid="stFileUploader"] > div > div:not(:has(button)) {
    display: none !important;
}

.stSuccess {
    display: none !important;
}

.stAlert[data-baseweb="notification"] {
    display: none !important;
}

/* 옵션 영역 */
.options-section {
    background: #ffffff;
    border-radius: 20px;
    padding: 2rem;
    border: 2px solid #e2e8f0;
    margin-bottom: 2rem;
}

/* 셀렉트 박스 */
[data-testid="stSelectbox"] label {
    font-weight: 600 !important;
    color: #1e293b !important;
    font-size: 1rem !important;
    margin-bottom: 0.5rem !important;
}

[data-testid="stSelectbox"] > div > div {
    border-radius: 12px !important;
    border: 2px solid #e2e8f0 !important;
    background: #ffffff !important;
    padding: 0.7rem 1rem !important;
    font-size: 0.95rem !important;
    transition: all 0.2s ease;
}

[data-testid="stSelectbox"] > div > div:hover,
[data-testid="stSelectbox"] > div > div:focus-within {
    border-color: #3b82f6 !important;
}

/* 텍스트 입력 */
input[type="text"] {
    border-radius: 12px !important;
    padding: 0.85rem 1.2rem !important;
    border: 2px solid #e2e8f0 !important;
    font-size: 0.95rem !important;
    transition: all 0.2s ease;
    background: #ffffff !important;
}

input[type="text"]:focus {
    border-color: #3b82f6 !important;
    box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1) !important;
    outline: none !important;
}

/* 버튼 - 상단 바 */
.top-bar-inner .stButton > button {
    height: 44px !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    font-size: 0.95rem !important;
    background: #ffffff !important;
    color: #1e293b !important;
    border: none !important;
    padding: 0 2rem !important;
    transition: all 0.2s ease;
}

.top-bar-inner .stButton > button:hover {
    background: #f1f5f9 !important;
    transform: translateY(-1px);
}

/* 버튼 - 일반 */
.stButton > button,
[data-testid="stDownloadButton"] > button {
    height: 52px !important;
    border-radius: 12px !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    background: linear-gradient(135deg, #6366f1, #8b5cf6) !important;
    color: #ffffff !important;
    border: none !important;
    padding: 0 2.5rem !important;
    transition: all 0.25s ease;
    box-shadow: 0 4px 14px rgba(99, 102, 241, 0.3);
}

.stButton > button:hover,
[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(99, 102, 241, 0.4);
}

/* 성공/에러 메시지 - 성공 완전히 숨김 */
.stSuccess,
[data-testid="stNotificationContentSuccess"] {
    display: none !important;
}

.stError {
    background: #fef2f2 !important;
    border-left: 4px solid #ef4444 !important;
    color: #991b1b !important;
    border-radius: 12px !important;
    padding: 1rem 1.2rem !important;
    font-weight: 500 !important;
}

.stInfo {
    background: #eff6ff !important;
    border-left: 4px solid #3b82f6 !important;
    color: #1e40af !important;
    border-radius: 12px !important;
    padding: 1rem 1.2rem !important;
    font-weight: 500 !important;
}

.stInfo {
    background: #eff6ff !important;
    border-left: 4px solid #3b82f6 !important;
    color: #1e40af !important;
    border-radius: 12px !important;
    padding: 1rem 1.2rem !important;
    font-weight: 500 !important;
}

/* 진행바 */
[data-testid="stProgress"] > div {
    background-color: #e5e7eb !important;
    border-radius: 999px !important;
    height: 6px !important;
}

[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, #6366f1, #8b5cf6) !important;
    border-radius: 999px !important;
}

/* 구분선 제거 */
hr {
    display: none !important;
}

/* 컬럼 간격 */
[data-testid="column"] {
    padding: 0 0.75rem !important;
}

[data-testid="column"]:first-child {
    padding-left: 0 !important;
}

[data-testid="column"]:last-child {
    padding-right: 0 !important;
}

/* 빈 요소 제거 */
div[data-testid="stMarkdownContainer"] p:empty {
    display: none !important;
}

/* 여백 정리 */
.stMarkdown {
    margin-bottom: 0 !important;
}

/* 스피너 */
[data-testid="stSpinner"] > div {
    border-color: #6366f1 !important;
}

/* 애니메이션 */
@keyframes slideUp {
    from {
        opacity: 0;
        transform: translateY(30px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.excel-card, .word-card {
    animation: slideUp 0.5s ease-out backwards;
}

.excel-card {
    animation-delay: 0.1s;
}

.word-card {
    animation-delay: 0.2s;
}

/* 다크 모드 */
@media (prefers-color-scheme: dark) {
    html, body, [data-testid="stAppViewContainer"] {
        background: #0f172a;
    }
    
    .block-container {
        background: #0f172a;
    }
    
    h1 {
        color: #f1f5f9 !important;
    }
    
    .app-subtitle {
        color: #94a3b8;
    }
    
    .top-bar-inner {
        background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.4);
    }
    
    /* 엑셀 카드 - 다크 모드 */
    .excel-card {
        background: #1e293b;
        border-color: #065f46;
        box-shadow: 0 4px 12px rgba(16, 185, 129, 0.15);
    }
    
    .excel-card:hover {
        border-color: #10b981;
        box-shadow: 0 8px 24px rgba(16, 185, 129, 0.25);
    }
    
    .excel-card .card-title {
        color: #6ee7b7;
    }
    
    .excel-card .card-description {
        color: #6ee7b7;
    }
    
    .excel-card [data-testid="stFileUploader"] {
        background: rgba(16, 185, 129, 0.05);
        border-color: #065f46;
    }
    
    .excel-card [data-testid="stFileUploader"]:hover {
        background: rgba(16, 185, 129, 0.1);
        border-color: #10b981;
    }
    
    /* 워드 카드 - 다크 모드 */
    .word-card {
        background: #1e293b;
        border-color: #1e40af;
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.15);
    }
    
    .word-card:hover {
        border-color: #3b82f6;
        box-shadow: 0 8px 24px rgba(59, 130, 246, 0.25);
    }
    
    .word-card .card-title {
        color: #93c5fd;
    }
    
    .word-card .card-description {
        color: #93c5fd;
    }
    
    .word-card [data-testid="stFileUploader"] {
        background: rgba(59, 130, 246, 0.05);
        border-color: #1e40af;
    }
    
    .word-card [data-testid="stFileUploader"]:hover {
        background: rgba(59, 130, 246, 0.1);
        border-color: #3b82f6;
    }
    
    /* 옵션 섹션 - 다크 모드 */
    .options-section {
        background: #1e293b;
        border-color: #334155;
    }
    
    [data-testid="stSelectbox"] label {
        color: #f1f5f9 !important;
    }
    
    [data-testid="stSelectbox"] > div > div {
        background: #0f172a !important;
        border-color: #334155 !important;
        color: #f1f5f9 !important;
    }
    
    [data-testid="stSelectbox"] > div > div:hover {
        border-color: #3b82f6 !important;
    }
    
    input[type="text"] {
        background: #0f172a !important;
        border-color: #334155 !important;
        color: #f1f5f9 !important;
    }
    
    input[type="text"]:focus {
        border-color: #3b82f6 !important;
    }
}

/* 반응형 */
@media (max-width: 768px) {
    .upload-grid {
        grid-template-columns: 1fr;
    }
    
    .top-bar-inner {
        flex-direction: column;
        gap: 1rem;
        text-align: center;
    }
}
"""

def inject():
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

def h4(text: str):
    st.markdown(f'<div class="card-title">{text}</div>', unsafe_allow_html=True)

def section_caption(text: str):
    st.markdown(f'<div class="card-description">{text}</div>', unsafe_allow_html=True)

def small_note(text: str):
    st.markdown(f'<div class="card-description">{text}</div>', unsafe_allow_html=True)
