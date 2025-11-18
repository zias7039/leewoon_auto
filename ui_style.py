import streamlit as st

BASE_CSS = """
/* 전체 폰트 및 배경 */
@import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@400;500;600;700;800&display=swap');

* {
    font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, sans-serif;
}

html, body, [data-testid="stAppViewContainer"] {
    background: #f5f7fa;
}

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container {
    padding: 1.5rem 1rem 3rem;
    max-width: 1400px !important;
    margin: auto;
}

/* 헤더 영역 - 좌측 정렬 + 배지 스타일 */
h1 {
    font-size: 1.8rem !important;
    font-weight: 800 !important;
    color: #0f172a !important;
    margin-bottom: 0.3rem !important;
    letter-spacing: -0.03em;
}

.app-subtitle {
    font-size: 0.95rem;
    color: #64748b;
    margin-bottom: 2rem !important;
    font-weight: 500;
}

/* 상단 고정 바 - 가로로 길게 */
.top-bar {
    position: sticky;
    top: 0;
    z-index: 100;
    padding: 0 0 1.5rem;
    margin-bottom: 1.5rem;
}

.top-bar-inner {
    background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
    border-radius: 16px;
    padding: 1.2rem 2rem;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
    display: flex;
    align-items: center;
    justify-content: space-between;
}

.top-bar-title {
    font-size: 1.1rem;
    font-weight: 700;
    color: #ffffff;
    letter-spacing: -0.02em;
}

/* 메인 그리드 레이아웃 - 3단 구성 */
.app-card {
    display: none !important; /* 기본 카드 숨김 */
}

/* 커스텀 그리드 컨테이너 */
.grid-container {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 1.5rem;
    margin-bottom: 1.5rem;
}

.grid-full {
    grid-column: 1 / -1;
}

/* 카드 스타일 - 플랫하고 미니멀 */
.upload-card {
    background: #ffffff;
    border-radius: 16px;
    padding: 2rem;
    border: 2px solid #e2e8f0;
    transition: all 0.3s ease;
    height: 100%;
}

.upload-card:hover {
    border-color: #3b82f6;
    box-shadow: 0 8px 24px rgba(59, 130, 246, 0.12);
}

/* 아이콘 + 제목 조합 */
.card-header {
    display: flex;
    align-items: center;
    gap: 0.8rem;
    margin-bottom: 1rem;
}

.card-icon {
    width: 44px;
    height: 44px;
    background: linear-gradient(135deg, #3b82f6, #2563eb);
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.3rem;
}

.card-title {
    font-size: 1.15rem;
    font-weight: 700;
    color: #0f172a;
    margin: 0;
}

.card-description {
    font-size: 0.88rem;
    color: #64748b;
    line-height: 1.5;
    margin-bottom: 1.5rem;
}

/* 파일 업로더 - 심플하게 */
[data-testid="stFileUploader"] {
    border: 2px dashed #cbd5e1;
    border-radius: 14px;
    padding: 2rem 1.5rem;
    background: #f8fafc;
    transition: all 0.25s ease;
    text-align: center;
}

[data-testid="stFileUploader"]:hover {
    border-color: #3b82f6;
    background: #eff6ff;
}

[data-testid="stFileUploader"] section {
    border: none !important;
    padding: 0 !important;
    background: transparent !important;
}

[data-testid="stFileUploader"] label {
    display: none !important;
}

[data-testid="stFileUploader"] button {
    background: #3b82f6 !important;
    color: white !important;
    border: none !important;
    padding: 0.7rem 2rem !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    transition: all 0.2s ease !important;
    margin: 0 auto !important;
}

[data-testid="stFileUploader"] button:hover {
    background: #2563eb !important;
    transform: translateY(-1px);
}

/* 업로드 완료 표시 숨김 */
[data-testid="stUploadedFile"],
[data-testid="stUploadedFileName"] {
    display: none !important;
}

/* 셀렉트 박스 */
[data-testid="stSelectbox"] label {
    font-weight: 600 !important;
    color: #1e293b !important;
    font-size: 0.95rem !important;
    margin-bottom: 0.5rem !important;
}

[data-testid="stSelectbox"] > div > div {
    border-radius: 12px !important;
    border: 2px solid #e2e8f0 !important;
    background: #ffffff !important;
    padding: 0.65rem 1rem !important;
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

/* 버튼 - 두 가지 스타일 */
/* 상단 바 버튼 */
.top-bar-inner .stButton > button {
    height: 44px !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    font-size: 0.9rem !important;
    background: #ffffff !important;
    color: #1e293b !important;
    border: none !important;
    padding: 0 1.8rem !important;
    transition: all 0.2s ease;
}

.top-bar-inner .stButton > button:hover {
    background: #f1f5f9 !important;
    transform: translateY(-1px);
}

/* 일반 버튼 */
.stButton > button,
[data-testid="stDownloadButton"] > button {
    height: 50px !important;
    border-radius: 12px !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    background: linear-gradient(135deg, #3b82f6, #2563eb) !important;
    color: #ffffff !important;
    border: none !important;
    padding: 0 2.5rem !important;
    transition: all 0.25s ease;
    box-shadow: 0 4px 14px rgba(59, 130, 246, 0.3);
}

.stButton > button:hover,
[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(59, 130, 246, 0.4);
}

/* 성공/에러 메시지 - 모서리 배지 스타일 */
.stAlert {
    border-radius: 12px !important;
    border-left: 4px solid !important;
    padding: 1rem 1.2rem !important;
    font-size: 0.9rem !important;
    font-weight: 500 !important;
}

.stSuccess {
    background: #f0fdf4 !important;
    border-color: #22c55e !important;
    color: #166534 !important;
}

.stError {
    background: #fef2f2 !important;
    border-color: #ef4444 !important;
    color: #991b1b !important;
}

.stInfo {
    background: #eff6ff !important;
    border-color: #3b82f6 !important;
    color: #1e40af !important;
}

/* 진행바 - 슬림 */
[data-testid="stProgress"] > div {
    background-color: #e5e7eb !important;
    border-radius: 999px !important;
    height: 6px !important;
}

[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, #3b82f6, #2563eb) !important;
    border-radius: 999px !important;
}

/* 구분선 */
hr {
    margin: 2rem 0 !important;
    border: none !important;
    height: 1px !important;
    background: #e2e8f0 !important;
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

/* 스피너 */
[data-testid="stSpinner"] > div {
    border-color: #3b82f6 !important;
}

/* 반응형 */
@media (max-width: 768px) {
    .grid-container {
        grid-template-columns: 1fr;
    }
    
    .top-bar-inner {
        flex-direction: column;
        gap: 1rem;
        text-align: center;
    }
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

.upload-card {
    animation: slideUp 0.4s ease-out backwards;
}

.upload-card:nth-child(1) {
    animation-delay: 0.1s;
}

.upload-card:nth-child(2) {
    animation-delay: 0.2s;
}

.upload-card:nth-child(3) {
    animation-delay: 0.3s;
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
