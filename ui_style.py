import streamlit as st

BASE_CSS = """
/* 전체 폰트 및 배경 */
@import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@400;500;600;700&display=swap');

* {
    font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans KR', sans-serif;
}

html, body, [data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
}

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container {
    padding: 2rem 1rem 3rem;
    max-width: 1100px;
    margin: auto;
}

/* 메인 타이틀 스타일링 */
h1 {
    font-size: 2.5rem !important;
    font-weight: 800 !important;
    color: #ffffff !important;
    margin-bottom: 0.8rem !important;
    letter-spacing: -0.03em;
    text-shadow: 0 2px 20px rgba(0, 0, 0, 0.2);
}

/* 서브타이틀 */
.app-subtitle {
    font-size: 1.05rem;
    color: rgba(255, 255, 255, 0.9);
    margin-bottom: 2.5rem !important;
    line-height: 1.6;
    font-weight: 400;
}

/* 상단 고정 바 - 글래스모피즘 */
.top-bar {
    position: sticky;
    top: 0;
    z-index: 100;
    padding: 0 0 1.5rem;
    margin-bottom: 2rem;
}

.top-bar-inner {
    background: rgba(255, 255, 255, 0.95);
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    border-radius: 20px;
    padding: 1rem 2rem;
    border: 1px solid rgba(255, 255, 255, 0.3);
    box-shadow: 0 8px 32px rgba(31, 38, 135, 0.15),
                0 1px 3px rgba(0, 0, 0, 0.1);
}

.top-bar-title {
    font-size: 1.05rem;
    font-weight: 700;
    color: #1f2937;
    letter-spacing: -0.02em;
}

/* 메인 카드 - 화이트 배경 */
.app-card {
    position: relative;
    background: #ffffff;
    border-radius: 28px;
    padding: 3rem;
    margin-top: 1rem;
    box-shadow: 0 20px 60px rgba(0, 0, 0, 0.15),
                0 8px 20px rgba(0, 0, 0, 0.08);
    border: none;
}

/* 컬럼 간격 조정 */
[data-testid="column"] {
    padding: 0 1rem !important;
}

[data-testid="column"]:first-child {
    padding-left: 0 !important;
}

[data-testid="column"]:last-child {
    padding-right: 0 !important;
}

/* 섹션 제목 - 심플한 스타일 */
.h4 {
    font-size: 1.2rem;
    font-weight: 700;
    margin: 2rem 0 0.8rem;
    color: #111827;
    letter-spacing: -0.02em;
}

.h4:first-child {
    margin-top: 0;
}

/* 설명 텍스트 */
.section-caption {
    font-size: 0.95rem;
    color: #6b7280;
    margin-bottom: 1.2rem !important;
    line-height: 1.6;
}

/* 파일 업로더 영역 개선 */
[data-testid="stFileUploader"] {
    border: 2px solid #e5e7eb;
    border-radius: 20px;
    padding: 2.5rem 2rem;
    background: #f9fafb;
    transition: all 0.3s ease;
    margin-bottom: 1rem;
}

[data-testid="stFileUploader"]:hover {
    border-color: #6366f1;
    background: #eef2ff;
    transform: translateY(-2px);
    box-shadow: 0 8px 20px rgba(99, 102, 241, 0.1);
}

[data-testid="stFileUploader"] section {
    border: none !important;
    padding: 0 !important;
    background: transparent !important;
}

[data-testid="stFileUploader"] label {
    font-weight: 600 !important;
    color: #374151 !important;
    font-size: 0.95rem !important;
    margin-bottom: 0.5rem !important;
}

/* 업로드 버튼 스타일 */
[data-testid="stFileUploader"] button {
    background: #6366f1 !important;
    color: white !important;
    border: none !important;
    padding: 0.7rem 1.8rem !important;
    border-radius: 12px !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    transition: all 0.2s ease !important;
}

[data-testid="stFileUploader"] button:hover {
    background: #4f46e5 !important;
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(99, 102, 241, 0.3) !important;
}

/* 업로드된 파일 정보는 숨기기 */
[data-testid="stUploadedFile"],
[data-testid="stUploadedFileName"],
[data-testid="stFileUploader"] div div:nth-child(2) {
    display: none !important;
}

/* Select Box 스타일 */
[data-testid="stSelectbox"] {
    margin-top: 0.5rem;
}

[data-testid="stSelectbox"] label {
    font-weight: 600 !important;
    color: #111827 !important;
    font-size: 1rem !important;
    margin-bottom: 0.5rem !important;
}

[data-testid="stSelectbox"] > div > div {
    border-radius: 14px !important;
    border: 2px solid #e5e7eb !important;
    background: #ffffff !important;
    padding: 0.6rem 1rem !important;
    font-size: 0.95rem !important;
    transition: all 0.2s ease;
}

[data-testid="stSelectbox"] > div > div:hover {
    border-color: #6366f1 !important;
    box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.1);
}

/* 텍스트 입력 */
input[type="text"] {
    border-radius: 14px !important;
    padding: 0.85rem 1.2rem !important;
    border: 2px solid #e5e7eb !important;
    font-size: 0.95rem !important;
    transition: all 0.2s ease;
    background: #ffffff !important;
}

input[type="text"]:focus {
    border-color: #6366f1 !important;
    box-shadow: 0 0 0 4px rgba(99, 102, 241, 0.1) !important;
    outline: none !important;
}

/* 버튼 스타일 - 통일된 디자인 */
.stButton > button,
[data-testid="stDownloadButton"] > button {
    height: 52px !important;
    border-radius: 14px !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%) !important;
    color: #ffffff !important;
    border: none !important;
    padding: 0 2.5rem !important;
    transition: all 0.3s ease;
    box-shadow: 0 4px 20px rgba(99, 102, 241, 0.4);
    letter-spacing: -0.01em;
}

.stButton > button:hover,
[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 30px rgba(99, 102, 241, 0.5);
}

.stButton > button:active,
[data-testid="stDownloadButton"] > button:active {
    transform: translateY(-1px);
}

/* 성공/에러 메시지 개선 */
.stAlert {
    border-radius: 14px !important;
    border: none !important;
    padding: 1rem 1.3rem !important;
    font-size: 0.95rem !important;
    font-weight: 500 !important;
}

[data-testid="stNotification"] {
    border-radius: 14px !important;
}

.stSuccess {
    background: #ecfdf5 !important;
    color: #065f46 !important;
}

.stError {
    background: #fef2f2 !important;
    color: #991b1b !important;
}

.stInfo {
    background: #eff6ff !important;
    color: #1e40af !important;
}

/* 진행바 */
[data-testid="stProgress"] > div {
    background-color: #e5e7eb !important;
    border-radius: 999px !important;
    height: 8px !important;
    overflow: hidden;
}

[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, #6366f1, #8b5cf6) !important;
    border-radius: 999px !important;
}

/* 구분선 */
hr {
    margin: 2.5rem 0 !important;
    border: none !important;
    height: 1px !important;
    background: linear-gradient(90deg, transparent, #d1d5db 20%, #d1d5db 80%, transparent) !important;
}

/* 빈 요소 제거 */
div[data-testid="stMarkdownContainer"] p:empty {
    display: none !important;
}

/* 스피너 */
[data-testid="stSpinner"] > div {
    border-color: #6366f1 !important;
}

/* 라벨 스타일 통일 */
label {
    font-weight: 600 !important;
    color: #111827 !important;
    font-size: 1rem !important;
}

/* 애니메이션 */
@keyframes fadeInUp {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.app-card {
    animation: fadeInUp 0.5s ease-out;
}

/* 추가 여백 조정 */
.stMarkdown {
    margin-bottom: 0 !important;
}

/* 다크모드 제거 - 라이트 테마만 사용 */
"""

def inject():
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

def h4(text: str):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)

def section_caption(text: str):
    st.markdown(f'<div class="section-caption">{text}</div>', unsafe_allow_html=True)

def small_note(text: str):
    st.markdown(f'<div class="section-caption">{text}</div>', unsafe_allow_html=True)
