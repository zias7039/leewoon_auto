import streamlit as st

BASE_CSS = """
/* 기본 레이아웃 */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container {
    padding-top: 1.6rem;
    max-width: 900px;
    margin: auto;
}

/* 상단 고정 바 */
.top-bar {
    position: sticky;
    top: 0;
    z-index: 50;
    padding: 0.35rem 0 0.75rem;
}

.top-bar-inner {
    background: rgba(15, 23, 42, 0.96);
    border-radius: 999px;
    padding: 0.4rem 0.9rem;
    border: 1px solid rgba(51, 65, 85, 0.95);
    box-shadow: 0 10px 30px rgba(15, 23, 42, 0.45);
}

.top-bar-title {
    font-size: 0.9rem;
    font-weight: 600;
    color: rgba(226, 232, 240, 0.96);
}

/* 서브타이틀 */
.app-subtitle {
    font-size: 0.92rem;
    color: rgba(148, 163, 184, 0.95);
    margin-bottom: 1.2rem !important;
}

/* 카드 스타일 (금융사 대시보드 느낌) */
.app-card {
    background: rgba(248, 250, 252, 0.96);
    border-radius: 18px;
    padding: 22px 24px 20px;
    border: 1px solid rgba(226, 232, 240, 0.9);
    box-shadow: 0 6px 22px rgba(15, 23, 42, 0.06);
    margin-top: 0.8rem;
}

/* 섹션 제목 */
.h4 {
    font-size: 1.0rem;
    font-weight: 700;
    margin: 0 0 0.35rem;
    color: rgba(15, 23, 42, 0.96);
}

/* 설명 텍스트 */
.section-caption {
    font-size: 0.82rem;
    color: rgba(100, 116, 139, 0.9);
    margin-bottom: 0.6rem !important;
}

/* 업로더 스타일 */
[data-testid="stFileUploader"] {
    margin-top: 0.15rem;
}
[data-testid="stFileUploader"] section {
    padding: 12px 10px;
    border-radius: 14px;
}

/* 텍스트 입력 */
input[type="text"] {
    border-radius: 10px !important;
    padding: 10px 12px !important;
    border: 1px solid rgba(203, 213, 225, 0.9) !important;
}
input[type="text"]:focus {
    border-color: rgba(37, 99, 235, 0.85) !important;
    box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.25) !important;
}

/* 버튼 (상단 ZIP, 하단 ZIP, 다운로드 전부 공통) */
.stButton > button,
[data-testid="stDownloadButton"] > button {
    height: 42px !important;
    border-radius: 999px !important;
    font-weight: 600 !important;
    font-size: 0.92rem !important;
    background: linear-gradient(135deg, #0f3b82, #1d4ed8) !important;
    color: #f9fafb !important;
    border: none !important;
    padding: 0.1rem 1.4rem !important;
    transition: all 0.16s ease-out;
}
.stButton > button:hover,
[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 10px 24px rgba(15, 23, 42, 0.25);
}

/* 진행바 슬림하게 */
[data-testid="stProgress"] > div > div {
    border-radius: 999px;
}

/* 빈 p 태그 제거 (애매한 막대 문제 방지) */
div[data-testid="stMarkdownContainer"] p:empty {
    display: none !important;
}

/* 다크 모드 */
@media (prefers-color-scheme: dark) {
    .top-bar-inner {
        background: rgba(15, 23, 42, 0.98);
        border-color: rgba(51, 65, 85, 0.95);
    }
    .app-card {
        background: rgba(15, 23, 42, 0.9);
        border-color: rgba(51, 65, 85, 0.85);
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.5);
    }
    .h4 {
        color: rgba(248, 250, 252, 0.98);
    }
    .section-caption {
        color: rgba(148, 163, 184, 0.95);
    }
}
"""

def inject():
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

def h4(text: str):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)

def section_caption(text: str):
    st.markdown(f'<div class="section-caption">{text}</div>', unsafe_allow_html=True)

def small_note(text: str):
    st.markdown(f'<div class="section-caption">{text}</div>', unsafe_allow_html=True)
