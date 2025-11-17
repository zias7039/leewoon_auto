import streamlit as st

BASE_CSS = """
/* 기본 레이아웃 */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.stApp {
    background-color: #0b1120;
}

/* 컨테이너 폭 */
.block-container {
    padding-top: 1.2rem;
    max-width: 900px;
    margin: auto;
}

/* 상단 고정 바 */
.fixed-top-bar {
    position: sticky;
    top: 0;
    z-index: 999;
    padding: 10px 14px;
    margin-bottom: 1.0rem;
    border-radius: 14px;
    background: rgba(15, 23, 42, 0.96);
    border: 1px solid rgba(51, 65, 85, 0.9);
    box-shadow: 0 8px 30px rgba(15, 23, 42, 0.5);
    backdrop-filter: blur(10px);
}
.top-bar-inner {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 1rem;
}
.top-bar-title {
    font-size: 1.05rem;
    font-weight: 700;
    color: #e5e7eb;
}
.top-bar-subtitle {
    font-size: 0.8rem;
    color: #9ca3af;
}

/* 카드 */
.app-card {
    background: rgba(15, 23, 42, 0.92);
    border-radius: 18px;
    padding: 20px 22px 18px;
    border: 1px solid rgba(31, 41, 55, 0.95);
    box-shadow: 0 10px 26px rgba(15, 23, 42, 0.6);
    margin-bottom: 1.3rem;
}

/* 섹션 제목 */
.h4 {
    font-size: 1rem;
    font-weight: 700;
    margin: 0 0 0.35rem;
    color: #e5e7eb;
}

/* 설명 캡션 */
.section-caption {
    font-size: 0.78rem;
    color: #9ca3af;
    margin-bottom: 0.65rem;
}

/* 업로더 박스 */
[data-testid="stFileUploader"] section {
    padding: 10px 12px;
    border-radius: 14px;
    border: 1px solid rgba(55, 65, 81, 0.9);
    background: rgba(15, 23, 42, 0.95);
}
[data-testid="stFileUploader"] label {
    font-size: 0.86rem;
    color: #e5e7eb;
}

/* 텍스트 입력 */
input[type="text"] {
    border-radius: 10px !important;
    padding: 9px 11px !important;
    border: 1px solid rgba(55, 65, 81, 0.9) !important;
    background: rgba(15, 23, 42, 0.9) !important;
    color: #e5e7eb !important;
}
input[type="text"]:focus {
    border-color: #2563eb !important;
    box-shadow: 0 0 0 1px rgba(37, 99, 235, 0.8) !important;
}

/* 버튼 – 금융사 느낌의 네이비 블루 */
.stButton > button,
[data-testid="stDownloadButton"] > button {
    height: 44px !important;
    border-radius: 999px !important;
    font-weight: 600 !important;
    font-size: 0.93rem !important;
    border: none !important;
    background: linear-gradient(135deg, #2563eb, #1d4ed8) !important;
    color: #f9fafb !important;
    box-shadow: 0 10px 24px rgba(37, 99, 235, 0.3);
    transition: all 0.18s ease-out;
}
.stButton > button:hover,
[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 14px 30px rgba(37, 99, 235, 0.4);
}

/* 프로그레스 바 색감 */
[data-testid="stProgressBar"] > div > div {
    background: linear-gradient(90deg, #22c55e, #16a34a);
}

/* 빈 p → 보이지 않게 (애매한 막대 제거) */
div[data-testid="stMarkdownContainer"] p:empty {
    display: none !important;
}

/* 라이트모드 대응 */
@media (prefers-color-scheme: light) {
    .stApp {
        background-color: #f3f4f6;
    }
    .fixed-top-bar {
        background: #ffffff;
        border-color: rgba(209, 213, 219, 0.9);
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.08);
    }
    .app-card {
        background: #ffffff;
        border-color: rgba(209, 213, 219, 0.9);
        box-shadow: 0 10px 26px rgba(148, 163, 184, 0.3);
    }
    .h4 {
        color: #111827;
    }
    .section-caption {
        color: #6b7280;
    }
    [data-testid="stFileUploader"] section {
        background: #f9fafb;
        border-color: rgba(209, 213, 219, 0.9);
    }
    input[type="text"] {
        background: #ffffff !important;
        color: #111827 !important;
        border-color: rgba(209, 213, 219, 0.9) !important;
    }
}

/* 스크롤바 */
::-webkit-scrollbar {
    width: 7px;
}
::-webkit-scrollbar-thumb {
    background: rgba(148, 163, 184, 0.7);
    border-radius: 4px;
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
