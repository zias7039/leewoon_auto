import streamlit as st

BASE_CSS = """
/* 기본 레이아웃 */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container {
    padding-top: 1.6rem;
    max-width: 860px;
    margin: auto;
}

/* 서브타이틀 */
.app-subtitle {
    font-size: 0.92rem;
    color: rgba(148, 163, 184, 0.95);
    margin-bottom: 1.2rem !important;
}

/* 카드 스타일 */
.app-card {
    background: rgba(248, 250, 252, 0.9);
    border-radius: 18px;
    padding: 22px 25px 20px;
    border: 1px solid rgba(226, 232, 240, 0.9);
    box-shadow: 0 6px 22px rgba(0, 0, 0, 0.06);
    margin-bottom: 1.4rem;
}

/* 제목 */
.h4 {
    font-size: 1.02rem;
    font-weight: 700;
    margin: 0 0 0.45rem;
    color: rgba(30, 41, 59, 0.95);
}

/* 설명 캡션 */
.section-caption {
    font-size: 0.82rem;
    color: rgba(100, 116, 139, 0.85);
    margin-bottom: 0.7rem !important;
}

/* 업로더 스타일 안정화 */
[data-testid="stFileUploader"] {
    margin-top: 0.1rem;
}
[data-testid="stFileUploader"] section {
    padding: 12px 10px;
    border-radius: 14px;
}

/* 텍스트 입력 폭 조정 */
input[type="text"] {
    border-radius: 10px !important;
    padding: 10px 12px !important;
    border: 1px solid rgba(203, 213, 225, 0.9) !important;
}
input[type="text"]:focus {
    border-color: rgba(37, 99, 235, 0.75) !important;
    box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.25) !important;
}

/* 버튼 */
.stButton > button, 
[data-testid="stDownloadButton"] > button {
    height: 46px !important;
    border-radius: 999px !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    background: linear-gradient(135deg, #2563EB, #1D4ED8) !important;
    color: white !important;
    border: none !important;
    transition: all .18s ease-out;
}
.stButton > button:hover,
[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 10px 24px rgba(37, 99, 235, .25);
}

/* 빈 p 태그 제거 (애매한 막대 문제 해결) */
div[data-testid="stMarkdownContainer"] p:empty {
    display: none !important;
}

/* 다크모드 */
@media (prefers-color-scheme: dark) {
    .app-card {
        background: rgba(30, 41, 59, 0.65);
        backdrop-filter: blur(4px);
        border-color: rgba(71, 85, 105, 0.5);
    }
    .h4 {
        color: rgba(248, 250, 252, 0.95);
    }
    .section-caption {
        color: rgba(148, 163, 184, 0.9);
    }
}
"""

def inject():
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

def h4(text):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)

def section_caption(text):
    st.markdown(f'<div class="section-caption">{text}</div>', unsafe_allow_html=True)

def small_note(text):
    st.markdown(f'<div class="section-caption">{text}</div>', unsafe_allow_html=True)
