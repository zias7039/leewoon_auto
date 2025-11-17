import streamlit as st

EXCEL_GREEN = "#217346"
WORD_BLUE = "#185ABD"

BASE_CSS = """
/* 기본 레이아웃 */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container {
    padding-top: 1.5rem;
    max-width: 900px;
}

/* 제목 아래 설명 간격 */
h1 {
    margin-bottom: 0.2rem;
}
.app-subtitle {
    font-size: 0.9rem;
    color: rgba(100, 116, 139, 0.95);
    margin-bottom: 1.2rem;
}

/* 카드 레이아웃 */
.app-card {
    background: rgba(248, 250, 252, 0.95);
    border-radius: 16px;
    padding: 18px 18px 14px;
    border: 1px solid rgba(226, 232, 240, 0.9);
    box-shadow: 0 4px 18px rgba(15, 23, 42, 0.06);
    margin-bottom: 1rem;
}

/* 섹션 제목 */
.h4 {
    font-weight: 700;
    font-size: 1rem;
    margin: 0 0 0.35rem;
    color: rgba(15, 23, 42, 0.9);
}
.section-caption {
    font-size: 0.8rem;
    color: rgba(100, 116, 139, 0.9);
    margin-bottom: 0.6rem;
}

/* 업로더 라벨 색감 약간 조정 */
[data-testid="stFileUploader"] label {
    font-size: 0.88rem;
    color: rgba(71, 85, 105, 0.98);
}

/* 버튼 스타일 */
.stButton > button {
    height: 44px;
    border-radius: 999px;
    font-weight: 600;
    font-size: 0.95rem;
    border: none;
    background: linear-gradient(135deg, #2563EB, #1D4ED8);
    color: white;
    box-shadow: 0 8px 18px rgba(37, 99, 235, 0.25);
    transition: all 0.18s ease-out;
}
.stButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 10px 22px rgba(37, 99, 235, 0.28);
}

/* ZIP 다운로드 버튼도 동일 스타일 */
[data-testid="stDownloadButton"] > button {
    height: 44px;
    border-radius: 999px;
    font-weight: 600;
    font-size: 0.95rem;
}

/* 텍스트 입력 */
input[type="text"] {
    border-radius: 10px !important;
    border: 1px solid rgba(203, 213, 225, 0.9) !important;
    padding: 9px 11px !important;
}
input[type="text"]:focus {
    border-color: rgba(37, 99, 235, 0.8) !important;
    box-shadow: 0 0 0 1px rgba(37, 99, 235, 0.5) !important;
}

/* 구분선 간격 */
hr {
    margin: 0.9rem 0;
}

/* 다크모드 간단 대응 */
@media (prefers-color-scheme: dark) {
    .app-card {
        background: rgba(15, 23, 42, 0.9);
        border-color: rgba(51, 65, 85, 0.9);
        box-shadow: 0 4px 18px rgba(0, 0, 0, 0.6);
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

def h4(text: str):
    st.markdown(f'<div class="h4">{text}</div>', unsafe_allow_html=True)

def section_caption(text: str):
    st.markdown(f'<div class="section-caption">{text}</div>', unsafe_allow_html=True)

def small_note(text: str):
    st.markdown(f'<div class="section-caption">{text}</div>', unsafe_allow_html=True)
