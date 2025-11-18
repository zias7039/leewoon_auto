import streamlit as st

BASE_CSS = """
/* 전체 폰트 및 배경 */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

* {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
}

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

.block-container {
    padding-top: 2rem;
    padding-bottom: 3rem;
    max-width: 1000px;
    margin: auto;
}

/* 메인 타이틀 스타일링 */
h1 {
    font-size: 2.2rem !important;
    font-weight: 700 !important;
    background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    margin-bottom: 0.5rem !important;
    letter-spacing: -0.02em;
}

/* 서브타이틀 */
.app-subtitle {
    font-size: 1rem;
    color: #64748b;
    margin-bottom: 2rem !important;
    line-height: 1.6;
}

/* 상단 고정 바 - 글래스모피즘 */
.top-bar {
    position: sticky;
    top: 0;
    z-index: 100;
    padding: 0.5rem 0 1rem;
    margin-bottom: 1.5rem;
}

.top-bar-inner {
    background: rgba(255, 255, 255, 0.85);
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    border-radius: 16px;
    padding: 0.8rem 1.5rem;
    border: 1px solid rgba(226, 232, 240, 0.6);
    box-shadow: 0 8px 32px rgba(15, 23, 42, 0.08),
                0 1px 3px rgba(15, 23, 42, 0.06);
}

.top-bar-title {
    font-size: 1rem;
    font-weight: 600;
    color: #0f172a;
    letter-spacing: -0.01em;
}

/* 메인 카드 - 그라데이션 테두리 */
.app-card {
    position: relative;
    background: #ffffff;
    border-radius: 24px;
    padding: 2rem 2.5rem;
    margin-top: 1rem;
    box-shadow: 0 4px 6px rgba(15, 23, 42, 0.03),
                0 10px 40px rgba(15, 23, 42, 0.08);
    border: 1px solid rgba(226, 232, 240, 0.8);
    transition: all 0.3s ease;
}

.app-card:hover {
    box-shadow: 0 8px 12px rgba(15, 23, 42, 0.05),
                0 16px 48px rgba(15, 23, 42, 0.12);
    transform: translateY(-2px);
}

/* 섹션 제목 - 아이콘 스타일 */
.h4 {
    font-size: 1.1rem;
    font-weight: 700;
    margin: 1.5rem 0 0.5rem;
    color: #0f172a;
    letter-spacing: -0.01em;
    display: flex;
    align-items: center;
}

.h4:first-child {
    margin-top: 0;
}

.h4::before {
    content: '';
    display: inline-block;
    width: 4px;
    height: 20px;
    background: linear-gradient(180deg, #3b82f6, #1e40af);
    border-radius: 2px;
    margin-right: 0.6rem;
}

/* 설명 텍스트 */
.section-caption {
    font-size: 0.9rem;
    color: #64748b;
    margin-bottom: 0.8rem !important;
    line-height: 1.5;
}

/* 파일 업로더 영역 개선 */
[data-testid="stFileUploader"] {
    border: 2px dashed #cbd5e1;
    border-radius: 16px;
    padding: 1.5rem;
    background: #f8fafc;
    transition: all 0.2s ease;
}

[data-testid="stFileUploader"]:hover {
    border-color: #3b82f6;
    background: #eff6ff;
}

[data-testid="stFileUploader"] section {
    border: none !important;
    padding: 0 !important;
}

[data-testid="stFileUploader"] label {
    font-weight: 500 !important;
    color: #475569 !important;
    font-size: 0.9rem !important;
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

[data-testid="stSelectbox"] > div > div {
    border-radius: 12px !important;
    border-color: #e2e8f0 !important;
    transition: all 0.2s ease;
}

[data-testid="stSelectbox"] > div > div:hover {
    border-color: #3b82f6 !important;
}

/* 텍스트 입력 */
input[type="text"] {
    border-radius: 12px !important;
    padding: 0.75rem 1rem !important;
    border: 1px solid #e2e8f0 !important;
    font-size: 0.95rem !important;
    transition: all 0.2s ease;
}

input[type="text"]:focus {
    border-color: #3b82f6 !important;
    box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1) !important;
    outline: none !important;
}

/* 버튼 스타일 - 모던한 느낌 */
.stButton > button,
[data-testid="stDownloadButton"] > button {
    height: 48px !important;
    border-radius: 12px !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%) !important;
    color: #ffffff !important;
    border: none !important;
    padding: 0 2rem !important;
    transition: all 0.2s ease;
    box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3);
    letter-spacing: -0.01em;
}

.stButton > button:hover,
[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 20px rgba(59, 130, 246, 0.4);
}

.stButton > button:active,
[data-testid="stDownloadButton"] > button:active {
    transform: translateY(0);
}

/* 성공/에러 메시지 개선 */
.stAlert {
    border-radius: 12px !important;
    border: none !important;
    padding: 0.75rem 1rem !important;
    font-size: 0.9rem !important;
}

[data-testid="stNotification"] {
    border-radius: 12px !important;
}

/* 진행바 */
[data-testid="stProgress"] > div {
    background-color: #e2e8f0 !important;
    border-radius: 999px !important;
    overflow: hidden;
}

[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, #3b82f6, #1e40af) !important;
    border-radius: 999px !important;
}

/* 구분선 */
hr {
    margin: 1.5rem 0 !important;
    border: none !important;
    height: 1px !important;
    background: linear-gradient(90deg, transparent, #e2e8f0, transparent) !important;
}

/* 빈 요소 제거 */
div[data-testid="stMarkdownContainer"] p:empty {
    display: none !important;
}

/* 스피너 */
[data-testid="stSpinner"] > div {
    border-color: #3b82f6 !important;
}

/* 다크 모드 */
@media (prefers-color-scheme: dark) {
    .block-container {
        background: #0f172a;
    }
    
    h1 {
        background: linear-gradient(135deg, #60a5fa 0%, #93c5fd 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }
    
    .app-subtitle {
        color: #94a3b8;
    }
    
    .top-bar-inner {
        background: rgba(15, 23, 42, 0.9);
        backdrop-filter: blur(20px);
        border-color: rgba(51, 65, 85, 0.6);
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
    }
    
    .top-bar-title {
        color: #f1f5f9;
    }
    
    .app-card {
        background: rgba(15, 23, 42, 0.8);
        border-color: rgba(51, 65, 85, 0.6);
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1),
                    0 10px 40px rgba(0, 0, 0, 0.3);
    }
    
    .app-card:hover {
        box-shadow: 0 8px 12px rgba(0, 0, 0, 0.2),
                    0 16px 48px rgba(0, 0, 0, 0.4);
    }
    
    .h4 {
        color: #f1f5f9;
    }
    
    .section-caption {
        color: #94a3b8;
    }
    
    [data-testid="stFileUploader"] {
        border-color: #334155;
        background: rgba(30, 41, 59, 0.5);
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: #60a5fa;
        background: rgba(30, 41, 59, 0.8);
    }
    
    input[type="text"] {
        background: rgba(30, 41, 59, 0.5) !important;
        border-color: #334155 !important;
        color: #f1f5f9 !important;
    }
    
    input[type="text"]:focus {
        background: rgba(30, 41, 59, 0.8) !important;
        border-color: #60a5fa !important;
    }
    
    [data-testid="stSelectbox"] > div > div {
        background: rgba(30, 41, 59, 0.5) !important;
        border-color: #334155 !important;
        color: #f1f5f9 !important;
    }
}

/* 애니메이션 */
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

.app-card {
    animation: fadeIn 0.4s ease-out;
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
