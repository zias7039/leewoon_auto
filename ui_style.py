# ui_style.py
import streamlit as st

def inject_style():
    st.markdown("""
    <style>
    /* 다크 테마 기본 설정 */
    :root {
        --bg-primary: #1e2433;
        --bg-secondary: #2d3548;
        --bg-card: #2a3142;
        --text-primary: #e5e7eb;
        --text-secondary: #9ca3af;
        --border-color: #374357;
        --excel-color: #6dd3a8;
        --word-color: #6eb5ea;
        --success-color: #6dd3a8;
        --warning-color: #f59e0b;
        --error-color: #ef4444;
    }
    
    /* 전체 배경 */
    .stApp {
        background: linear-gradient(135deg, #1a1f2e 0%, #252b3d 100%);
    }
    
    /* 메인 컨테이너 */
    .block-container {
        padding-top: 2rem;
        max-width: 1400px;
    }
    
    /* 헤더 */
    .header-container {
        display: flex;
        justify-content: flex-end;
        padding: 1rem 2rem;
        margin-bottom: 2rem;
    }
    
    .user-profile {
        display: flex;
        align-items: center;
        gap: 0.75rem;
    }
    
    .avatar {
        width: 40px;
        height: 40px;
        border-radius: 50%;
        background: linear-gradient(135deg, var(--excel-color), var(--word-color));
        display: flex;
        align-items: center;
        justify-content: center;
        color: white;
        font-weight: 600;
        font-size: 0.875rem;
    }
    
    .username {
        color: var(--text-primary);
        font-size: 0.95rem;
    }
    
    /* 타이틀 */
    .main-title {
        font-size: 2.5rem;
        font-weight: 700;
        letter-spacing: 0.05em;
        color: var(--text-primary);
        text-align: center;
        margin-bottom: 0.5rem;
    }
    
    .subtitle {
        text-align: center;
        color: var(--text-secondary);
        font-size: 1.1rem;
        margin-bottom: 3rem;
    }
    
    /* 업로드 카드 */
    .upload-card {
        background: var(--bg-card);
        border: 2px solid var(--border-color);
        border-radius: 16px;
        padding: 3rem 2rem;
        text-align: center;
        transition: all 0.3s ease;
        margin-bottom: 1.5rem;
        min-height: 280px;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
    }
    
    .upload-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
    }
    
    .excel-card {
        border-color: rgba(109, 211, 168, 0.3);
    }
    
    .excel-card:hover {
        border-color: var(--excel-color);
        box-shadow: 0 8px 32px rgba(109, 211, 168, 0.2);
    }
    
    .word-card {
        border-color: rgba(110, 181, 234, 0.3);
    }
    
    .word-card:hover {
        border-color: var(--word-color);
        box-shadow: 0 8px 32px rgba(110, 181, 234, 0.2);
    }
    
    .upload-icon {
        width: 80px;
        height: 80px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        margin: 0 auto 1.5rem;
    }
    
    .excel-icon {
        background: rgba(109, 211, 168, 0.15);
        color: var(--excel-color);
    }
    
    .word-icon {
        background: rgba(110, 181, 234, 0.15);
        color: var(--word-color);
    }
    
    .upload-title {
        font-size: 1.25rem;
        font-weight: 700;
        color: var(--text-primary);
        margin-bottom: 0.5rem;
        letter-spacing: 0.05em;
    }
    
    .upload-subtitle {
        color: var(--text-secondary);
        font-size: 0.95rem;
    }
    
    /* 파일 업로더 숨기기 및 스타일링 */
    [data-testid="stFileUploader"] {
        opacity: 0;
        height: 0;
        overflow: hidden;
    }
    
    /* 설정 섹션 */
    .settings-section {
        margin: 2rem 0;
    }
    
    /* 입력 필드 */
    input[type="text"], .stSelectbox {
        background: var(--bg-secondary) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 8px !important;
        color: var(--text-primary) !important;
        padding: 0.75rem !important;
    }
    
    /* 버튼 */
    .stButton > button {
        background: linear-gradient(135deg, var(--excel-color), var(--word-color)) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.75rem 2rem !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 16px rgba(109, 211, 168, 0.3);
    }
    
    /* Recent Generations 섹션 */
    .recent-section {
        background: var(--bg-card);
        border-radius: 16px;
        padding: 2rem;
        margin: 3rem 0;
    }
    
    .section-title {
        font-size: 1.25rem;
        font-weight: 700;
        color: var(--text-primary);
        margin-bottom: 1.5rem;
        letter-spacing: 0.05em;
    }
    
    /* 상태 아이템 */
    .status-item {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        padding: 1rem;
        border-radius: 8px;
        background: var(--bg-secondary);
        border: 1px solid var(--border-color);
    }
    
    .status-item svg {
        width: 24px;
        height: 24px;
    }
    
    .status-item span {
        font-size: 0.9rem;
        font-weight: 500;
    }
    
    .status-complete {
        color: var(--success-color);
        border-color: rgba(109, 211, 168, 0.3);
    }
    
    .status-pending {
        color: var(--warning-color);
        border-color: rgba(245, 158, 11, 0.3);
    }
    
    .status-error {
        color: var(--error-color);
        border-color: rgba(239, 68, 68, 0.3);
    }
    
    .status-inactive {
        color: var(--text-secondary);
        opacity: 0.5;
    }
    
    /* 진행률 바 */
    .progress-container {
        margin-top: 2rem;
    }
    
    .progress-label {
        color: var(--text-secondary);
        font-size: 0.9rem;
        margin-bottom: 0.5rem;
    }
    
    .progress-bar {
        width: 100%;
        height: 8px;
        background: var(--bg-secondary);
        border-radius: 4px;
        overflow: hidden;
        margin-bottom: 0.5rem;
    }
    
    .progress-fill {
        height: 100%;
        background: linear-gradient(90deg, var(--excel-color), var(--word-color));
        transition: width 0.3s ease;
    }
    
    .progress-percentage {
        text-align: right;
        color: var(--text-primary);
        font-weight: 600;
        font-size: 0.9rem;
    }
    
    /* 사이드바 숨기기 */
    [data-testid="stSidebar"] {
        display: none;
    }
    
    /* MainMenu 숨기기 */
    #MainMenu {
        visibility: hidden;
    }
    
    footer {
        visibility: hidden;
    }
    
    /* 다운로드 버튼 */
    [data-testid="stDownloadButton"] > button {
        background: var(--bg-secondary) !important;
        border: 1px solid var(--border-color) !important;
    }
    
    [data-testid="stDownloadButton"] > button:hover {
        background: var(--bg-card) !important;
        border-color: var(--excel-color) !important;
    }
    </style>
    """, unsafe_allow_html=True)
