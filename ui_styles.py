# -*- coding: utf-8 -*-
# styles.py
from __future__ import annotations
import textwrap

# 팔레트(엑셀/워드/중립)
PALETTES = {
    "excel": {
        "brand": "#217346",  # Excel
        "brand-weak": "rgba(33,115,70,.25)",
        "brand-strong": "rgba(33,115,70,.55)",
        "text": "rgba(255,255,255,.92)",
    },
    "word": {# -*- coding: utf-8 -*-
# ui_style.py
# Streamlit 전용 스타일 주입 + 편의 함수

from __future__ import annotations
from datetime import datetime

# 업로더 타입별 대표 색상 (엑셀/워드)
EXCEL_COLOR = "#1D6F42"   # 초록 (엑셀)
WORD_COLOR  = "#185ABD"   # 파랑 (워드)

def inject(st, *, glass=True):
    """
    페이지 공통 CSS를 주입한다.
    - glass=True: 유리(Glassmorphism) 느낌의 카드 배경
    """
    st.markdown(f"""
<style>
:root {{
  --card-bg: rgba(2,6,23,{0.65 if glass else 0.08});
  --card-bd: rgba(148,163,184,{0.25 if glass else 0.3});
  --muted: rgba(0,0,0,.6);
  --excel: {EXCEL_COLOR};
  --word:  {WORD_COLOR};
}}
#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}

.block-container {{ padding-top: 1.2rem; }}

.card {{
  background: var(--card-bg);
  border: 1px solid var(--card-bd);
  border-radius: 16px;
  padding: 16px;
  box-shadow: 0 8px 30px rgba(0,0,0,.25);
}}

.small-note {{ font-size:.85rem; color: var(--muted); }}

.stButton>button {{ height: 44px; border-radius: 10px; }}
[data-testid="stDownloadButton"] > button {{ min-width: 220px; }}

div[data-testid="stForm"].form-card {{
  border: 1px solid var(--card-bd);
  border-radius: 12px;
  padding: 16px 16px 8px 16px;
  background: rgba(17,24,39,{0.55 if glass else 0.03});
}}

.upload-tight [data-testid="stFileUploader"] section {{ gap: 6px; }}
.upload-tight [data-testid="stFileUploader"] button {{ border-radius: 10px; }}

.excel-upload [data-testid="stFileUploaderDropzone"] {{
  border: 1px dashed color-mix(in srgb, var(--excel) 65%, white);
  background: color-mix(in srgb, var(--excel) 8%, transparent);
}}
.excel-upload [data-testid="stFileUploaderDropzone"] p, 
.excel-upload [data-testid="stFileUploaderDropzone"] span {{
  color: color-mix(in srgb, var(--excel) 85%, white);
}}

.word-upload [data-testid="stFileUploaderDropzone"] {{
  border: 1px dashed color-mix(in srgb, var(--word) 65%, white);
  background: color-mix(in srgb, var(--word) 8%, transparent);
}}
.word-upload [data-testid="stFileUploaderDropzone"] p, 
.word-upload [data-testid="stFileUploaderDropzone"] span {{
  color: color-mix(in srgb, var(--word) 85%, white);
}}

/* 토큰 미리보기 code 박스 가독성 */
.token-sample code {{
  white-space: wrap;
  line-height: 1.45;
}}
</style>
""", unsafe_allow_html=True)

def open_div(st, cls: str):
    """任의 CSS 클래스로 div 오픈 (닫기는 close_div)."""
    st.markdown(f'<div class="{cls}">', unsafe_allow_html=True)

def close_div(st):
    """open_div로 연 div 닫기."""
    st.markdown('</div>', unsafe_allow_html=True)

def h4(st, text: str):
    st.markdown(f"#### {text}")

def caption_now(st):
    st.caption(datetime.now().strftime("%Y-%m-%d %H:%M"))

        "brand": "#185ABD",  # Word
        "brand-weak": "rgba(24,90,189,.25)",
        "brand-strong": "rgba(24,90,189,.55)",
        "text": "rgba(255,255,255,.92)",
    },
    "neutral": {
        "brand": "#334155",
        "brand-weak": "rgba(51,65,85,.25)",
        "brand-strong": "rgba(51,65,85,.55)",
        "text": "rgba(255,255,255,.92)",
    },
}

def _base_css(max_width: int = 1080) -> str:
    # 전역 기본 스타일(크롬 숨김, 컨테이너 패딩, 버튼 등)
    return textwrap.dedent(f"""
    <style>
      /* 상단 햄버거/푸터 숨김 */
      #MainMenu, footer {{ visibility: hidden; }}

      /* 레이아웃 */
      .block-container {{ padding-top: 1.2rem; max-width: {max_width}px; }}

      /* 공통 버튼 */
      .stButton>button {{
        height: 44px; border-radius: 10px;
      }}

      /* 다운로드 버튼 너비 고정 */
      [data-testid="stDownloadButton"] > button {{ min-width: 220px; }}

      /* 작은 캡션 */
      .small-note {{ font-size:.85rem; color: rgba(0,0,0,.6); }}
    </style>
    """)

def _glass_css() -> str:
    # 글래스모피즘 카드 / 업로더 스킨
    return """
    <style>
      /* 카드 래퍼 */
      .ui-card {
        border-radius: 16px;
        border: 1px solid rgba(148,163,184,.25);
        background: rgba(2,6,23,.65);
        box-shadow: 0 10px 40px rgba(0,0,0,.25);
        padding: 18px 18px 12px 18px;
        backdrop-filter: saturate(120%) blur(12px);
        -webkit-backdrop-filter: saturate(120%) blur(12px);
        margin-bottom: 14px;
      }
      .ui-card .ui-title {
        font-weight: 800; font-size: 18px; margin-bottom: 8px; color: #e5e7eb;
      }

      /* Streamlit 업로더 박스 */
      .ui-upload [data-testid="stFileUploaderDropzone"]{
        border-radius: 12px;
        border: 1px solid rgba(148,163,184,.25);
        background: rgba(17,24,39,.55);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
      }
      /* 업로더 내부 여백/텍스트 */
      .ui-upload [data-testid="stFileUploader"] section { gap: 6px; }
      .ui-upload [data-testid="stFileUploader"] button { border-radius: 10px; }

      /* 업로더 라벨 줄여서 */
      .ui-label { font-weight: 700; margin-bottom: 6px; color: #cbd5e1; }
    </style>
    """

def _upload_theme_css(palette: dict, class_name: str) -> str:
    # 업로더 색상 테마(엑셀/워드/중립)
    brand = palette["brand"]
    weak = palette["brand-weak"]
    strong = palette["brand-strong"]
    text = palette["text"]
    return f"""
    <style>
      /* 라벨 색 */
      .{class_name} .ui-label {{ color: {text}; }}

      /* 드롭존 테두리/배경에 브랜드 컬러 가미 */
      .{class_name} [data-testid="stFileUploaderDropzone"] {{
        border-color: {weak};
        background: linear-gradient(0deg, {strong}, rgba(17,24,39,.55));
      }}

      /* 드롭존 hover */
      .{class_name} [data-testid="stFileUploaderDropzone"]:hover {{
        border-color: {brand};
        box-shadow: 0 0 0 2px {weak} inset;
      }}

      /* 업로드 버튼에 브랜드 */
      .{class_name} [data-testid="stFileUploader"] button {{
        background: {brand};
        color: white;
        border: 1px solid {weak};
      }}
      .{class_name} [data-testid="stFileUploader"] button:hover {{
        filter: brightness(1.05);
      }}
    </style>
    """

def inject_base(st, *, max_width: int = 1080):
    """페이지 전역 기본 스타일 주입"""
    st.markdown(_base_css(max_width), unsafe_allow_html=True)
    st.markdown(_glass_css(), unsafe_allow_html=True)

def inject_uploader_skins(st):
    """엑셀/워드/중립 업로더 테마 모두 주입 (필요한 곳에서 클래스만 붙여 쓰면 됨)"""
    for cls, pal in [("upload-excel", PALETTES["excel"]),
                     ("upload-word", PALETTES["word"]),
                     ("upload-neutral", PALETTES["neutral"])]:
        st.markdown(_upload_theme_css(pal, cls), unsafe_allow_html=True)

def open_card(st, title: str | None = None):
    """카드 열기 (div 시작)"""
    st.markdown('<div class="ui-card">', unsafe_allow_html=True)
    if title:
        st.markdown(f'<div class="ui-title">{title}</div>', unsafe_allow_html=True)

def close_card(st):
    """카드 닫기 (div 종료)"""
    st.markdown("</div>", unsafe_allow_html=True)

def open_upload(st, variant: str = "neutral", label: str | None = None):
    """
    업로더 영역 열기. variant: excel | word | neutral
    사용 후 반드시 close_upload 호출.
    """
    cls = {
        "excel": "upload-excel",
        "word": "upload-word",
        "neutral": "upload-neutral",
    }.get(variant, "upload-neutral")
    st.markdown(f'<div class="ui-upload {cls}">', unsafe_allow_html=True)
    if label:
        st.markdown(f'<div class="ui-label">{label}</div>', unsafe_allow_html=True)

def close_upload(st):
    """업로더 영역 닫기"""
    st.markdown("</div>", unsafe_allow_html=True)
