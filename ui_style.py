# ui_style.py
import streamlit as st

EXCEL_GREEN = "#22c55e"   # 밝은 네온 그린
WORD_BLUE   = "#60a5fa"   # 밝은 네온 블루

BASE_CSS = f"""
/* ---------- Reset / Base ---------- */
#MainMenu, footer {{visibility: hidden;}}
.block-container {{
  padding-top: 1rem;
  max-width: 1080px;
}}
html, body, [data-testid="stAppViewContainer"] {{
  background: radial-gradient(1200px 800px at 10% 10%, rgba(59,130,246,.08) 0%, rgba(2,6,23,1) 55%) no-repeat,
              radial-gradient(900px 700px at 90% 90%, rgba(34,197,94,.06) 0%, rgba(2,6,23,1) 60%) no-repeat,
              #020617;
  color: rgba(241,245,249,.96);
}}

/* ---------- Dashboard Header ---------- */
.page-title {{
  font-size: 26px; font-weight: 800; letter-spacing:.2px;
}}
.page-sub {{
  margin-top:.25rem; color: rgba(148,163,184,.9);
}}

/* ---------- Card ---------- */
.card {{
  background: rgba(15,23,42,.55);
  border: 1px solid rgba(148,163,184,.18);
  border-radius: 18px;
  padding: 22px;
  box-shadow: 0 12px 60px rgba(0,0,0,.35), inset 0 1px 0 rgba(255,255,255,.02);
  backdrop-filter: blur(8px);
  transition: transform .2s ease, box-shadow .2s ease, border-color .2s ease;
}}
.card:hover {{
  transform: translateY(-3px);
  box-shadow: 0 18px 70px rgba(0,0,0,.45);
  border-color: rgba(148,163,184,.28);
}}

/* ---------- Upload cards (icon + heading) ---------- */
.upload-head {{
  display:flex; gap:12px; align-items:center; margin-bottom:10px;
}}
.upload-icon {{
  width:36px; height:36px; display:grid; place-items:center;
  border-radius:12px; font-weight:700;
  background: rgba(2,6,23,.6); border:1px solid rgba(148,163,184,.2);
}}
.upload-title {{ font-weight:800; font-size:15px; letter-spacing:.2px; }}
.upload-sub {{ font-size:12.5px; color:rgba(148,163,184,.9); margin-top:2px; }}

/* Excel / Word accent ring */
.card.excel {{ box-shadow: 0 0 0 1px rgba(34,197,94,.35) inset, 0 10px 50px rgba(34,197,94,.10); }}
.card.word  {{ box-shadow: 0 0 0 1px rgba(96,165,250,.35) inset, 0 10px 50px rgba(96,165,250,.10); }}

/* ---------- Streamlit uploader zone ---------- */
[data-testid="stFileUploader"] section {{ gap:10px !important; }}
[data-testid="stFileUploaderDropzone"] {{
  background: linear-gradient(180deg, rgba(17,24,39,.55), rgba(2,6,23,.55)) !important;
  border: 2px dashed rgba(148,163,184,.30) !important;
  border-radius: 14px !important;
  min-height: 140px !important;
  transition: all .25s ease !important;
}}
.excel [data-testid="stFileUploaderDropzone"] {{
  border-color: rgba(34,197,94,.45) !important;
}}
.word [data-testid="stFileUploaderDropzone"] {{
  border-color: rgba(96,165,250,.45) !important;
}}
[data-testid="stFileUploaderDropzone"]:hover {{
  transform: translateY(-2px) !important;
  box-shadow: 0 12px 40px rgba(0,0,0,.35) !important;
}}
.excel [data-testid="stFileUploaderDropzone"]:hover {{
  box-shadow: 0 0 0 1px rgba(34,197,94,.35) inset, 0 14px 60px rgba(34,197,94,.18) !important;
}}
.word [data-testid="stFileUploaderDropzone"]:hover {{
  box-shadow: 0 0 0 1px rgba(96,165,250,.35) inset, 0 14px 60px rgba(96,165,250,.18) !important;
}}
[data-testid="stFileUploaderDropzone"] p {{ color: rgba(203,213,225,.95) !important; }}
[data-testid="stFileUploaderDropzone"] small {{ color: rgba(148,163,184,.9) !important; }}

/* Buttons */
.stButton>button, [data-testid="stFileUploader"] button {{
  border-radius: 10px !important;
  border:1px solid rgba(148,163,184,.25) !important;
  background: rgba(30,41,59,.65) !important;
  color: #e5e7eb !important;
  padding: 8px 18px !important;
}}
.stButton>button:hover, [data-testid="stFileUploader"] button:hover {{
  transform: translateY(-1px);
  box-shadow: 0 8px 24px rgba(0,0,0,.35) !important;
}}

/* ---------- Status ---------- */
.status-wrap {{ display:flex; gap:12px; align-items:center; }}
.badge {{
  display:inline-flex; gap:6px; align-items:center;
  padding:6px 10px; border-radius:999px; font-size:12.5px; font-weight:700;
  border:1px solid rgba(148,163,184,.25); background: rgba(15,23,42,.6);
}}
.badge .dot {{width:8px;height:8px;border-radius:999px; display:inline-block;}
.badge.ok .dot {{background: #22c55e;}}
.badge.wait .dot {{background: #facc15;}}
.badge.err .dot {{background: #ef4444;}}

/* progress (override) */
[data-testid="stProgress"] > div > div {{
  background: rgba(15,23,42,.6); border-radius: 999px; border:1px solid rgba(148,163,184,.25);
}}
[data-testid="stProgress"] > div > div > div {{
  border-radius: 999px; background: linear-gradient(90deg, {EXCEL_GREEN}, {WORD_BLUE});
}}

/* Sidebar icons */
[data-testid="stSidebar"] {{
  background: rgba(15,23,42,.65);
  border-right:1px solid rgba(148,163,184,.18);
}}
.sidebar-item {{
  display:flex; align-items:center; gap:10px; padding:10px 12px; border-radius:12px;
}}
.sidebar-item:hover {{ background: rgba(30,41,59,.65); }}
.sidebar-ico {{
  width:28px;height:28px;border-radius:10px; display:grid; place-items:center;
  background: rgba(2,6,23,.6); border:1px solid rgba(148,163,184,.2);
}}

/* scrollbar */
::-webkit-scrollbar {{ width: 8px; height: 8px; }}
::-webkit-scrollbar-thumb {{ background: rgba(148,163,184,.35); border-radius: 8px; }}
"""

def inject():
    st.markdown(f"<style>{BASE_CSS}</style>", unsafe_allow_html=True)

def h4(text: str):
    st.markdown(f'<div class="upload-title">{text}</div>', unsafe_allow_html=True)

def sub(text: str):
    st.markdown(f'<div class="upload-sub">{text}</div>', unsafe_allow_html=True)

def excel_icon():
    return f"""
    <div class="upload-icon" style="border-color:rgba(34,197,94,.35); box-shadow:0 0 24px rgba(34,197,94,.25) inset;">
      <span style="color:{EXCEL_GREEN};font-weight:900;">X</span>
    </div>
    """

def word_icon():
    return f"""
    <div class="upload-icon" style="border-color:rgba(96,165,250,.35); box-shadow:0 0 24px rgba(96,165,250,.25) inset;">
      <span style="color:{WORD_BLUE};font-weight:900;">W</span>
    </div>
    """
