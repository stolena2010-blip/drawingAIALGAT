"""
Brand — Shared theming, logo, and CSS for all DrawingAI Pro Streamlit pages.
Color palette based on the Green Coat / Algat app (orange-green gradient).
Futuristic glassmorphism design with animated accents.
"""
import base64
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
_LOGO_PATH = PROJECT_ROOT / "company_logo.png"


def _logo_b64() -> str:
    if _LOGO_PATH.exists():
        return base64.b64encode(_LOGO_PATH.read_bytes()).decode()
    return ""


# ── Color palette ────────────────────────────────────────────────
ORANGE = "#FF8C00"
ORANGE_LIGHT = "#FFA726"
ORANGE_GLOW = "#ff9f1c"
GREEN = "#4CAF50"
GREEN_LIGHT = "#66BB6A"
GREEN_DARK = "#2E7D32"
TEAL = "#00d4aa"
GOLD = "#FFD700"
CARD_BG = "#0d1117"
CARD_BG_GLASS = "rgba(13, 17, 23, 0.65)"
CARD_BORDER = "#1e2a3a"
DARK_BG = "#0a0e14"
SURFACE = "#111827"
RED = "#ef4444"
BLUE = "#3b82f6"
PURPLE = "#8b5cf6"
YELLOW = "#fbbf24"
CONSOLE_GREEN = "#00ff88"
TEXT_WHITE = "#f0f4f8"
TEXT_MUTED = "#94a3b8"
ACCENT_GRADIENT = f"linear-gradient(135deg, {ORANGE} 0%, {GREEN} 100%)"
GLOW_ORANGE = "rgba(255, 140, 0, 0.15)"
GLOW_GREEN = "rgba(76, 175, 80, 0.12)"


def sidebar_logo() -> None:
    """Render the company logo and brand name in the Streamlit sidebar."""
    import streamlit as st
    logo = _logo_b64()
    if logo:
        st.sidebar.markdown(
            f"""
            <div style="
                display: flex;
                flex-direction: column;
                align-items: center;
                padding: 18px 8px 8px 8px;
            ">
                <div style="
                    position: relative;
                    width: 110px; height: 110px;
                    border-radius: 20px;
                    background: {ACCENT_GRADIENT};
                    padding: 3px;
                    box-shadow: 0 0 25px {GLOW_ORANGE}, 0 0 50px rgba(76,175,80,0.1);
                ">
                    <img src="data:image/png;base64,{logo}"
                         style="width:100%; height:100%; border-radius:18px; object-fit: contain;
                                background: white; padding: 4px;">
                </div>
                <div style="
                    margin-top: 12px;
                    font-size: 15px;
                    font-weight: 800;
                    letter-spacing: 2px;
                    text-transform: uppercase;
                    background: {ACCENT_GRADIENT};
                    -webkit-background-clip: text;
                    -webkit-text-fill-color: transparent;
                    text-align: center;
                    line-height: 1.3;
                ">GREEN COAT ALGAT</div>
                <div style="
                    font-size: 10px;
                    color: {TEXT_MUTED};
                    letter-spacing: 3px;
                    text-transform: uppercase;
                    margin-top: 2px;
                ">DrawingAI Pro</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.sidebar.markdown(
            f'<div style="text-align:center; padding:16px; font-size:15px; font-weight:800; '
            f'letter-spacing:2px; text-transform:uppercase; '
            f'background:{ACCENT_GRADIENT}; '
            f'-webkit-background-clip:text; -webkit-text-fill-color:transparent;">'
            f'GREEN COAT ALGAT</div>',
            unsafe_allow_html=True,
        )
    st.sidebar.markdown("---")


def brand_header(title: str = "DrawingAI Pro", subtitle: str = "") -> str:
    """Return HTML for a branded header with animated gradient background."""

    sub_html = ""
    if subtitle:
        sub_html = f'<div style="font-size:13px; color:rgba(255,255,255,0.75); margin-top:4px; letter-spacing:0.5px;">{subtitle}</div>'

    return f"""
    <div style="
        background: {ACCENT_GRADIENT};
        border-radius: 16px;
        padding: 22px 28px;
        margin-bottom: 20px;
        display: flex;
        align-items: center;
        justify-content: center;
        direction: rtl;
        position: relative;
        overflow: hidden;
        box-shadow: 0 4px 30px {GLOW_ORANGE}, 0 0 60px rgba(76,175,80,0.08);
    ">
        <div style="
            position: absolute; top: -50%; right: -20%;
            width: 300px; height: 300px;
            background: radial-gradient(circle, rgba(255,255,255,0.08) 0%, transparent 70%);
            border-radius: 50%;
        "></div>
        <div style="position: relative; z-index: 1; text-align: center;">
            <div style="font-size:30px; font-weight:800; color:white; letter-spacing:1px;
                        text-shadow: 0 2px 10px rgba(0,0,0,0.3);">{title}</div>
            {sub_html}
        </div>
    </div>
    """


def hero_card(icon: str, title: str, value: str, color: str = ORANGE, subtitle: str = "") -> str:
    """Return HTML for a futuristic hero stat card with glow effect."""
    sub = f'<div style="font-size:11px; color:{TEXT_MUTED}; margin-top:2px;">{subtitle}</div>' if subtitle else ""
    return f"""
    <div style="
        background: {CARD_BG_GLASS};
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(255,140,0,0.15);
        border-radius: 14px;
        padding: 20px;
        text-align: center;
        position: relative;
        overflow: hidden;
        transition: transform 0.2s, box-shadow 0.2s;
        box-shadow: 0 0 20px rgba(0,0,0,0.3);
    ">
        <div style="
            position: absolute; top: 0; left: 0; right: 0; height: 3px;
            background: linear-gradient(90deg, transparent, {color}, transparent);
        "></div>
        <div style="font-size: 28px; margin-bottom: 6px;">{icon}</div>
        <div style="font-size: 28px; font-weight: 800; color: {color};
                    text-shadow: 0 0 20px {color}44;">{value}</div>
        <div style="font-size: 12px; color: {TEXT_MUTED}; margin-top: 4px;
                    text-transform: uppercase; letter-spacing: 1px;">{title}</div>
        {sub}
    </div>
    """


BRAND_CSS = f"""
<style>
    /* ── Keyframes ── */
    @keyframes gradientShift {{
        0% {{ background-position: 0% 50%; }}
        50% {{ background-position: 100% 50%; }}
        100% {{ background-position: 0% 50%; }}
    }}
    @keyframes pulseGlow {{
        0%, 100% {{ box-shadow: 0 0 15px {GLOW_ORANGE}; }}
        50% {{ box-shadow: 0 0 25px {GLOW_ORANGE}, 0 0 40px {GLOW_GREEN}; }}
    }}
    @keyframes fadeInUp {{
        from {{ opacity: 0; transform: translateY(10px); }}
        to {{ opacity: 1; transform: translateY(0); }}
    }}

    /* ── RTL ── */
    .stApp, .stMarkdown, .stText, .stAlert,
    [data-testid="stExpander"],
    .stTextInput label, .stSelectbox label, .stCheckbox label,
    .stRadio label, .stNumberInput label, .stMultiSelect label,
    .stTabs [data-baseweb="tab-list"] {{
        direction: rtl; text-align: right;
    }}
    h1, h2, h3, h4 {{ direction: rtl; text-align: right; }}

    /* ── Global background enhancement ── */
    .stApp {{
        background: {DARK_BG};
        background-image:
            radial-gradient(ellipse at 20% 50%, rgba(255,140,0,0.04) 0%, transparent 50%),
            radial-gradient(ellipse at 80% 20%, rgba(76,175,80,0.03) 0%, transparent 50%);
    }}

    /* ── Hide default Streamlit title (we use branded header) ── */
    header[data-testid="stHeader"] {{ background: transparent; }}

    /* ── Metric cards: glassmorphism ── */
    [data-testid="stMetric"] {{
        background: {CARD_BG_GLASS};
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(255,140,0,0.12);
        border-radius: 14px;
        padding: 18px;
        position: relative;
        overflow: hidden;
        animation: fadeInUp 0.4s ease-out;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }}
    [data-testid="stMetric"]:hover {{
        transform: translateY(-2px);
        box-shadow: 0 8px 30px rgba(255,140,0,0.12);
        border-color: rgba(255,140,0,0.25);
    }}
    [data-testid="stMetric"]::before {{
        content: "";
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 2px;
        background: {ACCENT_GRADIENT};
        opacity: 0.6;
    }}
    [data-testid="stMetric"] [data-testid="stMetricValue"] {{
        color: {ORANGE_LIGHT};
        font-weight: 700;
        font-size: 1.6rem !important;
    }}
    [data-testid="stMetric"] [data-testid="stMetricLabel"] {{
        color: {TEXT_MUTED};
        font-size: 0.8rem !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}
    [data-testid="stMetric"] [data-testid="stMetricDelta"] {{
        font-weight: 600;
    }}

    /* ── Tabs: modern underline ── */
    .stTabs [data-baseweb="tab-list"] {{
        background: {CARD_BG_GLASS};
        backdrop-filter: blur(8px);
        border-radius: 12px;
        padding: 4px;
        border: 1px solid rgba(255,140,0,0.08);
        gap: 2px;
    }}
    .stTabs [data-baseweb="tab"] {{
        color: {TEXT_MUTED};
        border-radius: 8px;
        padding: 8px 16px;
        font-weight: 500;
        transition: all 0.2s ease;
    }}
    .stTabs [data-baseweb="tab"]:hover {{
        color: {TEXT_WHITE};
        background: rgba(255,140,0,0.06);
    }}
    .stTabs [aria-selected="true"] {{
        background: rgba(255,140,0,0.12) !important;
        border-bottom-color: {ORANGE} !important;
        color: {ORANGE_LIGHT} !important;
        font-weight: 700;
        box-shadow: 0 0 15px rgba(255,140,0,0.08);
    }}

    /* ── Buttons: gradient + glow ── */
    .stButton > button {{
        border-radius: 10px;
        font-weight: 600;
        border: 1px solid rgba(255,140,0,0.25);
        background: {CARD_BG_GLASS};
        backdrop-filter: blur(8px);
        color: {TEXT_WHITE};
        transition: all 0.25s ease;
        letter-spacing: 0.3px;
    }}
    .stButton > button:hover {{
        border-color: {ORANGE};
        box-shadow: 0 0 20px rgba(255,140,0,0.15);
        transform: translateY(-1px);
        color: {ORANGE_LIGHT};
    }}
    .stButton > button:active {{
        transform: translateY(0px);
    }}

    /* ── Primary buttons ── */
    .stButton > button[kind="primary"],
    .stButton > button[data-testid="stBaseButton-primary"] {{
        background: {ACCENT_GRADIENT} !important;
        border: none !important;
        color: white !important;
        font-weight: 700;
    }}
    .stButton > button[kind="primary"]:hover,
    .stButton > button[data-testid="stBaseButton-primary"]:hover {{
        box-shadow: 0 4px 25px rgba(255,140,0,0.3);
    }}

    /* ── Status indicators ── */
    .status-running {{ color: {GREEN_LIGHT}; font-weight: bold; text-shadow: 0 0 8px rgba(102,187,106,0.4); }}
    .status-stopped {{ color: {RED}; font-weight: bold; }}
    .status-processing {{ color: {YELLOW}; font-weight: bold; text-shadow: 0 0 8px rgba(251,191,36,0.3); }}

    /* ── Run-type badge ── */
    .run-badge {{
        display: inline-block;
        padding: 6px 18px;
        border-radius: 20px;
        font-size: 15px;
        font-weight: 700;
        direction: rtl;
    }}
    .run-badge-active {{
        background: linear-gradient(135deg, rgba(102,187,106,0.22) 0%, rgba(0,200,83,0.12) 100%);
        border: 1.5px solid rgba(102,187,106,0.5);
        color: {GREEN_LIGHT};
        text-shadow: 0 0 10px rgba(102,187,106,0.5);
        animation: badgePulse 2s ease-in-out infinite;
    }}
    .run-badge-inactive {{
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.08);
        color: rgba(255,255,255,0.35);
    }}
    .run-badge-heavy {{
        background: linear-gradient(135deg, rgba(225,112,85,0.22) 0%, rgba(214,48,49,0.12) 100%);
        border: 1.5px solid rgba(225,112,85,0.5);
        color: #e17055;
        text-shadow: 0 0 10px rgba(225,112,85,0.5);
        animation: badgePulse 2s ease-in-out infinite;
    }}
    @keyframes badgePulse {{
        0%, 100% {{ opacity: 1; }}
        50% {{ opacity: 0.75; }}
    }}
    .email-progress {{
        display: inline-block;
        padding: 4px 14px;
        border-radius: 14px;
        font-size: 18px;
        font-weight: 800;
        color: #ffd93d;
        background: rgba(255,217,61,0.10);
        border: 1.5px solid rgba(255,217,61,0.35);
        margin-right: 10px;
        direction: ltr;
        letter-spacing: 1px;
    }}

    /* ── Log viewer: terminal feel ── */
    .log-container {{
        background: linear-gradient(180deg, #0d1117 0%, #0a0e14 100%);
        color: {CONSOLE_GREEN};
        font-family: 'JetBrains Mono', 'Cascadia Code', Consolas, monospace;
        font-size: 12px;
        padding: 16px;
        border-radius: 12px;
        height: 420px;
        overflow-y: scroll;
        direction: ltr;
        text-align: left;
        white-space: pre-wrap;
        border: 1px solid rgba(0,255,136,0.08);
        box-shadow: inset 0 0 30px rgba(0,0,0,0.5), 0 0 20px rgba(0,255,136,0.03);
    }}

    /* ── Status bar (above log) ── */
    .status-bar {{
        background: {CARD_BG_GLASS};
        backdrop-filter: blur(12px);
        border: 1px solid rgba(255,140,0,0.12);
        border-radius: 12px;
        padding: 12px 18px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        direction: rtl;
        font-size: 14px;
        margin-bottom: 10px;
        animation: pulseGlow 3s ease-in-out infinite;
    }}
    .status-bar .phase {{ color: {GREEN_LIGHT}; font-weight: bold; text-shadow: 0 0 6px rgba(102,187,106,0.3); }}
    .status-bar .progress {{ color: {GOLD}; }}
    .status-bar .timer {{ color: {ORANGE_LIGHT}; font-family: 'JetBrains Mono', Consolas, monospace; }}

    /* ── Subheaders ── */
    .stApp h2 {{
        color: {TEXT_WHITE} !important;
        font-weight: 700;
        position: relative;
        padding-bottom: 8px;
    }}
    .stApp h3 {{
        color: {ORANGE_LIGHT};
        font-weight: 600;
    }}

    /* ── Accuracy color helpers ── */
    .card-green {{ color: {GREEN_LIGHT}; font-weight: bold; text-shadow: 0 0 8px rgba(102,187,106,0.3); }}
    .card-yellow {{ color: {YELLOW}; font-weight: bold; text-shadow: 0 0 8px rgba(251,191,36,0.3); }}
    .card-red {{ color: {RED}; font-weight: bold; text-shadow: 0 0 8px rgba(239,68,68,0.3); }}

    /* ── Sidebar ── */
    [data-testid="stSidebar"] {{
        background: linear-gradient(180deg, #0d1117 0%, #0a0e14 100%);
        border-right: 1px solid rgba(255,140,0,0.1);
    }}
    [data-testid="stSidebar"] [data-testid="stSidebarNav"] {{
        padding-top: 0;
    }}
    [data-testid="stSidebar"] > div:first-child {{
        display: flex;
        flex-direction: column;
    }}
    [data-testid="stSidebar"] [data-testid="stSidebarNav"] {{
        order: 2;
    }}
    [data-testid="stSidebar"] [data-testid="stSidebarNav"] a {{
        font-size: 16px !important;
        padding: 8px 12px !important;
        border-radius: 8px;
        transition: all 0.2s ease;
        margin: 2px 8px;
    }}
    [data-testid="stSidebar"] [data-testid="stSidebarNav"] a:hover {{
        background: rgba(255,140,0,0.08);
    }}
    [data-testid="stSidebar"] [data-testid="stSidebarNav"] span {{
        font-size: 16px !important;
    }}
    [data-testid="stSidebar"] .stMarkdown h1 {{
        background: {ACCENT_GRADIENT};
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }}

    /* ── Expander ── */
    [data-testid="stExpander"] {{
        background: {CARD_BG_GLASS};
        backdrop-filter: blur(8px);
        border: 1px solid rgba(255,140,0,0.08);
        border-radius: 12px;
        overflow: hidden;
    }}
    [data-testid="stExpander"]:hover {{
        border-color: rgba(255,140,0,0.18);
    }}

    /* ── Dataframe styling ── */
    .stDataFrame {{
        border: 1px solid rgba(255,140,0,0.1);
        border-radius: 12px;
        overflow: hidden;
    }}
    .stDataFrame [data-testid="stDataFrameResizable"] {{
        border-radius: 12px;
    }}

    /* ── Info / Warning / Error boxes ── */
    .stAlert {{
        border-radius: 12px;
        backdrop-filter: blur(8px);
    }}

    /* ── Selectbox / Inputs ── */
    .stSelectbox [data-baseweb="select"],
    .stTextInput input,
    .stNumberInput input {{
        border-radius: 8px !important;
        border-color: rgba(255,140,0,0.15) !important;
        background: rgba(13,17,23,0.6) !important;
    }}
    .stSelectbox [data-baseweb="select"]:hover,
    .stTextInput input:focus,
    .stNumberInput input:focus {{
        border-color: {ORANGE} !important;
        box-shadow: 0 0 10px rgba(255,140,0,0.1) !important;
    }}

    /* ── Radio buttons (period filter) ── */
    .stRadio [role="radiogroup"] {{
        gap: 4px;
    }}
    .stRadio [role="radiogroup"] label {{
        border-radius: 8px;
        padding: 4px 10px;
        transition: all 0.2s;
    }}

    /* ── Dividers ── */
    hr {{
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(255,140,0,0.2), transparent);
        margin: 16px 0;
    }}

    /* ── Scrollbar ── */
    ::-webkit-scrollbar {{
        width: 6px;
        height: 6px;
    }}
    ::-webkit-scrollbar-track {{
        background: {DARK_BG};
    }}
    ::-webkit-scrollbar-thumb {{
        background: rgba(255,140,0,0.2);
        border-radius: 3px;
    }}
    ::-webkit-scrollbar-thumb:hover {{
        background: rgba(255,140,0,0.4);
    }}
</style>
"""
