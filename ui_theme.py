"""Shared Streamlit UI theme for NAV_FETCHER apps."""

from __future__ import annotations

import streamlit as st


def inject_custom_css() -> None:
    st.markdown(
        """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;1,9..40,400&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
    --bg-deep: #080d14;
    --bg-surface: #0f1724;
    --bg-elevated: #162032;
    --bg-card: rgba(22, 32, 50, 0.72);
    --border-subtle: rgba(99, 179, 237, 0.12);
    --border-accent: rgba(56, 189, 248, 0.35);
    --text-primary: #f1f5f9;
    --text-secondary: #94a3b8;
    --text-muted: #64748b;
    --accent-blue: #38bdf8;
    --accent-emerald: #34d399;
    --gradient-hero: linear-gradient(135deg, #0c1929 0%, #132238 45%, #0f1f33 100%);
    --gradient-accent: linear-gradient(90deg, #38bdf8, #34d399);
    --shadow-lg: 0 20px 50px rgba(0, 0, 0, 0.45);
    --radius-lg: 16px;
    --radius-md: 12px;
}

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 1.5rem; padding-bottom: 3rem; max-width: 1280px; }

.stApp {
    background:
        radial-gradient(ellipse 80% 50% at 50% -20%, rgba(56, 189, 248, 0.08), transparent),
        radial-gradient(ellipse 60% 40% at 100% 0%, rgba(52, 211, 153, 0.05), transparent),
        linear-gradient(180deg, var(--bg-deep) 0%, var(--bg-surface) 40%, #0a1018 100%);
}

.hero-card {
    background: var(--gradient-hero);
    border: 1px solid var(--border-subtle);
    border-radius: var(--radius-lg);
    padding: 2.25rem 2.5rem;
    margin-bottom: 1.75rem;
    box-shadow: var(--shadow-lg);
    position: relative;
    overflow: hidden;
}
.hero-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 1px;
    background: linear-gradient(90deg, transparent, var(--accent-blue), var(--accent-emerald), transparent);
    opacity: 0.6;
}
.hero-badge {
    display: inline-block;
    background: rgba(56, 189, 248, 0.12);
    border: 1px solid rgba(56, 189, 248, 0.3);
    color: var(--accent-blue);
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    padding: 0.35rem 0.85rem;
    border-radius: 100px;
    margin-bottom: 0.85rem;
}
.hero-title {
    font-size: 2.1rem;
    font-weight: 700;
    letter-spacing: -0.02em;
    background: var(--gradient-accent);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    margin: 0 0 0.5rem 0;
    line-height: 1.2;
}
.hero-sub {
    color: var(--text-secondary);
    font-size: 1.02rem;
    margin: 0 0 1.25rem 0;
    line-height: 1.55;
    max-width: 680px;
}
.stat-row { display: flex; gap: 10px; flex-wrap: wrap; }
.stat-chip {
    background: rgba(15, 23, 36, 0.6);
    border: 1px solid var(--border-subtle);
    border-radius: 100px;
    padding: 0.45rem 1rem;
    color: var(--text-secondary);
    font-size: 0.82rem;
    font-weight: 500;
}
.stat-chip strong { color: var(--accent-emerald); font-weight: 600; }

.section-header {
    display: flex;
    align-items: center;
    gap: 0.65rem;
    margin: 1.5rem 0 1rem 0;
}
.section-icon {
    width: 36px; height: 36px;
    background: rgba(56, 189, 248, 0.1);
    border: 1px solid var(--border-subtle);
    border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    font-size: 1rem;
}
.section-title {
    color: var(--text-primary);
    font-size: 1.15rem;
    font-weight: 600;
    margin: 0;
    letter-spacing: -0.01em;
}
.section-desc {
    color: var(--text-muted);
    font-size: 0.82rem;
    margin: 0;
}

.info-card {
    background: var(--bg-card);
    backdrop-filter: blur(12px);
    border: 1px solid var(--border-subtle);
    border-left: 3px solid var(--accent-blue);
    border-radius: var(--radius-md);
    padding: 1rem 1.25rem;
    margin-bottom: 1.25rem;
    color: var(--text-secondary);
    font-size: 0.9rem;
    line-height: 1.65;
}
.info-card strong { color: var(--text-primary); }

.metrics-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 12px;
    margin-bottom: 1.5rem;
}
@media (max-width: 900px) { .metrics-grid { grid-template-columns: repeat(2, 1fr); } }
.metric-card {
    background: var(--bg-card);
    border: 1px solid var(--border-subtle);
    border-radius: var(--radius-md);
    padding: 1.1rem 1.25rem;
    transition: border-color 0.2s;
}
.metric-card:hover { border-color: var(--border-accent); }
.metric-label {
    display: block;
    color: var(--text-muted);
    font-size: 0.75rem;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-bottom: 0.35rem;
}
.metric-value {
    display: block;
    color: var(--text-primary);
    font-size: 1.05rem;
    font-weight: 600;
    line-height: 1.3;
    word-break: break-word;
}
.metric-value.accent { color: var(--accent-emerald); font-family: 'JetBrains Mono', monospace; }

.panel-divider {
    border: none;
    border-top: 1px solid var(--border-subtle);
    margin: 1.75rem 0;
}

.stButton > button[kind="primary"], .stButton > button[data-testid="stBaseButton-primary"] {
    background: linear-gradient(135deg, #0ea5e9, #0284c7) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    padding: 0.55rem 1.4rem !important;
    box-shadow: 0 4px 16px rgba(14, 165, 233, 0.35) !important;
    transition: transform 0.15s, box-shadow 0.15s !important;
}
.stButton > button[kind="primary"]:hover, .stButton > button[data-testid="stBaseButton-primary"]:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 22px rgba(14, 165, 233, 0.45) !important;
}
.stButton > button[kind="secondary"], .stButton > button:not([kind="primary"]) {
    background: rgba(22, 32, 50, 0.8) !important;
    border: 1px solid var(--border-subtle) !important;
    color: var(--text-secondary) !important;
    border-radius: 10px !important;
    font-weight: 500 !important;
}
.stDownloadButton > button {
    background: linear-gradient(135deg, #059669, #047857) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    box-shadow: 0 4px 14px rgba(5, 150, 105, 0.35) !important;
}

.stTextInput input, .stTextArea textarea, .stDateInput input, .stNumberInput input {
    background: rgba(12, 20, 32, 0.85) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: 10px !important;
    color: var(--text-primary) !important;
}
.stTextInput input:focus, .stTextArea textarea:focus {
    border-color: var(--accent-blue) !important;
    box-shadow: 0 0 0 3px rgba(56, 189, 248, 0.12) !important;
}

.stRadio > div { gap: 8px; flex-wrap: wrap; }
.stRadio label {
    background: rgba(15, 23, 36, 0.7) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: 10px !important;
    padding: 0.5rem 1.1rem !important;
    color: var(--text-secondary) !important;
    font-weight: 500 !important;
    transition: all 0.15s !important;
}
.stRadio label:has(input:checked) {
    background: rgba(14, 165, 233, 0.15) !important;
    border-color: var(--accent-blue) !important;
    color: var(--accent-blue) !important;
}

.stCheckbox label { color: var(--text-secondary) !important; font-size: 0.9rem !important; }

.stMultiSelect [data-baseweb="tag"] {
    background: rgba(14, 165, 233, 0.15) !important;
    border: 1px solid rgba(56, 189, 248, 0.3) !important;
    color: var(--accent-blue) !important;
}

div[data-testid="stMetric"] {
    background: var(--bg-card);
    border: 1px solid var(--border-subtle);
    border-radius: var(--radius-md);
    padding: 0.85rem 1rem;
}
div[data-testid="stMetric"] label { color: var(--text-muted) !important; font-size: 0.78rem !important; }
div[data-testid="stMetric"] [data-testid="stMetricValue"] { color: var(--text-primary) !important; }

.stDataFrame, [data-testid="stDataFrame"] {
    border: 1px solid var(--border-subtle) !important;
    border-radius: var(--radius-md) !important;
    overflow: hidden;
}

.stSuccess, .stWarning, .stError, .stInfo {
    border-radius: 10px !important;
    border-width: 1px !important;
}
.stSuccess { background: rgba(52, 211, 153, 0.08) !important; border-color: rgba(52, 211, 153, 0.25) !important; }
.stWarning { background: rgba(251, 191, 36, 0.08) !important; border-color: rgba(251, 191, 36, 0.25) !important; }
.stError { background: rgba(248, 113, 113, 0.08) !important; border-color: rgba(248, 113, 113, 0.25) !important; }
.stInfo { background: rgba(56, 189, 248, 0.08) !important; border-color: rgba(56, 189, 248, 0.25) !important; }

label, .stMarkdown p { color: var(--text-secondary) !important; }
h1, h2, h3, h4, [data-testid="stMarkdownContainer"] h3 {
    color: var(--text-primary) !important;
    font-weight: 600 !important;
    letter-spacing: -0.01em !important;
}

details[data-testid="stExpander"] {
    background: var(--bg-card) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: var(--radius-md) !important;
}
details[data-testid="stExpander"] summary { color: var(--text-primary) !important; font-weight: 500 !important; }

.stSpinner > div { color: var(--accent-blue) !important; }

.app-footer {
    text-align: center;
    color: var(--text-muted);
    font-size: 0.78rem;
    margin-top: 2.5rem;
    padding-top: 1.5rem;
    border-top: 1px solid var(--border-subtle);
}
</style>
""",
        unsafe_allow_html=True,
    )


def render_hero(
    title: str,
    subtitle: str,
    badge: str = "Live AMFI Data",
    chips: list[str] | None = None,
) -> None:
    chips_html = ""
    if chips:
        chips_html = '<div class="stat-row">' + "".join(f'<span class="stat-chip">{chip}</span>' for chip in chips) + "</div>"
    st.markdown(
        f"""
        <div class="hero-card">
            <span class="hero-badge">{badge}</span>
            <h1 class="hero-title">{title}</h1>
            <p class="hero-sub">{subtitle}</p>
            {chips_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_section_header(icon: str, title: str, description: str = "") -> None:
    desc_html = f'<p class="section-desc">{description}</p>' if description else ""
    st.markdown(
        f"""
        <div class="section-header">
            <div class="section-icon">{icon}</div>
            <div>
                <p class="section-title">{title}</p>
                {desc_html}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_info_card(content: str) -> None:
    st.markdown(f'<div class="info-card">{content}</div>', unsafe_allow_html=True)


def render_app_footer() -> None:
    st.markdown(
        '<div class="app-footer">Data sourced from AMFI India · Built for research &amp; portfolio analysis</div>',
        unsafe_allow_html=True,
    )
