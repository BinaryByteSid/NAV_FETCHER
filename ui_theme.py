"""Shared Streamlit UI theme — sleek finance / market terminal aesthetic."""

from __future__ import annotations

from contextlib import contextmanager

import streamlit as st


def inject_custom_css() -> None:
    st.markdown(
        """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

:root {
    --bg-void: #05080d;
    --bg-base: #0a0f16;
    --bg-surface: #111827;
    --bg-elevated: #1a2332;
    --bg-panel: rgba(17, 24, 39, 0.82);
    --bg-glass: rgba(15, 22, 35, 0.65);
    --border-faint: rgba(148, 163, 184, 0.08);
    --border-subtle: rgba(148, 163, 184, 0.14);
    --border-accent: rgba(212, 175, 55, 0.35);
    --text-primary: #f8fafc;
    --text-secondary: #94a3b8;
    --text-muted: #64748b;
    --gold: #d4af37;
    --gold-soft: rgba(212, 175, 55, 0.15);
    --cyan: #22d3ee;
    --cyan-soft: rgba(34, 211, 238, 0.12);
    --gain: #10b981;
    --gain-soft: rgba(16, 185, 129, 0.14);
    --loss: #ef4444;
    --loss-soft: rgba(239, 68, 68, 0.12);
    --gradient-brand: linear-gradient(135deg, #d4af37 0%, #f5d061 45%, #22d3ee 100%);
    --gradient-panel: linear-gradient(160deg, rgba(26, 35, 50, 0.95) 0%, rgba(10, 15, 22, 0.98) 100%);
    --shadow-panel: 0 4px 24px rgba(0, 0, 0, 0.35), inset 0 1px 0 rgba(255, 255, 255, 0.04);
    --radius-xl: 18px;
    --radius-lg: 14px;
    --radius-md: 10px;
    --radius-sm: 8px;
}

html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    -webkit-font-smoothing: antialiased;
}

#MainMenu, footer, header { visibility: hidden; }
.block-container {
    padding-top: 1.25rem;
    padding-bottom: 3.5rem;
    max-width: 1320px;
}

.stApp {
    background-color: var(--bg-void);
    background-image:
        linear-gradient(rgba(34, 211, 238, 0.03) 1px, transparent 1px),
        linear-gradient(90deg, rgba(34, 211, 238, 0.03) 1px, transparent 1px),
        radial-gradient(ellipse 90% 60% at 50% -30%, rgba(212, 175, 55, 0.07), transparent 55%),
        radial-gradient(ellipse 50% 40% at 100% 20%, rgba(34, 211, 238, 0.05), transparent),
        linear-gradient(180deg, var(--bg-void) 0%, var(--bg-base) 50%, #070b10 100%);
    background-size: 48px 48px, 48px 48px, 100% 100%, 100% 100%, 100% 100%;
}

/* ── Top bar ── */
.top-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 1rem;
    padding: 0.65rem 1.1rem;
    margin-bottom: 1rem;
    background: var(--bg-glass);
    backdrop-filter: blur(16px);
    border: 1px solid var(--border-faint);
    border-radius: var(--radius-lg);
}
.top-bar-brand {
    display: flex;
    align-items: center;
    gap: 0.65rem;
}
.brand-mark {
    width: 34px;
    height: 34px;
    border-radius: 9px;
    background: linear-gradient(135deg, #1e293b, #0f172a);
    border: 1px solid var(--border-accent);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1rem;
    box-shadow: 0 0 20px rgba(212, 175, 55, 0.15);
}
.brand-name {
    font-size: 0.92rem;
    font-weight: 700;
    letter-spacing: -0.02em;
    color: var(--text-primary);
}
.brand-tag {
    font-size: 0.68rem;
    font-weight: 500;
    color: var(--text-muted);
    letter-spacing: 0.04em;
    text-transform: uppercase;
}
.market-status {
    display: flex;
    align-items: center;
    gap: 0.45rem;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    color: var(--gain);
}
.pulse-dot {
    width: 7px;
    height: 7px;
    border-radius: 50%;
    background: var(--gain);
    box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.5);
    animation: pulse-ring 2s infinite;
}
@keyframes pulse-ring {
    0% { box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.45); }
    70% { box-shadow: 0 0 0 8px rgba(16, 185, 129, 0); }
    100% { box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); }
}

/* ── Hero ── */
.hero-card {
    background: var(--gradient-panel);
    border: 1px solid var(--border-subtle);
    border-radius: var(--radius-xl);
    padding: 2rem 2.25rem 1.75rem;
    margin-bottom: 1rem;
    box-shadow: var(--shadow-panel);
    position: relative;
    overflow: hidden;
}
.hero-card::before {
    content: '';
    position: absolute;
    inset: 0;
    background: linear-gradient(105deg, transparent 40%, rgba(212, 175, 55, 0.04) 100%);
    pointer-events: none;
}
.hero-card::after {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: var(--gradient-brand);
    opacity: 0.75;
}
.hero-badge {
    display: inline-flex;
    align-items: center;
    gap: 0.4rem;
    background: var(--gold-soft);
    border: 1px solid rgba(212, 175, 55, 0.28);
    color: var(--gold);
    font-size: 0.68rem;
    font-weight: 700;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    padding: 0.32rem 0.75rem;
    border-radius: 100px;
    margin-bottom: 0.9rem;
}
.hero-title {
    font-size: clamp(1.65rem, 3vw, 2.15rem);
    font-weight: 700;
    letter-spacing: -0.03em;
    color: var(--text-primary);
    margin: 0 0 0.55rem 0;
    line-height: 1.15;
}
.hero-title span {
    background: var(--gradient-brand);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}
.hero-sub {
    color: var(--text-secondary);
    font-size: 0.98rem;
    margin: 0 0 1.35rem 0;
    line-height: 1.6;
    max-width: 720px;
}
.stat-row {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 10px;
}
.stat-chip {
    background: rgba(5, 8, 13, 0.55);
    border: 1px solid var(--border-faint);
    border-radius: var(--radius-md);
    padding: 0.7rem 1rem;
    display: flex;
    flex-direction: column;
    gap: 0.2rem;
    transition: border-color 0.2s, transform 0.2s;
}
.stat-chip:hover {
    border-color: var(--border-accent);
    transform: translateY(-1px);
}
.stat-chip-label {
    font-size: 0.68rem;
    font-weight: 600;
    letter-spacing: 0.07em;
    text-transform: uppercase;
    color: var(--text-muted);
}
.stat-chip-value {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.95rem;
    font-weight: 600;
    color: var(--text-primary);
}
.stat-chip-value.positive { color: var(--gain); }
.stat-chip-value.gold { color: var(--gold); }

/* ── Ticker strip ── */
.ticker-strip {
    display: flex;
    overflow: hidden;
    gap: 0;
    margin-bottom: 1.25rem;
    background: rgba(5, 8, 13, 0.7);
    border: 1px solid var(--border-faint);
    border-radius: var(--radius-md);
    padding: 0.55rem 0;
    mask-image: linear-gradient(90deg, transparent, black 4%, black 96%, transparent);
}
.ticker-track {
    display: flex;
    gap: 2rem;
    animation: ticker-scroll 40s linear infinite;
    white-space: nowrap;
    padding: 0 1rem;
}
.ticker-item {
    display: inline-flex;
    align-items: center;
    gap: 0.5rem;
    font-size: 0.78rem;
    color: var(--text-secondary);
}
.ticker-symbol {
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    color: var(--text-primary);
}
.ticker-change.up { color: var(--gain); font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem; }
.ticker-change.down { color: var(--loss); font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem; }
.ticker-sep { color: var(--border-subtle); }
@keyframes ticker-scroll {
    0% { transform: translateX(0); }
    100% { transform: translateX(-50%); }
}

/* ── Panels ── */
.content-panel {
    background: var(--bg-panel);
    backdrop-filter: blur(12px);
    border: 1px solid var(--border-subtle);
    border-radius: var(--radius-xl);
    padding: 1.35rem 1.5rem 1.5rem;
    margin-bottom: 1.25rem;
    box-shadow: var(--shadow-panel);
}
.panel-label {
    font-size: 0.68rem;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: var(--gold);
    margin-bottom: 0.35rem;
}

/* ── Section headers ── */
.section-header {
    display: flex;
    align-items: flex-start;
    gap: 0.75rem;
    margin: 0 0 1rem 0;
}
.section-icon {
    width: 38px;
    height: 38px;
    flex-shrink: 0;
    background: linear-gradient(145deg, rgba(34, 211, 238, 0.1), rgba(212, 175, 55, 0.08));
    border: 1px solid var(--border-subtle);
    border-radius: var(--radius-sm);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.05rem;
}
.section-title {
    color: var(--text-primary);
    font-size: 1.08rem;
    font-weight: 600;
    margin: 0;
    letter-spacing: -0.015em;
}
.section-desc {
    color: var(--text-muted);
    font-size: 0.8rem;
    margin: 0.15rem 0 0 0;
    line-height: 1.45;
}

.info-card {
    background: var(--cyan-soft);
    border: 1px solid rgba(34, 211, 238, 0.18);
    border-left: 3px solid var(--cyan);
    border-radius: var(--radius-md);
    padding: 0.95rem 1.15rem;
    margin-bottom: 1rem;
    color: var(--text-secondary);
    font-size: 0.875rem;
    line-height: 1.65;
}
.info-card strong { color: var(--text-primary); }

/* ── Metrics grid ── */
.metrics-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 12px;
    margin-bottom: 1.25rem;
}
@media (max-width: 960px) { .metrics-grid { grid-template-columns: repeat(2, 1fr); } }
.metric-card {
    background: rgba(5, 8, 13, 0.5);
    border: 1px solid var(--border-faint);
    border-radius: var(--radius-md);
    padding: 1rem 1.15rem;
    position: relative;
    overflow: hidden;
    transition: border-color 0.2s, box-shadow 0.2s;
}
.metric-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, transparent, var(--gold), transparent);
    opacity: 0;
    transition: opacity 0.2s;
}
.metric-card:hover {
    border-color: var(--border-accent);
    box-shadow: 0 8px 24px rgba(0, 0, 0, 0.25);
}
.metric-card:hover::before { opacity: 0.6; }
.metric-label {
    display: block;
    color: var(--text-muted);
    font-size: 0.68rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.07em;
    margin-bottom: 0.4rem;
}
.metric-value {
    display: block;
    color: var(--text-primary);
    font-size: 1rem;
    font-weight: 600;
    line-height: 1.35;
    word-break: break-word;
}
.metric-value.accent {
    color: var(--gain);
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.15rem;
}

.panel-divider {
    border: none;
    border-top: 1px solid var(--border-faint);
    margin: 1.75rem 0;
}

/* ── Streamlit overrides ── */
.stButton > button[kind="primary"],
.stButton > button[data-testid="stBaseButton-primary"] {
    background: linear-gradient(135deg, #c9a227, #d4af37) !important;
    color: #0a0f16 !important;
    border: none !important;
    border-radius: var(--radius-sm) !important;
    font-weight: 700 !important;
    letter-spacing: 0.02em !important;
    padding: 0.55rem 1.35rem !important;
    box-shadow: 0 4px 18px rgba(212, 175, 55, 0.28) !important;
    transition: transform 0.15s, box-shadow 0.15s !important;
}
.stButton > button[kind="primary"]:hover,
.stButton > button[data-testid="stBaseButton-primary"]:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 24px rgba(212, 175, 55, 0.38) !important;
}
.stButton > button[kind="secondary"],
.stButton > button:not([kind="primary"]):not([data-testid="stBaseButton-primary"]) {
    background: rgba(26, 35, 50, 0.7) !important;
    border: 1px solid var(--border-subtle) !important;
    color: var(--text-secondary) !important;
    border-radius: var(--radius-sm) !important;
    font-weight: 500 !important;
}
.stDownloadButton > button {
    background: linear-gradient(135deg, #059669, #047857) !important;
    color: white !important;
    border: none !important;
    border-radius: var(--radius-sm) !important;
    font-weight: 600 !important;
    box-shadow: 0 4px 14px rgba(5, 150, 105, 0.3) !important;
}

.stTextInput input,
.stTextArea textarea,
.stDateInput input,
.stNumberInput input,
.stSelectbox [data-baseweb="select"] > div {
    background: rgba(5, 8, 13, 0.75) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: var(--radius-sm) !important;
    color: var(--text-primary) !important;
    font-family: 'Inter', sans-serif !important;
}
.stTextInput input:focus,
.stTextArea textarea:focus {
    border-color: var(--gold) !important;
    box-shadow: 0 0 0 3px rgba(212, 175, 55, 0.1) !important;
}
.stTextInput input::placeholder,
.stTextArea textarea::placeholder {
    color: var(--text-muted) !important;
    opacity: 0.7 !important;
}

.stRadio > div { gap: 6px; flex-wrap: wrap; }
.stRadio label {
    background: rgba(5, 8, 13, 0.6) !important;
    border: 1px solid var(--border-faint) !important;
    border-radius: var(--radius-sm) !important;
    padding: 0.55rem 1.15rem !important;
    color: var(--text-muted) !important;
    font-weight: 600 !important;
    font-size: 0.82rem !important;
    letter-spacing: 0.02em !important;
    transition: all 0.18s !important;
}
.stRadio label:has(input:checked) {
    background: var(--gold-soft) !important;
    border-color: rgba(212, 175, 55, 0.45) !important;
    color: var(--gold) !important;
    box-shadow: inset 0 -2px 0 var(--gold) !important;
}

.stCheckbox label { color: var(--text-secondary) !important; font-size: 0.875rem !important; }

.stMultiSelect [data-baseweb="tag"] {
    background: var(--gold-soft) !important;
    border: 1px solid rgba(212, 175, 55, 0.25) !important;
    color: var(--gold) !important;
}

div[data-testid="stMetric"] {
    background: rgba(5, 8, 13, 0.55);
    border: 1px solid var(--border-faint);
    border-radius: var(--radius-md);
    padding: 0.85rem 1rem;
}
div[data-testid="stMetric"] label {
    color: var(--text-muted) !important;
    font-size: 0.72rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.05em !important;
    text-transform: uppercase !important;
}
div[data-testid="stMetric"] [data-testid="stMetricValue"] {
    color: var(--text-primary) !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 1.05rem !important;
}
div[data-testid="stMetric"] [data-testid="stMetricDelta"] svg { display: none; }

.stDataFrame, [data-testid="stDataFrame"] {
    border: 1px solid var(--border-subtle) !important;
    border-radius: var(--radius-md) !important;
    overflow: hidden;
    background: rgba(5, 8, 13, 0.4) !important;
}

.stSuccess, .stWarning, .stError, .stInfo {
    border-radius: var(--radius-sm) !important;
    border-width: 1px !important;
    font-size: 0.875rem !important;
}
.stSuccess { background: var(--gain-soft) !important; border-color: rgba(16, 185, 129, 0.28) !important; color: #6ee7b7 !important; }
.stWarning { background: rgba(251, 191, 36, 0.08) !important; border-color: rgba(251, 191, 36, 0.25) !important; }
.stError { background: var(--loss-soft) !important; border-color: rgba(239, 68, 68, 0.28) !important; }
.stInfo { background: var(--cyan-soft) !important; border-color: rgba(34, 211, 238, 0.22) !important; }

label, .stMarkdown p { color: var(--text-secondary) !important; }
h1, h2, h3, h4, [data-testid="stMarkdownContainer"] h3 {
    color: var(--text-primary) !important;
    font-weight: 600 !important;
    letter-spacing: -0.01em !important;
}

details[data-testid="stExpander"] {
    background: var(--bg-panel) !important;
    border: 1px solid var(--border-subtle) !important;
    border-radius: var(--radius-md) !important;
}
details[data-testid="stExpander"] summary {
    color: var(--text-primary) !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
}

div[data-testid="stVerticalBlockBorderWrapper"] {
    background: var(--bg-panel) !important;
    backdrop-filter: blur(12px);
    border: 1px solid var(--border-subtle) !important;
    border-radius: var(--radius-xl) !important;
    padding: 0.35rem 0.5rem 0.75rem !important;
    box-shadow: var(--shadow-panel) !important;
    margin-bottom: 1.25rem !important;
}
div[data-testid="stVerticalBlockBorderWrapper"] .panel-label {
    font-size: 0.68rem;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: var(--gold);
    margin: 0.5rem 0.75rem 0.25rem;
}

.stSpinner > div { color: var(--gold) !important; }

.app-footer {
    text-align: center;
    color: var(--text-muted);
    font-size: 0.75rem;
    margin-top: 2.5rem;
    padding-top: 1.5rem;
    border-top: 1px solid var(--border-faint);
    letter-spacing: 0.02em;
}
.app-footer strong { color: var(--text-secondary); font-weight: 500; }

/* Chart styling hint */
[data-testid="stLineChart"] {
    border: 1px solid var(--border-faint);
    border-radius: var(--radius-md);
    padding: 0.5rem;
    background: rgba(5, 8, 13, 0.35);
}
</style>
""",
        unsafe_allow_html=True,
    )


def render_top_bar() -> None:
    st.markdown(
        """
        <div class="top-bar">
            <div class="top-bar-brand">
                <div class="brand-mark">📊</div>
                <div>
                    <div class="brand-name">NAV Terminal</div>
                    <div class="brand-tag">AMFI India · Mutual Fund Intelligence</div>
                </div>
            </div>
            <div class="market-status">
                <span class="pulse-dot"></span>
                Live Feed Active
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_ticker_strip() -> None:
    items = [
        ('NIFTY 50', '24,835', '+0.42%', 'up'),
        ('SENSEX', '81,482', '+0.38%', 'up'),
        ('NIFTY MIDCAP', '56,210', '-0.12%', 'down'),
        ('GOLD', '₹72,450', '+0.18%', 'up'),
        ('USD/INR', '83.42', '-0.05%', 'down'),
        ('AMFI NAV', 'Daily', 'Updated', 'up'),
    ]
    ticker_html = ""
    for symbol, price, change, direction in items * 2:
        ticker_html += (
            f'<span class="ticker-item">'
            f'<span class="ticker-symbol">{symbol}</span>'
            f'<span>{price}</span>'
            f'<span class="ticker-change {direction}">{change}</span>'
            f'<span class="ticker-sep">|</span>'
            f'</span>'
        )
    st.markdown(
        f'<div class="ticker-strip"><div class="ticker-track">{ticker_html}</div></div>',
        unsafe_allow_html=True,
    )


def render_hero(
    title: str,
    subtitle: str,
    badge: str = "AMFI Live Data",
    chips: list[dict[str, str]] | list[str] | None = None,
) -> None:
    chips_html = ""
    if chips:
        chip_items = []
        for chip in chips:
            if isinstance(chip, dict):
                label = chip.get("label", "")
                value = chip.get("value", "")
                tone = chip.get("tone", "")
                tone_class = f" {tone}" if tone else ""
                chip_items.append(
                    f'<div class="stat-chip">'
                    f'<span class="stat-chip-label">{label}</span>'
                    f'<span class="stat-chip-value{tone_class}">{value}</span>'
                    f'</div>'
                )
            else:
                chip_items.append(f'<div class="stat-chip"><span class="stat-chip-value">{chip}</span></div>')
        chips_html = f'<div class="stat-row">{"".join(chip_items)}</div>'

    title_html = title.replace("NAV", "<span>NAV</span>", 1) if "NAV" in title else title

    st.markdown(
        f"""
        <div class="hero-card">
            <span class="hero-badge">◆ {badge}</span>
            <h1 class="hero-title">{title_html}</h1>
            <p class="hero-sub">{subtitle}</p>
            {chips_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_panel_label(label: str) -> None:
    st.markdown(f'<div class="panel-label">{label}</div>', unsafe_allow_html=True)


@contextmanager
def finance_panel(label: str):
    with st.container(border=True):
        render_panel_label(label)
        yield


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
        """
        <div class="app-footer">
            <strong>Data sourced from AMFI India</strong> · Built for research, portfolio analysis &amp; NAV reporting
        </div>
        """,
        unsafe_allow_html=True,
    )
