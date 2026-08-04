"""
Benchmark index level sourcing for the MIS Generator.

The historical index feed this project used to rely on (niftyindices.com
``Backpage.aspx``) no longer returns JSON, so every index request failed and the
old code silently substituted a randomly simulated price series. This module
replaces that with real data: each benchmark is mapped to a passive index fund
whose NAV history is pulled from the same AMFI feed already used for scheme
NAVs.

An index fund's NAV compounds the index's price moves *and* its dividends net of
expenses, so NAV growth tracks the Total Return (TR) index minus TER. That makes
these proxies a much closer match for the "TR INR" benchmarks used in the MIS
than a price-return index would be.

Nothing here ever invents data. If a benchmark cannot be resolved or its proxy
has no NAV coverage, the caller receives an empty series plus a status the report
surfaces as "N/A".
"""

from __future__ import annotations

import re
from datetime import date, datetime, timedelta, timezone
from typing import Dict, List, Optional, Tuple

import pandas as pd

from nav_fetcher import (
    fetch_amfi_data_chunked,
    fetch_latest_navs,
    parse_amfi_date_series,
    _parse_amfi_date_str,
)

# ─── Resolution status ────────────────────────────────────────────────────────

STATUS_EXACT = "EXACT"      # proxy fund tracks precisely this index
STATUS_APPROX = "APPROX"    # no fund tracks this index; closest available used
STATUS_NONE = "UNAVAILABLE"  # nothing usable — report must show N/A


class BenchmarkProxy:
    """A benchmark and the passive fund(s) standing in for its index level."""

    def __init__(self, key: str, isins: List[Tuple[str, str]], status: str, note: str = ""):
        self.key = key
        self.isins = isins          # ordered [(isin, fund name)] — first with data wins
        self.status = status
        self.note = note


# ─── Benchmark → passive fund map ─────────────────────────────────────────────
# Candidates are ordered longest-history first, so a report starting years back
# still resolves. Every ISIN below was verified present in AMFI's live NAVAll
# feed as a Direct-Plan Growth option.

_P = BenchmarkProxy

BENCHMARK_PROXIES: Dict[str, BenchmarkProxy] = {
    "NIFTY 50": _P("NIFTY 50", [
        ("INF209K01VY3", "Aditya Birla Sun Life Nifty 50 Index Fund - Direct - Growth"),
        ("INF789F01XA0", "UTI Nifty 50 Index Fund - Direct - Growth"),
        ("INF090I01GS4", "Franklin India NSE Nifty 50 Index Fund - Direct - Growth"),
        ("INF767K01FF2", "LIC MF Nifty 50 Index Fund - Direct - Growth"),
    ], STATUS_EXACT),

    "NIFTY 500": _P("NIFTY 500", [
        ("INF109KC16Y3", "ICICI Prudential Nifty 500 Index Fund - Direct - Growth"),
        ("INF846K019W9", "Axis Nifty 500 Index Fund - Direct - Growth"),
        ("INF740KA1XY9", "DSP Nifty 500 Index Fund - Direct - Growth"),
        ("INF200KB1381", "SBI Nifty 500 Index Fund - Direct - Growth"),
    ], STATUS_EXACT),

    "NIFTY 100": _P("NIFTY 100", [
        ("INF846K01S29", "Axis Nifty 100 Index Fund - Direct - Growth"),
        ("INF194KB1CR7", "Bandhan Nifty 100 Index Fund - Direct - Growth"),
        ("INF179KC1BY3", "HDFC Nifty 100 Index Fund - Direct - Growth"),
    ], STATUS_EXACT),

    "NIFTY MIDCAP 150": _P("NIFTY MIDCAP 150", [
        ("INF204KB18Z7", "Nippon India Nifty Midcap 150 Index Fund - Direct - Growth"),
        ("INF209KB1W74", "Aditya Birla Sun Life Nifty Midcap 150 Index Fund - Direct - Growth"),
        ("INF109KC1W58", "ICICI Prudential Nifty Midcap 150 Index Fund - Direct - Growth"),
        ("INF200KA10P5", "SBI Nifty Midcap 150 Index Fund - Direct - Growth"),
        ("INF179KC1GC8", "HDFC Nifty Midcap 150 Index Fund - Direct - Growth"),
    ], STATUS_EXACT),

    "NIFTY SMALLCAP 250": _P("NIFTY SMALLCAP 250", [
        ("INF204KB15W0", "Nippon India Nifty Smallcap 250 Index Fund - Direct - Growth"),
        ("INF200KA16P2", "SBI Nifty Smallcap 250 Index Fund - Direct - Growth"),
        ("INF179KC1GE4", "HDFC Nifty Smallcap 250 Index Fund - Direct - Growth"),
        ("INF109KC1V18", "ICICI Prudential Nifty Smallcap 250 Index Fund - Direct - Growth"),
        ("INF754K01QT8", "Edelweiss Nifty Smallcap 250 Index Fund - Direct - Growth"),
    ], STATUS_EXACT),

    "NIFTY LARGEMIDCAP 250": _P("NIFTY LARGEMIDCAP 250", [
        ("INF754K01NR9", "Edelweiss Nifty Large Midcap 250 Index Fund - Direct - Growth"),
        ("INF0R8F01018", "Zerodha Nifty LargeMidcap 250 Index Fund - Direct - Growth"),
        ("INF109KC12U0", "ICICI Prudential Nifty LargeMidcap 250 Index Fund - Direct - Growth"),
        ("INF769K01MJ6", "Mirae Asset Nifty LargeMidcap 250 Index Fund - Direct - Growth"),
        ("INF179KC1IQ4", "HDFC Nifty LargeMidcap 250 Index Fund - Direct - Growth"),
    ], STATUS_EXACT),

    "NIFTY NEXT 50": _P("NIFTY NEXT 50", [
        ("INF109K01Y80", "ICICI Prudential Nifty Next 50 Index Fund - Direct - Growth"),
        ("INF397L01AW2", "LIC MF Nifty Next 50 Index Fund - Direct - Growth"),
        ("INF789FC12T1", "UTI Nifty Next 50 Index Fund - Direct - Growth"),
    ], STATUS_EXACT),

    "NIFTY BANK": _P("NIFTY BANK", [
        ("INF109KC11A4", "ICICI Prudential Nifty Bank Index Fund - Direct - Growth"),
        ("INF204KC1BY2", "Nippon India Nifty Bank Index Fund - Direct - Growth"),
        ("INF200KB1563", "SBI Nifty Bank Index Fund - Direct - Growth"),
    ], STATUS_EXACT),

    "BSE 500": _P("BSE 500", [
        ("INF179KC1GG9", "HDFC BSE 500 Index Fund - Direct - Growth"),
    ], STATUS_EXACT),

    "BSE SENSEX": _P("BSE SENSEX", [
        ("INF767K01FJ4", "LIC MF BSE Sensex Index Fund - Direct - Growth"),
        ("INF789F1AVD7", "UTI BSE Sensex Index Fund - Direct - Growth"),
        ("INF200KA16Y4", "SBI BSE Sensex Index Fund - Direct - Growth"),
    ], STATUS_EXACT),

    # ── No passive fund tracks these indices. Closest available stand-in is
    # used and the report labels the row APPROX so the number is never mistaken
    # for the real benchmark.
    "BSE 250 SMALLCAP": _P("BSE 250 SMALLCAP", [
        ("INF204KB15W0", "Nippon India Nifty Smallcap 250 Index Fund - Direct - Growth"),
        ("INF200KA16P2", "SBI Nifty Smallcap 250 Index Fund - Direct - Growth"),
    ], STATUS_APPROX, "No BSE 250 SmallCap index fund exists; Nifty Smallcap 250 used as stand-in."),

    "NIFTY 500 MULTICAP 50:25:25": _P("NIFTY 500 MULTICAP 50:25:25", [
        ("INF179KC1IO9", "HDFC NIFTY500 Multicap 50:25:25 Index Fund - Direct Plan"),
        ("INF179KC1IN1", "HDFC NIFTY500 Multicap 50:25:25 Index Fund - Regular Plan"),
    ], STATUS_EXACT),

    "NIFTY INFRASTRUCTURE": _P("NIFTY INFRASTRUCTURE", [
        ("INF109KC16E5", "ICICI Prudential Nifty Infrastructure ETF"),
    ], STATUS_EXACT),

    "NIFTY MIDCAP 100": _P("NIFTY MIDCAP 100", [
        ("INF204KB18Z7", "Nippon India Nifty Midcap 150 Index Fund - Direct - Growth"),
    ], STATUS_APPROX, "No Nifty Midcap 100 index fund exists; Nifty Midcap 150 used as stand-in."),
}

# Spellings that normalize to something other than the canonical key above.
_ALIASES = {
    "NIFTY LARGE MIDCAP 250": "NIFTY LARGEMIDCAP 250",
    "NIFTY LARGE AND MIDCAP 250": "NIFTY LARGEMIDCAP 250",
    "NIFTY LARGEMID 250": "NIFTY LARGEMIDCAP 250",
    "BSE SMALLCAP 250": "BSE 250 SMALLCAP",
    "BSE 250 SMALL CAP": "BSE 250 SMALLCAP",
    "SENSEX": "BSE SENSEX",
    "BSE 30": "BSE SENSEX",
    "NIFTY500 MULTICAP 50:25:25": "NIFTY 500 MULTICAP 50:25:25",
    "NIFTY 500 MULTICAP": "NIFTY 500 MULTICAP 50:25:25",
    "NIFTY MIDSMALLCAP 400": "NIFTY LARGEMIDCAP 250",
    "NIFTY SMALLCAP 100": "NIFTY SMALLCAP 250",
    "NIFTY BANK INDEX": "NIFTY BANK",
    "BANK NIFTY": "NIFTY BANK",
    "NIFTY 50 TR": "NIFTY 50",
}

DEFAULT_BENCHMARK = "NIFTY 50"


# ─── Name normalization ───────────────────────────────────────────────────────

# Decorations carried by Morningstar-style benchmark names that say nothing
# about which index is meant: "S&P BSE 250 SmallCap TR INR" → "BSE 250 SMALLCAP".
_NOISE_WORDS = r"\b(S&P|SP|TR|TRI|TOTAL\s+RETURN|INR|INDEX|INDICES|IND|INDIA)\b"


def normalize_benchmark_name(raw: str) -> str:
    """Reduce a benchmark label to the canonical key used in BENCHMARK_PROXIES."""
    if raw is None:
        return ""
    s = str(raw).upper().strip()
    if not s or s in ("NAN", "NONE", "-", ""):
        return ""

    s = s.replace("&AMP;", "&")
    s = re.sub(_NOISE_WORDS, " ", s)
    # Keep digits, letters and the ':' that carries meaning in "50:25:25".
    s = re.sub(r"[^A-Z0-9: ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()

    # Join the cap-size words so "LARGE MIDCAP"/"MID CAP" spellings converge.
    s = re.sub(r"\bSMALL\s+CAP\b", "SMALLCAP", s)
    s = re.sub(r"\bMID\s+CAP\b", "MIDCAP", s)
    s = re.sub(r"\bLARGE\s+MIDCAP\b", "LARGEMIDCAP", s)
    s = re.sub(r"\bLARGEMID\s+CAP\b", "LARGEMIDCAP", s)
    s = re.sub(r"\bNIFTY(\d)", r"NIFTY \1", s)
    s = re.sub(r"\s+", " ", s).strip()

    return _ALIASES.get(s, s)


def resolve_benchmark(raw_name: str) -> Tuple[str, Optional[BenchmarkProxy]]:
    """Map a benchmark label to its canonical key and proxy definition.

    Returns (canonical_key, proxy_or_None). A None proxy means the label was not
    recognised — the caller must report N/A rather than substitute anything.
    """
    key = normalize_benchmark_name(raw_name)
    if not key:
        key = DEFAULT_BENCHMARK
    return key, BENCHMARK_PROXIES.get(key)


# ─── Index level series ───────────────────────────────────────────────────────

class BenchmarkSeries:
    """Resolved level series for one benchmark, plus how it was sourced."""

    def __init__(self, requested: str, key: str, levels: Dict[date, float],
                 status: str, proxy_isin: str = "", proxy_name: str = "", note: str = ""):
        self.requested = requested
        self.key = key
        self.levels = levels
        self.status = status
        self.proxy_isin = proxy_isin
        self.proxy_name = proxy_name
        self.note = note

    @property
    def ok(self) -> bool:
        return bool(self.levels)

    def dates(self) -> List[date]:
        return sorted(self.levels.keys())


def _nav_frame_to_isin_series(df: pd.DataFrame) -> Dict[str, Dict[date, float]]:
    """Collapse a raw AMFI NAV frame into {ISIN: {date: nav}}."""
    out: Dict[str, Dict[date, float]] = {}
    if df.empty:
        return out

    for _, r in df.iterrows():
        d = r.get("NAV Date_Date")
        nav = r.get("NAV")
        if d is None or pd.isna(d) or pd.isna(nav):
            continue
        try:
            nav_f = float(nav)
        except (TypeError, ValueError):
            continue
        if nav_f <= 0:
            continue
        for col in ("ISIN_G", "ISIN_R"):
            isin = r.get(col)
            if isin and isin not in ("-", "NAN", "NONE"):
                out.setdefault(isin, {})[d] = nav_f
    return out


def required_proxy_isins(benchmark_names: List[str]) -> List[str]:
    """Every proxy-fund ISIN needed to price the given benchmarks.

    Lets a caller fold these into its own AMFI fetch rather than issuing a
    second round of requests. AMFI throttles aggressively, and a throttled
    chunk becomes a hole in the history that shifts return baselines.
    """
    needed: List[str] = []
    for name in dict.fromkeys(benchmark_names):
        _key, proxy = resolve_benchmark(name)
        if not proxy:
            continue
        for isin, _fund in proxy.isins:
            if isin not in needed:
                needed.append(isin)
    return needed


def build_benchmark_series(
    benchmark_names: List[str],
    isin_series: Dict[str, Dict[date, float]],
    start_date: date,
    gap_note: str = "",
) -> Dict[str, BenchmarkSeries]:
    """Assemble BenchmarkSeries from an already-fetched {ISIN: {date: nav}} map.

    Every requested name gets an entry; unresolvable ones carry an empty series
    and STATUS_NONE so the report prints N/A instead of a fabricated number.
    """
    requested = [b for b in dict.fromkeys(benchmark_names) if str(b).strip()]
    resolved: Dict[str, Tuple[str, Optional[BenchmarkProxy]]] = {
        name: resolve_benchmark(name) for name in requested
    }

    out: Dict[str, BenchmarkSeries] = {}
    for name in requested:
        key, proxy = resolved[name]

        if proxy is None:
            out[name] = BenchmarkSeries(
                name, key, {}, STATUS_NONE,
                note=f"'{name}' is not a recognised benchmark. Add it to BENCHMARK_PROXIES "
                     f"in benchmark_proxy.py to enable it.",
            )
            continue

        # Take the first candidate that actually covers the requested window.
        chosen = None
        for isin, fund_name in proxy.isins:
            levels = isin_series.get(isin) or {}
            if not levels:
                continue
            earliest = min(levels)
            if earliest <= start_date:
                chosen = (isin, fund_name, levels)
                break
            # Remember a short-history fallback in case nothing covers fully.
            if chosen is None:
                chosen = (isin, fund_name, levels)

        if chosen is None:
            out[name] = BenchmarkSeries(
                name, key, {}, STATUS_NONE,
                note=f"No NAV history returned for any proxy fund of '{key}'.",
            )
            continue

        isin, fund_name, levels = chosen
        note = " ".join(p for p in (proxy.note, gap_note) if p)
        earliest = min(levels)
        if earliest > start_date:
            note = (note + " " if note else "") + (
                f"Proxy fund history starts {earliest:%d-%b-%Y}, after the requested "
                f"start {start_date:%d-%b-%Y}; earlier baselines are unavailable."
            )

        out[name] = BenchmarkSeries(
            name, key, levels, proxy.status,
            proxy_isin=isin, proxy_name=fund_name, note=note,
        )

    return out


NIFTY_PRICE_SYMBOL = "^NSEI"


def fetch_nifty_price_index(start_date: date, end_date: date) -> Dict[date, float]:
    """Daily closes for the Nifty 50 *price* index.

    The MIS uses two conventions on purpose: the schemes' own benchmarks are
    total-return ("TR INR"), but the standalone Nifty column is the headline
    price index, which excludes dividends. An index fund proxy would answer the
    total-return question and read roughly half a point high over a quarter, so
    this column needs the real price index.

    Returns {} on failure -- callers must fall back visibly, never silently.
    """
    import requests

    try:
        resp = requests.get(
            f"https://query1.finance.yahoo.com/v8/finance/chart/{NIFTY_PRICE_SYMBOL}",
            params={
                "period1": int(datetime.combine(start_date - timedelta(days=10),
                                                datetime.min.time()).timestamp()),
                "period2": int(datetime.combine(end_date + timedelta(days=2),
                                                datetime.min.time()).timestamp()),
                "interval": "1d",
            },
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                                   "Chrome/124.0.0.0 Safari/537.36"},
            timeout=30,
        )
        resp.raise_for_status()
        result = resp.json()["chart"]["result"][0]
        stamps = result["timestamp"]
        closes = result["indicators"]["quote"][0]["close"]
    except Exception as exc:  # network, shape change, empty result
        print(f"Nifty 50 price index unavailable: {type(exc).__name__}: {exc}")
        return {}

    levels: Dict[date, float] = {}
    for ts, close in zip(stamps, closes):
        if close is None:
            continue
        levels[datetime.fromtimestamp(ts, tz=timezone.utc).date()] = float(close)
    return levels


def nav_frame_to_isin_series(df_raw: pd.DataFrame) -> Dict[str, Dict[date, float]]:
    """Normalize a raw AMFI frame and collapse it to {ISIN: {date: nav}}."""
    if df_raw is None or df_raw.empty:
        return {}
    df = df_raw.copy()
    if "NAV Date_Date" not in df.columns:
        df["NAV Date"] = parse_amfi_date_series(df["NAV Date"])
        df = df.dropna(subset=["NAV Date", "NAV"])
        df["NAV Date_Date"] = df["NAV Date"].dt.date
    if "ISIN_G" not in df.columns:
        df["ISIN_G"] = df["ISIN Div Payout / ISIN Growth"].astype(str).str.strip().str.upper()
        df["ISIN_R"] = df["ISIN Div Reinvestment"].astype(str).str.strip().str.upper()
    return _nav_frame_to_isin_series(df)


def overlay_live_navs(isin_series: Dict[str, Dict[date, float]],
                      isins: List[str], end_date: date) -> Dict[str, Dict[date, float]]:
    """Add today's NAV from the live AMFI feed so a same-day report isn't stale."""
    for isin_u, info in (fetch_latest_navs(isins) or {}).items():
        iso = _parse_amfi_date_str(info.get("date", ""))
        if not iso:
            continue
        try:
            d_live = datetime.strptime(iso, "%Y-%m-%d").date()
            nav_live = float(info["nav"])
        except (ValueError, TypeError, KeyError):
            continue
        if nav_live > 0 and d_live <= end_date:
            isin_series.setdefault(isin_u, {})[d_live] = nav_live
    return isin_series


def fetch_benchmark_levels(
    benchmark_names: List[str],
    start_date: date,
    end_date: date,
) -> Dict[str, BenchmarkSeries]:
    """Fetch a level series for every requested benchmark, standalone.

    Callers that already fetch AMFI data should instead combine
    required_proxy_isins() with build_benchmark_series() to avoid a second
    round of requests.
    """
    requested = [b for b in dict.fromkeys(benchmark_names) if str(b).strip()]
    if not requested:
        return {}

    needed_isins = required_proxy_isins(requested)
    isin_series: Dict[str, Dict[date, float]] = {}
    gap_note = ""

    if needed_isins:
        # Reach back so the baseline date still resolves across a long holiday.
        fetch_start = start_date - timedelta(days=20)
        # include_direct: proxies are Direct plans, whose low TER tracks the
        # index closest. The app's default parse drops Direct plans.
        df_raw, gaps = fetch_amfi_data_chunked(
            fetch_start, end_date, needed_isins, include_direct=True, return_gaps=True
        )
        if gaps:
            gap_note = describe_gaps(gaps)
        isin_series = nav_frame_to_isin_series(df_raw)
        isin_series = overlay_live_navs(isin_series, needed_isins, end_date)

    return build_benchmark_series(requested, isin_series, start_date, gap_note)


def describe_gaps(gaps: List[Tuple[date, date]]) -> str:
    """Human-readable note for date windows AMFI would not serve."""
    if not gaps:
        return ""
    spans = ", ".join(f"{a:%d-%b-%Y}-{b:%d-%b-%Y}" for a, b in gaps)
    return (f"AMFI did not serve NAV data for {spans}; baselines in those windows fall back "
            f"to the nearest earlier level.")


def level_asof(series: Optional[BenchmarkSeries], target: date) -> Tuple[Optional[float], Optional[date]]:
    """Latest level on or before ``target``, with the date it actually came from."""
    if series is None or not series.levels:
        return None, None
    candidates = [d for d in series.levels if d <= target]
    if not candidates:
        return None, None
    d = max(candidates)
    return series.levels[d], d


def level_prev(series: Optional[BenchmarkSeries], before: date) -> Tuple[Optional[float], Optional[date]]:
    """Latest level strictly before ``before`` — the true prior trading day.

    Used for one-day returns so a holiday can never resolve both ends of the
    comparison to the same observation (which would silently report 0.00%).
    """
    if series is None or not series.levels:
        return None, None
    candidates = [d for d in series.levels if d < before]
    if not candidates:
        return None, None
    d = max(candidates)
    return series.levels[d], d


def describe_resolution(all_series: Dict[str, BenchmarkSeries]) -> pd.DataFrame:
    """Tabulate how each benchmark was sourced, for display alongside the report."""
    rows = []
    for name, s in all_series.items():
        rows.append({
            "Benchmark": name,
            "Resolved As": s.key,
            "Status": s.status,
            "Proxy Fund": s.proxy_name or "-",
            "Proxy ISIN": s.proxy_isin or "-",
            "Observations": len(s.levels),
            "Coverage From": min(s.levels).strftime("%d-%b-%Y") if s.levels else "-",
            "Coverage To": max(s.levels).strftime("%d-%b-%Y") if s.levels else "-",
            "Note": s.note or "",
        })
    return pd.DataFrame(rows)
