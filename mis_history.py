"""
Saved MIS reports.

A model portfolio changes rarely -- often only once a year -- so a generated
report is worth keeping rather than regenerating from a feed that may since
have moved. Each save records the workbook itself plus the inputs that produced
it, so an old report can be read back exactly as issued and, if wanted, rebuilt.

Storage is a directory of JSON sidecars alongside their .xlsx files. No
database: the volume is a handful of reports a year, and a plain directory
survives inspection, backup and hand-editing.
"""

from __future__ import annotations

import json
import re
import shutil
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

HISTORY_DIR = Path(__file__).resolve().parent / "mis_history"

# Deployed hosts give the app an ephemeral filesystem: anything written here is
# lost when the container restarts. The UI says so rather than implying reports
# are archived forever.
EPHEMERAL_NOTE = (
    "Saved reports live on the server's disk. On a hosted deployment that disk "
    "is wiped when the app restarts, so download anything you need to keep."
)

_ID_SAFE = re.compile(r"[^0-9A-Za-z_-]+")


def _ensure_dir() -> Path:
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    return HISTORY_DIR


def _portfolio_records(df: Optional[pd.DataFrame]) -> List[Dict[str, Any]]:
    """Reduce a portfolio to the fields needed to redisplay or rebuild it."""
    if df is None or df.empty:
        return []
    cols = [c for c in ("Scheme Name", "ISIN", "Allocation (%)", "Benchmark") if c in df.columns]
    out: List[Dict[str, Any]] = []
    for _, r in df[cols].iterrows():
        rec = {}
        for c in cols:
            v = r[c]
            rec[c] = float(v) if isinstance(v, (int, float)) and c == "Allocation (%)" else str(v)
        out.append(rec)
    return out


def _benchmark_rows(report: Any) -> List[Dict[str, str]]:
    """Normalise the benchmark-source table to JSON-safe rows.

    Callers hand this over as either a DataFrame or a list of dicts depending
    on where it was built, and a DataFrame has no usable truth value.
    """
    if report is None:
        return []
    if isinstance(report, pd.DataFrame):
        if report.empty:
            return []
        records = report.head(40).to_dict("records")
    elif isinstance(report, list):
        records = [r for r in report[:40] if isinstance(r, dict)]
    else:
        return []
    return [{str(k): str(v) for k, v in rec.items()} for rec in records]


def _as_iso(value: Any) -> Optional[str]:
    if isinstance(value, (date, datetime)):
        return value.isoformat()[:10]
    return str(value) if value is not None else None


def save_report(
    mis_data: Dict[str, Any],
    excel_bytes: bytes,
    portfolio_df: pd.DataFrame,
    previous_df: Optional[pd.DataFrame],
    label: str = "",
) -> str:
    """Persist one generated report. Returns its id."""
    _ensure_dir()
    dates = mis_data.get("dates", {}) or {}
    now = datetime.now()

    start = _as_iso(dates.get("start_date"))
    end = _as_iso(dates.get("end_date"))
    entry_id = _ID_SAFE.sub("-", f"{now:%Y%m%d-%H%M%S}_{start}_to_{end}")

    xlsx_path = HISTORY_DIR / f"{entry_id}.xlsx"
    xlsx_path.write_bytes(excel_bytes)

    meta = {
        "id": entry_id,
        "label": label or f"MIS {start} to {end}",
        "saved_at": now.isoformat(timespec="seconds"),
        "period_start": start,
        "period_end": end,
        "fy_start": _as_iso(dates.get("fy_start")),
        "mis3_start": _as_iso(dates.get("mis3_start")),
        "mis3_end": _as_iso(dates.get("mis3_end")),
        "include_flows": bool(mis_data.get("include_flows")),
        "scheme_count": len(portfolio_df) if portfolio_df is not None else 0,
        "portfolio": _portfolio_records(portfolio_df),
        "previous_portfolio": _portfolio_records(previous_df),
        "warnings": [str(w) for w in (mis_data.get("warnings") or [])][:20],
        "benchmark_report": _benchmark_rows(mis_data.get("benchmark_report")),
        "xlsx": xlsx_path.name,
        "xlsx_bytes": len(excel_bytes),
    }
    (HISTORY_DIR / f"{entry_id}.json").write_text(
        json.dumps(meta, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    return entry_id


def list_reports() -> List[Dict[str, Any]]:
    """Every saved report, newest first. Unreadable sidecars are skipped."""
    if not HISTORY_DIR.exists():
        return []
    out: List[Dict[str, Any]] = []
    for path in HISTORY_DIR.glob("*.json"):
        try:
            out.append(json.loads(path.read_text(encoding="utf-8")))
        except (OSError, ValueError):
            continue  # a corrupt sidecar must not hide the rest of the history
    out.sort(key=lambda m: m.get("saved_at", ""), reverse=True)
    return out


def load_report(entry_id: str) -> Tuple[Optional[Dict[str, Any]], Optional[bytes]]:
    """Return (metadata, workbook bytes) for one saved report."""
    meta_path = HISTORY_DIR / f"{entry_id}.json"
    if not meta_path.exists():
        return None, None
    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None, None
    xlsx_path = HISTORY_DIR / str(meta.get("xlsx", ""))
    data = xlsx_path.read_bytes() if xlsx_path.exists() else None
    return meta, data


def delete_report(entry_id: str) -> bool:
    """Remove a saved report and its workbook."""
    ok = False
    for suffix in (".json", ".xlsx"):
        p = HISTORY_DIR / f"{entry_id}{suffix}"
        if p.exists():
            try:
                p.unlink()
                ok = True
            except OSError:
                pass
    return ok


def portfolio_frame(records: List[Dict[str, Any]]) -> pd.DataFrame:
    """Rebuild an editable portfolio frame from a saved report."""
    cols = ["Scheme Name", "ISIN", "Allocation (%)", "Benchmark"]
    if not records:
        return pd.DataFrame(columns=cols)
    df = pd.DataFrame(records)
    for c in cols:
        if c not in df.columns:
            df[c] = "" if c != "Allocation (%)" else 0.0
    df["Allocation (%)"] = pd.to_numeric(df["Allocation (%)"], errors="coerce").fillna(0.0)
    return df[cols]


def history_size() -> Tuple[int, float]:
    """(report count, total megabytes on disk)."""
    if not HISTORY_DIR.exists():
        return 0, 0.0
    total = sum(p.stat().st_size for p in HISTORY_DIR.glob("*") if p.is_file())
    return len(list(HISTORY_DIR.glob("*.json"))), total / (1024 * 1024)


def clear_history() -> None:
    """Delete every saved report."""
    if HISTORY_DIR.exists():
        shutil.rmtree(HISTORY_DIR, ignore_errors=True)
