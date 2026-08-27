"""
Resumable AUM backfill against AMFI's fund-performance API.

AMFI rate-limits by IP and answers 403 or 503 once it decides a caller is being
greedy. A report that fetches on demand therefore fails at exactly the moment
someone needs it. This collects the data slowly and permanently instead: it
fetches until AMFI cuts it off, waits for the block to lift, resumes, and writes
every answered query to disk.

Past-date AUM never changes, so the cache is write-once. After a first pass the
only work left is the newest day.

Progress is the cache itself -- there is no separate state file to fall out of
step. Stopping the job and starting it again picks up exactly where it left off.
"""

from __future__ import annotations

import json
import os
import random
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Tuple

import requests

CACHE_DIR = Path(__file__).resolve().parent / "api_cache"

_URL = "https://www.amfiindia.com/gateway/pollingsebi/api/amfi/fundperformance"
_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"),
    "Content-Type": "application/json",
    "Accept": "application/json, text/plain, */*",
    "Origin": "https://www.amfiindia.com",
    "Referer": "https://www.amfiindia.com/research-information/other-data/fund-performance",
    "X-Requested-With": "XMLHttpRequest",
}

# AMFI's own categorisation. Equity is the default because that is what the MIS
# and the fund-performance screens are built around.
EQUITY_SUBCATEGORIES = list(range(1, 13))
CATEGORY_EQUITY = 1
MATURITY_OPEN_ENDED = 1

# The performance dataset starts around SEBI recategorisation; earlier dates
# return nothing and are not worth a request.
EARLIEST_USEFUL = date(2018, 1, 1)


class Blocked(Exception):
    """AMFI refused the request. Transient -- the caller should wait, not skip."""


def cache_key(day: date, maturity: int, category: int, subcategory: int) -> str:
    return f"{day:%d-%b-%Y}_{maturity}_{category}_{subcategory}"


def cache_path(day: date, maturity: int, category: int, subcategory: int) -> Path:
    return CACHE_DIR / f"{cache_key(day, maturity, category, subcategory)}.json"


def is_cached(day: date, maturity: int, category: int, subcategory: int) -> bool:
    return cache_path(day, maturity, category, subcategory).exists()


def _fetch_one(day: date, maturity: int, category: int, subcategory: int,
               timeout: int = 30) -> List[dict]:
    """One query. Returns rows, or raises Blocked if AMFI refused.

    An empty list from a successful response is a real answer -- AMFI has no
    data for that date and category -- and is cached so it is never asked again.
    """
    payload = {
        "maturityType": maturity,
        "category": category,
        "subCategory": subcategory,
        "mfid": 0,
        "reportDate": f"{day:%d-%b-%Y}",
    }
    try:
        resp = requests.post(_URL, json=payload, headers=_HEADERS, timeout=timeout)
    except requests.exceptions.RequestException as exc:
        raise Blocked(f"network error: {type(exc).__name__}") from exc

    if resp.status_code in (403, 429, 503, 502, 504):
        raise Blocked(f"HTTP {resp.status_code}")
    if resp.status_code != 200:
        raise Blocked(f"HTTP {resp.status_code}")

    try:
        body = resp.json()
    except ValueError:
        # A 200 carrying an HTML error page is a block wearing a disguise.
        raise Blocked("non-JSON response")

    if body.get("validationMsg") != "SUCCESS":
        raise Blocked(f"validationMsg={body.get('validationMsg')}")

    return body.get("data") or []


def _write_cache(day: date, maturity: int, category: int, subcategory: int,
                 rows: List[dict]) -> None:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    path = cache_path(day, maturity, category, subcategory)
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(rows), encoding="utf-8")
    tmp.replace(path)


def plan_work(start: date, end: date,
              subcategories: Iterable[int] = EQUITY_SUBCATEGORIES,
              category: int = CATEGORY_EQUITY,
              maturity: int = MATURITY_OPEN_ENDED,
              skip_weekends: bool = True) -> List[Tuple[date, int, int, int]]:
    """Every query needed for the range that is not already cached."""
    work: List[Tuple[date, int, int, int]] = []
    day = max(start, EARLIEST_USEFUL)
    while day <= end:
        if not (skip_weekends and day.weekday() >= 5):
            for sub in subcategories:
                if not is_cached(day, maturity, category, sub):
                    work.append((day, maturity, category, sub))
        day += timedelta(days=1)
    return work


def backfill(
    start: date,
    end: date,
    subcategories: Iterable[int] = EQUITY_SUBCATEGORIES,
    category: int = CATEGORY_EQUITY,
    maturity: int = MATURITY_OPEN_ENDED,
    polite_delay: float = 1.5,
    max_wait: float = 1800.0,
    total_budget: Optional[float] = None,
    progress: Optional[Callable[[dict], None]] = None,
    should_stop: Optional[Callable[[], bool]] = None,
) -> dict:
    """Fetch every uncached query in the range, waiting out blocks.

    ``polite_delay`` spaces requests so a long run does not look like an attack;
    this job has time, and being refused costs far more than being slow.
    ``max_wait`` caps a single back-off, and ``total_budget`` caps the whole run
    for callers that cannot block indefinitely. Neither is required.

    Returns counts. Safe to interrupt: every answered query is already on disk.
    """
    work = plan_work(start, end, subcategories, category, maturity)
    stats = {"planned": len(work), "fetched": 0, "empty": 0, "blocked_waits": 0,
             "skipped": 0, "elapsed": 0.0, "stopped_early": False}
    started = time.monotonic()
    backoff = 30.0

    for i, (day, mat, cat, sub) in enumerate(work):
        if should_stop and should_stop():
            stats["stopped_early"] = True
            break
        if total_budget is not None and time.monotonic() - started > total_budget:
            stats["stopped_early"] = True
            break

        while True:
            try:
                rows = _fetch_one(day, mat, cat, sub)
                _write_cache(day, mat, cat, sub, rows)
                stats["fetched"] += 1
                if not rows:
                    stats["empty"] += 1
                backoff = 30.0  # recovered; forget the previous escalation
                break
            except Blocked as exc:
                if total_budget is not None and time.monotonic() - started > total_budget:
                    stats["stopped_early"] = True
                    break
                if should_stop and should_stop():
                    stats["stopped_early"] = True
                    break
                stats["blocked_waits"] += 1
                wait = min(backoff, max_wait) * (0.75 + random.random() * 0.5)
                if progress:
                    progress({**stats, "index": i, "waiting": round(wait),
                              "reason": str(exc), "current": cache_key(day, mat, cat, sub)})
                time.sleep(wait)
                backoff = min(backoff * 2, max_wait)

        if stats["stopped_early"]:
            break
        if progress and i % 5 == 0:
            progress({**stats, "index": i, "waiting": 0, "reason": "",
                      "current": cache_key(day, mat, cat, sub)})
        time.sleep(polite_delay)

    stats["elapsed"] = round(time.monotonic() - started, 1)
    return stats


def cache_summary() -> dict:
    """What the cache currently holds."""
    if not CACHE_DIR.exists():
        return {"files": 0, "mb": 0.0, "dates": 0, "earliest": None, "latest": None}
    files = [p for p in CACHE_DIR.glob("*.json")]
    days = set()
    for p in files:
        try:
            days.add(datetime.strptime(p.stem.split("_")[0], "%d-%b-%Y").date())
        except (ValueError, IndexError):
            continue
    total = sum(p.stat().st_size for p in files)
    return {
        "files": len(files),
        "mb": round(total / (1024 * 1024), 2),
        "dates": len(days),
        "earliest": min(days).isoformat() if days else None,
        "latest": max(days).isoformat() if days else None,
    }


def _cli() -> None:
    import argparse

    ap = argparse.ArgumentParser(description="Backfill AMFI AUM into the local cache.")
    ap.add_argument("--start", help="YYYY-MM-DD (default: 90 days before end)")
    ap.add_argument("--end", help="YYYY-MM-DD (default: today)")
    ap.add_argument("--daily", action="store_true",
                    help="fetch only the most recent 5 days, for a scheduled top-up")
    ap.add_argument("--delay", type=float, default=1.5, help="seconds between requests")
    ap.add_argument("--max-wait", type=float, default=1800.0,
                    help="cap on a single back-off, in seconds")
    ap.add_argument("--budget", type=float, default=None,
                    help="stop after this many seconds; omit to run to completion")
    args = ap.parse_args()

    end = date.fromisoformat(args.end) if args.end else date.today()
    if args.daily:
        start = end - timedelta(days=5)
    elif args.start:
        start = date.fromisoformat(args.start)
    else:
        start = end - timedelta(days=90)

    before = cache_summary()
    print(f"cache before: {before['files']} files, {before['mb']} MB, {before['dates']} dates")
    print(f"range: {start} to {end}")
    planned = plan_work(start, end)
    print(f"queries to fetch: {len(planned)}")
    if not planned:
        print("nothing to do — everything in range is already cached.")
        return

    def show(s: dict) -> None:
        if s.get("waiting"):
            print(f"  [{s['fetched']}/{s['planned']}] blocked ({s['reason']}) — "
                  f"waiting {s['waiting']}s, then resuming at {s['current']}", flush=True)
        else:
            print(f"  [{s['fetched']}/{s['planned']}] {s['current']}", flush=True)

    stats = backfill(start, end, polite_delay=args.delay, max_wait=args.max_wait,
                     total_budget=args.budget, progress=show)
    after = cache_summary()
    print()
    print(f"fetched {stats['fetched']} ({stats['empty']} with no data), "
          f"waited out {stats['blocked_waits']} block(s), {stats['elapsed']}s")
    print(f"cache after: {after['files']} files, {after['mb']} MB, "
          f"{after['dates']} dates ({after['earliest']} to {after['latest']})")
    if stats["stopped_early"]:
        print("stopped before finishing — re-run to continue from here.")


if __name__ == "__main__":
    _cli()
