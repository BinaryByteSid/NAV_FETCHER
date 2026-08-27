"""
Publish the local AUM cache to a Hugging Face Dataset, and pull it on the app.

The collector runs where AMFI will answer -- a residential connection with a
disk that persists. The deployed app runs where neither is true. This moves the
cache between them so the Space never calls AMFI at all, and therefore cannot
be throttled or left hanging by it.

Datasets are free and versioned. The payload is a few MB of JSON, well inside
what the free tier is meant for.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Optional, Tuple

CACHE_DIR = Path(__file__).resolve().parent / "api_cache"

# Override with AUM_CACHE_REPO to point at a different dataset.
DEFAULT_REPO = "Sidh17/nav-fetcher-aum-cache"


def repo_id() -> str:
    return os.environ.get("AUM_CACHE_REPO", DEFAULT_REPO)


def _token() -> Optional[str]:
    """A write token, from the environment only.

    Never read from the git remote or hard-coded: a token in source outlives
    every place you meant it to be used.
    """
    return os.environ.get("HF_TOKEN") or os.environ.get("HUGGING_FACE_HUB_TOKEN")


def available() -> Tuple[bool, str]:
    """Whether huggingface_hub is importable."""
    try:
        import huggingface_hub  # noqa: F401
        return True, ""
    except ImportError:
        return False, "huggingface_hub is not installed (pip install huggingface_hub)"


def push(private: bool = True, token: Optional[str] = None) -> str:
    """Upload the local cache to the dataset. Returns a summary line.

    Defaults to a private dataset. The contents are AMFI's published AUM
    figures, so a public one is defensible, but that is the user's call to make
    deliberately rather than a default they inherit.
    """
    ok, why = available()
    if not ok:
        raise RuntimeError(why)
    from huggingface_hub import HfApi

    tok = token or _token()
    if not tok:
        raise RuntimeError(
            "No HF token found. Set HF_TOKEN in the environment; do not paste it into a file."
        )
    if not CACHE_DIR.exists() or not any(CACHE_DIR.glob("*.json")):
        raise RuntimeError("Local cache is empty — run the backfill before pushing.")

    api = HfApi(token=tok)
    api.create_repo(repo_id=repo_id(), repo_type="dataset", private=private, exist_ok=True)
    api.upload_folder(
        folder_path=str(CACHE_DIR),
        repo_id=repo_id(),
        repo_type="dataset",
        allow_patterns=["*.json"],       # never ship .tmp files from an interrupted write
        commit_message="Update AMFI AUM cache",
    )
    n = len(list(CACHE_DIR.glob("*.json")))
    return f"pushed {n} cached queries to {repo_id()} ({'private' if private else 'public'})"


def pull(token: Optional[str] = None) -> str:
    """Download the dataset into the local cache directory.

    Called at app start-up on the deployment. A failure here is not fatal: the
    app falls back to whatever it already has, which is the same position it
    would be in without the dataset at all.
    """
    ok, why = available()
    if not ok:
        raise RuntimeError(why)
    from huggingface_hub import snapshot_download

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    snapshot_download(
        repo_id=repo_id(),
        repo_type="dataset",
        local_dir=str(CACHE_DIR),
        token=token or _token(),         # unnecessary for a public dataset
        allow_patterns=["*.json"],
    )
    n = len(list(CACHE_DIR.glob("*.json")))
    return f"pulled {n} cached queries from {repo_id()}"


def pull_quietly() -> str:
    """pull() that reports its failure instead of raising, for start-up use."""
    try:
        return pull()
    except Exception as exc:                     # noqa: BLE001 - start-up must not die here
        return f"AUM cache not synced: {type(exc).__name__}: {exc}"


def _cli() -> None:
    import argparse

    ap = argparse.ArgumentParser(description="Sync the AUM cache with a HF Dataset.")
    ap.add_argument("action", choices=["push", "pull", "status"])
    ap.add_argument("--public", action="store_true",
                    help="create the dataset as public (default: private)")
    args = ap.parse_args()

    if args.action == "status":
        n = len(list(CACHE_DIR.glob("*.json"))) if CACHE_DIR.exists() else 0
        ok, why = available()
        print(f"repo         : {repo_id()}")
        print(f"local cache  : {n} json files")
        print(f"hub library  : {'installed' if ok else why}")
        print(f"token in env : {'yes' if _token() else 'no'}")
        return

    print(push(private=not args.public) if args.action == "push" else pull())


if __name__ == "__main__":
    _cli()
