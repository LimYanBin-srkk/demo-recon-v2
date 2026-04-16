"""Persistent page-usage tracker for the OCR and Reconciliation pipelines.

Pages are counted only when PDFs are actually processed through a pipeline.
Tracking includes a per-source breakdown and a run history log.
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Literal

MAX_PAGES   = 1_000
_STORE_FILE = Path(__file__).resolve().parents[1] / "docs" / "page_usage.json"

Source = Literal["ocr", "recon"]

_EMPTY: dict = {
    "total": 0,
    "by_source": {"ocr": 0, "recon": 0},
    "history": [],   # [{ts, source, pages}]  — newest first
}


def _load() -> dict:
    if _STORE_FILE.exists():
        try:
            data = json.loads(_STORE_FILE.read_text(encoding="utf-8"))
            # Back-compat: migrate old format
            if "pages_used" in data and "total" not in data:
                data = {**_EMPTY, "total": int(data.get("pages_used", 0))}
            return data
        except (json.JSONDecodeError, OSError):
            pass
    return dict(_EMPTY)


def _save(data: dict) -> None:
    _STORE_FILE.parent.mkdir(parents=True, exist_ok=True)
    _STORE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def get_usage() -> int:
    """Total pages consumed across both pipelines."""
    return int(_load().get("total", 0))


def get_usage_by_source() -> dict[str, int]:
    """Per-source page counts: {"ocr": n, "recon": n}."""
    data = _load()
    bs = data.get("by_source", {})
    return {"ocr": int(bs.get("ocr", 0)), "recon": int(bs.get("recon", 0))}


def get_history(limit: int = 50) -> list[dict]:
    """Return the most recent *limit* run records (newest first)."""
    return _load().get("history", [])[:limit]


def get_remaining() -> int:
    """Pages remaining in the quota."""
    return max(0, MAX_PAGES - get_usage())


def add_pages(n: int, source: Source = "ocr") -> int:
    """Record *n* pages from *source* ('ocr' or 'recon'). Returns new total."""
    data = _load()
    data.setdefault("total", 0)
    data.setdefault("by_source", {"ocr": 0, "recon": 0})
    data.setdefault("history", [])

    data["total"] = int(data["total"]) + n
    data["by_source"][source] = int(data["by_source"].get(source, 0)) + n
    data["history"].insert(0, {
        "ts":     datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "source": source,
        "pages":  n,
    })
    # Keep at most 200 history entries
    data["history"] = data["history"][:200]

    _save(data)
    return data["total"]


def reset() -> None:
    """Reset all counters to zero (fresh start)."""
    _save(dict(_EMPTY))
