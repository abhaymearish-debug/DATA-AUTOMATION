"""Days the trade was shut.

A per-day figure is only honest if the divisor is the number of days the thing
could actually sell. A dry day, a hartal, a local festival closure - none of
them are bad trading, but divided across the whole calendar they read as
exactly that.

Two calendars, because two different things close for different reasons:

    shop        a KSBC outlet is shut - dry day, hartal, local closure. This
                is what the per-day rate in Shop Sales Analysis divides by.
    warehouse   a KSBC warehouse does not issue. Recorded here; no report
                divides by it yet.

A closure is never one shop or one warehouse: when the trade shuts, it shuts
together. So an entry is a date and a reason, and it applies to all of them -
one line for a dry day, not 291. The reason is free text, there to be read by
a person next March wondering why the 2nd looks light, not to be counted.
"""

from __future__ import annotations

import json
import os
import threading
import uuid
from datetime import date, datetime, timedelta
from pathlib import Path

from . import config

KINDS = ("shop", "warehouse")
_LOCK = threading.Lock()


def _file() -> Path:
    return config.WORKSPACE_ROOT / "_leaves" / "leaves.json"


def _read() -> dict:
    f = _file()
    if not f.is_file():
        return {k: [] for k in KINDS}
    try:
        got = json.loads(f.read_text())
    except (json.JSONDecodeError, OSError):
        return {k: [] for k in KINDS}
    return {k: list(got.get(k) or []) for k in KINDS}


def _write(data: dict) -> None:
    f = _file()
    f.parent.mkdir(parents=True, exist_ok=True)
    tmp = f.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2))
    os.replace(tmp, f)


def _day(value) -> date | None:
    try:
        return date.fromisoformat(str(value)[:10])
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------------------
# Reading
# ---------------------------------------------------------------------------


def load(kind: str) -> list[dict]:
    """Every recorded leave of one kind, newest date first."""
    if kind not in KINDS:
        return []
    rows = [r for r in _read().get(kind, []) if _day(r.get("date"))]
    rows.sort(key=lambda r: (r["date"], r.get("added_at", "")), reverse=True)
    return rows


def in_month(kind: str, year: int, month: int) -> list[dict]:
    stamp = f"{year:04d}-{month:02d}"
    return [r for r in load(kind) if str(r.get("date", "")).startswith(stamp)]


def closed_days(kind: str, start: date, end: date) -> set:
    """The dates in [start, end] the trade was shut."""
    out: set = set()
    for row in load(kind):
        day = _day(row.get("date"))
        if day and start <= day <= end:
            out.add(day)
    return out


def span_days(start: date, end: date) -> int:
    """The window's length, as the reports have always counted it."""
    return max(1, (end - start).days)


def open_days(start: date, end: date) -> int:
    """Trading days in the window: its length, less the days shops were shut."""
    return max(1, span_days(start, end) - len(closed_days("shop", start, end)))


# ---------------------------------------------------------------------------
# Writing
# ---------------------------------------------------------------------------


def add(kind: str, days: list[str], reason: str, user: str = "",
        monthly: bool = False) -> dict:
    """Record closed days. Every leave applies to the whole trade.

    `monthly` repeats each picked date on the same day of every later month in
    its year - the 1st is a dry day every month, and clicking twelve of them
    is twelve chances to miss one.
    """
    if kind not in KINDS:
        return {"error": "Unknown kind of leave."}

    picked = sorted({d for d in (_day(x) for x in days) if d})
    if not picked:
        return {"error": "Pick at least one day on the calendar."}
    if monthly:
        picked = sorted(set(picked) | {r for d in picked for r in _rest_of_year(d)})

    reason = (reason or "").strip()[:200]
    stamp = datetime.now().isoformat(timespec="seconds")

    fresh = 0
    with _LOCK:
        data = _read()
        rows = data.get(kind, [])
        for day in picked:
            # The same day twice is a duplicate, not a second closure - it
            # would show as two lines saying one thing.
            same = next((r for r in rows if r.get("date") == day.isoformat()), None)
            if same:
                if reason:
                    same["reason"] = reason
                same["scope"] = "all"
                same["names"] = []
                continue
            rows.append({"id": uuid.uuid4().hex[:8], "date": day.isoformat(),
                         "scope": "all", "names": [], "reason": reason,
                         "added_by": user, "added_at": stamp})
            fresh += 1
        data[kind] = rows
        _write(data)

    return {"added": fresh, "already": len(picked) - fresh}


def _rest_of_year(day: date) -> list:
    """The same day-of-month in every later month of that year.

    A 31st simply has no February, so months that cannot hold the date are
    skipped rather than rounded to a day the shops were open.
    """
    out = []
    for month in range(day.month + 1, 13):
        try:
            out.append(date(day.year, month, day.day))
        except ValueError:
            continue
    return out


def remove(kind: str, entry_id: str) -> dict:
    if kind not in KINDS:
        return {"error": "Unknown kind of leave."}
    with _LOCK:
        data = _read()
        rows = data.get(kind, [])
        keep = [r for r in rows if r.get("id") != entry_id]
        if len(keep) == len(rows):
            return {"error": "That leave is no longer there."}
        data[kind] = keep
        _write(data)
    return {"removed": 1}


def month_days(year: int, month: int) -> list[date]:
    first = date(year, month, 1)
    nxt = date(year + (month == 12), (month % 12) + 1, 1)
    return [first + timedelta(days=i) for i in range((nxt - first).days)]
