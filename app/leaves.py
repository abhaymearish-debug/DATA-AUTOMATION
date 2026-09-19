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

Each entry is one date, and says who it applies to: everybody (`scope: all`),
or the named warehouses / shops it closed. A statewide dry day is one entry,
not 291. The reason is free text - it is there to be read by a person next
March wondering why the 2nd looks light, not to be counted.
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


def closed_days(kind: str, start: date, end: date, name: str = "") -> set:
    """The dates in [start, end] this entity was shut.

    With no name, only the days everybody was shut - which is the divisor the
    network total uses. With one, those days plus the days that warehouse or
    shop was closed on its own.
    """
    want = (name or "").strip().upper()
    out: set = set()
    for row in load(kind):
        day = _day(row.get("date"))
        if not day or day < start or day > end:
            continue
        if row.get("scope") == "all":
            out.add(day)
        elif want and want in {str(n).strip().upper() for n in (row.get("names") or [])}:
            out.add(day)
    return out


def shop_closed_days(start: date, end: date, bond: str = "") -> set:
    """The dates a bond could not trade at all.

    Everybody-shut days always count. A bond also loses a day when every one
    of its shops is individually marked closed on it - a local hartal that
    took out one bond and nobody else. One shop of twenty being shut does not
    stop the bond trading, so it does not move the bond's divisor; it is still
    recorded, and still shows on the leaves page.
    """
    out = closed_days("shop", start, end)
    bond = (bond or "").strip().upper()
    if not bond:
        return out

    from . import reports_api

    mine = {code for code, info in reports_api.load_master().items()
            if str(info.get("bond") or "").strip().upper() == bond}
    if not mine:
        return out

    named = {str(info.get("name") or "").strip().upper()
             for code, info in reports_api.load_master().items() if code in mine}
    for row in load("shop"):
        day = _day(row.get("date"))
        if not day or day < start or day > end or row.get("scope") == "all":
            continue
        shut = {str(n).strip().upper() for n in (row.get("names") or [])}
        if named and named <= shut:
            out.add(day)
    return out


def span_days(start: date, end: date) -> int:
    """The window's length, as the reports have always counted it."""
    return max(1, (end - start).days)


def open_days(start: date, end: date, bond: str = "") -> int:
    """Trading days in the window: its length, less the days nothing opened."""
    return max(1, span_days(start, end) - len(shop_closed_days(start, end, bond)))


# ---------------------------------------------------------------------------
# Writing
# ---------------------------------------------------------------------------


def add(kind: str, days: list[str], names: list[str], reason: str,
        user: str = "") -> dict:
    """Record one or more closed days. Empty `names` means everybody."""
    if kind not in KINDS:
        return {"error": "Unknown kind of leave."}

    picked = sorted({d for d in (_day(x) for x in days) if d})
    if not picked:
        return {"error": "Pick at least one day on the calendar."}

    clean = sorted({str(n).strip() for n in (names or []) if str(n).strip()})
    scope = "some" if clean else "all"
    reason = (reason or "").strip()[:200]
    stamp = datetime.now().isoformat(timespec="seconds")

    fresh = 0
    with _LOCK:
        data = _read()
        rows = data.get(kind, [])
        for day in picked:
            # The same day, the same people, twice is a duplicate, not a second
            # closure - it would show as two lines saying one thing.
            same = next((r for r in rows
                         if r.get("date") == day.isoformat()
                         and r.get("scope") == scope
                         and sorted(r.get("names") or []) == clean), None)
            if same:
                if reason:
                    same["reason"] = reason
                continue
            rows.append({"id": uuid.uuid4().hex[:8], "date": day.isoformat(),
                         "scope": scope, "names": clean, "reason": reason,
                         "added_by": user, "added_at": stamp})
            fresh += 1
        data[kind] = rows
        _write(data)

    who = "everyone" if scope == "all" else f"{len(clean)} named"
    return {"added": fresh, "already": len(picked) - fresh,
            "scope": scope, "who": who}


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


# ---------------------------------------------------------------------------
# Who a leave can be set against
# ---------------------------------------------------------------------------


def warehouses() -> list[str]:
    from . import reports_api
    return sorted({w for whs in reports_api.WAREHOUSE_CLUSTERS.values() for w in whs})


def shops() -> list[dict]:
    """Every shop in master, with the bond it belongs to."""
    from . import reports_api
    out = []
    for code, info in reports_api.load_master().items():
        name = str(info.get("name") or "").strip()
        if not name:
            continue
        out.append({"name": name, "code": str(code),
                    "bond": str(info.get("bond") or "").strip().upper()})
    out.sort(key=lambda s: (s["bond"], s["name"]))
    return out


def bonds() -> list[str]:
    from . import reports_api
    return sorted({b for bs in reports_api.bond_clusters().values() for b in bs})


def month_days(year: int, month: int) -> list[date]:
    first = date(year, month, 1)
    nxt = date(year + (month == 12), (month % 12) + 1, 1)
    return [first + timedelta(days=i) for i in range((nxt - first).days)]
