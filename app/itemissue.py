"""Item issue consolidation - the secondary sales analysis, month against month.

WHERE THE FIGURES COME FROM
---------------------------
KSBC's ERP exports one "Item Wise Issue Consolidation Report" per warehouse:
an HTML table saved with a .xls extension, carrying every item that moved in a
date range. Each row splits its quantity across FTN, STN, GTN, INTER-WH(Out),
Consumerfed, Other and Total, and each of those into cases and bottles.

The report the office circulates uses five of those columns:

    STN      transfers in from another supplier's stock
    GTN      the ordinary issue
    TOTAL    STN + GTN
    C FED    Consumerfed
    BAR      the export's Other column
    <as on>  TOTAL + C FED + BAR

which is the export's own Total column, verified on the September pull:
Kollam 299 GTN + 191 C FED = 490, and 490 is what the file totals. FTN and
INTER-WH(Out) are not part of it.

The markup is malformed in a way that matters: data rows carry no opening
<tr>, so the file cannot be walked as rows of a table. It is split on the
closing tag instead, and any run of eighteen cells whose first is a number is
a data row.
"""

from __future__ import annotations

import json
import re
import shutil
import threading
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

from . import config

# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------

_CELL = re.compile(r"<td[^>]*>(.*?)</td>", re.S)
_TAG = re.compile(r"<[^>]+>")
_WAREHOUSE = re.compile(r"Warehouse\s*:\s*<b>([^<]+)</b>", re.I)
_PERIOD = re.compile(r"Report Period\s*:\s*<b>\s*([\d]{2}-[A-Za-z]{3}-[\d]{4})\s*/\s*"
                     r"([\d]{2}-[A-Za-z]{3}-[\d]{4})\s*</b>", re.I)

_MON = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun",
     "jul", "aug", "sep", "oct", "nov", "dec"], start=1)}

MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July",
               "August", "September", "October", "November", "December"]

# The five figures the report is made of, in the order it prints them.
LEGS = [("stn", "STN"), ("gtn", "GTN"), ("cfed", "C FED"), ("bar", "BAR")]


def _cells_of(chunk: str) -> list[str]:
    return [_TAG.sub("", c).replace("&amp;", "&").strip() for c in _CELL.findall(chunk)]


def _num(text: str) -> int:
    text = (text or "").replace(",", "").strip()
    try:
        return int(float(text))
    except ValueError:
        return 0


def _day(text: str) -> date | None:
    try:
        d, m, y = text.strip().split("-")
        return date(int(y), _MON[m[:3].lower()], int(d))
    except (ValueError, KeyError):
        return None


def warehouse_of(raw: str) -> str:
    """'WH-KOLLAM FL9-KLM-01/2026-27' -> 'KOLLAM'.

    The licence number rides along in the same string and changes every year,
    so the name is everything up to the first run of two spaces or the licence
    token, whichever comes first.
    """
    name = re.sub(r"^\s*WH[-\s]*", "", raw or "", flags=re.I).strip()
    name = re.split(r"\s+(?=(?:R?FL\s*-?\s*9|FL\s*-?\s*09|FL09|RFL9))", name, maxsplit=1,
                    flags=re.I)[0]
    return re.sub(r"\s+", " ", name).strip().upper()


def parse_file(path) -> dict:
    """One warehouse's export: its name, its period, and its four legs."""
    text = Path(path).read_text(encoding="utf-8", errors="replace")

    wh = _WAREHOUSE.search(text)
    period = _PERIOD.search(text)
    legs = {"stn": 0, "gtn": 0, "cfed": 0, "bar": 0}
    stated = 0
    rows = 0

    for chunk in text.split("</tr>"):
        cells = _cells_of(chunk)
        if len(cells) < 18 or not cells[0].isdigit():
            continue
        rows += 1
        legs["stn"] += _num(cells[6])
        legs["gtn"] += _num(cells[8])
        legs["cfed"] += _num(cells[12])
        legs["bar"] += _num(cells[14])
        stated += _num(cells[16])

    return {
        "warehouse": warehouse_of(wh.group(1)) if wh else "",
        "licence": (wh.group(1).strip() if wh else ""),
        "start": _day(period.group(1)) if period else None,
        "end": _day(period.group(2)) if period else None,
        "legs": legs,
        # The export's own Total, kept so a parse that drifts from the file it
        # came out of announces itself instead of quietly disagreeing.
        "stated_total": stated,
        "rows": rows,
    }


def total_of(legs: dict) -> int:
    return sum(int(legs.get(k, 0) or 0) for k, _ in LEGS)


# ---------------------------------------------------------------------------
# Storage - one folder per period
# ---------------------------------------------------------------------------


def root() -> Path:
    return config.CLAUDE_ROOT / "ITEM ISSUE" / "_periods"


def period_key(start: date, end: date) -> str:
    return f"{start.isoformat()}_{end.isoformat()}"


def period_of(key: str) -> tuple[date, date] | None:
    try:
        a, b = key.split("_")
        return date.fromisoformat(a), date.fromisoformat(b)
    except ValueError:
        return None


def period_label(key: str) -> str:
    span = period_of(key)
    if not span:
        return key
    start, end = span
    if (start.year, start.month) == (end.year, end.month):
        return (f"{start.day}–{end.day} {MONTH_NAMES[end.month - 1]} {end.year}")
    return (f"{start.day} {MONTH_NAMES[start.month - 1]} – "
            f"{end.day} {MONTH_NAMES[end.month - 1]} {end.year}")


def store(paths: list[Path], expect: tuple | None = None) -> dict:
    """File a batch under the period it covers.

    Every file in one pull covers the same range, so a batch that does not is
    two pulls mixed together and is refused rather than half-filed.

    `expect` is the period the person picked on the upload screen. It is a
    check, not an override: files that name a different range are refused with
    both stated, so a September pull cannot be filed under August by a slip of
    the picker. It is also the answer for an export that names no range at
    all - those used to be dropped on the floor with nothing to file them
    under, which is the whole reason the picker exists.
    """
    want = period_key(expect[0], expect[1]) if expect else ""
    by_period: dict[str, list[tuple[Path, dict]]] = defaultdict(list)
    undated: list[tuple[Path, dict]] = []
    skipped: list[str] = []

    for p in paths:
        parsed = parse_file(p)
        if not parsed["warehouse"]:
            skipped.append(Path(p).name)
            continue
        if not parsed["start"] or not parsed["end"]:
            undated.append((p, parsed))
            continue
        by_period[period_key(parsed["start"], parsed["end"])].append((p, parsed))

    if len(by_period) > 1:
        spans = ", ".join(period_label(k) for k in sorted(by_period))
        return {"error": f"Those files cover more than one period ({spans}). "
                         "Upload one pull at a time."}

    if by_period:
        key, items = next(iter(by_period.items()))
        if want and key != want:
            return {"error": f"You picked {period_label(want)}, but those files "
                             f"are for {period_label(key)}. Change the period or "
                             f"upload the files for {period_label(want)}."}
    elif want and undated:
        key, items = want, []
    elif not by_period:
        return {"error": "None of those files look like a KSBC item issue "
                         "consolidation export - no warehouse or report period "
                         "could be read out of them."}

    # An export that named no range still belongs to the pull: it goes in under
    # the period the rest of the batch - or the person - answered for.
    items = list(items) + undated
    folder = root() / key
    if folder.exists():
        shutil.rmtree(folder)
    folder.mkdir(parents=True, exist_ok=True)
    for p, _parsed in items:
        shutil.copy2(p, folder / Path(p).name)

    _CACHE.pop(key, None)
    return {"period": key, "label": period_label(key), "files": len(items),
            "warehouses": len({d["warehouse"] for _p, d in items}),
            "undated": len(undated), "skipped": skipped}


def source_for(key: str, prior: str = "") -> dict:
    """Which pulls answered this report - one per period, not one per file.

    A pull is fourteen-odd warehouse exports that all cover the same range,
    so naming every file would be fourteen lines saying the same thing. The
    period, the count and the warehouses behind it is the whole answer.
    """
    from . import reports_api as _r

    def leg(k: str, label: str, tone: str = "") -> dict:
        span = period_of(k) if k else None
        if not span:
            return {}
        folder = root() / k
        files = sorted(folder.glob("*.xls*")) if folder.is_dir() else []
        if not files:
            return {}
        houses = sorted({warehouse_of(parse_file(f).get("warehouse") or "")
                         for f in files} - {""})
        return _r.src_leg(label, [{
            "from": span[0].isoformat(), "to": span[1].isoformat(),
            "kind": "item issue", "stream": _r.STREAM["item_issue"],
            # The count itself is on the row, said the same way it is said on
            # every other report; what only this leg can add is how many
            # warehouses those exports came off.
            "files": len(files),
            "detail": f"across {len(houses)} warehouses" if len(houses) > 1 else "",
        }], tone=tone)

    legs = [l for l in (leg(key, "This period"),
                        leg(prior, "Compared against", "green")) if l]
    return _r.src_block(legs)


_CACHE: dict = {}
_LOCK = threading.Lock()


def periods() -> list[dict]:
    """Stored pulls, newest end date first."""
    folder = root()
    if not folder.is_dir():
        return []
    out = []
    for path in folder.iterdir():
        span = period_of(path.name)
        if not path.is_dir() or not span:
            continue
        out.append({"key": path.name, "label": period_label(path.name),
                    "start": span[0].isoformat(), "end": span[1].isoformat(),
                    "files": len(list(path.glob("*.xls")))})
    return sorted(out, key=lambda r: (r["end"], r["start"]), reverse=True)


def load(key: str) -> dict[str, dict]:
    """Every warehouse's legs for one stored period, keyed by warehouse."""
    folder = root() / key
    if not folder.is_dir():
        return {}
    stamp = max((p.stat().st_mtime_ns for p in folder.glob("*.xls")), default=0)
    with _LOCK:
        hit = _CACHE.get(key)
        if hit and hit[0] == stamp:
            return hit[1]

    out: dict[str, dict] = {}
    for p in sorted(folder.glob("*.xls")):
        parsed = parse_file(p)
        name = parsed["warehouse"]
        if not name:
            continue
        cell = out.setdefault(name, {"stn": 0, "gtn": 0, "cfed": 0, "bar": 0,
                                     "stated_total": 0})
        for k, _label in LEGS:
            cell[k] += parsed["legs"][k]
        cell["stated_total"] += parsed["stated_total"]

    with _LOCK:
        _CACHE[key] = (stamp, out)
    return out


# ---------------------------------------------------------------------------
# The industry figure - typed in, because it is not in any export
# ---------------------------------------------------------------------------


def _industry_file() -> Path:
    return config.WORKSPACE_ROOT / "_industry" / "industry.json"


def industry_all() -> dict:
    path = _industry_file()
    if not path.is_file():
        return {}
    try:
        return json.loads(path.read_text())
    except (json.JSONDecodeError, OSError):
        return {}


def industry_get(key: str) -> dict:
    return industry_all().get(key, {})


def industry_set(key: str, cases: float, prior: float, by: str = "") -> dict:
    from datetime import datetime, timezone
    store_all = industry_all()
    store_all[key] = {"cases": float(cases or 0), "prior": float(prior or 0),
                      "updated": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                      "by": by}
    path = _industry_file()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(store_all, indent=2, sort_keys=True))
    return store_all[key]
