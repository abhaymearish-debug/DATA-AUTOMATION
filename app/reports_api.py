"""Read-only report queries over the live workbooks.

The upload side of this app builds workbooks. This side reads them back so a
report can be browsed in the page — filtered, re-grouped and exported — without
anyone opening Excel.

Nothing here writes. The only mutation path in the service is promote.py.
"""

from __future__ import annotations

import csv
import re
import threading
from collections import defaultdict
from datetime import date, datetime, timedelta
from functools import lru_cache
from pathlib import Path

import openpyxl

from . import config

BPC = {"1000 ML": 9, "750 ML": 12, "500 ML": 18, "375 ML": 24, "180 ML": 48}

# Bond -> ASM cluster. The locked definition: Cluster 1 is the six southern
# bonds, Cluster 2 the five central, Cluster 3 the four northern.
BOND_CLUSTERS: dict[int, list[str]] = {
    1: ["ALAPPUZHA", "ATTINGAL", "KOLLAM", "KOTTARAKARA", "NEDUMANGAD", "PATHANAMTHITTA"],
    2: ["ALUVA", "KOTTAYAM", "THODUPUZHA", "THRISSUR", "TRIPUNITHURA"],
    3: ["KANNUR", "KOZHIKODE", "PALAKKAD", "PERINTHALMANNA"],
}

# ---------------------------------------------------------------------------
# Where a report's figures came from
# ---------------------------------------------------------------------------
#
# Every report on this app answers from files somebody uploaded, and rarely
# from just one. The KSBC portal caps a pull at sixteen days, so any window
# longer than that is the first-half file stitched to the ones after it; and
# a liquidation or target figure reads two streams at once - what the shops
# sold on one side, what was dispatched on the other.
#
# Which file answered which part is a question you ask occasionally and want
# answered exactly. These build the block the Source chip reads: legs, each
# holding the files it drew on, and one plain sentence at the end.

SRC_KIND = {"cumulative": "cumulative upload", "daily": "daily upload"}


def src_windows(chain: list) -> list:
    """Chain segments - the files that tile a window - as the chip reads them."""
    out = []
    for seg in chain or []:
        name = seg.get("name") or Path(seg["path"]).name
        out.append({"from": str(seg["from"])[:10], "to": str(seg["to"])[:10],
                    "kind": SRC_KIND.get(seg.get("kind", ""), seg.get("kind", "")),
                    "name": name})
    return out


def src_files(names, kind: str = "raw upload", title: str = "") -> list:
    """Loose files with no window of their own - a workbook, a month's raws."""
    return [{"title": title, "kind": kind, "name": str(n)}
            for n in (names or []) if n]


def src_leg(label: str, items: list, note: str = "", tone: str = "") -> dict:
    return {"leg": label, "note": note, "tone": tone, "items": items}


def src_block(legs: list, say: str = "") -> dict:
    """One report's provenance. Empty legs drop out; an empty block hides."""
    kept = [l for l in legs if l and l.get("items")]
    return {"legs": kept, "say": say} if kept else {"legs": [], "say": ""}


def src_secondary(sources: list) -> dict:
    """The dispatch leg: the month's built workbook, then the raws after it."""
    book = find_secondary_workbook()
    items = [{"kind": "built workbook" if (book is not None and n == book.name)
                      else "raw upload", "name": n} for n in (sources or [])]
    return src_leg("Secondary dispatch", items, "what left the warehouses", "gold")


def src_stitched(n: int) -> str:
    """The sentence under a window answered by more than one file."""
    if n > 1:
        return (f"{n} uploads, stitched end to end - KSBC caps a pull at 16 days, "
                f"so a longer window is the first-half file plus the ones after it.")
    return "One upload answers this whole window."



def bond_clusters() -> dict[int, list[str]]:
    """The live split, which Settings can change, falling back to the above."""
    from . import bondmap
    try:
        return {int(k): list(v) for k, v in bondmap.load_clusters().items()}
    except Exception:
        return {c: list(b) for c, b in BOND_CLUSTERS.items()}


def cluster_of_bond() -> dict[str, int]:
    return {b: c for c, bonds in bond_clusters().items() for b in bonds}


CLUSTER_OF_BOND = {b: c for c, bonds in BOND_CLUSTERS.items() for b in bonds}

# Bevco warehouse -> cluster, for the warehouse view.
WAREHOUSE_CLUSTERS: dict[int, list[str]] = {
    1: ["ALAPPUZHA", "ATTINGAL", "NEDUMANGAD", "KOLLAM", "KOTTARAKARA",
        "PATHANAMTHITTA", "KARUNAGAPPALLY", "BALARAMAPURAM", "MENAMKULAM", "THIRUVALLA"],
    2: ["THRISSUR", "TRIPUNITHURA", "THODUPUZHA", "KOTTAYAM", "ALUVA",
        "PERUMBAVOOR", "KOTHAMANGALAM", "AYARKKUNNAM", "CHALAKUDY", "KADAVANTHRA"],
    3: ["KANNUR", "KOZHIKODE", "PALAKKAD", "PERINTHALMANNA", "MENONPARA",
        "BATTATHUR", "KALPETTA", "NADUVANNUR"],
}
CLUSTER_OF_WAREHOUSE = {w: c for c, whs in WAREHOUSE_CLUSTERS.items() for w in whs}

_CACHE_LOCK = threading.Lock()
_CACHE: dict[str, tuple[float, object]] = {}


# ---------------------------------------------------------------------------
# Parse a raw file once, ever
# ---------------------------------------------------------------------------
#
# Every report here is a read over workbooks somebody uploaded, and openpyxl
# is the whole cost: about half a second for one day's shop export on a fast
# machine, two or three times that on the box this runs on. A month of days is
# therefore twenty to thirty seconds of parsing - which is what the daily
# report was spending every time the process came up cold, and every time a
# new day was uploaded, because the in-memory cache was keyed on the whole set
# of files rather than on each one.
#
# A raw never changes once it is uploaded. So each file is parsed once and its
# result is written beside the workspace, keyed on the file's own identity -
# name, mtime, size - plus the master data's, since what master says about a
# shop decides which rows are kept. Change either and the key changes and the
# file is parsed again; change nothing and a cold start reads JSON instead of
# a workbook, which is milliseconds.

_CACHE_DIR = config.WORKSPACE_ROOT / "_cache"
_CACHE_KEEP = 600          # files, after which the oldest are swept


def _master_stamp() -> str:
    """What master data is right now, as one short token."""
    path = config.MASTER_DATA
    try:
        return str(path.stat().st_mtime_ns)
    except OSError:
        return "0"


def _parsed(kind: str, path: Path, build, *, uses_master: bool = True):
    """`build()`'s result for this file, from disk if it has been built before.

    Best-effort throughout: a cache that cannot be read or written is simply
    not used, and the report is right either way - only slower.
    """
    import hashlib
    import json

    try:
        st = path.stat()
    except OSError:
        return build()
    stamp = f"{kind}|{path.name}|{st.st_mtime_ns}|{st.st_size}"
    if uses_master:
        stamp += "|" + _master_stamp()
    name = hashlib.sha1(stamp.encode("utf-8")).hexdigest()[:24] + ".json"
    into = _CACHE_DIR / name

    try:
        if into.is_file():
            return json.loads(into.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        pass

    got = build()
    try:
        _CACHE_DIR.mkdir(parents=True, exist_ok=True)
        tmp = into.with_suffix(".part")
        tmp.write_text(json.dumps(got), encoding="utf-8")
        tmp.replace(into)
        _sweep_cache()
    except (OSError, TypeError, ValueError):
        pass
    return got


def _sweep_cache() -> None:
    """Keep the folder from growing for ever: oldest out, newest kept."""
    try:
        files = sorted(_CACHE_DIR.glob("*.json"), key=lambda f: f.stat().st_mtime)
    except OSError:
        return
    for f in files[:-_CACHE_KEEP] if len(files) > _CACHE_KEEP else []:
        try:
            f.unlink()
        except OSError:
            pass


def warm_cache() -> None:
    """Parse anything not parsed yet, off the critical path.

    A deploy comes up with an empty cache, and the first person to open the
    daily report would otherwise pay for every day file in it. This runs on a
    thread at startup: by the time anybody has signed in, the reading is done.
    Best-effort - a failure here costs nothing but the speed it was buying.
    """
    steps = (
        ("master data", lambda: load_master()),
        ("shop day files", lambda: load_shop_daily_raws()),
        ("secondary dispatch", lambda: secondary_lines()),
        ("cumulative periods", lambda: [cumulative_lines(p["path"])
                                        for p in cumulative_periods()]),
        # A window that runs past the last cumulative pull is stitched from
        # day files, which are read a second way - so warm that reading too.
        ("day files, stitched", lambda: [_day_lines(f, load_master())
                                         for f in shop_day_files().values()]),
        ("warehouse stock", lambda: load_stock_rows()),
    )
    for what, run in steps:
        try:
            began = datetime.now()
            run()
            took = (datetime.now() - began).total_seconds()
            if took > 0.5:
                print(f"[warm] {what}: {took:.1f}s")
        except Exception as exc:                       # noqa: BLE001
            print(f"[warm] {what} skipped — {exc}")


def canonical_warehouse(raw) -> str:
    """'WH-KOLLAM FL9-KLM-01/2026-27' -> 'KOLLAM'. Handles FL9 and RFL9."""
    if not raw:
        return ""
    txt = str(raw).strip().upper()
    # The licence code is written a dozen ways - 'FL9-KLM-03', 'RFL9/PTA-02',
    # and 'LFL9/2024/00018' with no space before it - so strip anything from
    # the first FL onwards rather than trying to spell each of them out.
    m = re.match(r"^WH[-\s]+([A-Z .]+?)\s*[A-Z]{0,2}FL\d?[-/].*$", txt) \
        or re.match(r"^WH[-\s]+([A-Z .]+?)(?:\s+R?FL.*)?$", txt)
    name = (m.group(1) if m else txt).strip().rstrip(".")
    return re.sub(r"\s+", " ", name)


def _cases(c, b, pack) -> float:
    out = float(c or 0)
    if b:
        out += float(b) / BPC.get(str(pack or "").strip().upper(), 1)
    return out


def _parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(value).strip(), fmt).date()
        except (ValueError, TypeError):
            continue
    return None


# ---------------------------------------------------------------------------
# Master data
# ---------------------------------------------------------------------------


def load_master() -> dict[str, dict]:
    """Shop code -> {bond, warehouse, name, cat, status}. Cached on file mtime."""
    path = config.MASTER_DATA
    if not path.is_file():
        return {}
    key = f"master:{path.stat().st_mtime_ns}"
    with _CACHE_LOCK:
        hit = _CACHE.get(key)
    if hit:
        return hit[1]

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["16-4-25"] if "16-4-25" in wb.sheetnames else wb[wb.sheetnames[-1]]

    out: dict[str, dict] = {}
    header_seen = False
    idx: dict[str, int] = {}
    for row in ws.iter_rows(values_only=True):
        if not header_seen:
            if row and "Shop Code" in [str(v).strip() if v else "" for v in row]:
                idx = {str(v).strip(): i for i, v in enumerate(row) if v}
                header_seen = True
            continue
        code = row[idx["Shop Code"]] if "Shop Code" in idx else None
        if code is None:
            continue
        out[str(code).strip()] = {
            "bond": str(row[idx.get("Bond", 7)] or "").strip().upper(),
            # 'WH-KOLLAM FL9-KLM-01/2026-27' is nobody's idea of a filter label.
            "warehouse": canonical_warehouse(row[idx["Warehouse Name"]]
                                             if "Warehouse Name" in idx else ""),
            "name": str(row[idx.get("Shop Name", 4)] or "").strip(),
            "cat": str(row[idx.get("CAT", 5)] or "").strip().upper(),
            "status": str(row[idx.get("Status", 8)] or "").strip(),
        }
    wb.close()

    with _CACHE_LOCK:
        _CACHE[key] = (0.0, out)
    return out


# ---------------------------------------------------------------------------
# Brandwise secondary sales
# ---------------------------------------------------------------------------


def find_secondary_workbook() -> Path | None:
    folder = config.CLAUDE_ROOT / "Secondary sales"
    if not folder.is_dir():
        return None
    books = sorted(folder.glob("*SECONDARY SALES ANALYSIS.xlsx"), key=_workbook_rank)
    return books[-1] if books else None


_MONTHS_UP = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY",
              "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]
_MONTH_IN_NAME = re.compile(r"\b(" + "|".join(_MONTHS_UP) + r")\b")
_RANGE_IN_NAME = re.compile(r"(\d{1,2})(?:st|nd|rd|th)\s*-\s*(\d{1,2})(?:st|nd|rd|th)",
                            re.IGNORECASE)


def _workbook_rank(p: Path) -> tuple:
    """Order workbooks by the period they cover, not by mtime.

    Timestamps lie: copying a folder rewrites them, and the copy runs
    newest-first, so the OLDEST workbook ends up looking like the newest. That
    is how a September report can quietly serve August. The month is written on
    the file, so read it; mtime is only the tie-break.
    """
    name = p.name.upper()
    m = _MONTH_IN_NAME.search(name)
    month = _MONTHS_UP.index(m.group(1)) + 1 if m else 0
    r = _RANGE_IN_NAME.search(name)
    end_day = int(r.group(2)) if r else 31
    mtime = p.stat().st_mtime
    return (datetime.fromtimestamp(mtime).year, month, end_day, mtime)


def load_dispatch_lines(workbook: Path) -> list[dict]:
    """Every dispatch line, decorated with bond and canonical warehouse.

    Cached on the workbook's mtime, because a build replaces the file and the
    page must never serve figures from the version before it.
    """
    key = f"lines:{workbook}:{workbook.stat().st_mtime_ns}:{_master_stamp()}"
    with _CACHE_LOCK:
        hit = _CACHE.get(key)
    if hit:
        return hit[1]
    got = [dict(l, date=date.fromisoformat(l["date"]))
           for l in _parsed("dispatch", workbook,
                            lambda: [dict(l, date=l["date"].isoformat())
                                     for l in _dispatch_lines_read(workbook)])]
    with _CACHE_LOCK:
        _CACHE[key] = (0.0, got)
    return got


def _dispatch_lines_read(workbook: Path) -> list[dict]:
    """The workbook itself, read start to finish. Cached by its caller."""
    master = load_master()
    wb = openpyxl.load_workbook(workbook, read_only=True, data_only=True)
    sheets = [s for s in wb.sheetnames if "COMBINED DISPATCH" in s.upper()]

    # A month's analysis workbook keeps its dispatch on a COMBINED DISPATCHES
    # sheet; a KSBC raw export is the same columns on a sheet of its own name.
    # Read either, so a day answers from the file that was uploaded whether or
    # not the month's workbook has been built from it yet.
    ws = rows = header = None
    for name in (sheets or wb.sheetnames):
        cand = wb[name]
        got = cand.iter_rows(values_only=True)
        try:
            head = [str(h).strip() if h else "" for h in next(got)]
        except StopIteration:
            continue
        if {"Issue Cases", "Issue Bottles", "Pack"} <= set(head):
            ws, rows, header = cand, got, head
            break
    if header is None:
        wb.close()
        return []
    ix = {name: i for i, name in enumerate(header)}

    lines: list[dict] = []
    for r in rows:
        if r is None or r[ix.get("Warehouse Name", 1)] is None:
            continue
        # KSBC's export ends with a Total row - 'Total' where the invoice date
        # belongs, the day's cases beside it. It has no warehouse, so it used
        # to fall out by luck; say it outright, because a total counted as a
        # dispatch line would double the day.
        day = _parse_date(r[ix.get("Inv/GTN Date", 11)])
        if day is None:
            continue
        qty = _cases(r[ix["Issue Cases"]], r[ix["Issue Bottles"]], r[ix["Pack"]])
        if qty == 0:
            continue
        lic = str(r[ix.get("Licensee No.", 6)] or "").strip()
        info = master.get(lic, {})
        lines.append({
            "warehouse": canonical_warehouse(r[ix.get("Warehouse Name", 1)]),
            # Fall back to the warehouse when an outlet is not in master, rather
            # than dropping the line — an unmatched outlet is flagged, never lost.
            "bond": info.get("bond") or "",
            "shop": str(r[ix.get("Licensee Name", 7)] or "").strip(),
            "shop_code": lic,
            "cat": info.get("cat", ""),
            "brand": str(r[ix.get("Item Name", 4)] or "").strip().upper(),
            "date": day,
            "cases": qty,
        })
    wb.close()
    return lines


# 'RAW DATA -SEPTEMBER 15TH SECONDARY SALES.xlsx' is what this app writes, but
# the office's own downloads land as 'RAW DATA SEPTEMBER 15TH-SECONDARY SALES'
# and other near misses, so the separators are loose on purpose.
_RAW_SEC_RE = re.compile(
    r"^RAW DATA[\s-]*([A-Z]+)[\s-]*(\d{1,2})\s*(?:ST|ND|RD|TH)?[\s-]*SECONDARY SALES\.xlsx$",
    re.IGNORECASE)


def secondary_raw_files() -> list:
    """Every secondary raw still on disk, newest window first.

    Two places hold them: the live input folder the build reads, and each job's
    own archive - the only copy that survives a folder the build cleans out.
    Both are searched, because a day answered for is a day somebody uploaded,
    whether or not the month's workbook was ever rebuilt from it.
    """
    seen: dict = {}
    folders = [config.CLAUDE_ROOT / "Secondary sales"]
    jobs = config.JOBS_ROOT
    if jobs.is_dir():
        folders += sorted(jobs.glob("*/raw"))

    for folder in folders:
        if not folder.is_dir():
            continue
        for path in folder.glob("RAW DATA*SECONDARY SALES.xlsx"):
            m = _RAW_SEC_RE.match(path.name)
            if not m or path.name.startswith("~$"):
                continue
            month = _MONTH_NUM.get(m.group(1).upper())
            if not month:
                continue
            year = datetime.fromtimestamp(path.stat().st_mtime).year
            try:
                end = date(year, month, int(m.group(2)))
            except ValueError:
                continue
            # One upload per day; the input folder's copy wins over an archive.
            seen.setdefault(end, path)
    return [(end, seen[end]) for end in sorted(seen, reverse=True)]


def secondary_covered_days() -> set:
    """The days secondary uploads answer for - not the days that had dispatch.

    A workbook names the window it was built for ('SEPTEMBER 1st - 16th'), and
    every dated raw adds its own day. Inside that, a day with no line is a day
    nothing moved; outside it, a zero means nobody has uploaded that day yet.
    """
    days: set = set()

    book = find_secondary_workbook()
    if book is not None:
        dated = [l["date"] for l in load_dispatch_lines(book) if l.get("date")]
        if dated:
            hi = max(dated)
            m = _re.search(r"(\d{1,2})(?:st|nd|rd|th)\s*-\s*(\d{1,2})(?:st|nd|rd|th)",
                           book.name, _re.IGNORECASE)
            first, last = (int(m.group(1)), int(m.group(2))) if m else (1, hi.day)
            for day in range(first, last + 1):
                try:
                    days.add(date(hi.year, hi.month, day))
                except ValueError:
                    break

    for end, _path in secondary_raw_files():
        days.add(end)
    return days


def secondary_lines() -> tuple[list, list]:
    """Every dispatch line the app can answer for, and what it read.

    The month's analysis workbook is the authority where it exists. Days it
    does not cover - a month's first upload, or a build that never finished -
    are filled from the raws themselves, newest first, each raw contributing
    only the days no earlier source already claimed. That keeps a cumulative
    export from counting the same invoice twice.
    """
    claimed: set = set()
    lines: list = []
    sources: list = []

    book = find_secondary_workbook()
    if book is not None:
        got = load_dispatch_lines(book)
        if got:
            lines += got
            claimed |= {l["date"] for l in got if l.get("date")}
            sources.append(book.name)

    for _end, path in secondary_raw_files():
        got = load_dispatch_lines(path)
        days = {l["date"] for l in got if l.get("date")}
        fresh = [l for l in got if l.get("date") and l["date"] not in claimed]
        if not fresh:
            continue
        lines += fresh
        claimed |= days
        sources.append(path.name)

    return lines, sources


def brandwise(view: str = "bond", date_from=None, date_to=None,
              bond: str = "", warehouse: str = "", round_off: bool = False) -> dict:
    """Pivot the dispatch lines into the requested view.

    view: 'bond' | 'warehouse' | 'shop'
    """
    lines, sources = secondary_lines()
    if not lines:
        return {"error": "Nothing uploaded for secondary sales yet. "
                         "Upload a day's raw export on the Raw Data Upload page."}
    workbook = find_secondary_workbook() or Path(sources[0])

    dates = [l["date"] for l in lines if l["date"]]
    span = {"min": min(dates).isoformat() if dates else None,
            "max": max(dates).isoformat() if dates else None}
    # The span is everything on disk, which is what the picker offers. It is
    # not what the page should open on: last month's raws are still there, and
    # opening on 1 August in September is a report nobody asked for.
    if dates:
        lo, hi = _latest_month(dates)
        opens = {"from": lo.isoformat(), "to": hi.isoformat()}
    else:
        opens = {"from": None, "to": None}

    sel = []
    for l in lines:
        if date_from and l["date"] and l["date"] < date_from:
            continue
        if date_to and l["date"] and l["date"] > date_to:
            continue
        if bond and l["bond"] != bond:
            continue
        if warehouse and l["warehouse"] != warehouse:
            continue
        sel.append(l)

    brands = sorted({l["brand"] for l in sel})

    if view == "warehouse":
        key_of, cluster_of, label = (lambda l: l["warehouse"]), CLUSTER_OF_WAREHOUSE, "Warehouse"
    elif view == "shop":
        key_of, cluster_of, label = (lambda l: l["shop"] or l["shop_code"]), None, "Shop"
    else:
        key_of, cluster_of, label = (lambda l: l["bond"] or "(unmapped)"), cluster_of_bond(), "Bond"

    grid: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    extra: dict[str, dict] = {}
    for l in sel:
        k = key_of(l)
        grid[k][l["brand"]] += l["cases"]
        if view == "shop":
            extra.setdefault(k, {"bond": l["bond"], "warehouse": l["warehouse"]})

    def rounded(v: float) -> float:
        return round(v) if round_off else round(v, 2)

    rows = []
    if cluster_of is not None:
        for cl in (1, 2, 3):
            members = sorted(k for k in grid if cluster_of.get(k) == cl)
            if not members:
                continue
            ct: dict[str, float] = defaultdict(float)
            for k in members:
                cells = {b: rounded(grid[k].get(b, 0.0)) for b in brands}
                total = rounded(sum(grid[k].get(b, 0.0) for b in brands))
                rows.append({"kind": "row", "name": k, "cells": cells, "total": total})
                for b in brands:
                    ct[b] += grid[k].get(b, 0.0)
            rows.append({
                "kind": "cluster",
                "name": f"CLUSTER {cl}",
                "cells": {b: rounded(ct[b]) for b in brands},
                "total": rounded(sum(ct.values())),
            })
        orphans = sorted(k for k in grid if cluster_of.get(k) is None)
        for k in orphans:
            rows.append({
                "kind": "row", "name": k,
                "cells": {b: rounded(grid[k].get(b, 0.0)) for b in brands},
                "total": rounded(sum(grid[k].get(b, 0.0) for b in brands)),
            })
    else:
        for k in sorted(grid, key=lambda k: -sum(grid[k].values())):
            rows.append({
                "kind": "row", "name": k,
                "meta": extra.get(k, {}),
                "cells": {b: rounded(grid[k].get(b, 0.0)) for b in brands},
                "total": rounded(sum(grid[k].get(b, 0.0) for b in brands)),
            })

    gt: dict[str, float] = defaultdict(float)
    for k in grid:
        for b in brands:
            gt[b] += grid[k].get(b, 0.0)

    return {
        "view": view,
        "label": label,
        "brands": brands,
        "rows": rows,
        "grand": {"cells": {b: rounded(gt[b]) for b in brands},
                  "total": rounded(sum(gt.values()))},
        "span": span,
        "opens": opens,
        # The days an upload stands behind, so the picker can mark them.
        "days_with_data": sorted({l["date"].isoformat() for l in lines if l.get("date")}),
        "workbook": workbook.name,
        "bonds": sorted({l["bond"] for l in lines if l["bond"]}),
        "warehouses": sorted({l["warehouse"] for l in lines if l["warehouse"]}),
        "unmapped": sorted({l["shop"] or l["shop_code"] for l in sel if not l["bond"]}),
        "source_block": src_block(
            [src_secondary(sources)],
            "The month's analysis workbook is the authority where it exists; days it "
            "does not cover are filled from the raws themselves, newest first, so no "
            "invoice is counted twice."),
    }


def brandwise_shops(view: str = "bond", key: str = "", date_from=None, date_to=None,
                    round_off: bool = False, brands: list | None = None) -> dict:
    """The shops inside one bond or warehouse, brand by brand.

    The drill-down under a group row. It reuses the columns the table is
    already showing rather than deriving its own, so a shop's figures sit on
    the same axis as the bond's above it - a column that is empty for this
    bond stays in place instead of shifting everything left.
    """
    lines, _sources = secondary_lines()
    if not lines:
        return {"error": "Nothing uploaded for secondary sales yet."}

    want = (key or "").strip().upper()
    of = (lambda l: (l["warehouse"] or "").upper()) if view == "warehouse" \
        else (lambda l: (l["bond"] or "(unmapped)").upper())

    sel = [l for l in lines
           if of(l) == want
           and not (date_from and l["date"] and l["date"] < date_from)
           and not (date_to and l["date"] and l["date"] > date_to)]
    if not sel:
        return {"shops": [], "brands": brands or []}

    cols = brands or sorted({l["brand"] for l in sel})
    grid: dict = defaultdict(lambda: defaultdict(float))
    names: dict = {}
    for l in sel:
        code = l["shop_code"] or l["shop"]
        grid[code][l["brand"]] += l["cases"]
        names[code] = l["shop"] or code

    def rounded(v: float) -> float:
        return round(v) if round_off else round(v, 2)

    shops = []
    for code in sorted(grid, key=lambda c: -sum(grid[c].values())):
        shops.append({
            "code": code,
            "name": names.get(code, code),
            "cells": {b: rounded(grid[code].get(b, 0.0)) for b in cols},
            "total": rounded(sum(grid[code].values())),
        })
    return {"shops": shops, "brands": cols}


# ---------------------------------------------------------------------------
# Warehouse stock
# ---------------------------------------------------------------------------


def stock_history_path() -> Path:
    return config.CLAUDE_ROOT / "Warehouse stock" / "_history" / "brand_pack_history.csv"


def load_stock_rows() -> list[dict]:
    """Warehouse x brand x pack stock, from the canonical history CSV.

    The daily build already writes this file, so the report reads it rather
    than re-parsing the Bevco exports — which are deleted after ingest anyway.
    Cached on the file's mtime so a fresh build is picked up immediately.
    """
    path = stock_history_path()
    if not path.is_file():
        return []
    key = f"stock:{path.stat().st_mtime_ns}"
    with _CACHE_LOCK:
        hit = _CACHE.get(key)
    if hit:
        return hit[1]

    import csv
    out: list[dict] = []
    with path.open(newline="") as fh:
        for r in csv.DictReader(fh):
            try:
                out.append({
                    "date": r["date"],
                    "warehouse": (r["warehouse"] or "").strip().upper(),
                    "brand": (r["brand"] or "").strip().upper(),
                    "pack": (r["pack"] or "").strip().upper(),
                    "physical": float(r["physical"] or 0),
                    "allotable": float(r["allotable"] or 0),
                    "pending": float(r["pending"] or 0),
                })
            except (KeyError, ValueError):
                # A malformed line is skipped, never silently zeroed.
                continue

    with _CACHE_LOCK:
        _CACHE[key] = (0.0, out)
    return out


def warehouse_stock(as_of: str = "", cluster: int | None = None,
                    warehouse: str = "") -> dict:
    """Stock position per warehouse for one snapshot date."""
    rows = load_stock_rows()
    if not rows:
        return {"error": "No warehouse stock history yet. Upload a day's Bevco exports first."}

    dates = sorted({r["date"] for r in rows})
    day = as_of if as_of in dates else dates[-1]
    sel = [r for r in rows if r["date"] == day]

    if cluster in (1, 2, 3):
        sel = [r for r in sel if CLUSTER_OF_WAREHOUSE.get(r["warehouse"]) == cluster]
    if warehouse:
        sel = [r for r in sel if r["warehouse"] == warehouse]

    by_wh: dict[str, list[dict]] = defaultdict(list)
    for r in sel:
        by_wh[r["warehouse"]].append(r)

    def order(wh: str) -> tuple:
        cl = CLUSTER_OF_WAREHOUSE.get(wh)
        members = WAREHOUSE_CLUSTERS.get(cl, [])
        return (cl or 9, members.index(wh) if wh in members else 99, wh)

    warehouses = []
    grand = {"physical": 0.0, "allotable": 0.0, "pending": 0.0}
    for wh in sorted(by_wh, key=order):
        items = sorted(by_wh[wh], key=lambda r: (r["brand"], r["pack"]))
        totals = {k: sum(i[k] for i in items) for k in ("physical", "allotable", "pending")}
        for k in grand:
            grand[k] += totals[k]
        warehouses.append({
            "name": wh,
            "cluster": CLUSTER_OF_WAREHOUSE.get(wh),
            "rows": [{"brand": i["brand"], "pack": i["pack"],
                      "physical": i["physical"], "allotable": i["allotable"],
                      "pending": i["pending"]} for i in items],
            "totals": totals,
        })

    return {
        "date": day,
        "dates": dates,
        "warehouses": warehouses,
        "grand": grand,
        "all_warehouses": sorted({r["warehouse"] for r in rows if r["date"] == day}),
        "source_block": src_block(
            [src_leg("Warehouse stock (Bevco)",
                     [{"from": day, "to": day, "kind": "daily snapshot",
                       "name": stock_history_path().name}])],
            "A stock position is a photograph, not a total: this is the snapshot "
            "the daily build filed for that date. The Bevco exports themselves are "
            "deleted after ingest, so the history file is the record."),
    }


# ---------------------------------------------------------------------------
# Daily reports (Shop Sales - Daily, Secondary Sales - Daily)
#
# Both are the same grid: bonds down the side, one column per day of the
# period, cluster subtotals and a grand total. Only the source differs.
# ---------------------------------------------------------------------------

import decimal as _decimal
import re as _re


def _as_date(v):
    """Accept an ISO string, a date, or nothing."""
    if not v:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def half_up(v: float) -> int:
    """Round .5 away from zero, the way the office's own reports do.

    Python's round() is banker's rounding: round(448.5) is 448, and the
    published Secondary Sales - Daily shows 449 for exactly that cell. Every
    figure in these grids is rounded from its OWN unrounded aggregate, which is
    why a cluster row can differ by one from the sum of the rows above it.
    """
    return int(_decimal.Decimal(repr(float(v or 0.0))).quantize(0, rounding=_decimal.ROUND_HALF_UP))


def find_shop_workbook() -> Path | None:
    folder = config.CLAUDE_ROOT / "KSBC shop sales"
    if not folder.is_dir():
        return None
    books = sorted(folder.glob("*ANALYSIS.xlsx"), key=_workbook_rank)
    return books[-1] if books else None


_MONTH_NUM = {m.upper(): i for i, m in enumerate(
    ["January", "February", "March", "April", "May", "June",
     "July", "August", "September", "October", "November", "December"], 1)}


_RAW_DAY_RE = _re.compile(r"^([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)\.xlsx$", _re.IGNORECASE)

SHOP_OUT_COLUMNS = ("Shop Code", "Shop Out Cases", "Shop Out Bottles", "Bottle Per Case")


def _shop_rows(ws, day, master: dict) -> list[dict]:
    """Cases sold out of shops on one day, with our mapping applied."""
    return [{"bond": master.get(code, {}).get("bond", ""),
             "warehouse": wh or master.get(code, {}).get("warehouse", ""),
             "shop_code": code, "date": day, "cases": qty}
            for code, wh, qty in _shop_raw_sheet(ws)]


def _shop_raw_sheet(ws) -> list:
    """[shop code, warehouse, cases] off a SupplierWiseShopSaleReport grid."""
    rows = ws.iter_rows(values_only=True)
    try:
        header = [str(h).strip() if h else "" for h in next(rows)]
    except StopIteration:
        return []
    ix = {h: i for i, h in enumerate(header)}
    if any(c not in ix for c in SHOP_OUT_COLUMNS):
        return []
    cS, cC, cB, cP = (ix[c] for c in SHOP_OUT_COLUMNS)
    cW = ix.get("Warehouse Name")

    out = []
    for r in rows:
        if r is None or len(r) <= cS or r[cS] is None:
            continue
        bpc = r[cP] if len(r) > cP else 0
        qty = (r[cC] or 0) + ((r[cB] or 0) / bpc if bpc else 0)
        if not qty:
            continue
        code = str(r[cS]).strip()
        # Which warehouse a shop drew from is the raw's own column - a fact
        # KSBC reports and can change between months - and master is only the
        # fallback for an export that does not carry it.
        wh = canonical_warehouse(r[cW]) if (cW is not None and len(r) > cW) else ""
        out.append([code, wh, float(qty)])
    return out


def load_shop_daily_raws() -> list[dict]:
    """Every day that has been uploaded as its own raw export.

    The daily report is per-day by nature, so it reads the day's own raw file
    rather than the month workbook. That keeps one day independent of every
    other: uploading the 17th gives you the 17th, whether or not the 1st to the
    16th have ever been through the app. The month workbook is a different
    thing, built by the cumulative pipeline, and the cumulative reports use it.
    """
    folder = config.CLAUDE_ROOT / "KSBC shop sales"
    if not folder.is_dir():
        return []

    files = []
    for path in folder.glob("*.xlsx"):
        m = _RAW_DAY_RE.match(path.name)
        if not m:
            continue
        month = _MONTH_NUM.get(m.group(1).upper())
        if month:
            files.append((path, month, int(m.group(2))))
    if not files:
        return []

    # Cached per file rather than per set: uploading the 18th used to
    # re-parse the 1st to the 17th as well, because the key was every file's
    # mtime joined together.
    master = load_master()
    out: list[dict] = []
    for path, month, dom in files:
        year = datetime.fromtimestamp(path.stat().st_mtime).year
        try:
            day = date(year, month, dom)
        except ValueError:
            continue
        for code, wh, qty in _parsed("shopday", path,
                                     lambda p=path: _shop_raw_rows(p),
                                     uses_master=False):
            known = master.get(code, {})
            out.append({"bond": known.get("bond", ""),
                        "warehouse": wh or known.get("warehouse", ""),
                        "shop_code": code, "date": day, "cases": qty})
    return out


def _shop_raw_rows(path: Path) -> list:
    """[shop code, warehouse, cases] out of one day's export, and nothing else.

    Only what the file itself says. Which bond a shop belongs to is our own
    mapping and can change in Settings, so it is looked up fresh each time
    rather than frozen into the cache.
    """
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except (OSError, ValueError):
        return []
    try:
        for sheet in wb.worksheets:
            rows = _shop_raw_sheet(sheet)
            if rows:
                return rows
        return []
    finally:
        wb.close()


def load_shop_daily(workbook: Path) -> list[dict]:
    """Per-day shop sales out of the month's analysis workbook.

    Kept as the fallback for months whose days were ingested before the app
    started storing each day's raw of its own.
    """
    key = f"shopdaily:{workbook}:{workbook.stat().st_mtime_ns}"
    with _CACHE_LOCK:
        hit = _CACHE.get(key)
    if hit:
        return hit[1]

    master = load_master()
    wb = openpyxl.load_workbook(workbook, read_only=True, data_only=True)
    year = datetime.fromtimestamp(workbook.stat().st_mtime).year

    out: list[dict] = []
    for name in wb.sheetnames:
        m = _re.fullmatch(r"([A-Za-z]+)\s+(\d{1,2})", name.strip())
        if not m or m.group(1).upper() not in _MONTH_NUM:
            continue
        try:
            day = date(year, _MONTH_NUM[m.group(1).upper()], int(m.group(2)))
        except ValueError:
            continue
        out.extend(_shop_rows(wb[name], day, master))
    wb.close()

    with _CACHE_LOCK:
        _CACHE[key] = (0.0, out)
    return out


def _declared_period(book: Path | None, lines: list[dict]) -> tuple[date, date]:
    """The period a month workbook covers, not the days that had data.

    A dry Sunday is a column of zeros in the published report, and the 1st is
    shown even when nothing moved. The workbook names its own window
    ('SEPTEMBER 1st - 16th ...'), so that is read first; failing that the window
    runs from the 1st of the data's month to its last dated row.
    """
    seen_lo = min(l["date"] for l in lines)
    seen_hi = max(l["date"] for l in lines)

    # No workbook yet - the raws alone are answering - so the window is the 1st
    # to the last day uploaded, which is what those raws cover.
    if book is None:
        return (seen_lo.replace(day=1), seen_hi)

    m = _re.search(r"(\d{1,2})(?:st|nd|rd|th)\s*-\s*(\d{1,2})(?:st|nd|rd|th)",
                   book.name, _re.IGNORECASE)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        try:
            # A day uploaded after the workbook was last built sits past the
            # window written on it. The report has that day's figures, so the
            # window stretches to hold it rather than hiding it.
            return (seen_hi.replace(day=a), max(seen_hi.replace(day=b), seen_hi))
        except ValueError:
            pass
    return (seen_lo.replace(day=1), seen_hi)


DAILY_KINDS = {
    "shop":      {"title": "Shop Sales - Daily",      "unit": "cases sold"},
    "secondary": {"title": "Secondary Sales - Daily", "unit": "cases invoiced"},
}


def _latest_month(days) -> tuple:
    """This month as far as it actually goes, else the newest month with data.

    Both ends are days that exist, not the 1st and the 31st. A window is a
    claim about what it covers: opening on 1-19 when only the 17th, 18th and
    19th have been uploaded reads as a month-to-date figure that is missing
    half the month. Starting at the first day uploaded says what the report
    can actually answer for, and the gaps inside it are still marked on the
    screen as days nobody has pulled.
    """
    today = date.today()
    want = (today.year, today.month)
    if not any((d.year, d.month) == want for d in days):
        want = max((d.year, d.month) for d in days)
    inside = [d for d in days if (d.year, d.month) == want]
    return min(inside), max(inside)


def _daily_source(kind: str, sources: list, book, lo: date, hi: date) -> dict:
    """What answered a daily grid - one day at a time, or a month's raws."""
    if kind == "secondary":
        return src_block(
            [src_secondary(sources)],
            "A day reports as soon as its raw is uploaded; the month's workbook "
            "takes over for the days it covers.")

    # Shop sales are stored a day at a time, so the grid's provenance is
    # literally the files for the days on screen - which also makes a gap in
    # the middle of a window visible as a day with no file behind it.
    files = {d: path for d, path in shop_day_files().items() if lo <= d <= hi}
    if files:
        items = [{"from": d.isoformat(), "to": d.isoformat(), "kind": "daily upload",
                  "name": files[d].name} for d in sorted(files)]
        gaps = (hi - lo).days + 1 - len(items)
        return src_block(
            [src_leg("Shop sales (KSBC)", items, "what the shops sold")],
            f"{len(items)} day file{'' if len(items) == 1 else 's'} across this window"
            + (f"; {gaps} day{'' if gaps == 1 else 's'} in it has nothing uploaded yet, "
               f"which is why {'it reads' if gaps == 1 else 'they read'} zero." if gaps > 0
               else " - every day in it is accounted for."))
    if book is not None and not book.is_dir():
        return src_block([src_leg("Shop sales (KSBC)",
                                  [{"kind": "month workbook", "name": book.name}])],
                         "Read from the month's workbook: these days predate per-day storage.")
    return src_block([])


def daily_grid(kind: str = "secondary", date_from=None, date_to=None,
               cluster: int | None = None, view: str = "bond") -> dict:
    """Bond (or warehouse) x day grid with cluster subtotals and a grand total."""
    if kind not in DAILY_KINDS:
        return {"error": f"Unknown report '{kind}'."}

    if kind == "secondary":
        # The month's workbook where it exists, the uploaded raws for whatever
        # it does not cover - so a day reports as soon as it is uploaded.
        lines, sources = secondary_lines()
        book = find_secondary_workbook()
        if not lines:
            return {"error": "Nothing uploaded for secondary sales yet. "
                             "Upload a day's raw export on the Raw Data Upload page."}
    else:
        # A day's own raw is the source of truth for the daily report, so one
        # uploaded day stands on its own. Only fall back to the month workbook
        # for months whose days predate per-day storage.
        lines = load_shop_daily_raws()
        book = None
        if lines:
            book = config.CLAUDE_ROOT / "KSBC shop sales"
        else:
            book = find_shop_workbook()
            if not book:
                return {"error": "No shop sales uploaded yet. "
                                 "Upload a day on the Shop Sales - Daily card."}
            lines = load_shop_daily(book)

    lines = [l for l in lines if l.get("date")]
    if not lines:
        return {"error": "That workbook has no dated rows to report on."}

    if kind == "shop" and book is not None and book.is_dir():
        # The window is the days that actually exist, so uploading only the
        # 17th gives a report about the 17th — not fifteen empty columns.
        span_lo = min(l["date"] for l in lines)
        span_hi = max(l["date"] for l in lines)
    else:
        span_lo, span_hi = _declared_period(book, lines)
    # The days an upload actually answers for. Everything outside this is a
    # zero because nobody has pulled that day yet, which is a different fact
    # from a day that traded nothing, and the screen says which is which.
    if kind == "secondary":
        covered = secondary_covered_days()
    elif book is not None and book.is_dir():
        covered = set(shop_day_files())
    else:
        covered = set()
        d = span_lo
        while d <= span_hi:
            covered.add(d)
            d += timedelta(days=1)

    # Opened with no window of its own, the report opens on THIS month. Running
    # across everything uploaded meant a September report opened on 1 August,
    # because August's raws are still there - and nobody opens a daily report
    # to read last month. If this month holds nothing yet, it falls back to the
    # newest month that does, so the page is never blank on purpose.
    days = covered or {span_lo, span_hi}
    if not date_from and not date_to:
        lo, hi = _latest_month(days)
    else:
        lo = _as_date(date_from) or (min(days) if days else span_lo)
        hi = _as_date(date_to) or (max(days) if days else span_hi)

    # Bond is our own mapping; warehouse is KSBC's. Grouping by either asks
    # the same question of the same lines, so only the key and the cluster
    # lookup change - the figures underneath are identical, which is what
    # makes the two views reconcile to the same grand total.
    by_warehouse = view == "warehouse"
    if by_warehouse:
        live = {c: list(w) for c, w in WAREHOUSE_CLUSTERS.items()}
        of_group = dict(CLUSTER_OF_WAREHOUSE)
        field = "warehouse"
    else:
        live = bond_clusters()
        of_group = cluster_of_bond()
        field = "bond"
    of_bond = of_group

    present = {l.get(field) for l in lines if lo <= l["date"] <= hi and l.get(field)}
    groups = [g for c in (1, 2, 3) for g in live.get(c, [])
              if (cluster in (None, 0) or of_group.get(g) == cluster)
              and (not by_warehouse or g in present)]
    bonds = groups

    exact: dict[tuple[str, date], float] = defaultdict(float)
    seen_days: set = set()
    for l in lines:
        key = l.get(field) or ""
        if lo <= l["date"] <= hi and key in of_group:
            exact[(key, l["date"])] += l["cases"]
            seen_days.add(l["date"])

    # Every day in the window gets a column, including days nothing moved: a
    # dry Sunday is information. What the report must not do is let a day with
    # no raw behind it read the same as a day that genuinely sold nothing, so
    # the days an upload actually answers for travel with the figures and the
    # screen says which zeros are which.
    days = []
    d = lo
    while d <= hi:
        days.append(d)
        d += timedelta(days=1)

    def row(label: str, members: list[str], kind_: str, cl: int | None = None) -> dict:
        cells = [half_up(sum(exact[(b, d)] for b in members)) for d in days]
        return {"kind": kind_, "label": label, "cluster": cl, "cells": cells,
                "total": half_up(sum(exact[(b, d)] for b in members for d in days))}

    rows: list[dict] = []
    for c in (1, 2, 3):
        members = [b for b in live.get(c, []) if b in bonds]
        if not members:
            continue
        rows += [row(b, [b], "bond", c) for b in members]
        rows.append(row(f"CLUSTER {c}", members, "cluster", c))
    rows.append(row("Grand Total", bonds, "grand"))

    return {
        "kind": kind,
        "view": view,
        "group_label": "Warehouse" if by_warehouse else "Bond",
        "title": DAILY_KINDS[kind]["title"],
        "unit": DAILY_KINDS[kind]["unit"],
        "days": [{"iso": d.isoformat(), "dom": str(d.day),
                  "dow": d.strftime("%a").upper()} for d in days],
        "rows": rows,
        "period": {"from": lo.isoformat(), "to": hi.isoformat(),
                   "label": f"{lo.day} {lo.strftime('%B %Y')} - {hi.day} {hi.strftime('%B %Y')}"},
        "span": {"min": span_lo.isoformat(), "max": span_hi.isoformat()},
        # Which days actually have rows behind them, so the picker can mark
        # them. Not the same as the span: a stream can have gaps in it.
        "have": sorted({l["date"].isoformat() for l in lines}),
        # And which days an upload answers for at all. A zero inside this set
        # is a day nothing moved; a zero outside it is a day not uploaded yet.
        "covered": sorted(d.isoformat() for d in covered),
        "source": ("day files in KSBC shop sales" if book is not None and book.is_dir()
                   else book.name if book is not None
                   else "uploaded raw files"),
        "unmapped": sorted({l.get("shop_code", "") for l in lines
                            if l["bond"] not in of_bond and l.get("shop_code")}),
        "source_block": _daily_source(kind, sources if kind == "secondary" else [],
                                      book, lo, hi),
    }

# ---------------------------------------------------------------------------
# Shop sales - cumulative periods
# ---------------------------------------------------------------------------
#
# Every cumulative raw the office has ever uploaded stays in this folder, so a
# period from six months ago still opens. The filename carries the window, and
# the window is what both cumulative reports are keyed on.

_CUM_FILE_RE = re.compile(
    r"^(?P<year>\d{4})-(?P<mm>\d{2}) [A-Z]+ (?P<start>\d{1,2})-(?P<end>\d{1,2}) CUMULATIVE",
    re.IGNORECASE,
)


def cumulative_folder() -> Path:
    return config.CLAUDE_ROOT / "KSBC shop sales" / "_cumulative"


def cumulative_periods() -> list[dict]:
    """Stored cumulative periods, newest first."""
    folder = cumulative_folder()
    out: list[dict] = []
    if not folder.is_dir():
        return out
    for path in sorted(folder.glob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        m = _CUM_FILE_RE.match(path.name)
        if not m:
            continue
        year, month = int(m.group("year")), int(m.group("mm"))
        start_d, end_d = int(m.group("start")), int(m.group("end"))
        try:
            start = date(year, month, start_d)
            end = date(year, month, end_d)
        except ValueError:
            continue
        out.append({
            "key": f"{start.isoformat()}..{end.isoformat()}",
            "start": start, "end": end, "path": path,
            "days": max(1, (end - start).days),
            "long": f"{start.day} {start.strftime('%B')} {year} - {end.day} {end.strftime('%B')} {year}",
            "short": f"{start.day} {start.strftime('%b')} {year} to {end.day} {end.strftime('%b')} {year}",
        })
    out.sort(key=lambda p: (p["start"], p["end"]), reverse=True)
    return out


def cumulative_period(key: str) -> dict | None:
    periods = cumulative_periods()
    if not periods:
        return None
    for p in periods:
        if p["key"] == key:
            return p
    return periods[0]


def previous_cumulative(period: dict) -> dict | None:
    """The same day-window one month back, when the office has uploaded it.

    Matched on the day numbers rather than on elapsed days: 'the first half of
    last month' is the comparison the trade actually makes, and a 1-16 window
    should line up with 1-16, not with a rolling 15 days.
    """
    start = period["start"]
    month = start.month - 1 or 12
    year = start.year - (1 if start.month == 1 else 0)
    for p in cumulative_periods():
        if (p["start"].year, p["start"].month) != (year, month):
            continue
        if p["start"].day == start.day and p["end"].day == period["end"].day:
            return p
    return None

_CUM_MEASURES = [("Shop Opening Cases", "Shop Opening Bottles"),
                 ("Shop In Cases", "Shop In Bottles"),
                 ("Shop Out Cases", "Shop Out Bottles"),
                 ("Shop Closing Cases", "Shop Closing Bottles")]


def cumulative_lines(path: Path) -> dict:
    """(shop|brand|pack) -> [opening, receipt, sales, closing] from one cumulative raw.

    Brand and pack granularity, because the per-bond PDF prints at that level
    and every coarser figure is just this summed. Cached on mtime: the file
    never changes once uploaded, so the second read costs nothing.
    """
    key = f"cumline:{path}:{path.stat().st_mtime_ns}:{_master_stamp()}"
    with _CACHE_LOCK:
        hit = _CACHE.get(key)
    if hit:
        return hit[1]
    got = _parsed("cumline", path, lambda: _cumulative_lines_read(path))
    with _CACHE_LOCK:
        _CACHE[key] = (0.0, got)
    return got


def _cumulative_lines_read(path: Path) -> dict:
    """The workbook itself, read start to finish. Cached by its caller."""
    master = load_master()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    header = [str(h).strip() if h else "" for h in next(rows)]
    ix = {h: i for i, h in enumerate(header)}
    need = ["Shop Code", "Shop Name", "Brand Name", "Packing", "Bottle Per Case"]
    absent = [c for c in need + [c for pr in _CUM_MEASURES for c in pr] if c not in ix]
    if absent:
        wb.close()
        raise KeyError(absent[0])

    out: dict = {"lines": {}, "names": {}, "warehouses": {}}
    for r in rows:
        if r is None or r[ix["Shop Code"]] is None:
            continue
        code = str(r[ix["Shop Code"]]).strip()
        if not master.get(code, {}).get("bond"):
            continue
        out["names"][code] = str(r[ix["Shop Name"]] or code).strip().upper()
        # The warehouse a shop draws from is the raw file's own column, not
        # ours: bond mapping is a decision we make, warehouse is a fact KSBC
        # reports, and it can move between periods.
        if "Warehouse Name" in ix:
            wh = canonical_warehouse(r[ix["Warehouse Name"]])
            if wh:
                out["warehouses"][code] = wh
        k3 = f'{code}|{str(r[ix["Brand Name"]] or "").strip().upper()}|' \
             f'{str(r[ix["Packing"]] or "").strip().upper()}'
        cell = out["lines"].setdefault(k3, [0.0, 0.0, 0.0, 0.0])
        bpc = r[ix["Bottle Per Case"]] or 0
        for k, (c_col, b_col) in enumerate(_CUM_MEASURES):
            cell[k] += float(r[ix[c_col]] or 0) + (
                float(r[ix[b_col]] or 0) / bpc if bpc else 0.0)
    wb.close()
    return out


def _shops_from_lines(lines: dict, names: dict, warehouses: dict | None = None) -> list:
    master = load_master()
    shops: dict = {}
    for k, v in lines.items():
        code = k.split("|", 1)[0]
        shop = shops.get(code)
        if shop is None:
            shop = shops[code] = {
                "code": code, "name": names.get(code, code),
                "bond": str(master[code]["bond"]).strip().upper(),
                # The raw file's warehouse when it has one, master's only as a
                # fallback for a shop the export did not name.
                "warehouse": (warehouses or {}).get(code)
                             or master[code].get("warehouse", ""),
                "v": [0.0, 0.0, 0.0, 0.0]}
        for i in range(4):
            shop["v"][i] += v[i]
    rank = {code: i for i, code in enumerate(master)}
    return sorted(shops.values(), key=lambda s: (rank.get(s["code"], 10_000), s["name"]))


def shop_cumulative(period: dict, cluster: int | None = None, bond: str = "",
                    warehouse: str = "", group_by: str = "bond") -> dict:
    """Bond -> its shops, for one stored cumulative period."""
    try:
        got = cumulative_lines(period["path"])
    except KeyError as exc:
        return {"error": f"That cumulative file has no '{exc.args[0]}' column."}
    slim = {k: period[k] for k in ("key", "short", "long", "days")}
    return _group_shops(_shops_from_lines(got["lines"], got["names"],
                                         got.get("warehouses")), slim,
                        cluster, bond, warehouse, group_by)


def shop_day_files() -> dict:
    """date -> the daily shop-sales raw that covers it."""
    folder = config.CLAUDE_ROOT / "KSBC shop sales"
    out: dict = {}
    if not folder.is_dir():
        return out
    for path in folder.glob("*.xlsx"):
        m = _RAW_DAY_RE.match(path.name)
        if not m:
            continue
        month = _MONTH_NUM.get(m.group(1).upper())
        if not month:
            continue
        year = datetime.fromtimestamp(path.stat().st_mtime).year
        try:
            out[date(year, month, int(m.group(2)))] = path
        except ValueError:
            continue
    return out


def _day_lines(path: Path, master: dict) -> dict:
    """The same shape as cumulative_lines, for one day's export."""
    key = f"dayline:{path}:{path.stat().st_mtime_ns}:{_master_stamp()}"
    with _CACHE_LOCK:
        hit = _CACHE.get(key)
    if hit:
        return hit[1]
    got = _parsed("dayline", path, lambda: _day_lines_read(path, master))
    with _CACHE_LOCK:
        _CACHE[key] = (0.0, got)
    return got


def _day_lines_read(path: Path, master: dict) -> dict:
    """The day's export, read start to finish. Cached by its caller."""
    out: dict = {"lines": {}, "names": {}, "warehouses": {}}
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except (OSError, ValueError):
        return out
    need = ["Shop Code", "Shop Name", "Brand Name", "Packing", "Bottle Per Case"]
    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        try:
            header = [str(h).strip() if h else "" for h in next(rows)]
        except StopIteration:
            continue
        ix = {h: i for i, h in enumerate(header)}
        if any(c not in ix for c in need + [c for pr in _CUM_MEASURES for c in pr]):
            continue
        for r in rows:
            if r is None or r[ix["Shop Code"]] is None:
                continue
            code = str(r[ix["Shop Code"]]).strip()
            if not master.get(code, {}).get("bond"):
                continue
            out["names"][code] = str(r[ix["Shop Name"]] or code).strip().upper()
            if "Warehouse Name" in ix:
                wh = canonical_warehouse(r[ix["Warehouse Name"]])
                if wh:
                    out["warehouses"][code] = wh
            k3 = f'{code}|{str(r[ix["Brand Name"]] or "").strip().upper()}|' \
                 f'{str(r[ix["Packing"]] or "").strip().upper()}'
            cell = out["lines"].setdefault(k3, [0.0, 0.0, 0.0, 0.0])
            bpc = r[ix["Bottle Per Case"]] or 0
            for k, (c_col, b_col) in enumerate(_CUM_MEASURES):
                cell[k] += float(r[ix[c_col]] or 0) + (
                    float(r[ix[b_col]] or 0) / bpc if bpc else 0.0)
    wb.close()
    return out


def segments() -> list:
    """Every window this app can read, as (start, end, kind, path).

    The KSBC portal caps one pull at sixteen days, so the office's files always
    start on the 1st or the 17th: 1-1, 1-2 ... 1-16, then 17-17, 17-18 ... to
    month end. A window longer than that is therefore never one file - it is
    the first-half file plus the second-half one - which is what the planner
    below exists to assemble.
    """
    segs = [(p["start"], p["end"], "cumulative", p["path"]) for p in cumulative_periods()]
    segs += [(d, d, "daily", path) for d, path in shop_day_files().items()]
    return segs


def plan_window(start: date, end: date) -> dict:
    """A chain of stored files that tiles [start, end] exactly, or the gap.

    Longest segment first at each step: fewest reads, and a window the office
    pulled whole is answered by that very file rather than reassembled from a
    fortnight of days.
    """
    by_start: dict = {}
    for s, e, kind, path in segments():
        by_start.setdefault(s, []).append((e, kind, path))

    chain, cursor = [], start
    while cursor <= end:
        options = [o for o in by_start.get(cursor, []) if o[0] <= end]
        if not options:
            return {"error": _gap_message(cursor, end, by_start)}
        e, kind, path = max(options)
        chain.append({"from": cursor, "to": e, "kind": kind, "path": path})
        cursor = e + timedelta(days=1)
    return {"chain": chain}


def _gap_message(cursor: date, end: date, by_start: dict) -> str:
    """Name the file that is missing, in the terms the office pulls them in."""
    overshoot = [e for e, _, _ in by_start.get(cursor, [])]
    if overshoot:
        best = max(overshoot)
        return (f"The only upload starting {cursor.strftime('%d %b').lstrip('0')} runs to "
                f"{best.strftime('%d %b').lstrip('0')}, past the "
                f"{end.strftime('%d %b').lstrip('0')} you asked for. Pull "
                f"{cursor.day}-{end.day} {end.strftime('%B')} from KSBC and upload it, or "
                f"pick {cursor.strftime('%d %b').lstrip('0')} to "
                f"{best.strftime('%d %b').lstrip('0')}.")
    return (f"Nothing uploaded covers {cursor.strftime('%d %b %Y').lstrip('0')}. A KSBC pull "
            f"is capped at 16 days, so a longer window is the first-half file plus the "
            f"second-half one - upload the file starting {cursor.day} "
            f"{cursor.strftime('%B')}, or the daily exports for those days.")


def window_lines(start: date, end: date) -> dict:
    """The whole window, stitched from whatever files cover it.

    Opening comes from the first file a line appears in and closing from the
    last - never a sum, or a shop that held stock all month would read as
    though it opened once per file. Receipts and sales do sum.
    """
    plan = plan_window(start, end)
    if "error" in plan:
        return plan

    master = load_master()
    lines: dict = {}
    names: dict = {}
    warehouses: dict = {}
    opened: set = set()
    for seg in plan["chain"]:
        got = (cumulative_lines(seg["path"]) if seg["kind"] == "cumulative"
               else _day_lines(seg["path"], master))
        names.update(got["names"])
        # Later files win: a shop moved to another warehouse should read as
        # where it draws from now, not where it drew from a fortnight ago.
        warehouses.update(got.get("warehouses", {}))
        for k, v in got["lines"].items():
            cell = lines.setdefault(k, [0.0, 0.0, 0.0, 0.0])
            if k not in opened:
                cell[0] = v[0]
                opened.add(k)
            cell[1] += v[1]
            cell[2] += v[2]
            cell[3] = v[3]
    return {"lines": lines, "names": names, "warehouses": warehouses,
            "chain": [{"from": seg["from"].isoformat(), "to": seg["to"].isoformat(),
                       "kind": seg["kind"], "name": seg["path"].name}
                      for seg in plan["chain"]]}


def shop_range_lines(start: date, end: date) -> dict:
    """window_lines, as the JSON feed the PDF builder takes."""
    return window_lines(start, end)


def shop_detail(start: date, end: date, code: str) -> dict:
    """One shop's brand and pack lines for a window.

    Fetched a shop at a time rather than shipped with the page: the full book
    is five and a half thousand lines, and nobody opens all of them.
    """
    got = window_lines(start, end)
    if "error" in got:
        return got
    brands: dict = {}
    for k, v in got["lines"].items():
        shop, brand, pack = k.split("|", 2)
        if shop != code:
            continue
        brands.setdefault(brand, {})[pack] = [round(x, 2) for x in v]
    if not brands:
        return {"brands": [], "name": got["names"].get(code, code)}

    out = []
    for brand in sorted(brands):
        packs = brands[brand]
        sub = [0.0, 0.0, 0.0, 0.0]
        for pack in packs:
            for i in range(4):
                sub[i] += packs[pack][i]
        out.append({
            "brand": brand,
            "values": [round(x, 2) for x in sub],
            # Packs sort on the label, the way the office's own books read:
            # 1000 ML files before 500 ML.
            "packs": [{"pack": pk, "values": packs[pk]} for pk in sorted(packs)],
        })
    return {"brands": out, "name": got["names"].get(code, code)}


def _group_shops(shops: list, period: dict, cluster: int | None, bond: str,
                 warehouse: str = "", group_by: str = "bond") -> dict:
    """Group the shops one way or the other - never both at once.

    Bond is our own mapping, set in Settings; warehouse is the raw file's own
    column. They cut the same shops differently, so the screen picks a side and
    the headings, the filter and the totals all follow it.

    Cluster is not a third grouping - it is a filter on the shops, and it
    applies whichever side is showing. Every shop belongs to a bond and every
    bond to a cluster, so 'Cluster 1, by warehouse' is a perfectly good
    question: the warehouses Cluster 1's shops draw from.
    """
    by_warehouse = group_by == "warehouse"
    of_cluster = cluster_of_bond()
    want_bond = (bond or "").strip().upper()
    want_wh = (warehouse or "").strip().upper()

    grouped: dict = {}
    bonds_seen: set = set()
    warehouses_seen: set = set()

    for shop in shops:
        b = shop["bond"]
        wh = (shop.get("warehouse") or "").upper()
        if b:
            bonds_seen.add(b)
        if wh:
            warehouses_seen.add(wh)

        if cluster and of_cluster.get(b) != cluster:
            continue

        if by_warehouse:
            if want_wh and wh != want_wh:
                continue
            key, tag = (wh or "UNMAPPED"), 0
        else:
            if want_bond and b != want_bond:
                continue
            key, tag = b, of_cluster.get(b, 0)

        g = grouped.get(key)
        if g is None:
            g = grouped[key] = {"bond": key, "cluster": tag,
                                "shops": [], "totals": [0.0, 0.0, 0.0, 0.0],
                                "clusters": set()}
        if by_warehouse:
            g["clusters"].add(of_cluster.get(b, 0))
        g["shops"].append({"code": shop["code"], "name": shop["name"],
                           "warehouse": shop.get("warehouse", ""),
                           "values": [round(v, 2) for v in shop["v"]]})
        for k in range(4):
            g["totals"][k] += shop["v"][k]

    groups, grand = [], [0.0, 0.0, 0.0, 0.0]
    for key in sorted(grouped):
        g = grouped[key]
        # A warehouse carries a cluster chip only when all of its shops sit in
        # one - which is always true once a cluster is picked, and true of many
        # warehouses anyway. One that straddles two gets no chip rather than a
        # misleading one.
        seen = g.pop("clusters", set())
        if by_warehouse:
            g["cluster"] = seen.pop() if len(seen) == 1 else 0
        for k in range(4):
            grand[k] += g["totals"][k]
        g["totals"] = [round(v, 2) for v in g["totals"]]
        groups.append(g)

    return {"period": period, "bonds": groups, "total": [round(v, 2) for v in grand],
            "shop_count": sum(len(g["shops"]) for g in groups),
            "group_by": "warehouse" if by_warehouse else "bond",
            "all_bonds": sorted(b for b in bonds_seen if b),
            "warehouses": sorted(w for w in warehouses_seen if w)}


def period_for(start: date, end: date) -> dict:
    return {
        "key": f"{start.isoformat()}..{end.isoformat()}",
        "start": start, "end": end, "days": max(1, (end - start).days),
        "long": f"{start.day} {start.strftime('%B')} {start.year} - "
                f"{end.day} {end.strftime('%B')} {end.year}",
        "short": f"{start.day} {start.strftime('%b')} {start.year} to "
                 f"{end.day} {end.strftime('%b')} {end.year}",
    }


def widest_window():
    """The window to open on: this month, as wide as the uploads can answer.

    Not simply the newest file: with 1-16 and 17-23 on disk the useful default
    is 1-23, because that is the whole of what has been uploaded. Searched
    rather than taken greedily - the longest file starting on a day is not
    always the one that reaches furthest, if a shorter one is the only link to
    the next stretch.

    It never crosses a month. Daily files tile one day at a time, so an
    unbroken run of them chained 1 August straight through to 18 September and
    opened every report on a window nobody asked for. A month is the unit the
    office reports in, so that is the unit a default is cut to - and the month
    it picks is this one, falling back to the newest month with anything in
    it. Asking for a window by hand is untouched: plan_window() still tiles
    across months.
    """
    per_month: dict = {}
    for s, e, _kind, _path in segments():
        per_month.setdefault((s.year, s.month), {}).setdefault(s, []).append(e)
    if not per_month:
        return None

    today = date.today()
    key = (today.year, today.month)
    if key not in per_month:
        key = max(per_month)
    by_start = per_month[key]

    memo: dict = {}

    def reach(day):
        if day in memo:
            return memo[day]
        memo[day] = day          # guards a cycle; segments cannot overlap back
        best = None
        for e in by_start.get(day, []):
            nxt = e + timedelta(days=1)
            far = reach(nxt) if nxt in by_start else e
            if best is None or far > best:
                best = far
        memo[day] = best if best is not None else day
        return memo[day]

    # Furthest reach first, then the earliest start that gets there: 1-16 plus
    # 17-23 opens as 1-23 rather than 17-23.
    winner = None
    for start in sorted(by_start):
        end = reach(start)
        if winner is None:
            winner = (start, end)
            continue
        better = (end, (end - start).days)
        against = (winner[1], (winner[1] - winner[0]).days)
        if better > against:
            winner = (start, end)
    return winner


def resolve_window(start: date, end: date) -> dict:
    """Everything a screen needs for one window: the figures and their provenance.

    An exact single-file match is still the fast path, but the general answer
    is the chain - 1-16 plus 17-24 for a window to the 24th, because that is
    how KSBC lets the office pull it.
    """
    for p in cumulative_periods():
        if p["start"] == start and p["end"] == end:
            return {"period": p, "source": "cumulative", "chain":
                    [{"from": start.isoformat(), "to": end.isoformat(),
                      "kind": "cumulative", "name": p["path"].name}]}

    built = window_lines(start, end)
    if "error" in built:
        return {"error": built["error"]}

    period = period_for(start, end)
    period["path"] = None
    chain = built["chain"]
    source = "stitched" if len(chain) > 1 else chain[0]["kind"]
    return {"period": period, "source": source, "chain": chain,
            "shops": _shops_from_lines(built["lines"], built["names"])}

# ---------------------------------------------------------------------------
# Liquidation: what actually left the trade, this month against last
# ---------------------------------------------------------------------------
#
# Three streams, and only two of them are liquidation:
#
#   SHOP LIQUIDATION   cases sold out of KSBC shops - the shop-sales export.
#   SECONDARY SALES    every dispatch line, whoever it went to. Stock moving
#                      down the chain rather than leaving it, so it is context
#                      and is deliberately left out of the total.
#   FED / BAR INVOICE  the part of that dispatch invoiced straight to Consumer
#                      Fed outlets and bars. Nobody sells these on, so the
#                      invoice IS the liquidation - which is why it is broken
#                      out of secondary rather than sitting only inside it.
#
# Hence TOTAL = SHOP + FED/BAR. Adding all of secondary would double-count
# every case a shop both received and sold.


def _same_window_last_month(start: date, end: date):
    month = start.month - 1 or 12
    year = start.year - (1 if start.month == 1 else 0)
    try:
        return date(year, month, start.day), date(year, month, end.day)
    except ValueError:
        return None, None


def _shop_sales_by_bond(start: date, end: date) -> dict:
    got = window_lines(start, end)
    if "error" in got:
        return {}
    master = load_master()
    out: dict = defaultdict(float)
    for k, v in got["lines"].items():
        bond = str(master.get(k.split("|", 1)[0], {}).get("bond", "")).strip().upper()
        if bond:
            out[bond] += v[2]
    return out


def _dispatch_by_bond(start: date, end: date) -> tuple[dict, dict, int]:
    """(all dispatch, the Fed/Bar part of it) per bond, plus the days it ran on.

    Fed and Bar are a SUBSET of secondary, not a sibling of it - the office's
    own scorecard shows ATTINGAL at 376 secondary with 10 of that Fed/Bar, and
    both figures reconcile to this workbook line for line.
    """
    every: dict = defaultdict(float)
    fedbar: dict = defaultdict(float)
    days: set = set()
    lines, _sources = secondary_lines()
    for line in lines:
        day = line.get("date")
        if not day or not (start <= day <= end):
            continue
        bond = str(line.get("bond") or "").strip().upper()
        if not bond:
            continue
        every[bond] += line["cases"]
        days.add(day)
        if line.get("cat") in ("FED", "BAR"):
            fedbar[bond] += line["cases"]
    return every, fedbar, len(days)


def _liq_source(start: date, end: date, p_start, p_end) -> dict:
    """Both windows, both legs. A scorecard reads four sets of files at once.

    Shop and dispatch are different streams stored in different places, and
    each side of the comparison is answered by its own files - which is worth
    being able to see when the two sides are not the same length.
    """
    _l, sec = secondary_lines()

    def shop(a, b):
        plan = plan_window(a, b)
        return src_windows(plan.get("chain", []))

    legs = [src_leg("Shop sales (KSBC)", shop(start, end),
                    f"{period_for(start, end)['short']}")]
    if p_start:
        was = shop(p_start, p_end)
        if was:
            legs.append(src_leg("Compared against",
                                was, period_for(p_start, p_end)["short"], "green"))
    legs.append(src_secondary(sec))
    return src_block(
        legs,
        "Shop liquidation comes off the KSBC cumulative files; secondary and "
        "Fed/Bar off the dispatch raws. Total liquidation is shop plus Fed/Bar, "
        "never shop plus secondary.")


def liquidation(start: date, end: date, prev: tuple | None = None) -> dict:
    """The liquidation scorecard: this window against another.

    The comparison defaults to the same day-range a month back, which is what
    the trade asks for nine times in ten. `prev` overrides it with any window
    at all - 1-16 August against 1-20 October - so the two sides no longer
    have to be the same length, and the per-day averages divide each side by
    its own span rather than by the current one twice.
    """
    p_start, p_end = prev if prev else _same_window_last_month(start, end)
    span = max(1, (end - start).days)
    p_span = max(1, (p_end - p_start).days) if p_start else 1

    shop_now = _shop_sales_by_bond(start, end)
    shop_was = _shop_sales_by_bond(p_start, p_end) if p_start else {}
    sec_now, fed_now, disp_days_now = _dispatch_by_bond(start, end)
    if p_start:
        sec_was, fed_was, disp_days_was = _dispatch_by_bond(p_start, p_end)
    else:
        sec_was, fed_was, disp_days_was = {}, {}, 0

    if not shop_now and not sec_now and not fed_now:
        return {"error": "Nothing uploaded covers that window yet."}

    clusters = bond_clusters()
    of_cluster = cluster_of_bond()
    known = {b for bs in clusters.values() for b in bs}
    seen = set(shop_now) | set(shop_was) | set(sec_now) | set(sec_was) | set(fed_now) | set(fed_was)

    def block(bonds, src_now, src_was):
        return [sum(src_now.get(b, 0.0) for b in bonds), sum(src_was.get(b, 0.0) for b in bonds)]

    def row(label, bonds, kind):
        shop = block(bonds, shop_now, shop_was)
        sec = block(bonds, sec_now, sec_was)
        fed = block(bonds, fed_now, fed_was)
        total = [shop[0] + fed[0], shop[1] + fed[1]]
        return {"label": label, "kind": kind,
                "blocks": [shop, sec, fed, total]}

    rows = []
    for cid in sorted(clusters):
        members = [b for b in clusters[cid] if b in seen]
        if not members:
            continue
        for bond in sorted(members):
            rows.append(row(bond, [bond], "bond"))
        rows.append(row(f"CLUSTER - {cid}", members, "cluster"))
    loose = sorted(b for b in seen if b and b not in known)
    for bond in loose:
        rows.append(row(bond, [bond], "bond"))

    everything = sorted(seen)
    rows.append(row("TOTAL", everything, "total"))

    # Averages: the shop and total columns divide by the days the window spans,
    # the dispatch columns by the days a dispatch actually went out - a
    # warehouse that shut for three days did not average across them.
    grand = rows[-1]["blocks"]
    # Both dispatch columns divide by the days a dispatch actually went out -
    # the office's August column reads 421 and 98 against 4,632 and 1,079,
    # which is 11 days for each, not one count per column.
    divisors = [(span, p_span), (disp_days_now or 1, disp_days_was or 1),
                (disp_days_now or 1, disp_days_was or 1), (span, p_span)]
    rows.append({
        "label": "AVERAGE DAILY SALE", "kind": "average",
        "blocks": [[grand[i][0] / divisors[i][0], grand[i][1] / divisors[i][1]]
                   for i in range(4)],
    })

    return {
        "rows": rows,
        "period": period_for(start, end),
        "previous": period_for(p_start, p_end) if p_start else None,
        "days": {"span": span, "prev_span": p_span,
                 "dispatch": [disp_days_now, disp_days_was]},
        "source_block": _liq_source(start, end, p_start, p_end),
    }

# ---------------------------------------------------------------------------
# Status calendar: what has been uploaded, and for which day
# ---------------------------------------------------------------------------
#
# Read off the files themselves rather than off the job log. A job record says
# somebody pressed upload; a file on disk says the reports can actually answer
# for that day, which is the question this screen exists to settle.


def _secondary_days() -> set:
    lines, _sources = secondary_lines()
    return {l["date"] for l in lines if l.get("date")}


def _warehouse_days() -> set:
    """The stock snapshots, from the history the build appends to."""
    out: set = set()
    hist = config.CLAUDE_ROOT / "Warehouse stock" / "_history" / "stock_history.csv"
    if hist.is_file():
        try:
            with hist.open(newline="", encoding="utf-8-sig") as fh:
                reader = csv.DictReader(fh)
                for row in reader:
                    for key in ("Report Date", "report_date", "Date", "date", "As On"):
                        if key in row:
                            d = _parse_date(row[key])
                            if d:
                                out.add(d)
                            break
        except (OSError, csv.Error):
            pass
    # Anything still sitting in the input folder counts too - it has not been
    # consumed by a build yet, but the day is covered.
    folder = config.CLAUDE_ROOT / "Warehouse stock"
    if folder.is_dir():
        for path in folder.glob("Report*"):
            m = re.match(r"^Report[ _](\d{1,2})-([A-Za-z]{3})-(\d{4})", path.name, re.I)
            if not m:
                continue
            abbr = m.group(2).upper()
            month = next((v for k, v in _MONTH_NUM.items() if k.startswith(abbr)), 0)
            if month:
                try:
                    out.add(date(int(m.group(3)), month, int(m.group(1))))
                except ValueError:
                    pass
    return out


def upload_calendar() -> dict:
    """day -> which streams can answer for it, plus the periods behind them."""
    cover: dict = defaultdict(set)

    for d in shop_day_files():
        cover[d].add("shop_daily")
    for d in _secondary_days():
        cover[d].add("secondary")
    for d in _warehouse_days():
        cover[d].add("warehouse")

    periods = []
    for p in cumulative_periods():
        periods.append({"from": p["start"].isoformat(), "to": p["end"].isoformat(),
                        "short": p["short"], "name": p["path"].name})
        day = p["start"]
        while day <= p["end"]:
            cover[day].add("shop_cumulative")
            day += timedelta(days=1)

    days = {d.isoformat(): sorted(v) for d, v in cover.items() if v}
    return {
        "days": days,
        "periods": periods,
        "streams": [
            {"key": "shop_cumulative", "label": "Shop Sales - Cumulative", "colour": "#2563EB"},
            {"key": "shop_daily", "label": "Shop Sales - Daily", "colour": "#0EA5E9"},
            {"key": "secondary", "label": "Secondary Sales", "colour": "#10B981"},
            {"key": "warehouse", "label": "Warehouse Stock", "colour": "#F59E0B"},
        ],
        "span": {"first": min(days) if days else "", "last": max(days) if days else ""},
    }


# ---------------------------------------------------------------------------
# Target vs achievement
# ---------------------------------------------------------------------------
#
# ACHIEVEMENT IS TOTAL LIQUIDATION, AND IT HAS TWO LEGS
# ----------------------------------------------------
# A case leaves this company through one of two doors, and the target is set
# against both:
#
#   * the tertiary leg - what KSBC shops actually sold (Shop Out on the
#     cumulative raws, cases plus loose bottles over the pack size);
#   * the invoice leg  - what was invoiced straight to FED and BAR outlets,
#     which never passes through a KSBC shop and so appears nowhere in the
#     first leg.
#
# KSBC-category dispatch rows are deliberately NOT counted. Those are stock
# moving into the shops whose sales the tertiary leg already counts, so adding
# them would count the same case twice - in August that would have inflated
# the network by 4,213 cases against a true 4,989.
#
# Every figure is carried unrounded to the very last step. The August 2026
# workbook rounded each brand cell first and then added the rounded cells in
# some places and the true total in others, which left nine cluster cells
# that did not equal the column added up by hand. Round once, for display.


def _tva_source(sources: dict) -> dict:
    """Achievement is two legs added together, so its source is two legs.

    What the shops sold comes off the KSBC cumulative files; what went
    straight to Fed and Bar outlets never passes through a shop and comes off
    the dispatch raws. Targets are typed in rather than uploaded, so they are
    named but carry no file.
    """
    legs = [src_leg("Shop sales (KSBC)", src_windows(sources.get("chain", [])),
                    "what the shops sold")]
    invoice = sources.get("invoice") or []
    if invoice:
        legs.append(src_secondary(invoice))
    return src_block(
        legs,
        "Achievement is shop sales plus Fed/Bar invoice - KSBC-category dispatch "
        "is left out, since that stock is counted where the shop sold it. "
        "Targets are entered in the app, not uploaded.")


def _tva_legs(start: date, end: date) -> dict:
    """Achievement per (bond, family), and what each leg contributed."""
    master = load_master()
    from . import targets as targets_mod

    grid: dict = defaultdict(lambda: defaultdict(float))
    legs = {"tertiary": 0.0, "fed": 0.0, "bar": 0.0}
    sources: dict = {"tertiary": [], "invoice": []}

    win = window_lines(start, end)
    if "error" in win:
        return {"error": win["error"]}
    sources["tertiary"] = [seg["name"] for seg in win.get("chain", [])]
    sources["chain"] = win.get("chain", [])
    for key, cell in win["lines"].items():
        sold = cell[2]
        if not sold:
            continue
        code, brand, _pack = key.split("|", 2)
        bond = master.get(code, {}).get("bond", "")
        if not bond:
            continue
        grid[bond][targets_mod.family_of(brand)] += sold
        legs["tertiary"] += sold

    lines, sec_sources = secondary_lines()
    used = False
    for line in lines:
        day = line.get("date")
        if day is None or not (start <= day <= end):
            continue
        cat = (line.get("cat") or "").upper()
        if cat not in ("FED", "BAR"):
            continue
        bond = line.get("bond") or ""
        if not bond:
            continue
        grid[bond][targets_mod.family_of(line.get("brand", ""))] += line["cases"]
        legs["fed" if cat == "FED" else "bar"] += line["cases"]
        used = True
    if used:
        sources["invoice"] = sec_sources

    return {"grid": grid, "legs": legs, "sources": sources}


def _tva_row(label: str, kind: str, cluster: int | None,
             bonds: list, grid: dict, tgt: dict, cols: list) -> dict:
    """One printed line. Every cell is the true sum, rounded only here."""
    ach_true = {k: sum(grid.get(b, {}).get(k, 0.0) for b in bonds) for k in cols}
    tgt_true = {k: sum(tgt.get(b, {}).get(k, 0.0) for b in bonds) for k in cols}
    ach_total = sum(ach_true.values())
    tgt_total = sum(tgt_true.values())
    return {
        "kind": kind,
        "label": label,
        "cluster": cluster,
        "bond": bonds[0] if kind == "bond" else "",
        "ach": {k: round(v) for k, v in ach_true.items()},
        "tgt": {k: round(v) for k, v in tgt_true.items()},
        "ach_total": round(ach_total),
        "tgt_total": round(tgt_total),
        "ach_exact": round(ach_total, 3),
        "pct": round(ach_total / tgt_total * 100, 2) if tgt_total else None,
    }


def target_vs_achievement(start: date, end: date, cluster: int | None = None,
                          month: str = "") -> dict:
    """The bond x brand grid, in target-and-achieved pairs, by cluster."""
    from . import targets as targets_mod

    got = _tva_legs(start, end)
    if "error" in got:
        return got
    grid = got["grid"]

    month = month or start.strftime("%Y-%m")
    tgt = targets_mod.load(month)

    # Other earns a column only by selling something. A dead brand that comes
    # back shows up here rather than quietly leaving the total short.
    other_sold = any(cells.get(targets_mod.OTHER, 0.0) for cells in grid.values())
    cols = targets_mod.FAMILY_KEYS + ([targets_mod.OTHER] if other_sold else [])

    clusters = bond_clusters()
    known = {b for members in clusters.values() for b in members}
    # A bond that sold or was targeted but sits in no cluster still has to
    # appear, or the grand total stops being the network.
    loose = sorted((set(grid) | set(tgt)) - known)

    rows: list = []
    shown: list = []
    for cid in sorted(clusters):
        if cluster in (1, 2, 3) and cid != cluster:
            continue
        members = sorted(clusters[cid])
        if not members:
            continue
        for bond in members:
            rows.append(_tva_row(bond, "bond", cid, [bond], grid, tgt, cols))
        rows.append(_tva_row(f"CLUSTER {cid}", "cluster", cid, members, grid, tgt, cols))
        shown += members

    if cluster in (None, 0):
        for bond in loose:
            rows.append(_tva_row(bond, "bond", None, [bond], grid, tgt, cols))
        shown += loose
        rows.append(_tva_row("GRAND TOTAL", "grand", None, shown, grid, tgt, cols))

    legs = got["legs"]
    return {
        "period": {"from": start.isoformat(), "to": end.isoformat(),
                   "short": f"{start.day} {start.strftime('%b')} {start.year} to "
                            f"{end.day} {end.strftime('%b')} {end.year}"},
        "month": month,
        "month_label": f"{_MONTHS_UP[int(month[5:7]) - 1].title()} {month[:4]}",
        "columns": [{"key": k, "label": targets_mod.FAMILY_LABEL[k]} for k in cols],
        "rows": rows,
        "legs": {k: round(v, 3) for k, v in legs.items()},
        "liquidation": round(sum(legs.values()), 3),
        "sources": got["sources"],
        "source_block": _tva_source(got["sources"]),
        "targets_set": bool(tgt),
        "target_meta": targets_mod.meta(month),
        "cluster": cluster or 0,
    }


# ---------------------------------------------------------------------------
# Secondary sales analysis - item issue, month against month
# ---------------------------------------------------------------------------
#
# The export spells one warehouse differently from the master data, and a name
# that matches nothing would drop a whole warehouse out of the network total
# without saying so. Known differences are mapped; anything else is reported.
ITEM_ISSUE_ALIASES = {
    "PATHANAMTHITA": "PATHANAMTHITTA",
}


def _ii_name(raw: str) -> str:
    name = (raw or "").strip().upper()
    return ITEM_ISSUE_ALIASES.get(name, name)


def _ii_row(label: str, kind: str, members: list[str], cur: dict, prior: dict,
            last: dict) -> dict:
    """One line: the period, the one before it, and the distance between them."""
    from . import itemissue as ii

    def block(source):
        legs = {k: 0 for k, _ in ii.LEGS}
        for name in members:
            cell = source.get(name)
            if not cell:
                continue
            for k, _ in ii.LEGS:
                legs[k] += int(cell.get(k, 0) or 0)
        legs["total"] = legs["stn"] + legs["gtn"]
        legs["all"] = legs["total"] + legs["cfed"] + legs["bar"]
        return legs

    a, b = block(cur), block(prior)
    lastmonth = block(last)["all"] if last else None

    diff = a["all"] - b["all"]
    # A rise from nothing has no percentage - there is nothing to be a
    # percentage of - so it is left empty rather than printed as infinity.
    pct = (diff / b["all"] * 100) if b["all"] else None

    return {"label": label, "kind": kind, "cur": a, "prior": b,
            "diff": diff, "pct": pct, "last_month": lastmonth}


def item_issue(period: str = "", prior: str = "", cluster: int | None = None) -> dict:
    """The secondary sales analysis: one pull against the month before it."""
    from . import itemissue as ii

    have = ii.periods()
    if not have:
        return {"error": "No item issue export has been uploaded yet. Upload a "
                         "pull on the Raw Data Upload page."}

    keys = [p["key"] for p in have]
    period = period if period in keys else keys[0]
    span = ii.period_of(period)
    cur_end = span[1] if span else None

    # Default comparison: the newest stored pull that ended before this one's
    # month began - which is last month, however far into it they pulled.
    if prior not in keys:
        prior = ""
        if cur_end:
            first_of_month = cur_end.replace(day=1)
            for p in have:
                if p["key"] == period:
                    continue
                if date.fromisoformat(p["end"]) < first_of_month:
                    prior = p["key"]
                    break

    # LAST MONTH is the prior month taken whole, where such a pull exists.
    last_key = ""
    if cur_end:
        prev_end = cur_end.replace(day=1) - timedelta(days=1)
        prev_start = prev_end.replace(day=1)
        want = ii.period_key(prev_start, prev_end)
        if want in keys:
            last_key = want

    cur_raw = {_ii_name(k): v for k, v in ii.load(period).items()}
    prior_raw = {_ii_name(k): v for k, v in ii.load(prior).items()} if prior else {}
    last_raw = {_ii_name(k): v for k, v in ii.load(last_key).items()} if last_key else {}

    clusters = WAREHOUSE_CLUSTERS
    known = {w for members in clusters.values() for w in members}
    unknown = sorted((set(cur_raw) | set(prior_raw)) - known)

    rows: list[dict] = []
    shown: list[str] = []
    for cid in sorted(clusters):
        if cluster in (1, 2, 3) and cid != cluster:
            continue
        members = [w for w in clusters[cid]]
        for name in members:
            rows.append(_ii_row(name, "warehouse", [name], cur_raw, prior_raw, last_raw))
        rows.append(_ii_row(f"CLUSTER - {cid}", "cluster", members,
                            cur_raw, prior_raw, last_raw))
        shown += members

    if cluster in (None, 0):
        for name in unknown:
            rows.append(_ii_row(name, "warehouse", [name], cur_raw, prior_raw, last_raw))
        shown += unknown
        rows.append(_ii_row("TOTAL", "grand", shown, cur_raw, prior_raw, last_raw))

    industry = ii.industry_get(period)

    return {
        "period": period,
        "period_label": ii.period_label(period),
        "prior": prior,
        "prior_label": ii.period_label(prior) if prior else "",
        "last_key": last_key,
        "last_label": ii.period_label(last_key) if last_key else "",
        "periods": have,
        "columns": [{"key": k, "label": label} for k, label in ii.LEGS],
        "rows": rows,
        "as_on": cur_end.isoformat() if cur_end else "",
        "prior_as_on": (ii.period_of(prior)[1].isoformat() if prior else ""),
        "day_sale": _ii_day_sale(period, prior),
        "industry": {"cases": industry.get("cases", 0),
                     "prior": industry.get("prior", 0),
                     "updated": industry.get("updated", ""),
                     "by": industry.get("by", "")},
        "unknown": unknown,
        "cluster": cluster or 0,
        "source_block": ii.source_for(period, prior),
    }


def _ii_day_sale(period: str, prior: str) -> dict:
    """The last day of each window, taken from the secondary sales already held.

    The item issue export is cumulative over its range and says nothing about
    any single day, so this comes from the daily secondary figures the app
    already carries rather than from the pull.
    """
    from . import itemissue as ii

    def one(key: str):
        span = ii.period_of(key) if key else None
        if not span:
            return None
        day = span[1]
        try:
            grid = daily_grid("secondary", day, day)
        except Exception:
            return None
        if "error" in grid:
            return None
        for row in grid.get("rows", []):
            if row.get("kind") == "grand":
                return round(float(row.get("total", 0) or 0))
        return None

    a, b = one(period), one(prior)
    diff = (a - b) if (a is not None and b is not None) else None
    pct = (diff / b * 100) if (diff is not None and b) else None
    return {"cur": a, "prior": b, "diff": diff, "pct": pct}
