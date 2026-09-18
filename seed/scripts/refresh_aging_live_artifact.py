#!/usr/bin/env python3
"""
Refresh the `ksd-aging-live` Cowork artifact from the current AGING STOCK
ANALYSIS workbook.

Design contract (mirrors refresh_warehouse_live_artifact.py /
refresh_liquidation_live_artifact.py / refresh_commitment_tracker.py):

  * `.claude/scripts/aging_live_template.html` is the SINGLE SOURCE OF TRUTH
    for the artifact's HTML/CSS/JS.  Every layout change goes there.  This
    script only computes the data and injects it as a JS `PAYLOAD` constant.
    NEVER hand-build the artifact HTML and push it via update_artifact.

  * The artifact iframe is sandboxed and cannot call bash, so the data is
    EMBEDDED rather than fetched.

  * FULL DETAIL is the canonical row source.  It carries every shop x SKU
    position with its severity tier already classified by build_aging_stock.py,
    and reconciles exactly to the SUMMARY sheet's tier totals.  We never
    re-derive the classification here -- a second implementation of the tier
    logic would drift from build_aging_stock.py's `classify_position()`.

  * Aging tiers only (Abhay, 8 Aug 2026): NON-MOVING / CRITICAL / SLOW are the
    browsable exposure.  HEALTHY and NEW STOCK are carried as single context
    numbers so the denominator ("% of closing stock") stays honest, but they
    are not drillable.

History: every build appends a row to `Aging stock/_history/aging_history.csv`
so the artifact can show whether the exposure is actually moving.  The archive
folder is back-filled on first run.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

# --------------------------------------------------------------------------
# Paths
# --------------------------------------------------------------------------

SCRIPT_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = SCRIPT_DIR / "aging_live_template.html"


def _default_root() -> Path:
    """Walk up from this script to find the Claude folder root.

    Same probe as the commitment PDF builders -- assuming `parents[2]` breaks
    the moment the script is called from a different depth.
    """
    for p in [SCRIPT_DIR] + list(SCRIPT_DIR.parents):
        if (p / "MASTER DATA CONFIRMED.xlsx").exists():
            return p
        if (p / "Aging stock").is_dir() and (p / "KSBC shop sales").is_dir():
            return p
    return SCRIPT_DIR.parent.parent


def _default_out() -> Path:
    """Prefer the session outputs dir for scheduled runs."""
    home = os.environ.get("HOME", "")
    if home:
        outputs = Path(home) / "mnt" / "outputs"
        if outputs.is_dir():
            return outputs / "aging_live.html"
    return SCRIPT_DIR / "aging_live.html"


# --------------------------------------------------------------------------
# Canonical cluster / ASM map -- imported so it cannot drift
# --------------------------------------------------------------------------

try:
    sys.path.insert(0, str(SCRIPT_DIR))
    from pace_engine import BOND_CLUSTER, ASM_BY_CLUSTER  # type: ignore
except Exception:  # pragma: no cover - fallback mirrors CLAUDE.md exactly
    _CLUSTERS = {
        1: ["ALAPPUZHA", "ATTINGAL", "NEDUMANGAD", "KOLLAM", "KOTTARAKARA",
            "PATHANAMTHITTA"],
        2: ["THRISSUR", "TRIPUNITHURA", "THODUPUZHA", "KOTTAYAM", "ALUVA"],
        3: ["KANNUR", "KOZHIKODE", "PALAKKAD", "PERINTHALMANNA"],
    }
    BOND_CLUSTER = {b: c for c, bs in _CLUSTERS.items() for b in bs}
    ASM_BY_CLUSTER = {1: "Dinesh", 2: "Sojan", 3: "Haridasan"}


# --------------------------------------------------------------------------
# Tier vocabulary
# --------------------------------------------------------------------------

# The severity strings build_aging_stock.py writes into FULL DETAIL col A.
# Keyed on a normalised token so a glyph change upstream cannot silently drop a
# whole tier into the "unknown" bucket (it would abort instead -- see below).
TIERS = {
    "NON-MOVING":     {"key": "nm",  "label": "Non-Moving",     "icon": "●"},
    "CRITICAL AGING": {"key": "cr",  "label": "Critical Aging", "icon": "▲"},
    "SLOW MOVING":    {"key": "sl",  "label": "Slow Moving",    "icon": "◆"},
    "NEW STOCK":      {"key": "new", "label": "New Stock",      "icon": "✦"},
    "HEALTHY":        {"key": "ok",  "label": "Healthy",        "icon": "✓"},
}
AGING_KEYS = ("nm", "cr", "sl")


def tier_key(raw) -> str | None:
    """Map a FULL DETAIL severity cell to a tier key, ignoring the icon."""
    if not raw:
        return None
    txt = re.sub(r"[^A-Z \-]", "", str(raw).upper()).strip()
    for name, meta in TIERS.items():
        if name in txt:
            return meta["key"]
    return None


# --------------------------------------------------------------------------
# Workbook discovery
# --------------------------------------------------------------------------

WB_RE = re.compile(r"AGING STOCK ANALYSIS - (\w{3}) to (\w{3}) (\d{1,2})\.xlsx$", re.I)
MONTHS = {m: i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
     "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], start=1)}


def parse_wb_name(path: Path):
    """-> (from_month, to_month, day) or None."""
    m = WB_RE.search(path.name)
    if not m:
        return None
    a, b, d = m.group(1).title(), m.group(2).title(), int(m.group(3))
    if a not in MONTHS or b not in MONTHS:
        return None
    return a, b, d


def _anchor_year(to_month: str, day: int, mtime: float) -> int:
    """Anchor the window's end date to the file's own mtime.

    The workbook name carries no year, so a Dec/Jan window would flip to the
    wrong side of the boundary if we anchored to `today`.  mtime is when the
    build actually ran, which is the same clock build_aging_stock.py used.
    """
    stamp = datetime.fromtimestamp(mtime)
    yr = stamp.year
    # A Dec-dated workbook written in January belongs to the prior year.
    if MONTHS[to_month] == 12 and stamp.month == 1:
        yr -= 1
    return yr


def find_workbook(root: Path, explicit: str | None) -> Path:
    if explicit:
        p = Path(explicit)
        if not p.exists():
            raise FileNotFoundError(f"workbook not found: {p}")
        return p
    folder = root / "Aging stock"
    cands = [p for p in folder.glob("AGING STOCK ANALYSIS - *.xlsx")
             if not p.name.startswith("~$") and parse_wb_name(p)]
    if not cands:
        raise FileNotFoundError(
            f"no AGING STOCK ANALYSIS workbook in {folder}. "
            "Run build_aging_stock.py first.")
    # Newest by (end month, day) then mtime -- never by name sort, which would
    # put 'Jul 3' after 'Jul 30'.
    def sort_key(p: Path):
        _, to_m, day = parse_wb_name(p)
        return (_anchor_year(to_m, day, p.stat().st_mtime), MONTHS[to_m], day,
                p.stat().st_mtime)
    return max(cands, key=sort_key)


# --------------------------------------------------------------------------
# Extraction
# --------------------------------------------------------------------------

def _num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return 0.0


def _cover(v):
    """Months of cover -> float, or None for the 'NO SALES' badge."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().upper()
    if "NO SALES" in s or s in {"-", "—", "INF", "∞"}:
        return None
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return None


def _trend(v):
    """Normalise the within-window closing-direction column to a token."""
    s = (str(v) if v is not None else "").strip()
    if "▲" in s or "rising" in s.lower():
        return "up"
    if "▼" in s or "falling" in s.lower():
        return "down"
    if "new" in s.lower():
        return "new"
    return "flat"


def _shop_label(raw) -> str:
    """'7017-Ravipuram (PALARIVATTOM)' -> 'Ravipuram (PALARIVATTOM)'.

    The code is a separate column, so repeating it inside the name wastes the
    width the drill-down needs for the actual outlet.
    """
    s = (str(raw) if raw is not None else "").strip()
    s = re.sub(r"^\d{4,6}\s*[-–]\s*", "", s)
    return s.strip() or "(unnamed)"


BRAND_SHORT = [
    (r"MORNING WALKER",    "MORNING WALKERS"),
    (r"MAGIC BLEND",       "MAGIC BLEND"),
    (r"OLD PEARL",         "OLD PEARL"),
    (r"BCB",               "BCB NO.1"),
    (r"BLENDER",           "BLENDER'S CHOICE"),
    (r"KS ?99",            "K.S 99"),
    (r"ROYAL OLD FORT",    "ROYAL OLD FORT"),
    (r"CHAIRMAN",          "CHAIRMAN'S CHOICE"),
]


def brand_short(name: str) -> str:
    """Canonical short brand name.

    PUNCTUATION IS STRIPPED BEFORE MATCHING, and that matters: the 14 May
    workbook spells it "MORNING WALKER'S XO BRANDY" while every later one uses
    "MORNING WALKERS XO BRANDY". Matching the raw string sent the apostrophe
    spelling down the `.title()` fallback, so the long comparison window saw
    one brand vanish and an unrelated one appear -- a fabricated -200 cs
    alongside a fabricated +200 cs. Any brand-name drift upstream now
    normalises to the same key.
    """
    up = re.sub(r"[^A-Z0-9 ]", "", (name or "").upper())
    for pat, short in BRAND_SHORT:
        if re.search(pat, up):
            return short
    return (name or "").title()


def read_workbook(path: Path) -> dict:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if "FULL DETAIL" not in wb.sheetnames:
            raise ValueError(f"{path.name} has no FULL DETAIL sheet -- is this "
                             "an aging stock analysis workbook?")
        ws = wb["FULL DETAIL"]

        # Resolve columns by HEADER NAME, not position.  Column order has
        # changed twice in this workbook's history (Trend was appended 22 May
        # 2026); positional indexing would have silently read the wrong field.
        header_row, header = None, {}
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=8,
                                             values_only=True), start=1):
            cells = {str(c).strip().lower(): j
                     for j, c in enumerate(row) if c is not None}
            if "severity" in cells and "bond" in cells:
                header_row, header = i, cells
                break
        if header_row is None:
            raise ValueError("FULL DETAIL header row not found (no 'Severity')")

        def col(*names, required=True):
            for n in names:
                if n in header:
                    return header[n]
            if required:
                raise ValueError(
                    f"FULL DETAIL is missing a column matching {names!r}. "
                    f"Found: {sorted(header)}")
            return None

        C = {
            "sev":   col("severity"),
            "bond":  col("bond"),
            "staff": col("field staff"),
            "code":  col("shop code"),
            "shop":  col("shop name"),
            "brand": col("brand"),
            "pack":  col("pack"),
            "sales": col("total sales (cs)", "total sales"),
            "close": col("latest closing", "latest closing (cs)"),
            "cover": col("months of cover"),
            "zero":  col("zero months"),
            "trend": col("trend", required=False),
            "pcode": col("product code", required=False),
        }

        lines, ctx = [], {k: {"cs": 0.0, "n": 0} for k in
                          ("nm", "cr", "sl", "new", "ok")}
        # EVERY line, all five tiers, for the line-level history. The changelog
        # has to tell "this went healthy" apart from "this sold out" apart from
        # "this shop stopped reporting", and only the healthy rows can do that.
        all_rows = []
        unknown = 0
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            if row[C["sev"]] is None and row[C["bond"]] is None:
                continue
            key = tier_key(row[C["sev"]])
            if key is None:
                unknown += 1
                continue
            closing = _num(row[C["close"]])
            ctx[key]["cs"] += closing
            ctx[key]["n"] += 1
            bond = str(row[C["bond"]] or "").strip().upper()
            brand = str(row[C["brand"]] or "").strip()
            code = str(row[C["code"]] or "").strip()
            pcode = str(row[C["pcode"]] or "").strip() if C["pcode"] else ""
            all_rows.append({
                "code": code, "pcode": pcode,
                "shop": _shop_label(row[C["shop"]]),
                "bond": bond, "br": brand_short(brand),
                "p": str(row[C["pack"]] or "").strip(),
                "t": key, "cl": closing, "sa": _num(row[C["sales"]]),
                "z": int(_num(row[C["zero"]])),
                # cover is needed on EVERY tier, not just the aging ones: the
                # cover buckets and the proven-buyer test both read it off
                # healthy rows. Omitting it here silently put all 775 aging
                # lines in the "unclearable" bucket.
                "cv": _cover(row[C["cover"]]),
            })
            if key not in AGING_KEYS:
                continue  # Healthy / New Stock stay as context only
            lines.append({
                "t":  key,
                "b":  bond,
                "st": str(row[C["staff"]] or "").strip().upper() or "VACANT",
                "c":  str(row[C["code"]] or "").strip(),
                "s":  _shop_label(row[C["shop"]]),
                "br": brand_short(brand),
                "p":  str(row[C["pack"]] or "").strip(),
                # NOT rounded: every bond/brand/cluster total is a sum of these,
                # and rounding 710 lines to 2dp before summing drifted the bond
                # totals ~0.05 cs off the workbook's own figures.  Rounding
                # happens once, at display time, in the template.
                "sa": _num(row[C["sales"]]),
                "cl": closing,
                "cv": _cover(row[C["cover"]]),
                "z":  int(_num(row[C["zero"]])),
                "tr": _trend(row[C["trend"]]) if C["trend"] is not None else "flat",
            })

        if unknown:
            # Never silently drop rows -- an unrecognised tier means the
            # upstream vocabulary changed and every total below is wrong.
            raise ValueError(
                f"{unknown} FULL DETAIL rows carry an unrecognised severity. "
                "TIERS in this script is out of sync with build_aging_stock.py.")
        if not lines:
            raise ValueError("no aging positions found in FULL DETAIL")

        meta = _read_summary_meta(wb)
        return {"lines": lines, "ctx": ctx, "meta": meta, "all": all_rows}
    finally:
        wb.close()


def _read_summary_meta(wb) -> dict:
    """Pull the period strap and the active-shop count off SUMMARY.

    Cosmetic only -- a missing SUMMARY never gates the build.  `network_shops`
    is the MASTER DATA active-KSBC count that build_aging_stock.py stamps into
    its TOTAL row; it is the honest denominator for "how many shops are
    carrying aging stock", which is NOT the same as the count of shops that
    happen to appear in FULL DETAIL.
    """
    out = {"period": "", "days": None, "network_shops": None}
    if "SUMMARY" not in wb.sheetnames:
        return out
    ws = wb["SUMMARY"]
    for row in ws.iter_rows(min_row=1, max_row=40, values_only=True):
        for c in row:
            if not isinstance(c, str):
                continue
            if not out["period"] and "·" in c and re.search(r"\d+\s*days", c):
                out["period"] = c.split("·")[0].strip()
                m = re.search(r"(\d+)\s*days", c)
                if m:
                    out["days"] = int(m.group(1))
            if out["network_shops"] is None:
                m = re.search(r"(\d{2,4})\s*shops?\b", c)
                if m and "network" in c.lower():
                    out["network_shops"] = int(m.group(1))
        if out["period"] and out["network_shops"] is not None:
            break
    return out


# --------------------------------------------------------------------------
# Aggregation
# --------------------------------------------------------------------------

def _blank_tiers():
    return {k: 0.0 for k in AGING_KEYS} | {f"{k}_n": 0 for k in AGING_KEYS}


def _add(acc, ln):
    acc[ln["t"]] += ln["cl"]
    acc[f"{ln['t']}_n"] += 1


def _finish(acc):
    acc["aging"] = round(sum(acc[k] for k in AGING_KEYS), 2)
    acc["rows"] = sum(acc[f"{k}_n"] for k in AGING_KEYS)
    for k in AGING_KEYS:
        acc[k] = round(acc[k], 2)
    return acc


def aggregate(lines, ctx) -> dict:
    bonds, brands, packs, staff, bp = {}, {}, {}, {}, {}
    for ln in lines:
        for store, key in ((bonds, ln["b"]), (brands, ln["br"]),
                           (packs, ln["p"]), (staff, ln["st"]),
                           (bp, (ln["br"], ln["p"]))):
            if key not in store:
                store[key] = _blank_tiers()
            _add(store[key], ln)

    # --- bonds -----------------------------------------------------------
    bond_rows = []
    for name, acc in bonds.items():
        mine = [l for l in lines if l["b"] == name]
        _finish(acc)
        # Worst single position in the bond, by CASES STUCK.
        #
        # BUG FIXED 9 Aug 2026: the key was `(t == "nm", cl)`, which sorts
        # every Non-Moving line above every Critical one -- so a 10 cs
        # non-mover beat a 25 cs critical line and the "worst position" was
        # wrong on 10 of 15 bonds. Tier is not the ranking here; exposure is.
        worst = max(mine, key=lambda l: l["cl"])
        staff_names = sorted({l["st"] for l in mine})
        bond_rows.append({
            "bond": name,
            "cluster": BOND_CLUSTER.get(name, 0),
            "staff": " · ".join(staff_names),
            "vacant": all(s == "VACANT" for s in staff_names),
            "shops": len({l["c"] for l in mine}),
            "skus": len({(l["br"], l["p"]) for l in mine}),
            "worst": {"shop": worst["s"], "brand": worst["br"],
                      "pack": worst["p"], "cs": worst["cl"], "t": worst["t"]},
            **{k: acc[k] for k in AGING_KEYS},
            **{f"{k}_n": acc[f"{k}_n"] for k in AGING_KEYS},
            "aging": acc["aging"], "rows": acc["rows"],
        })
    bond_rows.sort(key=lambda r: -r["aging"])

    # --- clusters --------------------------------------------------------
    clusters = []
    for cn in (1, 2, 3):
        mem = [b for b in bond_rows if b["cluster"] == cn]
        if not mem:
            continue
        clusters.append({
            "n": cn,
            "asm": ASM_BY_CLUSTER.get(cn, "—"),
            "bonds": len(mem),
            "aging": round(sum(b["aging"] for b in mem), 2),
            "rows": sum(b["rows"] for b in mem),
            "shops": sum(b["shops"] for b in mem),
            "vacant": sum(1 for b in mem if b["vacant"]),
            **{k: round(sum(b[k] for b in mem), 2) for k in AGING_KEYS},
            "worst_bond": max(mem, key=lambda b: b["aging"])["bond"],
        })
    clusters.sort(key=lambda c: -c["aging"])

    def flat(store, label):
        out = []
        for name, acc in store.items():
            _finish(acc)
            out.append({label: name, **{k: acc[k] for k in AGING_KEYS},
                        **{f"{k}_n": acc[f"{k}_n"] for k in AGING_KEYS},
                        "aging": acc["aging"], "rows": acc["rows"]})
        out.sort(key=lambda r: -r["aging"])
        return out

    matrix = []
    for (br, pk), acc in bp.items():
        _finish(acc)
        matrix.append({"brand": br, "pack": pk, "aging": acc["aging"],
                       "rows": acc["rows"],
                       **{k: acc[k] for k in AGING_KEYS}})
    matrix.sort(key=lambda r: -r["aging"])

    aging_cs = round(sum(ctx[k]["cs"] for k in AGING_KEYS), 2)
    closing_cs = round(sum(v["cs"] for v in ctx.values()), 2)

    return {
        "totals": {
            **{k: round(ctx[k]["cs"], 2) for k in ctx},
            **{f"{k}_n": ctx[k]["n"] for k in ctx},
            "aging": aging_cs,
            "aging_n": sum(ctx[k]["n"] for k in AGING_KEYS),
            "closing": closing_cs,
            "aging_pct": round(aging_cs / closing_cs * 100, 2) if closing_cs else 0,
            "shops": len({l["c"] for l in lines}),
            "bonds": len(bond_rows),
            "skus": len({(l["br"], l["p"]) for l in lines}),
        },
        "bonds": bond_rows,
        "clusters": clusters,
        "brands": flat(brands, "brand"),
        "packs": flat(packs, "pack"),
        "staff": flat(staff, "staff"),
        "matrix": matrix,
    }


# --------------------------------------------------------------------------
# History
# --------------------------------------------------------------------------

HIST_COLS = ["date", "workbook", "nm", "cr", "sl", "aging", "healthy",
             "newstock", "closing", "rows", "shops"]
BOND_HIST_COLS = ["date", "bond", "nm", "cr", "sl", "aging", "rows", "shops"]
LINE_HIST_COLS = ["date", "code", "pcode", "shop", "bond", "brand", "pack",
                  "tier", "cl", "sa", "z"]

# How many snapshots of line-level detail to retain. Every snapshot is ~2,800
# rows, so this bounds the file at roughly 70k rows / 6 MB no matter how long
# the stream runs. Older network and per-bond history is NOT pruned -- those
# are small, and the trend needs the full span.
LINE_HIST_KEEP = 24

# How many recent comparison windows to precompute (the earliest snapshot is
# always added on top, so "since the beginning" is always available).
DIFF_WINDOWS = 6


def _line_key(r) -> str:
    """Identity of a position across snapshots.

    (shop, product code) rather than (shop, brand, pack): build_aging_stock.py
    already aggregates duplicate shop x SKU source lines, so the product code
    is unique per shop, and it survives a brand being renamed upstream.
    """
    return f"{r['code']}|{r['pcode'] or r['br'] + ':' + r['p']}"


def _hist_row(path: Path, agg: dict) -> dict:
    _, to_m, day = parse_wb_name(path)
    yr = _anchor_year(to_m, day, path.stat().st_mtime)
    t = agg["totals"]
    return {
        "date": f"{yr}-{MONTHS[to_m]:02d}-{day:02d}",
        "workbook": path.name,
        "nm": t["nm"], "cr": t["cr"], "sl": t["sl"], "aging": t["aging"],
        "healthy": t["ok"], "newstock": t["new"], "closing": t["closing"],
        "rows": t["aging_n"], "shops": t["shops"],
    }


def _bond_hist_rows(date: str, agg: dict) -> list[dict]:
    return [{"date": date, "bond": b["bond"], "nm": b["nm"], "cr": b["cr"],
             "sl": b["sl"], "aging": b["aging"], "rows": b["rows"],
             "shops": b["shops"]} for b in agg["bonds"]]


def _line_hist_rows(date: str, all_rows: list) -> list[dict]:
    """One row per position.

    Cases are twelfths and forty-eighths, so they are non-terminating
    decimals. Storing them at 3dp cost ~0.0005 cs each, which over 400+ lines
    in a tier summed to a visible 0.1 cs disagreement against the workbook.
    Full precision costs a few bytes a row and removes the whole class of
    problem.
    """
    return [{"date": date, "code": r["code"], "pcode": r["pcode"],
             "shop": r["shop"], "bond": r["bond"], "brand": r["br"],
             "pack": r["p"], "tier": r["t"], "cl": repr(float(r["cl"])),
             "sa": repr(float(r["sa"])), "z": r["z"]} for r in all_rows]


def update_history(root: Path, current: Path, agg: dict, all_rows: list,
                   backfill: bool = True, rebuild_lines: bool = False):
    """Append/refresh this snapshot; back-fill from _archive on first run.

    TWO files, deliberately.  The network file answers "is the total moving";
    the per-bond file answers "who is getting worse", which is the question
    that actually puts a name on an action.  A network total that holds steady
    while one bond doubles and another halves looks like nothing happened.

    Re-running a build for the same window date UPDATES that snapshot rather
    than duplicating it -- the same rule build_warehouse_stock.py uses for its
    stock history (a skip-as-duplicate froze an incomplete first set).
    """
    hdir = root / "Aging stock" / "_history"
    hdir.mkdir(parents=True, exist_ok=True)
    hpath = hdir / "aging_history.csv"
    bpath = hdir / "aging_bond_history.csv"
    lpath = hdir / "aging_line_history.csv"

    rows, brows, lrows = {}, {}, {}
    if hpath.exists():
        with hpath.open(newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                if r.get("date"):
                    rows[r["date"]] = r
    if bpath.exists():
        with bpath.open(newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                if r.get("date") and r.get("bond"):
                    brows.setdefault(r["date"], {})[r["bond"]] = r
    if lpath.exists() and not rebuild_lines:
        with lpath.open(newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                if r.get("date"):
                    lrows.setdefault(r["date"], []).append(r)

    if backfill:
        arch = root / "Aging stock" / "_archive"
        if arch.is_dir():
            for p in sorted(arch.glob("AGING STOCK ANALYSIS - *.xlsx")):
                if p.name.startswith("~$") or not parse_wb_name(p):
                    continue
                _, to_m, day = parse_wb_name(p)
                d = f"{_anchor_year(to_m, day, p.stat().st_mtime)}-{MONTHS[to_m]:02d}-{day:02d}"
                if d in rows and d in brows and d in lrows:
                    continue
                try:
                    raw = read_workbook(p)
                    a = aggregate(raw["lines"], raw["ctx"])
                    rows[d] = _hist_row(p, a)
                    brows[d] = {r["bond"]: r for r in _bond_hist_rows(d, a)}
                    lrows[d] = _line_hist_rows(d, raw["all"])
                except Exception as e:  # a stale archive file must not block
                    print(f"  ! history back-fill skipped {p.name}: {e}")

    cur = _hist_row(current, agg)
    rows[cur["date"]] = cur
    brows[cur["date"]] = {r["bond"]: r for r in _bond_hist_rows(cur["date"], agg)}
    lrows[cur["date"]] = _line_hist_rows(cur["date"], all_rows)

    # keep the line file bounded; network/bond history stay complete
    for d in sorted(lrows)[:-LINE_HIST_KEEP]:
        lrows.pop(d, None)

    with hpath.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=HIST_COLS)
        w.writeheader()
        for d in sorted(rows):
            w.writerow({k: rows[d].get(k, "") for k in HIST_COLS})
    with bpath.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=BOND_HIST_COLS)
        w.writeheader()
        for d in sorted(brows):
            for bond in sorted(brows[d]):
                w.writerow({k: brows[d][bond].get(k, "") for k in BOND_HIST_COLS})
    with lpath.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=LINE_HIST_COLS)
        w.writeheader()
        for d in sorted(lrows):
            for r in lrows[d]:
                w.writerow({k: r.get(k, "") for k in LINE_HIST_COLS})

    dates = sorted(rows)
    hist = [{k: (v if k in ("date", "workbook") else _num(v))
             for k, v in rows[d].items() if k in HIST_COLS} for d in dates]

    # Per-bond series aligned to `dates`. A bond absent from an older snapshot
    # gets None, NOT 0 -- a gap in coverage is not the same as "no aging
    # stock", and plotting it as zero would invent a recovery that never
    # happened.
    bond_hist = {}
    for b in agg["bonds"]:
        bond_hist[b["bond"]] = [
            (_num(brows[d][b["bond"]]["aging"]) if b["bond"] in brows.get(d, {})
             else None) for d in dates]

    # line history, normalised back to floats for the diff engine
    lines_by_date = {}
    for d, rs in lrows.items():
        lines_by_date[d] = [{**r, "cl": _num(r["cl"]), "sa": _num(r["sa"]),
                             "z": int(_num(r["z"]))} for r in rs]

    # VOCABULARY TRIPWIRE. A brand or pack that exists in one snapshot and not
    # another produces a diff that looks like a huge real move but is really a
    # spelling change upstream (this fired for MORNING WALKER'S vs MORNING
    # WALKERS on the 14 May workbook). Warn rather than abort -- a brand
    # genuinely can drop to zero aging -- but say so loudly.
    if lines_by_date:
        newest = max(lines_by_date)
        base = {r["brand"] for r in lines_by_date[newest]}
        for d in sorted(lines_by_date):
            if d == newest:
                continue
            odd = {r["brand"] for r in lines_by_date[d]} - base
            if odd:
                print(f"  ! {d}: brand name(s) not seen in the latest snapshot "
                      f"-- {', '.join(sorted(odd))}. Check brand_short(); a "
                      "spelling drift will read as a phantom brand move.")
    return hist, bond_hist, lines_by_date


# --------------------------------------------------------------------------
# Diff engine -- the whole point of the changelog
# --------------------------------------------------------------------------

# A line has to move by more than this to be reported at all. Aging cases are
# twelfths and forty-eighths of a case, so tiny float residue would otherwise
# fill the feed with "+0.0 cs" entries that are really just rounding.
EPS = 0.05
AGING_SET = set(AGING_KEYS)
TIERS_ORDER = ("nm", "cr", "sl", "new", "ok")


def _tier_of(r):
    return r["tier"] if "tier" in r else r["t"]


def build_diff(prev_rows: list, cur_rows: list, prev_date: str,
               cur_date: str) -> dict:
    """Compare two line-level snapshots.

    Everything the changelog shows comes from here: nothing is carried over
    from the current snapshot alone, so a panel can never quietly display a
    static figure dressed up as a change.
    """
    P = {_line_key(_norm(r)): _norm(r) for r in prev_rows}
    C = {_line_key(_norm(r)): _norm(r) for r in cur_rows}

    def aging_cs(store):
        return sum(r["cl"] for r in store.values() if r["t"] in AGING_SET)

    def tier_cs(store, k):
        return sum(r["cl"] for r in store.values() if r["t"] == k)

    newly, cleared, up, down = [], [], [], []

    for k, c in C.items():
        p = P.get(k)
        c_ag = c["t"] in AGING_SET
        p_ag = bool(p) and p["t"] in AGING_SET

        if c_ag and not p_ag:
            # A position that was healthy, brand-new, or simply absent last
            # time and is now aging. `was` is what makes it readable: "went
            # from healthy" is a different story from "arrived and stalled".
            newly.append({**_disp(c), "cl": c["cl"],
                          "was": (p["t"] if p else "absent"),
                          "was_cl": (p["cl"] if p else 0.0)})
        elif p_ag and c_ag and abs(c["cl"] - p["cl"]) > EPS:
            d = c["cl"] - p["cl"]
            (up if d > 0 else down).append({**_disp(c), "cl": c["cl"],
                                            "was_cl": p["cl"], "d": d})

    for k, p in P.items():
        if p["t"] not in AGING_SET:
            continue
        c = C.get(k)
        if c and c["t"] in AGING_SET:
            continue
        # Left the aging population. Three different endings, and they are
        # not equally good news -- "sold out" and "now selling" are wins,
        # "stopped reporting" is a data gap wearing a win's clothes.
        if c is None:
            how = "gone"
        elif c["cl"] <= 0.01:
            how = "soldout"
        else:
            how = "healthy" if c["t"] == "ok" else c["t"]
        cleared.append({**_disp(p), "was_cl": p["cl"], "how": how,
                        "now_cl": (c["cl"] if c else 0.0)})

    # ---- per-bond from / to / lines, for EVERY bond -------------------
    # Not just the ones that moved: a bond sitting still is a fact the table
    # has to be able to state, and "absent from the list" is not the same
    # statement as "flat".
    def roll(store, keyfn):
        out = {}
        for r in store.values():
            if r["t"] not in AGING_SET:
                continue
            k = keyfn(r)
            g = out.setdefault(k, {"cs": 0.0, "n": 0})
            g["cs"] += r["cl"]
            g["n"] += 1
        return out

    def movement(keyfn, label):
        pv, cv = roll(P, keyfn), roll(C, keyfn)
        rows = []
        for k in sorted(set(pv) | set(cv)):
            a = pv.get(k, {"cs": 0.0, "n": 0})
            b = cv.get(k, {"cs": 0.0, "n": 0})
            rows.append({label: k, "from": round(a["cs"], 2),
                         "to": round(b["cs"], 2),
                         "d": round(b["cs"] - a["cs"], 2),
                         "n_from": a["n"], "n_to": b["n"]})
        rows.sort(key=lambda r: -r["d"])
        return rows

    bond_mv = movement(lambda r: r["bond"], "bond")
    brand_mv = movement(lambda r: r["br"], "brand")
    pack_mv = movement(lambda r: r["p"], "pack")

    # ---- TIER FLOW -----------------------------------------------------
    # Where stock moved BETWEEN tiers. A net "+244 into Critical" does not say
    # whether it arrived from Healthy, decayed out of Slow, or is new stock
    # landing badly -- and those need different responses. Cases are measured
    # at the destination, except for lines that left the report entirely,
    # which are measured where they were.
    flow_acc = {}
    for k in set(P) | set(C):
        p, c = P.get(k), C.get(k)
        src = p["t"] if p else "absent"
        dst = c["t"] if c else "gone"
        if src == dst:
            continue
        cs = (c["cl"] if c else p["cl"])
        g = flow_acc.setdefault((src, dst), {"lines": 0, "cs": 0.0})
        g["lines"] += 1
        g["cs"] += cs
    flow = sorted(({"from": a, "to": b, "lines": v["lines"],
                    "cs": round(v["cs"], 2)} for (a, b), v in flow_acc.items()
                   if v["cs"] > EPS), key=lambda r: -r["cs"])

    # Bond deltas come from the CLEAN from/to rollup above, not from
    # accumulating per-line deltas.
    #
    # BUG FIXED 9 Aug 2026: the accumulate approach double-subtracted any line
    # that left the aging population but still existed in the current snapshot
    # -- once in the forward pass (as `0 - p.cl`) and again in the cleared
    # pass. KOZHIKODE read +76.5 against a true +84.4. The rollup cannot drift
    # this way because it never looks at a line twice, and the assertion below
    # would now catch it.
    bonds = sorted(({"bond": r["bond"], "d": r["d"]} for r in bond_mv
                    if abs(r["d"]) > EPS), key=lambda r: -r["d"])
    _net = sum(r["d"] for r in bond_mv)
    _true = aging_cs(C) - aging_cs(P)
    if abs(_net - _true) > 0.5:
        raise ValueError(
            f"bond movement does not reconcile to the network total for "
            f"{prev_date}->{cur_date}: bonds sum {_net:.2f} vs {_true:.2f}. "
            "The diff is wrong -- do not publish it.")

    # name the shops driving each bond's move, worst first
    shop_d = {}
    for src, sign in ((C, 1), (P, -1)):
        for k, r in src.items():
            if r["t"] not in AGING_SET:
                continue
            other = (P if sign == 1 else C).get(k)
            o_cs = other["cl"] if other and other["t"] in AGING_SET else 0.0
            if sign == 1:
                d = r["cl"] - o_cs
            else:
                d = 0.0 if (other and other["t"] in AGING_SET) else -r["cl"]
            if abs(d) > EPS:
                key = (r["bond"], r["code"], r["shop"])
                shop_d[key] = shop_d.get(key, 0.0) + d
    by_bond_shops = {}
    for (bond, code, shop), d in shop_d.items():
        by_bond_shops.setdefault(bond, []).append(
            {"code": code, "shop": shop, "d": round(d, 2)})
    for b in by_bond_shops:
        by_bond_shops[b].sort(key=lambda r: -abs(r["d"]))
    for b in bonds:
        b["shops"] = by_bond_shops.get(b["bond"], [])[:6]

    a_prev, a_cur = aging_cs(P), aging_cs(C)
    cl_prev = sum(r["cl"] for r in P.values())
    cl_cur = sum(r["cl"] for r in C.values())

    up.sort(key=lambda r: -r["d"])
    down.sort(key=lambda r: r["d"])
    newly.sort(key=lambda r: -r["cl"])
    cleared.sort(key=lambda r: -r["was_cl"])

    return {
        "from": prev_date, "to": cur_date,
        "aging_from": round(a_prev, 2), "aging_to": round(a_cur, 2),
        "d_aging": round(a_cur - a_prev, 2),
        "closing_from": round(cl_prev, 2), "closing_to": round(cl_cur, 2),
        "share_from": round(a_prev / cl_prev * 100, 2) if cl_prev else 0,
        "share_to": round(a_cur / cl_cur * 100, 2) if cl_cur else 0,
        # ALL FIVE tiers, not just the aging three: the KPI strip shows New
        # Stock and Healthy too, and a tier tile with no movement figure is
        # exactly the static panel this page is not supposed to contain.
        "tiers": {k: {"from": round(tier_cs(P, k), 2),
                      "to": round(tier_cs(C, k), 2),
                      "n_from": sum(1 for r in P.values() if r["t"] == k),
                      "n_to": sum(1 for r in C.values() if r["t"] == k)}
                  for k in TIERS_ORDER},
        "bond_mv": bond_mv, "brand_mv": brand_mv, "pack_mv": pack_mv,
        "flow": flow,
        "bonds": bonds,
        "worse": [b for b in bonds if b["d"] > EPS],
        "better": [b for b in bonds if b["d"] < -EPS],
        "newly": newly, "cleared": cleared,
        "up": up[:40], "down": down[:40],
        "n_newly": len(newly), "cs_newly": round(sum(r["cl"] for r in newly), 2),
        "n_cleared": len(cleared),
        "cs_cleared": round(sum(r["was_cl"] for r in cleared), 2),
        "n_up": len(up), "cs_up": round(sum(r["d"] for r in up), 2),
        "n_down": len(down), "cs_down": round(sum(r["d"] for r in down), 2),
        "lines_from": sum(1 for r in P.values() if r["t"] in AGING_SET),
        "lines_to": sum(1 for r in C.values() if r["t"] in AGING_SET),
    }


def longitudinal(lines_by_date: dict) -> dict:
    """Analysis that needs the WHOLE series, not a pair of snapshots.

    Two questions a window diff cannot answer:
      * what has been stuck the entire time (chronic exposure -- the stock
        that no amount of normal trading has shifted), and
      * what cleared and then came back (churn -- an incentive or a push that
        did not hold, which reads as a win in any single window).
    """
    dates = sorted(lines_by_date)
    if len(dates) < 2:
        return {"chronic": [], "chronic_cs": 0, "chronic_n": 0,
                "churn": [], "churn_n": 0, "snapshots": len(dates)}

    seen, states = {}, {}
    for d in dates:
        for r in lines_by_date[d]:
            r = _norm(r)
            k = _line_key(r)
            states.setdefault(k, []).append(r["t"] in AGING_SET)
            seen[k] = r

    latest = {_line_key(_norm(r)) for r in lines_by_date[dates[-1]]
              if _norm(r)["t"] in AGING_SET}

    chronic, churn = [], []
    for k, hist in states.items():
        if k not in latest:
            continue
        # chronic: aging in every snapshot it appeared in, and present
        # throughout -- a line first seen last week cannot be chronic
        if len(hist) == len(dates) and all(hist):
            chronic.append(seen[k])
        # churn: aging, then not, then aging again
        elif True in hist and False in hist:
            first = hist.index(True)
            if any(not x for x in hist[first:]) and hist[-1]:
                churn.append(seen[k])

    chronic.sort(key=lambda r: -r["cl"])
    churn.sort(key=lambda r: -r["cl"])
    return {
        "chronic": [_disp(r) | {"cl": round(r["cl"], 2)} for r in chronic[:12]],
        "chronic_n": len(chronic),
        "chronic_cs": round(sum(r["cl"] for r in chronic), 2),
        "churn": [_disp(r) | {"cl": round(r["cl"], 2)} for r in churn[:12]],
        "churn_n": len(churn),
        "churn_cs": round(sum(r["cl"] for r in churn), 2),
        "snapshots": len(dates),
        "span_from": dates[0], "span_to": dates[-1],
    }


def _norm(r):
    """Line-history CSV rows and in-memory rows use different key names."""
    if "t" in r:
        return r
    return {"code": r["code"], "pcode": r.get("pcode", ""), "shop": r["shop"],
            "bond": r["bond"], "br": r["brand"], "p": r["pack"],
            "t": r["tier"], "cl": r["cl"], "sa": r.get("sa", 0),
            "z": r.get("z", 0)}


def _disp(r):
    return {"code": r["code"], "shop": r["shop"], "bond": r["bond"],
            "br": r["br"], "p": r["p"], "t": r["t"], "z": r.get("z", 0),
            "sa": round(r.get("sa", 0), 2)}


# --------------------------------------------------------------------------
# Render
# --------------------------------------------------------------------------

def load_template() -> str:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"artifact template missing: {TEMPLATE_PATH}\n"
            "The live artifact cannot be rendered without it.")
    tpl = TEMPLATE_PATH.read_text(encoding="utf-8")
    if "__PAYLOAD__" not in tpl:
        raise ValueError(f"template is missing the __PAYLOAD__ placeholder: "
                         f"{TEMPLATE_PATH}")
    return tpl


def write_artifact(payload: dict, out_html: Path):
    html = load_template().replace(
        "__PAYLOAD__", json.dumps(payload, ensure_ascii=False))
    out_html.parent.mkdir(parents=True, exist_ok=True)
    out_html.write_text(html, encoding="utf-8")


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def build(root: Path, workbook: str | None, out: Path,
          backfill: bool = True, rebuild_lines: bool = False) -> dict:
    wbp = find_workbook(root, workbook)
    raw = read_workbook(wbp)
    agg = aggregate(raw["lines"], raw["ctx"])
    hist, bond_hist, lines_by_date = update_history(
        root, wbp, agg, raw["all"], backfill=backfill,
        rebuild_lines=rebuild_lines)

    # Attach each bond's own movement. Computed here rather than in the
    # template because it needs the history file, and the template only ever
    # sees the payload.
    for b in agg["bonds"]:
        series = [v for v in bond_hist.get(b["bond"], []) if v is not None]
        b["hist"] = bond_hist.get(b["bond"], [])
        b["prev"] = series[-2] if len(series) > 1 else None
        b["first"] = series[0] if series else None
        b["d_prev"] = round(b["aging"] - b["prev"], 2) if b["prev"] is not None else None
        b["d_first"] = round(b["aging"] - b["first"], 2) if b["first"] is not None else None

    _, to_m, day = parse_wb_name(wbp)

    # One precomputed diff per available comparison window. Doing this in
    # Python rather than shipping every snapshot's lines to the browser keeps
    # the payload small -- a diff is a few hundred rows, a snapshot is 2,800.
    snap_dates = sorted(lines_by_date)
    cur_date = snap_dates[-1] if snap_dates else None
    diffs = []
    if cur_date:
        priors = [d for d in snap_dates if d != cur_date]
        # every recent window, plus the earliest on record for the long view
        keep = priors[-DIFF_WINDOWS:]
        if priors and priors[0] not in keep:
            keep = [priors[0]] + keep
        for d in keep:
            diffs.append(build_diff(lines_by_date[d], lines_by_date[cur_date],
                                    d, cur_date))
    diffs.sort(key=lambda x: x["from"], reverse=True)

    # Rate of change, so the page can say whether the deterioration is
    # speeding up or easing -- two windows with the same sign can be very
    # different stories.
    for d in diffs:
        days = max(1, (datetime.strptime(d["to"], "%Y-%m-%d")
                       - datetime.strptime(d["from"], "%Y-%m-%d")).days)
        d["days"] = days
        d["per_day"] = round(d["d_aging"] / days, 2)

    longi = longitudinal(lines_by_date)

    # A snapshot taken before the NEW STOCK tier existed reports 0 cs for it,
    # which is "not measured", not "none". Plotting that as a real zero makes
    # the tier look like it grew from nothing. Flag the affected windows so
    # the page can suppress the comparison instead of lying about it.
    for d in diffs:
        d["new_tier_missing"] = (d["tiers"]["new"]["from"] <= 0.01
                                 and d["tiers"]["new"]["to"] > 1)

    # ---- deep analysis (the four levers) --------------------------------
    insights = {}
    try:
        import aging_insights as AI
        wb2 = openpyxl.load_workbook(wbp, read_only=True, data_only=True)
        try:
            bond_names = [b["bond"] for b in agg["bonds"]]
            bond_rows = AI.read_bond_sheets(wb2, bond_names, short=brand_short)
            months = max(1.0, (raw["meta"].get("days") or 180) / 30.0)
            # per-bond TOTAL shelf stock, so aging can be shown as a share of
            # the bond's OWN shelf -- ALUVA is 4th by cases but 1st by share
            closing_by_bond = {}
            for r in raw["all"]:
                closing_by_bond[r["bond"]] = closing_by_bond.get(r["bond"], 0.0) + r["cl"]
            for b in agg["bonds"]:
                b["closing"] = round(closing_by_bond.get(b["bond"], 0.0), 2)
                b["intensity"] = round(b["aging"] / b["closing"] * 100, 1) \
                    if b["closing"] else 0.0
            insights = {
                "months": round(months, 2),
                "stall": AI.stall_profile(bond_rows),
                "redispatch": AI.redispatch(bond_rows),
                "stockouts": AI.stockouts(raw["all"], months),
                "routing": AI.routing(raw["all"], months),
                "cover": AI.cover_buckets(raw["all"]),
                "newstock": AI.newstock(wb2),
                "incentives": AI.incentives(root, raw["all"], brand_short,
                                            bond_rows),
            }
        finally:
            wb2.close()
    except Exception as e:      # analysis must never break the core artifact
        print(f"  ! deep insights skipped: {e}")
        insights = {}

    payload = {
        "refresh_ts": datetime.now().isoformat(timespec="minutes"),
        "workbook": wbp.name,
        "as_of": f"{day} {to_m}",
        "period": raw["meta"].get("period") or "",
        "days": raw["meta"].get("days"),
        "network_shops": raw["meta"].get("network_shops"),
        "lines": raw["lines"],
        "history": hist,
        "hist_dates": [h["date"] for h in hist],
        "diffs": diffs,
        "snap_dates": snap_dates,
        "longi": longi,
        "iq": insights,
        **agg,
    }
    write_artifact(payload, out)
    return payload


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--base", help="Claude folder root")
    ap.add_argument("--workbook", help="explicit aging workbook path")
    ap.add_argument("--out", help="output HTML path")
    ap.add_argument("--no-backfill", action="store_true",
                    help="skip reading _archive workbooks into history")
    ap.add_argument("--rebuild-lines", action="store_true",
                    help="discard and regenerate aging_line_history.csv from "
                         "the workbooks (use after a precision or schema change)")
    args = ap.parse_args()

    root = Path(args.base) if args.base else _default_root()
    out = Path(args.out) if args.out else _default_out()

    p = build(root, args.workbook, out, backfill=not args.no_backfill,
              rebuild_lines=args.rebuild_lines)
    t = p["totals"]
    print(f"aging live artifact written: {out}")
    print(f"  source      : {p['workbook']}  (as of {p['as_of']})")
    print(f"  aging       : {t['aging']:,.2f} cs  "
          f"({t['aging_pct']:.1f}% of {t['closing']:,.2f} cs closing)")
    print(f"  non-moving  : {t['nm']:,.2f} cs / {t['nm_n']} lines")
    print(f"  critical    : {t['cr']:,.2f} cs / {t['cr_n']} lines")
    print(f"  slow        : {t['sl']:,.2f} cs / {t['sl_n']} lines")
    print(f"  context     : healthy {t['ok']:,.2f} cs · new stock "
          f"{t['new']:,.2f} cs")
    print(f"  scope       : {t['bonds']} bonds · {t['shops']} shops · "
          f"{t['aging_n']} aging positions")
    print(f"  history     : {len(p['history'])} snapshots")
    return 0


if __name__ == "__main__":
    sys.exit(main())
