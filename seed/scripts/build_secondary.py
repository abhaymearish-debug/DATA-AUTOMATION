#!/usr/bin/env python3
"""Build K.S. Distillery Secondary Sales monthly analysis."""
import os, sys
from collections import defaultdict
from datetime import datetime, date, timedelta
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

# Auto-detect session path from this file's location:
# <session>/mnt/Claude/.claude/scripts/build_secondary.py
_THIS = os.path.abspath(__file__)
CLAUDE = os.path.abspath(os.path.join(os.path.dirname(_THIS), "..", ".."))
SESSION_ROOT = os.path.abspath(os.path.join(CLAUDE, "..", ".."))
SEC_DIR = f"{CLAUDE}/Secondary sales"
MASTER = f"{CLAUDE}/MASTER DATA CONFIRMED.xlsx"
SCRATCH = f"{SESSION_ROOT}/sec_scratch.xlsx"

# MONTH / YEAR / MONTH_NUM are AUTO-DETECTED from the newest raw file on disk
# (see "AUTO-DETECT" block below) — audit fix #1. No manual constant edit on
# month rollover; they are assigned before any raw is parsed.
MONTH = None
YEAR = None
MONTH_NUM = None

MONTHS = {"JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4, "MAY": 5, "JUNE": 6,
          "JULY": 7, "AUGUST": 8, "SEPTEMBER": 9, "OCTOBER": 10, "NOVEMBER": 11, "DECEMBER": 12}

# Bottles-per-case map (audit fix #4 — loose bottles were dropped from case
# totals). Keys match the COMBINED/raw 'Pack' strings exactly. Canonical values
# from .claude/scripts/industry_analysis/industry_lib.py BPC.
BPC = {"1000 ML": 9, "750 ML": 12, "500 ML": 18, "375 ML": 24, "180 ML": 48}

# Palette
NAVY = "FF1A237E"
LIGHT_BG = "FFF8FAFC"
ALT_ROW = "FFF1F5F9"
WHITE = "FFFFFFFF"
DARK_GRAY = "FF374151"
TEXT_DARK = "FF111827"
CYAN_ACCENT = "FF00D5FF"
GOLD_ACCENT = "FFFFD700"
KSBC_FILL = "FFBBDEFB"
FED_FILL = "FFDCEDC8"
BAR_FILL = "FFFFE0B2"

FONT_NAME = "Trebuchet MS"

thin = Side(border_style="thin", color="FF9CA3AF")
thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)

gold_top = Side(border_style="medium", color=GOLD_ACCENT)
cyan_bottom = Side(border_style="medium", color=CYAN_ACCENT)


def title_style(cell):
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.font = Font(name=FONT_NAME, size=14, bold=True, color=WHITE)
    cell.alignment = Alignment(horizontal="center", vertical="center")


def header_style(cell):
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.font = Font(name=FONT_NAME, size=10, bold=True, color=WHITE)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(left=thin, right=thin, top=thin, bottom=cyan_bottom)


def data_style(cell, alt=False, numeric=False):
    cell.fill = PatternFill("solid", fgColor=ALT_ROW if alt else WHITE)
    cell.font = Font(name=FONT_NAME, size=10, color=TEXT_DARK)
    cell.alignment = Alignment(horizontal=("right" if numeric else "left"), vertical="center")
    cell.border = thin_border


def total_style(cell, numeric=False):
    cell.fill = PatternFill("solid", fgColor=DARK_GRAY)
    cell.font = Font(name=FONT_NAME, size=10, bold=True, color=WHITE)
    cell.alignment = Alignment(horizontal=("right" if numeric else "left"), vertical="center")
    cell.border = Border(left=thin, right=thin, top=gold_top, bottom=thin)


# =========================================================================
# Step 0-1: Read raw files
# =========================================================================

import os, re, glob

# ---------------------------------------------------------------------------
# AUTO-DETECT the target month (audit fix #1 — MONTH/YEAR/MONTH_NUM were
# hardcoded and silently built the wrong month on rollover). Scan for any
# 'RAW DATA ... SECONDARY SALES.xlsx' (in Secondary sales/ or Claude root),
# take the month name from the NEWEST file, and the YEAR from its row dates.
# ---------------------------------------------------------------------------
_GENERIC_RAW = re.compile(r"raw\s*data.*secondary\s*sales\s*\.xlsx$", re.IGNORECASE)
_scan = []
for _d in [SEC_DIR, CLAUDE]:
    if not os.path.isdir(_d):
        continue
    for _fn in os.listdir(_d):
        if _fn.startswith("~$"):
            continue
        if _GENERIC_RAW.search(_fn):
            _scan.append(os.path.join(_d, _fn))
if not _scan:
    raise SystemExit("No secondary sales raw files found (looked in 'Secondary sales/' and Claude root).")


def _month_token(path):
    up = os.path.basename(path).upper()
    for _mn, _num in MONTHS.items():
        if _mn in up:
            return _mn, _num
    return None, None


_scan.sort(key=lambda p: os.path.getmtime(p), reverse=True)  # newest first
_newest_raw = None
for _p in _scan:
    _mn, _num = _month_token(_p)
    if _mn:
        MONTH, MONTH_NUM, _newest_raw = _mn, _num, _p
        break
if _newest_raw is None:
    raise SystemExit("Raw file(s) present but none name a recognised month: "
                     + ", ".join(os.path.basename(p) for p in _scan))


def _detect_year(path):
    try:
        _wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        _ws = _wb["SupplierSalesAnalysisReport"]
        for _r in _ws.iter_rows(min_row=2, values_only=True):
            if _r[11] is None or str(_r[11]).strip().lower() == "total":
                continue
            try:
                _d, _m, _y = str(_r[11]).split("-")
                _wb.close()
                return int(_y)
            except Exception:
                continue
        _wb.close()
    except Exception:
        pass
    return datetime.fromtimestamp(os.path.getmtime(path)).year


YEAR = _detect_year(_newest_raw)
print(f"Auto-detected target: {MONTH} {YEAR} (MONTH_NUM={MONTH_NUM}) from {os.path.basename(_newest_raw)}")

# Auto-detect raw files for the target MONTH in SEC_DIR (and Claude root as fallback).
# Pattern (case-insensitive): anything containing 'RAW DATA' and the MONTH name,
# ending with ' SECONDARY SALES.xlsx'. Handles:
#   'RAW DATA -APRIL SECONDARY SALES.xlsx'
#   'RAW DATA -APRIL 12TH SECONDARY SALES.xlsx'
#   'RAW DATA -APRIL 5-19TH SECONDARY SALES.xlsx'
#   'RAW DATA APRIL 1-19TH SECONDARY SALES.xlsx'
raw_pattern = re.compile(r"raw\s*data.*" + MONTH.lower() + r".*secondary\s*sales\s*\.xlsx$", re.IGNORECASE)
candidate_dirs = [SEC_DIR, CLAUDE]
raw_files = []
seen = set()
for d in candidate_dirs:
    if not os.path.isdir(d):
        continue
    for fn in sorted(os.listdir(d)):
        if fn.startswith("~$"):
            continue
        if raw_pattern.search(fn):
            full = os.path.join(d, fn)
            # Read the raw IN PLACE — do NOT move it before the approval gate
            # (audit fix #9). If a raw was dropped at the Claude root, note it so
            # the operator/cleanup knows where it is.
            if d == CLAUDE and d != SEC_DIR:
                print(f"NOTE: raw found at Claude root (read in place, not moved): {fn}")
            if full not in seen:
                seen.add(full)
                raw_files.append(full)

if not raw_files:
    raise SystemExit(f"No secondary sales raw files found for {MONTH} in {SEC_DIR}")

print(f"Raw files found for {MONTH}:")
for rf in raw_files:
    print(f"  - {rf}")

# Parse the export DATE WINDOW from the raw filename(s) so the labelled period
# reflects the calendar window covered -- NOT just the first/last day that
# happened to have dispatches. KSD's day 1 is always a dry day (zero dispatches),
# and any mid-month day can legitimately be empty; those are valid "0-dispatch"
# data points and must never shorten the labelled range.
#   'raw data may 31st-secondary sales.xlsx'      -> end 31
#   'RAW DATA -MAY 1-13TH SECONDARY SALES.xlsx'    -> end 13
#   'RAW DATA -MAY 5-19TH SECONDARY SALES.xlsx'    -> end 19
#   'RAW DATA -MAY SECONDARY SALES.xlsx'           -> (no window in name)
_RANGE_RE = re.compile(r"(\d{1,2})\s*-\s*(\d{1,2})\s*(?:st|nd|rd|th)", re.IGNORECASE)
_SINGLE_RE = re.compile(r"(?<!\d)(\d{1,2})\s*(?:st|nd|rd|th)", re.IGNORECASE)
window_end = 0
for rf in raw_files:
    nm = os.path.basename(rf).lower().replace(MONTH.lower(), " ")
    mrange = _RANGE_RE.search(nm)
    if mrange:
        window_end = max(window_end, int(mrange.group(2)))
    else:
        for m1 in _SINGLE_RE.finditer(nm):
            window_end = max(window_end, int(m1.group(1)))
if window_end:
    print(f"Date window parsed from filename(s): up to day {window_end}")

source_meta = []
all_rows = []  # list of tuple rows
# DAY-LEVEL REPLACEMENT store (12 Jun 2026 fix). The old per-row dedupe key
# (inv_no, product_code) SILENTLY DROPPED legitimate split lines -- one GTN can
# list the same product code on multiple lines, and within a single raw (equal
# mtime) only the first line survived. Proven on the 11-Jun raw: 151 rows read,
# 147 retained, 13 cs lost (+2 cs on 9-Jun) vs the developer's independent
# report (2,975 vs 2,960). Now the dedupe unit is a whole DAY: seeded days carry
# priority 0.0; any real raw covering day D replaces the ENTIRE seeded day D
# (re-upload/correction semantics preserved); among raw files covering the same
# day, later mtime wins the whole day. Within the winning source every row is
# kept -- nothing is dropped row-wise.
day_store = {}    # day -> (priority_mtime, list_of_row_tuples, source_name)
_read_tally = {}  # (source_name, day) -> [rows_read_for_month, issue_cases_sum]

def _offer_day(day, mtime, rows, src):
    if day not in day_store or mtime > day_store[day][0]:
        day_store[day] = (mtime, rows, src)

def _tally(src, day, rows):
    s = 0.0
    for x in rows:
        try: s += float(x[12] or 0)
        except Exception: pass
    _read_tally[(src, day)] = [len(rows), s]

# ---------------------------------------------------------------------------
# SEED from the existing live workbook's COMBINED DISPATCHES (accumulator model).
# Secondary raws are uploaded ONE DAY AT A TIME and deleted after ingest, so the
# month's prior days live only inside COMBINED DISPATCHES. Seed the dedupe store
# from there FIRST, then layer the new single-day raw(s) on top. Dedupe key
# (inv_no, product_code) with mtime preference lets a re-uploaded/corrected day
# overwrite the seeded copy (seed mtime = 0, any real raw row mtime > 0 wins).
# Brand-new month (no <MONTH> analysis workbook yet) -> nothing to seed.
def _find_prior_workbook():
    best=None; best_n=-1
    for fn in sorted(os.listdir(SEC_DIR)):
        if fn.startswith("~$"): continue
        low=fn.lower()
        if not low.endswith(".xlsx"): continue
        if not low.startswith(MONTH.lower()): continue
        if "secondary sales analysis" not in low: continue
        path=os.path.join(SEC_DIR, fn)
        try:
            wb=openpyxl.load_workbook(path, data_only=True, read_only=True)
        except Exception as _e:
            # This file IS the month's live analysis workbook - the name says so.
            # Skipping it means _find_prior_workbook returns None, the build
            # calls the month "fresh", seeds nothing, and the coverage guard
            # below (which compares against this same workbook) is disabled with
            # it. One upload then replaces the whole month with whatever days
            # its raw happens to cover, and the job reports success. An
            # unreadable live workbook is a stop, not a fresh month.
            print(f"ERROR: prior month workbook {fn} is unreadable: {_e}")
            print("       Refusing to rebuild: this would wipe the month's "
                  "accumulated days. Restore it from a backup and retry.")
            sys.exit(3)
        sn=next((x for x in wb.sheetnames if "COMBINED DISPATCHES" in x.upper()), None)
        n=(wb[sn].max_row if sn else -1)
        wb.close()
        if sn and n>best_n:
            best_n=n; best=path
    return best

seed_count=0
_prior_wb_path=_find_prior_workbook()
if _prior_wb_path:
    pwb=openpyxl.load_workbook(_prior_wb_path, data_only=True, read_only=True)
    psn=next(x for x in pwb.sheetnames if "COMBINED DISPATCHES" in x.upper())
    _seed_by_day={}
    for row in pwb[psn].iter_rows(min_row=2, values_only=True):
        if row[11] is None or str(row[11]).strip().lower().startswith("total"):
            continue
        try:
            d,m,y=str(row[11]).split("-"); d,m,y=int(d),int(m),int(y)
        except Exception:
            continue
        if m!=MONTH_NUM or y!=YEAR:
            continue
        _seed_by_day.setdefault(d, []).append(tuple(row))
        seed_count+=1
    pwb.close()
    for _d, _rows in _seed_by_day.items():
        _tally("SEED", _d, _rows)
        _offer_day(_d, 0.0, _rows, "SEED")
    print(f"Seeded {seed_count} prior rows from {os.path.basename(_prior_wb_path)} (COMBINED DISPATCHES)")
else:
    print("No prior month workbook to seed from (fresh month).")

for rf in raw_files:
    mtime = os.path.getmtime(rf)
    wb = openpyxl.load_workbook(rf, data_only=True)
    ws = wb["SupplierSalesAnalysisReport"]
    _file_by_day = {}
    rows_read = 0
    rows_kept_for_month = 0
    min_day = 99; max_day = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[11] is None or str(row[11]).strip().lower() == "total":
            continue
        rows_read += 1
        inv_date = str(row[11])
        try:
            d, m, y = inv_date.split("-")
            d, m, y = int(d), int(m), int(y)
        except Exception:
            continue
        if m != MONTH_NUM or y != YEAR:
            continue
        rows_kept_for_month += 1
        min_day = min(min_day, d); max_day = max(max_day, d)
        _file_by_day.setdefault(d, []).append(tuple(row))
    for _d, _rows in _file_by_day.items():
        _tally(os.path.basename(rf), _d, _rows)
        _offer_day(_d, mtime, _rows, os.path.basename(rf))
    source_meta.append({
        "file": os.path.basename(rf),
        "date_range": f"{MONTH} {min_day}–{max_day}, {YEAR}" if min_day != 99 else "(none)",
        "rows_read": rows_read,
        "mtime": mtime,
    })

retained_rows = [r for _mt, _rows, _src in day_store.values() for r in _rows]
print(f"Retained after day-level merge: {len(retained_rows)}")

# ---------------------------------------------------------------------------
# RAW-VS-RETAINED RECONCILIATION GUARD (locked 12 Jun 2026, Abhay-approved).
# Every day in the final merge must contain EXACTLY the rows its winning
# source supplied -- same row count, same Issue-Cases sum, recounted
# independently from retained_rows. Permanent tripwire for the 11-Jun-2026
# silent-row-loss bug class (raw read 151 rows, only 147 retained, total 15 cs
# short vs the developer's report). ANY mismatch aborts BEFORE the live
# workbook is touched.
_guard_fail = []
_ret_tally = {}
for _r in retained_rows:
    try:
        _dd = int(str(_r[11]).split("-")[0])
    except Exception:
        _guard_fail.append(f"retained row with unparseable date: {_r[11]!r}")
        continue
    _t = _ret_tally.setdefault(_dd, [0, 0.0])
    _t[0] += 1
    try: _t[1] += float(_r[12] or 0)
    except Exception: pass
for _d in sorted(day_store):
    _mt, _rows, _src = day_store[_d]
    _exp = _read_tally.get((_src, _d))
    _got = _ret_tally.get(_d, [0, 0.0])
    if _exp is None:
        _guard_fail.append(f"day {_d}: no read tally for winning source {_src}")
    elif _got[0] != _exp[0] or abs(_got[1] - _exp[1]) > 0.001:
        _guard_fail.append(
            f"day {_d} (source {_src}): source supplied {_exp[0]} rows / "
            f"{_exp[1]:.2f} cs but retained has {_got[0]} rows / {_got[1]:.2f} cs")
if sorted(_ret_tally) != sorted(day_store):
    _guard_fail.append(f"day sets differ: retained {sorted(_ret_tally)} vs merged {sorted(day_store)}")
if _guard_fail:
    print("RECONCILIATION GUARD FAILED -- aborting, live workbook NOT touched:")
    for _f in _guard_fail:
        print("  X", _f)
    raise SystemExit(1)
print(f"Reconciliation guard OK: {len(day_store)} day(s); every winning source's rows fully retained.")

# Per-file kept counts for the SOURCE FILES sheet (days this file won).
_kept_by_src = {}
for _d, (_mt, _rows, _src) in day_store.items():
    _kept_by_src[_src] = _kept_by_src.get(_src, 0) + len(_rows)
for _m in source_meta:
    _m["rows_kept"] = _kept_by_src.get(_m["file"], 0)

# Zero-dispatch raw (e.g. a lone dry-day export with only a header row): there is
# NO month data on disk to rebuild from, so overwriting the live workbook would
# wipe it. Treat as a clean no-op -- exit WITHOUT saving and let the operator
# promote / keep the existing workbook. A genuine cumulative raw always carries
# data, so this branch only fires on an empty single-day drop.
if not retained_rows:
    raise SystemExit(
        "ZERO-DISPATCH RAW: 0 dispatch rows for " + MONTH + ".\n"
        "This is a valid zero-dispatch day, but there is no month data in the raw to\n"
        "rebuild from. The existing workbook is left UNTOUCHED (not overwritten).\n"
        "If this is month-end, promote the live workbook to its full-month name."
    )

# Period END = the later of (last day with dispatches) and (window end parsed from
# the filename). This keeps trailing empty days (incl. a dry day 1) inside the
# labelled range, so a '1-13th' export with an empty day 13 still reads 1-13.
data_max = max(int(str(r[11]).split("-")[0]) for r in retained_rows)
latest_day = max(data_max, window_end)
print(f"latest_day = {latest_day} (data_max={data_max}, window_end={window_end})")

# =========================================================================
# Coverage guard — abort if raw covers fewer days than the prior live workbook
# =========================================================================
# Prevents the April 30 regression where a single-day incremental raw
# (only Apr 29) was rebuilt over a month-to-date workbook (Apr 1-28),
# wiping all prior days. If the prior live workbook's COMBINED DISPATCHES
# contains any date NOT in the new build, abort with instructions.

new_dates = {str(r[11]) for r in retained_rows}


def _read_dates(path):
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        for sn in wb.sheetnames:
            if "COMBINED DISPATCHES" in sn.upper():
                ws = wb[sn]
                ds = set()
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row[11] is None:
                        continue
                    if str(row[11]).strip().lower().startswith("total"):
                        continue
                    ds.add(str(row[11]))
                return ds
    except Exception as _e:
        print(f"WARN: could not read prior dates from {os.path.basename(path)}: {_e}")
    return set()


# SEED / COVERAGE GUARD (audit fix #3 — now a REAL post-seed assertion).
# The seed injected the prior live workbook's COMBINED dates into retained_rows.
# Verify every prior date actually made it into the final build, sourced from the
# SAME workbook the seed used (_prior_wb_path). This catches a silent partial
# seed-load (corrupt/locked prior) OR a raw that shrinks coverage BEFORE the live
# workbook is touched. Previously it re-scanned for a 'best' prior workbook, which
# was tautological against the seed and could never fire.
if _prior_wb_path:
    prior_dates = _read_dates(_prior_wb_path)
    missing = prior_dates - new_dates
    if missing:
        missing_sorted = sorted(missing)
        raise SystemExit(
            "\n=== SEED/COVERAGE GUARD ABORTED THE BUILD ===\n"
            f"Seed source workbook: {os.path.basename(_prior_wb_path)}\n"
            f"  prior dates ({len(prior_dates)}): {sorted(prior_dates)[0]} -> {sorted(prior_dates)[-1]}\n"
            f"  new build dates ({len(new_dates)}): {sorted(new_dates)[0] if new_dates else '(none)'} -> {sorted(new_dates)[-1] if new_dates else '(none)'}\n"
            f"  MISSING from new build ({len(missing)}): {missing_sorted}\n\n"
            "The prior live workbook covered days the new build does not — the seed\n"
            "load was partial/corrupt OR the raw shrinks coverage. Rebuilding now\n"
            "would WIPE those days from COMBINED DISPATCHES.\n\n"
            "To proceed, do ONE of:\n"
            "  (a) Repair/re-export the prior workbook so the seed loads every day, or\n"
            "  (b) Drop the missing days' raw(s) alongside the new one so the build sees them.\n"
            "Then rerun build_secondary.py.\n"
        )
    else:
        print(f"Seed/coverage guard OK: all {len(prior_dates)} prior date(s) present in the new build "
              f"(+{len(new_dates - prior_dates)} new).")
else:
    print("Seed/coverage guard: no prior same-month workbook (fresh month) — nothing to assert.")

# =========================================================================
# Step 2: Load master# =========================================================================
# Step 2: Load master
# =========================================================================

mwb = openpyxl.load_workbook(MASTER, data_only=True)
mws = mwb["16-4-25"]

master = {}  # shop_code (int or str, normalized) -> dict
for row in mws.iter_rows(min_row=3, values_only=True):
    if row[0] is None: continue
    slno, wcode, whname, shop_code, shop_name, cat, staff, bond, status = row[:9]
    if cat == "CFD":
        cat = "FED"
    if shop_code is None or str(shop_code).strip() == "":
        # Blank shop code (3 BAR placeholder rows) — skip so they don't all
        # collapse to a single master["None"] entry (audit fix #6).
        continue
    code_str = str(shop_code).strip()
    master[code_str] = {
        "slno": slno, "w_code": wcode, "wh_name": whname,
        "shop_code": shop_code, "shop_name": shop_name,
        "cat": cat, "staff": staff, "bond": bond, "status": status,
    }

print(f"Master entries: {len(master)}")

# =========================================================================
# Step 3: Aggregate
# =========================================================================

outlet_agg = defaultdict(lambda: {"cases":0.0, "bottles":0, "name":None})
bond_cat = defaultdict(lambda: defaultdict(float))  # bond -> cat -> cases
brand_cat = defaultdict(lambda: defaultdict(float))  # brand -> cat -> cases
pack_cat = defaultdict(lambda: defaultdict(float))
day_cases = defaultdict(float)
unmatched = defaultdict(lambda: {"name": None, "cases": 0.0})

grand = defaultdict(float)
invoices_set = set()

for row in retained_rows:
    (wh_code, wh_name, supplier, prod_code, item_name, pack,
     lic_no, lic_name, indent_no, indent_date, inv_no, inv_date,
     issue_cases, issue_bottles) = row

    try:
        cases = float(issue_cases) if issue_cases not in (None, "") else 0.0
    except:
        cases = 0.0
    try:
        bottles = int(float(issue_bottles)) if issue_bottles not in (None, "") else 0
    except:
        bottles = 0
    # Fold loose bottles into case-decimal (audit fix #4 — they were dropped from
    # every case total). Pack strings match the BPC keys exactly.
    _bpc = BPC.get(str(pack).strip().upper())
    if bottles and _bpc:
        cases += bottles / _bpc
    elif bottles and not _bpc:
        print(f"WARN: {bottles} loose bottle(s) for pack {pack!r} with no BPC mapping — not folded (undercount).")

    code_str = str(lic_no).strip() if lic_no not in (None, "") else "(blank)"
    invoices_set.add(inv_no)

    # Outlet
    outlet_agg[code_str]["cases"] += cases
    outlet_agg[code_str]["bottles"] += bottles
    outlet_agg[code_str]["name"] = lic_name

    # Master lookup
    info = master.get(code_str)
    if info is None:
        unmatched[code_str]["name"] = lic_name
        unmatched[code_str]["cases"] += cases
        cat = "BAR"  # classify as BAR for totals
        bond = None
    else:
        cat = info["cat"]
        bond = info["bond"]

    grand[cat] += cases
    grand["TOTAL"] += cases

    if bond:
        bond_cat[bond][cat] += cases

    brand_cat[item_name][cat] += cases
    pack_cat[pack][cat] += cases

    # Day
    d, m, y = str(inv_date).split("-")
    day_cases[int(d)] += cases

print(f"Grand KSBC={grand['KSBC']:.2f} FED={grand['FED']:.2f} BAR={grand['BAR']:.2f} TOTAL={grand['TOTAL']:.2f}")
print(f"Invoice cashflow (FED+BAR): {grand['FED']+grand['BAR']:.2f}")
print(f"Unmatched outlets: {len(unmatched)}")
for code, u in sorted(unmatched.items(), key=lambda x: -x[1]['cases']):
    print(f"   {code} {u['name']} : {u['cases']:.2f}")

# =========================================================================
# Step 4: Build workbook
# =========================================================================

out = Workbook()
# Remove default sheet
out.remove(out.active)

# =========================================================================
# Aggregate totals + per-bond rollup. The DASHBOARD sheet itself is built by
# restyle_dashboard.py (invoked automatically at the end of this script —
# audit fix #2). The old build-side light dashboard with rejected labels
# ("TOTAL DISPATCHED" / "OUTLETS INVOICED" ...) was removed so it can never
# ship if restyle is skipped.
# =========================================================================
total = grand["TOTAL"]
cashflow = grand["FED"] + grand["BAR"]
grand_k = grand["KSBC"]; grand_f = grand["FED"]; grand_b_inc_u = grand["BAR"]

all_bonds = ["KOLLAM","KOZHIKODE","ATTINGAL","PALAKKAD","KOTTARAKARA","KANNUR","ALAPPUZHA",
             "NEDUMANGAD","ALUVA","PERINTHALMANNA","THODUPUZHA","PATHANAMTHITTA",
             "KOTTAYAM","THRISSUR","TRIPUNITHURA"]

bond_totals = []
for bond in all_bonds:
    k = bond_cat[bond]["KSBC"]
    f = bond_cat[bond]["FED"]
    b = bond_cat[bond]["BAR"]
    tot = k + f + b
    bond_totals.append({"bond": bond, "k": k, "f": f, "b": b, "tot": tot, "cash": f + b})

bond_totals.sort(key=lambda x: -x["tot"])

# ---- BOND PERFORMANCE ----
bp = out.create_sheet("BOND PERFORMANCE")
bp.sheet_view.showGridLines = False

bp.merge_cells("A1:H1")
c = bp["A1"]
c.value = f"BOND-WISE SECONDARY SALES — {MONTH} {YEAR}"
title_style(c)
bp.row_dimensions[1].height = 26

bp_headers = ["Bond", "KSBC (cases)", "Consumer fed", "BAR", "Total", "KSBC %", "FED %", "BAR %"]
for i, h in enumerate(bp_headers):
    cell = bp.cell(row=3, column=i+1)
    cell.value = h
    header_style(cell)
bp.row_dimensions[3].height = 24

r = 4
for i, bd in enumerate(bond_totals):
    k, f, b, tot = bd["k"], bd["f"], bd["b"], bd["tot"]
    kp = k/tot if tot else 0
    fp = f/tot if tot else 0
    bp_pct = b/tot if tot else 0
    vals = [bd["bond"], k, f, b, tot, kp, fp, bp_pct]
    for col, v in enumerate(vals, 1):
        cell = bp.cell(row=r, column=col)
        cell.value = v
        data_style(cell, alt=(i%2==0), numeric=(col!=1))
        if col in (2,3,4,5):
            cell.number_format = "#,##0.00"
        if col in (6,7,8):
            cell.number_format = "0.0%"
    r += 1

# Unmatched row
if unmatched:
    un_cases = sum(u["cases"] for u in unmatched.values())
    vals = ["UNMATCHED / Pending master", 0, 0, un_cases, un_cases, 0, 0, 1.0]
    for col, v in enumerate(vals, 1):
        cell = bp.cell(row=r, column=col)
        cell.value = v
        cell.fill = PatternFill("solid", fgColor=BAR_FILL)
        cell.font = Font(name=FONT_NAME, size=10, italic=True, color=TEXT_DARK)
        cell.alignment = Alignment(horizontal=("right" if col!=1 else "left"), vertical="center")
        cell.border = thin_border
        if col in (2,3,4,5):
            cell.number_format = "#,##0.00"
        if col in (6,7,8):
            cell.number_format = "0.0%"
    r += 1

# TOTAL
tot_row = ["TOTAL", grand_k, grand_f, grand_b_inc_u, total,
           grand_k/total if total else 0,
           grand_f/total if total else 0,
           grand_b_inc_u/total if total else 0]
for col, v in enumerate(tot_row, 1):
    cell = bp.cell(row=r, column=col)
    cell.value = v
    total_style(cell, numeric=(col!=1))
    if col in (2,3,4,5):
        cell.number_format = "#,##0.00"
    if col in (6,7,8):
        cell.number_format = "0.0%"

bp.column_dimensions["A"].width = 22
for cw, w in zip("BCDEFGH", [14, 14, 12, 14, 10, 10, 10]):
    bp.column_dimensions[cw].width = w

# ---- 15 Region sheets ----
# Build per-bond outlet list from master
# Include ALL master outlets (any status) so a dispatched Closed/inactive shop
# still appears on its region sheet and the region TOTAL reconciles to BOND
# PERFORMANCE (audit fix #5). Active shops always show; non-Active only if they
# actually dispatched (filtered in the per-bond loop below).
bond_outlets = defaultdict(list)
for code, info in master.items():
    bond_outlets[info["bond"]].append((code, info))

cat_order = {"KSBC": 0, "FED": 1, "BAR": 2}
grand_region_total = 0.0

for bond in all_bonds:
    ws = out.create_sheet(bond)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"

    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value = f"{bond} — SECONDARY SALES ({MONTH} {YEAR})"
    title_style(c)
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:E2")
    c = ws["A2"]
    c.value = "All figures in cases. Grouped by outlet type."
    c.fill = PatternFill("solid", fgColor=LIGHT_BG)
    c.font = Font(name=FONT_NAME, size=10, italic=True, color=TEXT_DARK)
    c.alignment = Alignment(horizontal="center", vertical="center")

    headers = ["Shop Code", "Outlet Name", "Field Staff", "CAT", "Secondary Sales"]
    for i, h in enumerate(headers):
        cell = ws.cell(row=4, column=i+1)
        cell.value = h
        header_style(cell)
    ws.row_dimensions[4].height = 24

    outlets = bond_outlets[bond]
    # Aggregate cases per outlet for this bond. Show every Active outlet; show a
    # non-Active (Closed) outlet only when it has dispatches (audit fix #5).
    enriched = []
    for code, info in outlets:
        cases = outlet_agg.get(code, {}).get("cases", 0.0)
        if info["status"] != "Active" and cases <= 0:
            continue
        enriched.append((code, info, cases))
    # Sort: KSBC → FED → BAR, then cases desc
    enriched.sort(key=lambda x: (cat_order.get(x[1]["cat"], 99), -x[2]))

    r = 5
    total_bond = 0.0
    for idx, (code, info, cases) in enumerate(enriched):
        vals = [code, info["shop_name"], info["staff"], info["cat"], cases]
        cat_fill = {"KSBC": KSBC_FILL, "FED": FED_FILL, "BAR": BAR_FILL}.get(info["cat"], WHITE)
        for col, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=col)
            cell.value = v
            data_style(cell, alt=(idx%2==0), numeric=(col==5 or col==1))
            if col == 1:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if col == 4:
                cell.fill = PatternFill("solid", fgColor=cat_fill)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(name=FONT_NAME, size=10, bold=True, color=TEXT_DARK)
            if col == 5:
                cell.number_format = "#,##0.00"
        total_bond += cases
        r += 1

    # TOTAL
    vals = ["", "TOTAL", "", "", total_bond]
    for col, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=col)
        cell.value = v
        total_style(cell, numeric=(col==5))
        if col == 2:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        if col == 5:
            cell.number_format = "#,##0.00"

    grand_region_total += total_bond

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 16

# Region reconciliation tripwire (audit fix #5): the 15 region sheets list every
# Active outlet plus any dispatched non-Active one, so Sigma region TOTALs +
# unmatched cases must equal the grand total. A mismatch means a dispatched outlet
# fell through the cracks — abort before writing the live workbook.
_un_total = sum(u["cases"] for u in unmatched.values())
_recon = grand_region_total + _un_total
if abs(_recon - total) > 0.01:
    raise SystemExit(
        "\n=== REGION RECONCILIATION GUARD ABORTED THE BUILD ===\n"
        f"  region TOTALs sum = {grand_region_total:.2f}\n"
        f"  + unmatched       = {_un_total:.2f}\n"
        f"  = {_recon:.2f}  but grand total = {total:.2f}  (diff {_recon - total:+.2f})\n"
        "A dispatched outlet is missing from the region sheets. Investigate before saving.\n"
    )
print(f"Region reconciliation OK: region {grand_region_total:.2f} + unmatched {_un_total:.2f} = {total:.2f}")

# ---- BRAND PERFORMANCE ----
bpf = out.create_sheet("BRAND PERFORMANCE")
bpf.sheet_view.showGridLines = False
bpf.merge_cells("A1:F1")
c = bpf["A1"]; c.value = f"BRAND-WISE SECONDARY SALES — {MONTH} {YEAR}"; title_style(c)
bpf.row_dimensions[1].height = 26

bhdrs = ["Brand", "KSBC", "Consumer fed", "BAR", "Total", "% of Grand Total"]
for i,h in enumerate(bhdrs):
    cell = bpf.cell(row=3, column=i+1); cell.value = h; header_style(cell)
bpf.row_dimensions[3].height = 24

brand_rows = []
for brand, cats in brand_cat.items():
    k = cats.get("KSBC", 0); f = cats.get("FED", 0); b = cats.get("BAR", 0); tot = k+f+b
    brand_rows.append((brand, k, f, b, tot))
brand_rows.sort(key=lambda x: -x[4])

r = 4
for i, (brand, k, f, b, tot) in enumerate(brand_rows):
    vals = [brand, k, f, b, tot, (tot/total if total else 0)]
    for col, v in enumerate(vals, 1):
        cell = bpf.cell(row=r, column=col); cell.value = v
        data_style(cell, alt=(i%2==0), numeric=(col!=1))
        if col in (2,3,4,5): cell.number_format = "#,##0.00"
        if col == 6: cell.number_format = "0.0%"
    r += 1

vals = ["TOTAL", grand_k, grand_f, grand_b_inc_u, total, 1.0]
for col, v in enumerate(vals, 1):
    cell = bpf.cell(row=r, column=col); cell.value = v; total_style(cell, numeric=(col!=1))
    if col in (2,3,4,5): cell.number_format = "#,##0.00"
    if col == 6: cell.number_format = "0.0%"

bpf.column_dimensions["A"].width = 42
for cw, w in zip("BCDEF", [12,14,12,12,14]):
    bpf.column_dimensions[cw].width = w

# ---- PACK PERFORMANCE ----
pp = out.create_sheet("PACK PERFORMANCE")
pp.sheet_view.showGridLines = False
pp.merge_cells("A1:F1")
c = pp["A1"]; c.value = f"PACK-WISE SECONDARY SALES — {MONTH} {YEAR}"; title_style(c)
pp.row_dimensions[1].height = 26
phdrs = ["Pack", "KSBC", "Consumer fed", "BAR", "Total", "% of Grand Total"]
for i,h in enumerate(phdrs):
    cell = pp.cell(row=3, column=i+1); cell.value = h; header_style(cell)
pp.row_dimensions[3].height = 24

pack_rows = []
for pack, cats in pack_cat.items():
    k = cats.get("KSBC", 0); f = cats.get("FED", 0); b = cats.get("BAR", 0); tot = k+f+b
    pack_rows.append((pack, k, f, b, tot))
pack_rows.sort(key=lambda x: -x[4])

r = 4
for i, (pack, k, f, b, tot) in enumerate(pack_rows):
    vals = [pack, k, f, b, tot, (tot/total if total else 0)]
    for col, v in enumerate(vals, 1):
        cell = pp.cell(row=r, column=col); cell.value = v
        data_style(cell, alt=(i%2==0), numeric=(col!=1))
        if col in (2,3,4,5): cell.number_format = "#,##0.00"
        if col == 6: cell.number_format = "0.0%"
    r += 1

vals = ["TOTAL", grand_k, grand_f, grand_b_inc_u, total, 1.0]
for col, v in enumerate(vals, 1):
    cell = pp.cell(row=r, column=col); cell.value = v; total_style(cell, numeric=(col!=1))
    if col in (2,3,4,5): cell.number_format = "#,##0.00"
    if col == 6: cell.number_format = "0.0%"

pp.column_dimensions["A"].width = 20
for cw, w in zip("BCDEF", [12,14,12,12,14]):
    pp.column_dimensions[cw].width = w

# ---- DAILY TREND ----
dt = out.create_sheet("DAILY TREND")
dt.sheet_view.showGridLines = False
dt.merge_cells("A1:C1")
c = dt["A1"]; c.value = f"DAILY DISPATCH TREND — {MONTH} {YEAR}"; title_style(c)
dt.row_dimensions[1].height = 26
dhdrs = ["Date", "Day of Week", "Cases Invoiced"]
for i,h in enumerate(dhdrs):
    cell = dt.cell(row=3, column=i+1); cell.value = h; header_style(cell)
dt.row_dimensions[3].height = 24

# Days in target month (handles 28/29/30/31)
import calendar as _cal
days_in_month = _cal.monthrange(YEAR, MONTH_NUM)[1]
r = 4
daily_total = 0.0
for d in range(1, days_in_month+1):
    dt_obj = date(YEAR, MONTH_NUM, d)
    dow = dt_obj.strftime("%A")
    cases = day_cases.get(d, 0.0)
    daily_total += cases
    vals = [f"{d:02d}-{MONTH_NUM:02d}-{YEAR}", dow, cases]
    for col, v in enumerate(vals, 1):
        cell = dt.cell(row=r, column=col); cell.value = v
        data_style(cell, alt=((d-1)%2==0), numeric=(col==3))
        if col == 3: cell.number_format = "#,##0.00"
    r += 1

vals = ["TOTAL", "", daily_total]
for col, v in enumerate(vals, 1):
    cell = dt.cell(row=r, column=col); cell.value = v; total_style(cell, numeric=(col==3))
    if col == 3: cell.number_format = "#,##0.00"

dt.column_dimensions["A"].width = 14
dt.column_dimensions["B"].width = 14
dt.column_dimensions["C"].width = 18

# ---- COMBINED DISPATCHES ----
cd = out.create_sheet(f"{MONTH} COMBINED DISPATCHES")
cd.sheet_view.showGridLines = False
raw_headers = ['Warehouse Code', 'Warehouse Name', 'Supplier Name', 'Product Code', 'Item Name',
               'Pack', 'Licensee No.', 'Licensee Name', 'Indent No', 'Indent Approve Date',
               'Inv/GTN No.', 'Inv/GTN Date', 'Issue Cases', 'Issue Bottles']
for i,h in enumerate(raw_headers):
    cell = cd.cell(row=1, column=i+1); cell.value = h; header_style(cell)
cd.row_dimensions[1].height = 24
cd.freeze_panes = "A2"  # keep header visible on this long sheet (audit fix #12)
# Sort retained rows by date then invoice
def row_sortkey(row):
    d, m, y = str(row[11]).split("-")
    return (int(y), int(m), int(d), str(row[10]))
retained_sorted = sorted(retained_rows, key=row_sortkey)
for i, row in enumerate(retained_sorted, 2):
    for col, v in enumerate(row, 1):
        cell = cd.cell(row=i, column=col)
        # Issue Cases (13) / Issue Bottles (14): write real numbers, not text
        # (audit fix #7). Dates stay DD-MM-YYYY text — the seed parser here and
        # downstream rely on that string format.
        if col == 13:
            try: cell.value = float(v) if v not in (None, "") else 0.0
            except Exception: cell.value = v
        elif col == 14:
            try: cell.value = int(float(v)) if v not in (None, "") else 0
            except Exception: cell.value = v
        else:
            cell.value = v
        data_style(cell, alt=(i%2==0), numeric=(col in (13,14)))
        if col == 13: cell.number_format = "#,##0.00"
        if col == 14: cell.number_format = "#,##0"

widths = [10, 28, 30, 12, 40, 10, 10, 32, 22, 14, 24, 14, 10, 10]
for i, w in enumerate(widths, 1):
    cd.column_dimensions[get_column_letter(i)].width = w

# ---- UNMATCHED ----
if unmatched:
    um = out.create_sheet("UNMATCHED")
    um.sheet_view.showGridLines = False
    um.merge_cells("A1:C1")
    c = um["A1"]
    c.value = "⚠ Licensees in raw data but not in MASTER DATA CONFIRMED"
    title_style(c)
    um.row_dimensions[1].height = 26
    um.merge_cells("A2:C2")
    c = um["A2"]
    c.value = "These licensees need to be added to master. Cases ARE counted in grand totals but not bond-mapped."
    c.fill = PatternFill("solid", fgColor=LIGHT_BG)
    c.font = Font(name=FONT_NAME, size=10, italic=True, color=TEXT_DARK)
    c.alignment = Alignment(horizontal="center", vertical="center")

    for i,h in enumerate(["Licensee Code", "Licensee Name", "Cases"]):
        cell = um.cell(row=4, column=i+1); cell.value = h; header_style(cell)
    um.row_dimensions[4].height = 24

    rows = sorted(unmatched.items(), key=lambda x: -x[1]["cases"])
    r = 5
    for idx, (code, u) in enumerate(rows):
        vals = [code, u["name"], u["cases"]]
        for col, v in enumerate(vals, 1):
            cell = um.cell(row=r, column=col); cell.value = v
            data_style(cell, alt=(idx%2==0), numeric=(col==3))
            if col == 3: cell.number_format = "#,##0.00"
        r += 1

    um.column_dimensions["A"].width = 16
    um.column_dimensions["B"].width = 42
    um.column_dimensions["C"].width = 14

# ---- SOURCE FILES ----
sf = out.create_sheet("SOURCE FILES")
sf.sheet_view.showGridLines = False
sf.merge_cells("A1:E1")
c = sf["A1"]; c.value = "SOURCE FILES — raw data combined into this analysis"; title_style(c)
sf.row_dimensions[1].height = 26
for i,h in enumerate(["File Name","Date Range Covered","Rows Read","Rows Retained","Last Modified"]):
    cell = sf.cell(row=3, column=i+1); cell.value = h; header_style(cell)
sf.row_dimensions[3].height = 24

r = 4
total_read = 0
for m in source_meta:
    dt_mod = datetime.fromtimestamp(m["mtime"]).strftime("%Y-%m-%d %H:%M:%S")
    vals = [m["file"], m["date_range"], m["rows_read"], m.get("rows_kept", ""), dt_mod]
    total_read += m["rows_read"]
    for col, v in enumerate(vals, 1):
        cell = sf.cell(row=r, column=col); cell.value = v
        data_style(cell, alt=((r-4)%2==0), numeric=(col==3))
    r += 1

# Summary row showing total after dedupe
vals = ["TOTAL (retained)", f"{MONTH} 1–{latest_day}", total_read, len(retained_rows), ""]
for col, v in enumerate(vals, 1):
    cell = sf.cell(row=r, column=col); cell.value = v; total_style(cell, numeric=(col in (3,4)))

sf.column_dimensions["A"].width = 46
sf.column_dimensions["B"].width = 22
sf.column_dimensions["C"].width = 12
sf.column_dimensions["D"].width = 14
sf.column_dimensions["E"].width = 22

# ---- MASTER DATA ----
md = out.create_sheet("MASTER DATA")
md.sheet_view.showGridLines = False
# Copy values + styles from master
src = mws
for row in src.iter_rows():
    for cell in row:
        new = md.cell(row=cell.row, column=cell.column, value=cell.value)
        if cell.has_style:
            _of = cell.font
            new.font = Font(name=FONT_NAME, size=_of.size, bold=_of.bold,
                            italic=_of.italic, color=_of.color, underline=_of.underline,
                            strike=_of.strike, vertAlign=_of.vertAlign)
            new.fill = copy(cell.fill)
            new.border = copy(cell.border)
            new.alignment = copy(cell.alignment)
            new.number_format = cell.number_format
            new.protection = copy(cell.protection)

# Copy column widths
for col_letter, col_dim in src.column_dimensions.items():
    md.column_dimensions[col_letter].width = col_dim.width
# Copy row heights
for row_num, row_dim in src.row_dimensions.items():
    md.row_dimensions[row_num].height = row_dim.height
# Copy merged cells
for mr in src.merged_cells.ranges:
    md.merge_cells(str(mr))
# Freeze panes
md.freeze_panes = src.freeze_panes

# Save
out.save(SCRATCH)
print(f"\nSaved scratch: {SCRATCH}")
print(f"Sheets: {out.sheetnames}")

# Print totals for verification
print(f"\n=== VERIFICATION TOTALS ===")
print(f"grand_total = {total:.2f}")
print(f"invoice_cashflow = {cashflow:.2f}")
print(f"Sum bond_totals = {sum(bd['tot'] for bd in bond_totals):.2f}")
print(f"grand - (sum bond totals) = {total - sum(bd['tot'] for bd in bond_totals):.2f} (should == unmatched)")
print(f"Brand sum = {sum(r[4] for r in brand_rows):.2f}")
print(f"Pack sum = {sum(r[4] for r in pack_rows):.2f}")
print(f"Daily sum = {daily_total:.2f}")

# Top items
print(f"\nTop bond by total: {bond_totals[0]['bond']} = {bond_totals[0]['tot']:.2f}")
bond_by_cash = sorted(bond_totals, key=lambda x: -x["cash"])
print(f"Top bond by cashflow: {bond_by_cash[0]['bond']} = {bond_by_cash[0]['cash']:.2f}")
print(f"Top brand: {brand_rows[0][0]} = {brand_rows[0][4]:.2f}")
print(f"Top pack: {pack_rows[0][0]} = {pack_rows[0][4]:.2f}")
peak_day = max(day_cases.items(), key=lambda x: x[1])
print(f"Peak day: {peak_day[0]} = {peak_day[1]:.2f}")

# Month-complete notice (audit fix #13): if the final calendar day is ingested,
# remind the operator to promote the workbook to the clean full-month name.
_days_in_month = _cal.monthrange(YEAR, MONTH_NUM)[1]
if latest_day >= _days_in_month:
    print(f"\n*** MONTH COMPLETE: {MONTH} {YEAR} fully ingested (day {latest_day} of {_days_in_month}). "
          f"Promote the saved workbook to full-month name '{MONTH} SECONDARY SALES ANALYSIS.xlsx'. ***")

# =========================================================================
# Build the DASHBOARD via restyle_dashboard.py (audit fix #2) so a single run
# of this script ALWAYS yields the locked dark dashboard. Fail loudly on error
# and verify the locked layout before declaring success — never ship without it.
# =========================================================================
import subprocess
_restyle = os.path.join(os.path.dirname(_THIS), "restyle_dashboard.py")
print("\nRunning restyle_dashboard.py to build the locked DASHBOARD ...")
_res = subprocess.run([sys.executable, _restyle], capture_output=True, text=True)
if _res.stdout:
    print(_res.stdout)
if _res.returncode != 0:
    print(_res.stderr)
    raise SystemExit("restyle_dashboard.py FAILED — DASHBOARD not built; scratch NOT fit to promote.")
_chk = openpyxl.load_workbook(SCRATCH)
_ok = ("DASHBOARD" in _chk.sheetnames and _chk.sheetnames[0] == "DASHBOARD")
_b7 = _chk["DASHBOARD"]["B7"].value if "DASHBOARD" in _chk.sheetnames else None
_chk.close()
if not _ok or _b7 != "TOTAL SECONDARY SALES":
    raise SystemExit(
        f"Post-restyle check FAILED: DASHBOARD present={_ok}, B7={_b7!r} "
        "(expected first sheet 'DASHBOARD' with KPI 'TOTAL SECONDARY SALES'). "
        "The locked dark dashboard did not render.")
print("DASHBOARD locked-layout check OK (dark dashboard is first sheet, KPI labels correct).")