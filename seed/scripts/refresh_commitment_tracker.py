#!/usr/bin/env python3
"""
Refresh the K.S. Distillery Commitment Tracker.

What it does in one pass:
  1. Opens   Commitment tracking/COMMITMENT TRACKER - <MONTH>.xlsx
  2. For each bond tab, reads the meeting date (cell C3) and every commit row
     (shop_code, brand, pack, target_cases).
  3. Opens the current KSBC daily-sales workbook in  KSBC shop sales/
     (auto-discovers MAY 1st - Nth ANALYSIS.xlsx or MAY SHOP SALES ANALYSIS.xlsx),
     parses the daily sheets (MAY 1, MAY 2, ...), builds a fast lookup
     { (shop_code, brand, pack) : { date : cases_sold } }.
  4. For every commit row in every bond tab:
        sold  = SUM(cases_sold across dates > aging-snapshot date)  # col-G baseline date, not meeting date
        days  = meeting_date + REVIEW_CYCLE_DAYS - today
     writes Sold into col K and Days remaining into col M.
  5. Saves the workbook (formulas in cols L and N recompute on open).
  6. Builds the live-artifact JSON payload, writes
        outputs/commitment_tracker_payload.json    + summary
        outputs/commitment_tracker_live.html       full standalone artifact
     and prints a console summary.

Manual trigger:
    python3 refresh_commitment_tracker.py [--as-of YYYY-MM-DD]
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl

REVIEW_CYCLE_DAYS = 30   # fallback only — if the Next Meeting Date cell (G3) is
                         # empty on a bond tab, we assume the next meeting is this
                         # many days after the meeting date. Filling G3 explicitly
                         # overrides this fallback per-bond.

# Auto-detect the Claude data folder from this file's location
#   <claude>/.claude/scripts/refresh_commitment_tracker.py
_THIS = os.path.abspath(__file__)
CLAUDE_DIR  = Path(os.path.abspath(os.path.join(os.path.dirname(_THIS), "..", "..")))
AGING_DIR   = CLAUDE_DIR / "Commitment tracking"
KSBC_DIR    = CLAUDE_DIR / "KSBC shop sales"
OUTPUTS_DIR = AGING_DIR / ".artifact"

# Status tiers
STATUS_ACHIEVED       = "✅ Achieved"
STATUS_ON_TRACK       = "🚀 On track"
STATUS_AT_RISK        = "⚠ At risk"
STATUS_BEHIND         = "🚫 Behind"
STATUS_AWAITING       = "Awaiting refresh"
STATUS_NO_COMMIT      = "— No commit yet"

BONDS = ["KOTTAYAM", "THODUPUZHA", "TRIPUNITHURA", "PATHANAMTHITTA", "ALUVA",
         "ATTINGAL", "THRISSUR", "KOTTARAKARA", "NEDUMANGAD", "PALAKKAD",
         "KANNUR", "KOZHIKODE", "PERINTHALMANNA", "KOLLAM", "ALAPPUZHA"]


# ----------------------------------------------------------------------
# Discovery
# ----------------------------------------------------------------------

def find_commitment_workbook():
    """Find the latest COMMITMENT TRACKER - <MONTH>.xlsx."""
    candidates = sorted(AGING_DIR.glob("COMMITMENT TRACKER - *.xlsx"),
                        key=lambda p: p.stat().st_mtime, reverse=True)
    candidates = [p for p in candidates if not p.name.startswith("~$")]
    if not candidates:
        raise SystemExit(f"No COMMITMENT TRACKER workbook in {AGING_DIR}")
    return candidates[0]


def find_ksbc_workbook(month_hint: str = None):
    """Find the most recent KSBC daily-sales workbook for the current month.
    Prefer the full-month file (`MAY SHOP SALES ANALYSIS.xlsx`) if present,
    otherwise the latest mid-month (`MAY 1st - Nth ANALYSIS.xlsx`)."""
    if month_hint:
        m = month_hint.upper()
    else:
        m = datetime.today().strftime("%B").upper()

    # Full-month
    full = KSBC_DIR / f"{m} SHOP SALES ANALYSIS.xlsx"
    if full.exists():
        return full
    # Mid-month — newest by mtime
    mid = sorted(KSBC_DIR.glob(f"{m} 1st - * ANALYSIS.xlsx"),
                 key=lambda p: p.stat().st_mtime, reverse=True)
    mid = [p for p in mid if not p.name.startswith("~$")]
    if mid:
        return mid[0]
    # Try previous month as fallback
    # Walk BACK through prior months rather than aborting. On the 1st of a new
    # month (or any day before that month's first raw lands) there is no
    # current-month workbook yet, and this call is only the fallback "latest
    # closing source" — main() resolves the real per-month workbooks via
    # find_ksbc_workbook_for_month(). Hard-exiting here made every 1st-of-month
    # refresh fail with "No KSBC analysis workbook found for SEPTEMBER".
    # (Fix 1 Sep 2026, Abhay-requested.)
    _ORDER = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY",
              "AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"]
    if m in _ORDER:
        i = _ORDER.index(m)
        for back in range(1, 13):
            prev = _ORDER[(i - back) % 12]
            alt = find_ksbc_workbook_for_month(prev)
            if alt is not None:
                print(f"  NOTE: no {m.title()} KSBC workbook yet — "
                      f"falling back to {alt.name}")
                return alt
    raise SystemExit(f"No KSBC analysis workbook found for {m} in {KSBC_DIR}")


def find_ksbc_workbook_for_month(month_name: str):
    """Return the best KSBC workbook for a given month name, or None.
    Full-month (`MAY SHOP SALES ANALYSIS.xlsx`) preferred, else latest mid-month."""
    m = month_name.upper()
    full = KSBC_DIR / f"{m} SHOP SALES ANALYSIS.xlsx"
    if full.exists():
        return full
    mid = sorted(KSBC_DIR.glob(f"{m} 1st - * ANALYSIS.xlsx"),
                 key=lambda p: p.stat().st_mtime, reverse=True)
    mid = [p for p in mid if not p.name.startswith("~$")]
    return mid[0] if mid else None


def months_in_window(start_date, end_date):
    """List of (month_name, year) from start_date's month through end_date's
    month inclusive. Lets the commitment window span month boundaries -- e.g. a
    May-16 meeting refreshed in June must read BOTH May and June daily sheets;
    without this the May sales silently drop out (the bug Abhay flagged 3 Jun 2026)."""
    inv = {v: k for k, v in MONTHS.items()}
    out = []
    y, mo = start_date.year, start_date.month
    while (y, mo) <= (end_date.year, end_date.month):
        out.append((inv[mo], y))
        mo += 1
        if mo > 12:
            mo = 1
            y += 1
    return out


# ----------------------------------------------------------------------
# KSBC daily-sales lookup
# ----------------------------------------------------------------------

DAILY_SHEET_RE = re.compile(r"^([A-Z]+)\s+(\d{1,2})$")  # e.g. "MAY 14"
MONTHS = {"JANUARY":1,"FEBRUARY":2,"MARCH":3,"APRIL":4,"MAY":5,"JUNE":6,
          "JULY":7,"AUGUST":8,"SEPTEMBER":9,"OCTOBER":10,"NOVEMBER":11,"DECEMBER":12}

# Month-name lookup covering full ("MAY") and 3-letter abbreviated ("MAY"/"MAR") forms.
_MONTH_LOOKUP = {}
for _full, _num in MONTHS.items():
    _MONTH_LOOKUP[_full] = _num
    _MONTH_LOOKUP[_full[:3]] = _num
_ASOF_DATE_RE = re.compile(r"([A-Za-z]{3,9})\s+(\d{1,2})")


def parse_aging_snapshot_date(subtitle, year):
    """Extract the aging-snapshot ("as of") date from a bond tab's row-2 subtitle,
    e.g. 'Field Staff: X · Aging Stock as of May 14 · Generated 15 May 2026'.

    Returns the LAST '<Month> <day>' found inside the 'as of ...' segment -- so a
    range like 'as of Mar 1 to May 14' yields the closing date (May 14) -- or None
    if no date can be parsed.

    The "May Closing (at meeting)" baseline (col G) is the closing balance as of
    END-OF-DAY this snapshot date, so cases-sold must be counted from the DAY AFTER
    it (the filter uses ``d > snapshot``). Anchoring to the meeting date instead
    left the sales between the snapshot and the meeting uncounted -- the May-15/16
    blind spot Abhay flagged on 27 May 2026.
    """
    if not subtitle:
        return None
    seg = None
    for part in str(subtitle).split("·"):
        if "as of" in part.lower():
            seg = part
            break
    if seg is None:
        return None
    seg = seg.lower().split("as of", 1)[1]
    last = None
    for m in _ASOF_DATE_RE.finditer(seg):
        mon = _MONTH_LOOKUP.get(m.group(1).upper())
        if mon is None:
            continue
        try:
            last = date(year, mon, int(m.group(2)))
        except ValueError:
            continue
    return last


def load_latest_closing(ksbc_path: Path):
    """Pull the latest closing balance per (shop, brand, pack) from the
    COMBINED sheet of the current KSBC analysis workbook. Returns:
        closing[(shop_code, brand, pack)] = closing_cases (float)
        combined_label: the COMBINED sheet name (e.g. "MAY 1-14 COMBINED")
    """
    wb = openpyxl.load_workbook(ksbc_path, data_only=True, read_only=True)
    combined_label = next((s for s in wb.sheetnames if s.endswith("COMBINED")), None)
    if not combined_label:
        return {}, None

    ws = wb[combined_label]
    closing = {}
    # Schema:
    #  A Warehouse  B Shop Code  C Shop Name  D Product Code
    #  E Brand Name  F Packing  G Bottle Per Case
    #  H Opening   I Receipts   J Sales   K Closing
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 11:
            continue
        shop_code, brand, pack, closing_cs = row[1], row[4], row[5], row[10]
        if shop_code is None or brand is None or pack is None:
            continue
        try:
            cs = float(closing_cs) if closing_cs is not None else 0.0
        except (TypeError, ValueError):
            continue
        key = (str(shop_code).strip(), str(brand).strip(), str(pack).strip())
        closing[key] = cs

    return closing, combined_label


def load_daily_sales(workbooks, year: int = None):
    """Walk every daily sheet across one or more KSBC workbooks and build a lookup.

    ``workbooks`` may be a single Path (back-compat) or a list of Paths. Daily
    sheets carry their own month name (e.g. "MAY 14", "JUNE 1") so merging across
    a May + June workbook never collides. The year for each sheet is inferred from
    the month relative to today so a window spanning a year boundary stays correct.

    Returns:
        sales[(shop_code, brand, pack)][date_iso]  =  cases_sold (float)
        available_dates: sorted list of dates that have data
    """
    if isinstance(workbooks, (str, Path)):
        workbooks = [workbooks]
    sales = defaultdict(lambda: defaultdict(float))
    available_dates = set()
    today = date.today()
    for ksbc_path in workbooks:
        _accumulate_daily_sales(ksbc_path, sales, available_dates, today)
    return dict(sales), sorted(available_dates)


def _accumulate_daily_sales(ksbc_path, sales, available_dates, today):
    wb = openpyxl.load_workbook(ksbc_path, data_only=True, read_only=True)

    for sheet_name in wb.sheetnames:
        m = DAILY_SHEET_RE.match(sheet_name)
        if not m:
            continue
        month_name, day_str = m.group(1), m.group(2)
        month_num = MONTHS.get(month_name)
        if month_num is None:
            continue
        # Infer year: a month later than today's month belongs to the prior year
        # (handles a Dec->Jan window). Same/earlier month = current year.
        yr = today.year - 1 if month_num > today.month else today.year
        try:
            d = date(yr, month_num, int(day_str))
        except ValueError:
            continue
        ws = wb[sheet_name]
        # Daily-sheet schema: r1 = header, data from r2
        # cols: A=Warehouse B=Shop Code C=Shop Name D=Product Code
        #       E=Brand Name F=Packing G=Bottle Per Case
        #       H=Opening cs I=Opening btl J=In cs K=In btl L=Out cs M=Out btl
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) < 13:
                continue
            shop_code, _shop_name = row[1], row[2]
            brand, pack = row[4], row[5]
            bpc = row[6]
            out_cs = row[11]
            out_btl = row[12]
            if shop_code is None or brand is None or pack is None:
                continue
            try:
                out_cs = float(out_cs) if out_cs is not None else 0.0
            except (TypeError, ValueError):
                out_cs = 0.0
            try:
                out_btl = float(out_btl) if out_btl is not None else 0.0
            except (TypeError, ValueError):
                out_btl = 0.0
            try:
                bpc = float(bpc) if bpc is not None else 0.0
            except (TypeError, ValueError):
                bpc = 0.0
            # Convert loose bottles to case-decimal; full sale = cases + bottles/bpc.
            # Without this, partial cases (loose bottles < a full case) are silently dropped.
            sold_cases = out_cs + (out_btl / bpc if bpc else 0.0)
            if sold_cases == 0:
                continue
            key = (str(shop_code).strip(), str(brand).strip(), str(pack).strip())
            sales[key][d.isoformat()] += sold_cases
        available_dates.add(d)


# ----------------------------------------------------------------------
# Workbook refresh
# ----------------------------------------------------------------------

def classify(target, sold, days_remaining):
    if target is None or not isinstance(target, (int, float)) or target == 0:
        return STATUS_NO_COMMIT
    if sold is None or not isinstance(sold, (int, float)):
        return STATUS_AWAITING
    if sold >= target:
        return STATUS_ACHIEVED
    pct = sold / target if target > 0 else 0
    if pct >= 0.75:
        return STATUS_ON_TRACK
    if pct >= 0.40:
        return STATUS_AT_RISK
    return STATUS_BEHIND


def refresh_bond_sheet(ws, sales_by_key, latest_closing_by_key, available_dates, today):
    """Refresh one bond tab's Sold + Days Remaining columns.
    Returns:
        bond_summary dict with totals and per-row records for the artifact.
    """
    # Parse a date from a workbook cell. Accepts datetime, date, or ISO string.
    def _parse_date_cell(cell):
        v = cell.value
        if isinstance(v, datetime):
            return v.date()
        if isinstance(v, date):
            return v
        if isinstance(v, str) and v.strip():
            for fmt in ("%Y-%m-%d", "%d-%b-%Y", "%d/%m/%Y", "%d-%m-%Y"):
                try:
                    return datetime.strptime(v.strip(), fmt).date()
                except ValueError:
                    continue
        return None

    # Meeting date  → cell C3
    # Next meeting  → cell G3 (optional; fallback to meeting + REVIEW_CYCLE_DAYS)
    meeting_date = _parse_date_cell(ws["C3"])
    next_meeting = _parse_date_cell(ws["G3"])
    if meeting_date and not next_meeting:
        next_meeting = meeting_date + timedelta(days=REVIEW_CYCLE_DAYS)

    # ---- Sold-cases window anchor = the aging-snapshot ("as of") date ----
    # The "May Closing (at meeting)" baseline (col G) is the aging-snapshot close,
    # so we count sold cases from the DAY AFTER that snapshot -- not the meeting
    # date. Anchoring to the meeting left sales between the snapshot and the meeting
    # uncounted (the May-15/16 blind spot, fixed 27 May 2026). Falls back to the
    # meeting date if the snapshot date can't be parsed from the row-2 subtitle.
    subtitle = ws.cell(row=2, column=1).value or ""
    anchor_year = meeting_date.year if meeting_date else today.year
    sold_anchor = parse_aging_snapshot_date(subtitle, anchor_year) or meeting_date

    # Keep the visible labels honest about the anchor date.
    if sold_anchor is not None:
        ws.cell(row=6, column=7,
                value=f"Snapshot Closing\n(as of {sold_anchor.day} {sold_anchor:%b})")
        ws.cell(row=6, column=12,
                value=f"Cases Sold\nsince {sold_anchor.day} {sold_anchor:%b}")
        ws.cell(row=3, column=9,
                value=(f"←  Sold cases counted from the aging-snapshot close "
                       f"({sold_anchor.day} {sold_anchor:%b}) onwards; "
                       f"Days remaining counted to Next Meeting."))

    field_staff = ""
    sub = ws.cell(row=2, column=1).value or ""
    for part in str(sub).split("·"):
        part = part.strip()
        if part.lower().startswith("field staff:"):
            field_staff = part.split(":", 1)[1].strip()
            break

    rows = []
    total_target = 0.0
    total_sold = 0.0

    # Data starts at row 7 (header row 6, banner intermixed).
    # New column layout:
    #   B shop_code · D severity · E brand · F pack
    #   G May Closing · H Latest Closing (write) · I MoC · J Zero Months
    #   K Target Cases · L Cases Sold (write) · M % achieved · N Days Remaining (write) · O Status
    total_latest_closing = 0.0
    for r in range(7, ws.max_row + 1):
        shop_code = ws.cell(row=r, column=2).value
        severity  = ws.cell(row=r, column=4).value
        if shop_code is None or severity is None:
            continue
        brand = ws.cell(row=r, column=5).value
        pack  = ws.cell(row=r, column=6).value
        target = ws.cell(row=r, column=11).value   # col K
        if brand is None or pack is None:
            continue

        key = (str(shop_code).strip(), str(brand).strip(), str(pack).strip())

        # ---- Latest Closing (col H = 8) ----
        latest_closing = latest_closing_by_key.get(key)
        if latest_closing is not None:
            ws.cell(row=r, column=8, value=round(latest_closing, 4))
            total_latest_closing += latest_closing

        # ---- Cases sold since the aging-snapshot date (col L = 12) + Days remaining (col N = 14) ----
        sold = None
        if meeting_date is not None:
            day_dict = sales_by_key.get(key, {})
            s = 0.0
            for iso_d, cs in day_dict.items():
                d = date.fromisoformat(iso_d)
                if d > sold_anchor and d <= today:
                    s += cs
            sold = s
            ws.cell(row=r, column=12, value=round(sold, 4))
            days_remaining = (next_meeting - today).days
            ws.cell(row=r, column=14, value=days_remaining)

        try:
            tgt_val = float(target) if target is not None else None
        except (TypeError, ValueError):
            tgt_val = None

        rows.append({
            "shop_code": str(shop_code),
            "shop_name": ws.cell(row=r, column=3).value or "",
            "severity":  ws.cell(row=r, column=4).value or "",
            "brand": str(brand),
            "pack":  str(pack),
            "may_closing":    float(ws.cell(row=r, column=7).value or 0) or None,
            "latest_closing": round(latest_closing, 2) if latest_closing is not None else None,
            "target": tgt_val,
            "sold":   round(sold, 2) if sold is not None else None,
            "status": classify(tgt_val, sold,
                               (next_meeting - today).days if next_meeting else None),
        })
        if tgt_val is not None:
            total_target += tgt_val
        if sold is not None:
            total_sold += sold

    days_remaining = (next_meeting - today).days if next_meeting else None
    # May-Closing total across all data rows (positive only)
    may_total = sum((row["may_closing"] or 0) for row in rows)
    return {
        "bond": ws.title,
        "field_staff": field_staff,
        "meeting_date": meeting_date.isoformat() if meeting_date else None,
        "next_meeting": next_meeting.isoformat() if next_meeting else None,
        "days_remaining": days_remaining,
        "target_total": round(total_target, 2),
        "sold_total":   round(total_sold, 2),
        "may_closing_total":    round(may_total, 2),
        "latest_closing_total": round(total_latest_closing, 2),
        "rows": rows,
    }


# ----------------------------------------------------------------------
# Artifact payload + HTML
# ----------------------------------------------------------------------

LIVE_HTML = """<!doctype html>
<html lang="en"><head>
<meta charset="utf-8">
<title>K.S. Distillery — Commitment Tracker</title>
<style>
  :root {
    /* dark palette */
    --bg:#0A0E1A;             /* page background */
    --bg-2:#0F1424;           /* alt panels */
    --card:#141A2E;           /* card background */
    --card-hi:#1A2138;        /* hover */
    --line:#222A44;           /* borders */
    --line-hi:#2D375A;        /* heavier divider */
    --ink:#E8ECF5;            /* primary text */
    --soft:#8B97B5;           /* muted text */
    --grey:#6B7591;           /* fainter */
    --gold:#FFB300;
    --gold-dim:#C99100;

    /* status palettes (bright on dark) */
    --done-bg:#163E1F;  --done-fg:#9CCC65;  --done-bar:#66BB6A;
    --on-bg:#0E3A4A;    --on-fg:#4DD0E1;    --on-bar:#26C6DA;
    --risk-bg:#4A2E0E;  --risk-fg:#FFB74D;  --risk-bar:#FFA726;
    --behind-bg:#4A1818;--behind-fg:#EF9A9A;--behind-bar:#EF5350;
    --none-bg:#252B3D;  --none-fg:#90A4BE;  --none-bar:#5C6884;
  }
  * { box-sizing: border-box; }
  body { font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Inter,sans-serif;
         margin:0; padding:24px 28px 40px;
         background:
           radial-gradient(900px 600px at 88% -8%, rgba(255,179,0,0.05), transparent 60%),
           radial-gradient(700px 500px at -10% 110%, rgba(38,198,218,0.06), transparent 60%),
           var(--bg);
         color:var(--ink); -webkit-font-smoothing:antialiased; }
  header { display:flex; align-items:baseline; gap:14px; padding-bottom:18px;
           border-bottom:1px solid var(--line); }
  header .badge { font-size:10px; font-weight:700; letter-spacing:1.5px; color:var(--gold);
                  padding:3px 8px; border:1px solid var(--gold-dim); border-radius:4px;
                  text-transform:uppercase; }
  h1 { margin:0; font-size:22px; font-weight:600; letter-spacing:-0.2px; color:var(--ink); }
  h1 .accent { color:var(--gold); }
  .sub { color:var(--soft); font-size:12px; margin-left:auto; font-variant-numeric:tabular-nums; }

  .kpis { display:grid; grid-template-columns:repeat(4,1fr); gap:14px; margin:22px 0 28px; }
  .kpi { background:linear-gradient(180deg, var(--card) 0%, var(--bg-2) 100%);
         border:1px solid var(--line); border-radius:12px; padding:16px 18px;
         box-shadow: inset 0 1px 0 rgba(255,255,255,0.03), 0 1px 0 rgba(0,0,0,0.4); }
  .kpi .label { font-size:11px; color:var(--soft); text-transform:uppercase; letter-spacing:1px; font-weight:600;}
  .kpi .value { font-size:28px; font-weight:700; margin-top:8px; color:var(--ink);
                font-variant-numeric:tabular-nums; letter-spacing:-0.5px; }
  .kpi .hint  { font-size:11px; color:var(--grey); margin-top:3px; }
  .kpi.k-target .value { color:var(--gold); }
  .kpi.k-pct .value { color:var(--on-fg); }
  .kpi.k-days .value { color:var(--ink); }

  .toolbar { display:flex; gap:10px; align-items:center; font-size:12px;
             color:var(--soft); margin-bottom:10px; }
  .toolbar .label { font-weight:700; color:var(--ink); font-size:12px; text-transform:uppercase;
                    letter-spacing:1px; }
  .toolbar .hint { margin-left:auto; font-style:italic; color:var(--grey); }

  .bond { display:grid; grid-template-columns: 170px 90px 1fr 140px 110px;
          align-items:center; gap:14px; padding:13px 16px;
          background:var(--card); border:1px solid var(--line);
          border-radius:10px; margin-bottom:7px; font-size:13px;
          cursor:pointer; transition: background .12s, border-color .12s, transform .12s; }
  .bond:hover { background:var(--card-hi); border-color:var(--line-hi); }
  .bond.active { background:var(--card-hi); border-color:var(--gold-dim);
                 box-shadow: 0 0 0 1px rgba(255,179,0,0.25), 0 8px 24px rgba(0,0,0,0.3); }
  .name { font-weight:700; color:var(--ink); letter-spacing:0.2px; text-align:left; }
  .name .staff { font-size:10.5px; font-weight:400; color:var(--soft);
                 margin-top:3px; text-transform:uppercase; letter-spacing:0.8px; }
  .bar-wrap { position:relative; height:8px; background:var(--line);
              border-radius:4px; overflow:hidden; }
  .bar { position:absolute; top:0; left:0; bottom:0; border-radius:4px;
         box-shadow: 0 0 8px currentColor; }
  .pill { font-size:10.5px; font-weight:700; padding:4px 11px; border-radius:11px;
          text-align:center; letter-spacing:0.3px; white-space:nowrap;
          border:1px solid transparent; }
  .pill-on    { background:var(--on-bg);     color:var(--on-fg);     border-color:rgba(38,198,218,0.3); }
  .pill-risk  { background:var(--risk-bg);   color:var(--risk-fg);   border-color:rgba(255,167,38,0.3); }
  .pill-behind{ background:var(--behind-bg); color:var(--behind-fg); border-color:rgba(239,83,80,0.35); }
  .pill-done  { background:var(--done-bg);   color:var(--done-fg);   border-color:rgba(102,187,106,0.3); }
  .pill-none  { background:var(--none-bg);   color:var(--none-fg);   border-color:rgba(140,150,180,0.18); }
  .right { text-align:center; font-variant-numeric:tabular-nums; }
  .ratio { color:var(--soft); margin-right:14px; }
  .dot   { color:var(--grey); margin-right:14px; }
  .pctval { color:var(--ink); }


  .drill { background:var(--card); border:1px solid var(--line); border-radius:12px;
           padding:20px 22px; margin-top:22px;
           box-shadow: 0 8px 32px rgba(0,0,0,0.4); }
  .drill h3 { margin:0 0 4px; font-size:16px; font-weight:700; color:var(--ink); letter-spacing:-0.2px; }
  .drill .sub { font-size:11px; color:var(--soft); margin:0 0 14px 0; text-transform:uppercase;
                letter-spacing:1px; }
  table { width:100%; border-collapse:collapse; font-size:12.5px; }
  thead th { text-align:center; padding:8px 10px; font-weight:600; color:var(--soft);
             border-bottom:1px solid var(--line-hi); text-transform:uppercase;
             font-size:10.5px; letter-spacing:0.8px; }
  tbody tr.shop-row td { padding:10px 10px 6px 10px;
                         background:linear-gradient(180deg, rgba(255,179,0,0.04), transparent);
                         border-top:1px solid var(--line-hi);
                         border-bottom:1px solid var(--line); }
  tbody tr.shop-row .shop-name { font-size:12.5px; font-weight:700; color:var(--gold);
                                  text-transform:uppercase; letter-spacing:0.6px; }
  tbody tr.shop-row .shop-meta { font-size:10px; font-weight:500; color:var(--soft);
                                  text-transform:uppercase; letter-spacing:0.8px; margin-left:8px; }
  tbody tr.line-row td { padding:7px 10px; border-bottom:1px solid var(--line); color:var(--ink);
                          font-variant-numeric:tabular-nums; }
  tbody tr.line-row:hover { background:rgba(255,255,255,0.02); }
  td.r { text-align:center; }
  /* default centre every tbody cell; left-align only the explicit "left" cells */
  tbody td { text-align:center; }
  tbody td.left { text-align:left; }
  thead th:nth-child(2) { text-align:left; }
  td.indent { padding-left:24px !important; color:var(--grey); }
  .muted { color:var(--soft); }
  .arrow { color:var(--grey); font-size:11px; }
</style>
</head>
<body>
<header>
  <span class="badge">K.S. Distillery</span>
  <h1>Commitment Tracker <span class="accent">·</span> Live</h1>
  <div class="sub" id="meta">Refresh ${REFRESH_TS}</div>
</header>

<div class="kpis">
  <div class="kpi k-target"><div class="label">Total committed</div><div class="value" id="kpi-target">—</div><div class="hint" id="kpi-target-sub">—</div></div>
  <div class="kpi k-sold"><div class="label">Sold so far</div>   <div class="value" id="kpi-sold">—</div>   <div class="hint" id="kpi-sold-sub">committed lines · since snapshot</div></div>
  <div class="kpi k-pct"><div class="label">% achieved</div>    <div class="value" id="kpi-pct">—</div>    <div class="hint" id="kpi-pct-sub">vs total target</div></div>
  <div class="kpi k-days"><div class="label">Days remaining</div><div class="value" id="kpi-days">—</div>   <div class="hint" id="kpi-days-sub">in the review cycle</div></div>
</div>

<div class="toolbar">
  <span class="label">Bonds by commitment progress</span>
  <span class="hint">click a bond to drill down</span>
</div>

<div id="bonds"></div>
<div id="drill" class="drill" style="display:none;"></div>

<script>
const PAYLOAD = __PAYLOAD__;

const TIER = {
  "${STATUS_ACHIEVED}": {cls:"pill-done",   bar:"var(--done-bar)"},
  "${STATUS_ON_TRACK}": {cls:"pill-on",     bar:"var(--on-bar)" },
  "${STATUS_AT_RISK}":  {cls:"pill-risk",   bar:"var(--risk-bar)"},
  "${STATUS_BEHIND}":   {cls:"pill-behind", bar:"var(--behind-bar)"  },
  "${STATUS_NO_COMMIT}":{cls:"pill-none",   bar:"var(--none-bar)"   },
  "${STATUS_AWAITING}": {cls:"pill-none",   bar:"var(--none-bar)"   },
};

const fmt = (n) => n == null ? "—" : Math.round(n * 100) / 100;
const pct = (s, t) => (t > 0 ? Math.round(s / t * 100) : 0);

// Committed-line basis: sum sold ONLY on lines the ASM put a target on
// (target > 0). Movement on aging lines with no commitment is excluded, so
// every dashboard figure matches the bond/cluster commitment PDFs and answers
// "of what was committed, how much has moved?".
const committedSold = (b) => (b.rows || []).reduce((a, r) =>
    a + ((r.target != null && r.target > 0) ? (r.sold || 0) : 0), 0);

function bondTotalStatus(b) {
  if (!b.target_total) return "${STATUS_NO_COMMIT}";
  const s = committedSold(b);
  if (s >= b.target_total) return "${STATUS_ACHIEVED}";
  const p = b.target_total > 0 ? s / b.target_total : 0;
  if (p >= 0.75) return "${STATUS_ON_TRACK}";
  if (p >= 0.40) return "${STATUS_AT_RISK}";
  return "${STATUS_BEHIND}";
}

// KPI tiles
const totT = PAYLOAD.bonds.reduce((a,b)=>a+(b.target_total||0),0);
const totS = PAYLOAD.bonds.reduce((a,b)=>a+committedSold(b),0);
const daysR = PAYLOAD.bonds.find(b => b.days_remaining != null)?.days_remaining;
document.getElementById("kpi-target").textContent = fmt(totT) + " cs";
document.getElementById("kpi-target-sub").textContent = PAYLOAD.commit_lines + " commit lines · " + PAYLOAD.bonds.length + " bonds";
document.getElementById("kpi-sold").textContent   = fmt(totS) + " cs";
document.getElementById("kpi-pct").textContent    = pct(totS, totT) + "%";
document.getElementById("kpi-days").textContent   = daysR != null ? daysR : "—";

// Bond rows (sorted by descending target)
const bondsBox = document.getElementById("bonds");
const sortedBonds = [...PAYLOAD.bonds].sort((a,b) => (b.target_total||0) - (a.target_total||0));
sortedBonds.forEach(b => {
  const bs = committedSold(b);
  const p = pct(bs, b.target_total || 0);
  const st = TIER[bondTotalStatus(b)];
  const row = document.createElement("div");
  row.className = "bond";
  row.dataset.bond = b.bond;
  row.innerHTML = `
    <div class="name">${b.bond}<div class="staff">${b.field_staff||"—"}</div></div>
    <div class="right muted">${fmt(b.target_total)} cs</div>
    <div class="bar-wrap"><div class="bar" style="width:${Math.min(p,100)}%; background:${st.bar}; color:${st.bar};"></div></div>
    <div class="right"><span class="ratio">${fmt(bs)} / ${fmt(b.target_total)}</span><span class="dot">·</span><b class="pctval">${p}%</b></div>
    <div class="pill ${st.cls}">${bondTotalStatus(b)}</div>
  `;
  row.addEventListener("click", () => showDrill(b.bond));
  bondsBox.appendChild(row);
});

function showDrill(bond) {
  document.querySelectorAll(".bond").forEach(r => r.classList.toggle("active", r.dataset.bond === bond));
  const b = PAYLOAD.bonds.find(x => x.bond === bond);
  const drill = document.getElementById("drill");
  const linesWithTarget = b.rows.filter(r => r.target != null && r.target > 0);
  if (linesWithTarget.length === 0) {
    drill.style.display = "block";
    drill.innerHTML = `<h3>${bond}</h3><div class="sub">${b.field_staff || "—"}</div><p class="muted">No commitments captured yet for this bond.</p>`;
    return;
  }
  // sort by shop name, then within shop by target desc
  linesWithTarget.sort((x, y) => {
    if (x.shop_name === y.shop_name) return (y.target || 0) - (x.target || 0);
    return (x.shop_name || "").localeCompare(y.shop_name || "");
  });
  const shopTotals = {};
  for (const r of linesWithTarget) {
    if (!shopTotals[r.shop_name]) shopTotals[r.shop_name] = { target: 0, sold: 0, count: 0 };
    shopTotals[r.shop_name].target += (r.target || 0);
    shopTotals[r.shop_name].sold += (r.sold || 0);
    shopTotals[r.shop_name].count += 1;
  }
  const bondPct = pct(committedSold(b), b.target_total || 0);
  let html = `<h3>${bond} <span class="accent" style="color:var(--gold);">·</span> <span style="color:var(--soft); font-weight:500;">${b.field_staff || "—"}</span></h3>
    <div class="sub">${linesWithTarget.length} commit lines · ${Object.keys(shopTotals).length} shops · ${fmt(b.target_total)} cs target · ${bondPct}% achieved</div>
    <table>
      <thead><tr><th style="width:36px;"></th><th>Item</th><th class="r">Commit</th><th class="r">Sold</th><th class="r">%</th><th>Status</th></tr></thead><tbody>`;
  let currentShop = null;
  for (const r of linesWithTarget) {
    if (r.shop_name !== currentShop) {
      currentShop = r.shop_name;
      const st = shopTotals[currentShop];
      const sp = pct(st.sold, st.target);
      const spColor = sp >= 100 ? "var(--done-fg)" : sp >= 75 ? "var(--on-fg)" : sp >= 40 ? "var(--risk-fg)" : sp > 0 ? "var(--behind-fg)" : "var(--soft)";
      html += `<tr class="shop-row"><td colspan="2" class="left"><span class="shop-name">${shortName(currentShop)}</span><span class="shop-meta">${st.count} line${st.count>1?"s":""}</span></td><td class="r" style="font-weight:700; color:var(--ink);">${fmt(st.target)}</td><td class="r" style="font-weight:700; color:var(--ink);">${fmt(st.sold)}</td><td class="r" style="font-weight:700; color:${spColor};">${sp}%</td><td></td></tr>`;
    }
    const p = pct(r.sold || 0, r.target);
    const st = TIER[r.status] || TIER["${STATUS_AWAITING}"];
    html += `<tr class="line-row"><td class="indent"><span class="arrow">↳</span></td>
              <td class="muted left">${shortBrand(r.brand)} <span style="color:var(--grey);">·</span> ${shortPack(r.pack)}</td>
              <td class="r">${fmt(r.target)}</td>
              <td class="r"><b>${fmt(r.sold)}</b></td>
              <td class="r" style="color:${st.bar}; font-weight:700;">${p}%</td>
              <td><span class="pill ${st.cls}">${r.status}</span></td></tr>`;
  }
  html += `</tbody></table>`;
  drill.innerHTML = html;
  drill.style.display = "block";
}

function shortName(n) { if (!n) return ""; const p = n.split("-", 2); return p.length === 2 ? p[1] : n; }
function shortBrand(b) { return (b || "").replace("NO.1", "No.1").replace("XXX RUM", "XXX").substring(0, 28); }
function shortPack(p) { return (p || "").replace(" ML", "ml"); }

// Auto-open first bond with commitments captured
const firstWithCommit = PAYLOAD.bonds.find(b => (b.target_total || 0) > 0);
if (firstWithCommit) showDrill(firstWithCommit.bond);
</script>
</body></html>
"""


def build_payload(bond_summaries):
    return {
        "refresh_ts": datetime.now().isoformat(timespec="minutes"),
        "bonds": bond_summaries,
        "commit_lines": sum(len([r for r in b["rows"] if r["target"] is not None and r["target"] > 0])
                            for b in bond_summaries),
    }


def write_artifact(payload, out_html):
    html = (LIVE_HTML
            .replace("${REFRESH_TS}", payload["refresh_ts"])
            .replace("${STATUS_ACHIEVED}", STATUS_ACHIEVED)
            .replace("${STATUS_ON_TRACK}", STATUS_ON_TRACK)
            .replace("${STATUS_AT_RISK}",  STATUS_AT_RISK)
            .replace("${STATUS_BEHIND}",   STATUS_BEHIND)
            .replace("${STATUS_NO_COMMIT}", STATUS_NO_COMMIT)
            .replace("${STATUS_AWAITING}",  STATUS_AWAITING)
            .replace("__PAYLOAD__", json.dumps(payload, ensure_ascii=False)))
    Path(out_html).parent.mkdir(parents=True, exist_ok=True)
    Path(out_html).write_text(html, encoding="utf-8")


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--as-of", help="Override today's date for testing (YYYY-MM-DD)")
    args = p.parse_args()

    today = (datetime.strptime(args.as_of, "%Y-%m-%d").date()
             if args.as_of else date.today())

    track_path = find_commitment_workbook()
    ksbc_path  = find_ksbc_workbook()      # current-month workbook: latest closing source

    # ---- Determine the earliest commitment-window anchor across all bond tabs ----
    # The sold window starts at each bond's aging-snapshot date (col-2 subtitle)
    # or, failing that, its meeting date (C3). We must read every KSBC daily sheet
    # from the EARLIEST such anchor through today -- which can span >1 month, hence
    # >1 workbook. (3 Jun 2026 fix: a May-16 meeting refreshed in June was only
    # reading the June workbook, dropping all of May 15-31.)
    _peek = openpyxl.load_workbook(track_path, data_only=True, read_only=True)
    earliest_anchor = None
    for bond in BONDS:
        if bond not in _peek.sheetnames:
            continue
        bws = _peek[bond]
        c3 = bws["C3"].value
        mdate = c3.date() if isinstance(c3, datetime) else (c3 if isinstance(c3, date) else None)
        ay = mdate.year if mdate else today.year
        anc = parse_aging_snapshot_date(bws.cell(row=2, column=1).value or "", ay) or mdate
        if anc and (earliest_anchor is None or anc < earliest_anchor):
            earliest_anchor = anc
    _peek.close()
    if earliest_anchor is None:
        earliest_anchor = date(today.year, today.month, 1)

    # Collect one workbook per month in the window (current-month preferred form each)
    ksbc_paths, missing_months = [], []
    seen = set()
    for mname, _yr in months_in_window(earliest_anchor, today):
        wbp = find_ksbc_workbook_for_month(mname)
        if wbp and wbp not in seen:
            ksbc_paths.append(wbp); seen.add(wbp)
        elif wbp is None:
            missing_months.append(mname.title())
    if not ksbc_paths:
        ksbc_paths = [ksbc_path]

    print(f"  Tracker workbook : {track_path.name}")
    print(f"  Window anchor    : {earliest_anchor.isoformat()} → {today.isoformat()}")
    print(f"  KSBC workbooks   : {', '.join(w.name for w in ksbc_paths)}")
    if missing_months:
        print(f"  WARN: no KSBC workbook for: {', '.join(missing_months)} (those days excluded)")
    print(f"  Refreshing as of : {today.isoformat()}")
    print()

    print("Loading KSBC daily sales...")
    sales_by_key, available_dates = load_daily_sales(ksbc_paths, today.year)
    print(f"  {len(sales_by_key):,} (shop, brand, pack) keys with non-zero sales")
    print(f"  date range       : {available_dates[0] if available_dates else '—'} → "
          f"{available_dates[-1] if available_dates else '—'}")

    print("Loading latest closing balances from COMBINED sheet...")
    latest_closing_by_key, combined_label = load_latest_closing(ksbc_path)
    if combined_label:
        print(f"  COMBINED sheet   : {combined_label}")
        print(f"  {len(latest_closing_by_key):,} (shop, brand, pack) closing positions")
    else:
        print(f"  WARN: no COMBINED sheet found in {ksbc_path.name} — Latest Closing column will be empty")
    print()

    wb = openpyxl.load_workbook(track_path)
    bond_summaries = []
    for bond in BONDS:
        if bond not in wb.sheetnames:
            print(f"  WARN: bond '{bond}' missing from tracker workbook")
            continue
        ws = wb[bond]
        summary = refresh_bond_sheet(ws, sales_by_key, latest_closing_by_key,
                                     available_dates, today)
        bond_summaries.append(summary)

    wb.save(track_path)
    print(f"Wrote refreshed Sold + Days columns to {track_path.name}")
    print()

    payload = build_payload(bond_summaries)
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    payload_path = OUTPUTS_DIR / "commitment_tracker_payload.json"
    payload_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    html_path = OUTPUTS_DIR / "commitment_tracker_live.html"
    write_artifact(payload, html_path)
    print(f"Live artifact HTML : {html_path}")
    print(f"Payload JSON       : {payload_path}")
    print()

    # ---- Console summary ----
    print("Status by bond:")
    print("-" * 70)
    print(f"  {'Bond':<16} {'Target':>9} {'Sold':>9} {'%':>5}  Status")
    print("-" * 70)
    for b in sorted(bond_summaries, key=lambda x: -(x["target_total"] or 0)):
        if not b["target_total"]:
            print(f"  {b['bond']:<16} {'—':>9} {'—':>9} {'—':>5}  {STATUS_NO_COMMIT}")
            continue
        pct_v = round(b["sold_total"] / b["target_total"] * 100) if b["target_total"] else 0
        if b["sold_total"] >= b["target_total"]:
            status = STATUS_ACHIEVED
        elif pct_v >= 75:
            status = STATUS_ON_TRACK
        elif pct_v >= 40:
            status = STATUS_AT_RISK
        else:
            status = STATUS_BEHIND
        print(f"  {b['bond']:<16} {b['target_total']:>9.1f} {b['sold_total']:>9.1f} {pct_v:>4}%  {status}")
    print()
    total_t = sum(b['target_total'] or 0 for b in bond_summaries)
    total_s = sum(b['sold_total']   or 0 for b in bond_summaries)
    if total_t:
        print(f"  GRAND TOTAL      {total_t:>9.1f} {total_s:>9.1f} {round(total_s/total_t*100):>4}%")
    else:
        print("  GRAND TOTAL      no commitments captured yet")


if __name__ == "__main__":
    main()
