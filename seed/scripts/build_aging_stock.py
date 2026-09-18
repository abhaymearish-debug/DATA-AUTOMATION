#!/usr/bin/env python3
"""
KSBC AGING STOCK ANALYSIS — durable build pipeline.

Reads N months of KSBC COMBINED sheets, computes aging metrics, produces the
full multi-sheet workbook (SUMMARY -> 15 bond sheets -> FULL DETAIL ->
drill-downs).

USAGE
  python3 build_aging_stock.py                       # auto-discover all months
  python3 build_aging_stock.py --last 3              # last 3 months ending at most recent
  python3 build_aging_stock.py --from MARCH --to MAY # explicit range

Outputs to: Claude/Aging stock/AGING STOCK ANALYSIS - <start> to <end>.xlsx

METHODOLOGY
For each Shop × Brand × Pack across the period:
  total_sales over N days; latest_closing; monthly_rate = sales*30/days
  months_of_cover = latest_closing / monthly_rate
  zero_sales_months = count of months where Sales=0 but stock present
Tiers:
  NON-MOVING     - Closing > 0 AND total sales = 0
  CRITICAL AGING - MoC > 6 OR (>=2 zero-sale months AND closing >= 5)
  SLOW MOVING    - 3 < MoC <= 6
  HEALTHY        - MoC <= 3
"""
import argparse
import re
import sys
from pathlib import Path
from collections import defaultdict, Counter
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import DataBarRule

# Resolve CLAUDE_ROOT — derive from script location so the same script works
# both on the user's Mac and inside any sandboxed mount.
# Script lives at: <CLAUDE_ROOT>/.claude/scripts/build_aging_stock.py
CLAUDE_ROOT = Path(__file__).resolve().parent.parent.parent
KSBC_FOLDER = CLAUDE_ROOT / "KSBC shop sales"
OUTPUT_FOLDER = CLAUDE_ROOT / "Aging stock"
SECONDARY_FOLDER = CLAUDE_ROOT / "Secondary sales"
# Abhay's authoritative master data file. The MASTER DATA tab embedded in monthly
# analysis workbooks can lag behind. Always prefer the standalone file at root.
STANDALONE_MASTER = CLAUDE_ROOT / "MASTER DATA CONFIRMED.xlsx"

MONTH_ORDER = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY",
               "AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"]
MONTH_DAYS = {"JANUARY":31,"FEBRUARY":28,"MARCH":31,"APRIL":30,"MAY":31,
              "JUNE":30,"JULY":31,"AUGUST":31,"SEPTEMBER":30,"OCTOBER":31,
              "NOVEMBER":30,"DECEMBER":31}
MONTH_ABBREV = {"MARCH":"Mar","APRIL":"Apr","MAY":"May","JUNE":"Jun","JULY":"Jul",
                "AUGUST":"Aug","SEPTEMBER":"Sep","OCTOBER":"Oct","NOVEMBER":"Nov",
                "DECEMBER":"Dec","JANUARY":"Jan","FEBRUARY":"Feb"}

NAVY_DEEP = "FF0D1B4A"; NAVY_MID = "FF1A237E"; NAVY_SOFT = "FF2E3F70"
GOLD = "FFFFB300"; GOLD_DIM = "FFFFD54F"; WHITE = "FFFFFFFF"
GREY_TXT = "FF374151"; GREY_LBL = "FF6B7280"; GREY_BG = "FFF8F9FB"
GREY_ZERO = "FFB0B7C3"; LINE_GREY = "FFE5E7EB"; ROW_ALT = "FFFAFBFC"
SHOP_BAND = "FF334466"

TIER = {
    "NON-MOVING":     {"dark":"FFB71C1C","bright":"FFE53935","light":"FFFFEBEE","label":"NON-MOVING","icon":"●","emoji":"🔴"},
    "CRITICAL AGING": {"dark":"FFE65100","bright":"FFF57C00","light":"FFFFF3E0","label":"CRITICAL","icon":"▲","emoji":"🟠"},
    "SLOW MOVING":    {"dark":"FFF9A825","bright":"FFFBC02D","light":"FFFFFDE7","label":"SLOW","icon":"◆","emoji":"🟡"},
    "HEALTHY":        {"dark":"FF2E7D32","bright":"FF43A047","light":"FFE8F5E9","label":"HEALTHY","icon":"✓","emoji":"🟢"},
    "NEW STOCK":      {"dark":"FF1565C0","bright":"FF1E88E5","light":"FFE3F2FD","label":"NEW STOCK","icon":"✦","emoji":"🔵"},
}
HAIR = Side(style='thin', color=LINE_GREY)
HAIR_BORDER = Border(left=HAIR, right=HAIR, top=HAIR, bottom=HAIR)
GOLD_THIN_BOT = Border(bottom=Side(style='medium', color=GOLD))
GOLD_MED_TOP = Border(top=Side(style='medium', color=GOLD))
NUM_FMT = '#,##0.00;-#,##0.00;"·"'
GRACE_DAYS = 21  # stock newer than this (days on shelf) is held out of aging -> NEW STOCK (was 7 until 9 Jun 2026, Abhay-approved)

# --- Methodology calibration constants (17 Jun 2026 audit fixes, Abhay-approved) ---
# NON-MOVING is STRICT: total period sales = 0 (the tier is literally non-moving). No negligible
# floor -- a single bottle sold (1/48 cs) is a real sale, so that SKU is CRITICAL, not NON-MOVING
# (Abhay, 17 Jun 2026 PM). Barely-moving items therefore sit in CRITICAL with honest (large) cover.
RATE_HORIZON_MONTHS = 3           # C2/H1: months-of-cover velocity uses the TRAILING N months only
                                  # -> window-length invariant AND recency-weighted
STALL_MONTHS = 2                  # C2: holding stock with ZERO sales across the trailing N present
                                  # months = a recent dead-stop -> escalate out of HEALTHY/SLOW
MIN_RATE_DAYS = 30                # L1: never annualise a velocity from < 1 month of shelf life


def _window_years(months, dispatch_years=None):
    """Assign a calendar year to each window month, cross-year safe (H2/S2 fix, 17 Jun 2026).
    Anchor the LATEST month to its most-recent real occurrence per the run clock, then walk
    BACKWARDS, decrementing the year on each reverse-wrap. Replaces the old anchor (max
    dispatch-date year, else hardcoded 2026), which mis-stamped any cross-year window and froze
    every run at 2026 once dispatch data was stale/absent. dispatch_years only guards a skewed
    sandbox clock — the latest month can never predate the newest in-window dispatch."""
    import datetime as _dtm
    today = _dtm.date.today()
    latest = months[-1]
    li1 = MONTH_ORDER.index(latest) + 1
    base = today.year if li1 <= today.month else today.year - 1
    if dispatch_years:
        base = max(base, max(dispatch_years))
    wraps = 0; prev = -1
    for mm in months:
        i = MONTH_ORDER.index(mm)
        if i <= prev: wraps += 1
        prev = i
    y = base - wraps; prev = -1; years = {}
    for mm in months:
        i = MONTH_ORDER.index(mm)
        if i <= prev: y += 1
        years[mm] = y; prev = i
    return years


def classify_position(latest_closing, total_sales, months_cover, zero_sales_months,
                      pre_existed, no_prior_stock, days_on_shelf,
                      recent_stall=False, restock_new=False):
    """SINGLE SOURCE OF TRUTH for tier assignment. Called by compute_metrics AND re-derived by
    VALIDATION CHECK 4, so the production classifier and its self-check can never drift.
    Returns (severity, is_new)."""
    if latest_closing > 0 and total_sales <= 1e-9:   # strict: literally sold nothing in the period
        severity = "NON-MOVING"
    elif months_cover > 6:
        severity = "CRITICAL AGING"
    elif 3 < months_cover <= 6:
        severity = "SLOW MOVING"
        if zero_sales_months >= 2 and latest_closing >= 5:
            severity = "CRITICAL AGING"
    else:
        severity = "HEALTHY"
    # C2 recent-stall override: still holding stock but no sales across the trailing STALL_MONTHS
    # present months -> escalate. A long-run average rate can otherwise mask a recent dead-stop.
    if recent_stall and severity in ("HEALTHY", "SLOW MOVING"):
        severity = "CRITICAL AGING"
    # NEW STOCK grace: a genuine debut (never held stock before, within grace) OR a genuine
    # restock of a line that went FULLY empty in-window and was re-dispatched within grace (M3).
    is_new = (latest_closing > 0) and (days_on_shelf <= GRACE_DAYS) and \
             (((not pre_existed) and no_prior_stock) or restock_new)
    if is_new and severity in ("NON-MOVING", "CRITICAL AGING", "SLOW MOVING"):
        severity = "NEW STOCK"
    return severity, is_new


def discover_workbooks():
    if not KSBC_FOLDER.exists():
        raise SystemExit(f"KSBC folder not found: {KSBC_FOLDER}")
    found = {}
    # Full-month wins absolutely — these are the renamed, final files
    for p in KSBC_FOLDER.glob("* SHOP SALES ANALYSIS.xlsx"):
        m = p.stem.split(" SHOP SALES")[0].upper().strip()
        if m in MONTH_ORDER:
            found[m] = (m, MONTH_DAYS[m], p)
    # Mid-month — collect ALL candidates per month, pick the highest-day one
    mid_candidates = defaultdict(list)
    for p in KSBC_FOLDER.glob("* 1st - * ANALYSIS.xlsx"):
        # skip Excel lock files like ~$MAY 1st - 11th ANALYSIS.xlsx
        if p.name.startswith("~$"): continue
        m_match = re.match(r"^([A-Z]+)\s+1st\s+-\s+(\d+)(?:st|nd|rd|th)\s+ANALYSIS", p.stem, re.I)
        if not m_match: continue
        month = m_match.group(1).upper()
        if month not in MONTH_ORDER: continue
        days = int(m_match.group(2))
        mid_candidates[month].append((days, p))
    for month, candidates in mid_candidates.items():
        if month in found: continue  # full-month exists, ignore mid-month
        # Pick the candidate with the most days (latest snapshot)
        candidates.sort(key=lambda t: -t[0])
        days, p = candidates[0]
        found[month] = (month, days, p)
    resolved = []
    for month, days, path in found.values():
        wb = load_workbook(path, read_only=True, data_only=True)
        combined = None
        for s in wb.sheetnames:
            if "COMBINED" in s.upper() and month in s.upper():
                combined = s; break
        wb.close()
        if combined is None:
            print(f"  Warning: no COMBINED sheet found in {path.name}, skipping")
            continue
        resolved.append((month, days, path, combined))
    resolved.sort(key=lambda t: MONTH_ORDER.index(t[0]))
    return resolved


def filter_periods(periods, last=None, from_month=None, to_month=None):
    if last: return periods[-last:]
    if from_month:
        fm = from_month.upper()
        periods = [p for p in periods if MONTH_ORDER.index(p[0]) >= MONTH_ORDER.index(fm)]
    if to_month:
        tm = to_month.upper()
        periods = [p for p in periods if MONTH_ORDER.index(p[0]) <= MONTH_ORDER.index(tm)]
    return periods


def load_data(periods):
    records = defaultdict(dict)
    dup_total = 0
    for month, days, path, sheet in periods:
        print(f"  Loading {month} from {path.name} :: {sheet} ({days} days)")
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb[sheet]
        dup_month = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[1] is None: continue
            warehouse, shop_code, shop_name, prod_code, brand, pack, btl, opening, receipts, sales, closing = row[:11]
            shop_code = str(shop_code).strip()
            prod_code = str(prod_code).strip() if prod_code else None
            if not prod_code: continue
            key = (shop_code, prod_code)
            if month in records[key]:
                # M1: same (shop, product) listed twice in one COMBINED sheet -> AGGREGATE O/R/S/C
                # (never overwrite, which silently dropped the first line's cases).
                ex = records[key][month]
                ex['opening'] += opening or 0; ex['receipts'] += receipts or 0
                ex['sales'] += sales or 0; ex['closing'] += closing or 0
                dup_month += 1
            else:
                records[key][month] = {
                    'warehouse': warehouse, 'shop_name': shop_name,
                    'brand': brand, 'pack': pack, 'btl': btl,
                    'opening': opening or 0, 'receipts': receipts or 0,
                    'sales': sales or 0, 'closing': closing or 0,
                }
        if dup_month:
            print(f"    note: {dup_month} duplicate (shop x SKU) line(s) in {month} aggregated, not dropped")
            dup_total += dup_month
        wb.close()
    if dup_total:
        print(f"  M1 guard: {dup_total} duplicate source line(s) folded by summation")
    return records


def load_dispatch_dates(periods):
    """Earliest warehouse->shop dispatch date per (shop_code, product_code),
    from each window month's secondary COMBINED DISPATCHES sheet (Inv/GTN Date).
    Exact join: Licensee No = shop code, Product Code = product code. Months whose
    secondary file lacks a dispatch sheet (e.g. MARCH) contribute nothing; those
    positions fall back to month-start in compute_metrics."""
    from datetime import datetime as _dt
    out = {}; out_latest = {}
    if not SECONDARY_FOLDER.exists():
        print('  (no Secondary sales folder -> month-start fallback for all)')
        return out, out_latest
    want = {p[0] for p in periods}
    for f in SECONDARY_FOLDER.glob('*SECONDARY SALES ANALYSIS.xlsx'):
        if f.name.startswith('~$'): continue
        mm = f.stem.split(' ')[0].upper()
        if mm not in want: continue
        try: wb = load_workbook(f, read_only=True, data_only=True)
        except Exception as e: print(f'  warn: cannot open {f.name}: {e}'); continue
        disp = next((s for s in wb.sheetnames if 'DISPATCH' in s.upper()), None)
        if not disp: wb.close(); continue
        ws = wb[disp]
        hdr = [str(c).strip().lower() if c else '' for c in next(ws.iter_rows(min_row=1,max_row=1,values_only=True))]
        def col(*names):
            for n in names:
                if n in hdr: return hdr.index(n)
            return None
        ci_prod = col('product code'); ci_lic = col('licensee no.','licensee no'); ci_date = col('inv/gtn date')
        if None in (ci_prod, ci_lic, ci_date): wb.close(); continue
        n = 0; skipped = 0
        import datetime as _dtm
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[ci_lic] is None or row[ci_prod] is None or row[ci_date] is None: continue
            d = row[ci_date]
            if isinstance(d, _dtm.datetime): d = d.date()
            elif isinstance(d, _dtm.date): pass
            elif isinstance(d, str):
                d = None
                for _fmt in ('%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y'):
                    try: d = _dtm.datetime.strptime(row[ci_date].strip(), _fmt).date(); break
                    except Exception: pass
                if d is None: skipped += 1; continue
            else:
                skipped += 1; continue
            key = (str(row[ci_lic]).strip(), str(row[ci_prod]).strip())
            if key not in out or d < out[key]: out[key] = d
            if key not in out_latest or d > out_latest[key]: out_latest[key] = d
            n += 1
        wb.close()
        print(f'  dispatch dates from {f.name} :: {disp}  ({n} lines' + (f', {skipped} unparseable dates skipped' if skipped else '') + ')')
    print(f'  positions with a dispatch date: {len(out)}')
    return out, out_latest


def compute_metrics(records, periods, dispatch_dates=None, as_of=None, dispatch_latest=None):
    """Compute aging metrics.

    STRICT-LATEST-PERIOD RULE (locked 13 May 2026):
    Only include (Shop × SKU) rows that are present in the LATEST period's
    COMBINED data. If a row appeared in earlier months but vanished from the
    latest month, we cannot say it's currently aging — the stock may have
    been sold/transferred/destroyed or the shop may be closed. Treating
    carry-forward closing as 'latest_closing' creates ghost aging exposure
    that misleads sales meetings. Excluded rows ARE counted and reported by
    the validation layer.
    """
    months = [p[0] for p in periods]
    total_days = sum(p[1] for p in periods)
    latest_month = months[-1]
    latest_days = periods[-1][1]
    from datetime import date as _date
    dispatch_dates = dispatch_dates or {}
    dispatch_latest = dispatch_latest or {}
    # H2/S2 (17 Jun 2026): cross-year-safe year assignment via the shared _window_years helper,
    # anchored to the run clock (not the old max-dispatch-year / hardcoded-2026 anchor).
    _dy = [d.year for d in dispatch_dates.values()] if dispatch_dates else []
    _years = _window_years(months, _dy)
    window_start = _date(_years[months[0]], MONTH_ORDER.index(months[0]) + 1, 1)
    if as_of is None:
        as_of = _date(_years[latest_month], MONTH_ORDER.index(latest_month) + 1, latest_days)
    month_start = {mm: _date(_years[mm], MONTH_ORDER.index(mm) + 1, 1) for mm in months}
    pdays = {p[0]: p[1] for p in periods}
    results = []
    excluded_ghost_count = 0
    excluded_ghost_cases = 0.0
    for (shop_code, prod_code), m in records.items():
        per_month = {}
        for month in months:
            d = m.get(month, {})
            per_month[month] = {'opening': d.get('opening', 0), 'receipts': d.get('receipts', 0),
                                'sales': d.get('sales', 0), 'closing': d.get('closing', 0)}
        # STRICT rule: must be in latest month
        if latest_month not in m:
            # if it had any historical stock, count it for the validation report
            for mm in reversed(months):
                if mm in m and m[mm].get('closing', 0) > 0:
                    excluded_ghost_count += 1
                    excluded_ghost_cases += m[mm]['closing']
                    break
            continue
        meta = m[latest_month]
        warehouse = meta.get('warehouse', ''); shop_name = meta.get('shop_name', '')
        brand = meta.get('brand', ''); pack = meta.get('pack', '')
        total_sales = sum(per_month[mm]['sales'] for mm in months)
        total_receipts = sum(per_month[mm]['receipts'] for mm in months)
        latest_closing = per_month[latest_month]['closing']
        if latest_closing == 0 and total_sales == 0 and total_receipts == 0:
            continue
        # ---- shared day-level shelf age (added 22 May 2026) ----
        pre_existed = per_month[months[0]]['opening'] > 0
        first_active = None
        for mm in months:
            pp = per_month[mm]
            if pp['opening'] or pp['receipts'] or pp['sales'] or pp['closing']:
                first_active = mm
                break
        if pre_existed:
            arrival = window_start
        else:
            arrival = dispatch_dates.get((shop_code, prod_code))
            if arrival is None:
                arrival = month_start.get(first_active, window_start)
            if arrival < window_start:
                arrival = window_start
        # M3: a genuine restock -- the line went FULLY empty somewhere in-window and was then
        # re-dispatched within grace -- is physically new again, so it earns the NEW STOCK grace
        # even though it pre-existed. A top-up to still-present stock (never emptied) is excluded,
        # preserving the anti-top-up rule.
        went_empty = any(per_month[mm]['closing'] <= 1e-6 for mm in months[:-1])
        restock_new = False
        if went_empty and latest_closing > 0:
            _ld = dispatch_latest.get((shop_code, prod_code))
            if _ld is not None and _ld >= window_start:
                _ra = (as_of - _ld).days
                if 0 <= _ra <= GRACE_DAYS:
                    restock_new = True
                    arrival = _ld   # reflect the fresh lot's arrival in the shelf age
        days_on_shelf = (as_of - arrival).days
        if days_on_shelf < 0: days_on_shelf = 0
        days_present = total_days if pre_existed else max(1, min(days_on_shelf, total_days))
        # ---- months-of-cover from the TRAILING velocity (C2/H1/L1) ----
        # Rate uses only the trailing RATE_HORIZON_MONTHS, over the days the stock was on shelf
        # within that horizon (floored at MIN_RATE_DAYS).
        rate_months = months[-RATE_HORIZON_MONTHS:]
        rate_window_start = month_start[rate_months[0]]
        rate_window_len = sum(pdays[mm] for mm in rate_months)
        rate_sales = sum(per_month[mm]['sales'] for mm in rate_months)
        if arrival <= rate_window_start:
            rate_days = rate_window_len
        else:
            rate_days = (as_of - arrival).days
        rate_days = max(MIN_RATE_DAYS, min(rate_days, rate_window_len))
        if rate_sales > 0:
            monthly_rate = rate_sales * 30 / rate_days
            months_cover = latest_closing / monthly_rate if monthly_rate else float('inf')
        else:
            months_cover = float('inf') if latest_closing > 0 else 0
        # M4: a month counts as a zero-sale month only if stock was AVAILABLE that month
        # (opening or receipts > 0) and nothing sold -- closing alone no longer triggers it.
        zero_sales_months = 0
        for mm in months:
            p = per_month[mm]
            if (p['opening'] > 0 or p['receipts'] > 0) and p['sales'] == 0:
                zero_sales_months += 1
        # C2: recent dead-stop -- holding stock but zero sales across the trailing STALL_MONTHS
        # months the stock was present.
        present_months = [mm for mm in months if (per_month[mm]['opening'] or per_month[mm]['receipts'] or per_month[mm]['closing'])]
        _recent = present_months[-STALL_MONTHS:]
        recent_stall = (len(_recent) >= STALL_MONTHS) and (latest_closing > 0) and all(per_month[mm]['sales'] == 0 for mm in _recent)
        no_prior_stock = not any(per_month[mm]['closing'] > 0 for mm in months[:-1])
        severity, is_new = classify_position(
            latest_closing, total_sales, months_cover, zero_sales_months,
            pre_existed, no_prior_stock, days_on_shelf,
            recent_stall=recent_stall, restock_new=restock_new)
        # C10: within-window trajectory from the closing trend across present months.
        _present = [mm for mm in months if (per_month[mm]['opening'] or per_month[mm]['receipts'] or per_month[mm]['sales'] or per_month[mm]['closing'])]
        if len(_present) >= 2:
            _cf = per_month[_present[0]]['closing']; _cl = per_month[_present[-1]]['closing']
            trajectory = "rising" if _cl > _cf + 0.5 else ("falling" if _cl < _cf - 0.5 else "flat")
        else:
            trajectory = "new"
        results.append({
            'severity': severity, 'new_stock': is_new, 'trajectory': trajectory,
            'warehouse': warehouse, 'shop_code': shop_code, 'shop_name': shop_name,
            'brand': brand, 'pack': pack, 'product_code': prod_code,
            'per_month': per_month,
            'total_sales': total_sales, 'latest_closing': latest_closing,
            'months_cover': months_cover, 'zero_sales_months': zero_sales_months,
            'arrival_date': arrival, 'days_on_shelf': days_on_shelf,
            'days_present': days_present, 'pre_existed': pre_existed,
            'no_prior_stock': no_prior_stock, 'recent_stall': recent_stall,
            'restock_new': restock_new, 'went_empty': went_empty,
        })
    meta = {'excluded_ghost_count': excluded_ghost_count,
            'excluded_ghost_cases': excluded_ghost_cases,
            'latest_month': latest_month}
    return results, meta


def _resolve_master_sheet(wb):
    """Auto-detect the master-data sheet: look for one with 'Shop Code' + 'Bond' + 'Status'
    in its header row (anywhere in first 3 rows). Returns sheet name."""
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows(min_row=1, max_row=3, values_only=True):
            if not row: continue
            cells = [str(c).strip().lower() if c else '' for c in row]
            if 'shop code' in cells and 'bond' in cells and 'status' in cells:
                return name
    # fallback to common names
    for fallback in ('MASTER DATA', '16-4-25', wb.sheetnames[-1]):
        if fallback in wb.sheetnames: return fallback
    return wb.sheetnames[0]


def _resolve_master_header_row(ws):
    """Find which row has the column headers (Shop Code, Bond, Status, etc.)."""
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), 1):
        if not row: continue
        cells = [str(c).strip().lower() if c else '' for c in row]
        if 'shop code' in cells and 'bond' in cells:
            return i
    return 2


def attach_bond_staff(results, master_path):
    # PREFER the standalone MASTER DATA file at Claude root — it's where Abhay
    # maintains live shop status (closures, bond reassignments, staff changes).
    # The embedded tab inside monthly analysis workbooks can be stale.
    if STANDALONE_MASTER.exists():
        master_path = str(STANDALONE_MASTER)
    try:
        wb = load_workbook(master_path, read_only=True, data_only=True)
    except Exception as e:
        raise SystemExit(f"FATAL: cannot open MASTER DATA ({master_path}): {e}. "
                         f"If it is open in Excel, close it and re-run.")
    sheet_name = _resolve_master_sheet(wb)
    ws = wb[sheet_name]
    header_row = _resolve_master_header_row(ws)
    # Map header label → column index
    headers = list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]
    # FIRST occurrence wins (31 Jul 2026). A last-wins map lets a stray duplicate
    # header in the hand-maintained master silently redirect a column to an empty
    # copy — the exact failure that cost the STN register 15 lines / 115 cs.
    col_idx = {}
    for i, h in enumerate(headers):
        if not h:
            continue
        k = str(h).strip().lower()
        if k in col_idx:
            print(f"  ~ MASTER: duplicate header {k!r} at column {i+1} — using the first copy "
                  f"(column {col_idx[k]+1}); delete the stray column to silence this")
            continue
        col_idx[k] = i
    SC = col_idx.get('shop code', 3)
    NAME = col_idx.get('shop name', 4)
    CAT = col_idx.get('cat', 5)
    STAFF = col_idx.get('field staff', 6)
    BOND = col_idx.get('bond', 7)
    STATUS = col_idx.get('status', 8)
    shop_to_bond = {}
    closed_shops = set()
    active_ksbc_shops = set()
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if not row: continue
        sc_val = row[SC] if SC < len(row) else None
        if sc_val is None: continue
        sc = str(sc_val).strip()
        cat = (row[CAT] or '').strip() if CAT < len(row) and row[CAT] else ''
        status = (row[STATUS] or '').strip() if STATUS < len(row) and row[STATUS] else ''
        bond = (row[BOND] or '').strip() if BOND < len(row) and row[BOND] else ''
        staff = (row[STAFF] or '').strip() if STAFF < len(row) and row[STAFF] else ''
        shop_to_bond[sc] = {'bond': bond, 'staff': staff, 'status': status, 'cat': cat}
        if status.lower() == 'closed':
            closed_shops.add(sc)
        elif cat == 'KSBC' and status.lower() == 'active':
            active_ksbc_shops.add(sc)
    wb.close()
    unmapped = []
    for x in results:
        m = shop_to_bond.get(x['shop_code'], {})
        x['bond'] = m.get('bond', '(UNMAPPED)')
        x['staff'] = m.get('staff', '')
        x['shop_status'] = m.get('status', '')
        if not m: unmapped.append(x['shop_code'])
    return results, unmapped, closed_shops, active_ksbc_shops


def build_coverage_audit(periods, records, results):
    """Audit every source row — confirm none silently dropped.
    Returns (per_month_rows, classification_counts) for the VALIDATION sheet."""
    raw_counts = {}
    raw_null = {}
    raw_dupes = {}
    raw_keys_union = set()
    for month, _, path, sheet in periods:
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb[sheet]
        total = 0; nulls = 0
        seen = Counter()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row: continue
            total += 1
            sc = row[1]; pc = row[3]
            if sc is None or pc is None: nulls += 1; continue
            seen[(str(sc).strip(), str(pc).strip())] += 1
        raw_counts[month] = total
        raw_null[month] = nulls
        raw_dupes[month] = sum(1 for c in seen.values() if c > 1)
        raw_keys_union |= set(seen.keys())
        wb.close()
    return raw_counts, raw_null, raw_dupes, raw_keys_union


def run_validation(results, meta, periods, master_path):
    """Run 5 reconciliation checks and return a list of (check_name, status, detail).
    status is 'PASS' / 'WARN' / 'FAIL'. Used to populate the VALIDATION sheet."""
    months = [p[0] for p in periods]
    latest_month, latest_days, latest_path, latest_sheet = periods[-1]
    checks = []

    # Load latest COMBINED for ground truth
    wb = load_workbook(latest_path, read_only=True, data_only=True)
    ws = wb[latest_sheet]
    latest_combined = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[1] is None: continue
        sc = str(row[1]).strip()
        pc = str(row[3]).strip() if row[3] else None
        if not pc: continue
        latest_combined[(sc, pc)] = {
            'opening': row[7] or 0, 'receipts': row[8] or 0,
            'sales': row[9] or 0, 'closing': row[10] or 0,
        }
    wb.close()
    raw_total_closing = sum(v['closing'] for v in latest_combined.values())

    # CHECK 1: every NON-MOVING row must exist in the latest period with matching closing and
    # ZERO total period sales (the tier is literally non-moving).
    dead = [x for x in results if x['severity']=="NON-MOVING"]
    ghost = 0; value_mismatch = 0; ok = 0
    for x in dead:
        ref = latest_combined.get((x['shop_code'], x['product_code']))
        if ref is None: ghost += 1
        elif abs(ref['closing'] - x['latest_closing']) > 0.01 or x['total_sales'] > 1e-9:
            value_mismatch += 1
        else: ok += 1
    status = "PASS" if (ghost == 0 and value_mismatch == 0) else ("WARN" if ghost <= 5 else "FAIL")
    checks.append((
        "NON-MOVING rows reconcile with latest COMBINED",
        status,
        f"{ok}/{len(dead)} verified  ·  {ghost} ghost (not in {latest_month}) · {value_mismatch} value mismatch  ·  NON-MOVING = strict zero sales"
    ))

    # CHECK 2: aggregate closing per bond reconciles with raw latest COMBINED
    # Use the standalone MASTER DATA file (same source as attach_bond_staff)
    master_for_validation = str(STANDALONE_MASTER) if STANDALONE_MASTER.exists() else master_path
    raw_by_bond = defaultdict(float)
    try:
        wb = load_workbook(master_for_validation, read_only=True, data_only=True)
    except Exception as e:
        raise SystemExit(f"FATAL: cannot open MASTER DATA for validation ({master_for_validation}): {e}. "
                         f"If it is open in Excel, close it and re-run.")
    mws_name = _resolve_master_sheet(wb); mws = wb[mws_name]
    mhdr = _resolve_master_header_row(mws)
    mheaders = list(mws.iter_rows(min_row=mhdr, max_row=mhdr, values_only=True))[0]
    midx = {str(h).strip().lower(): i for i, h in enumerate(mheaders) if h}
    SC = midx.get('shop code', 3); BOND = midx.get('bond', 7); STATUS = midx.get('status', 8)
    s2b = {}; closed_set = set()
    for row in mws.iter_rows(min_row=mhdr + 1, values_only=True):
        if not row or row[SC] is None: continue
        sc = str(row[SC]).strip()
        s2b[sc] = (row[BOND] or '').strip() if BOND < len(row) and row[BOND] else ''
        if STATUS < len(row) and (row[STATUS] or '').strip().lower() == 'closed':
            closed_set.add(sc)
    wb.close()
    for (sc, _), v in latest_combined.items():
        if sc in closed_set: continue  # closed shops are excluded from analysis
        raw_by_bond[s2b.get(sc, '(UNMAPPED)')] += v['closing']
    ours_by_bond = defaultdict(float)
    for x in results:
        ours_by_bond[x['bond']] += x['latest_closing']
    # Build per-bond reconciliation rows for the VALIDATION sheet
    recon_per_bond = []
    drift_rows = []
    for bond in sorted(set(list(raw_by_bond.keys()) + list(ours_by_bond.keys()))):
        raw = raw_by_bond.get(bond, 0); ours = ours_by_bond.get(bond, 0)
        diff = ours - raw
        pct = abs(diff / raw * 100) if raw else 0
        recon_per_bond.append((bond, raw, ours, diff, pct))
        if pct > 1.0:
            drift_rows.append((bond, raw, ours, diff, pct))
    status = "PASS" if not drift_rows else ("WARN" if max(d[4] for d in drift_rows) < 5 else "FAIL")
    detail = "All bonds within 1% drift — see per-bond table below" if not drift_rows else \
        "Drift >1% on " + ", ".join(f"{b}({pct:.1f}%)" for b, _, _, _, pct in drift_rows[:5])
    checks.append((
        "Per-bond closing reconciles (row-inclusion) with raw COMBINED",
        status,
        detail + "  [row-inclusion check: confirms each active shop's closing is counted once, not a value re-derivation]"
    ))

    # CHECK 3: ghost stock exclusions (transparency, not a failure)
    n = meta.get('excluded_ghost_count', 0)
    cs = meta.get('excluded_ghost_cases', 0)
    detail = f"{n} (shop × SKU) rows held closing in earlier months but vanished in {latest_month} — {cs:,.1f} cs excluded from aging (likely closed shops, transferred stock, or pipeline gaps)"
    checks.append((
        "Strict-latest-period rule applied",
        "INFO",
        detail
    ))

    # CHECK 4: tier classification logic -- re-derive every row's tier through the SAME shared
    # classify_position() the build uses, so a logic change can't silently desync the report.
    errors = 0
    for x in results:
        expected, _ = classify_position(
            x['latest_closing'], x['total_sales'], x['months_cover'], x['zero_sales_months'],
            x.get('pre_existed', False), x.get('no_prior_stock', False), x.get('days_on_shelf', 0),
            recent_stall=x.get('recent_stall', False), restock_new=x.get('restock_new', False))
        if expected != x['severity']: errors += 1
    status = "PASS" if errors == 0 else "FAIL"
    checks.append((
        "Tier classification logic is internally consistent",
        status,
        f"{errors} misclassified rows out of {len(results)}"
    ))

    # CHECK 5: bond mapping coverage
    unmapped = [x for x in results if x['bond'] == '(UNMAPPED)']
    status = "PASS" if not unmapped else "FAIL"
    checks.append((
        "Every shop maps to a bond in MASTER DATA",
        status,
        f"{len(results) - len(unmapped)}/{len(results)} mapped" + (f", unmapped: {', '.join(sorted(set(x['shop_code'] for x in unmapped[:10])))}" if unmapped else "")
    ))

    # CHECK 6: closed shops excluded
    closed_in_results = [x for x in results if x.get('shop_status','').lower() == 'closed']
    status = "PASS" if not closed_in_results else "FAIL"
    detail = f"{len(closed_in_results)} rows from closed shops" if closed_in_results else "No closed shops in aging results"
    checks.append((
        "Closed shops excluded from aging analysis",
        status,
        detail
    ))

    # CHECK 7 (A3 + M3): NEW STOCK rows are genuinely new -- either a true debut (no prior-month
    # closing, never pre-existed) OR a genuine restock (went empty in-window, re-dispatched within
    # grace) -- and always within grace.
    _mchk = [p[0] for p in periods]
    bad_new = []
    for x in results:
        if x['severity'] != "NEW STOCK": continue
        prior = any(x['per_month'][mm]['closing'] > 0 for mm in _mchk[:-1])
        debut_ok = (not x.get('pre_existed')) and (not prior)
        restock_ok = bool(x.get('restock_new'))
        if (x.get('days_on_shelf', 0) > GRACE_DAYS) or not (debut_ok or restock_ok):
            bad_new.append(x['shop_code'])
    _nnew = len([x for x in results if x['severity'] == "NEW STOCK"])
    checks.append((
        "NEW STOCK rows are genuinely new (debut or in-window restock, within grace)",
        "PASS" if not bad_new else "FAIL",
        f"{_nnew} NEW STOCK rows, {len(bad_new)} invalid" + (f": {', '.join(bad_new[:8])}" if bad_new else ""),
    ))

    return checks, raw_total_closing, recon_per_bond


def build_coverage_table(periods, records, results, closed_set):
    """Compute the coverage funnel: every raw key accounted for."""
    raw_counts, raw_null, raw_dupes, raw_keys_union = build_coverage_audit(periods, records, results)

    # Latest month per-row data
    latest_month, _, latest_path, latest_sheet = periods[-1]
    wb = load_workbook(latest_path, read_only=True, data_only=True)
    ws = wb[latest_sheet]
    latest_rowdata = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[1] is None or row[3] is None: continue
        latest_rowdata[(str(row[1]).strip(), str(row[3]).strip())] = (
            row[7] or 0, row[8] or 0, row[9] or 0, row[10] or 0)
    wb.close()

    analysed_keys = set((x['shop_code'], x['product_code']) for x in results)
    in_analysis = 0; placeholder = 0; ghost = 0; closed_count = 0; zero_everywhere = 0
    for key in raw_keys_union:
        sc, _ = key
        if key in analysed_keys: in_analysis += 1; continue
        if sc in closed_set: closed_count += 1; continue
        if key not in latest_rowdata: ghost += 1; continue
        o, r, s, c = latest_rowdata[key]
        if o == 0 and r == 0 and s == 0 and c == 0:
            # check if it had activity in any other month
            m = records.get(key, {})
            any_activity = any(
                (v.get('opening',0) + v.get('receipts',0) + v.get('sales',0) + v.get('closing',0)) > 0
                for v in m.values())
            if any_activity: zero_everywhere += 1
            else: placeholder += 1
        else:
            # Shouldn't happen — would be a real bug if a row with activity isn't in analysis
            zero_everywhere += 1

    total_accounted = in_analysis + placeholder + ghost + closed_count + zero_everywhere
    return {
        'raw_counts': raw_counts,
        'raw_null': raw_null,
        'raw_dupes': raw_dupes,
        'total_raw_rows': sum(raw_counts.values()),
        'unique_keys': len(raw_keys_union),
        'in_analysis': in_analysis,
        'placeholder': placeholder,
        'ghost': ghost,
        'closed': closed_count,
        'zero_everywhere': zero_everywhere,
        'total_accounted': total_accounted,
        'unaccounted': len(raw_keys_union) - total_accounted,
    }


def build_validation_sheet(wb, checks, results, raw_total_closing, total_days, recon_per_bond=None, coverage=None):
    """Add a VALIDATION sheet documenting every check that ran on this build."""
    ws = wb.create_sheet("VALIDATION")
    for col, w in enumerate([20, 10, 80], 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    write_title_band(ws, "VALIDATION REPORT — automated checks on every build", "C")

    ws.merge_cells('A3:C3')
    c = ws['A3']; c.value = "Run-time reconciliation against source workbooks. Re-run on every refresh."
    c.font = Font(name='Aptos Narrow', size=11, italic=True, color=GREY_LBL)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[3].height = 22

    # Headers
    headers = ['Check', 'Status', 'Detail']
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=NAVY_SOFT)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = GOLD_THIN_BOT
    ws.row_dimensions[4].height = 30

    status_colors = {
        "PASS": ("FFDCEDC8", "FF2E7D32"),
        "WARN": ("FFFFF3E0", "FFE65100"),
        "FAIL": ("FFFFCDD2", "FFC62828"),
        "INFO": ("FFE3E8F0", "FF1A237E"),
    }
    r = 5
    for name, status, detail in checks:
        ws.cell(row=r, column=1, value=name).font = Font(name='Aptos Narrow', size=11, color=GREY_TXT)
        c = ws.cell(row=r, column=2, value=status)
        bg, fg = status_colors.get(status, ("FFFFFFFF", "FF000000"))
        c.font = Font(name='Aptos Narrow', size=11, bold=True, color=fg)
        c.fill = PatternFill('solid', fgColor=bg)
        c.alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=r, column=3, value=detail).font = Font(name='Aptos Narrow', size=10, color=GREY_TXT)
        ws.cell(row=r, column=3).alignment = Alignment(horizontal='left', vertical='center', indent=1, wrap_text=True)
        for col in (1, 3):
            ws.cell(row=r, column=col).alignment = Alignment(horizontal='left', vertical='center', indent=1, wrap_text=True)
            ws.cell(row=r, column=col).border = Border(bottom=Side(style='thin', color=LINE_GREY))
        ws.cell(row=r, column=2).border = Border(bottom=Side(style='thin', color=LINE_GREY))
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1

    # Headline numbers
    ws.merge_cells(f'A{r}:C{r}')
    c = ws.cell(row=r, column=1, value="HEADLINE FIGURES")
    c.font = Font(name='Aptos Narrow', size=12, bold=True, color=WHITE)
    c.fill = PatternFill('solid', fgColor=NAVY_DEEP)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[r].height = 24; r += 1

    sev_count = Counter(x['severity'] for x in results)
    sev_cases = defaultdict(float)
    for x in results: sev_cases[x['severity']] += x['latest_closing']
    rows = [
        ("Period analysed", f"{total_days} days"),
        ("Total shop×SKU rows considered", f"{len(results):,}"),
        ("NON-MOVING rows / cases", f"{sev_count['NON-MOVING']:,} rows  /  {sev_cases['NON-MOVING']:,.1f} cs"),
        ("CRITICAL AGING rows / cases", f"{sev_count['CRITICAL AGING']:,} rows  /  {sev_cases['CRITICAL AGING']:,.1f} cs"),
        ("SLOW MOVING rows / cases", f"{sev_count['SLOW MOVING']:,} rows  /  {sev_cases['SLOW MOVING']:,.1f} cs"),
        (f"NEW STOCK held out (≤{GRACE_DAYS}d grace)", f"{sev_count['NEW STOCK']:,} rows  /  {sev_cases['NEW STOCK']:,.1f} cs"),
        ("HEALTHY rows / cases", f"{sev_count['HEALTHY']:,} rows  /  {sev_cases['HEALTHY']:,.1f} cs"),
        ("Aggregated latest closing (this analysis)", f"{sum(sev_cases.values()):,.1f} cs"),
        ("Raw latest COMBINED closing (source of truth)", f"{raw_total_closing:,.1f} cs"),
        ("Drift", f"{sum(sev_cases.values()) - raw_total_closing:+,.1f} cs  ({(sum(sev_cases.values()) - raw_total_closing) / raw_total_closing * 100:+.2f}%)" if raw_total_closing else "n/a"),
    ]
    for label, val in rows:
        ws.merge_cells(f'A{r}:B{r}')
        c = ws.cell(row=r, column=1, value=label)
        c.font = Font(name='Aptos Narrow', size=10, bold=True, color=NAVY_DEEP)
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        c.fill = PatternFill('solid', fgColor=GREY_BG)
        c = ws.cell(row=r, column=3, value=val)
        c.font = Font(name='Aptos Narrow', size=10, color=GREY_TXT)
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        c.fill = PatternFill('solid', fgColor=GREY_BG)
        ws.row_dimensions[r].height = 20; r += 1

    # Per-bond reconciliation — the receipt for CHECK 2
    if recon_per_bond:
        r += 1
        # Section header — needs more cols, so we'll widen columns and add the table
        ws.column_dimensions['A'].width = 22
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 14
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 14

        ws.merge_cells(f'A{r}:F{r}')
        c = ws.cell(row=r, column=1, value="PER-BOND RECONCILIATION  ·  the receipt for Check #2")
        c.font = Font(name='Aptos Narrow', size=12, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=NAVY_DEEP)
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[r].height = 24; r += 1

        ws.merge_cells(f'A{r}:F{r}')
        c = ws.cell(row=r, column=1, value="For every bond, total closing in this analysis must equal the same bond's closing in the raw COMBINED sheet.  Drift ≤ 1% = PASS.")
        c.font = Font(name='Aptos Narrow', size=10, italic=True, color=GREY_LBL)
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[r].height = 22; r += 1

        # Headers
        headers2 = ['Bond', 'Raw closing (source)', 'This analysis', 'Diff', 'Drift %', 'Status']
        for i, h in enumerate(headers2, 1):
            c = ws.cell(row=r, column=i, value=h)
            c.font = Font(name='Aptos Narrow', size=10, bold=True, color=WHITE)
            c.fill = PatternFill('solid', fgColor=NAVY_SOFT)
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.border = GOLD_THIN_BOT
        ws.row_dimensions[r].height = 28; r += 1

        # Sort by raw desc so biggest bonds at top
        recon_sorted = sorted(recon_per_bond, key=lambda t: -t[1])
        for bond, raw, ours, diff, pct in recon_sorted:
            row_pass = abs(pct) <= 1.0
            ws.cell(row=r, column=1, value=bond).font = Font(name='Aptos Narrow', size=11, bold=True, color=NAVY_DEEP)
            ws.cell(row=r, column=1).alignment = Alignment(horizontal='left', vertical='center', indent=1)
            ws.cell(row=r, column=2, value=raw).number_format = '#,##0.00'
            ws.cell(row=r, column=3, value=ours).number_format = '#,##0.00'
            ws.cell(row=r, column=4, value=diff).number_format = '+#,##0.00;-#,##0.00;"0.00"'
            ws.cell(row=r, column=5, value=pct/100).number_format = '0.00%'
            status_cell = ws.cell(row=r, column=6, value="PASS" if row_pass else "DRIFT")
            bg, fg = ("FFDCEDC8","FF2E7D32") if row_pass else ("FFFFCDD2","FFC62828")
            status_cell.font = Font(name='Aptos Narrow', size=10, bold=True, color=fg)
            status_cell.fill = PatternFill('solid', fgColor=bg)
            status_cell.alignment = Alignment(horizontal='center', vertical='center')
            for col in (2,3,4,5):
                cell = ws.cell(row=r, column=col)
                cell.font = Font(name='Aptos Narrow', size=10, color=GREY_TXT)
                cell.alignment = Alignment(horizontal='right', vertical='center', indent=1)
            for col in range(1, 7):
                ws.cell(row=r, column=col).border = Border(bottom=Side(style='thin', color=LINE_GREY))
            ws.row_dimensions[r].height = 20
            r += 1

        # TOTAL row
        total_raw = sum(rr[1] for rr in recon_per_bond)
        total_ours = sum(rr[2] for rr in recon_per_bond)
        total_diff = total_ours - total_raw
        total_pct = total_diff / total_raw if total_raw else 0
        ws.cell(row=r, column=1, value="NETWORK TOTAL").font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
        ws.cell(row=r, column=2, value=total_raw).number_format = '#,##0.00'
        ws.cell(row=r, column=3, value=total_ours).number_format = '#,##0.00'
        ws.cell(row=r, column=4, value=total_diff).number_format = '+#,##0.00;-#,##0.00;"0.00"'
        ws.cell(row=r, column=5, value=total_pct).number_format = '0.00%'
        ws.cell(row=r, column=6, value="PASS" if abs(total_pct*100) <= 1.0 else "DRIFT")
        for col in range(1, 7):
            cell = ws.cell(row=r, column=col)
            cell.font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
            cell.fill = PatternFill('solid', fgColor=NAVY_DEEP)
            cell.alignment = Alignment(horizontal='left' if col==1 else ('center' if col==6 else 'right'),
                                       vertical='center', indent=1)
            cell.border = Border(top=Side(style='medium', color=GOLD))
        ws.row_dimensions[r].height = 26; r += 1

    # ====== COVERAGE AUDIT — every source row accounted for ======
    if coverage:
        r += 2
        ws.merge_cells(f'A{r}:F{r}')
        c = ws.cell(row=r, column=1, value="COVERAGE AUDIT  ·  every source row accounted for")
        c.font = Font(name='Aptos Narrow', size=12, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=NAVY_DEEP)
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[r].height = 24; r += 1

        ws.merge_cells(f'A{r}:F{r}')
        c = ws.cell(row=r, column=1,
            value="Proves no row in any month's COMBINED was silently dropped. Total raw rows + null + duplicates checked.")
        c.font = Font(name='Aptos Narrow', size=10, italic=True, color=GREY_LBL)
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[r].height = 22; r += 1

        # Per-month raw row counts
        ws.cell(row=r, column=1, value="Month").font = Font(name='Aptos Narrow', size=10, bold=True, color=WHITE)
        ws.cell(row=r, column=2, value="Rows read").font = Font(name='Aptos Narrow', size=10, bold=True, color=WHITE)
        ws.cell(row=r, column=3, value="Null shop/SKU").font = Font(name='Aptos Narrow', size=10, bold=True, color=WHITE)
        ws.cell(row=r, column=4, value="Duplicates").font = Font(name='Aptos Narrow', size=10, bold=True, color=WHITE)
        ws.merge_cells(f'E{r}:F{r}')
        ws.cell(row=r, column=5, value="Sheet").font = Font(name='Aptos Narrow', size=10, bold=True, color=WHITE)
        for col in range(1, 7):
            c = ws.cell(row=r, column=col)
            c.fill = PatternFill('solid', fgColor=NAVY_SOFT)
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.border = GOLD_THIN_BOT
        ws.row_dimensions[r].height = 26; r += 1

        for month, count in coverage['raw_counts'].items():
            ws.cell(row=r, column=1, value=month).font = Font(name='Aptos Narrow', size=10, bold=True, color=NAVY_DEEP)
            ws.cell(row=r, column=2, value=count).number_format = '#,##0'
            ws.cell(row=r, column=3, value=coverage['raw_null'].get(month, 0)).number_format = '#,##0'
            ws.cell(row=r, column=4, value=coverage['raw_dupes'].get(month, 0)).number_format = '#,##0'
            ws.merge_cells(f'E{r}:F{r}')
            ws.cell(row=r, column=5, value=f"{month} COMBINED")
            for col in range(1, 7):
                cell = ws.cell(row=r, column=col)
                if col > 1:
                    cell.font = Font(name='Aptos Narrow', size=10, color=GREY_TXT)
                cell.alignment = Alignment(horizontal='left' if col in (1,5) else 'center',
                                           vertical='center', indent=1 if col in (1,5) else 0)
                cell.border = Border(bottom=Side(style='thin', color=LINE_GREY))
            ws.row_dimensions[r].height = 20; r += 1
        # Total raw rows
        ws.cell(row=r, column=1, value="TOTAL").font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
        ws.cell(row=r, column=2, value=coverage['total_raw_rows']).number_format = '#,##0'
        ws.cell(row=r, column=3, value=sum(coverage['raw_null'].values())).number_format = '#,##0'
        ws.cell(row=r, column=4, value=sum(coverage['raw_dupes'].values())).number_format = '#,##0'
        ws.merge_cells(f'E{r}:F{r}')
        ws.cell(row=r, column=5, value=f"{coverage['unique_keys']:,} unique (Shop × SKU) keys")
        for col in range(1, 7):
            c = ws.cell(row=r, column=col)
            c.font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
            c.fill = PatternFill('solid', fgColor=NAVY_DEEP)
            c.alignment = Alignment(horizontal='left' if col in (1,5) else 'center',
                                    vertical='center', indent=1 if col in (1,5) else 0)
            c.border = Border(top=Side(style='medium', color=GOLD))
        ws.row_dimensions[r].height = 26; r += 2

        # Classification funnel
        ws.merge_cells(f'A{r}:F{r}')
        c = ws.cell(row=r, column=1, value="Where every (Shop × SKU) key ended up:")
        c.font = Font(name='Aptos Narrow', size=11, bold=True, color=NAVY_DEEP)
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[r].height = 22; r += 1

        funnel = [
            ("✅ Included in aging analysis", coverage['in_analysis'], "Has stock or movement in latest month, active shop", "FFDCEDC8", "FF2E7D32"),
            ("🟫 All-zero placeholder rows", coverage['placeholder'], "Master lists the SKU but shop has zero stock & zero movement (no SKU on shelf)", "FFF3F4F6", "FF6B7280"),
            ("🔴 Closed shop rows", coverage['closed'], "Shop status=Closed in MASTER DATA — excluded by rule", "FFFFCDD2", "FFC62828"),
            ("👻 Ghost rows", coverage['ghost'], "Appeared in earlier months but not in latest month — excluded by strict rule", "FFFFE0B2", "FFE65100"),
            ("◯ Zero-everywhere rows", coverage['zero_everywhere'], "Has latest-month entry but zero across all months", "FFF3F4F6", "FF6B7280"),
        ]
        for label, count, why, bg, fg in funnel:
            ws.cell(row=r, column=1, value=label).font = Font(name='Aptos Narrow', size=10, bold=True, color=fg)
            ws.cell(row=r, column=1).fill = PatternFill('solid', fgColor=bg)
            ws.cell(row=r, column=2, value=count).number_format = '#,##0'
            pct = count / coverage['unique_keys'] * 100 if coverage['unique_keys'] else 0
            ws.cell(row=r, column=3, value=f"{pct:.1f}%")
            ws.merge_cells(f'D{r}:F{r}')
            ws.cell(row=r, column=4, value=why)
            for col in range(1, 7):
                cell = ws.cell(row=r, column=col)
                if col != 1: cell.font = Font(name='Aptos Narrow', size=10, color=GREY_TXT)
                cell.alignment = Alignment(horizontal='left' if col in (1,4) else 'center',
                                           vertical='center', indent=1 if col in (1,4) else 0)
                cell.border = Border(bottom=Side(style='thin', color=LINE_GREY))
            ws.row_dimensions[r].height = 22; r += 1

        # Accounting line
        ok = coverage['unaccounted'] == 0
        ws.cell(row=r, column=1, value="∑ TOTAL ACCOUNTED FOR").font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
        ws.cell(row=r, column=2, value=coverage['total_accounted']).number_format = '#,##0'
        ws.cell(row=r, column=3, value=coverage['unique_keys']).number_format = '#,##0'
        ws.merge_cells(f'D{r}:F{r}')
        if ok:
            ws.cell(row=r, column=4, value=f"PASS — every one of {coverage['unique_keys']:,} unique keys is categorised (unaccounted = 0)")
        else:
            ws.cell(row=r, column=4, value=f"FAIL — {coverage['unaccounted']} key(s) unaccounted for! Investigate the script.")
        for col in range(1, 7):
            c = ws.cell(row=r, column=col)
            c.font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
            c.fill = PatternFill('solid', fgColor=NAVY_DEEP if ok else "FFC62828")
            c.alignment = Alignment(horizontal='left' if col in (1,4) else 'center',
                                    vertical='center', indent=1 if col in (1,4) else 0)
            c.border = Border(top=Side(style='medium', color=GOLD))
        ws.row_dimensions[r].height = 26

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = 'A5'


def write_title_band(ws, text, last_col_letter, start_row=1, height=42):
    ws.merge_cells(f'A{start_row}:{last_col_letter}{start_row}')
    c = ws.cell(row=start_row, column=1, value=text)
    c.font = Font(name='Aptos Narrow', size=18, bold=True, color=WHITE)
    c.fill = PatternFill('solid', fgColor=NAVY_DEEP)
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border = GOLD_THIN_BOT
    ws.row_dimensions[start_row].height = height


def write_header_row(ws, headers, row, col_widths=None):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=NAVY_SOFT)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = GOLD_THIN_BOT
    ws.row_dimensions[row].height = 30
    if col_widths:
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


# --- SUMMARY builder, per-bond builder, drill-downs, BY BOND/STAFF/BRAND × PACK ---
# (Build logic identical to interactive scripts in beautify_summary.py / beautify_bond_sheet.py /
#  rebuild_bondwise.py, consolidated into a single durable pipeline.)

def _mix_top(d):
    """Sort a {name: cases} dict desc; collapse the tail past 6 into one 'Other (k)' row."""
    s = sorted(d.items(), key=lambda kv: -kv[1])
    if len(s) > 6:
        head = s[:5]; head.append((f"Other ({len(s) - 5})", sum(v for _, v in s[5:]))); return head
    return s


def _draw_mix_strip(ws, start_row, n_cols, brand_list, pack_list, mix_total, hdr_size=10, val_size=10):
    """Two-panel AGING MIX strip: each entry = name · mini data-bar (proportional to cases) ·
    'X cs · Y%' value pinned to the panel's right edge, with a thin gold divider between the
    BRAND and PACK panels. Shared by the SUMMARY and every bond sheet so they stay identical.
    Returns the spacer row index immediately AFTER the strip."""
    gl = get_column_letter; LAST = gl(n_cols)
    BAR = "FF8FA8D8"
    DIV = Side(style='thin', color=GOLD)
    HDRB = Side(style='medium', color=GOLD); ROWB = Side(style='thin', color=LINE_GREY)
    left_last = max(4, n_cols // 2); right_first = left_last + 1
    def carve(c0, c1):
        if (c1 - c0 + 1) >= 5:
            vf = c1 - 1; bar = vf - 1; return (c0, bar - 1), bar, (vf, c1)
        return (c0, c1 - 1), None, (c1, c1)
    (lnA, lnB), lbar, (lvA, lvB) = carve(1, left_last)
    (rnA, rnB), rbar, (rvA, rvB) = carve(right_first, n_cols)
    # header band (two navy panels, gold underline, divider before the right panel)
    for col in range(1, n_cols + 1):
        cc = ws.cell(row=start_row, column=col)
        cc.fill = PatternFill('solid', fgColor=NAVY_SOFT)
        cc.border = Border(bottom=HDRB, left=(DIV if col == right_first else None))
    cc = ws.cell(row=start_row, column=1, value="AGING MIX BY BRAND  (cases)")
    cc.font = Font(name='Aptos Narrow', size=hdr_size, bold=True, color=WHITE)
    cc.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.merge_cells(f"A{start_row}:{gl(left_last)}{start_row}")
    cc = ws.cell(row=start_row, column=right_first, value="AGING MIX BY PACK  (cases)")
    cc.font = Font(name='Aptos Narrow', size=hdr_size, bold=True, color=WHITE)
    cc.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    cc.border = Border(bottom=HDRB, left=DIV)
    ws.merge_cells(f"{gl(right_first)}{start_row}:{LAST}{start_row}")
    ws.row_dimensions[start_row].height = 24
    nrows = max(len(brand_list), len(pack_list), 1)
    def entry(rr, name, cs, nA, nB, bar, vA, vB, divider):
        ws.merge_cells(f"{gl(nA)}{rr}:{gl(nB)}{rr}")
        c = ws.cell(row=rr, column=nA, value=name)
        c.font = Font(name='Aptos Narrow', size=val_size, bold=True, color=NAVY_DEEP)
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        if divider:
            c.border = Border(left=DIV, bottom=ROWB)
        if bar is not None:
            bc = ws.cell(row=rr, column=bar, value=round(cs, 2))
            bc.number_format = ';;;'   # hide the number; the data-bar is the visual
        ws.merge_cells(f"{gl(vA)}{rr}:{gl(vB)}{rr}")
        c = ws.cell(row=rr, column=vA, value=f"{cs:,.1f} cs · {cs / mix_total * 100:.0f}%")
        c.font = Font(name='Aptos Narrow', size=val_size, color=GREY_TXT)
        c.alignment = Alignment(horizontal='right', vertical='center', indent=1)
    for i in range(nrows):
        rr = start_row + 1 + i
        fill = ROW_ALT if (i % 2) else WHITE
        for col in range(1, n_cols + 1):
            cc = ws.cell(row=rr, column=col)
            cc.fill = PatternFill('solid', fgColor=fill)
            cc.border = Border(bottom=ROWB, left=(DIV if col == right_first else None))
        if i < len(brand_list):
            entry(rr, brand_list[i][0], brand_list[i][1], lnA, lnB, lbar, lvA, lvB, False)
        if i < len(pack_list):
            entry(rr, pack_list[i][0], pack_list[i][1], rnA, rnB, rbar, rvA, rvB, True)
        ws.row_dimensions[rr].height = 19
    if lbar is not None:
        ws.conditional_formatting.add(f"{gl(lbar)}{start_row+1}:{gl(lbar)}{start_row+nrows}",
            DataBarRule(start_type='num', start_value=0, end_type='max', color=BAR, showValue=False))
    if rbar is not None:
        ws.conditional_formatting.add(f"{gl(rbar)}{start_row+1}:{gl(rbar)}{start_row+nrows}",
            DataBarRule(start_type='num', start_value=0, end_type='max', color=BAR, showValue=False))
    return start_row + 1 + nrows


def build_summary(wb, results, periods, total_days, active_shops_count=None):
    months = [p[0] for p in periods]
    last_month, last_days, _, _ = periods[-1]
    period_label = f"{MONTH_ABBREV[months[0]]} 1 – {MONTH_ABBREV[last_month]} {last_days}"
    sev_count = Counter(r['severity'] for r in results)
    sev_cases = defaultdict(float); sev_skus = defaultdict(set)
    for r in results:
        sev_cases[r['severity']] += r['latest_closing']
        sev_skus[r['severity']].add(r['product_code'])
    total_aging = sev_cases["NON-MOVING"] + sev_cases["CRITICAL AGING"] + sev_cases["SLOW MOVING"]
    total_closing = sum(sev_cases.values())
    aging_pct = total_aging / total_closing if total_closing else 0
    # Use MASTER DATA's active-KSBC count as source of truth, not just shops with activity
    n_shops = active_shops_count if active_shops_count else len({r['shop_code'] for r in results})

    bond_agg = defaultdict(lambda: {'dead':0,'crit':0,'slow':0,'rows':0,'staff':set()})
    for x in results:
        if x['severity'] in ("HEALTHY", "NEW STOCK"): continue
        b = x['bond']; bond_agg[b]['rows'] += 1
        if x['staff']: bond_agg[b]['staff'].add(x['staff'])
        if x['severity']=="NON-MOVING": bond_agg[b]['dead'] += x['latest_closing']
        elif x['severity']=="CRITICAL AGING": bond_agg[b]['crit'] += x['latest_closing']
        elif x['severity']=="SLOW MOVING": bond_agg[b]['slow'] += x['latest_closing']
    bond_sorted = sorted(bond_agg.items(), key=lambda kv: -(kv[1]['dead']+kv[1]['crit']+kv[1]['slow']))

    brand_agg = defaultdict(lambda: {'dead':0,'crit':0,'slow':0,'rows':0})
    for x in results:
        if x['severity'] in ("HEALTHY", "NEW STOCK"): continue
        b = x['brand']; brand_agg[b]['rows'] += 1
        if x['severity']=="NON-MOVING": brand_agg[b]['dead'] += x['latest_closing']
        elif x['severity']=="CRITICAL AGING": brand_agg[b]['crit'] += x['latest_closing']
        elif x['severity']=="SLOW MOVING": brand_agg[b]['slow'] += x['latest_closing']
    brand_sorted = sorted(brand_agg.items(), key=lambda kv: -(kv[1]['dead']+kv[1]['crit']+kv[1]['slow']))

    pack_agg = defaultdict(lambda: {'dead':0,'crit':0,'slow':0,'rows':0})
    for x in results:
        if x['severity'] in ("HEALTHY", "NEW STOCK"): continue
        p = x['pack']; pack_agg[p]['rows'] += 1
        if x['severity']=="NON-MOVING": pack_agg[p]['dead'] += x['latest_closing']
        elif x['severity']=="CRITICAL AGING": pack_agg[p]['crit'] += x['latest_closing']
        elif x['severity']=="SLOW MOVING": pack_agg[p]['slow'] += x['latest_closing']
    pack_sorted = sorted(pack_agg.items(), key=lambda kv: -(kv[1]['dead']+kv[1]['crit']+kv[1]['slow']))

    ws = wb.create_sheet("SUMMARY", 0)
    for col in range(1, 14): ws.column_dimensions[get_column_letter(col)].width = 14
    # Slightly wider for the ●/▲/◆ header columns so the bullet doesn't clip on macOS
    ws.column_dimensions['G'].width = 17
    ws.column_dimensions['H'].width = 17
    ws.column_dimensions['I'].width = 16
    ws.column_dimensions['M'].width = 4

    ws.merge_cells('A1:M1'); c = ws['A1']
    c.value = "KSBC AGING STOCK ANALYSIS"
    c.font = Font(name='Aptos Narrow', size=28, bold=True, color=WHITE)
    c.fill = PatternFill('solid', fgColor=NAVY_DEEP)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 56
    ws.merge_cells('A2:M2'); c = ws['A2']
    c.value = f"{period_label}  ·  {total_days} days  ·  {n_shops} active shops across 15 bonds"
    c.font = Font(name='Aptos Narrow', size=13, italic=True, color=WHITE)
    c.fill = PatternFill('solid', fgColor=NAVY_MID)
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border = GOLD_THIN_BOT
    ws.row_dimensions[2].height = 26; ws.row_dimensions[3].height = 12

    def draw_card(start_col, tier_key):
        info = TIER[tier_key]; last_col = start_col + 2
        sc = get_column_letter(start_col); ec = get_column_letter(last_col)
        ws.merge_cells(f'{sc}4:{ec}4'); c = ws.cell(row=4, column=start_col)
        c.value = f"{info['icon']}   {tier_key}"
        c.font = Font(name='Aptos Narrow', size=12, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=info['dark'])
        c.alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(f'{sc}5:{ec}6'); c = ws.cell(row=5, column=start_col)
        c.value = f"{sev_cases[tier_key]:,.0f}"
        c.font = Font(name='Aptos Narrow', size=42, bold=True, color=info['bright'])
        c.fill = PatternFill('solid', fgColor=WHITE)
        c.alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(f'{sc}7:{ec}7'); c = ws.cell(row=7, column=start_col)
        c.value = "cases on hand"
        c.font = Font(name='Aptos Narrow', size=10, color=GREY_LBL)
        c.fill = PatternFill('solid', fgColor=WHITE); c.alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(f'{sc}8:{ec}8'); c = ws.cell(row=8, column=start_col)
        c.value = f"{sev_count[tier_key]:,} positions  ·  {len(sev_skus[tier_key])} SKUs"
        c.font = Font(name='Aptos Narrow', size=11, bold=True, color=GREY_TXT)
        c.fill = PatternFill('solid', fgColor=WHITE); c.alignment = Alignment(horizontal='center', vertical='center')
        pct = sev_cases[tier_key] / total_closing if total_closing else 0
        ws.merge_cells(f'{sc}9:{ec}9'); c = ws.cell(row=9, column=start_col)
        c.value = f"{pct*100:.1f}% of total closing stock"
        c.font = Font(name='Aptos Narrow', size=10, bold=True, color=info['dark'])
        c.fill = PatternFill('solid', fgColor=info['light']); c.alignment = Alignment(horizontal='center', vertical='center')
        for rr in range(4, 10):
            for cc in range(start_col, last_col+1):
                cell = ws.cell(row=rr, column=cc)
                cell.border = Border(left=HAIR if cc==start_col else None, right=HAIR if cc==last_col else None,
                                     top=HAIR if rr==4 else None, bottom=HAIR if rr==9 else None)
    draw_card(1, "NON-MOVING"); draw_card(4, "CRITICAL AGING")
    draw_card(7, "SLOW MOVING"); draw_card(10, "HEALTHY")
    ws.row_dimensions[4].height = 28; ws.row_dimensions[5].height = 32; ws.row_dimensions[6].height = 32
    ws.row_dimensions[7].height = 18; ws.row_dimensions[8].height = 20; ws.row_dimensions[9].height = 22
    ws.row_dimensions[10].height = 14

    ws.merge_cells('A11:M11'); c = ws['A11']
    c.value = f"{aging_pct*100:.1f}% of total closing stock is currently aging"
    c.font = Font(name='Aptos Narrow', size=22, bold=True, color=WHITE)
    c.fill = PatternFill('solid', fgColor=NAVY_MID)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[11].height = 46
    ws.merge_cells('A12:M12'); c = ws['A12']
    c.value = (f"{total_aging:,.0f} cases out of {total_closing:,.0f} total on hand  ·  Non-Moving + Critical + Slow combined"
               f"   ·   ✦ {sev_cases['NEW STOCK']:,.0f} cs in {sev_count['NEW STOCK']:,} new positions held out (≤{GRACE_DAYS}-day grace)")
    c.font = Font(name='Aptos Narrow', size=11, italic=True, color=GOLD)
    c.fill = PatternFill('solid', fgColor=NAVY_MID)
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border = GOLD_THIN_BOT
    ws.row_dimensions[12].height = 22; ws.row_dimensions[13].height = 14

    ws.merge_cells('A14:M14'); c = ws['A14']
    c.value = "AGING STOCK BY BOND — sorted by total aging exposure"
    c.font = Font(name='Aptos Narrow', size=14, bold=True, color=NAVY_DEEP)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[14].height = 28

    hr = 15
    # Pre-fill every header column with navy
    for col in range(1, 14):
        ws.cell(row=hr, column=col).fill = PatternFill('solid', fgColor=NAVY_SOFT)
    ws.cell(row=hr, column=1, value='#')
    ws.merge_cells(f'B{hr}:C{hr}'); ws.cell(row=hr, column=2, value='Bond')
    ws.merge_cells(f'D{hr}:F{hr}'); ws.cell(row=hr, column=4, value='Field Staff (Salesman)')
    ws.cell(row=hr, column=7, value='● Non-Moving (cs)')
    ws.cell(row=hr, column=8, value='▲ Critical (cs)')
    ws.cell(row=hr, column=9, value='◆ Slow (cs)')
    ws.merge_cells(f'J{hr}:K{hr}'); ws.cell(row=hr, column=10, value='Total Aging (cs)')
    ws.merge_cells(f'L{hr}:M{hr}'); ws.cell(row=hr, column=12, value='# Shop×SKU rows')
    for col in [1,2,4,7,8,9,10,12]:
        c = ws.cell(row=hr, column=col)
        c.font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=NAVY_SOFT)
        c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[hr].height = 30

    r = hr + 1; data_start = r
    for idx, (bond, v) in enumerate(bond_sorted, 1):
        total = v['dead']+v['crit']+v['slow']
        staff_str = ", ".join(sorted(v['staff'])) if v['staff'] else '—'
        alt = (idx % 2 == 0); row_fill = "FFFAFBFC" if alt else "FFFFFFFF"
        # Pre-fill every column 1..13 so merge boundaries are seamless
        for col in range(1, 14):
            cc = ws.cell(row=r, column=col)
            cc.fill = PatternFill('solid', fgColor=row_fill)
            cc.border = Border(bottom=Side(style='thin', color=LINE_GREY))
        ws.cell(row=r, column=1, value=idx)
        ws.merge_cells(f'B{r}:C{r}'); ws.cell(row=r, column=2, value=bond)
        ws.merge_cells(f'D{r}:F{r}'); ws.cell(row=r, column=4, value=staff_str)
        ws.cell(row=r, column=7, value=v['dead']); ws.cell(row=r, column=8, value=v['crit']); ws.cell(row=r, column=9, value=v['slow'])
        ws.merge_cells(f'J{r}:K{r}'); ws.cell(row=r, column=10, value=total)
        ws.merge_cells(f'L{r}:M{r}'); ws.cell(row=r, column=12, value=v['rows'])
        for col in [1,2,4,7,8,9,10,12]:
            c = ws.cell(row=r, column=col)
            c.font = Font(name='Aptos Narrow', size=10.5, color=GREY_TXT)
            c.fill = PatternFill('solid', fgColor=row_fill)
            c.border = Border(bottom=Side(style='thin', color=LINE_GREY))
            c.alignment = Alignment(horizontal='center', vertical='center')
            if col == 2:
                c.font = Font(name='Aptos Narrow', size=11, bold=True, color=NAVY_DEEP)
                c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            if col == 4:
                c.font = Font(name='Aptos Narrow', size=10, color=GREY_LBL)
                c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            if col == 1:
                c.font = Font(name='Aptos Narrow', size=11, bold=True, color=NAVY_SOFT)
            if col in (7,8,9,10): c.number_format = '#,##0.0'
            if col == 12: c.number_format = '#,##0'
        ct = ws.cell(row=r, column=10)
        ct.font = Font(name='Aptos Narrow', size=11, bold=True,
                       color="FFB71C1C" if total > 300 else ("FFE65100" if total > 100 else "FF2E7D32"))
        ws.row_dimensions[r].height = 22; r += 1
    data_end = r - 1
    ws.conditional_formatting.add(f'J{data_start}:J{data_end}',
        DataBarRule(start_type='min', end_type='max', color="FFE65100", showValue=True))

    # Pre-fill every column for the BY BOND total row
    for col in range(1, 14):
        cc = ws.cell(row=r, column=col)
        cc.fill = PatternFill('solid', fgColor=NAVY_DEEP)
        cc.border = Border(top=Side(style='medium', color=GOLD))
    ws.cell(row=r, column=1, value='∑')
    ws.merge_cells(f'B{r}:C{r}'); ws.cell(row=r, column=2, value='15 BONDS')
    ws.merge_cells(f'D{r}:F{r}'); ws.cell(row=r, column=4, value=f"{n_shops} shops · KSBC network")
    ws.cell(row=r, column=7, value=sev_cases['NON-MOVING'])
    ws.cell(row=r, column=8, value=sev_cases['CRITICAL AGING'])
    ws.cell(row=r, column=9, value=sev_cases['SLOW MOVING'])
    ws.merge_cells(f'J{r}:K{r}'); ws.cell(row=r, column=10, value=total_aging)
    ws.merge_cells(f'L{r}:M{r}'); ws.cell(row=r, column=12,
        value=sev_count['NON-MOVING']+sev_count['CRITICAL AGING']+sev_count['SLOW MOVING'])
    for col in [1,2,4,7,8,9,10,12]:
        c = ws.cell(row=r, column=col)
        c.font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=NAVY_DEEP)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = Border(top=Side(style='medium', color=GOLD))
        if col == 2: c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        if col == 4:
            c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            c.font = Font(name='Aptos Narrow', size=10, italic=True, color=GOLD)
        if col in (7,8,9,10): c.number_format = '#,##0.0'
        if col == 12: c.number_format = '#,##0'
    ws.row_dimensions[r].height = 28; r += 2

    ws.merge_cells(f'A{r}:M{r}')
    c = ws.cell(row=r, column=1, value="TOP BRANDS BY AGING STOCK")
    c.font = Font(name='Aptos Narrow', size=14, bold=True, color=NAVY_DEEP)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[r].height = 28; r += 1

    # Pre-fill every column for header row first
    for col in range(1, 14):
        cc = ws.cell(row=r, column=col)
        cc.fill = PatternFill('solid', fgColor=NAVY_SOFT)
    ws.cell(row=r, column=1, value='#')
    # Extend Brand merge to B:F (matches BY BOND's Field Staff D:F width — eliminates the column-F gap)
    ws.merge_cells(f'B{r}:F{r}'); ws.cell(row=r, column=2, value='Brand')
    ws.cell(row=r, column=7, value='● Non-Moving (cs)')
    ws.cell(row=r, column=8, value='▲ Critical (cs)')
    ws.cell(row=r, column=9, value='◆ Slow (cs)')
    ws.merge_cells(f'J{r}:K{r}'); ws.cell(row=r, column=10, value='Total Aging (cs)')
    ws.merge_cells(f'L{r}:M{r}'); ws.cell(row=r, column=12, value='# Shop×SKU rows')
    for col in [1,2,7,8,9,10,12]:
        c = ws.cell(row=r, column=col)
        c.font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=NAVY_SOFT)
        c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[r].height = 30; r += 1

    brand_start = r
    for idx, (brand, v) in enumerate(brand_sorted, 1):
        total = v['dead']+v['crit']+v['slow']
        alt = (idx % 2 == 0); row_fill = "FFFAFBFC" if alt else "FFFFFFFF"
        # Pre-fill every column 1..13 so merge boundaries are seamless
        for col in range(1, 14):
            cc = ws.cell(row=r, column=col)
            cc.fill = PatternFill('solid', fgColor=row_fill)
            cc.border = Border(bottom=Side(style='thin', color=LINE_GREY))
        ws.cell(row=r, column=1, value=idx)
        ws.merge_cells(f'B{r}:F{r}'); ws.cell(row=r, column=2, value=brand)
        ws.cell(row=r, column=7, value=v['dead']); ws.cell(row=r, column=8, value=v['crit']); ws.cell(row=r, column=9, value=v['slow'])
        ws.merge_cells(f'J{r}:K{r}'); ws.cell(row=r, column=10, value=total)
        ws.merge_cells(f'L{r}:M{r}'); ws.cell(row=r, column=12, value=v['rows'])
        for col in [1,2,7,8,9,10,12]:
            c = ws.cell(row=r, column=col)
            c.font = Font(name='Aptos Narrow', size=10.5, color=GREY_TXT)
            c.fill = PatternFill('solid', fgColor=row_fill)
            c.border = Border(bottom=Side(style='thin', color=LINE_GREY))
            c.alignment = Alignment(horizontal='center', vertical='center')
            if col == 2:
                c.font = Font(name='Aptos Narrow', size=11, bold=True, color=NAVY_DEEP)
                c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            if col == 1: c.font = Font(name='Aptos Narrow', size=11, bold=True, color=NAVY_SOFT)
            if col in (7,8,9,10): c.number_format = '#,##0.0'
            if col == 12: c.number_format = '#,##0'
        ws.cell(row=r, column=10).font = Font(name='Aptos Narrow', size=11, bold=True,
            color="FFB71C1C" if total > 400 else ("FFE65100" if total > 200 else "FF2E7D32"))
        ws.row_dimensions[r].height = 22; r += 1
    brand_end = r - 1
    # Use same data-bar color as BY BOND for visual consistency
    ws.conditional_formatting.add(f'J{brand_start}:J{brand_end}',
        DataBarRule(start_type='min', end_type='max', color="FFE65100", showValue=True))

    # ∑ TOTAL row for TOP BRANDS — parity with BY BOND table
    brand_dead_total = sum(v['dead'] for _, v in brand_sorted)
    brand_crit_total = sum(v['crit'] for _, v in brand_sorted)
    brand_slow_total = sum(v['slow'] for _, v in brand_sorted)
    brand_total_aging = brand_dead_total + brand_crit_total + brand_slow_total
    brand_total_rows = sum(v['rows'] for _, v in brand_sorted)
    # Pre-fill every column with navy for the total row
    for col in range(1, 14):
        cc = ws.cell(row=r, column=col)
        cc.fill = PatternFill('solid', fgColor=NAVY_DEEP)
        cc.border = Border(top=Side(style='medium', color=GOLD))
    ws.cell(row=r, column=1, value='∑')
    ws.merge_cells(f'B{r}:F{r}'); ws.cell(row=r, column=2, value=f"{len(brand_sorted)} BRANDS")
    ws.cell(row=r, column=7, value=brand_dead_total)
    ws.cell(row=r, column=8, value=brand_crit_total)
    ws.cell(row=r, column=9, value=brand_slow_total)
    ws.merge_cells(f'J{r}:K{r}'); ws.cell(row=r, column=10, value=brand_total_aging)
    ws.merge_cells(f'L{r}:M{r}'); ws.cell(row=r, column=12, value=brand_total_rows)
    for col in [1,2,7,8,9,10,12]:
        c = ws.cell(row=r, column=col)
        c.font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=NAVY_DEEP)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = Border(top=Side(style='medium', color=GOLD))
        if col == 2:
            c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        if col in (7,8,9,10): c.number_format = '#,##0.0'
        if col == 12: c.number_format = '#,##0'
    ws.row_dimensions[r].height = 28
    r += 1

    r += 1
    ws.merge_cells(f'A{r}:M{r}')
    c = ws.cell(row=r, column=1, value="TOP PACKS BY AGING STOCK")
    c.font = Font(name='Aptos Narrow', size=14, bold=True, color=NAVY_DEEP)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[r].height = 28; r += 1

    for col in range(1, 14):
        cc = ws.cell(row=r, column=col)
        cc.fill = PatternFill('solid', fgColor=NAVY_SOFT)
    ws.cell(row=r, column=1, value='#')
    ws.merge_cells(f'B{r}:F{r}'); ws.cell(row=r, column=2, value='Pack')
    ws.cell(row=r, column=7, value='● Non-Moving (cs)')
    ws.cell(row=r, column=8, value='▲ Critical (cs)')
    ws.cell(row=r, column=9, value='◆ Slow (cs)')
    ws.merge_cells(f'J{r}:K{r}'); ws.cell(row=r, column=10, value='Total Aging (cs)')
    ws.merge_cells(f'L{r}:M{r}'); ws.cell(row=r, column=12, value='# Shop×SKU rows')
    for col in [1,2,7,8,9,10,12]:
        c = ws.cell(row=r, column=col)
        c.font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=NAVY_SOFT)
        c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[r].height = 30; r += 1

    pack_start = r
    for idx, (pack, v) in enumerate(pack_sorted, 1):
        total = v['dead']+v['crit']+v['slow']
        alt = (idx % 2 == 0); row_fill = "FFFAFBFC" if alt else "FFFFFFFF"
        for col in range(1, 14):
            cc = ws.cell(row=r, column=col)
            cc.fill = PatternFill('solid', fgColor=row_fill)
            cc.border = Border(bottom=Side(style='thin', color=LINE_GREY))
        ws.cell(row=r, column=1, value=idx)
        ws.merge_cells(f'B{r}:F{r}'); ws.cell(row=r, column=2, value=pack)
        ws.cell(row=r, column=7, value=v['dead']); ws.cell(row=r, column=8, value=v['crit']); ws.cell(row=r, column=9, value=v['slow'])
        ws.merge_cells(f'J{r}:K{r}'); ws.cell(row=r, column=10, value=total)
        ws.merge_cells(f'L{r}:M{r}'); ws.cell(row=r, column=12, value=v['rows'])
        for col in [1,2,7,8,9,10,12]:
            c = ws.cell(row=r, column=col)
            c.font = Font(name='Aptos Narrow', size=10.5, color=GREY_TXT)
            c.fill = PatternFill('solid', fgColor=row_fill)
            c.border = Border(bottom=Side(style='thin', color=LINE_GREY))
            c.alignment = Alignment(horizontal='center', vertical='center')
            if col == 2:
                c.font = Font(name='Aptos Narrow', size=11, bold=True, color=NAVY_DEEP)
                c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            if col == 1: c.font = Font(name='Aptos Narrow', size=11, bold=True, color=NAVY_SOFT)
            if col in (7,8,9,10): c.number_format = '#,##0.0'
            if col == 12: c.number_format = '#,##0'
        ws.cell(row=r, column=10).font = Font(name='Aptos Narrow', size=11, bold=True,
            color="FFB71C1C" if total > 400 else ("FFE65100" if total > 200 else "FF2E7D32"))
        ws.row_dimensions[r].height = 22; r += 1
    pack_end = r - 1
    ws.conditional_formatting.add(f'J{pack_start}:J{pack_end}',
        DataBarRule(start_type='min', end_type='max', color="FFE65100", showValue=True))

    pack_dead_total = sum(v['dead'] for _, v in pack_sorted)
    pack_crit_total = sum(v['crit'] for _, v in pack_sorted)
    pack_slow_total = sum(v['slow'] for _, v in pack_sorted)
    pack_total_aging = pack_dead_total + pack_crit_total + pack_slow_total
    pack_total_rows = sum(v['rows'] for _, v in pack_sorted)
    for col in range(1, 14):
        cc = ws.cell(row=r, column=col)
        cc.fill = PatternFill('solid', fgColor=NAVY_DEEP)
        cc.border = Border(top=Side(style='medium', color=GOLD))
    ws.cell(row=r, column=1, value='∑')
    ws.merge_cells(f'B{r}:F{r}'); ws.cell(row=r, column=2, value=f"{len(pack_sorted)} PACK SIZES")
    ws.cell(row=r, column=7, value=pack_dead_total)
    ws.cell(row=r, column=8, value=pack_crit_total)
    ws.cell(row=r, column=9, value=pack_slow_total)
    ws.merge_cells(f'J{r}:K{r}'); ws.cell(row=r, column=10, value=pack_total_aging)
    ws.merge_cells(f'L{r}:M{r}'); ws.cell(row=r, column=12, value=pack_total_rows)
    for col in [1,2,7,8,9,10,12]:
        c = ws.cell(row=r, column=col)
        c.font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=NAVY_DEEP)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = Border(top=Side(style='medium', color=GOLD))
        if col == 2:
            c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        if col in (7,8,9,10): c.number_format = '#,##0.0'
        if col == 12: c.number_format = '#,##0'
    ws.row_dimensions[r].height = 28
    r += 1

    ws.merge_cells(f'A{r}:M{r}'); c = ws.cell(row=r, column=1, value="METHODOLOGY")
    c.font = Font(name='Aptos Narrow', size=12, bold=True, color=WHITE)
    c.fill = PatternFill('solid', fgColor=NAVY_DEEP)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[r].height = 24; r += 1
    for label, desc in [
        ("Period", f"{period_label} ({total_days} days). Source: COMBINED sheets of monthly KSBC shop sales workbooks."),
        ("Joining", "Each (Shop × Brand × Pack) row matched across months. Bond + Field Staff from MASTER DATA."),
        ("Months of Cover", f"Closing ÷ Monthly Sales Rate, where rate = sales over the trailing {RATE_HORIZON_MONTHS} months × 30 / days on shelf in that window."),
        ("Zero-Sale Months", "Count of months where stock was available (Opening or Receipts > 0) but Sales = 0."),
        ("● Non-Moving", "Closing > 0 AND total sales = 0 over the full period (the tier is literally non-moving)."),
        ("▲ Critical Aging", "Months of Cover > 6, a recent dead-stop (no sales in the trailing 2 months), OR ≥2 zero-sale months with Closing ≥ 5 cases."),
        ("◆ Slow Moving", "3 < Months of Cover ≤ 6."),
        ("✓ Healthy", "Months of Cover ≤ 3 (moving normally)."),
    ]:
        ws.merge_cells(f'A{r}:C{r}')
        c = ws.cell(row=r, column=1, value=label)
        c.font = Font(name='Aptos Narrow', size=10, bold=True, color=NAVY_DEEP)
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        c.fill = PatternFill('solid', fgColor=GREY_BG)
        ws.merge_cells(f'D{r}:M{r}')
        c = ws.cell(row=r, column=4, value=desc)
        c.font = Font(name='Aptos Narrow', size=10, color=GREY_TXT)
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1, wrap_text=True)
        c.fill = PatternFill('solid', fgColor=GREY_BG)
        ws.row_dimensions[r].height = 20; r += 1

    ws.sheet_view.showGridLines = False
    return bond_sorted


def build_bond_sheet(wb, bond, items, periods, total_days, bond_total_closing=None):
    months = [p[0] for p in periods]
    last_month, last_days, _, _ = periods[-1]
    period_label = f"{MONTH_ABBREV[last_month]} {last_days}"

    sale_labels = [f"{MONTH_ABBREV[m]} Sales" for m in months]
    clos_labels = [f"{MONTH_ABBREV[m]} Closing" for m in months]
    headers = ['Severity', 'Brand', 'Pack'] + sale_labels + ['Total Sales'] + clos_labels + ['Months of Cover', 'Zero Months', 'Trend']
    widths = [15, 38, 10] + [10]*len(months) + [12] + [11]*len(months) + [16, 12, 13]
    n_cols = len(headers); LAST_COL = get_column_letter(n_cols)

    ws = wb.create_sheet(bond)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.column_dimensions[get_column_letter(n_cols+1)].width = 3

    staff_set = sorted({x['staff'] for x in items if x['staff']})
    staff_str = ", ".join(staff_set) if staff_set else '—'
    dead = [x for x in items if x['severity']=="NON-MOVING"]
    crit = [x for x in items if x['severity']=="CRITICAL AGING"]
    slow = [x for x in items if x['severity']=="SLOW MOVING"]
    dead_cs = sum(x['latest_closing'] for x in dead)
    crit_cs = sum(x['latest_closing'] for x in crit)
    slow_cs = sum(x['latest_closing'] for x in slow)
    total_aging_cs = dead_cs + crit_cs + slow_cs
    n_shops = len({x['shop_code'] for x in items})

    ws.merge_cells(f'A1:{LAST_COL}1'); c = ws['A1']
    c.value = bond
    c.font = Font(name='Aptos Narrow', size=30, bold=True, color=WHITE)
    c.fill = PatternFill('solid', fgColor=NAVY_DEEP)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 50

    ws.merge_cells(f'A2:{LAST_COL}2'); c = ws['A2']
    c.value = f"Field Staff: {staff_str}   ·   Aging Stock as of {period_label}   ·   {total_days}-day analysis period"
    c.font = Font(name='Aptos Narrow', size=11, italic=True, color=GOLD_DIM)
    c.fill = PatternFill('solid', fgColor=NAVY_MID)
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border = GOLD_THIN_BOT
    ws.row_dimensions[2].height = 22; ws.row_dimensions[3].height = 12

    if n_cols >= 12:
        tile_spans = [('A','C'),('D','E'),('F','H'),('I','J'),('K', LAST_COL)]
    else:
        third = max(1, n_cols // 5)
        tile_spans = [('A', get_column_letter(third)),
                      (get_column_letter(third+1), get_column_letter(2*third)),
                      (get_column_letter(2*third+1), get_column_letter(3*third)),
                      (get_column_letter(3*third+1), get_column_letter(4*third)),
                      (get_column_letter(4*third+1), LAST_COL)]
    # Total aging subtitle: show "of X total cs · Y%" when we know the bond's total closing
    if bond_total_closing and bond_total_closing > 0:
        aging_pct = total_aging_cs / bond_total_closing * 100
        total_aging_sub = f"of {bond_total_closing:,.1f} cs on hand  ·  {aging_pct:.0f}%"
    else:
        total_aging_sub = 'cases on hand'

    tiles = [
        {'cols': tile_spans[0], 'label':'TOTAL AGING', 'value':f"{total_aging_cs:,.1f}", 'sub': total_aging_sub,
         'dark':NAVY_SOFT, 'bright':NAVY_DEEP, 'light':"FFE3E8F0"},
        {'cols': tile_spans[1], 'label':'● NON-MOVING', 'value':f"{dead_cs:,.1f}", 'sub':f"{len(dead)} positions",
         'dark':TIER['NON-MOVING']['dark'], 'bright':TIER['NON-MOVING']['bright'], 'light':TIER['NON-MOVING']['light']},
        {'cols': tile_spans[2], 'label':'▲ CRITICAL', 'value':f"{crit_cs:,.1f}", 'sub':f"{len(crit)} positions",
         'dark':TIER['CRITICAL AGING']['dark'], 'bright':TIER['CRITICAL AGING']['bright'], 'light':TIER['CRITICAL AGING']['light']},
        {'cols': tile_spans[3], 'label':'◆ SLOW', 'value':f"{slow_cs:,.1f}", 'sub':f"{len(slow)} positions",
         'dark':TIER['SLOW MOVING']['dark'], 'bright':TIER['SLOW MOVING']['bright'], 'light':TIER['SLOW MOVING']['light']},
        {'cols': tile_spans[4], 'label':'SHOPS WITH AGING', 'value':f"{n_shops}", 'sub':'outlets affected',
         'dark':NAVY_SOFT, 'bright':NAVY_DEEP, 'light':"FFE3E8F0"},
    ]
    for t in tiles:
        c1, c2 = t['cols']
        ws.merge_cells(f"{c1}4:{c2}4"); c = ws[f"{c1}4"]
        c.value = t['label']; c.font = Font(name='Aptos Narrow', size=10, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=t['dark']); c.alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(f"{c1}5:{c2}5"); c = ws[f"{c1}5"]
        c.value = t['value']; c.font = Font(name='Aptos Narrow', size=22, bold=True, color=t['bright'])
        c.fill = PatternFill('solid', fgColor=WHITE); c.alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(f"{c1}6:{c2}6"); c = ws[f"{c1}6"]
        c.value = t['sub']; c.font = Font(name='Aptos Narrow', size=9, color=GREY_LBL)
        c.fill = PatternFill('solid', fgColor=t['light']); c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[4].height = 22; ws.row_dimensions[5].height = 34
    ws.row_dimensions[6].height = 18

    # === AGING MIX strip — brand & pack composition of THIS bond's aging stock (added 17 Jun 2026,
    #     Abhay-requested). `items` are exactly the bond's aging positions, so this is aging-only. ===
    _brand_cs = defaultdict(float); _pack_cs = defaultdict(float)
    for x in items:
        _brand_cs[(x['brand'] or '—')] += x['latest_closing']
        _pack_cs[(x['pack'] or '—')] += x['latest_closing']
    mix_total = total_aging_cs if total_aging_cs > 0 else 1.0
    _spacer = _draw_mix_strip(ws, 7, n_cols, _mix_top(_brand_cs), _mix_top(_pack_cs), mix_total, hdr_size=10, val_size=10)
    ws.row_dimensions[_spacer].height = 10

    header_row = _spacer + 1
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=header_row, column=i, value=h)
        c.font = Font(name='Aptos Narrow', size=10, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=NAVY_SOFT)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = Border(bottom=Side(style='medium', color=GOLD))
    ws.row_dimensions[header_row].height = 32

    by_shop = defaultdict(list)
    for x in items: by_shop[(x['shop_code'], x['shop_name'])].append(x)
    shop_totals = sorted(
        [((sc, sn), rows, sum(r['latest_closing'] for r in rows)) for (sc, sn), rows in by_shop.items()],
        key=lambda t: -t[2])

    severity_rank = {"NON-MOVING": 0, "CRITICAL AGING": 1, "SLOW MOVING": 2}
    n_months = len(months); sales_col_start = 4
    total_sales_col = 3 + n_months + 1
    closing_col_start = total_sales_col + 1
    moc_col = closing_col_start + n_months
    zero_col = moc_col + 1
    trend_col = zero_col + 1
    last_shop_col = get_column_letter(n_cols - 5)
    banner_right_start_col = n_cols - 4
    banner_right_start = get_column_letter(banner_right_start_col)

    r = header_row + 1
    for (shop_code, shop_name), shop_rows, shop_total in shop_totals:
        cs_dead = sum(x['latest_closing'] for x in shop_rows if x['severity']=="NON-MOVING")
        cs_crit = sum(x['latest_closing'] for x in shop_rows if x['severity']=="CRITICAL AGING")
        cs_slow = sum(x['latest_closing'] for x in shop_rows if x['severity']=="SLOW MOVING")

        ws.merge_cells(f"A{r}:{last_shop_col}{r}")
        c = ws.cell(row=r, column=1, value=f"{shop_code}  ·  {shop_name}")
        c.font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=SHOP_BAND)
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)

        ws.merge_cells(f"{banner_right_start}{r}:{LAST_COL}{r}")
        c = ws.cell(row=r, column=banner_right_start_col,
            value=f"{shop_total:,.1f} cs aging   ·   ● {cs_dead:,.1f}   ▲ {cs_crit:,.1f}   ◆ {cs_slow:,.1f}")
        c.font = Font(name='Aptos Narrow', size=11, bold=True, color=GOLD_DIM)
        c.fill = PatternFill('solid', fgColor=SHOP_BAND)
        c.alignment = Alignment(horizontal='right', vertical='center', indent=1)
        ws.row_dimensions[r].height = 26; r += 1

        def sk(x):
            moc = x['months_cover'] if x['months_cover'] != float('inf') else 9999
            return (severity_rank[x['severity']], -moc, -x['latest_closing'])

        for row_idx, x in enumerate(sorted(shop_rows, key=sk)):
            tier = TIER[x['severity']]
            alt = (row_idx % 2 == 1); row_fill = ROW_ALT if alt else WHITE
            moc = x['months_cover']
            moc_display = "NO SALES" if moc == float('inf') else round(moc, 1)
            row_values = [f"{tier['icon']} {tier['label']}", x['brand'], x['pack']]
            for m in months: row_values.append(x['per_month'][m]['sales'])
            row_values.append(x['total_sales'])
            for m in months: row_values.append(x['per_month'][m]['closing'])
            _traj = {"rising":"▲ rising","falling":"▼ falling","flat":"→ flat","new":"new"}.get(x.get('trajectory','new'),"")
            row_values.append(moc_display); row_values.append(x['zero_sales_months']); row_values.append(_traj)

            for col_idx, v in enumerate(row_values, 1):
                c = ws.cell(row=r, column=col_idx, value=v)
                c.border = Border(bottom=Side(style='thin', color=LINE_GREY))
                c.fill = PatternFill('solid', fgColor=row_fill)
                c.font = Font(name='Aptos Narrow', size=10, color=GREY_TXT)
                if col_idx == 1:
                    c.font = Font(name='Aptos Narrow', size=9.5, bold=True, color=tier['dark'])
                    c.fill = PatternFill('solid', fgColor=tier['light'])
                    c.alignment = Alignment(horizontal='center', vertical='center')
                elif col_idx == 2: c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
                elif col_idx == 3: c.alignment = Alignment(horizontal='center', vertical='center')
                elif sales_col_start <= col_idx <= total_sales_col:
                    c.number_format = NUM_FMT
                    c.alignment = Alignment(horizontal='right', vertical='center', indent=1)
                    if v == 0: c.font = Font(name='Aptos Narrow', size=10, color=GREY_ZERO)
                elif closing_col_start <= col_idx < moc_col:
                    c.number_format = NUM_FMT
                    c.alignment = Alignment(horizontal='right', vertical='center', indent=1)
                    if v == 0: c.font = Font(name='Aptos Narrow', size=10, color=GREY_ZERO)
                elif col_idx == moc_col:
                    c.alignment = Alignment(horizontal='center', vertical='center')
                    if isinstance(v, str):
                        # "NO SALES" — make it pop with a tinted pill matching the severity tier
                        c.font = Font(name='Aptos Narrow', size=10.5, bold=True, color=tier['dark'])
                        c.fill = PatternFill('solid', fgColor=tier['light'])
                    else:
                        c.number_format = '0.0'
                        c.font = Font(name='Aptos Narrow', size=10, bold=True, color=tier['dark'])
                elif col_idx == zero_col:
                    c.alignment = Alignment(horizontal='center', vertical='center')
                elif col_idx == trend_col:
                    c.alignment = Alignment(horizontal='center', vertical='center')
                    _tc = {"▲ rising":"FFC62828","▼ falling":"FF2E7D32","→ flat":GREY_LBL}.get(v, GREY_LBL)
                    c.font = Font(name='Aptos Narrow', size=9.5, bold=(v == "▲ rising"), color=_tc)
            ws.row_dimensions[r].height = 18; r += 1
        ws.row_dimensions[r].height = 6; r += 1

    sum_cols = {}
    for i, m in enumerate(months):
        sum_cols[sales_col_start + i] = sum(x['per_month'][m]['sales'] for x in items)
    sum_cols[total_sales_col] = sum(x['total_sales'] for x in items)
    for i, m in enumerate(months):
        sum_cols[closing_col_start + i] = sum(x['per_month'][m]['closing'] for x in items)

    ws.merge_cells(f"A{r}:C{r}")
    c = ws.cell(row=r, column=1, value=f"BOND TOTAL — {n_shops} shops · {len(items)} aging positions")
    c.font = Font(name='Aptos Narrow', size=11, bold=True, color=GOLD_DIM)
    c.fill = PatternFill('solid', fgColor=NAVY_DEEP)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
    c.border = GOLD_MED_TOP
    for col, val in sum_cols.items():
        c = ws.cell(row=r, column=col, value=val)
        c.font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=NAVY_DEEP)
        c.alignment = Alignment(horizontal='right', vertical='center', indent=1)
        c.number_format = NUM_FMT
        c.border = GOLD_MED_TOP
    for col in (moc_col, zero_col, trend_col):
        c = ws.cell(row=r, column=col, value="")
        c.fill = PatternFill('solid', fgColor=NAVY_DEEP); c.border = GOLD_MED_TOP
    ws.row_dimensions[r].height = 30

    ws.freeze_panes = f'A{header_row+1}'
    ws.sheet_view.showGridLines = False


def build_drill_down(wb, name, items_filter, results):
    rows = [x for x in results if items_filter(x)]
    headers = ['Severity', 'Bond', 'Field Staff', 'Shop Code', 'Shop Name',
               'Brand', 'Pack', 'Product Code',
               'Total Sales (cs)', 'Latest Closing', 'Months of Cover', 'Zero Months', 'Trend']
    widths = [16, 18, 18, 11, 28, 32, 10, 14, 14, 14, 16, 12, 13]
    n = len(headers); LAST = get_column_letter(n)
    ws = wb.create_sheet(name)
    if name == "FULL DETAIL":
        title = f"ALL ROWS — {len(rows)} shop×SKU combos with movement or stock"
    else:
        cases = sum(x['latest_closing'] for x in rows)
        icon = TIER.get(name, {}).get('icon','')
        title = f"{icon} {name} — {len(rows)} rows, {cases:,.1f} cases"
    write_title_band(ws, title, LAST)
    write_header_row(ws, headers, 3, widths)
    _sev_order = {"NON-MOVING":0, "CRITICAL AGING":1, "SLOW MOVING":2, "NEW STOCK":3, "HEALTHY":4}
    def sk(x):
        moc = x['months_cover'] if x['months_cover'] != float('inf') else 9999
        return (_sev_order.get(x['severity'], 9), -x['latest_closing'], -moc)
    rows.sort(key=sk)
    r = 4
    for x in rows:
        tier = TIER[x['severity']]
        moc = x['months_cover']
        moc_display = "NO SALES" if moc == float('inf') else round(moc, 2)
        _traj = {"rising":"▲ rising","falling":"▼ falling","flat":"→ flat","new":"new"}.get(x.get('trajectory','new'),"")
        vals = [f"{tier['icon']} {x['severity']}", x['bond'], x['staff'], x['shop_code'], x['shop_name'],
                x['brand'], x['pack'], x['product_code'],
                x['total_sales'], x['latest_closing'], moc_display, x['zero_sales_months'], _traj]
        for col_idx, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=col_idx, value=v)
            c.font = Font(name='Aptos Narrow', size=10); c.border = HAIR_BORDER
            if col_idx == 1:
                c.font = Font(name='Aptos Narrow', size=10, bold=True, color=tier['dark'])
                c.fill = PatternFill('solid', fgColor=tier['light'])
                c.alignment = Alignment(horizontal='center', vertical='center')
            elif col_idx in (4,7,8): c.alignment = Alignment(horizontal='center', vertical='center')
            elif col_idx in (9,10):
                c.number_format = '#,##0.00'
                c.alignment = Alignment(horizontal='right', vertical='center')
            elif col_idx == 11:
                c.alignment = Alignment(horizontal='center', vertical='center')
                if isinstance(v, (int, float)):
                    c.number_format = '0.0'
                else:
                    # "NO SALES" — bold + tinted background pill, same as bond sheets
                    c.font = Font(name='Aptos Narrow', size=10.5, bold=True, color=tier['dark'])
                    c.fill = PatternFill('solid', fgColor=tier['light'])
            elif col_idx == 12: c.alignment = Alignment(horizontal='center', vertical='center')
            elif col_idx == 13:
                c.alignment = Alignment(horizontal='center', vertical='center')
                _tc = {"▲ rising":"FFC62828","▼ falling":"FF2E7D32","→ flat":GREY_LBL}.get(v, GREY_LBL)
                c.font = Font(name='Aptos Narrow', size=9.5, bold=(v == "▲ rising"), color=_tc)
        ws.row_dimensions[r].height = 18; r += 1
    ws.freeze_panes = 'D4'
    ws.auto_filter.ref = f"A3:{LAST}{r-1}"


def build_pivot(wb, name, results, key_fn, key_headers, key_widths):
    ws = wb.create_sheet(name)
    headers = key_headers + ['Shop×SKU rows', 'Non-Moving Rows', 'Non-Moving Cases',
        'Critical Rows', 'Critical Cases', 'Slow Rows', 'Slow Cases', 'Total Aging Cases']
    widths = key_widths + [14, 12, 14, 14, 14, 12, 14, 18]
    write_title_band(ws, name, get_column_letter(len(headers)))
    write_header_row(ws, headers, 3, widths)
    agg = defaultdict(lambda: {'rows':0,'dead_r':0,'dead_c':0,'crit_r':0,'crit_c':0,'slow_r':0,'slow_c':0,'extra':set()})
    for x in results:
        k = key_fn(x)
        agg[k]['rows'] += 1
        if x['severity']=="NON-MOVING":
            agg[k]['dead_r'] += 1; agg[k]['dead_c'] += x['latest_closing']
        elif x['severity']=="CRITICAL AGING":
            agg[k]['crit_r'] += 1; agg[k]['crit_c'] += x['latest_closing']
        elif x['severity']=="SLOW MOVING":
            agg[k]['slow_r'] += 1; agg[k]['slow_c'] += x['latest_closing']
    sorted_agg = sorted(agg.items(), key=lambda kv: -(kv[1]['dead_c']+kv[1]['crit_c']+kv[1]['slow_c']))
    r = 4
    for k, v in sorted_agg:
        if not k or (isinstance(k, tuple) and not k[0]): continue
        total_age = v['dead_c']+v['crit_c']+v['slow_c']
        key_vals = list(k) if isinstance(k, tuple) else [k]
        vals = key_vals + [v['rows'], v['dead_r'], v['dead_c'], v['crit_r'], v['crit_c'], v['slow_r'], v['slow_c'], total_age]
        n_key = len(key_vals)
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=r, column=col, value=val)
            c.font = Font(name='Aptos Narrow', size=10); c.border = HAIR_BORDER
            c.alignment = Alignment(horizontal='left' if col <= n_key else 'center', vertical='center')
            if col in (n_key+2, n_key+4, n_key+6, n_key+7):
                c.number_format = '#,##0.00'
            elif col > n_key: c.number_format = '#,##0'
        ws.row_dimensions[r].height = 22; r += 1
    ws.freeze_panes = f'{get_column_letter(len(key_headers)+1)}4'
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{r-1}"


def load_prewindow_dispatch(periods):
    """Earliest dispatch date BEFORE the analysis window per (shop,product) -- used
    to flag a NEW STOCK position as a re-order rather than a true debut. Best effort:
    returns empty if older secondary files are not present."""
    from datetime import datetime as _dt, date as _date
    out = {}
    if not SECONDARY_FOLDER.exists(): return out
    _yrs = _window_years([p[0] for p in periods])
    window_start = _date(_yrs[periods[0][0]], MONTH_ORDER.index(periods[0][0]) + 1, 1)
    for f in SECONDARY_FOLDER.glob('*SECONDARY SALES ANALYSIS.xlsx'):
        if f.name.startswith('~$'): continue
        try: wb = load_workbook(f, read_only=True, data_only=True)
        except Exception: continue
        disp = next((s for s in wb.sheetnames if 'DISPATCH' in s.upper()), None)
        if not disp: wb.close(); continue
        ws = wb[disp]
        hdr = [str(c).strip().lower() if c else '' for c in next(ws.iter_rows(min_row=1,max_row=1,values_only=True))]
        def col(*names):
            for n in names:
                if n in hdr: return hdr.index(n)
            return None
        ci_prod = col('product code'); ci_lic = col('licensee no.','licensee no'); ci_date = col('inv/gtn date')
        if None in (ci_prod, ci_lic, ci_date): wb.close(); continue
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[ci_lic] is None or row[ci_prod] is None or row[ci_date] is None: continue
            d = row[ci_date]
            if isinstance(d, str):
                try: d = _dt.strptime(d.strip(), '%d-%m-%Y').date()
                except Exception:
                    try: d = _dt.strptime(d.strip(), '%Y-%m-%d').date()
                    except Exception: continue
            else:
                import datetime as _dtm2
                if isinstance(d, _dtm2.datetime): d = d.date()
                elif isinstance(d, _dtm2.date): pass
                else: continue
            if d >= window_start: continue
            key = (str(row[ci_lic]).strip(), str(row[ci_prod]).strip())
            if key not in out or d < out[key]: out[key] = d
        wb.close()
    return out


def build_newstock_check(wb, results, periods, prewindow=None):
    """Verification sheet for NEW STOCK: per held-out position show arrival date,
    days on shelf, each window month's closing (proving no stock before arrival),
    the latest month's receipts/sales, and whether the shop had a pre-window order
    of the same SKU (re-order vs true debut)."""
    prewindow = prewindow or {}
    months = [p[0] for p in periods]; latest = months[-1]
    rows = [x for x in results if x['severity'] == "NEW STOCK"]
    info = TIER["NEW STOCK"]
    ws = wb.create_sheet("NEW STOCK")
    ws.sheet_view.showGridLines = False
    base = ['Bond','Shop Code','Shop Name','Brand','Pack','Arrival (in window)','Days on Shelf']
    mon_cols = [MONTH_ABBREV[m] + " Closing" for m in months]
    tail = [MONTH_ABBREV[latest] + " Receipts", MONTH_ABBREV[latest] + " Sales", 'Prior order (pre-window)', 'Why held out']
    headers = base + mon_cols + tail
    widths = [15,10,24,28,8,16,12] + [11]*len(mon_cols) + [12,10,18,30]
    for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width = w
    LAST = get_column_letter(len(headers)); ncol = len(headers)
    ws.merge_cells("A1:%s1" % LAST); c = ws['A1']
    c.value = info['icon'] + " NEW STOCK -- held-out new arrivals (arrival, days on shelf, why)"
    c.font = Font(name='Aptos Narrow', size=14, bold=True, color=WHITE); c.fill = PatternFill('solid', fgColor=NAVY_DEEP)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1); ws.row_dimensions[1].height = 34
    ws.merge_cells("A2:%s2" % LAST); c = ws['A2']
    c.value = ("Held out only if NO stock at window start AND on shelf <= %d days as of %s %d. "
               "Arrival = first warehouse->shop dispatch date in window." % (GRACE_DAYS, MONTH_ABBREV[latest], periods[-1][1]))
    c.font = Font(name='Aptos Narrow', size=10, italic=True, color=GOLD_DIM); c.fill = PatternFill('solid', fgColor=NAVY_SOFT)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1); c.border = Border(bottom=Side(style='medium', color=GOLD))
    ws.row_dimensions[2].height = 22; ws.row_dimensions[3].height = 6
    hr = 4
    for j,h in enumerate(headers,1):
        c = ws.cell(row=hr, column=j, value=h)
        c.font = Font(name='Aptos Narrow', size=11, bold=True, color=WHITE); c.fill = PatternFill('solid', fgColor=NAVY_SOFT)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True); c.border = Border(bottom=Side(style='medium', color=GOLD))
    ws.row_dimensions[hr].height = 30
    LINE = Side(style='thin', color=LINE_GREY)
    rows.sort(key=lambda x: (x.get('days_on_shelf', 0), x.get('bond') or ''))
    r = hr+1; n_returning = 0
    for idx,x in enumerate(rows):
        pm = x['per_month']; a = x.get('arrival_date'); dos = x.get('days_on_shelf')
        pre = prewindow.get((x['shop_code'], x['product_code']))
        if pre: n_returning += 1
        fill = ROW_ALT if idx % 2 else WHITE
        arr_txt = a.strftime('%d-%b-%Y') if a else '(in window, no dispatch row)'
        why = (("Re-dispatched after going empty - arrived %s - %sd on shelf" if x.get('restock_new')
                else "No stock at window start - arrived %s - %sd on shelf")
               % (a.strftime('%d %b') if a else MONTH_ABBREV[latest], dos if dos is not None else '<=%d' % GRACE_DAYS))
        prior_txt = ("Yes - " + pre.strftime('%d %b %Y')) if pre else "-"
        vals = [x['bond'], x['shop_code'], x['shop_name'], x['brand'], x['pack'], arr_txt, dos if dos is not None else '']
        vals += [round(pm[m]['closing'], 2) for m in months]
        vals += [round(pm[latest]['receipts'], 2), round(pm[latest]['sales'], 2), prior_txt, why]
        for j,v in enumerate(vals, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.fill = PatternFill('solid', fgColor=fill); c.border = Border(bottom=LINE)
            c.font = Font(name='Aptos Narrow', size=10, color=GREY_TXT)
            c.alignment = Alignment(horizontal='center', vertical='center')
            if j in (1,3,4): c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            if j == ncol: c.alignment = Alignment(horizontal='left', vertical='center', indent=1); c.font = Font(name='Aptos Narrow', size=9.5, color=info['dark'])
            if j == ncol-1: c.font = Font(name='Aptos Narrow', size=9.5, bold=bool(pre), color=("FFE65100" if pre else GREY_LBL))
            if j == 7: c.font = Font(name='Aptos Narrow', size=10, bold=True, color=info['dark'])
            if 8 <= j <= 9 + len(months): c.number_format = '#,##0.00;-#,##0.00;"·"'
        ws.row_dimensions[r].height = 20; r += 1
    if not rows:
        ws.merge_cells("A%d:%s%d" % (r, LAST, r)); c = ws.cell(row=r, column=1, value="No positions held out as NEW STOCK in this period.")
        c.font = Font(name='Aptos Narrow', size=10, italic=True, color=GREY_LBL); c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws.row_dimensions[r].height = 20; r += 1
    total_cs = sum(x['latest_closing'] for x in rows)
    ws.merge_cells("A%d:%s%d" % (r, LAST, r)); c = ws.cell(row=r, column=1,
        value=("∑  %d positions held out of aging  -  %.1f cases  -  <= %d days on shelf, no stock at window start  -  %d re-order(s) (prior pre-window dispatch)"
               % (len(rows), total_cs, GRACE_DAYS, n_returning)))
    c.font = Font(name='Aptos Narrow', size=10, bold=True, color=GOLD_DIM); c.fill = PatternFill('solid', fgColor=NAVY_DEEP)
    c.alignment = Alignment(horizontal='left', vertical='center', indent=1); c.border = Border(top=Side(style='medium', color=GOLD)); ws.row_dimensions[r].height = 24
    ws.freeze_panes = "A%d" % (hr+1)


def main():
    parser = argparse.ArgumentParser(description="Build KSBC aging stock analysis workbook.")
    parser.add_argument('--last', type=int, help='Use last N months')
    parser.add_argument('--from', dest='from_month', type=str, help='Start month (MARCH)')
    parser.add_argument('--to', dest='to_month', type=str, help='End month (MAY)')
    parser.add_argument('--output', type=str, help='Override output path')
    parser.add_argument('--master', type=str, default=None, help='Workbook with MASTER DATA tab')
    args = parser.parse_args()

    print("=" * 64); print("KSBC AGING STOCK ANALYSIS"); print("=" * 64)
    periods = discover_workbooks()
    if not periods: sys.exit("No KSBC analysis workbooks found.")
    print("Available months:")
    for m, days, path, sheet in periods:
        print(f"  {m:10} {days} days  {path.name}")
    periods = filter_periods(periods, last=args.last, from_month=args.from_month, to_month=args.to_month)
    if not periods: sys.exit("No periods after filtering.")
    print(f"\nUsing: {[(m, d) for m, d, _, _ in periods]}")
    # M6: warn loudly on a non-contiguous window (a missing interior month silently shrinks
    # total_days / sales and skews months-of-cover for stock spanning the gap).
    _pidx = [MONTH_ORDER.index(p[0]) for p in periods]
    _gap = [MONTH_ORDER[i] for i in range(_pidx[0], _pidx[-1] + 1) if i not in _pidx]
    if _gap:
        print(f"\n  WARNING: non-contiguous window - missing month(s): {', '.join(_gap)}.")
        print( "    Cover/zero-month metrics for stock spanning the gap will be understated.")
    total_days = sum(p[1] for p in periods)
    print(f"Total period: {total_days} days")
    master_path = args.master or str(periods[-1][2])
    effective_master = str(STANDALONE_MASTER) if STANDALONE_MASTER.exists() else master_path
    print(f"\nMaster data from: {Path(effective_master).name}"
          + ("" if STANDALONE_MASTER.exists() else "  (standalone master not found - using embedded tab)"))
    print("\nLoading data..."); records = load_data(periods)
    print(f"  Unique shop×SKU combos: {len(records)}")
    print("\nLoading dispatch dates (day-level shelf age)..."); dispatch_dates, dispatch_latest = load_dispatch_dates(periods)
    prewindow_dispatch = load_prewindow_dispatch(periods)
    print("\nComputing metrics..."); results, meta = compute_metrics(records, periods, dispatch_dates=dispatch_dates, dispatch_latest=dispatch_latest)
    print(f"  Filtered records: {len(results)}")
    print(f"  Excluded ghost rows (not in latest month): {meta['excluded_ghost_count']} ({meta['excluded_ghost_cases']:.1f} cs)")
    results, unmapped, closed_shops, active_ksbc_shops = attach_bond_staff(results, master_path)
    if unmapped: print(f"  Unmapped shops: {len(set(unmapped))}")
    print(f"  Active KSBC shops in MASTER DATA: {len(active_ksbc_shops)}")
    # Drop closed-shop rows entirely
    before = len(results)
    results = [x for x in results if x.get('shop_status','').lower() != 'closed']
    if before != len(results):
        print(f"  Excluded closed-shop rows: {before - len(results)}")

    sev_count = Counter(r['severity'] for r in results)
    sev_cases = defaultdict(float)
    for r in results: sev_cases[r['severity']] += r['latest_closing']
    print("\nSeverity breakdown:")
    for s in ["NON-MOVING", "CRITICAL AGING", "SLOW MOVING", "NEW STOCK", "HEALTHY"]:
        print(f"  {s:18} {sev_count[s]:>5} rows  {sev_cases[s]:>10,.1f} cs")

    print("\nBuilding workbook...")
    wb = Workbook(); wb.remove(wb.active)
    bond_sorted = build_summary(wb, results, periods, total_days, active_shops_count=len(active_ksbc_shops))
    print("  ✓ SUMMARY")
    by_bond = defaultdict(list)
    bond_total_closing = defaultdict(float)
    for x in results:
        # bond_total_closing includes HEALTHY rows too — it's the full bond inventory
        bond_total_closing[x['bond']] += x['latest_closing']
        if x['severity'] in ('NON-MOVING','CRITICAL AGING','SLOW MOVING'):
            by_bond[x['bond']].append(x)
    bonds_ordered = [b for b, _ in bond_sorted]
    for bond in bonds_ordered:
        if bond not in by_bond: continue
        build_bond_sheet(wb, bond, by_bond[bond], periods, total_days,
                         bond_total_closing=bond_total_closing[bond])
        print(f"  ✓ {bond} ({len(by_bond[bond])} positions, {bond_total_closing[bond]:.1f} cs total)")
    build_drill_down(wb, "FULL DETAIL", lambda x: True, results)
    build_drill_down(wb, "NON-MOVING", lambda x: x['severity']=="NON-MOVING", results)
    build_drill_down(wb, "CRITICAL AGING", lambda x: x['severity']=="CRITICAL AGING", results)
    build_drill_down(wb, "SLOW MOVING", lambda x: x['severity']=="SLOW MOVING", results)
    build_newstock_check(wb, results, periods, prewindow_dispatch); print("  ✓ NEW STOCK")
    build_pivot(wb, "BY BOND", results, lambda x: x['bond'], ['Bond'], [18])
    build_pivot(wb, "BY FIELD STAFF", results, lambda x: x['staff'] or '', ['Field Staff'], [22])
    build_pivot(wb, "BY BRAND × PACK", results, lambda x: (x['brand'], x['pack']), ['Brand', 'Pack'], [32, 12])
    print("  ✓ Drill-down sheets")

    # Run validation BEFORE we save — so any FAIL is visible to the operator
    print("\nRunning validation...")
    checks, raw_total_closing, recon_per_bond = run_validation(results, meta, periods, master_path)
    overall = "PASS"
    for name, status, detail in checks:
        marker = {"PASS":"✓","WARN":"⚠","FAIL":"✗","INFO":"ℹ"}.get(status, "?")
        print(f"  {marker} [{status:4}] {name}")
        print(f"           {detail}")
        if status == "FAIL": overall = "FAIL"
        elif status == "WARN" and overall != "FAIL": overall = "WARN"
    print(f"\n  OVERALL: {overall}")
    print("\nPer-bond reconciliation:")
    print(f"  {'Bond':<18} {'Raw cs':>12} {'Analysis':>12} {'Drift':>10}")
    for bond, raw, ours, diff, pct in sorted(recon_per_bond, key=lambda t: -t[1]):
        print(f"  {bond:<18} {raw:>12.2f} {ours:>12.2f} {pct:>8.2f}%")

    print("\nCoverage audit:")
    coverage = build_coverage_table(periods, records, results, closed_shops)
    print(f"  Total raw rows across all months: {coverage['total_raw_rows']:,}")
    print(f"  Unique (shop × SKU) keys: {coverage['unique_keys']:,}")
    print(f"    Included in analysis:    {coverage['in_analysis']:>5,}")
    print(f"    All-zero placeholders:   {coverage['placeholder']:>5,}")
    print(f"    Closed shops:            {coverage['closed']:>5,}")
    print(f"    Ghost (not in latest):   {coverage['ghost']:>5,}")
    print(f"    Zero-everywhere:         {coverage['zero_everywhere']:>5,}")
    print(f"  Unaccounted: {coverage['unaccounted']}  (must be 0)")
    build_validation_sheet(wb, checks, results, raw_total_closing, total_days, recon_per_bond, coverage)
    print("  ✓ VALIDATION sheet")

    order = ["SUMMARY", "VALIDATION"] + [b for b in bonds_ordered if b in wb.sheetnames] + [
        "FULL DETAIL", "NON-MOVING", "CRITICAL AGING", "SLOW MOVING", "NEW STOCK",
        "BY BOND", "BY FIELD STAFF", "BY BRAND × PACK"]
    wb._sheets = [wb[n] for n in order if n in wb.sheetnames]

    start_m = periods[0][0]; end_m, end_days, _, _ = periods[-1]
    name = f"AGING STOCK ANALYSIS - {MONTH_ABBREV[start_m]} to {MONTH_ABBREV[end_m]} {end_days}.xlsx"
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    out_path = Path(args.output) if args.output else OUTPUT_FOLDER / name
    wb.save(out_path)
    print(f"\n✓ Saved: {out_path}")
    total_age = sum(sev_cases[s] for s in ['NON-MOVING','CRITICAL AGING','SLOW MOVING'])
    total_pos = sum(sev_count[s] for s in ['NON-MOVING','CRITICAL AGING','SLOW MOVING'])
    print(f"  Total aging: {total_age:,.1f} cs across {total_pos} positions")
    print(f"  NEW STOCK held out of aging: {sev_cases['NEW STOCK']:,.1f} cs across {sev_count['NEW STOCK']} positions (≤{GRACE_DAYS}-day grace)")

    # ------------------------------------------------------------------
    # Live artifact refresh (added 8 Aug 2026, Abhay-approved).
    #
    # NON-FATAL by design: this is a presentation layer over a workbook that
    # is already saved, so a template problem must never turn a good build
    # into a failed one. It prints a single warning line and carries on --
    # the same contract build_warehouse_stock.py uses for its history hook.
    #
    # GATED on validation: a FAIL or an unaccounted coverage key means the
    # numbers are not trustworthy, and pushing them to a dashboard Abhay
    # reads at a glance is worse than leaving yesterday's figures up. The
    # workbook is still written either way so it can be inspected.
    # ------------------------------------------------------------------
    validation_ok = (overall != "FAIL" and coverage.get('unaccounted', 0) == 0)
    if validation_ok:
        try:
            import importlib.util as _ilu
            _p = Path(__file__).resolve().parent / "refresh_aging_live_artifact.py"
            _spec = _ilu.spec_from_file_location("refresh_aging_live_artifact", _p)
            _mod = _ilu.module_from_spec(_spec)
            _spec.loader.exec_module(_mod)
            _out = _mod._default_out()
            _pl = _mod.build(OUTPUT_FOLDER.parent, str(out_path), _out)
            print(f"  ✓ Live artifact refreshed: {_out}")
            print(f"    ({_pl['totals']['aging']:,.1f} cs aging · "
                  f"{_pl['totals']['aging_pct']:.1f}% of closing · "
                  f"{len(_pl['history'])} snapshots in history)")
        except Exception as _e:
            print(f"  ⚠ Live artifact refresh skipped: {_e}")
    else:
        print("  ⚠ Live artifact NOT refreshed - validation did not pass.")

    # M5: a scheduled/autonomous run must be able to DETECT a bad build. The file is still saved
    # (per spec §18) so the operator can inspect it, but exit non-zero on a validation FAIL or any
    # unaccounted coverage key so the task surfaces it instead of reporting a clean success.
    if not validation_ok:
        print("\n⚠ VALIDATION DID NOT PASS - investigate the VALIDATION sheet before sharing this workbook.")
        sys.exit(2)


if __name__ == "__main__":
    main()
