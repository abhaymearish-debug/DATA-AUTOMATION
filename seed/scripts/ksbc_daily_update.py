"""
KSBC shop-sales daily-update pipeline (ingest ONE or more raw-day files into
the current-month analysis workbook).

Usage:
    python3 ksbc_daily_update.py <current_analysis.xlsx> <raw1.xlsx> [<raw2.xlsx> ...]

Notes / conventions:
- Raw filename convention: "<month> <N>(st|nd|rd|th).xlsx" or
  "<month> <M>-<N>(st|nd|rd|th).xlsx"; the raw sheet must be named
  "SupplierWiseShopSaleReport".
- The pipeline preserves the existing region-sheet/DETAIL-sheet shop order
  (sort-by-rating is run as a SEPARATE final step via ksbc_sort_by_rating.py).
- Legacy KOZHIKODE sheet label 111014 → canonical 11014 via ksbc_alias_guard.
- Closed KSBC shops (PUNTHALATHAZHAM, THIRUVALLA, NORTH PARAVOOR, PATTALAKUNNU,
  KANJIRAM, PONNANI) are NEVER auto-added.
- Filename rule: mid-month = "<MONTH> 1st - <N>th ANALYSIS.xlsx",
                 full-month = "<MONTH> SHOP SALES ANALYSIS.xlsx".
- Writes to a unique scratch path (override with KSBC_SCRATCH_XLSX env var,
  otherwise tempfile.mkstemp). Summary JSON written to KSBC_SUMMARY_JSON env
  var if set, otherwise a unique /tmp file. Caller handles approval gate +
  live copy and reads the printed paths from stdout.
"""

from __future__ import annotations
import os, sys, re, shutil, importlib.util, tempfile
from copy import copy
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter


# Scratch paths — env-overridable, with per-process tempfile fallback so
# stale prior-session files in /tmp (owned by a different user) can't block
# the script (regression Abhay flagged 10 May 2026 — /tmp/scratch.xlsx was
# owned by nobody:nogroup, couldn't rm, script crashed twice). Set
# KSBC_SCRATCH_XLSX and KSBC_SUMMARY_JSON to override.
def _scratch_xlsx_path() -> str:
    p = os.environ.get('KSBC_SCRATCH_XLSX')
    if p:
        return p
    fd, path = tempfile.mkstemp(prefix='ksbc_scratch_', suffix='.xlsx')
    os.close(fd)
    os.unlink(path)  # we just want a fresh unique name; the copy below creates it
    return path

def _summary_json_path() -> str:
    p = os.environ.get('KSBC_SUMMARY_JSON')
    if p:
        return p
    fd, path = tempfile.mkstemp(prefix='ksbc_summary_', suffix='.json')
    os.close(fd)
    os.unlink(path)
    return path

# Import alias guard
HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
import ksbc_alias_guard as _alias
import ksbc_ingest_cumulative as _cumagg

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
MONTHS = {"JANUARY":1,"FEBRUARY":2,"MARCH":3,"APRIL":4,"MAY":5,"JUNE":6,
          "JULY":7,"AUGUST":8,"SEPTEMBER":9,"OCTOBER":10,"NOVEMBER":11,"DECEMBER":12}
MONTH_LAST_DAY = {"JANUARY":31,"FEBRUARY":28,"MARCH":31,"APRIL":30,"MAY":31,
                  "JUNE":30,"JULY":31,"AUGUST":31,"SEPTEMBER":30,"OCTOBER":31,
                  "NOVEMBER":30,"DECEMBER":31}  # 2026 non-leap
MONTH_SHORT = {"JANUARY":"Jan","FEBRUARY":"Feb","MARCH":"Mar","APRIL":"Apr",
               "MAY":"May","JUNE":"Jun","JULY":"Jul","AUGUST":"Aug",
               "SEPTEMBER":"Sep","OCTOBER":"Oct","NOVEMBER":"Nov","DECEMBER":"Dec"}

CLOSED_KSBC_CODES = {
    102004,  # 2004-PUNTHALATHAZHAM (KOLLAM)
    103012,  # 3012-THIRUVALLA (PATHANAMTHITTA)
    107042,  # 7042-NORTH PARAVOOR (ALUVA)
    108002,  # 8002-PATTALAKUNNU (THRISSUR)
    109030,  # 9030-KANJIRAM (PALAKKAD)
    110001,  # 10001-PONNANI (PERINTHALMANNA)
    # 101042 (1042-KALLAMBALAM, ATTINGAL) — REOPENED 21 Jun 2026; flipped back to
    # Active in master, removed from this hardcoded list so it auto-adds on next ingest.
}

# Style palette for auto-added rows (keeps the canonical grey/gold look)
NAVY = 'FF1A237E'
LIGHT_BG = 'FFF8FAFC'
ALT_ROW = 'FFF1F5F9'
WHITE = 'FFFFFFFF'
DARK_GRAY = 'FF374151'
TEXT_DARK = 'FF111827'
TOTAL_RATING = 'FF4FC3F7'
GOLD = 'FFFFD700'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def parse_day_range(filename: str) -> tuple[int, int, str]:
    """Return (start_day, end_day, MONTH_UPPER) parsed from e.g. 'april 21st.xlsx'
    or 'april 17-19th.xlsx'."""
    base = os.path.basename(filename).lower().replace('.xlsx', '')
    parts = base.split()
    month_u = None; day_part = None
    for m in MONTHS:
        if m.lower() in parts[0]:
            month_u = m; break
    if month_u is None and len(parts) >= 2:
        for m in MONTHS:
            if m.lower() == parts[0]:
                month_u = m; break
    if month_u is None:
        raise ValueError(f'Could not parse month from filename: {filename}')
    rest = base.replace(month_u.lower(), '').strip()
    m = re.match(r'(\d+)(?:st|nd|rd|th)?(?:\s*-\s*(\d+)(?:st|nd|rd|th)?)?$', rest)
    if not m:
        raise ValueError(f'Could not parse day range from filename: {filename}')
    start = int(m.group(1))
    end = int(m.group(2)) if m.group(2) else start
    return start, end, month_u


def ordinal(n: int) -> str:
    if 10 <= n % 100 <= 20: return 'th'
    return {1:'st', 2:'nd', 3:'rd'}.get(n % 10, 'th')


def copy_raw_sheet(src_wb_path: str, dst_wb, dst_sheet_name: str) -> int:
    """Copy the 'SupplierWiseShopSaleReport' sheet from src into dst_wb as
    a new sheet `dst_sheet_name`. Returns number of rows copied."""
    src_wb = load_workbook(src_wb_path, read_only=True, data_only=False)
    src_ws = src_wb['SupplierWiseShopSaleReport']
    if dst_sheet_name in dst_wb.sheetnames:
        del dst_wb[dst_sheet_name]
    new_ws = dst_wb.create_sheet(dst_sheet_name)
    n = 0
    for row in src_ws.iter_rows(values_only=True):
        new_ws.append(list(row))
        n += 1
    src_wb.close()
    return n


def read_raw_sheet_rows(ws) -> list[dict]:
    """Return list of dicts (Shop Code, Product code, Brand Name, Packing,
    Bottle Per Case, opening_cases_total, in_cases_total, out_cases_total,
    closing_cases_total) with cases-as-decimal math using bottles/BPC."""
    out = []
    it = ws.iter_rows(values_only=True)
    header = next(it, None)
    if not header:
        return out
    # Standard columns:
    # 0 Warehouse Name, 1 Shop Code, 2 Shop Name, 3 Product code, 4 Brand Name,
    # 5 Packing, 6 Bottle Per Case, 7 Shop Opening Cases, 8 Shop Opening Bottles,
    # 9 Shop In Cases, 10 Shop In Bottles, 11 Shop Out Cases, 12 Shop Out Bottles,
    # 13 Shop Closing Cases, 14 Shop Closing Bottles
    for r in it:
        if r is None or r[1] is None:
            continue
        try:
            code = int(r[1])
        except Exception:
            continue
        bpc = r[6] or 0
        # Surface (don't silently drop) loose bottles we can't convert because
        # Bottle-Per-Case is missing/zero. Fix 1 Jun 2026 — previously such
        # bottles vanished into 0 cases with no warning.
        if not bpc:
            loose = [b for b in (r[8], r[10], r[12], r[14]) if (b or 0)]
            if loose:
                sys.stderr.write(
                    f"⚠ read_raw: shop {r[1]} product {r[3]} ({r[5]}) has loose "
                    f"bottles {loose} but Bottle-Per-Case is missing/0 — those "
                    f"bottles cannot be converted to cases and are treated as 0. "
                    f"Check the raw export.\n")
        def c_b(cases, bottles):
            cases = cases or 0
            bottles = bottles or 0
            return cases + (bottles / bpc if bpc else 0)
        out.append({
            'Warehouse Name': r[0],
            'Shop Code': code,
            'Shop Name': r[2],
            'Product Code': r[3],
            'Brand Name': r[4],
            'Packing': r[5],
            'Bottle Per Case': bpc,
            'Opening': c_b(r[7], r[8]),
            'Receipts': c_b(r[9], r[10]),
            'Sales': c_b(r[11], r[12]),
            'Closing': c_b(r[13], r[14]),
        })
    return out


def _parse_master_worksheet(ws) -> dict[int, dict]:
    """Parse a worksheet shaped like the KSBC master (header row containing
    'Shop Code') into {code: {Shop Name, CAT, Field Staff, Bond, Status}}."""
    master = {}
    rows = list(ws.iter_rows(values_only=True))
    header_idx = None
    for i, r in enumerate(rows):
        if r and 'Shop Code' in (r or ()):
            header_idx = i; break
    if header_idx is None:
        raise RuntimeError('Cannot find MASTER DATA header row')
    header = list(rows[header_idx])
    idx = {name: header.index(name) for name in ('Shop Code','Shop Name','CAT','Field Staff','Bond','Status')}
    for r in rows[header_idx + 1:]:
        if not r or r[idx['Shop Code']] is None:
            continue
        try:
            code = int(r[idx['Shop Code']])
        except Exception:
            continue
        master[code] = {
            'Shop Name': r[idx['Shop Name']],
            'CAT': r[idx['CAT']],
            'Field Staff': r[idx['Field Staff']],
            'Bond': r[idx['Bond']],
            'Status': r[idx['Status']],
        }
    return master


# Path to the EXTERNAL standalone master file — overridable via env var.
# Defaults are derived from this script's own location so it works across
# different sandbox sessions without hardcoding `amazing-serene-johnson` etc.
# Note the trailing space in the filename — that's how Abhay keeps it.
#
# The pipeline reads from the external file FIRST and falls back to the
# workbook's embedded MASTER DATA sheet only if the external file is
# unreadable. Embedded has historically lagged (KALLAMBALAM 101042 was closed
# in master 12 May 2026 but the embedded MASTER DATA in the May workbook
# still showed Active — produced the off-by-one "284 Active" dashboard
# scope Abhay flagged 12 May 2026).
def _default_external_master_path() -> str:
    # Script lives at <ROOT>/.claude/scripts/ksbc_daily_update.py
    here = Path(__file__).resolve()
    root = here.parent.parent.parent  # .claude/scripts → .claude → <ROOT>
    # The shared KSBC master in this folder is `MASTER DATA CONFIRMED.xlsx`
    # (a symlink to the trailing-space `MASTER DATA CONFIRMED .xlsx`). The old
    # default looked for a bare `MASTER DATA .xlsx` that has never existed here,
    # so EVERY run silently fell back to the (lagging) embedded master sheet —
    # exactly the staleness this external-first design was meant to avoid.
    # Probe the known candidate names in order and return the first that exists
    # (fixed 3 Jun 2026). Sheet name still `16-4-25` (present in CONFIRMED).
    candidates = [
        'MASTER DATA CONFIRMED.xlsx',
    ]
    for name in candidates:
        if (root / name).exists():
            return str(root / name)
    return str(root / candidates[0])

EXTERNAL_MASTER_PATH = os.environ.get('KSBC_MASTER_DATA_PATH', _default_external_master_path())
EXTERNAL_MASTER_SHEET = os.environ.get('KSBC_MASTER_DATA_SHEET', '16-4-25')


def load_master_data(wb) -> dict[int, dict]:
    """Return {shop_code: {name, staff, bond, cat, status}} for every KSBC /
    FED / BAR shop in master.

    SOURCE-OF-TRUTH ORDER (locked 12 May 2026):
      1. External standalone master file `MASTER DATA .xlsx` (sheet `16-4-25`).
         This is what Abhay edits when a shop opens/closes. It is the
         authoritative reference for bond / field staff / status.
      2. FALLBACK: the workbook's embedded MASTER DATA sheet, IFF the
         external file is missing or unreadable. Logs a warning so the run
         summary makes clear the data may be stale.

    Reading from the external file first prevents the off-by-one Active-shop
    count Abhay flagged on May (embedded master still had KALLAMBALAM as
    Active after it was closed in the external master).
    """
    try:
        ext_wb = load_workbook(EXTERNAL_MASTER_PATH, data_only=True)
        sheet_name = EXTERNAL_MASTER_SHEET if EXTERNAL_MASTER_SHEET in ext_wb.sheetnames else ext_wb.sheetnames[-1]
        master = _parse_master_worksheet(ext_wb[sheet_name])
        print(f'  · Master loaded from external file: {EXTERNAL_MASTER_PATH} '
              f'(sheet "{sheet_name}", {len(master)} shops)')
        return master
    except Exception as e:
        print(f'  ⚠ External master at {EXTERNAL_MASTER_PATH} unreadable ({e!r}); '
              f'FALLING BACK to workbook embedded MASTER DATA. The embedded '
              f'sheet may lag the external — closed-shop changes might not '
              f'be reflected.')
        return _parse_master_worksheet(wb['MASTER DATA'])


# ---------------------------------------------------------------------------
# Combined sheet rebuild + roll-ups
# ---------------------------------------------------------------------------
def aggregate_month(wb, day_ranges_sheetnames: list[tuple[int,int,str]]) -> list[dict]:
    """Aggregate all raw day sheets into combined product-level rows.
    day_ranges_sheetnames: list of (start_day, end_day, sheet_name) sorted by start_day.
    Returns list of rows keyed by (shop, product, packing):
      { 'Shop Code','Shop Name','Product Code','Brand Name','Packing',
        'Bottle Per Case','Opening','Receipts','Sales','Closing',
        'Warehouse Name' }
    Opening comes from FIRST sheet, Receipts/Sales summed, Closing from LAST.
    """
    first_sheet = day_ranges_sheetnames[0][2]
    last_sheet = day_ranges_sheetnames[-1][2]

    # Initialise from first day. Shop Code is canonicalised via the alias map
    # (111014 → 11014) HERE (fix 1 Jun 2026) so the COMBINED sheet itself is
    # alias-clean and a shop appearing under its legacy code in one source +
    # canonical code in another collapses into ONE row instead of splitting.
    # Use the accumulate pattern (not a bare assignment) so duplicate rows for
    # the same (canonical shop, product, pack) within a sheet sum their
    # Receipts/Sales instead of the later row overwriting the earlier one.
    agg: dict[tuple, dict] = {}
    first_rows = read_raw_sheet_rows(wb[first_sheet])
    for r in first_rows:
        code = _alias.canonical_code(r['Shop Code'])
        key = (code, r['Product Code'], r['Packing'])
        if key not in agg:
            agg[key] = {
                'Warehouse Name': r['Warehouse Name'],
                'Shop Code': code,
                'Shop Name': r['Shop Name'],
                'Product Code': r['Product Code'],
                'Brand Name': r['Brand Name'],
                'Packing': r['Packing'],
                'Bottle Per Case': r['Bottle Per Case'],
                'Opening': r['Opening'],
                'Receipts': 0.0,
                'Sales': 0.0,
                'Closing': 0.0,
            }
        agg[key]['Receipts'] += r['Receipts']
        agg[key]['Sales']    += r['Sales']
        agg[key]['Closing']   = r['Closing']
    # Later sheets: sum Receipts/Sales; overwrite Closing each time so final
    # value = last sheet's closing. Shop Code canonicalised the same way.
    for s_i, (_start, _end, sn) in enumerate(day_ranges_sheetnames[1:], start=1):
        for r in read_raw_sheet_rows(wb[sn]):
            code = _alias.canonical_code(r['Shop Code'])
            key = (code, r['Product Code'], r['Packing'])
            if key not in agg:
                # First time this SKU appears is in a LATER source (it wasn't in
                # the first period). H2 fix 1 Jun 2026: carry its Opening from
                # THIS source — that is the SKU's opening balance at the start of
                # the window it entered in. The old code hardcoded Opening=0,
                # which understated Opening (and the Opening+Receipts sell-through
                # denominator → wrong ratings) for any second-period-only SKU and
                # broke Opening+Receipts−Sales≈Closing reconciliation.
                agg[key] = {
                    'Warehouse Name': r['Warehouse Name'],
                    'Shop Code': code,
                    'Shop Name': r['Shop Name'],
                    'Product Code': r['Product Code'],
                    'Brand Name': r['Brand Name'],
                    'Packing': r['Packing'],
                    'Bottle Per Case': r['Bottle Per Case'],
                    'Opening': r['Opening'],   # opening at the start of the window it first appears in
                    'Receipts': 0,
                    'Sales': 0,
                    'Closing': 0,
                }
            agg[key]['Receipts'] += r['Receipts']
            agg[key]['Sales']    += r['Sales']
            agg[key]['Closing']  = r['Closing']
    return list(agg.values())


def write_combined_sheet(wb, sheet_name: str, rows: list[dict]) -> None:
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    header = ['Warehouse Name','Shop Code','Shop Name','Product Code',
              'Brand Name','Packing','Bottle Per Case','Opening (Cases)',
              'Receipts (Cases)','Sales (Cases)','Closing (Cases)']
    ws.append(header)
    for r in rows:
        ws.append([r['Warehouse Name'], r['Shop Code'], r['Shop Name'],
                   r['Product Code'], r['Brand Name'], r['Packing'],
                   r['Bottle Per Case'], r['Opening'], r['Receipts'],
                   r['Sales'], r['Closing']])


def build_rollups(combined_rows: list[dict]) -> tuple[dict, dict, dict]:
    """Return shop_totals, shop_brand, shop_brand_pack.
    Keys normalised via alias_guard.canonical_code.

    shop_brand_pack is keyed by (code, brand, pack) — required by the nested
    DETAIL updater (`ksbc_detail_nested.update_detail_sheet`). The previous
    `shop_pack` key was (code, pack) which worked for the legacy two-table
    layout, but the nested layout needs to disambiguate the same pack size
    appearing under multiple brands. Regression Abhay flagged 11 May 2026 —
    pack rows in DETAIL sheets were all stuck at 0 even though the brand
    totals were correct, because the (code, brand, pack) lookup missed every
    time on a dict keyed by (code, pack).
    """
    shop_totals = {}
    shop_brand = {}
    shop_brand_pack = {}
    for r in combined_rows:
        code = _alias.canonical_code(int(r['Shop Code']))
        t = shop_totals.setdefault(code, {'Opening':0.0,'Receipts':0.0,'Sales':0.0,'Closing':0.0})
        t['Opening']  += r['Opening']
        t['Receipts'] += r['Receipts']
        t['Sales']    += r['Sales']
        t['Closing']  += r['Closing']
        brand = r['Brand Name']
        pack = r['Packing']
        b = shop_brand.setdefault((code, brand), {'Opening':0.0,'Receipts':0.0,'Sales':0.0,'Closing':0.0})
        for k in b: b[k] += r[k]
        bp = shop_brand_pack.setdefault((code, brand, pack),
                                        {'Opening':0.0,'Receipts':0.0,'Sales':0.0,'Closing':0.0})
        for k in bp: bp[k] += r[k]
    return shop_totals, shop_brand, shop_brand_pack


# ---------------------------------------------------------------------------
# Region sheet update
# ---------------------------------------------------------------------------
def update_region_sheet(ws, shop_totals: dict, month_u: str, latest_day: int) -> list[int]:
    """Overwrite Opening/Receipts/Sales/Closing for each shop row (5..TOTAL-1).
    Update title A1. Returns list of shop codes present on this sheet."""
    # Update title
    t = ws.cell(1,1).value
    if isinstance(t, str):
        ws.cell(1,1).value = re.sub(r'APRIL 1-\d+', f'{month_u} 1-{latest_day}', t) if month_u.upper() == 'APRIL' else \
            re.sub(rf'{month_u} 1-\d+', f'{month_u} 1-{latest_day}', t)
    # Generic replace: any "<M> 1-N" or "<M> 1-NN"
    if isinstance(t, str):
        ws.cell(1,1).value = re.sub(r'([A-Z]+) 1-\d+', f'{month_u} 1-{latest_day}', t)

    # Find TOTAL row
    total_row = None
    for r in range(5, ws.max_row + 1):
        if ws.cell(r,1).value == 'TOTAL':
            total_row = r; break
    if not total_row:
        raise RuntimeError(f'TOTAL row not found in {ws.title}')

    codes_on_sheet = []
    for r in range(5, total_row):
        code = ws.cell(r,1).value
        if code is None:
            continue
        try:
            code_i = int(code)
        except Exception:
            continue
        codes_on_sheet.append(code_i)
        canon = _alias.canonical_code(code_i)
        t = shop_totals.get(canon, {'Opening':0,'Receipts':0,'Sales':0,'Closing':0})
        ws.cell(r,4).value = t['Opening']
        ws.cell(r,5).value = t['Receipts']
        ws.cell(r,6).value = t['Sales']
        ws.cell(r,7).value = t['Closing']
    return codes_on_sheet


# ---------------------------------------------------------------------------
# DETAIL sheet update
# ---------------------------------------------------------------------------
def update_detail_sheet(ws, shop_brand: dict, shop_brand_pack: dict) -> int:
    """Delegate to ksbc_detail_nested.update_detail_sheet — the nested
    Brand→Pack layout has been canonical since 23 Apr 2026.

    `shop_brand_pack` must be keyed by (canonical_code, brand_name, pack_name).
    `build_rollups` produces this dict directly. If a caller passes the
    legacy 2-tuple `(code, pack)` dict, the nested updater will miss every
    pack lookup and leave packs at 0 (the regression Abhay flagged
    11 May 2026 — pack rows stuck at 0 even though brand totals were
    correct, because the old shim defaulted to an empty dict on 2-tuple
    keys). This signature now uniformly accepts the 3-tuple-keyed dict.
    """
    # Lazy import to avoid a module-load cycle at file top.
    import ksbc_detail_nested as _nested
    return _nested.update_detail_sheet(ws, shop_brand, shop_brand_pack, _alias.canonical_code)


# ---------------------------------------------------------------------------
# BOND PERFORMANCE update
# ---------------------------------------------------------------------------
def update_bond_performance(ws, region_totals: dict[str,dict], month_u: str, latest_day: int):
    """Rewrite bond data rows with region totals. Update A1 title.

    CRITICAL — skips cluster banner rows (col A starts with "Cluster") and the
    TOTAL row. Both keep live SUM formulas across cols B-E which must NOT be
    overwritten with literal numbers (regression Abhay flagged 10 May 2026:
    Cluster 2 banner at r10 and Cluster 3 banner at r16 got zeroed because
    they weren't in region_totals dict, blanking out their =SUM(B11:B15) and
    =SUM(B17:B20) formulas).
    """
    # Title update
    t = ws.cell(1,1).value
    if isinstance(t, str):
        ws.cell(1,1).value = re.sub(r'([A-Z]+) 1[–-]\d+', f'{month_u} 1–{latest_day}', t)
    for r in range(4, ws.max_row + 1):
        name = ws.cell(r,1).value
        if name in (None, 'TOTAL'):
            continue
        # Skip cluster banner rows — they own SUM formulas in cols B-E.
        if isinstance(name, str) and name.strip().lower().startswith('cluster'):
            continue
        tt = region_totals.get(str(name).strip(), {'Opening':0,'Receipts':0,'Sales':0,'Closing':0})
        ws.cell(r,2).value = tt['Opening']
        ws.cell(r,3).value = tt['Receipts']
        ws.cell(r,4).value = tt['Sales']
        ws.cell(r,5).value = tt['Closing']


# ---------------------------------------------------------------------------
# DASHBOARD update
# ---------------------------------------------------------------------------
def update_dashboard(ws, region_totals: dict, region_non_perf: dict[str,int],
                     top5_shops: list, bottom5_shops: list,
                     shop_counts_by_bond: dict[str,int],
                     month_u: str, latest_day: int, total_shops_active: int):
    """Rewrite all dashboard cells to reflect latest aggregates."""
    # Row 3 scope line
    scope = f'{month_u} 1-{latest_day}, 2026  |  Scope: {total_shops_active} KSBC Shops across 15 Bonds  |  All figures in cases'
    ws.cell(3,2).value = scope

    # KPI band row 8
    tot_opening = sum(t['Opening'] for t in region_totals.values())
    tot_receipts = sum(t['Receipts'] for t in region_totals.values())
    tot_sales = sum(t['Sales'] for t in region_totals.values())
    tot_closing = sum(t['Closing'] for t in region_totals.values())
    denom = tot_opening + tot_receipts
    st_pct = (tot_sales / denom) if denom else 0
    ws.cell(8,2).value = total_shops_active
    ws.cell(8,5).value = tot_sales
    ws.cell(8,8).value = st_pct
    ws.cell(8,11).value = tot_opening
    ws.cell(8,14).value = tot_receipts
    ws.cell(8,17).value = tot_closing

    # Row 9 labels
    tot_non_perf = sum(region_non_perf.values())
    ws.cell(9,2).value = f'{tot_non_perf} non-performing'
    ws.cell(9,11).value = f'cases on {MONTH_SHORT[month_u]} 1'
    ws.cell(9,17).value = f'cases on {MONTH_SHORT[month_u]} {latest_day}'

    # Callouts rows 13-14
    by_sales = sorted(region_totals.items(), key=lambda kv: kv[1]['Sales'], reverse=True)
    by_st = sorted(region_totals.items(), key=lambda kv: (kv[1]['Sales']/((kv[1]['Opening']+kv[1]['Receipts']) or 1)), reverse=True)
    top_bond_s = by_sales[0][0]; top_bond_s_sales = by_sales[0][1]['Sales']
    top_bond_st = by_st[0][0]
    top_bond_st_pct = by_st[0][1]['Sales']/((by_st[0][1]['Opening']+by_st[0][1]['Receipts']) or 1)
    low_bond_st = by_st[-1][0]
    low_bond_st_pct = by_st[-1][1]['Sales']/((by_st[-1][1]['Opening']+by_st[-1][1]['Receipts']) or 1)

    ws.cell(13,2).value = top_bond_s
    ws.cell(14,2).value = f'{round(top_bond_s_sales)} cases sold'
    ws.cell(13,8).value = top_bond_st
    ws.cell(14,8).value = f'{top_bond_st_pct*100:.1f}% sell-through'
    ws.cell(13,14).value = low_bond_st
    ws.cell(14,14).value = f'{low_bond_st_pct*100:.1f}% sell-through'

    # Bond rankings — SALES (rows 20..34)
    for i, (bond, tt) in enumerate(by_sales, start=20):
        ws.cell(i,2).value = i - 19
        ws.cell(i,3).value = bond
        ws.cell(i,4).value = tt['Sales']
        ws.cell(i,9).value = tt['Sales']
    # Bond rankings — ST% (rows 20..34)
    for i, (bond, tt) in enumerate(by_st, start=20):
        denom_b = tt['Opening'] + tt['Receipts']
        st = tt['Sales']/denom_b if denom_b else 0
        ws.cell(i,11).value = i - 19
        ws.cell(i,12).value = bond
        ws.cell(i,13).value = st
        ws.cell(i,18).value = st

    # TOP 5 shops rows 40..44
    for i, shop in enumerate(top5_shops, start=40):
        ws.cell(i,2).value = i - 39
        ws.cell(i,3).value = shop['name']
        ws.cell(i,4).value = shop['bond']
        ws.cell(i,5).value = shop['sales']
        ws.cell(i,6).value = shop['st']
    # BOTTOM 5 shops rows 40..44
    for i, shop in enumerate(bottom5_shops, start=40):
        ws.cell(i,11).value = i - 39
        ws.cell(i,12).value = shop['name']
        ws.cell(i,13).value = shop['bond']
        ws.cell(i,14).value = shop['sales']
        ws.cell(i,15).value = shop['st']

    # Bond league table rows 49..63 (preserve existing bond ORDER in col C)
    for r in range(49, 64):
        bond = ws.cell(r,3).value
        if not bond:
            continue
        tt = region_totals.get(bond, {'Opening':0,'Receipts':0,'Sales':0,'Closing':0})
        denom_b = tt['Opening'] + tt['Receipts']
        st = tt['Sales']/denom_b if denom_b else 0
        if st >= 0.8: rating = '🚀 High Performance'
        elif st >= 0.6: rating = '✅ Balanced'
        elif st >= 0.4: rating = '⚠️ Inventory Heavy'
        else: rating = '🚫 Critical Overstock'
        ws.cell(r,4).value = shop_counts_by_bond.get(bond, 0)
        ws.cell(r,5).value = region_non_perf.get(bond, 0)
        ws.cell(r,6).value = tt['Opening']
        ws.cell(r,7).value = tt['Receipts']
        ws.cell(r,8).value = tt['Sales']
        ws.cell(r,9).value = tt['Closing']
        ws.cell(r,10).value = st
        ws.cell(r,11).value = rating

    # Footer row 66
    ws.cell(66,2).value = (f'Dashboard generated from {month_u} 1-{latest_day} COMBINED data  •  '
                           f'Drill into bond sheets for shop-level detail  •  '
                           f'Drill into DETAIL sheets for brand/pack breakdown')


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------
def _rewrite_region_formulas(ws):
    """Rewrite the per-row %/rating formulas and the TOTAL SUM range after rows
    have been deleted from a region sheet (openpyxl does not adjust formula
    text on delete_rows). Idempotent."""
    total_row = None
    for r in range(5, ws.max_row + 1):
        if ws.cell(r, 1).value == 'TOTAL':
            total_row = r; break
    if not total_row:
        return
    last = total_row - 1
    # 16 Jun 2026 fix (BUG 2): 5-tier - all-zero shops read No activity.
    _rating = ('=IF(AND(D{r}=0,E{r}=0,F{r}=0,G{r}=0),"\u2014 No activity",'
               'IF(H{r}>=0.8,"\U0001F680 High Performance",'
               'IF(H{r}>=0.6,"\u2705 Balanced",'
               'IF(H{r}>=0.4,"\u26A0\uFE0F Inventory Heavy",'
               '"\U0001F6AB Critical Overstock"))))')
    for r in range(5, total_row):
        if ws.cell(r, 1).value in (None, ''):
            continue
        ws.cell(r, 8).value = f'=IFERROR(F{r}/(D{r}+E{r}),0)'
        ws.cell(r, 9).value = f'=IFERROR(G{r}/F{r},0)'
        ws.cell(r, 10).value = _rating.format(r=r)
    for col, letter in ((4, 'D'), (5, 'E'), (6, 'F'), (7, 'G')):
        ws.cell(total_row, col).value = f'=SUM({letter}5:{letter}{last})'
    ws.cell(total_row, 8).value = f'=IFERROR(F{total_row}/(D{total_row}+E{total_row}),0)'
    ws.cell(total_row, 9).value = f'=IFERROR(G{total_row}/F{total_row},0)'
    ws.cell(total_row, 10).value = _rating.format(r=total_row)


def prune_closed_shops(wb, region_names, closed_codes):
    """Remove closed-shop rows from region sheets and their blocks from DETAIL
    sheets. Closed shops are surveilled for activity but must not occupy an
    active row in the derived sheets. Idempotent — an already-pruned shop is
    simply not found. Returns list of (canonical_code, bond) pruned."""
    import ksbc_detail_nested as _nested
    pruned = []
    for rn in region_names:
        ws = wb[rn]
        total_row = None
        for r in range(5, ws.max_row + 1):
            if ws.cell(r, 1).value == 'TOTAL':
                total_row = r; break
        if not total_row:
            continue
        del_rows = []
        for r in range(5, total_row):
            v = ws.cell(r, 1).value
            try:
                canon = _alias.canonical_code(int(v))
            except (TypeError, ValueError):
                continue
            if canon in closed_codes:
                del_rows.append((r, canon))
        if del_rows:
            for r, canon in sorted(del_rows, key=lambda x: x[0], reverse=True):
                ws.delete_rows(r, 1)
                pruned.append((canon, rn))
            _rewrite_region_formulas(ws)
        # DETAIL block(s)
        dsn = f'{rn} DETAIL'
        if dsn not in wb.sheetnames:
            continue
        dws = wb[dsn]
        spans = []
        for (h, tr) in _nested._find_shop_blocks(dws):
            code = _nested._parse_shop_code(dws.cell(h, 1).value)
            if code is not None and _alias.canonical_code(code) in closed_codes:
                # Include one trailing spacer row if present (blank col A).
                end = tr
                if end + 1 <= dws.max_row and dws.cell(end + 1, 1).value in (None, ''):
                    end += 1
                spans.append((h, end))
        for (h, end) in sorted(spans, key=lambda x: x[0], reverse=True):
            for mr in list(dws.merged_cells.ranges):
                if not (mr.max_row < h or mr.min_row > end):
                    dws.unmerge_cells(str(mr))
            dws.delete_rows(h, end - h + 1)
    return pruned


def main():
    if len(sys.argv) < 3:
        print('usage: ksbc_daily_update.py <analysis.xlsx> <raw.xlsx> [<raw.xlsx> ...]')
        sys.exit(2)
    analysis_path = sys.argv[1]
    raw_paths = sys.argv[2:]

    # Separate cumulative-period raws (`<month> M-N CUMULATIVE.xlsx`) from
    # daily/range raws. Cumulatives are the source of truth for their window
    # (CLAUDE.md "Cumulative-period rule") and must be INGESTED as workbook
    # sheets so aggregation_sources() can use them in place of the dailies —
    # they do NOT go through parse_day_range/copy_raw_sheet (the trailing
    # " CUMULATIVE" makes parse_day_range raise). C1 fix 1 Jun 2026: previously
    # main() never ingested cumulatives at all, so full-month totals silently
    # fell back to drift-prone daily sums.
    cumulative_info = []   # (start, end, MONTH_UPPER, path) — cumulative raws
    raw_info = []          # (start, end, path, MONTH_UPPER) — daily/range raws
    for p in raw_paths:
        cum = _cumagg.parse_cumulative_filename(p)
        if cum:
            cs, ce, cmonth = cum
            cumulative_info.append((cs, ce, cmonth, p))
        else:
            s, e, mu = parse_day_range(p)
            raw_info.append((s, e, p, mu))
    raw_info.sort(key=lambda x: (x[3], x[0]))
    cumulative_info.sort(key=lambda x: (x[2], x[0]))

    months_seen = {x[3] for x in raw_info} | {x[2] for x in cumulative_info}
    if not months_seen:
        print('ERROR: no parseable raw files (daily or cumulative) supplied')
        sys.exit(3)
    if len(months_seen) > 1:
        print(f'ERROR: multiple months in raws: {months_seen} — process one month per run')
        sys.exit(3)
    month_u = next(iter(months_seen))

    # Copy analysis to scratch
    scratch = _scratch_xlsx_path()
    shutil.copyfile(analysis_path, scratch)
    wb = load_workbook(scratch, data_only=False)

    # Figure out which day sheets already exist + what's being added
    existing_day_sheets = []  # list of (start, end, sheet_name)
    for sn in list(wb.sheetnames):
        m = re.match(rf'^{month_u}\s+(\d+)(?:\-(\d+))?$', sn)
        if m:
            a = int(m.group(1)); b = int(m.group(2)) if m.group(2) else a
            existing_day_sheets.append((a, b, sn))
    existing_day_sheets.sort(key=lambda x: x[0])

    # Validate contiguity: existing covers [1..N_existing]; new raws must start at N+1
    covered_last = max((e for (_,e,_) in existing_day_sheets), default=0)
    print(f'Existing covered: up to day {covered_last}')
    new_days = [(s,e,p) for (s,e,p,_) in raw_info]
    new_days.sort(key=lambda x: x[0])
    if new_days and new_days[0][0] > covered_last + 1:
        # BUG 5 fix (16 Jun 2026): a forward GAP (missing middle day) must ABORT,
        # not warn — else COMBINED is silently understated by the missing day(s)
        # and the live workbook gets overwritten with wrong totals.
        print(f'ERROR: gap between existing coverage (up to day {covered_last}) and '
              f'new raws (start at day {new_days[0][0]}). Missing day(s) '
              f'{covered_last + 1}..{new_days[0][0] - 1}. Upload the missing day(s) '
              f'before rebuilding — refusing to produce an understated COMBINED.')
        sys.exit(4)
    elif new_days and new_days[0][0] != covered_last + 1:
        print(f'WARNING: new raws start at day {new_days[0][0]} but existing covers '
              f'up to day {covered_last} (overlap / re-ingest of an existing day)')
    # Contiguity within new_days
    for i in range(1, len(new_days)):
        if new_days[i][0] != new_days[i-1][1] + 1:
            print(f'ERROR: gap in new raw days between {new_days[i-1][1]} and {new_days[i][0]}')
            sys.exit(4)

    # Copy each new raw as a day sheet
    added = []
    for (s, e, p) in new_days:
        sn = f'{month_u} {s}' if s == e else f'{month_u} {s}-{e}'
        n_rows = copy_raw_sheet(p, wb, sn)
        added.append(sn)
        print(f'Added raw sheet {sn} ({n_rows} rows)')

    # Ingest cumulative-period raws as workbook sheets BEFORE assembling the
    # aggregation sources. Each becomes `<MONTH> M-N CUMULATIVE` and REPLACES
    # the dailies for its window in aggregation_sources(). Daily sheets remain
    # in the workbook as audit trail (never deleted here). (C1 fix 1 Jun 2026.)
    ingested_cumulatives = []
    for (cs, ce, cmonth, cpath) in cumulative_info:
        sn, nrows = _cumagg.ingest_cumulative_file(wb, cpath, cs, ce, cmonth)
        ingested_cumulatives.append(sn)
        print(f'Ingested cumulative sheet {sn} ({nrows} rows) from {os.path.basename(cpath)}')

    # Now assemble the canonical aggregation sources via the cumulative-aware
    # helper. If a <MONTH> 1-16 CUMULATIVE / 17-N CUMULATIVE sheet is present,
    # it REPLACES the dailies for that window (source of truth — daily-sum
    # drift from case-conversion rounding is exactly what the cumulative
    # file exists to fix). See ksbc-cumulative-period.md. Daily sheets are
    # preserved in the workbook as audit trail regardless.
    agg_sources, agg_status = _cumagg.aggregation_sources(wb, month_u)
    latest_day = max(e for (_,e,_) in agg_sources)
    print(f'Latest day = {latest_day}')
    print(f"Period 1 (1-16): {agg_status['period1_source']}")
    print(f"Period 2 (17-end): {agg_status['period2_source']}")

    # --- Load master (moved ahead of COMBINED build 3 Jun 2026 so closed shops
    # can be filtered out of the derived COMBINED while still being surveilled).
    master = load_master_data(wb)
    # Closed KSBC shops to exclude from ALL derived sheets. Union of the
    # hardcoded canonical closed list and anything master flags Closed, so a
    # freshly-closed shop is dropped the moment master records it — no lingering
    # zero-row (the KALLAMBALAM 101042 artifact, fixed 3 Jun 2026). They stay in
    # `shop_totals` so closed-shop activity surveillance still works.
    closed_codes = set(CLOSED_KSBC_CODES) | {
        c for c, m in master.items()
        if str(m.get('CAT')) == 'KSBC' and str(m.get('Status')) == 'Closed'
    }

    # --- Aggregate (cumulative-aware). Keep the FULL roll-up for surveillance;
    # build the COMBINED sheet from the active-only subset.
    combined_rows_full = aggregate_month(wb, agg_sources)
    shop_totals, shop_brand, shop_brand_pack = build_rollups(combined_rows_full)
    combined_rows = [r for r in combined_rows_full
                     if _alias.canonical_code(int(r['Shop Code'])) not in closed_codes]

    # --- Write combined (active shops only)
    for sn in list(wb.sheetnames):
        if 'COMBINED' in sn.upper():
            del wb[sn]
    combined_sheet_name = f'{month_u} 1-{latest_day} COMBINED'
    write_combined_sheet(wb, combined_sheet_name, combined_rows)
    print(f'Wrote {combined_sheet_name} with {len(combined_rows)} rows '
          f'({len(combined_rows_full) - len(combined_rows)} closed-shop product '
          f'rows excluded)')

    # --- Update region sheets
    region_names = [n for n in wb.sheetnames if f'{n} DETAIL' in wb.sheetnames and n not in ('MASTER DATA',)]

    # --- Region-coverage guard (H3 fix 1 Jun 2026). A region is recognised
    # ONLY if a `<bond> DETAIL` sheet exists. If a DETAIL sheet is missing or
    # renamed, that whole bond silently vanishes from region totals, BOND
    # PERFORMANCE, and the DASHBOARD KPIs — a silent under-report with no error.
    # Assert every bond that master expects (Active KSBC shops) has BOTH its
    # region sheet AND its DETAIL sheet present; abort loudly otherwise.
    expected_bonds = {m['Bond'] for m in master.values()
                      if m.get('CAT') == 'KSBC' and m.get('Status') == 'Active' and m.get('Bond')}
    missing_cov = []
    for b in sorted(expected_bonds):
        has_region = b in wb.sheetnames
        has_detail = f'{b} DETAIL' in wb.sheetnames
        if not (has_region and has_detail):
            missing_cov.append((b, has_region, has_detail))
    if missing_cov:
        lines = [f"    {b}: region sheet {'present' if hr else 'MISSING'}, "
                 f"DETAIL sheet {'present' if hd else 'MISSING'}"
                 for b, hr, hd in missing_cov]
        raise RuntimeError(
            "Region-coverage guard tripped — a bond that master expects (Active "
            "KSBC) is missing a region and/or DETAIL sheet, so it would be "
            "silently dropped from all totals:\n" + "\n".join(lines)
            + "\nRestore the missing sheet(s) before building."
        )

    codes_present_per_region = {}
    for rn in region_names:
        codes = update_region_sheet(wb[rn], shop_totals, month_u, latest_day)
        codes_present_per_region[rn] = set(codes)

    # --- Auto-add shops that appear in raw but not in any region sheet, are
    #     Active+KSBC in master. (Closed-shop reappearances are WARNINGS only.)
    codes_in_any_region = set()
    for s in codes_present_per_region.values():
        codes_in_any_region.update(_alias.canonical_code(c) for c in s)
    raw_codes = set(shop_totals.keys())
    auto_added = []
    closed_reappeared = []
    unmatched = []

    # --- Closed-shop activity surveillance (added 12 May 2026 per Abhay).
    # Flag ANY closed shop whose raw aggregate is non-zero — even if it already
    # has a row in a region sheet (which is the case for all 7 currently-closed
    # shops, since they were Active when the sheets were built). Without this,
    # the existing `closed_reappeared` only catches brand-new closed-shop
    # appearances; activity on an existing closed-shop row would silently
    # update the row and never surface.
    closed_shop_activity = []
    for code in sorted(CLOSED_KSBC_CODES):
        if code not in raw_codes:
            continue
        tt = shop_totals.get(code, {})
        if not any(abs(tt.get(k, 0) or 0) > 0.001
                   for k in ('Opening', 'Receipts', 'Sales', 'Closing')):
            continue
        closed_shop_activity.append({
            'code': code,
            'opening':  tt.get('Opening', 0),
            'receipts': tt.get('Receipts', 0),
            'sales':    tt.get('Sales', 0),
            'closing':  tt.get('Closing', 0),
        })
    if closed_shop_activity:
        print('')
        print('⚠  CLOSED-SHOP ACTIVITY DETECTED — review before approving:')
        for row in closed_shop_activity:
            print(f"    code={row['code']:>6}  O={row['opening']:>7.2f}  "
                  f"R={row['receipts']:>6.2f}  S={row['sales']:>7.2f}  "
                  f"C={row['closing']:>7.2f}")
        print('')

    for code in sorted(raw_codes - codes_in_any_region):
        if code in CLOSED_KSBC_CODES:
            closed_reappeared.append(code)
            continue
        m = master.get(code)
        if not m:
            unmatched.append(code)
            continue
        if m['CAT'] != 'KSBC' or m['Status'] != 'Active':
            unmatched.append(code)
            continue
        # Insert into region sheet
        bond = m['Bond']
        if bond not in region_names:
            unmatched.append(code)
            continue
        _insert_shop_into_region(wb[bond], code, m, shop_totals[code])
        _append_shop_block_to_detail(wb[f'{bond} DETAIL'], code, m,
                                     shop_brand, shop_brand_pack)
        auto_added.append((code, m['Shop Name'], bond))

    # --- Update DETAIL sheets (for existing shop blocks)
    for rn in region_names:
        update_detail_sheet(wb[f'{rn} DETAIL'], shop_brand, shop_brand_pack)
        # Also update title A1 if it references a month — DETAIL A1 typically
        # says "<BOND> — SHOP DETAILS (Brand & Pack breakdown)" no month, so
        # no update needed.

    # --- Prune any carried closed-shop rows/blocks from region + DETAIL sheets
    # (3 Jun 2026). update_region_sheet / update_detail_sheet only touch EXISTING
    # rows, so a closed shop bootstrapped from the prior month lingers as a
    # zero-row unless explicitly removed. This makes the derived sheets match the
    # active master (KALLAMBALAM 101042 was the lone closed shop still carrying a
    # row; the other 6 already had none).
    pruned_closed = prune_closed_shops(wb, region_names, closed_codes)
    if pruned_closed:
        print('Pruned %d closed-shop row(s)/block(s): %s'
              % (len(pruned_closed),
                 ', '.join(f'{c} ({b})' for c, b in pruned_closed)))

    # --- Build region totals for BOND PERFORMANCE + DASHBOARD
    region_totals = {}
    region_non_perf = {}
    shop_counts_by_bond = {}
    all_shops_for_ranking = []   # for TOP5/BOTTOM5 + total active count
    for rn in region_names:
        ws = wb[rn]
        total_row = None
        for r in range(5, ws.max_row + 1):
            if ws.cell(r,1).value == 'TOTAL':
                total_row = r; break
        op = rc = sl = cl = 0.0; cnt = 0; non_perf = 0
        for r in range(5, total_row):
            code = ws.cell(r,1).value
            if code in (None, ''):
                continue
            cnt += 1
            o = ws.cell(r,4).value or 0
            i = ws.cell(r,5).value or 0
            s_ = ws.cell(r,6).value or 0
            c_ = ws.cell(r,7).value or 0
            op += o; rc += i; sl += s_; cl += c_
            # Epsilon compare (fix 1 Jun 2026): case-conversion rounding can
            # leave a true-zero shop at e.g. 0.0000001, which exact `== 0`
            # would miss. "Non-performing" = effectively zero sales.
            if abs(s_) < 1e-6:
                non_perf += 1
            name = ws.cell(r,2).value
            denom_s = o + i
            st = s_/denom_s if denom_s else 0
            all_shops_for_ranking.append({
                'code': code, 'name': name, 'bond': rn, 'sales': s_, 'st': st,
            })
        region_totals[rn] = {'Opening':op,'Receipts':rc,'Sales':sl,'Closing':cl}
        region_non_perf[rn] = non_perf
        shop_counts_by_bond[rn] = cnt

    # --- Alias guard: zero-drop across ALL region sheets
    all_region_rows = []
    for rn in region_names:
        ws = wb[rn]
        total_row = None
        for r in range(5, ws.max_row + 1):
            if ws.cell(r,1).value == 'TOTAL':
                total_row = r; break
        for r in range(5, total_row):
            c = ws.cell(r,1).value
            if c in (None, ''): continue
            all_region_rows.append({'Shop Code': c, 'Sales': ws.cell(r,6).value or 0})
    raw_rows_for_guard = [{'Shop Code': _alias.canonical_code(int(r['Shop Code'])),
                           'Sales (Cases)': r['Sales']} for r in combined_rows]
    _alias.validate_no_zero_drops(all_region_rows, raw_rows_for_guard)

    # --- BOND PERFORMANCE
    update_bond_performance(wb['BOND PERFORMANCE'], region_totals, month_u, latest_day)

    # --- Alias guard: double-count check (BP vs region totals)
    # Skip cluster banner rows (col A starts with "Cluster") and TOTAL row -
    # both own =SUM formulas that aren't readable as floats here.
    bp = wb['BOND PERFORMANCE']
    bp_aggs = {}
    def _num(v):
        try:
            return float(v or 0)
        except (TypeError, ValueError):
            return 0.0
    for r in range(4, bp.max_row + 1):
        name = bp.cell(r,1).value
        if name in (None, 'TOTAL'): continue
        if isinstance(name, str) and name.strip().lower().startswith('cluster'):
            continue
        bp_aggs[str(name).strip()] = {
            'Opening':  _num(bp.cell(r,2).value),
            'Receipts': _num(bp.cell(r,3).value),
            'Sales':    _num(bp.cell(r,4).value),
            'Closing':  _num(bp.cell(r,5).value),
        }
    _alias.validate_no_double_counts(bp_aggs, region_totals)

    # --- Alias guard (C2 fix 1 Jun 2026): independent ground-truth check.
    # validate_no_double_counts above compares BP against the region totals,
    # but BP is WRITTEN FROM region_totals, so any double-count present in the
    # region totals themselves passes vacuously. Reconcile BOTH paths against a
    # third, alias-collapsed source built straight from the COMBINED roll-up.
    #
    # bond_truth: per-bond sum of the canonical-coded shop_totals, mapped to
    # bonds via master — exactly one contribution per canonical code, so it can
    # never double-count an aliased shop (e.g. MUKKAM 111014 + 11014).
    bond_truth = {}
    for code, t in shop_totals.items():
        m = master.get(code)
        if not m:
            continue                       # unmatched — surfaced separately
        if code in closed_codes:
            continue                       # closed — excluded from derived totals
        bond = m['Bond']
        if bond not in region_names:
            continue
        bt = bond_truth.setdefault(bond, {'Opening':0.0,'Receipts':0.0,'Sales':0.0,'Closing':0.0})
        for k in bt:
            bt[k] += t.get(k, 0) or 0
    # Per-region raw code lists (row order, duplicates PRESERVED) so a duplicate
    # canonical code on a single region sheet is caught directly.
    region_code_lists = {}
    for rn in region_names:
        ws = wb[rn]
        total_row = None
        for r in range(5, ws.max_row + 1):
            if ws.cell(r,1).value == 'TOTAL':
                total_row = r; break
        codes = []
        for r in range(5, total_row or ws.max_row + 1):
            c = ws.cell(r,1).value
            if c not in (None, ''):
                codes.append(c)
        region_code_lists[rn] = codes
    dupe_report = _alias.build_region_dupe_report(region_code_lists)
    _alias.validate_against_independent_truth(region_totals, bp_aggs, bond_truth,
                                              dupe_report)

    # --- DASHBOARD: top5 / bottom5
    all_shops_for_ranking.sort(key=lambda s: s['sales'], reverse=True)
    top5 = all_shops_for_ranking[:5]
    bottom5 = sorted(all_shops_for_ranking, key=lambda s: (s['sales'], s['st']))[:5]

    total_active = sum(1 for m in master.values() if m['CAT']=='KSBC' and m['Status']=='Active')
    update_dashboard(wb['DASHBOARD'], region_totals, region_non_perf,
                     top5, bottom5, shop_counts_by_bond, month_u, latest_day,
                     total_active)

    # --- Save scratch
    wb.save(scratch)

    # Write summary side file for caller
    summary = {
        'month': month_u,
        'latest_day': latest_day,
        'added_day_sheets': added,
        'ingested_cumulatives': ingested_cumulatives,
        'combined_sheet': combined_sheet_name,
        'total_opening': round(sum(t['Opening'] for t in region_totals.values()), 2),
        'total_receipts': round(sum(t['Receipts'] for t in region_totals.values()), 2),
        'total_sales': round(sum(t['Sales'] for t in region_totals.values()), 2),
        'total_closing': round(sum(t['Closing'] for t in region_totals.values()), 2),
        'total_non_perf': sum(region_non_perf.values()),
        'auto_added': auto_added,
        'closed_reappeared': closed_reappeared,
        'closed_shop_activity': closed_shop_activity,
        'unmatched': unmatched,
        'total_active_master': total_active,
        'total_active_on_sheet': sum(shop_counts_by_bond.values()),
        'period1_source': agg_status['period1_source'],
        'period2_source': agg_status['period2_source'],
    }
    import json
    summary_path = _summary_json_path()
    with open(summary_path, 'w') as f:
        json.dump(summary, f, indent=2, default=str)

    print('\n=== PIPELINE SUMMARY ===')
    print(json.dumps(summary, indent=2, default=str))
    print(f'\nScratch saved to {scratch}')


# ---------------------------------------------------------------------------
# Shop insertion helpers (auto-add)
# ---------------------------------------------------------------------------
def _insert_shop_into_region(ws, code: int, m: dict, totals: dict):
    """Insert new shop row above TOTAL, copying style of existing shop row
    above. Translates formulas to new row."""
    total_row = None
    for r in range(5, ws.max_row + 1):
        if ws.cell(r,1).value == 'TOTAL':
            total_row = r; break
    if total_row is None or total_row < 6:
        # No existing shop to template from — fall back to basic insert
        ws.insert_rows(total_row if total_row else 5)
        new_row = total_row if total_row else 5
    else:
        template_row = total_row - 1
        ws.insert_rows(total_row)
        new_row = total_row
        # Copy formatting from template row
        for c in range(1, 11):
            src = ws.cell(template_row, c)
            dst = ws.cell(new_row, c)
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.alignment = copy(src.alignment)
            dst.border = copy(src.border)
            dst.number_format = src.number_format
    ws.cell(new_row, 1).value = code
    ws.cell(new_row, 2).value = m['Shop Name']
    ws.cell(new_row, 3).value = m['Field Staff']
    ws.cell(new_row, 4).value = totals['Opening']
    ws.cell(new_row, 5).value = totals['Receipts']
    ws.cell(new_row, 6).value = totals['Sales']
    ws.cell(new_row, 7).value = totals['Closing']
    ws.cell(new_row, 8).value = f'=IFERROR(F{new_row}/(D{new_row}+E{new_row}),0)'
    ws.cell(new_row, 9).value = f'=IFERROR(G{new_row}/F{new_row},0)'
    ws.cell(new_row, 10).value = (
        f'=IF(AND(D{new_row}=0,E{new_row}=0,F{new_row}=0,G{new_row}=0),"— No activity",'
        f'IF(H{new_row}>=0.8,"🚀 High Performance",'
        f'IF(H{new_row}>=0.6,"✅ Balanced",'
        f'IF(H{new_row}>=0.4,"⚠️ Inventory Heavy","🚫 Critical Overstock"))))'
    )
    # Update TOTAL SUM ranges — TOTAL row was shifted +1; new_row is now part of each SUM.
    # After insert_rows, openpyxl shifts references in existing formulas automatically
    # for SUM(D5:D<TOTAL-1>). Let's just re-write the TOTAL SUM formulas explicitly
    # for safety.
    new_total = new_row + 1
    for col_letter, col_idx in (('D',4),('E',5),('F',6),('G',7)):
        ws.cell(new_total, col_idx).value = f'=SUM({col_letter}5:{col_letter}{new_row})'
    ws.cell(new_total, 8).value = f'=IFERROR(F{new_total}/(D{new_total}+E{new_total}),0)'
    ws.cell(new_total, 9).value = f'=IFERROR(G{new_total}/F{new_total},0)'
    ws.cell(new_total,10).value = (
        f'=IF(AND(D{new_total}=0,E{new_total}=0,F{new_total}=0,G{new_total}=0),"— No activity",'
        f'IF(H{new_total}>=0.8,"🚀 High Performance",'
        f'IF(H{new_total}>=0.6,"✅ Balanced",'
        f'IF(H{new_total}>=0.4,"⚠️ Inventory Heavy","🚫 Critical Overstock"))))'
    )
    # Also patch A3 formula (="Number of shops: "&COUNTA(A5:A<prev>))
    a3 = ws.cell(3,1).value
    if isinstance(a3, str) and 'COUNTA' in a3:
        ws.cell(3,1).value = re.sub(r'COUNTA\(A5:A\d+\)', f'COUNTA(A5:A{new_row})', a3)


def _append_shop_block_to_detail(ws, code: int, m: dict, shop_brand: dict, shop_brand_pack: dict):
    """Append a canonical NESTED Brand → Pack shop block to the DETAIL sheet.

    16 Jun 2026 fix (BUG 6): this previously emitted the DEPRECATED two-table
    (BRAND-WISE + PACK-WISE) layout, which the nested value-updater and the
    styling sweep cannot process — a newly auto-added shop shipped a malformed
    block. It now builds the canonical nested layout:

        [start]   "<code> — <name>  |  Staff: <staff>"   (merged A:G banner)
        [start+1] "Brand / Pack" | Opening | Receipts | Sales | Closing | Sell-Through % | Rating
        [start+2] brand summary row, then its indented pack rows, per brand
        [last]    TOTAL (sum of brand rows only — never brand+pack)

    Only structural rows + values + canonical SELF-REFERENCING F/G formulas are
    written here; the downstream fix_detail_row_text styling sweep paints the
    final banner / pack / brand / TOTAL styling and spacing.
    """
    canon = _alias.canonical_code(int(code))
    packs_by_brand = {}
    for (c, b, p), _v in shop_brand_pack.items():
        if c == canon:
            packs_by_brand.setdefault(b, set()).add(p)
    brands_for_shop = sorted({b for (c, b), _ in shop_brand.items() if c == canon}
                             | set(packs_by_brand.keys()))

    def _f(rr):
        return (f'=IF(AND(B{rr}=0,C{rr}=0,D{rr}=0,E{rr}=0),"—",'
                f'IFERROR(D{rr}/(B{rr}+C{rr}),0))')

    def _g(rr):
        return (f'=IF(AND(B{rr}=0,C{rr}=0,D{rr}=0,E{rr}=0),"— No activity",'
                f'IF(F{rr}>=0.8,"🚀 High Performance",'
                f'IF(F{rr}>=0.6,"✅ Balanced",'
                f'IF(F{rr}>=0.4,"⚠️ Inventory Heavy","🚫 Critical Overstock"))))')

    start = ws.max_row + 2  # leave a blank spacer row before the block
    r = start
    ws.cell(r, 1).value = f"{code} — {m['Shop Name']}  |  Staff: {m['Field Staff']}"
    ws.merge_cells(start_row=r, end_row=r, start_column=1, end_column=7)
    ws.cell(r, 1).font = Font(name='Trebuchet MS', size=12, bold=True, color=WHITE)
    ws.cell(r, 1).fill = PatternFill(fill_type='solid', fgColor=NAVY)
    ws.cell(r, 1).alignment = Alignment(horizontal='center', vertical='center')
    r += 1
    for c, v in enumerate(['Brand / Pack', 'Opening', 'Receipts', 'Sales',
                           'Closing', 'Sell-Through %', 'Rating'], start=1):
        cell = ws.cell(r, c); cell.value = v
        cell.font = Font(name='Trebuchet MS', size=10, bold=True, color=WHITE)
        cell.fill = PatternFill(fill_type='solid', fgColor=NAVY)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    r += 1
    brand_rows = []
    for brand in brands_for_shop:
        bt = shop_brand.get((canon, brand), {'Opening': 0, 'Receipts': 0, 'Sales': 0, 'Closing': 0})
        ws.cell(r, 1).value = brand
        ws.cell(r, 2).value = bt['Opening']; ws.cell(r, 3).value = bt['Receipts']
        ws.cell(r, 4).value = bt['Sales'];   ws.cell(r, 5).value = bt['Closing']
        ws.cell(r, 6).value = _f(r); ws.cell(r, 7).value = _g(r)
        brand_rows.append(r); r += 1
        for pack in sorted(packs_by_brand.get(brand, set())):
            pt = shop_brand_pack.get((canon, brand, pack),
                                     {'Opening': 0, 'Receipts': 0, 'Sales': 0, 'Closing': 0})
            ws.cell(r, 1).value = pack
            ws.cell(r, 2).value = pt['Opening']; ws.cell(r, 3).value = pt['Receipts']
            ws.cell(r, 4).value = pt['Sales'];   ws.cell(r, 5).value = pt['Closing']
            ws.cell(r, 6).value = _f(r); ws.cell(r, 7).value = _g(r)
            r += 1
    ws.cell(r, 1).value = 'TOTAL'
    for col, letter in ((2, 'B'), (3, 'C'), (4, 'D'), (5, 'E')):
        ws.cell(r, col).value = ('=' + '+'.join(f'{letter}{br}' for br in brand_rows)) if brand_rows else 0
    ws.cell(r, 6).value = _f(r); ws.cell(r, 7).value = _g(r)
    for c in range(1, 8):
        ws.cell(r, c).fill = PatternFill(fill_type='solid', fgColor=DARK_GRAY)
        ws.cell(r, c).font = Font(name='Trebuchet MS', size=10, bold=True, color=WHITE)
        ws.cell(r, c).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(r, c).border = Border(top=Side(style='medium', color=GOLD))


if __name__ == '__main__':
    main()
