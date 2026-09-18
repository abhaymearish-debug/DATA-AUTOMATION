#!/usr/bin/env python3
"""
build_total_liquidation.py — produce TOTAL LIQUIDATION -<MONTH>.xlsx

Locked spec: .claude/memory/total-liquidation-monthly-statement.md (12 May 2026).

Inputs:
  - MASTER DATA CONFIRMED.xlsx (shop -> bond/staff/type/status lookup)
  - KSBC shop sales/<MONTH> SHOP SALES ANALYSIS.xlsx (KSBC tertiary sales by brand)
  - Secondary sales/<MONTH> SECONDARY SALES ANALYSIS.xlsx (CFD + BAR dispatches)

Output:
  - MONTHLY STATEMENT-LIQUIDATION/TOTAL LIQUIDATION -<MONTH>.xlsx

Usage:
  python3 build_total_liquidation.py --month APRIL
  python3 build_total_liquidation.py --month APRIL --exclude-bar
  python3 build_total_liquidation.py --month APRIL --output /tmp/scratch.xlsx
"""

import argparse
import os
import sys
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Locked layout constants
# ---------------------------------------------------------------------------
BRAND_COLS = [
    'BCB NO.1 CLASSIC BRANDY',
    "BLENDER'S CHOICE NO.1 BRANDY",
    "CHAIRMAN'S CHOICE XO BRANDY",
    'K.S 99 LIFE TIME MATURED XXX RUM',
    'MAGIC BLEND RESERVED XXX RUM',
    "MORNING WALKER'S XO BRANDY",
    'OLD PEARL NO.1 MATURED XXX RUM',
    'ROYAL OLD FORT NO.1 XXX RUM',
]
COL_HEADERS = ['Shop Code', 'Shop Name', 'Field Staff', 'Bond', 'Type'] + BRAND_COLS + ['TOTAL']
COL_WIDTHS = {
    'A': 12.5, 'B': 36.6640625, 'C': 17.0, 'D': 18.33203125, 'E': 11.6640625,
    'F': 18.5, 'G': 19.83203125, 'H': 18.5, 'I': 16.0, 'J': 17.1640625,
    'K': 17.0, 'L': 19.0, 'M': 18.1640625, 'N': 17.5,
}
HEADER_FILL = PatternFill('solid', fgColor='FF1F4E78')
TOTAL_FILL = PatternFill('solid', fgColor='FFD9E1F2')
THIN_GREY = Side(style='thin', color='FFBFBFBF')
THIN_BORDER = Border(top=THIN_GREY, bottom=THIN_GREY, left=THIN_GREY, right=THIN_GREY)
FONT_NAME = 'Aptos Narrow'

# Strip apostrophes AND dots entirely (curly, straight, or missing) and
# collapse whitespace, so "MORNING WALKER'S"/"MORNING WALKER’S"/"MORNING
# WALKERS" all map, and "BCB NO.1"/"B.C.B NO.1" both map to the BCB column.
# Dots are stripped on BOTH the BRAND_COLS keys and the source values, so the
# 8 brands stay mutually distinct (verified: no two collide after stripping).
APOSTROPHE_STRIP = str.maketrans({"'": "", "’": "", "`": "", ".": ""})

def norm_brand(s):
    """Match brand names robustly across apostrophe variants and whitespace."""
    if s is None:
        return ''
    return ' '.join(str(s).strip().upper().translate(APOSTROPHE_STRIP).split())

BRAND_LOOKUP = {norm_brand(b): b for b in BRAND_COLS}

# ---------------------------------------------------------------------------
# Master data
# ---------------------------------------------------------------------------
def load_master(path):
    """
    Returns {shop_code:int -> dict(name, staff, bond, type, active)}.
    Maps master CAT 'FED' to 'CFD' for the workbook's Type column.
    """
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb['16-4-25']
    out = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        code = r[3]
        if code is None:
            continue
        try:
            code = int(code)
        except (TypeError, ValueError):
            continue
        cat = (r[5] or '').strip().upper()
        type_label = {'FED': 'CFD', 'KSBC': 'KSBC', 'BAR': 'BAR'}.get(cat)
        if type_label is None:
            continue
        out[code] = {
            'name': (r[4] or '').strip() if isinstance(r[4], str) else r[4],
            'staff': (r[6] or '').strip() if isinstance(r[6], str) else r[6],
            'bond': (r[7] or '').strip() if isinstance(r[7], str) else r[7],
            'type': type_label,
            'active': (r[8] or '').strip().upper() == 'ACTIVE',
        }
    return out

# ---------------------------------------------------------------------------
# KSBC tertiary sales: read the AUTHORITATIVE '... COMBINED' roll-up sheet.
# NOT the DETAIL presentation sheets — they are a regenerated display layer
# that drifted ~319 cs low in May 2026 (see load_ksbc_brand_sales docstring).
# ---------------------------------------------------------------------------
def load_ksbc_brand_sales(ksbc_path):
    """
    KSBC tertiary brand x shop sales for the month.

    SOURCE OF TRUTH = the workbook's full-month roll-up sheet
    '<MONTH> 1-<lastday> COMBINED' (cols: Shop Code, Brand Name, Sales (Cases)).
    This reconciles exactly to BOND PERFORMANCE / the cumulative period totals.

    The DETAIL presentation sheets are deliberately NOT used as the source:
    in MAY 2026 they drifted ~319 cs BELOW COMBINED (almost the entire
    Morning Walker's column was missing from DETAIL), which silently
    undercounted the liquidation. Read the authoritative COMBINED sheet.

    Returns (out={shop_code:int -> {brand:str -> cases:float}}, seen_codes:set).
    """
    wb = load_workbook(ksbc_path, data_only=True, read_only=True)
    brand_norm = {norm_brand(b): b for b in BRAND_COLS}
    cands = [s for s in wb.sheetnames
             if s.upper().rstrip().endswith('COMBINED') and 'DISPATCH' not in s.upper()]
    if not cands:
        raise RuntimeError(f"No '... COMBINED' roll-up sheet in {ksbc_path}")
    # if several, take the one with the most rows (the full-month combined)
    sheet = max(cands, key=lambda sn: wb[sn].max_row)
    ws = wb[sheet]
    hdr = [str(c.value).strip() if c.value is not None else '' for c in ws[1]]

    def find_col(*names):
        for idx, h in enumerate(hdr):
            if h.lower() in names:
                return idx
        return None

    ci = find_col('shop code')
    bi = find_col('brand name', 'brand', 'item name')
    si = find_col('sales (cases)', 'sales')
    if None in (ci, bi, si):
        raise RuntimeError(
            f"COMBINED sheet '{sheet}' missing Shop Code / Brand Name / Sales (Cases); "
            f"headers seen: {hdr}")

    out = defaultdict(lambda: defaultdict(float))
    seen_codes = set()
    skipped = defaultdict(float)
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = row[ci]
        if code is None:
            continue
        try:
            code = int(code)
        except (TypeError, ValueError):
            continue
        seen_codes.add(code)
        key = norm_brand(row[bi])
        try:
            cases = float(row[si]) if row[si] is not None else 0.0
        except (TypeError, ValueError):
            cases = 0.0
        if key in brand_norm:
            out[code][brand_norm[key]] += cases
        elif cases:
            skipped[str(row[bi]).strip()] += cases
    if skipped:
        print(f"  [ksbc] skipped {len(skipped)} non-KSD brand name(s) with sales in "
              f"COMBINED: {sum(skipped.values()):.2f} cs")
    print(f"  [ksbc] source sheet: '{sheet}'")
    return out, seen_codes


# ---------------------------------------------------------------------------
# Secondary dispatches: parse COMBINED DISPATCHES for FED + BAR outlets
# ---------------------------------------------------------------------------
def load_secondary_brand_sales(sec_path, fed_codes, bar_codes):
    # Reads '<MONTH> COMBINED DISPATCHES' and aggregates Issue Cases by
    # (shop_code, brand) for FED + BAR outlets only. Columns are resolved BY
    # HEADER NAME (positional fallback) so a future column re-order in
    # build_secondary.py can't silently shift the read.
    # Returns (out={shop_code -> {brand -> cases}}, seen_codes), where
    # seen_codes is every licensee code seen (for main()'s unmatched audit).
    wb = load_workbook(sec_path, data_only=True, read_only=True)
    candidates = [s for s in wb.sheetnames if s.upper().rstrip().endswith('COMBINED DISPATCHES')]
    if not candidates:
        raise RuntimeError(f"No 'COMBINED DISPATCHES' sheet in {sec_path}")
    ws = max((wb[s] for s in candidates), key=lambda w: w.max_row)
    hdr = [str(c.value).strip().lower() if c.value is not None else '' for c in ws[1]]
    def col(default_idx, *names):
        for nm in names:
            for idx, h in enumerate(hdr):
                if h == nm:
                    return idx
        return default_idx
    li = col(6, 'licensee no.', 'licensee no', 'licensee number')
    ii = col(4, 'item name', 'brand name', 'brand')
    ci = col(12, 'issue cases', 'cases')
    bti = col(13, 'issue bottles', 'bottles')

    out = defaultdict(lambda: defaultdict(float))
    seen_codes = set()
    target = fed_codes | bar_codes
    brand_norm = {norm_brand(b): b for b in BRAND_COLS}
    skipped_brands = defaultdict(int)
    loose_bottles = []  # (code, item, bottles) on COUNTED CFD/BAR rows — H2 guard
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or li >= len(row) or row[li] is None:
            continue
        try:
            code = int(row[li])
        except (TypeError, ValueError):
            continue
        seen_codes.add(code)
        if code not in target:
            continue
        key = norm_brand(row[ii])
        if key not in brand_norm:
            if row[ii]:
                skipped_brands[row[ii]] += 1
            continue
        try:
            cases = float(row[ci]) if row[ci] is not None else 0.0
        except (TypeError, ValueError):
            cases = 0.0
        out[code][brand_norm[key]] += cases
        try:
            btl = float(row[bti]) if bti < len(row) and row[bti] is not None else 0.0
        except (TypeError, ValueError):
            btl = 0.0
        if btl:
            loose_bottles.append((code, row[ii], btl))
    if skipped_brands:
        print(f"  [secondary] skipped {sum(skipped_brands.values())} dispatch lines for non-KSD brands "
              f"({len(skipped_brands)} distinct names)")
    if loose_bottles:
        tot = sum(b for _, _, b in loose_bottles)
        print(f"  [secondary] WARNING: {len(loose_bottles)} CFD/BAR line(s) carry {tot:.0f} loose Issue "
              f"Bottles NOT folded into cases (no BPC in this sheet) — fractional cases may be undercounted:")
        for code, item, btl in loose_bottles[:10]:
            print(f"      shop {code}  {item}  +{btl:.0f} btl")
    return out, seen_codes

# ---------------------------------------------------------------------------
# Writer
# ---------------------------------------------------------------------------
def write_workbook(path, rows):
    """
    rows: list of dicts with keys: shop_code, shop_name, staff, bond, type, brand_cases (dict)
    """
    wb = Workbook()
    # Single data sheet (the empty 'Sheet1' placeholder from the April
    # reference is intentionally dropped — repurpose the default sheet).
    ws = wb.active
    ws.title = 'Brand Sales by Shop' 

    # Column widths
    for letter, w in COL_WIDTHS.items():
        ws.column_dimensions[letter].width = w

    # Header row
    header_font = Font(name=FONT_NAME, size=12, bold=True, color='FFFFFFFF')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for i, h in enumerate(COL_HEADERS, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = header_font
        c.fill = HEADER_FILL
        c.alignment = header_align
        c.border = THIN_BORDER
    ws.row_dimensions[1].height = 56.0

    # Data rows
    data_font = Font(name=FONT_NAME, size=12, bold=False)
    sum_font = Font(name=FONT_NAME, size=12, bold=True)
    data_align = Alignment(horizontal='center')
    for i, r in enumerate(rows, start=2):
        ws.cell(row=i, column=1, value=r['shop_code']).font = data_font
        ws.cell(row=i, column=2, value=r['shop_name']).font = data_font
        ws.cell(row=i, column=3, value=r['staff']).font = data_font
        ws.cell(row=i, column=4, value=r['bond']).font = data_font
        ws.cell(row=i, column=5, value=r['type']).font = data_font
        for j, brand in enumerate(BRAND_COLS, start=6):
            v = r['brand_cases'].get(brand, 0) or 0
            cell = ws.cell(row=i, column=j, value=round(float(v), 2))
            cell.font = data_font
            cell.number_format = '0.00'
        total_cell = ws.cell(row=i, column=14, value=f"=SUM(F{i}:M{i})")
        total_cell.font = sum_font
        total_cell.number_format = '0.00'
        for col in range(1, 15):
            c = ws.cell(row=i, column=col)
            c.alignment = data_align
            c.border = THIN_BORDER

    # A thin blank spacer row separates the data block from the TOTAL row.
    # Without it, Excel folds the (contiguous) TOTAL row into the auto-filter
    # region and hides it whenever a filter is applied. The gap keeps TOTAL
    # pinned and always visible.
    last_data_row = len(rows) + 1
    spacer_row = last_data_row + 1
    ws.row_dimensions[spacer_row].height = 6
    # Keep the table grid continuous through the spacer (no borderless gap):
    # give the thin spacer cells the same grey borders + Aptos Narrow font.
    for col in range(1, 15):
        sc = ws.cell(row=spacer_row, column=col)
        sc.font = data_font
        sc.border = THIN_BORDER

    # TOTAL row (uses SUBTOTAL so it reflects only the filtered/visible rows,
    # not a static grand total).
    tr = spacer_row + 1
    total_font = Font(name=FONT_NAME, size=12, bold=True)
    ws.cell(row=tr, column=1, value='TOTAL').alignment = Alignment(horizontal='right', vertical='center')
    for col in range(1, 15):
        c = ws.cell(row=tr, column=col)
        c.font = total_font
        c.fill = TOTAL_FILL
        c.border = THIN_BORDER
        if col >= 6:
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.number_format = '0.00'
        else:
            c.alignment = Alignment(horizontal=('right' if col == 1 else 'center'), vertical='center')
    ws.merge_cells(start_row=tr, end_row=tr, start_column=1, end_column=5)
    for col in range(6, 14):  # F..M
        letter = get_column_letter(col)
        ws.cell(row=tr, column=col, value=f"=SUBTOTAL(109,{letter}2:{letter}{last_data_row})")
    ws.cell(row=tr, column=14, value=f"=SUBTOTAL(109,N2:N{last_data_row})")

    # Freeze panes
    ws.freeze_panes = 'A2'

    # Auto-filter on the header row across the data rows only. The spacer row
    # below keeps the TOTAL row out of the filter region so it stays visible.
    ws.auto_filter.ref = f"A1:N{last_data_row}"

    wb.save(path)

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def find_analysis_file(folder, month, suffix_full, suffix_mid_pat):
    """Return path to <MONTH> SHOP SALES ANALYSIS.xlsx if it exists, else newest mid-month."""
    full = os.path.join(folder, f"{month} {suffix_full}")
    if os.path.exists(full):
        return full, 'full'
    # mid-month: look for "<MONTH> 1st - Nth ..."
    matches = []
    for f in os.listdir(folder):
        if f.startswith(f"{month} 1st - ") and f.endswith(suffix_mid_pat):
            matches.append(os.path.join(folder, f))
    if not matches:
        return None, None
    matches.sort(key=os.path.getmtime, reverse=True)
    return matches[0], 'mid'

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--month', required=True, help='Month in all-caps, e.g., APRIL')
    ap.add_argument('--exclude-bar', action='store_true',
                    help='Omit BAR rows (matches April reference layout)')
    ap.add_argument('--output', default=None,
                    help='Override output path (default: MONTHLY STATEMENT-LIQUIDATION/TOTAL LIQUIDATION -<MONTH>.xlsx)')
    ap.add_argument('--root', default=os.environ.get('KSD_ROOT', '.'),
                    help='Root folder (default: current dir or $KSD_ROOT)')
    args = ap.parse_args()

    month = args.month.strip().upper()
    root = args.root

    master_path = os.path.join(root, 'MASTER DATA CONFIRMED.xlsx')
    if not os.path.exists(master_path):
        print(f"ERROR: master not found: {master_path}", file=sys.stderr); sys.exit(2)

    ksbc_path, ksbc_kind = find_analysis_file(
        os.path.join(root, 'KSBC shop sales'), month,
        'SHOP SALES ANALYSIS.xlsx', 'ANALYSIS.xlsx')
    sec_path, sec_kind = find_analysis_file(
        os.path.join(root, 'Secondary sales'), month,
        'SECONDARY SALES ANALYSIS.xlsx', 'SECONDARY SALES ANALYSIS.xlsx')
    if not ksbc_path:
        print(f"ERROR: no KSBC analysis for {month}", file=sys.stderr); sys.exit(2)
    if not sec_path:
        print(f"ERROR: no Secondary analysis for {month}", file=sys.stderr); sys.exit(2)
    print(f"[inputs] master   = {master_path}")
    print(f"[inputs] KSBC     = {ksbc_path}  ({ksbc_kind})")
    print(f"[inputs] Secondary= {sec_path}  ({sec_kind})")
    if ksbc_kind == 'mid' or sec_kind == 'mid':
        print(f"  WARNING: building on mid-month source(s). Full-month preferred.")

    master = load_master(master_path)
    active_ksbc = {c for c, m in master.items() if m['type'] == 'KSBC' and m['active']}
    active_fed  = {c for c, m in master.items() if m['type'] == 'CFD'  and m['active']}
    active_bar  = {c for c, m in master.items() if m['type'] == 'BAR'  and m['active']}
    print(f"[master] active KSBC={len(active_ksbc)}, CFD={len(active_fed)}, BAR={len(active_bar)}")

    ksbc_sales, ksbc_seen = load_ksbc_brand_sales(ksbc_path)
    print(f"[ksbc]   shops with brand sales: {len(ksbc_sales)}")
    sec_sales, sec_seen = load_secondary_brand_sales(sec_path, active_fed, active_bar)
    print(f"[sec]    shops with dispatches: {len(sec_sales)} (CFD+BAR)")

    # --- Unmatched-outlet audit (CLAUDE.md: flag, never silently drop) -------
    master_codes = set(master)
    ksbc_active = {c for c, mm in master.items() if mm['type'] == 'KSBC' and mm['active']}
    ksbc_unknown = sorted(c for c in ksbc_seen if c not in master_codes)
    sec_unknown  = sorted(c for c in sec_seen  if c not in master_codes)
    ksbc_inactive_sales = sorted(
        (c, round(sum(ksbc_sales[c].values()), 2)) for c in ksbc_seen
        if c in master_codes and c not in ksbc_active and sum(ksbc_sales[c].values()) > 0)
    if ksbc_unknown:
        print(f"  [unmatched] {len(ksbc_unknown)} code(s) in KSBC COMBINED NOT in master -> DROPPED: {ksbc_unknown}")
    if ksbc_inactive_sales:
        print(f"  [unmatched] {len(ksbc_inactive_sales)} KSBC code(s) master marks closed/non-KSBC but carrying sales -> DROPPED: {ksbc_inactive_sales}")
    if sec_unknown:
        print(f"  [unmatched] {len(sec_unknown)} licensee code(s) in Secondary DISPATCHES NOT in master -> DROPPED: {sec_unknown}")
    if not (ksbc_unknown or ksbc_inactive_sales or sec_unknown):
        print(f"  [unmatched] none — every source outlet maps to an active master outlet")

    # Build rows: include every active outlet of the included types (KSBC + CFD; +BAR unless excluded)
    types_to_include = {'KSBC', 'CFD'}
    if not args.exclude_bar:
        types_to_include.add('BAR')

    rows = []
    for code, m in master.items():
        if not m['active'] or m['type'] not in types_to_include:
            continue
        if m['type'] == 'KSBC':
            brands = dict(ksbc_sales.get(code, {}))
        else:  # CFD or BAR
            brands = dict(sec_sales.get(code, {}))
        rows.append({
            'shop_code': code,
            'shop_name': m['name'],
            'staff':     m['staff'],
            'bond':      m['bond'],
            'type':      m['type'],
            'brand_cases': brands,
        })

    # Sort: bond ASC, type ASC (BAR<CFD<KSBC), shop_code ASC
    rows.sort(key=lambda r: (r['bond'] or '', r['type'], r['shop_code']))

    # Summary
    print()
    print(f"[summary] rows={len(rows)}")
    by_type = defaultdict(int)
    by_type_cases = defaultdict(float)
    brand_totals = defaultdict(float)
    grand = 0.0
    for r in rows:
        by_type[r['type']] += 1
        # Round each brand cell to 2 dp BEFORE summing so the printed summary
        # equals the workbook's displayed totals (cells are written rounded,
        # then summed by SUBTOTAL). Summing unrounded values drifted ~0.03 cs.
        row_total = 0.0
        for b, c in r['brand_cases'].items():
            cv = round(float(c or 0), 2)
            brand_totals[b] += cv
            row_total += cv
        by_type_cases[r['type']] += row_total
        grand += row_total
    for t in ('KSBC', 'CFD', 'BAR'):
        if by_type[t]:
            print(f"  {t:5s}: {by_type[t]:3d} shops, {by_type_cases[t]:>10.2f} cases")
    print(f"  GRAND TOTAL: {grand:.2f} cases")
    print()
    print("[brand totals]")
    for b in BRAND_COLS:
        print(f"  {b:42s} {brand_totals[b]:>10.2f}")

    # Output
    out_path = args.output or os.path.join(
        root, 'MONTHLY STATEMENT-LIQUIDATION', f'TOTAL LIQUIDATION -{month}.xlsx')
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    write_workbook(out_path, rows)
    print(f"\n[written] {out_path}")

if __name__ == '__main__':
    main()
