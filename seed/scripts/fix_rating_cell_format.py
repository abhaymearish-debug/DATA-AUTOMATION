"""
Recompute the tier from each row's data and re-apply the correct fill + font
on every Rating-bearing cell across the KSBC analysis workbook.

This is a MANDATORY final step of every KSBC build (per CLAUDE.md). Without
it, sort passes / data refreshes leave rating cells with stale fill+font that
no longer matches the row's recomputed tier — the regression Abhay flagged
on 27 Apr 2026 (KOLLAM R6 showing "🚀 High Performance" text in the green
"Balanced" colour, etc.).

The script is idempotent — re-running on a clean workbook is a no-op.

Coverage (all three rating-bearing sheet classes):

  1. Region sheets (KOLLAM, KOZHIKODE, ...): rating col J, header row 4,
     data D=Opening, E=Receipts, F=Sales, G=Closing.
  2. DETAIL sheets (KOLLAM DETAIL, ...): rating col G; nested Brand→Pack
     blocks. Data B=Opening, C=Receipts, D=Sales, E=Closing.
  3. BOND PERFORMANCE: rating col H. Bond rows have raw numbers in B-E;
     three cluster banner rows (Cluster 1/2/3) carry SUM-of-bonds; TOTAL
     row sums all 15 bonds.

Palette (canonical — aligned to CLAUDE.md 5-tier lock of 10 May 2026 on
1 Jun 2026, so the direct-paint fallback matches the CF layer exactly when
CF is suppressed; the older 27-Apr hexes are superseded):

  Data rows (region / DETAIL / non-banner BOND PERF / brand+pack rows):
    🚀 High Performance    fill FFBBDEFB / font FF1565C0
    ✅ Balanced            fill FFDCEDC8 / font FF2E7D32
    ⚠️ Inventory Heavy    fill FFFFE0B2 / font FFE65100
    🚫 Critical Overstock  fill FFFFCDD2 / font FFC62828
    — No activity          fill FFECEFF1 / font FF6B7280

  TOTAL rows (region/DETAIL/BOND PERF):
    Fill FF374151 (dark gray, kept), gold top border kept.
    Font = bright tier variant:
       high     FF4FC3F7   balanced FF81C784
       inv      FFFFB74D   crit     FFE57373   none FFB0BEC5

  Cluster banner rows (BOND PERFORMANCE only):
    Fill FF263F80 (navy, kept uniformly across all cells incl. Rating).
    Font = bright tier variant (same set as TOTAL rows).

Usage:
    python3 .claude/scripts/fix_rating_cell_format.py <workbook.xlsx>
"""

from __future__ import annotations
import re
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill


BONDS = ['KOLLAM', 'KOZHIKODE', 'ATTINGAL', 'PALAKKAD', 'KOTTARAKARA', 'KANNUR',
         'ALAPPUZHA', 'NEDUMANGAD', 'ALUVA', 'PERINTHALMANNA', 'THODUPUZHA',
         'PATHANAMTHITTA', 'KOTTAYAM', 'THRISSUR', 'TRIPUNITHURA']

# The five rating-label tokens (without emoji) used to identify a rating cell.
RATING_TOKENS = ('High Performance', 'Balanced', 'Inventory Heavy',
                 'Critical Overstock', 'No activity')

# ---- Canonical palettes -----------------------------------------------------

# (fill, font) — aligned to the CLAUDE.md / task-rules canonical 5-tier palette
# (locked 10 May 2026) on 1 Jun 2026 so this direct-paint fallback matches the
# CF layer when CF is suppressed. Was the older 27-Apr palette before.
DATA_FILL_FONT = {
    'high':     ('FFBBDEFB', 'FF1565C0'),
    'balanced': ('FFDCEDC8', 'FF2E7D32'),
    'inv':      ('FFFFE0B2', 'FFE65100'),
    'crit':     ('FFFFCDD2', 'FFC62828'),
    'none':     ('FFECEFF1', 'FF6B7280'),
}

# Bright variant used on dark backgrounds (TOTAL grey, cluster navy).
BRIGHT_FONT = {
    'high':     'FF4FC3F7',   # cyan
    'balanced': 'FF81C784',   # light green
    'inv':      'FFFFB74D',   # amber
    'crit':     'FFE57373',   # coral
    'none':     'FFB0BEC5',   # light grey
}

TOTAL_FILL = 'FF374151'    # dark gray
CLUSTER_FILL = 'FF263F80'  # navy

# ---- Tier resolution --------------------------------------------------------

def _tier(opening, receipts, sales, closing):
    """Compute tier from raw numeric values. Mirrors the rating formula."""
    o = opening or 0
    r = receipts or 0
    s = sales or 0
    c = closing or 0
    if o == 0 and r == 0 and s == 0 and c == 0:
        return 'none'
    denom = o + r
    st = (s / denom) if denom else 0.0
    if st >= 0.8:
        return 'high'
    if st >= 0.6:
        return 'balanced'
    if st >= 0.4:
        return 'inv'
    return 'crit'


def _coerce_num(v):
    """Return float for plain numbers; None for formulas/strings/None.

    Cluster + TOTAL rows store SUM formulas in B-E; we resolve them by
    summing the source rows directly (see _aggregate_range).
    """
    if isinstance(v, (int, float)):
        return float(v)
    return None


_RANGE_RE = re.compile(r'=SUM\(([A-Z])(\d+):([A-Z])(\d+)\)')
_LIST_RE  = re.compile(r'=SUM\(([^)]+)\)')


def _aggregate_range(ws, formula, col_letter):
    """Resolve a SUM(...) formula by reading the referenced cells.

    Handles every form the build pipeline produces, each comma-separated term
    being EITHER a range or a single cell:
      =SUM(B4:B9)                      → cluster banners (single range)
      =SUM(B4:B9,B11:B15,B17:B20)      → TOTAL row (multi-range)
      =SUM(B4,B5,B6,...)               → list of cells

    Fix 1 Jun 2026: the old list branch split on commas and then matched only
    the FIRST cell of each term, so a multi-range TOTAL like
    `=SUM(B4:B9,B11:B15,B17:B20)` read just B4+B11+B17 — a wrong partial sum
    that mis-tiered the BOND PERFORMANCE TOTAL row's rating colour (it showed
    "Inventory Heavy" text in the green "Balanced" colour). Now every term is
    parsed as a full range or a cell.
    """
    if not isinstance(formula, str) or not formula.startswith('='):
        return _coerce_num(formula)
    f = formula.replace(' ', '')
    if not (f.startswith('=SUM(') and f.endswith(')')):
        return None
    inner = f[5:-1]
    total = 0.0
    parsed_any = False
    for part in inner.split(','):
        rng = re.match(r'^([A-Z]+)(\d+):([A-Z]+)(\d+)$', part)
        one = re.match(r'^([A-Z]+)(\d+)$', part)
        if rng and rng.group(1) == rng.group(3) and len(rng.group(1)) == 1:
            col = ord(rng.group(1)) - ord('A') + 1
            for r in range(int(rng.group(2)), int(rng.group(4)) + 1):
                n = _coerce_num(ws.cell(r, col).value)
                if n is not None:
                    total += n
            parsed_any = True
        elif one and len(one.group(1)) == 1:
            col = ord(one.group(1)) - ord('A') + 1
            n = _coerce_num(ws.cell(int(one.group(2)), col).value)
            if n is not None:
                total += n
            parsed_any = True
    return total if parsed_any else None


# ---- Cell painting ----------------------------------------------------------

def _paint_data(cell, tier):
    fill, font = DATA_FILL_FONT[tier]
    cell.fill = PatternFill('solid', fgColor=fill)
    existing = cell.font
    cell.font = Font(
        name=existing.name or 'Trebuchet MS',
        size=existing.size or 10,
        bold=bool(existing.bold),
        italic=bool(existing.italic),
        color=font,
    )


def _paint_total(cell, tier):
    cell.fill = PatternFill('solid', fgColor=TOTAL_FILL)
    existing = cell.font
    cell.font = Font(
        name=existing.name or 'Trebuchet MS',
        size=existing.size or 10,
        bold=True,
        color=BRIGHT_FONT[tier],
    )


def _paint_cluster(cell, tier):
    cell.fill = PatternFill('solid', fgColor=CLUSTER_FILL)
    existing = cell.font
    cell.font = Font(
        name=existing.name or 'Trebuchet MS',
        size=existing.size or 11,
        bold=True,
        color=BRIGHT_FONT[tier],
    )


# ---- Sheet handlers ---------------------------------------------------------

def fix_region_sheet(ws):
    """Region sheet: rating col J. Data in D-G. TOTAL row at the bottom."""
    total_row = None
    for r in range(5, ws.max_row + 1):
        if ws.cell(r, 1).value == 'TOTAL':
            total_row = r
            break
    if not total_row:
        return 0
    fixed = 0
    for r in range(5, total_row):
        op = _coerce_num(ws.cell(r, 4).value)
        re_ = _coerce_num(ws.cell(r, 5).value)
        sa = _coerce_num(ws.cell(r, 6).value)
        cl = _coerce_num(ws.cell(r, 7).value)
        t = _tier(op, re_, sa, cl)
        _paint_data(ws.cell(r, 10), t)
        fixed += 1
    op = sum((_coerce_num(ws.cell(r, 4).value) or 0) for r in range(5, total_row))
    re_ = sum((_coerce_num(ws.cell(r, 5).value) or 0) for r in range(5, total_row))
    sa = sum((_coerce_num(ws.cell(r, 6).value) or 0) for r in range(5, total_row))
    cl = sum((_coerce_num(ws.cell(r, 7).value) or 0) for r in range(5, total_row))
    _paint_total(ws.cell(total_row, 10), _tier(op, re_, sa, cl))
    fixed += 1
    return fixed


def fix_detail_sheet(ws):
    """DETAIL sheet: nested Brand→Pack. Rating col G. Data in B-E.

    Every rating-bearing row uses the same logic — brand summaries, indented
    pack rows, and the per-shop TOTAL row at the bottom of each block.
    """
    fixed = 0
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 7).value
        # Repaint a cell IFF it carries ANY of the 5 rating labels. The old
        # guard matched only 'High Performance', so Balanced / Inventory Heavy /
        # Critical / No-activity rating cells on DETAIL sheets never got their
        # direct fill+font corrected after a sort (4 of 5 tiers drifted).
        # (Fix 1 Jun 2026.) Non-rating cells (banners, col-headers, spacers)
        # don't contain these tokens, so they're still skipped.
        if not isinstance(v, str) or not any(tok in v for tok in RATING_TOKENS):
            continue
        op = _coerce_num(ws.cell(r, 2).value)
        re_ = _coerce_num(ws.cell(r, 3).value)
        sa = _coerce_num(ws.cell(r, 4).value)
        cl = _coerce_num(ws.cell(r, 5).value)
        if op is None:
            op = _aggregate_range(ws, ws.cell(r, 2).value, 'B') or 0
        if re_ is None:
            re_ = _aggregate_range(ws, ws.cell(r, 3).value, 'C') or 0
        if sa is None:
            sa = _aggregate_range(ws, ws.cell(r, 4).value, 'D') or 0
        if cl is None:
            cl = _aggregate_range(ws, ws.cell(r, 5).value, 'E') or 0
        t = _tier(op, re_, sa, cl)
        if ws.cell(r, 1).value == 'TOTAL':
            _paint_total(ws.cell(r, 7), t)
        else:
            _paint_data(ws.cell(r, 7), t)
        fixed += 1
    return fixed


def fix_bond_performance(ws):
    """BOND PERFORMANCE: cluster banners + bond rows + TOTAL. Rating col H."""
    total_row = None
    for r in range(3, ws.max_row + 1):
        if ws.cell(r, 1).value == 'TOTAL':
            total_row = r
            break
    if not total_row:
        return 0
    fixed = 0
    for r in range(3, total_row):
        a = ws.cell(r, 1).value
        is_cluster = isinstance(a, str) and a.lower().startswith('cluster')
        if is_cluster:
            op = _aggregate_range(ws, ws.cell(r, 2).value, 'B') or 0
            re_ = _aggregate_range(ws, ws.cell(r, 3).value, 'C') or 0
            sa = _aggregate_range(ws, ws.cell(r, 4).value, 'D') or 0
            cl = _aggregate_range(ws, ws.cell(r, 5).value, 'E') or 0
            _paint_cluster(ws.cell(r, 8), _tier(op, re_, sa, cl))
        else:
            op = _coerce_num(ws.cell(r, 2).value)
            re_ = _coerce_num(ws.cell(r, 3).value)
            sa = _coerce_num(ws.cell(r, 4).value)
            cl = _coerce_num(ws.cell(r, 5).value)
            _paint_data(ws.cell(r, 8), _tier(op, re_, sa, cl))
        fixed += 1
    op = _aggregate_range(ws, ws.cell(total_row, 2).value, 'B') or 0
    re_ = _aggregate_range(ws, ws.cell(total_row, 3).value, 'C') or 0
    sa = _aggregate_range(ws, ws.cell(total_row, 4).value, 'D') or 0
    cl = _aggregate_range(ws, ws.cell(total_row, 5).value, 'E') or 0
    _paint_total(ws.cell(total_row, 8), _tier(op, re_, sa, cl))
    fixed += 1
    return fixed


# ---- Driver -----------------------------------------------------------------

def fix_dashboard(ws):
    """DASHBOARD league table: rating col K, data O/R/S/C in cols F-I.

    Direct-paint defence-in-depth so DASHBOARD ratings keep their tier colour
    even when CF is suppressed (print / filter). 16 Jun 2026 (BUG 14) —
    DASHBOARD was the only rating-bearing sheet with no direct-paint fallback.
    """
    fixed = 0
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 11).value
        if not (isinstance(v, str) and any(tok in v for tok in RATING_TOKENS)):
            continue
        op = _coerce_num(ws.cell(r, 6).value)
        re_ = _coerce_num(ws.cell(r, 7).value)
        sa = _coerce_num(ws.cell(r, 8).value)
        cl = _coerce_num(ws.cell(r, 9).value)
        if op is None and re_ is None and sa is None and cl is None:
            continue  # not a league-table data row (stray KPI text in col K)
        _paint_data(ws.cell(r, 11), _tier(op, re_, sa, cl))
        fixed += 1
    return fixed


def main():
    if len(sys.argv) < 2:
        print('Usage: python3 fix_rating_cell_format.py <workbook.xlsx>')
        sys.exit(1)
    target = sys.argv[1]
    wb = load_workbook(target, data_only=False)

    total_fixed = 0

    for bond in BONDS:
        if bond in wb.sheetnames:
            n = fix_region_sheet(wb[bond])
            total_fixed += n
            print(f'  ✓ {bond:18s}  {n} cells')
        sn = f'{bond} DETAIL'
        if sn in wb.sheetnames:
            n = fix_detail_sheet(wb[sn])
            total_fixed += n
            print(f'  ✓ {sn:18s}  {n} cells')

    if 'BOND PERFORMANCE' in wb.sheetnames:
        n = fix_bond_performance(wb['BOND PERFORMANCE'])
        total_fixed += n
        print(f'  ✓ BOND PERFORMANCE   {n} cells')

    if 'DASHBOARD' in wb.sheetnames:
        n = fix_dashboard(wb['DASHBOARD'])
        total_fixed += n
        print(f'  ✓ DASHBOARD          {n} cells')

    wb.save(target)
    print(f'\nFixed {total_fixed} rating cells. Saved {target}.')


if __name__ == '__main__':
    main()
