"""
Style the column-header row + TOTAL row on every KSBC region/bond sheet
(KOLLAM, KOZHIKODE, ATTINGAL, ...) to match the elegant pass already done
on DETAIL sheets — taller rows, bigger bold text, gold accent borders.

Region sheet layout (10 columns A-J, locked):
  r1: title band
  r2: subtitle ("All figures in cases.")
  r3: shop count
  r4: COLUMN-HEADER row (Shop Code | Shop Name | Field Staff | Opening |
      Receipts | Sales | Closing | Sell-Through % | Closing vs Sales % |
      Rating)
  r5..(TOTAL-1): data rows (one per shop)
  TOTAL row at the bottom

What this script does, per region sheet:
  0. Hero band (r1-r3): full-width A:J deep-navy FF0D1B4A hero — white bold
     18pt title, gold-dim italic subtitle, soft-steel shop count, medium-gold
     rule under r3 (locked 16 Jul 2026, Abhay-approved Option A). Sub-lines sit
     AT THE SIDES (Abhay 16 Jul PM): r2 subtitle left-aligned, r3 count right-aligned.
  1. Column-header (r4):
     - Height bumped to 30pt
     - Deep navy fill FF1A237E (kept)
     - 12pt bold white text, centred
     - Thin gold (FFFFD700) bottom-border accent for elegance
  2. TOTAL row:
     - Height bumped to 32pt
     - 14pt bold text (was 10pt) so the summary line pops
     - Cols A-I: white text on existing dark-grey fill
     - Col J: bright tier-color font on dark grey (kept)
     - Gold medium top border (kept)

Idempotent. Re-running on a clean sheet is a no-op.

Usage:
    python3 .claude/scripts/fix_region_sheet_styling.py <workbook.xlsx>
"""
from __future__ import annotations
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.views import Selection
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
from openpyxl.formatting.formatting import ConditionalFormattingList


BONDS = ['KOLLAM', 'KOZHIKODE', 'ATTINGAL', 'PALAKKAD', 'KOTTARAKARA', 'KANNUR',
         'ALAPPUZHA', 'NEDUMANGAD', 'ALUVA', 'PERINTHALMANNA', 'THODUPUZHA',
         'PATHANAMTHITTA', 'KOTTAYAM', 'THRISSUR', 'TRIPUNITHURA']

FONT_NAME = 'Trebuchet MS'

# Column-header row palette + typography
COLHDR_ROW          = 4
COLHDR_HEIGHT       = 30
COLHDR_FILL_HEX     = 'FF1A237E'   # deep navy
COLHDR_FONT_HEX     = 'FFFFFFFF'   # white
COLHDR_FONT_SIZE    = 12
COLHDR_ACCENT_HEX   = 'FFFFD700'   # gold thin bottom border

# TOTAL row styling — bumped from 10pt → 14pt so it stands out as the
# summary band. Cols A-I keep white bold text on the dark-grey fill, col J
# keeps the bright tier-color font (set by fix_rating_cell_format.py).
TOTAL_HEIGHT        = 26   # 16 Jun 2026 (BUG 13): was 24; spec + DETAIL TOTAL are 26
TOTAL_FILL_HEX      = 'FF374151'   # dark grey
TOTAL_FONT_HEX      = 'FFFFFFFF'   # white (cols A-I)
TOTAL_FONT_SIZE     = 14
TOTAL_BORDER_HEX    = 'FFFFD700'   # gold accent on the top edge

# Bright tier-color font for col J on the dark-grey TOTAL fill (mirrors
# fix_rating_cell_format.BRIGHT_FONT). Computed from the TOTAL row's data.
TOTAL_TIER_FONT = {
    'high':     'FF4FC3F7',
    'balanced': 'FF81C784',
    'inv':      'FFFFB74D',
    'crit':     'FFE57373',
    'none':     'FFB0BEC5',
}

# Per-tier CF styling for col J (Rating) on data rows. MUST match the
# palette in fix_detail_row_text.DATA_RATING_STYLES so DETAIL sheets and
# bond/region sheets render identically tier-by-tier (locked 10 May 2026).
#
# (formula_text, fill_hex, font_hex)
DATA_RATING_STYLES = [
    ('"🚀 High Performance"',  'FFBBDEFB', 'FF1565C0'),  # blue
    ('"✅ Balanced"',           'FFDCEDC8', 'FF2E7D32'),  # green
    ('"⚠️ Inventory Heavy"',   'FFFFE0B2', 'FFE65100'),  # amber-orange
    ('"🚫 Critical Overstock"', 'FFFFCDD2', 'FFC62828'),  # red
    ('"— No activity"',         'FFECEFF1', 'FF6B7280'),  # grey
]
DATA_FONT_SIZE = 10


def _compute_tier(opening, receipts, sales, closing):
    o = opening or 0
    r = receipts or 0
    s = sales or 0
    c = closing or 0
    if o == 0 and r == 0 and s == 0 and c == 0:
        return 'none'
    denom = o + r
    st = (s / denom) if denom else 0.0
    if st >= 0.8:  return 'high'
    if st >= 0.6:  return 'balanced'
    if st >= 0.4:  return 'inv'
    return 'crit'


# Hero band (rows 1-3) — navy hero, locked 16 Jul 2026 (Abhay-approved, Option A).
# Full-width A:J deep-navy band: white bold title, gold-dim italic subtitle,
# soft-steel shop-count line, medium-gold rule under row 3. Fixes the old
# A:I-only near-white band that left col J unstyled (grey patch). Cell VALUES
# (title text, live COUNTA formula in r3) are never touched — styling only.
HERO_FILL_HEX     = 'FF0D1B4A'    # deepest navy (house canvas)
HERO_TITLE_HEX    = 'FFFFFFFF'
HERO_SUB_HEX      = 'FFFFD54F'    # gold-dim italic
HERO_COUNT_HEX    = 'FFB9C4E8'    # soft steel-blue
HERO_RULE_HEX     = 'FFFFB300'    # medium gold rule under the band
HERO_HEIGHTS      = {1: 34, 2: 16, 3: 18}


def _style_detail_hero(ws):
    """DETAIL-sheet rows 1-2 → navy hero band (A:G), locked 16 Jul 2026.

    Lives HERE (not in fix_detail_row_text) because this script is the FINAL
    write step of the pipeline: openpyxl round-trips DROP style-only merged
    cells (verified on openpyxl 3.1.5, 16 Jul 2026), so any hero styled by an
    earlier step loses interior fills/borders when a later step re-saves the
    workbook. Styling applied in the last save persists fully. r1 = 18pt bold
    white title (text untouched), r2 = italic gold-dim subtitle left-aligned,
    medium-gold rule across the bottom of r2. Row 3 stays the white spacer.
    """
    fill = PatternFill('solid', fgColor=HERO_FILL_HEX)
    gold_rule = Side(style='medium', color=HERO_RULE_HEX)
    for m in list(ws.merged_cells.ranges):
        if m.max_row <= 2:
            ws.unmerge_cells(str(m))
    specs = {
        1: (34, Font(name=FONT_NAME, size=18, bold=True, color=HERO_TITLE_HEX),
            Alignment(horizontal='center', vertical='center'), Border()),
        2: (16, Font(name=FONT_NAME, size=10, italic=True, color=HERO_SUB_HEX),
            Alignment(horizontal='left', vertical='center', indent=1),
            Border(bottom=gold_rule)),
    }
    for r, (h, font, align, border) in specs.items():
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        ws.row_dimensions[r].height = h
        for c in range(1, 8):
            cell = ws.cell(r, c)
            cell.fill = fill
            cell.font = font
            cell.alignment = align
            cell.border = border


_MASTER_STAFF = None


def _master_staff():
    """Lazy-load canonical shop→field-staff from MASTER DATA CONFIRMED.xlsx
    (root, sheet '16-4-25' — the single source of truth per CLAUDE.md).
    Added 16 Jul 2026 PM: 9 bonds showed stale staff because region col C is
    only written at row creation and never refreshed on reassignment. The
    final pass now syncs col C + hero + DETAIL banners from master every
    build. Non-fatal: on load failure returns {} and the hero falls back to
    whatever col C holds."""
    global _MASTER_STAFF
    if _MASTER_STAFF is not None:
        return _MASTER_STAFF
    import os
    path = os.environ.get('KSBC_MASTER_DATA_PATH') or os.path.normpath(
        os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     '..', '..', 'MASTER DATA CONFIRMED.xlsx'))
    try:
        from ksbc_alias_guard import canonical_code
        mw = load_workbook(path, read_only=True)
        out = {}
        for r in mw['16-4-25'].iter_rows(min_row=3, values_only=True):
            if not r or r[3] is None:
                continue
            try:
                code = canonical_code(int(str(r[3]).strip()))
            except (ValueError, TypeError):
                continue
            if r[5] and str(r[5]).strip().upper() == 'KSBC' and r[6]:
                out[code] = str(r[6]).strip()
        mw.close()
        _MASTER_STAFF = out
    except Exception as e:
        print(f'  ⚠ master staff load failed ({e}) — staff sync skipped')
        _MASTER_STAFF = {}
    return _MASTER_STAFF


def _sync_region_staff(ws):
    """Rewrite hidden col C from master for every data row. Returns count."""
    staff_map = _master_staff()
    if not staff_map:
        return 0
    from ksbc_alias_guard import canonical_code
    n = 0
    for r in range(5, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.strip().upper() == 'TOTAL':
            break
        try:
            code = canonical_code(int(str(v).strip()))
        except (ValueError, TypeError):
            continue
        mf = staff_map.get(code)
        if mf and str(ws.cell(r, 3).value or '').strip() != mf:
            ws.cell(r, 3).value = mf
            n += 1
    return n


def _sync_detail_staff(ws):
    """Refresh the 'Staff: <name>' suffix on DETAIL shop banners from master."""
    staff_map = _master_staff()
    if not staff_map:
        return 0
    import re as _re
    from ksbc_alias_guard import canonical_code
    from ksbc_detail_nested import _find_shop_blocks, _parse_shop_code
    n = 0
    for h, _t in _find_shop_blocks(ws):
        cell = ws.cell(h, 1)
        val = str(cell.value)
        code = _parse_shop_code(val)
        if code is None:
            continue
        mf = staff_map.get(canonical_code(code))
        if not mf:
            continue
        new = _re.sub(r'(Staff:\s*).*$', lambda m: m.group(1) + mf, val)
        if new != val:
            cell.value = new
            n += 1
    return n


def _zebra(ws):
    """Re-stripe region data rows (added 16 Jul 2026 PM). The zebra fills are
    painted at row creation and the rating sort moves rows without
    re-striping, so stripes end up scrambled. Re-applied here every build:
    alternating FFF1F5F9 / white on cols A..I, rows 5..TOTAL-1. Col J is
    NEVER touched — rating cells carry direct tier fills + CF."""
    stripe = PatternFill('solid', fgColor='FFF1F5F9')
    plain = PatternFill('solid', fgColor='FFFFFFFF')
    for r in range(5, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.strip().upper() == 'TOTAL':
            break
        fill = stripe if (r - 5) % 2 == 0 else plain
        for c in range(1, 10):
            ws.cell(r, c).fill = fill


def _unfreeze(ws):
    """Remove freeze panes WITHOUT corrupting the sheet view.

    Plain `ws.freeze_panes = None` drops the <pane> element but leaves the
    stale <selection pane="bottomLeft" .../> records behind — Excel then
    reports 'Repaired Records: View from /xl/worksheets/sheetN.xml' on every
    such sheet and strips the views (caught 16 Jul 2026 PM from Abhay's
    repair log; all 30 region+DETAIL sheets flagged). Reset the pane AND the
    selection to a single default so the view serialises clean.
    """
    ws.freeze_panes = None
    sv = ws.sheet_view
    sv.pane = None
    sv.selection = [Selection(activeCell='A1', sqref='A1')]


def _style_bp_hero(ws):
    """BOND PERFORMANCE r1 → navy hero (A:H), locked 16 Jul 2026.

    Same family as the region-sheet hero: deep-navy FF0D1B4A band, 18pt bold
    white title, medium-gold bottom rule. Row 2 (column headers) and below are
    untouched. The clusterwise rebuild (pipeline Step 8) writes its own light
    title each run — this pass (Step 12) re-enforces the hero after it.
    MERGE FIRST, STYLE AFTER (openpyxl merge_cells drops non-anchor fills).
    """
    fill = PatternFill('solid', fgColor=HERO_FILL_HEX)
    gold_rule = Side(style='medium', color=HERO_RULE_HEX)
    for m in list(ws.merged_cells.ranges):
        if m.min_row == 1 and m.max_row == 1:
            ws.unmerge_cells(str(m))
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    ws.row_dimensions[1].height = 36
    for c in range(1, 9):
        cell = ws.cell(1, c)
        cell.fill = fill
        cell.font = Font(name=FONT_NAME, size=18, bold=True, color=HERO_TITLE_HEX)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(bottom=gold_rule)


def _style_hero(ws):
    """Rows 1-3 → full-width navy hero band (A:J), idempotent.

    Unmerges every merged range living entirely inside rows 1-3 (the legacy
    band merged A:I only), fills all 30 cells A1:J3 BEFORE re-merging A:J per
    row (anti-hairline convention), sets fonts/heights/alignment, and draws a
    medium-gold bottom rule across row 3. Values/formulas untouched.
    """
    fill = PatternFill('solid', fgColor=HERO_FILL_HEX)
    gold_rule = Side(style='medium', color=HERO_RULE_HEX)
    for m in list(ws.merged_cells.ranges):
        if m.min_row >= 1 and m.max_row <= 3:
            ws.unmerge_cells(str(m))
    # Field staff for the hero (Abhay 16 Jul PM: staff column hidden, name
    # shown in the header instead). Read from col C data rows — hidden but
    # always populated by the driver.
    staff = []
    for r in range(5, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if isinstance(a, str) and a.strip().upper() == 'TOTAL':
            break
        v = ws.cell(r, 3).value
        if isinstance(v, str) and v.strip() and v.strip() not in staff:
            staff.append(v.strip())
    fonts = {
        1: Font(name=FONT_NAME, size=18, bold=True, color=HERO_TITLE_HEX),
        2: Font(name=FONT_NAME, size=10, italic=True, color=HERO_SUB_HEX),
        3: Font(name=FONT_NAME, size=10, color=HERO_COUNT_HEX),
    }
    aligns = {
        1: Alignment(horizontal='center', vertical='center'),
        2: Alignment(horizontal='left', vertical='center', indent=1),
        3: Alignment(horizontal='right', vertical='center', indent=1),
    }
    for r in (1, 2, 3):
        ws.row_dimensions[r].height = HERO_HEIGHTS[r]
        # MERGE FIRST, style after: openpyxl's merge_cells DROPS non-anchor
        # cells (carrying only borders), so any fill set before the merge is
        # wiped and Excel renders the band white right of col A. Styling the
        # MergedCell objects AFTER the merge persists all 10 fills (verified
        # in saved XML, 16 Jul 2026).
        if r == 2:
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=5)
            ws.merge_cells(start_row=2, start_column=6, end_row=2, end_column=10)
        else:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
        for c in range(1, 11):
            cell = ws.cell(r, c)
            cell.fill = fill
            cell.font = fonts[r]
            cell.alignment = aligns[r]
            cell.border = Border(bottom=gold_rule) if r == 3 else Border()
    # r2 right half = field staff (bold gold-dim, right-aligned)
    f2 = ws.cell(2, 6)
    f2.value = ('Field Staff: ' + ' · '.join(staff)) if staff else None
    f2.font = Font(name=FONT_NAME, size=10, bold=True, color=HERO_SUB_HEX)
    f2.alignment = Alignment(horizontal='right', vertical='center', indent=1)


def _style_column_header(ws):
    """Style the column-header row (r4) — taller, bigger font, gold accent."""
    fill = PatternFill('solid', fgColor=COLHDR_FILL_HEX)
    gold_thin = Side(style='thin', color=COLHDR_ACCENT_HEX)
    ws.row_dimensions[COLHDR_ROW].height = COLHDR_HEIGHT
    for c in range(1, 11):  # A..J
        cell = ws.cell(COLHDR_ROW, c)
        cell.fill = fill
        cell.font = Font(name=FONT_NAME, size=COLHDR_FONT_SIZE,
                         bold=True, italic=False, color=COLHDR_FONT_HEX)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        existing = cell.border
        cell.border = Border(
            top=existing.top if existing and existing.top else None,
            bottom=gold_thin,
            left=existing.left if existing and existing.left else None,
            right=existing.right if existing and existing.right else None,
        )


def _resolve_sum(ws, formula, fallback_col):
    """Resolve an =SUM(...) formula to a numeric value by reading the source
    cells. Handles single-range, multi-range, and cell-list forms.

    16 Jun 2026 (BUG 15): was single-range only — a multi-range TOTAL (e.g.
    =SUM(D5:D9,D11:D15)) returned 0 and mis-coloured the region TOTAL rating.
    """
    import re
    if not isinstance(formula, str):
        return formula or 0
    m = re.match(r'^=SUM\((.*)\)$', formula.replace(' ', ''))
    if not m:
        return 0
    total = 0.0
    for part in m.group(1).split(','):
        rng = re.match(r'^([A-Z]+)(\d+):([A-Z]+)(\d+)$', part)
        one = re.match(r'^([A-Z]+)(\d+)$', part)
        if rng and rng.group(1) == rng.group(3) and len(rng.group(1)) == 1:
            col = ord(rng.group(1)) - ord('A') + 1
            for r in range(int(rng.group(2)), int(rng.group(4)) + 1):
                v = ws.cell(r, col).value
                if isinstance(v, (int, float)):
                    total += v
        elif one and len(one.group(1)) == 1:
            col = ord(one.group(1)) - ord('A') + 1
            v = ws.cell(int(one.group(2)), col).value
            if isinstance(v, (int, float)):
                total += v
    return total


def _style_total_row(ws):
    """Find the TOTAL row at the bottom and apply tall + bold + tier styling."""
    total_r = None
    for r in range(ws.max_row, 0, -1):
        if ws.cell(r, 1).value == 'TOTAL':
            total_r = r
            break
    if total_r is None:
        return None

    # Compute tier from the TOTAL's resolved B-E equivalents (D-G in this layout).
    op = _resolve_sum(ws, ws.cell(total_r, 4).value, 4)
    rc = _resolve_sum(ws, ws.cell(total_r, 5).value, 5)
    sa = _resolve_sum(ws, ws.cell(total_r, 6).value, 6)
    cl = _resolve_sum(ws, ws.cell(total_r, 7).value, 7)
    tier = _compute_tier(op, rc, sa, cl)
    j_font = TOTAL_TIER_FONT[tier]

    fill = PatternFill('solid', fgColor=TOTAL_FILL_HEX)
    top_border = Side(style='medium', color=TOTAL_BORDER_HEX)
    ws.row_dimensions[total_r].height = TOTAL_HEIGHT

    for c in range(1, 11):  # A..J
        cell = ws.cell(total_r, c)
        cell.fill = fill
        font_color = j_font if c == 10 else TOTAL_FONT_HEX
        cell.font = Font(name=FONT_NAME, size=TOTAL_FONT_SIZE,
                         bold=True, italic=False, color=font_color)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        existing = cell.border
        cell.border = Border(
            top=top_border,
            bottom=existing.bottom if existing and existing.bottom else None,
            left=existing.left if existing and existing.left else None,
            right=existing.right if existing and existing.right else None,
        )
    return total_r


def _rebuild_j_cf(ws):
    """Wipe ALL CF on the sheet and re-apply uniform col J (Rating) rules
    that match the DETAIL-sheet palette tier-for-tier. Excel quirk: solid
    CF fills MUST be set via bgColor inside the dxf, not fgColor — see
    .claude/memory/ksbc-detail-cf.md.

    Region sheet: data rows are 5..(TOTAL-1). TOTAL row col J is styled
    directly by _style_total_row above (dark-grey + bright tier font), so
    no CF is applied to the TOTAL row.
    """
    total_r = None
    for r in range(ws.max_row, 0, -1):
        if ws.cell(r, 1).value == 'TOTAL':
            total_r = r
            break
    if not total_r or total_r <= 5:
        return
    data_range = f'J5:J{total_r - 1}'
    ws.conditional_formatting = ConditionalFormattingList()
    for priority, (txt, fill_color, font_color) in enumerate(DATA_RATING_STYLES, start=1):
        dxf = DifferentialStyle(
            fill=PatternFill(bgColor=fill_color),
            font=Font(name=FONT_NAME, size=DATA_FONT_SIZE, bold=True, color=font_color),
        )
        rule = Rule(type='cellIs', operator='equal', formula=[txt],
                    stopIfTrue=False, dxf=dxf)
        rule.priority = priority
        ws.conditional_formatting.add(data_range, rule)


def _enforce_pct_whole_numbers(ws):
    """Region sheets: cols H (Sell-Through %) and I (Closing vs Sales %) on
    every data row + TOTAL row use whole-number percent format `0%`. No
    decimals (Abhay's rule, 10 May 2026)."""
    total_r = None
    for r in range(ws.max_row, 0, -1):
        if ws.cell(r, 1).value == 'TOTAL':
            total_r = r
            break
    if not total_r:
        return
    for r in range(5, total_r + 1):
        ws.cell(r, 8).number_format = '0%'   # Sell-Through %
        ws.cell(r, 9).number_format = '0%'   # Closing vs Sales %


def fix_region_sheet(ws):
    _sync_region_staff(ws)   # col C from master BEFORE the hero reads it
    _zebra(ws)               # re-stripe rows scrambled by the rating sort
    _style_hero(ws)
    _style_column_header(ws)
    _style_total_row(ws)
    _rebuild_j_cf(ws)
    _enforce_pct_whole_numbers(ws)
    # Cols A (Shop Code) + C (Field Staff) HIDDEN, not deleted (Abhay 16 Jul
    # 2026 PM). Every consumer (driver row-matching, alias guard, sorter, BP
    # builder, COUNTA shop count) still reads the values; only the view
    # changes. Staff shows in the hero (r2 right); the TOTAL label is
    # mirrored into col B so the anchor value in A stays for the scripts.
    ws.column_dimensions['A'].hidden = True
    ws.column_dimensions['C'].hidden = True
    _unfreeze(ws)   # no frozen panes on region sheets (Abhay 16 Jul 2026 PM)
    ws.sheet_view.showGridLines = False   # gridlines off (Abhay 16 Jul 2026 PM)
    for r in range(5, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if isinstance(a, str) and a.strip().upper() == 'TOTAL':
            b = ws.cell(r, 2)
            b.value = 'TOTAL'
            b.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            break


# ---------------------------------------------------------------------------
# BOND PERFORMANCE — col H is Rating. Layout (locked):
#   r1: title; r2: column headers
#   r3, r10, r16: cluster banner rows (navy fill, bright tier font — kept)
#   r4-r9, r11-r15, r17-r20: bond rows (data — CF applies here)
#   r21 (last): TOTAL row (dark grey + bright tier font — kept)
# ---------------------------------------------------------------------------
def _enforce_bp_pct_whole_numbers(ws):
    """BOND PERFORMANCE: cols F + G are percentages on every row r3..TOTAL.
    Force whole-number format (Abhay's rule, 10 May 2026)."""
    total_r = None
    for r in range(ws.max_row, 0, -1):
        if ws.cell(r, 1).value == 'TOTAL':
            total_r = r
            break
    if not total_r:
        return
    for r in range(3, total_r + 1):
        ws.cell(r, 6).number_format = '0%'
        ws.cell(r, 7).number_format = '0%'


def _rebuild_bond_performance_cf(ws):
    """Apply uniform col H CF rules to the BOND data rows only. Cluster
    banner rows + TOTAL row keep their navy / dark-grey fills set by other
    scripts (any CF fill on those would clash with the banner look)."""
    # Find TOTAL row + cluster banner rows
    total_r = None
    banner_rows = set()
    for r in range(2, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a == 'TOTAL':
            total_r = r
            break
        if isinstance(a, str) and a.strip().lower().startswith('cluster'):
            banner_rows.add(r)
    if total_r is None:
        return
    # Build contiguous data ranges (skip cluster banner rows)
    data_ranges = []
    cur_start = None
    for r in range(3, total_r):
        if r in banner_rows:
            if cur_start is not None:
                data_ranges.append(f'H{cur_start}:H{r-1}' if cur_start != r-1
                                   else f'H{cur_start}')
                cur_start = None
        else:
            if cur_start is None:
                cur_start = r
    if cur_start is not None:
        last = total_r - 1
        data_ranges.append(f'H{cur_start}:H{last}' if cur_start != last
                           else f'H{cur_start}')
    if not data_ranges:
        return
    ws.conditional_formatting = ConditionalFormattingList()
    range_str = ' '.join(data_ranges)
    for priority, (txt, fill_color, font_color) in enumerate(DATA_RATING_STYLES, start=1):
        dxf = DifferentialStyle(
            fill=PatternFill(bgColor=fill_color),
            font=Font(name=FONT_NAME, size=DATA_FONT_SIZE, bold=True, color=font_color),
        )
        rule = Rule(type='cellIs', operator='equal', formula=[txt],
                    stopIfTrue=False, dxf=dxf)
        rule.priority = priority
        ws.conditional_formatting.add(range_str, rule)


# ---------------------------------------------------------------------------
# DASHBOARD — find every cell that contains a rating string, infer the
# col, and apply uniform CF on that column's continuous range. The
# dashboard layout has its rating column at K49:K63 (locked May 2026 build),
# but we discover ranges dynamically so a layout shift doesn't break us.
# ---------------------------------------------------------------------------
RATING_TOKENS = ('High Performance', 'Balanced', 'Inventory Heavy',
                 'Critical Overstock', 'No activity')


def _rebuild_dashboard_cf(ws):
    """Find rating-bearing cells in the DASHBOARD and apply uniform CF."""
    # Group rating cells by column → contiguous row ranges.
    by_col = {}
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and any(t in v for t in RATING_TOKENS):
                by_col.setdefault(c, []).append(r)
    if not by_col:
        return
    ranges = []
    for c, rows in by_col.items():
        rows.sort()
        col_letter = chr(64 + c) if c <= 26 else None
        if col_letter is None:
            continue
        # Build contiguous ranges
        start = rows[0]
        prev = rows[0]
        for r in rows[1:] + [None]:
            if r is None or r != prev + 1:
                ranges.append(f'{col_letter}{start}:{col_letter}{prev}'
                              if start != prev else f'{col_letter}{start}')
                if r is not None:
                    start = r
            if r is not None:
                prev = r
    if not ranges:
        return
    # DASHBOARD has its own CF for KPI bars / data bars elsewhere — DON'T
    # nuke the entire ConditionalFormattingList. But DO remove any existing
    # rating-rule CF (cellIs equal to "🚀 High Performance" etc.) so we
    # don't end up with duplicate / conflicting rules from prior runs.
    rating_formulas = {style[0] for style in DATA_RATING_STYLES}
    new_cf = ConditionalFormattingList()
    for cf_range, rules in ws.conditional_formatting._cf_rules.items():
        kept = [r for r in rules
                if not (r.formula and len(r.formula) == 1
                        and r.formula[0] in rating_formulas)]
        if kept:
            sqref = str(cf_range.sqref) if hasattr(cf_range, 'sqref') else str(cf_range)
            for r in kept:
                new_cf.add(sqref, r)
    ws.conditional_formatting = new_cf
    range_str = ' '.join(ranges)
    for priority, (txt, fill_color, font_color) in enumerate(DATA_RATING_STYLES, start=1):
        dxf = DifferentialStyle(
            fill=PatternFill(bgColor=fill_color),
            font=Font(name=FONT_NAME, size=DATA_FONT_SIZE, bold=True, color=font_color),
        )
        rule = Rule(type='cellIs', operator='equal', formula=[txt],
                    stopIfTrue=False, dxf=dxf)
        rule.priority = priority
        ws.conditional_formatting.add(range_str, rule)


def apply_all(wb):
    """Apply the full styling pass to an ALREADY-LOADED workbook.

    Split out of main() on 16 Jul 2026 so ksbc_region_insights.py (the
    pipeline's final write step — openpyxl round-trips drop charts AND
    style-only merged cells) can re-apply every hero/CF in ITS session
    just before the one final save."""

    # 1. Region/bond sheets — col-header + TOTAL + rating CF
    for bond in BONDS:
        if bond not in wb.sheetnames:
            print(f'  · {bond:18s}  (missing)')
            continue
        fix_region_sheet(wb[bond])
        print(f'  ✓ {bond:18s}  styled (hero + col-header + TOTAL + rating CF)')

    # 1.5 DETAIL sheets — navy hero band rows 1-2 (final-writer pass)
    for bond in BONDS:
        dn = f'{bond} DETAIL'
        if dn in wb.sheetnames:
            _sync_detail_staff(wb[dn])
            _style_detail_hero(wb[dn])
            _unfreeze(wb[dn])   # no frozen panes on DETAIL (Abhay 16 Jul 2026 PM)
    print('  \u2713 DETAIL sheets       navy hero (rows 1-2) on all present')

    # 2. BOND PERFORMANCE — uniform col H rating CF on bond data rows + whole-% format
    if 'BOND PERFORMANCE' in wb.sheetnames:
        _style_bp_hero(wb['BOND PERFORMANCE'])
        _rebuild_bond_performance_cf(wb['BOND PERFORMANCE'])
        _enforce_bp_pct_whole_numbers(wb['BOND PERFORMANCE'])
        print(f'  ✓ BOND PERFORMANCE   navy hero + col H rating CF rebuilt + whole-% format')

    # 3. DASHBOARD — uniform rating CF wherever the rating column lives
    if 'DASHBOARD' in wb.sheetnames:
        _rebuild_dashboard_cf(wb['DASHBOARD'])
        print(f'  ✓ DASHBOARD          rating CF appended')


def main():
    if len(sys.argv) < 2:
        print('Usage: python3 fix_region_sheet_styling.py <workbook.xlsx>')
        sys.exit(1)
    target = sys.argv[1]
    wb = load_workbook(target, data_only=False)
    apply_all(wb)
    wb.save(target)
    print(f'\nSaved {target}.')


if __name__ == '__main__':
    main()
