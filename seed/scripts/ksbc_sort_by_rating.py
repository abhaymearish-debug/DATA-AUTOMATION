"""
KSBC sort-by-rating script.

Sorts every KSBC region sheet and its matching DETAIL sheet by sell-through
percentage, descending. This matches the rating tier order:

    🚀 High Performance (≥80%)  →  ✅ Balanced (≥60%)  →
    ⚠️ Inventory Heavy (≥40%)  →  🚫 Critical Overstock (<40%)

Region sheet: each shop row (between the header at row 4 and the TOTAL row)
is sorted by its sell-through ratio Sales/(Opening+Receipts) descending.
Formulas in the rating/sell-through columns are translated to the new row.

DETAIL sheet: each shop's variable-sized block (shop header → PACK-WISE
TOTAL row + trailing blanks) is sorted as a single unit using the block's
brand-TOTAL sell-through as key. Formulas inside each block are translated
by the row offset so internal references stay valid.

This is MANDATORY as the final step of any KSBC build — never skip it. See
MEMORY.md. Run standalone:

    python3 .claude/scripts/ksbc_sort_by_rating.py <workbook.xlsx>
"""

from __future__ import annotations
import sys
from copy import copy
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter


MAX_COL_REGION = 10   # A..J
MAX_COL_DETAIL = 7    # A..G
REGION_HEADER_ROW = 4
DETAIL_START_ROW = 4

# Canonical border colors for region sheets
GOLD_HEX    = "FFFFD700"
SILVER_HEX  = "FFC0C0C0"
TOTAL_SEP_SIDE = Side(style="thick", color=GOLD_HEX)
ROW_SEP_SIDE   = Side(style="thin",  color=SILVER_HEX)


def _capture_cell(cell):
    return {
        "value": cell.value,
        "number_format": cell.number_format,
        "font": copy(cell.font),
        "fill": copy(cell.fill),
        "alignment": copy(cell.alignment),
        "border": copy(cell.border),
        "protection": copy(cell.protection),
        "orig_coord": cell.coordinate,
    }


def _restore_cell(cell, data, translate_to=None):
    val = data["value"]
    if translate_to and isinstance(val, str) and val.startswith("="):
        val = Translator(val, origin=data["orig_coord"]).translate_formula(translate_to)
    cell.value = val
    cell.number_format = data["number_format"]
    cell.font = data["font"]
    cell.fill = data["fill"]
    cell.alignment = data["alignment"]
    cell.border = data["border"]
    cell.protection = data["protection"]


def _sellthrough(opening, receipts, sales):
    denom = (opening or 0) + (receipts or 0)
    return (sales or 0) / denom if denom else 0.0


def _code_key(v):
    """Deterministic, type-safe sort key for a shop code used as a tie-break.
    Numeric codes sort before non-numeric; both are internally ordered."""
    try:
        return (0, int(v))
    except (TypeError, ValueError):
        return (1, str(v))


import re as _re
def _header_code_key(header_value):
    """Tie-break key for a DETAIL shop block: the first 3-6 digit run in the
    banner label (the shop code), else the raw label string."""
    m = _re.search(r'(\d{3,6})', str(header_value)) if header_value is not None else None
    return (0, int(m.group(1))) if m else (1, str(header_value))


# ----------------------------------------------------------------------------
# Region sheet sort
# ----------------------------------------------------------------------------
def sort_region_sheet(ws_live, ws_values):
    """Sort shop rows in-place on ws_live; ws_values supplies cached numbers.

    Region table borders are POSITIONAL (the thick gold separator above the
    TOTAL row belongs to the last shop-row position, not to whichever shop
    currently sits there). We snapshot borders per position BEFORE sorting and
    re-apply them AFTER sorting so the gold lining stays on the row
    immediately above TOTAL no matter which shop gets sorted into it.
    """
    # Find TOTAL row
    total_row = None
    for r in range(REGION_HEADER_ROW + 1, ws_live.max_row + 1):
        v = ws_live.cell(row=r, column=1).value
        if v == "TOTAL":
            total_row = r
            break
    if total_row is None:
        return 0

    start_row = REGION_HEADER_ROW + 1
    end_row = total_row - 1
    if end_row < start_row:
        return 0

    # Snapshot positional borders (row, col) -> Border  for shop rows + TOTAL.
    positional_borders = {}
    for r in range(start_row, total_row + 1):
        for c in range(1, MAX_COL_REGION + 1):
            positional_borders[(r, c)] = copy(ws_live.cell(row=r, column=c).border)

    # Capture every row (as a list of cell dicts) + the sort key from ws_values
    rows = []
    for r in range(start_row, end_row + 1):
        cells = [_capture_cell(ws_live.cell(row=r, column=c))
                 for c in range(1, MAX_COL_REGION + 1)]
        # Use cached numeric values from ws_values for sort
        opening  = ws_values.cell(row=r, column=4).value   # D
        receipts = ws_values.cell(row=r, column=5).value   # E
        sales    = ws_values.cell(row=r, column=6).value   # F
        code     = ws_values.cell(row=r, column=1).value   # A (shop code, tie-break)
        rows.append((cells, _sellthrough(opening, receipts, sales), _code_key(code)))

    # Deterministic tie-break (fix 1 Jun 2026): equal sell-through rows are
    # ordered by shop code ascending so the sheet doesn't reshuffle build-to-
    # build with no data change. Primary key sell-through DESC.
    rows.sort(key=lambda t: (-(t[1] or 0), t[2]))

    # Write back in sorted order, translating formulas to new row
    for new_offset, (cells, _st, _ck) in enumerate(rows):
        new_row = start_row + new_offset
        for c_idx, data in enumerate(cells, start=1):
            cell = ws_live.cell(row=new_row, column=c_idx)
            coord = f"{get_column_letter(c_idx)}{new_row}"
            _restore_cell(cell, data, translate_to=coord)

    # Re-apply positional borders (overrides the ones carried over by restore)
    for (r, c), border in positional_borders.items():
        ws_live.cell(row=r, column=c).border = border

    # Idempotent enforcement: thick gold on last shop row's bottom, thin
    # silver on all other shop rows. This fixes any pre-existing drift and
    # guarantees the separator lives with the position, not the shop.
    _normalize_region_borders(ws_live, start_row, end_row)

    return len(rows)


def _normalize_region_borders(ws_live, start_row, end_row):
    """Ensure: last shop row has gold thick bottom; others have silver thin.

    Overwrites only the bottom side of each cell's border. Other sides are
    preserved so the cyan header-separator and outer frame stay intact.
    """
    for r in range(start_row, end_row + 1):
        target_side = TOTAL_SEP_SIDE if r == end_row else ROW_SEP_SIDE
        for c in range(1, MAX_COL_REGION + 1):
            cell = ws_live.cell(row=r, column=c)
            b = cell.border
            cell.border = Border(
                left=b.left, right=b.right, top=b.top, bottom=target_side,
                diagonal=b.diagonal, diagonal_direction=b.diagonal_direction,
                outline=b.outline, vertical=b.vertical, horizontal=b.horizontal,
            )
    # Also clear any gold bottom border that might remain on the first shop
    # row (row 5) top — which is the header-separator, handled elsewhere.


# ----------------------------------------------------------------------------
# DETAIL sheet sort
# ----------------------------------------------------------------------------
_PACK_RE_SORT = __import__('re').compile(r'^\s*[\d.,]+\s*ML\b', __import__('re').I)


def _detail_block_sellthrough(ws_values, start_row, end_row):
    """Compute the shop block's sell-through for sorting.

    NESTED layout (canonical since 23 Apr 2026):
        [start]     "<code> — ... | Staff: ..."   (banner)
        [start+1]   column headers (Brand / Pack, Opening, ...)
        [start+2..] brand summary rows, each followed by its nested pack rows,
                    until the "TOTAL" row.

    Sum ONLY brand summary rows — a brand row already aggregates its packs, so
    summing the pack rows too would double-count. Brand-row B/C/D are literal
    numbers (not formulas), so this is immune to the stale-cache problem that
    affects the TOTAL row's SUM formula.

    16 Jun 2026 fix (BUG 3): was `start_row + 3` with a flat-brand assumption
    carried over from the OLD two-table layout. That skipped each block's first
    brand row AND double-counted the rest (brand row + its packs), producing a
    wrong sort key and DETAIL blocks out of rating order.
    """
    opening = receipts = sales = 0.0
    for r in range(start_row + 2, end_row + 1):
        first = ws_values.cell(row=r, column=1).value
        if first == "TOTAL":
            break
        if first is None or (isinstance(first, str) and first.strip() == ''):
            continue
        if isinstance(first, str) and _PACK_RE_SORT.match(first):
            continue  # pack row — already counted in its brand summary
        opening  += ws_values.cell(row=r, column=2).value or 0
        receipts += ws_values.cell(row=r, column=3).value or 0
        sales    += ws_values.cell(row=r, column=4).value or 0
    return _sellthrough(opening, receipts, sales)


def _find_detail_blocks(ws):
    """Return list of (start_row, end_row_inclusive) for each shop block."""
    headers = []
    for r in range(DETAIL_START_ROW, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and " — " in v and "Staff:" in v:
            headers.append(r)
    blocks = []
    for i, h in enumerate(headers):
        end = (headers[i + 1] - 1) if i + 1 < len(headers) else ws.max_row
        blocks.append((h, end))
    return blocks


def sort_detail_sheet(ws_live, ws_values):
    blocks = _find_detail_blocks(ws_live)
    if not blocks:
        return 0

    # Record merged ranges per block (as (row_offset_from_block_start,
    # min_col, max_col, row_span)) so we can re-create them at new positions.
    merges_per_block: list[list[tuple[int, int, int, int]]] = [[] for _ in blocks]
    all_merges = list(ws_live.merged_cells.ranges)
    for mr in all_merges:
        for b_i, (start, end) in enumerate(blocks):
            if start <= mr.min_row and mr.max_row <= end:
                merges_per_block[b_i].append((
                    mr.min_row - start,
                    mr.min_col,
                    mr.max_col,
                    mr.max_row - mr.min_row,
                ))
                break
    # Unmerge everything in the blocks range so cells become writable
    for mr in all_merges:
        if blocks[0][0] <= mr.min_row <= blocks[-1][1]:
            ws_live.unmerge_cells(str(mr))

    # Capture block cells + sort key (after unmerging so MergedCell → Cell)
    captured = []
    for b_i, (start, end) in enumerate(blocks):
        block_cells = []
        for r in range(start, end + 1):
            row_cells = [_capture_cell(ws_live.cell(row=r, column=c))
                         for c in range(1, MAX_COL_DETAIL + 1)]
            block_cells.append(row_cells)
        key = _detail_block_sellthrough(ws_values, start, end)
        code_key = _header_code_key(ws_values.cell(row=start, column=1).value)
        captured.append((start, end, block_cells, key, merges_per_block[b_i], code_key))

    # Primary key sell-through DESC; deterministic shop-code ASC tie-break
    # (fix 1 Jun 2026) so equal-sell-through blocks don't reshuffle build-to-build.
    captured.sort(key=lambda t: (-(t[3] or 0), t[5]))

    # Clear all block rows
    first_start = blocks[0][0]
    last_end    = blocks[-1][1]
    for r in range(first_start, last_end + 1):
        for c in range(1, MAX_COL_DETAIL + 1):
            ws_live.cell(row=r, column=c).value = None

    # Rewrite blocks in new order + re-merge
    cursor = first_start
    for (old_start, old_end, block_cells, _, block_merges, _ck) in captured:
        block_len = old_end - old_start + 1
        for b_i, row_cells in enumerate(block_cells):
            new_row = cursor + b_i
            for c_idx, data in enumerate(row_cells, start=1):
                cell = ws_live.cell(row=new_row, column=c_idx)
                coord = f"{get_column_letter(c_idx)}{new_row}"
                _restore_cell(cell, data, translate_to=coord)
        for (row_off, min_col, max_col, row_span) in block_merges:
            new_r_start = cursor + row_off
            ws_live.merge_cells(
                start_row=new_r_start, end_row=new_r_start + row_span,
                start_column=min_col, end_column=max_col,
            )
        cursor += block_len

    return len(captured)


# ----------------------------------------------------------------------------
# Driver
# ----------------------------------------------------------------------------
def sort_workbook(path):
    wb_live = load_workbook(path)
    wb_vals = load_workbook(path, data_only=True)
    regions = [n for n in wb_live.sheetnames
               if f"{n} DETAIL" in wb_live.sheetnames and n != "MASTER DATA"]
    totals = {}
    for r in regions:
        n_rows = sort_region_sheet(wb_live[r], wb_vals[r])
        n_blocks = sort_detail_sheet(wb_live[f"{r} DETAIL"], wb_vals[f"{r} DETAIL"])
        totals[r] = (n_rows, n_blocks)
    wb_live.save(path)
    return totals


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("usage: ksbc_sort_by_rating.py <workbook.xlsx>")
        sys.exit(2)
    totals = sort_workbook(sys.argv[1])
    for region, (n_rows, n_blocks) in totals.items():
        print(f"  {region}: sorted {n_rows} shop rows, {n_blocks} DETAIL blocks")
    print("done.")
