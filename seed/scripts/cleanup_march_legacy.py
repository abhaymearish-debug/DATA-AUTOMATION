"""
ONE-SHOT cleanup for the March DETAIL sheets — converts the hybrid layout
(nested Brand→Pack section PLUS leftover legacy BRAND-WISE / PACK-WISE
tables) into pure nested layout matching April and May.

Per shop block, removes:
  - The "BRAND-WISE" label row (between navy header and col header)
  - The "Brand" col header is renamed to "Brand / Pack"
  - The entire PACK-WISE section after the shop's TOTAL:
      "PACK-WISE" label + "Pack" col headers + N pack rows + PACK-WISE TOTAL
      + any trailing blank rows up to the next block's header

Then for each remaining shop block:
  - Rewrites the BRAND-WISE TOTAL row's B-E as CONSTANTS equal to the sum of
    BRAND rows only (the legacy formula =SUM(B7:B27) double-counts because
    the nested layout has both brand rows AND their pack sub-rows; summing
    everything gives 2x the true total).
  - Rewrites F and G formulas to reference the TOTAL row's new position.

Batched formula-row-ref shift at the end so the operation is O(M) cells
rather than O(N × M) (where N is the number of deletions).

Run AFTER fix_detail_row_text.py — the deletes shift rows, then the styling
sweep re-paints the now-clean blocks.

Usage:
    python3 cleanup_march_legacy.py <workbook.xlsx>
"""
from __future__ import annotations
import re
import sys
import bisect
from openpyxl import load_workbook

BONDS = ['KOLLAM', 'KOZHIKODE', 'ATTINGAL', 'PALAKKAD', 'KOTTARAKARA', 'KANNUR',
         'ALAPPUZHA', 'NEDUMANGAD', 'ALUVA', 'PERINTHALMANNA', 'THODUPUZHA',
         'PATHANAMTHITTA', 'KOTTAYAM', 'THRISSUR', 'TRIPUNITHURA']

PACK_RE = re.compile(r'^\s*\d+\s*ML\s*$', re.IGNORECASE)
CELL_REF_RE = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def _is_header(v):
    return isinstance(v, str) and ' — ' in v and 'Staff:' in v


def _is_pack(v):
    return isinstance(v, str) and bool(PACK_RE.match(v))


def cleanup_sheet(ws):
    """Remove legacy artifacts from every block on this DETAIL sheet."""
    headers = [r for r in range(1, ws.max_row + 1)
               if _is_header(ws.cell(r, 1).value)]
    if not headers:
        return 0, 0

    # Per-block plan: collect rows to delete (BRAND-WISE label + everything
    # after the FIRST TOTAL up to the next header).
    rows_to_delete = set()
    for i, h in enumerate(headers):
        end = headers[i + 1] - 1 if i + 1 < len(headers) else ws.max_row
        # 1. Find BRAND-WISE label row (typically h+1)
        for r in range(h + 1, min(h + 4, end + 1)):
            if ws.cell(r, 1).value == 'BRAND-WISE':
                rows_to_delete.add(r)
                break
        # 2. Find the FIRST TOTAL row (BRAND-WISE TOTAL = the true shop total)
        first_total = None
        for r in range(h + 2, end + 1):
            if ws.cell(r, 1).value == 'TOTAL':
                first_total = r
                break
        if first_total is None:
            continue
        # 3. Mark all rows AFTER the first TOTAL up to end-of-block for deletion
        #    (these are: blank, PACK-WISE label, Pack col header, pack rows,
        #    PACK-WISE TOTAL, trailing blanks)
        for r in range(first_total + 1, end + 1):
            rows_to_delete.add(r)

    if not rows_to_delete:
        return 0, 0

    deleted_sorted = sorted(rows_to_delete)  # ascending

    # Delete bottom-up so positions don't shift mid-loop
    for r in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(r)

    # Batched formula-ref shift: any row reference R in any formula gets
    # decremented by count(deleted_rows <= R).
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not (isinstance(v, str) and v.startswith('=')):
                continue
            def _bump(m):
                cd, col, rd, n_str = m.group(1), m.group(2), m.group(3), m.group(4)
                n = int(n_str)
                shift = bisect.bisect_right(deleted_sorted, n)
                if shift > 0:
                    n -= shift
                return f"{cd}{col}{rd}{n}"
            new_v = CELL_REF_RE.sub(_bump, v)
            if new_v != v:
                cell.value = new_v

    # Re-find blocks at their new positions, fix col header text + TOTAL row.
    headers_new = [r for r in range(1, ws.max_row + 1)
                   if _is_header(ws.cell(r, 1).value)]
    blocks_fixed = 0
    for i, h in enumerate(headers_new):
        end = headers_new[i + 1] - 1 if i + 1 < len(headers_new) else ws.max_row
        # Col header row is now directly after header (h+1) since BRAND-WISE was deleted
        col_hdr_r = h + 1
        if col_hdr_r <= ws.max_row and ws.cell(col_hdr_r, 1).value == 'Brand':
            ws.cell(col_hdr_r, 1).value = 'Brand / Pack'

        # Find the TOTAL row (now the only TOTAL in the block)
        total_r = None
        for r in range(h + 2, end + 1):
            if ws.cell(r, 1).value == 'TOTAL':
                total_r = r
                break
        if total_r is None:
            continue

        # Recompute TOTAL as constants — sum of BRAND rows only (skip pack
        # sub-rows to avoid the double-count of the legacy =SUM(B7:B27)).
        op = rc = sl = cl = 0.0
        for r in range(h + 2, total_r):
            label = ws.cell(r, 1).value
            if label is None:
                continue
            if isinstance(label, str) and PACK_RE.match(label):
                continue   # skip pack sub-rows
            if label == 'Brand / Pack':  # skip col header
                continue
            op += float(ws.cell(r, 2).value or 0)
            rc += float(ws.cell(r, 3).value or 0)
            sl += float(ws.cell(r, 4).value or 0)
            cl += float(ws.cell(r, 5).value or 0)
        ws.cell(total_r, 2).value = op
        ws.cell(total_r, 3).value = rc
        ws.cell(total_r, 4).value = sl
        ws.cell(total_r, 5).value = cl
        for c in (2, 3, 4, 5):
            ws.cell(total_r, c).number_format = '0.00'
        ws.cell(total_r, 6).number_format = '0%'
        ws.cell(total_r, 6).value = (
            f'=IF(AND(B{total_r}=0,C{total_r}=0,D{total_r}=0,E{total_r}=0),"—",'
            f'IFERROR(D{total_r}/(B{total_r}+C{total_r}),0))'
        )
        ws.cell(total_r, 7).value = (
            f'=IF(AND(B{total_r}=0,C{total_r}=0,D{total_r}=0,E{total_r}=0),"— No activity",'
            f'IF(F{total_r}>=0.8,"🚀 High Performance",'
            f'IF(F{total_r}>=0.6,"✅ Balanced",'
            f'IF(F{total_r}>=0.4,"⚠️ Inventory Heavy","🚫 Critical Overstock"))))'
        )
        blocks_fixed += 1

    return len(rows_to_delete), blocks_fixed


def main():
    if len(sys.argv) < 2:
        print('Usage: python3 cleanup_march_legacy.py <workbook.xlsx>')
        sys.exit(1)
    target = sys.argv[1]
    wb = load_workbook(target, data_only=False)
    total_deleted = total_fixed = 0
    for bond in BONDS:
        sn = f'{bond} DETAIL'
        if sn not in wb.sheetnames:
            print(f'  · {sn:22s}  (missing)')
            continue
        deleted, fixed = cleanup_sheet(wb[sn])
        total_deleted += deleted
        total_fixed += fixed
        print(f'  ✓ {sn:22s}  deleted {deleted:>4} rows, fixed {fixed:>3} TOTALs')
    wb.save(target)
    print(f'\nTotal: {total_deleted} legacy rows deleted, {total_fixed} TOTAL rows fixed. Saved {target}.')


if __name__ == '__main__':
    main()
