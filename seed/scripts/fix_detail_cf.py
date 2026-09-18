"""
Rebuild CF + TOTAL-row canonical styling on every <BOND> DETAIL sheet.

Read .claude/memory/ksbc-detail-cf.md before editing this file.

Why this script exists: after any shop-block reorder/insert/delete in a DETAIL
sheet, hardcoded CF row ranges go stale. Also, Excel CF solid fills silently
render nothing unless the color is set via bgColor (NOT fgColor) in the dxf.
This script re-parses the current layout and rewrites CF with the correct
DifferentialStyle shape.

Usage:
    python3 .claude/scripts/fix_detail_cf.py /path/to/workbook.xlsx

If no path arg is given, prompts stdin.
"""
import sys
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
from openpyxl.formatting.formatting import ConditionalFormattingList

BONDS = ['KOLLAM', 'KOZHIKODE', 'ATTINGAL', 'PALAKKAD', 'KOTTARAKARA', 'KANNUR',
         'ALAPPUZHA', 'NEDUMANGAD', 'ALUVA', 'PERINTHALMANNA', 'THODUPUZHA',
         'PATHANAMTHITTA', 'KOTTAYAM', 'THRISSUR', 'TRIPUNITHURA']

# Canonical palette (see ksbc-detail-cf.md)
DARK_GRAY = 'FF374151'
GOLD_ACCENT = 'FFFFD700'
TOTAL_RATING_BASE = 'FF4FC3F7'

RATING_FILLS = {
    '"🚀 High Performance"':  'FFBBDEFB',
    '"✅ Balanced"':           'FFDCEDC8',
    '"⚠️ Inventory Heavy"':   'FFFFE0B2',
    '"🚫 Critical Overstock"': 'FFFFCDD2',
}
TOTAL_RATING_FONT_COLORS = {
    '"🚀 High Performance"':  'FF64B5F6',
    '"✅ Balanced"':           'FF81C784',
    '"⚠️ Inventory Heavy"':   'FFFFB74D',
    '"🚫 Critical Overstock"': 'FFE57373',
}


def parse_block_rows(ws):
    """Find every shop block's brand-data / brand-total / pack-data / pack-total rows.

    Block start = a merged cell on col A whose value contains '—' and 'Staff:'.
    Brand TOTAL = first cell with value 'TOTAL' after the block start.
    Pack TOTAL = second 'TOTAL' after the 'PACK-WISE' label row.
    """
    starts = []
    for mr in list(ws.merged_cells.ranges):
        if mr.min_col != 1:
            continue
        v = ws.cell(mr.min_row, 1).value
        if v and isinstance(v, str) and '—' in v and 'Staff:' in v:
            starts.append(mr.min_row)
    starts.sort()

    blocks = []
    for i, s in enumerate(starts):
        next_s = starts[i + 1] if i + 1 < len(starts) else ws.max_row + 1
        brand_total = None
        pack_label = None
        pack_total = None
        for r in range(s + 1, next_s):
            v = ws.cell(r, 1).value
            if v == 'TOTAL':
                if brand_total is None:
                    brand_total = r
                else:
                    pack_total = r
                    break
            elif v == 'PACK-WISE':
                pack_label = r
        if brand_total is None:
            continue
        brand_data_start = s + 3
        brand_data_end = brand_total - 1
        if pack_label and pack_total:
            pack_data_start = pack_label + 2
            pack_data_end = pack_total - 1
        else:
            pack_data_start = pack_data_end = None
        blocks.append({
            'start': s,
            'brand_data_start': brand_data_start,
            'brand_data_end': brand_data_end,
            'brand_total': brand_total,
            'pack_data_start': pack_data_start,
            'pack_data_end': pack_data_end,
            'pack_total': pack_total,
        })
    return blocks


def fix_sheet(wb, bond):
    sheet_name = f'{bond} DETAIL'
    if sheet_name not in wb.sheetnames:
        return False
    ws = wb[sheet_name]

    blocks = parse_block_rows(ws)
    if not blocks:
        return False

    dark_gray_fill = PatternFill('solid', fgColor=DARK_GRAY)

    data_ranges = []
    total_ranges = []
    for b in blocks:
        data_ranges.append(f'G{b["brand_data_start"]}:G{b["brand_data_end"]}')
        if b['pack_data_start']:
            data_ranges.append(f'G{b["pack_data_start"]}:G{b["pack_data_end"]}')
        total_ranges.append(f'G{b["brand_total"]}')
        if b['pack_total']:
            total_ranges.append(f'G{b["pack_total"]}')

        # Canonical TOTAL row styling (not CF — direct on cell)
        for tr in [b['brand_total'], b['pack_total']]:
            if tr is None:
                continue
            cell = ws.cell(tr, 7)
            cell.fill = dark_gray_fill
            existing = cell.border
            cell.border = Border(
                top=Side(style='medium', color=GOLD_ACCENT),
                bottom=existing.bottom if existing and existing.bottom else None,
                left=existing.left if existing and existing.left else None,
                right=existing.right if existing and existing.right else None,
            )
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = Font(name='Trebuchet MS', size=10, bold=True, color=TOTAL_RATING_BASE)

    # Wipe existing CF on this sheet
    ws.conditional_formatting = ConditionalFormattingList()

    # Data-row fills — MUST use bgColor (Excel CF quirk; see ksbc-detail-cf.md)
    data_range_str = ' '.join(data_ranges)
    for priority, (txt, fill_color) in enumerate(RATING_FILLS.items(), start=1):
        dxf = DifferentialStyle(fill=PatternFill(bgColor=fill_color))
        rule = Rule(type='cellIs', operator='equal', formula=[txt],
                    stopIfTrue=False, dxf=dxf)
        rule.priority = priority
        ws.conditional_formatting.add(data_range_str, rule)

    # Total-row font colors
    total_range_str = ' '.join(total_ranges)
    for priority, (txt, font_color) in enumerate(TOTAL_RATING_FONT_COLORS.items(), start=5):
        dxf = DifferentialStyle(
            font=Font(name='Trebuchet MS', size=10, bold=True, color=font_color)
        )
        rule = Rule(type='cellIs', operator='equal', formula=[txt],
                    stopIfTrue=False, dxf=dxf)
        rule.priority = priority
        ws.conditional_formatting.add(total_range_str, rule)

    return True


def main():
    if len(sys.argv) < 2:
        print('Usage: python3 fix_detail_cf.py <workbook.xlsx>')
        sys.exit(1)
    target = sys.argv[1]
    wb = load_workbook(target, data_only=False)
    for bond in BONDS:
        ok = fix_sheet(wb, bond)
        mark = '✓' if ok else '·'
        print(f'  {mark} {bond} DETAIL')
    wb.save(target)
    print(f'\n✓ Saved {target}')


if __name__ == '__main__':
    main()
