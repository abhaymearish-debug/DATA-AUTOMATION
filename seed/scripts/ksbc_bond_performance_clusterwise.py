"""
KSBC BOND PERFORMANCE sheet — clusterwise rebuild from scratch.

Locked layout (25 Apr 2026):
  r1: title band, merged A1:H1
  r2: column headers (Bond | Opening | Receipts | Sales | Closing |
                       Sell-Through % | Closing vs Sales % | Rating)
  r3:  Cluster 1 banner — navy `FF263F80`, SUM(B4:B9..E4:E9), tier-color font on H
  r4-r9: Cluster 1 bonds (6) sorted by rating tier then sell-through % desc
  r10: Cluster 2 banner — SUM(B11:B15..E11:E15)
  r11-r15: Cluster 2 bonds (5) sorted
  r16: Cluster 3 banner — SUM(B17:B20..E17:E20)
  r17-r20: Cluster 3 bonds (4) sorted
  r21: TOTAL — dark grey `FF374151`, white bold, =SUM across all 15 bond rows,
       bright tier-color font on H, gold medium top border

Cluster membership (locked 25 Apr 2026 per CLAUDE.md):
  Cluster 1 = ALAPPUZHA, ATTINGAL, NEDUMANGAD, KOLLAM, KOTTARAKARA, PATHANAMTHITTA
  Cluster 2 = THRISSUR, TRIPUNITHURA, THODUPUZHA, KOTTAYAM, ALUVA
  Cluster 3 = KANNUR, KOZHIKODE, PALAKKAD, PERINTHALMANNA

This script REBUILDS the sheet from scratch. It does NOT do per-row CF rule
application — that's owned by fix_region_sheet_styling.py downstream. It
DOES set direct cell fills + fonts (defense-in-depth that survives even
when CF is suppressed). The downstream fix_rating_cell_format.py will
overlay the bright tier-color font on cluster banners + TOTAL.

Usage:
    python3 ksbc_bond_performance_clusterwise.py <workbook.xlsx>

The script reads each of the 15 region sheets, computes its TOTAL row,
and uses those as the per-bond aggregates. The 15 region sheets must
exist and have a TOTAL row at the bottom with valid Opening/Receipts/
Sales/Closing values in cols D-G.
"""
from __future__ import annotations
import sys
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# Cluster membership (locked 25 Apr 2026)
CLUSTERS = [
    ('Cluster 1', ['ALAPPUZHA', 'ATTINGAL', 'NEDUMANGAD', 'KOLLAM',
                   'KOTTARAKARA', 'PATHANAMTHITTA']),
    ('Cluster 2', ['THRISSUR', 'TRIPUNITHURA', 'THODUPUZHA', 'KOTTAYAM', 'ALUVA']),
    ('Cluster 3', ['KANNUR', 'KOZHIKODE', 'PALAKKAD', 'PERINTHALMANNA']),
]

FONT_NAME = 'Trebuchet MS'

# Palette
NAVY_TITLE   = 'FF1A237E'   # title row text
NAVY_HEADER  = 'FF1A237E'   # column-header row fill
NAVY_BANNER  = 'FF263F80'   # cluster banner fill (slightly lighter)
WHITE        = 'FFFFFFFF'
BLACK        = 'FF000000'
LIGHT_BG     = 'FFF8FAFC'
ALT_ROW      = 'FFF1F5F9'   # alternating row fill
DARK_GRAY    = 'FF374151'   # TOTAL fill
GOLD         = 'FFFFD700'   # accent border colour
THIN_GREY    = 'FFE5E7EB'   # thin row separator

# Bright tier fonts for use on dark (banner / TOTAL) backgrounds
BRIGHT_FONT = {
    'high':     'FF4FC3F7',
    'balanced': 'FF81C784',
    'inv':      'FFFFB74D',
    'crit':     'FFE57373',
    'none':     'FFB0BEC5',
}

# Row heights
H_TITLE   = 36
H_HEADER  = 28
H_BANNER  = 32
H_BOND    = 24
H_TOTAL   = 30


def _tier(opening, receipts, sales, closing):
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


def _read_region_totals(wb) -> dict[str, dict]:
    """For each region sheet, read its TOTAL row's Opening/Receipts/Sales/Closing.
    Region layout: r4 col-header, data r5..total-1, TOTAL at bottom.
    Region columns: A=code, B=name, C=staff, D=Opening, E=Receipts, F=Sales, G=Closing.
    """
    all_bonds = [b for _, bonds in CLUSTERS for b in bonds]
    out = {}
    for bond in all_bonds:
        if bond not in wb.sheetnames:
            out[bond] = {'Opening': 0, 'Receipts': 0, 'Sales': 0, 'Closing': 0}
            continue
        ws = wb[bond]
        total_r = None
        for r in range(5, ws.max_row + 1):
            if ws.cell(r, 1).value == 'TOTAL':
                total_r = r
                break
        if total_r is None:
            out[bond] = {'Opening': 0, 'Receipts': 0, 'Sales': 0, 'Closing': 0}
            continue

        def _resolve(v):
            if isinstance(v, (int, float)):
                return float(v)
            if isinstance(v, str) and v.startswith('=SUM(') and v.endswith(')'):
                # Robust SUM parser (fix 1 Jun 2026): handles single-range
                # `=SUM(B4:B9)`, multi-range `=SUM(B4:B9,B11:B15)`, and a list
                # of single cells `=SUM(B4,B5,...)`. The old version only matched
                # a single range and silently returned 0.0 for anything else —
                # a silent-zero trap that would drop a bond's total.
                inner = v[5:-1].replace(' ', '')
                total = 0.0
                parsed_any = False
                for part in inner.split(','):
                    rng = re.match(r'^([A-Z]+)(\d+):([A-Z]+)(\d+)$', part)
                    one = re.match(r'^([A-Z]+)(\d+)$', part)
                    if rng and rng.group(1) == rng.group(3):
                        col = ord(rng.group(1)[-1]) - ord('A') + 1 if len(rng.group(1)) == 1 else None
                        if col is None:
                            sys.stderr.write(
                                f"⚠ _read_region_totals: multi-letter column in SUM "
                                f"term '{part}' of '{v}' — skipped (contributes 0); "
                                f"BOND PERFORMANCE totals may be understated.\n")
                            continue
                        for rr in range(int(rng.group(2)), int(rng.group(4)) + 1):
                            vv = ws.cell(rr, col).value
                            if isinstance(vv, (int, float)):
                                total += vv
                        parsed_any = True
                    elif one and len(one.group(1)) == 1:
                        col = ord(one.group(1)) - ord('A') + 1
                        vv = ws.cell(int(one.group(2)), col).value
                        if isinstance(vv, (int, float)):
                            total += vv
                        parsed_any = True
                    else:
                        sys.stderr.write(
                            f"⚠ _resolve: could not parse SUM term '{part}' in "
                            f"'{v}' — that term contributes 0 (check BOND "
                            f"PERFORMANCE formula layout).\n")
                if parsed_any:
                    return total
            return 0.0

        out[bond] = {
            'Opening':  _resolve(ws.cell(total_r, 4).value),
            'Receipts': _resolve(ws.cell(total_r, 5).value),
            'Sales':    _resolve(ws.cell(total_r, 6).value),
            'Closing':  _resolve(ws.cell(total_r, 7).value),
        }
    return out


def _sort_cluster_bonds(bonds: list[str], region_totals: dict[str, dict]) -> list[str]:
    """Sort bonds within a cluster by rating tier (high→balanced→inv→crit→none),
    then by sell-through % descending within tier."""
    tier_order = {'high': 0, 'balanced': 1, 'inv': 2, 'crit': 3, 'none': 4}
    def key(bond):
        t = region_totals.get(bond, {})
        o, r, s, c = t.get('Opening', 0), t.get('Receipts', 0), t.get('Sales', 0), t.get('Closing', 0)
        denom = o + r
        st = (s / denom) if denom else 0.0
        # Bond name as a final deterministic tie-break (fix 1 Jun 2026) so two
        # bonds at the same tier + sell-through don't reshuffle build-to-build.
        return (tier_order[_tier(o, r, s, c)], -st, str(bond))
    return sorted(bonds, key=key)


def rebuild_bond_performance(wb, month_u: str, latest_day: int):
    """Rebuild BOND PERFORMANCE from scratch using region-sheet TOTAL rows.

    `month_u` is the uppercase month ("MAY"), `latest_day` is the highest
    day covered (e.g. 9 for "MAY 1st - 9th"). These go into the title band.
    """
    region_totals = _read_region_totals(wb)

    if 'BOND PERFORMANCE' in wb.sheetnames:
        del wb['BOND PERFORMANCE']
    ws = wb.create_sheet('BOND PERFORMANCE')
    # Position the sheet right after DASHBOARD if present, else first.
    if 'DASHBOARD' in wb.sheetnames:
        sheets = wb.sheetnames
        dash_idx = sheets.index('DASHBOARD')
        new_idx = sheets.index('BOND PERFORMANCE')
        wb.move_sheet(ws, offset=(dash_idx + 1) - new_idx)

    # Column widths
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 11
    for col in ('C', 'D', 'E'):
        ws.column_dimensions[col].width = 11
    ws.column_dimensions['F'].width = 16
    ws.column_dimensions['G'].width = 18
    ws.column_dimensions['H'].width = 26

    # ---- r1: title band ----
    ws.row_dimensions[1].height = H_TITLE
    ws.cell(1, 1).value = f'KSBC  BOND-WISE  PERFORMANCE  —  {month_u} 1–{latest_day}'
    ws.merge_cells('A1:H1')
    c = ws.cell(1, 1)
    c.font = Font(name=FONT_NAME, size=14, bold=True, color=NAVY_TITLE)
    c.fill = PatternFill('solid', fgColor=LIGHT_BG)
    c.alignment = Alignment(horizontal='center', vertical='center')

    # ---- r2: column header ----
    ws.row_dimensions[2].height = H_HEADER
    headers = ['Bond', 'Opening', 'Receipts', 'Sales', 'Closing',
               'Sell-Through %', 'Closing vs Sales %', 'Rating']
    thin = Side(style='thin', color=THIN_GREY)
    for i, h in enumerate(headers, start=1):
        c = ws.cell(2, i)
        c.value = h
        c.font = Font(name=FONT_NAME, size=11, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=NAVY_HEADER)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = Border(top=thin, bottom=thin)

    # ---- Cluster banners + bond rows ----
    r = 3
    rating_formula = ('=IF(F{r}>=0.8,"🚀 High Performance",'
                      'IF(F{r}>=0.6,"✅ Balanced",'
                      'IF(F{r}>=0.4,"⚠️ Inventory Heavy",'
                      'IF(F{r}>=0,"🚫 Critical Overstock","— No activity"))))')
    pct_f = '=IFERROR(D{r}/(B{r}+C{r}),0)'
    closing_vs_sales_f = '=IFERROR(E{r}/D{r},0)'

    for cluster_label, bonds in CLUSTERS:
        banner_r = r
        first_bond_r = r + 1
        last_bond_r = first_bond_r + len(bonds) - 1

        # Banner row
        ws.row_dimensions[banner_r].height = H_BANNER
        ws.cell(banner_r, 1).value = cluster_label
        ws.cell(banner_r, 2).value = f'=SUM(B{first_bond_r}:B{last_bond_r})'
        ws.cell(banner_r, 3).value = f'=SUM(C{first_bond_r}:C{last_bond_r})'
        ws.cell(banner_r, 4).value = f'=SUM(D{first_bond_r}:D{last_bond_r})'
        ws.cell(banner_r, 5).value = f'=SUM(E{first_bond_r}:E{last_bond_r})'
        ws.cell(banner_r, 6).value = pct_f.format(r=banner_r)
        ws.cell(banner_r, 7).value = closing_vs_sales_f.format(r=banner_r)
        ws.cell(banner_r, 8).value = rating_formula.format(r=banner_r)

        # Compute cluster tier to set bright font on col H
        c_op = sum(region_totals[b]['Opening']  for b in bonds)
        c_rc = sum(region_totals[b]['Receipts'] for b in bonds)
        c_sa = sum(region_totals[b]['Sales']    for b in bonds)
        c_cl = sum(region_totals[b]['Closing']  for b in bonds)
        cluster_tier = _tier(c_op, c_rc, c_sa, c_cl)
        banner_h_font = BRIGHT_FONT[cluster_tier]

        for col in range(1, 9):
            cell = ws.cell(banner_r, col)
            cell.fill = PatternFill('solid', fgColor=NAVY_BANNER)
            font_color = banner_h_font if col == 8 else WHITE
            cell.font = Font(name=FONT_NAME, size=11, bold=True, color=font_color)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if col in (2, 3, 4, 5):
                cell.number_format = '#,##0'
            elif col in (6, 7):
                cell.number_format = '0%'

        # Bond rows — sorted by tier then ST% desc within cluster
        sorted_bonds = _sort_cluster_bonds(bonds, region_totals)
        for i, bond in enumerate(sorted_bonds):
            br = first_bond_r + i
            t = region_totals[bond]
            ws.row_dimensions[br].height = H_BOND
            row_fill = WHITE if (i % 2 == 0) else ALT_ROW
            ws.cell(br, 1).value = bond
            ws.cell(br, 2).value = t['Opening']
            ws.cell(br, 3).value = t['Receipts']
            ws.cell(br, 4).value = t['Sales']
            ws.cell(br, 5).value = t['Closing']
            ws.cell(br, 6).value = pct_f.format(r=br)
            ws.cell(br, 7).value = closing_vs_sales_f.format(r=br)
            ws.cell(br, 8).value = rating_formula.format(r=br)
            for col in range(1, 9):
                cell = ws.cell(br, col)
                cell.fill = PatternFill('solid', fgColor=row_fill)
                bold = (col in (1, 8))  # bond name + rating are bold
                cell.font = Font(name=FONT_NAME, size=10, bold=bold, color=BLACK)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = Border(top=thin, bottom=thin)
                if col in (2, 3, 4, 5):
                    cell.number_format = '#,##0'
                elif col in (6, 7):
                    cell.number_format = '0%'

        r = last_bond_r + 1

    # ---- TOTAL row ----
    total_r = r
    ws.row_dimensions[total_r].height = H_TOTAL
    # Build SUM across all bond rows, skipping cluster banners.
    # Cluster banner rows are: 3, (3 + 1 + 6) = 10, (10 + 1 + 5) = 16.
    # Bond ranges: B4:B9, B11:B15, B17:B20.
    ranges = []
    cur = 3
    for _, bonds in CLUSTERS:
        first = cur + 1
        last = first + len(bonds) - 1
        ranges.append(f'{{c}}{first}:{{c}}{last}')
        cur = last + 1
    sum_template = ','.join(ranges)

    ws.cell(total_r, 1).value = 'TOTAL'
    ws.cell(total_r, 2).value = '=SUM(' + sum_template.replace('{c}', 'B') + ')'
    ws.cell(total_r, 3).value = '=SUM(' + sum_template.replace('{c}', 'C') + ')'
    ws.cell(total_r, 4).value = '=SUM(' + sum_template.replace('{c}', 'D') + ')'
    ws.cell(total_r, 5).value = '=SUM(' + sum_template.replace('{c}', 'E') + ')'
    ws.cell(total_r, 6).value = pct_f.format(r=total_r)
    ws.cell(total_r, 7).value = closing_vs_sales_f.format(r=total_r)
    ws.cell(total_r, 8).value = rating_formula.format(r=total_r)

    # Compute overall tier for the bright font on col H
    tot_op = sum(t['Opening']  for t in region_totals.values())
    tot_rc = sum(t['Receipts'] for t in region_totals.values())
    tot_sa = sum(t['Sales']    for t in region_totals.values())
    tot_cl = sum(t['Closing']  for t in region_totals.values())
    total_tier = _tier(tot_op, tot_rc, tot_sa, tot_cl)
    total_h_font = BRIGHT_FONT[total_tier]

    gold_top = Side(style='medium', color=GOLD)
    for col in range(1, 9):
        cell = ws.cell(total_r, col)
        cell.fill = PatternFill('solid', fgColor=DARK_GRAY)
        font_color = total_h_font if col == 8 else WHITE
        cell.font = Font(name=FONT_NAME, size=10, bold=True, color=font_color)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(top=gold_top)
        if col in (2, 3, 4, 5):
            cell.number_format = '#,##0'
        elif col in (6, 7):
            cell.number_format = '0%'

    return ws


def main():
    if len(sys.argv) < 2:
        print('Usage: python3 ksbc_bond_performance_clusterwise.py <workbook.xlsx>')
        sys.exit(1)
    target = sys.argv[1]
    wb = load_workbook(target, data_only=False)
    # Infer month + latest_day from any sheet named "<MONTH> N" or COMBINED.
    month_u = None
    latest_day = 0
    for sn in wb.sheetnames:
        m = re.match(r'^([A-Z]+)\s+(\d+)(?:-(\d+))?(?:\s+COMBINED)?$', sn)
        if m:
            month_u = m.group(1)
            day_end = int(m.group(3) or m.group(2))
            if day_end > latest_day:
                latest_day = day_end
        m2 = re.match(r'^([A-Z]+)\s+1-(\d+)\s+COMBINED$', sn)
        if m2:
            month_u = m2.group(1)
            if int(m2.group(2)) > latest_day:
                latest_day = int(m2.group(2))
    if not month_u:
        month_u = 'MONTH'
    rebuild_bond_performance(wb, month_u, latest_day)
    wb.save(target)
    print(f'BOND PERFORMANCE rebuilt — {month_u} 1-{latest_day}, saved {target}.')


if __name__ == '__main__':
    main()
