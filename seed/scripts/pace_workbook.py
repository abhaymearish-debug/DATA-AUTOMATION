"""
Daily pace — the DAILY PACE TRACKER workbook.

KSD house style (navy / gold, Aptos Narrow). One target line per bond, split
flat across the month's selling days, with the catch-up number recomputed off
whatever has actually been sold.

The workbook is LIVE: each bond's MONTHLY TARGET is an editable yellow cell
and the daily grid, cumulatives, KPI tiles and the dashboard table are all
Excel formulas off it. Re-target a bond mid-month and everything reflows.

Sheets: DASHBOARD · DAILY GRID · 15 bond sheets · METHOD
"""

from __future__ import annotations

import datetime as _dt

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NAVY_DEEP, NAVY_MID, NAVY_SOFT = "FF0D1B4A", "FF1A237E", "FF263F80"
GOLD, GOLD_DIM, STEEL = "FFFFB300", "FFFFD54F", "FFB9C4E8"
GREY_DARK, WASH, EDIT, WHITE = "FF374151", "FFF5F7FA", "FFFFF3C4", "FFFFFFFF"
GREEN_F, GREEN_B = "FF2E7D32", "FFDCEDC8"
RED_F, RED_B = "FFC62828", "FFFFCDD2"
AMBER_F, AMBER_B = "FFE65100", "FFFFE0B2"
DRY_F, DRY_B = "FF6B7280", "FFECEFF1"

FONT = "Aptos Narrow"
THIN = Side(style="thin", color="FFD5DBE8")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
GOLD_RULE = Border(bottom=Side(style="medium", color=GOLD))

CS1 = '#,##0.0;-#,##0.0;"·"'
CS0 = '#,##0;-#,##0;"·"'
PCT = '0%'

FIRST = 10          # day 1 lands on this row, every day-grid sheet


def _f(sz=10, b=False, i=False, color="FF1F2937"):
    return Font(name=FONT, size=sz, bold=b, italic=i, color=color)


def _fill(c):
    return PatternFill("solid", fgColor=c)


def _c(ws, ref, val, font=None, fill=None, fmt=None, align="center",
       border=None, wrap=False, indent=0):
    cell = ws[ref]
    cell.value = val
    cell.font = font or _f()
    if fill:
        cell.fill = _fill(fill)
    if fmt:
        cell.number_format = fmt
    cell.alignment = Alignment(horizontal=align, vertical="center",
                               wrap_text=wrap, indent=indent)
    if border:
        cell.border = border
    return cell


def _band(ws, row, c1, c2, fill, height=None):
    """Merge FIRST, then style every cell -- openpyxl drops non-anchor fills."""
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    for c in range(c1, c2 + 1):
        ws.cell(row=row, column=c).fill = _fill(fill)
    if height:
        ws.row_dimensions[row].height = height


def _hero(ws, title, subtitle, ncols, note=None):
    _band(ws, 1, 1, ncols, NAVY_DEEP, 34)
    _c(ws, "A1", title, _f(18, True, color=WHITE), align="left", indent=1)
    _band(ws, 2, 1, ncols, NAVY_DEEP, 17)
    _c(ws, "A2", subtitle, _f(10, i=True, color=GOLD_DIM), align="left", indent=1)
    _band(ws, 3, 1, ncols, NAVY_DEEP, 16)
    if note:
        _c(ws, "A3", note, _f(9, color=STEEL), align="right", indent=1)
    for c in range(1, ncols + 1):
        ws.cell(row=3, column=c).border = GOLD_RULE
    ws.sheet_view.showGridLines = False


def _q(name):
    return f"'{name}'" if " " in name or "-" in name else name


def _status_cf(ws, rng):
    for txt, fc, bg in (("ON TARGET", GREEN_F, GREEN_B),
                        ("SHORT", RED_F, RED_B),
                        ("DRY", DRY_F, DRY_B)):
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="containsText", formula=[f'"{txt}"'],
                            fill=_fill(bg),
                            font=Font(name=FONT, size=10, bold=True, color=fc)))


def _day_grid(ws, plan, rows, target_ref, dry_ref_col=None):
    """Shared day table. target_ref = the cell holding this sheet's cs/day."""
    n = plan["ndays"]
    hdr = ["Day", "Wk", "Target/day", "Sold", "+/−", "Target to date",
           "Sold to date", "Behind by", "Status"]
    for i, h in enumerate(hdr, 1):
        _c(ws, f"{get_column_letter(i)}9", h, _f(10, True, color=WHITE),
           NAVY_SOFT, border=BOX, wrap=True)
    ws.row_dimensions[9].height = 30
    last = FIRST + n - 1
    for d in range(1, n + 1):
        r = FIRST + d - 1
        row = rows[d - 1]
        dry = row["dry"]
        _c(ws, f"A{r}", d, _f(10), border=BOX)
        _c(ws, f"B{r}", row["dow"], _f(10, color="FF6B7280"), border=BOX)
        _c(ws, f"C{r}", 0 if dry else f"={target_ref}", _f(10), fmt=CS1,
           border=BOX)
        act = row["actual"]
        _c(ws, f"D{r}", (round(act, 2) if act is not None else None),
           _f(10, True), fmt=CS1, border=BOX)
        _c(ws, f"E{r}", f'=IF(D{r}="","",D{r}-C{r})', _f(10), fmt=CS1, border=BOX)
        _c(ws, f"F{r}", f"=SUM($C${FIRST}:C{r})", _f(10, color="FF6B7280"),
           fmt=CS0, border=BOX)
        _c(ws, f"G{r}", f'=IF(D{r}="","",SUM($D${FIRST}:D{r}))', _f(10, True),
           fmt=CS0, border=BOX)
        _c(ws, f"H{r}", f'=IF(G{r}="","",G{r}-F{r})', _f(10, True), fmt=CS0,
           border=BOX)
        _c(ws, f"I{r}",
           f'=IF(C{r}=0,"⚪ DRY",IF(D{r}="","—",'
           f'IF(D{r}>=C{r},"🟢 ON TARGET","🔴 SHORT")))',
           _f(10, True), border=BOX)
        if dry:
            for col in "ABCDEFGH":
                ws[f"{col}{r}"].fill = _fill(DRY_B)
    tr = last + 1
    _band(ws, tr, 1, 2, GREY_DARK, 24)
    _c(ws, f"A{tr}", "TOTAL", _f(11, True, color=WHITE), align="left", indent=1)
    for col in "CDE":
        _c(ws, f"{col}{tr}", f"=SUM({col}{FIRST}:{col}{last})",
           _f(11, True, color=WHITE), GREY_DARK, fmt=CS0, border=BOX)
    for col in "FGHI":
        _c(ws, f"{col}{tr}", "", _f(11, True, color=WHITE), GREY_DARK, border=BOX)
    _status_cf(ws, f"I{FIRST}:I{last}")
    ws.conditional_formatting.add(
        f"H{FIRST}:H{last}",
        CellIsRule(operator="lessThan", formula=["0"],
                   font=Font(name=FONT, size=10, bold=True, color=RED_F)))
    ws.conditional_formatting.add(
        f"H{FIRST}:H{last}",
        CellIsRule(operator="greaterThanOrEqual", formula=["0"],
                   font=Font(name=FONT, size=10, bold=True, color=GREEN_F)))
    for col, w in zip("ABCDEFGHI", (6, 6, 12, 10, 10, 14, 13, 12, 15)):
        ws.column_dimensions[col].width = w
    return last, tr


def _bond_sheet(wb, plan, bond, targets):
    b = plan["bonds"][bond]
    ws = wb.create_sheet(bond[:31])
    _hero(ws, bond, f"Daily shop target · {plan['month'].title()} "
                    f"{plan['year']} · KSBC shop sales, all figures in cases", 9,
          f"data through {plan['as_of']:%d %b}")

    _c(ws, "A5", "MONTHLY TARGET", _f(9, True, color=NAVY_MID), align="left")
    ws.merge_cells("A5:B5")
    _c(ws, "C5", round(b["target_month"], 1), _f(12, True, color="FF7A5C00"),
       EDIT, fmt=CS0, border=BOX)
    _c(ws, "A6", "SELLING DAYS", _f(9, True, color=NAVY_MID), align="left")
    ws.merge_cells("A6:B6")
    _c(ws, "C6", plan["selling_days"], _f(11, color="FF6B7280"), WASH, fmt=CS0,
       border=BOX)
    _c(ws, "A7", "TARGET PER DAY", _f(9, True, color=NAVY_MID), align="left")
    ws.merge_cells("A7:B7")
    _c(ws, "C7", "=$C$5/$C$6", _f(14, True, color=NAVY_DEEP), WASH, fmt=CS1,
       border=BOX)
    _c(ws, "D5", "← edit the target; everything below reflows",
       _f(9, i=True, color="FF9AA3B2"), align="left")
    ws.merge_cells("D5:F5")

    last, tr = _day_grid(ws, plan, b["rows"], "$C$7")
    as_of_row = FIRST + plan["as_of"].day - 1

    kpis = [
        ("E6", "SOLD", "G6", f"=SUM(D{FIRST}:D{last})", CS0, GREEN_F),
        ("E7", "SHOULD BE AT", "G7", f"=F{as_of_row}", CS0, "FF6B7280"),
        ("H5", "DONE", "I5", f"=IFERROR(G6/$C$5,0)", PCT, NAVY_DEEP),
        ("H6", "BEHIND BY", "I6", f"=G6-G7", CS0, RED_F),
        ("H7", "NEED / DAY", "I7", f"=IFERROR(MAX($C$5-G6,0)/{max(b['days_left'],1)},0)",
         CS1, GOLD),
    ]
    for lab_ref, lab, val_ref, val, fmt, colr in kpis:
        _c(ws, lab_ref, lab, _f(9, True, color=NAVY_MID), align="left")
        _c(ws, val_ref, val, _f(13, True, color=colr), WASH, fmt=fmt, border=BOX)
    _c(ws, "F6", "", _f(), WASH, border=BOX)
    _c(ws, "F7", "", _f(), WASH, border=BOX)
    _c(ws, "D6", f"{b['days_left']} days left", _f(9, i=True, color="FF9AA3B2"),
       align="left")
    ws.merge_cells("D6:D7")

    shops = targets.get(bond, [])
    r = tr + 2
    _band(ws, r, 1, 9, NAVY_MID, 24)
    _c(ws, f"A{r}", "SALES PER SHOP — MONTH TO DATE",
       _f(12, True, color=GOLD_DIM), align="left", indent=1)
    r += 1
    ws.merge_cells(f"A{r}:I{r}")
    _c(ws, f"A{r}", f"Cumulative cases this month against the same number of "
                    f"selling days last month — like-for-like, not against "
                    f"last month's full total. KSBC shops only (FED and BAR "
                    f"are order-driven invoice dispatches).",
       _f(9, i=True, color="FF6B7280"), align="left", indent=1)
    r += 1
    hdr2 = ["#", "Shop", "This month", "Same days last month", "Change",
            "Cases per selling day"]
    for i, h in enumerate(hdr2, 1):
        _c(ws, f"{get_column_letter(i)}{r}", h, _f(10, True, color=WHITE),
           NAVY_SOFT, border=BOX, wrap=True)
    ws.merge_cells(f"F{r}:I{r}")
    for cc in range(6, 10):
        ws.cell(row=r, column=cc).fill = _fill(NAVY_SOFT)
        ws.cell(row=r, column=cc).border = BOX
    ws.row_dimensions[r].height = 24
    first_shop = r + 1
    for i, x in enumerate(shops, 1):
        r += 1
        _c(ws, f"A{r}", i, _f(10), border=BOX)
        _c(ws, f"B{r}", x["shop"], _f(10, True), align="left", border=BOX,
           indent=1)
        _c(ws, f"C{r}", round(x["now"], 1), _f(10, True, color="FF7A5C00"),
           "FFFFF9EC", fmt=CS1, border=BOX)
        _c(ws, f"D{r}", round(x["last"], 1), _f(10, color="FF6B7280"),
           fmt=CS1, border=BOX)
        _c(ws, f"E{r}", f"=C{r}-D{r}",
           _f(10, True, color=GREEN_F if x["delta"] >= 0 else RED_F),
           fmt='+#,##0.0;-#,##0.0;"·"', border=BOX)
        ws.merge_cells(f"F{r}:I{r}")
        _c(ws, f"F{r}", round(x["per_day"], 2), _f(10),
           fmt='#,##0.00;-#,##0.00;"·"', border=BOX, align="left", indent=1)
        for cc in range(7, 10):
            ws.cell(row=r, column=cc).border = BOX
        ws.row_dimensions[r].height = 17
    if shops:
        r += 1
        _band(ws, r, 1, 2, GREY_DARK, 22)
        _c(ws, f"A{r}", "BOND", _f(11, True, color=WHITE), align="left",
           indent=1)
        for col in "CDF":
            _c(ws, f"{col}{r}", f"=SUM({col}{first_shop}:{col}{r-1})",
               _f(11, True, color=WHITE), GREY_DARK,
               fmt=CS1 if col != "F" else '#,##0.00', border=BOX)
        _c(ws, f"E{r}", f"=C{r}-D{r}", _f(11, True, color=WHITE), GREY_DARK,
           fmt='+#,##0.0;-#,##0.0;"·"', border=BOX)
        ws.merge_cells(f"F{r}:I{r}")
        for cc in range(7, 10):
            ws.cell(row=r, column=cc).fill = _fill(GREY_DARK)
            ws.cell(row=r, column=cc).border = BOX

    ws.column_dimensions["B"].width = 24
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def _daily_grid(wb, plan):
    ws = wb.create_sheet("DAILY GRID")
    _hero(ws, "DAILY GRID — NETWORK",
          f"Every day of {plan['month'].title()} {plan['year']} · "
          f"target vs what actually sold", 9,
          f"data through {plan['as_of']:%d %b}")
    bonds = [_q(b) for b in plan["order"]]
    n = plan["ndays"]
    hdr = ["Day", "Wk", "Target/day", "Sold", "+/−", "Target to date",
           "Sold to date", "Behind by", "Status"]
    for i, h in enumerate(hdr, 1):
        _c(ws, f"{get_column_letter(i)}9", h, _f(10, True, color=WHITE),
           NAVY_SOFT, border=BOX, wrap=True)
    ws.row_dimensions[9].height = 30
    last = FIRST + n - 1
    for d in range(1, n + 1):
        r = FIRST + d - 1
        row = plan["network"]["rows"][d - 1]
        _c(ws, f"A{r}", d, _f(10), border=BOX)
        _c(ws, f"B{r}", row["dow"], _f(10, color="FF6B7280"), border=BOX)
        for col in ("C", "D"):
            _c(ws, f"{col}{r}", "=" + "+".join(f"{b}!{col}{r}" for b in bonds),
               _f(10, True if col == "D" else False), fmt=CS1, border=BOX)
        _c(ws, f"E{r}", f'=IF(D{r}=0,"",D{r}-C{r})', _f(10), fmt=CS1, border=BOX)
        _c(ws, f"F{r}", f"=SUM($C${FIRST}:C{r})", _f(10, color="FF6B7280"),
           fmt=CS0, border=BOX)
        _c(ws, f"G{r}", f'=IF(D{r}=0,"",SUM($D${FIRST}:D{r}))', _f(10, True),
           fmt=CS0, border=BOX)
        _c(ws, f"H{r}", f'=IF(G{r}="","",G{r}-F{r})', _f(10, True), fmt=CS0,
           border=BOX)
        _c(ws, f"I{r}", f'=IF(C{r}=0,"⚪ DRY",IF(D{r}=0,"—",'
                        f'IF(D{r}>=C{r},"🟢 ON TARGET","🔴 SHORT")))',
           _f(10, True), border=BOX)
        if row["dry"]:
            for col in "ABCDEFGH":
                ws[f"{col}{r}"].fill = _fill(DRY_B)
    tr = last + 1
    _band(ws, tr, 1, 2, GREY_DARK, 24)
    _c(ws, f"A{tr}", "TOTAL", _f(11, True, color=WHITE), align="left", indent=1)
    for col in "CDE":
        _c(ws, f"{col}{tr}", f"=SUM({col}{FIRST}:{col}{last})",
           _f(11, True, color=WHITE), GREY_DARK, fmt=CS0, border=BOX)
    for col in "FGHI":
        _c(ws, f"{col}{tr}", "", _f(11, True, color=WHITE), GREY_DARK, border=BOX)
    _status_cf(ws, f"I{FIRST}:I{last}")
    for col, w in zip("ABCDEFGHI", (6, 6, 12, 10, 10, 14, 13, 12, 15)):
        ws.column_dimensions[col].width = w

    ch = LineChart()
    ch.title = "Cumulative — target vs sold"
    ch.height, ch.width = 8.2, 20
    ch.y_axis.title = "cases"
    ch.add_data(Reference(ws, min_col=6, max_col=7, min_row=9, max_row=last),
                titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=FIRST, max_row=last))
    for s, colr, dash in zip(ch.series, (GOLD, "FF1565C0"), ("dash", None)):
        s.graphicalProperties.line.solidFill = colr
        s.graphicalProperties.line.width = 24000
        if dash:
            s.graphicalProperties.line.dashStyle = dash
        s.smooth = False
    ws.add_chart(ch, f"K{FIRST}")
    ws.freeze_panes = f"A{FIRST}"
    ws.page_setup.orientation = "landscape"
    return ws


def _dashboard(wb, plan, targets):
    ws = wb.create_sheet("DASHBOARD", 0)
    n = plan["network"]
    nxt = next((b["next_date"] for b in plan["bonds"].values()
                if b["next_date"]), None)
    _hero(ws, "DAILY PACE TRACKER",
          f"{plan['month'].title()} {plan['year']} · KSBC shop liquidation "
          f"· all figures in cases", 10,
          f"data through {plan['as_of']:%d %b}" +
          (f" · for {nxt:%A %d %b}" if nxt else ""))

    first, lastr = 12, 26
    tiles = [
        ("MONTHLY TARGET", f"=SUM(C{first}:C{lastr})", CS0, WHITE),
        ("SOLD SO FAR", f"=SUM(E{first}:E{lastr})", CS0, GOLD),
        ("% DONE", f"=IFERROR(SUM(E{first}:E{lastr})/SUM(C{first}:C{lastr}),0)",
         PCT, "FF81C784"),
        ("BEHIND BY", f"=SUM(G{first}:G{lastr})", CS0, "FFE57373"),
        ("DAYS LEFT", n["days_left"], CS0, STEEL),
    ]
    col = 1
    for lab, val, fmt, accent in tiles:
        _band(ws, 5, col, col + 1, NAVY_MID, 30)
        _band(ws, 6, col, col + 1, NAVY_MID, 28)
        L = get_column_letter(col)
        _c(ws, f"{L}5", lab, _f(9, True, color=STEEL))
        _c(ws, f"{L}6", val, _f(17, True, color=accent), fmt=fmt)
        col += 2
    _band(ws, 7, 1, 10, NAVY_DEEP, 26)
    _c(ws, "A7", f"NEEDS {n['needed_per_day']:,.0f} CASES A DAY "
                 f"FOR THE LAST {n['days_left']} DAYS",
       _f(13, True, color=GOLD_DIM))

    r = 10
    _band(ws, r, 1, 10, NAVY_MID, 24)
    _c(ws, f"A{r}", "BOND TABLE — ranked by % of target achieved",
       _f(12, True, color=GOLD_DIM), align="left", indent=1)
    r = 11
    hdr = ["#", "Bond", "Target", "Per day", "Sold", "Should be at",
           "Behind by", "% done", "NEED / DAY", "Streak"]
    for i, h in enumerate(hdr, 1):
        _c(ws, f"{get_column_letter(i)}{r}", h, _f(10, True, color=WHITE),
           NAVY_SOFT, border=BOX, wrap=True)
    ws.row_dimensions[r].height = 30

    for i, bond in enumerate(plan["order"]):
        b = plan["bonds"][bond]
        rr = first + i
        q = _q(bond)
        _c(ws, f"A{rr}", i + 1, _f(10), border=BOX)
        _c(ws, f"B{rr}", bond, _f(10, True), align="left", border=BOX, indent=1)
        _c(ws, f"C{rr}", f"={q}!$C$5", _f(10), fmt=CS0, border=BOX)
        _c(ws, f"D{rr}", f"={q}!$C$7", _f(10, color="FF6B7280"), fmt=CS1,
           border=BOX)
        _c(ws, f"E{rr}", f"={q}!$G$6", _f(10, True), fmt=CS0, border=BOX)
        _c(ws, f"F{rr}", f"={q}!$G$7", _f(10, color="FF6B7280"), fmt=CS0,
           border=BOX)
        _c(ws, f"G{rr}", f"={q}!$I$6", _f(10, True), fmt=CS0, border=BOX)
        _c(ws, f"H{rr}", f"={q}!$I$5", _f(10, True), fmt=PCT, border=BOX)
        _c(ws, f"I{rr}", f"={q}!$I$7", _f(11, True, color="FF7A5C00"), EDIT,
           fmt=CS1, border=BOX)
        _c(ws, f"J{rr}", b["streak"], _f(10), border=BOX)

    tr = lastr + 1
    _band(ws, tr, 1, 2, GREY_DARK, 24)
    _c(ws, f"A{tr}", "TOTAL", _f(11, True, color=WHITE), align="left", indent=1)
    for col in "CDEFGIJ":
        _c(ws, f"{col}{tr}", f"=SUM({col}{first}:{col}{lastr})",
           _f(11, True, color=WHITE), GREY_DARK,
           fmt=CS1 if col in "DI" else CS0, border=BOX)
    _c(ws, f"H{tr}", f"=IFERROR(E{tr}/C{tr},0)", _f(11, True, color=WHITE),
       GREY_DARK, fmt=PCT, border=BOX)

    ws.conditional_formatting.add(
        f"G{first}:G{lastr}",
        CellIsRule(operator="lessThan", formula=["0"], fill=_fill(RED_B),
                   font=Font(name=FONT, size=10, bold=True, color=RED_F)))
    ws.conditional_formatting.add(
        f"G{first}:G{lastr}",
        CellIsRule(operator="greaterThanOrEqual", formula=["0"],
                   fill=_fill(GREEN_B),
                   font=Font(name=FONT, size=10, bold=True, color=GREEN_F)))

    r = tr + 2
    _band(ws, r, 1, 10, NAVY_MID, 24)
    _c(ws, f"A{r}", "HOW IT WORKS", _f(12, True, color=GOLD_DIM), align="left",
       indent=1)
    notes = [
        f"TARGET PER DAY = monthly shop target ÷ {plan['selling_days']} selling "
        f"days. That is the one number the executive is held to.",
        "NEED / DAY is the catch-up number: (monthly target − sold so far) ÷ days "
        "left. Fall short today and it goes up tomorrow by itself — which is the "
        "point of running a daily target instead of a monthly one.",
        f"Dry days carry no target. The 1st is dry every month in Kerala, and any "
        f"other day that comes in network-zero is detected automatically and its "
        f"share spread over the days that remain. Dry so far this month: "
        f"{', '.join(str(d) for d in plan['dry_days'])}.",
        f"Basis is KSBC SHOP liquidation only. FED and BAR invoice still counts "
        f"as liquidation but is deliberately NOT in this target (Abhay, 27 Jul "
        f"2026) — invoice arrives as an event rather than a daily flow, and "
        f"executives were leaning on it instead of working the shops. "
        f"Month-to-date is reconciled to the KSBC BOND PERFORMANCE sheet.",
        f"Monthly shop target = 80% of the official bond target. The official "
        f"target covers TOTAL liquidation (shop + FED/BAR invoice); 80% of it "
        f"is treated as the shop leg. Each bond's target is an editable yellow "
        f"cell — change it and the whole grid recomputes.",
        "Every bond sheet's MONTHLY TARGET is an editable yellow cell. Change it "
        "and that bond's grid, this table and the tiles all recompute.",
    ]
    for note in notes:
        r += 1
        ws.merge_cells(f"A{r}:J{r}")
        _c(ws, f"A{r}", "•  " + note, _f(10), align="left", wrap=True, indent=1)
        ws.row_dimensions[r].height = 30

    for col, w in zip("ABCDEFGHIJ", (5, 18, 10, 10, 10, 13, 11, 9, 12, 8)):
        ws.column_dimensions[col].width = w
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def build(plan, targets, out_path: str) -> str:
    wb = Workbook()
    wb.remove(wb.active)
    for bond in plan["order"]:
        _bond_sheet(wb, plan, bond, targets)
    _daily_grid(wb, plan)
    _dashboard(wb, plan, targets)
    wb._sheets = [wb[s] for s in (["DASHBOARD", "DAILY GRID"] + list(plan["order"]))
                  if s in wb.sheetnames]
    wb.active = 0
    wb.save(out_path)
    return out_path
