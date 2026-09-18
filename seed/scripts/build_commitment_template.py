#!/usr/bin/env python3
"""
Build the commitment-tracker INPUT workbook for a single bond.

Mirrors the Brand × Pack × Shop rows shown on the signed Aging Stock Liquidation
Review PDF. Two cells per row are intentionally left empty for capture:
  - Target Cases (commit)  ← number the ASM/Sales Executive wrote in the PDF
  - Remarks                ← context they wrote
Vision OCR (or manual data-entry) will populate these two columns.

Four columns are placeholders for the nightly tracker (Stage 2) to fill:
  - Cases Sold Since Aging Snapshot
  - % Achieved
  - Days Remaining
  - Status

Layout: one header band, one column-header row, alternating shop blocks
(banner row + brand×pack rows). Compatible with openpyxl recalc.
"""

from __future__ import annotations

import argparse
import os
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule

# ---------------------- palette (matches KSD docs) ----------------------
NAVY_DEEP  = "FF0D1B4A"
NAVY_MID   = "FF1A237E"
NAVY_SOFT  = "FF334466"
GOLD       = "FFFFB300"
GOLD_DIM   = "FFFFD54F"

NM_DARK    = "FFB71C1C"
NM_LIGHT   = "FFFFEBEE"
CR_DARK    = "FFE65100"
CR_LIGHT   = "FFFFF3E0"
SL_DARK    = "FFF9A825"
SL_LIGHT   = "FFFFFDE7"

INPUT_FILL = "FFFFF59D"     # bright yellow — ASM/OCR fills these
REF_FILL   = "FFF5F7FB"     # light grey — reference data
OUT_FILL   = "FFE3F2FD"     # pale blue — tracker computes
GREEN_DARK = "FF2E7D32"
GREEN_LIGHT= "FFDCEDC8"

THIN  = Side(style="thin",  color="FFD0D7DE")
MED   = Side(style="medium",color="FFFFB300")


# ---------------------- columns ----------------------
# Layout (15 cols). REF = reference (locked, pre-filled at build time).
# INPUT = yellow (Abhay fills). OUTPUT = blue (refresh script writes).
# FORMULA = blue (auto-calculated by Excel).
#   A  #                                    REF
#   B  Shop Code                            REF
#   C  Shop Name                            REF
#   D  Severity                             REF
#   E  Brand                                REF
#   F  Pack                                 REF
#   G  May Closing  (at meeting, reference) REF
#   H  Latest Closing                       OUTPUT  ← refreshed from KSBC COMBINED
#   I  Months Cover                         REF
#   J  Zero Months                          REF
#   K  Target Cases (commit)                INPUT   ← yellow, ASM commit
#   L  Cases Sold since aging snapshot       OUTPUT  ← refresh writes
#   M  % achieved                           FORMULA = L / K
#   N  Days remaining                       OUTPUT  ← refresh writes
#   O  Status                               FORMULA tier
COLS = [
    ("#",                          5,   "center"),
    ("Shop Code",                  11,  "center"),
    ("Shop Name",                  24,  "left"),
    ("Severity",                   13,  "center"),
    ("Brand",                      30,  "left"),
    ("Pack",                       9,   "center"),
    ("May Closing\n(as of snapshot)", 13, "center"),
    ("Latest\nClosing",            11,  "center"),
    ("Months\nCover",              9,   "center"),
    ("Zero\nMonths",               8,   "center"),
    ("Target Cases\n(commit) ← FILL", 16, "center"),
    ("Cases Sold\nsince snapshot", 13,  "center"),
    ("%\nachieved",                10,  "center"),
    ("Days\nremaining",            10,  "center"),
    ("Status",                     17,  "center"),
]


# ---------------------- parse the aging workbook ----------------------
def parse_bond(xlsx_path, bond):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if bond not in wb.sheetnames:
        raise SystemExit(f"Bond '{bond}' not in workbook")
    ws = wb[bond]

    field_staff = ""
    period = ""
    line2 = ws.cell(row=2, column=1).value or ""
    for part in str(line2).split("·"):
        part = part.strip()
        if part.lower().startswith("field staff:"):
            field_staff = part.split(":", 1)[1].strip()
        elif "as of" in part.lower():
            period = part.strip()

    # ---- Locate the data header row + resolve columns BY NAME ----
    # The aging bond sheet grows every month (an extra Sales + Closing column per
    # month) and now carries an AGING MIX strip above the table, so neither the
    # header-row position nor the closing/MoC/zero column indices are fixed.
    # Hardcoding col 10/11/12 (the old behaviour) silently read Total Sales /
    # Feb Closing / Mar Closing on the current layout -- it only lined up by
    # coincidence for flat non-movers. Resolve everything by header text instead:
    # Severity, Brand, Pack, the LATEST "* Closing" column (the snapshot closing),
    # "Months of Cover", and "Zero Months".
    header_row = None
    for r in range(1, min(ws.max_row, 80) + 1):
        if str(ws.cell(row=r, column=1).value or "").strip().lower() == "severity":
            header_row = r
            break
    if header_row is None:
        raise SystemExit(f"Bond '{bond}': could not find the 'Severity' header row")

    col_of = {}
    closing_cols = []
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(row=header_row, column=c).value or "").strip()
        if not h:
            continue
        col_of[h.lower()] = c
        if h.lower().endswith("closing"):
            closing_cols.append(c)
    brand_col = col_of.get("brand", 2)
    pack_col  = col_of.get("pack", 3)
    moc_col   = col_of.get("months of cover")
    zero_col  = col_of.get("zero months")
    latest_closing_col = closing_cols[-1] if closing_cols else None  # rightmost = snapshot month
    if not (moc_col and zero_col and latest_closing_col):
        raise SystemExit(
            f"Bond '{bond}': missing expected header columns "
            f"(latest closing={latest_closing_col}, Months of Cover={moc_col}, "
            f"Zero Months={zero_col}) on header row {header_row}")

    shops = []
    current = None
    for r in range(header_row + 1, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        if a is None:
            continue
        a_str = str(a)
        if "·" in a_str and a_str.split("·")[0].strip().isdigit():
            if current:
                shops.append(current)
            code, name = [p.strip() for p in a_str.split("·", 1)]
            current = {"code": code, "name": name, "rows": []}
            continue
        if a_str.startswith(("●", "▲", "◆")):
            try:
                may_c = float(ws.cell(row=r, column=latest_closing_col).value or 0)
            except (TypeError, ValueError):
                may_c = 0.0
            try:
                zero_i = int(float(ws.cell(row=r, column=zero_col).value or 0))
            except (TypeError, ValueError):
                zero_i = 0
            row = {
                "severity": a_str.strip(),
                "brand":   str(ws.cell(row=r, column=brand_col).value or ""),
                "pack":    str(ws.cell(row=r, column=pack_col).value or ""),
                "may_c":   may_c,
                "moc":     ws.cell(row=r, column=moc_col).value,
                "zero":    zero_i,
            }
            if current:
                current["rows"].append(row)
    if current:
        shops.append(current)
    return {"field_staff": field_staff, "period": period, "shops": shops}


# ---------------------- formatting helpers ----------------------
def thin_border():
    return Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def fill(color):
    return PatternFill("solid", fgColor=color)

def sev_palette(sev):
    if sev.startswith("●"):
        return NM_DARK, NM_LIGHT, "Non-Moving"
    if sev.startswith("▲"):
        return CR_DARK, CR_LIGHT, "Critical"
    return SL_DARK, SL_LIGHT, "Slow"


# ---------------------- writer ----------------------
def add_bond_sheet(wb, data, bond):
    """Add one bond's sheet to an existing workbook. Same layout as before."""
    ws = wb.create_sheet(title=bond)

    n_cols = len(COLS)
    last_col = get_column_letter(n_cols)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A7"      # bond banner + meeting-date + column header pinned

    # ----- ROW 1: bond banner -----
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    cell = ws.cell(row=1, column=1, value=f"{bond}  —  Commitment Tracker")
    cell.font = Font(name="Calibri", size=22, bold=True, color="FFFFFFFF")
    cell.fill = fill(NAVY_DEEP)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 36

    # ----- ROW 2: sub-banner -----
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    sub = (f"Field Staff: {data['field_staff'] or '—'}   ·   "
           f"{data['period'] or 'Aging Stock review'}   ·   "
           f"Generated {date.today():%d %b %Y}")
    cell = ws.cell(row=2, column=1, value=sub)
    cell.font = Font(name="Calibri", size=10, italic=True, color=GOLD_DIM)
    cell.fill = fill(NAVY_DEEP)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18

    # ----- ROW 3: Meeting Date + Next Meeting Date inputs (Abhay fills these once) -----
    ws.row_dimensions[3].height = 22
    # "Meeting Date:" label  →  A3:B3
    label = ws.cell(row=3, column=1, value="Meeting Date:")
    label.font = Font(name="Calibri", size=11, bold=True, color="FF1F2937")
    label.alignment = Alignment(horizontal="right", vertical="center", indent=1)
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=2)
    # Yellow editable date cell — spans C3:D3
    ws.merge_cells(start_row=3, start_column=3, end_row=3, end_column=4)
    date_cell = ws.cell(row=3, column=3)
    date_cell.fill = fill(INPUT_FILL)
    date_cell.font = Font(name="Calibri", size=11, bold=True, color="FF1F2937")
    date_cell.alignment = Alignment(horizontal="center", vertical="center")
    date_cell.number_format = 'dd-mmm-yyyy'
    date_cell.border = thin_border()
    # "Next Meeting:" label  →  E3:F3
    next_label = ws.cell(row=3, column=5, value="Next Meeting:")
    next_label.font = Font(name="Calibri", size=11, bold=True, color="FF1F2937")
    next_label.alignment = Alignment(horizontal="right", vertical="center", indent=1)
    ws.merge_cells(start_row=3, start_column=5, end_row=3, end_column=6)
    # Yellow editable next-meeting date cell — spans G3:H3
    ws.merge_cells(start_row=3, start_column=7, end_row=3, end_column=8)
    next_cell = ws.cell(row=3, column=7)
    next_cell.fill = fill(INPUT_FILL)
    next_cell.font = Font(name="Calibri", size=11, bold=True, color="FF1F2937")
    next_cell.alignment = Alignment(horizontal="center", vertical="center")
    next_cell.number_format = 'dd-mmm-yyyy'
    next_cell.border = thin_border()
    # Helper text alongside — spans I3:N3
    hint = ws.cell(row=3, column=9,
                   value="←  fill both yellow cells. Sold cases counted from the aging-snapshot close onwards (the col-G closing date); Days remaining counted to Next Meeting.")
    hint.font = Font(name="Calibri", size=9, italic=True, color="FF6B7280")
    hint.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells(start_row=3, start_column=9, end_row=3, end_column=n_cols)

    # ----- ROW 4: legend strip -----
    legend = ("Yellow cells = INPUT (you fill).   "
              "Pale-blue cells = OUTPUT (the refresh script fills).   "
              "Run the refresh script manually whenever you want the latest status.")
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=n_cols)
    cell = ws.cell(row=4, column=1, value=legend)
    cell.font = Font(name="Calibri", size=9, italic=True, color="FF4B5563")
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[4].height = 16

    # ----- ROW 5: spacer -----
    ws.row_dimensions[5].height = 5

    # ----- ROW 6: column header -----
    header_row = 6
    for c_idx, (name, _, _) in enumerate(COLS, start=1):
        cell = ws.cell(row=header_row, column=c_idx, value=name)
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
        cell.fill = fill(NAVY_MID)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=MED)
    ws.row_dimensions[header_row].height = 30

    # Column widths
    for c_idx, (_, w, _) in enumerate(COLS, start=1):
        ws.column_dimensions[get_column_letter(c_idx)].width = w

    # ----- DATA rows + shop banners -----
    r = header_row + 1
    seq = 0   # global row counter across all shops
    for sh in data["shops"]:
        # Shop banner — a single merged row spanning all columns
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
        cell = ws.cell(row=r, column=1,
                       value=f"  {sh['code']}  ·  {sh['name']}")
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
        cell.fill = fill(NAVY_SOFT)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[r].height = 22
        r += 1

        for row in sh["rows"]:
            seq += 1
            sev_dark, sev_light, sev_label = sev_palette(row["severity"])

            # Reference cells (locked-feel light grey)
            ws.cell(row=r, column=1, value=seq).alignment = Alignment(horizontal="center")
            ws.cell(row=r, column=2, value=sh["code"]).alignment = Alignment(horizontal="center")
            ws.cell(row=r, column=3, value=sh["name"]).alignment = Alignment(horizontal="left", indent=1)
            # Severity badge cell
            sev_cell = ws.cell(row=r, column=4, value=sev_label)
            sev_cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
            sev_cell.fill = fill(sev_dark)
            sev_cell.alignment = Alignment(horizontal="center", vertical="center")
            # Brand / Pack
            ws.cell(row=r, column=5, value=row["brand"]).alignment = Alignment(horizontal="left", indent=1)
            ws.cell(row=r, column=6, value=row["pack"]).alignment = Alignment(horizontal="center")
            # May Closing reference (bold tier-color) — col 7 = G
            mc = ws.cell(row=r, column=7, value=row["may_c"])
            mc.number_format = '0.00'
            mc.font = Font(name="Calibri", size=10, bold=True, color=sev_dark)
            mc.alignment = Alignment(horizontal="center")
            # Latest Closing — col 8 = H (refresh script overwrites; blue OUTPUT cell)
            lc_cell = ws.cell(row=r, column=8)
            lc_cell.fill = fill(OUT_FILL)
            lc_cell.number_format = '0.00;-0.00;-'
            lc_cell.font = Font(name="Calibri", size=10, bold=True)
            lc_cell.alignment = Alignment(horizontal="center")
            lc_cell.border = thin_border()
            # Months Cover — col 9 = I
            moc_cell = ws.cell(row=r, column=9,
                               value=("NO SALES" if row["severity"].startswith("●")
                                      else (round(float(row["moc"]), 1) if isinstance(row["moc"], (int, float)) else row["moc"])))
            if row["severity"].startswith("●"):
                moc_cell.font = Font(name="Calibri", size=9, bold=True, color=NM_DARK)
            else:
                moc_cell.number_format = '0.0'
                moc_cell.font = Font(name="Calibri", size=10, color=sev_dark)
            moc_cell.alignment = Alignment(horizontal="center")
            # Zero Months — col 10 = J
            ws.cell(row=r, column=10, value=row["zero"]).alignment = Alignment(horizontal="center")

            # Apply reference-cell background fill to REF cols (1..7, 9..10).
            # Col 4 has its own severity fill; col 8 is OUTPUT (blue) so skip both.
            for c_idx in (1, 2, 3, 5, 6, 7, 9, 10):
                cell = ws.cell(row=r, column=c_idx)
                if cell.fill.fgColor.value in (None, "00000000"):
                    cell.fill = fill(REF_FILL)
                cell.border = thin_border()

            # INPUT cell — col 11 = K = Target Cases (commit), yellow editable
            input_cell = ws.cell(row=r, column=11)
            input_cell.fill = fill(INPUT_FILL)
            input_cell.font = Font(name="Calibri", size=11, bold=True, color="FF1F2937")
            input_cell.alignment = Alignment(horizontal="center", vertical="center")
            input_cell.border = thin_border()
            input_cell.number_format = '0.00;-0.00;-'

            # OUTPUT cells — refresh script writes 12 (Sold) + 14 (Days); formulas in 13 + 15
            target_addr = f"K{r}"
            sold_addr   = f"L{r}"
            pct_addr    = f"M{r}"
            days_addr   = f"N{r}"

            pct_formula = (f'=IF(AND(ISNUMBER({target_addr}),{target_addr}>0,'
                           f'ISNUMBER({sold_addr})),{sold_addr}/{target_addr},"")')
            status_formula = (
                f'=IF(NOT(ISNUMBER({target_addr})),"",'
                f'IF({target_addr}=0,"— No commit",'
                f'IF(NOT(ISNUMBER({sold_addr})),"Awaiting refresh",'
                f'IF({sold_addr}>={target_addr},"✅ Achieved",'
                f'IF({pct_addr}>=0.75,"🚀 On track",'
                f'IF({pct_addr}>=0.40,"⚠ At risk",'
                f'"🚫 Behind"))))))'
            )

            # col 12 = L — Cases Sold since aging snapshot (refresh script writes the number)
            sold_cell = ws.cell(row=r, column=12)
            sold_cell.fill = fill(OUT_FILL)
            sold_cell.number_format = '0.00;-0.00;-'
            sold_cell.alignment = Alignment(horizontal="center")
            sold_cell.border = thin_border()
            # col 13 = M — % achieved (formula)
            pct_cell = ws.cell(row=r, column=13, value=pct_formula)
            pct_cell.fill = fill(OUT_FILL)
            pct_cell.number_format = '0%'
            pct_cell.font = Font(name="Calibri", size=10, bold=True)
            pct_cell.alignment = Alignment(horizontal="center")
            pct_cell.border = thin_border()
            # col 14 = N — Days remaining (refresh script writes the number)
            days_cell = ws.cell(row=r, column=14)
            days_cell.fill = fill(OUT_FILL)
            days_cell.number_format = '0;-0;-'
            days_cell.alignment = Alignment(horizontal="center")
            days_cell.border = thin_border()
            # col 15 = O — Status (formula)
            status_cell = ws.cell(row=r, column=15, value=status_formula)
            status_cell.fill = fill(OUT_FILL)
            status_cell.font = Font(name="Calibri", size=10, bold=True)
            status_cell.alignment = Alignment(horizontal="center")
            status_cell.border = thin_border()

            ws.row_dimensions[r].height = 18
            r += 1

    # ----- TOTAL row at the bottom -----
    total_row = r + 1
    ws.row_dimensions[r].height = 6   # spacer
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=6)
    total_label = ws.cell(row=total_row, column=1, value="TOTAL (all shops · all SKUs)")
    total_label.font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
    total_label.fill = fill(NAVY_DEEP)
    total_label.alignment = Alignment(horizontal="right", vertical="center")
    # Sum May Closing (col 7 = G) — only positive (so reference cells with 0 don't pollute)
    sum_mc = ws.cell(row=total_row, column=7,
                     value=f"=SUMPRODUCT((G{header_row+1}:G{r-1}>0)*G{header_row+1}:G{r-1})")
    sum_mc.font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
    sum_mc.fill = fill(NAVY_DEEP); sum_mc.alignment = Alignment(horizontal="center")
    sum_mc.number_format = '#,##0.00'
    # Sum Latest Closing (col 8 = H)
    sum_lc = ws.cell(row=total_row, column=8,
                     value=f"=SUMPRODUCT((H{header_row+1}:H{r-1}>0)*H{header_row+1}:H{r-1})")
    sum_lc.font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
    sum_lc.fill = fill(NAVY_DEEP); sum_lc.alignment = Alignment(horizontal="center")
    sum_lc.number_format = '#,##0.00;-#,##0.00;-'
    # MoC, Zero — blank cells, navy fill (cols 9, 10)
    for c_idx in (9, 10):
        cell = ws.cell(row=total_row, column=c_idx)
        cell.fill = fill(NAVY_DEEP)
    # Total Target Cases (col 11 = K)
    total_target = ws.cell(row=total_row, column=11,
                           value=f"=SUM(K{header_row+1}:K{r-1})")
    total_target.font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
    total_target.fill = fill(NAVY_DEEP)
    total_target.number_format = '#,##0.00;-#,##0.00;-'
    total_target.alignment = Alignment(horizontal="center")
    # Total Sold so far (col 12 = L)
    total_sold = ws.cell(row=total_row, column=12,
                         value=f"=SUM(L{header_row+1}:L{r-1})")
    total_sold.font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
    total_sold.fill = fill(NAVY_DEEP)
    total_sold.number_format = '#,##0.00;-#,##0.00;-'
    total_sold.alignment = Alignment(horizontal="center")
    # Total % achieved (col 13 = M)
    total_pct = ws.cell(row=total_row, column=13,
                        value=f"=IF({total_target.coordinate}>0,{total_sold.coordinate}/{total_target.coordinate},\"\")")
    total_pct.font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
    total_pct.fill = fill(NAVY_DEEP); total_pct.number_format = '0%'
    total_pct.alignment = Alignment(horizontal="center")
    # Days remaining (col 14 = N) — blank
    ws.cell(row=total_row, column=14).fill = fill(NAVY_DEEP)
    # Status (col 15 = O)
    total_status = ws.cell(row=total_row, column=15,
                           value=(f'=IF({total_target.coordinate}=0,"— No commit yet",'
                                  f'IF({total_pct.coordinate}>=1,"✅ Achieved",'
                                  f'IF({total_pct.coordinate}>=0.75,"🚀 On track",'
                                  f'IF({total_pct.coordinate}>=0.40,"⚠ At risk","🚫 Behind"))))'))
    total_status.font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
    total_status.fill = fill(NAVY_DEEP); total_status.alignment = Alignment(horizontal="center")
    ws.row_dimensions[total_row].height = 26

    # ----- Conditional formatting on Status column (col 15 = O) -----
    data_range = f"O{header_row+1}:O{r-1}"
    # Reset existing CF (defensive — new workbook has none, but keep idempotent)
    for rule, font_color, bg_color in [
        ("✅ Achieved",        GREEN_DARK,  GREEN_LIGHT),
        ("🚀 On track",        "FF1565C0",  "FFBBDEFB"),
        ("⚠ At risk",          "FFE65100",  "FFFFE0B2"),
        ("🚫 Behind",          "FFC62828",  "FFFFCDD2"),
        ("Awaiting refresh",   "FF4B5563",  "FFECEFF1"),
        ("— No commit",        "FF4B5563",  "FFECEFF1"),
    ]:
        cf = CellIsRule(operator="equal", formula=[f'"{rule}"'],
                        fill=PatternFill("solid", fgColor=bg_color),
                        font=Font(name="Calibri", size=10, bold=True, color=font_color))
        ws.conditional_formatting.add(data_range, cf)

    # CF on % achieved column M — data bar based on value (0..100%)
    from openpyxl.formatting.rule import DataBarRule
    bar = DataBarRule(start_type="num", start_value=0,
                      end_type="num", end_value=1,
                      color="FF1565C0", showValue=True)
    ws.conditional_formatting.add(f"M{header_row+1}:M{r-1}", bar)

    return r - header_row - 1   # data row count for this bond sheet


# ---------------------- SUMMARY tab ----------------------
def build_summary(wb, bonds_in_order, sheet_meta):
    """SUMMARY tab — one row per bond, pulls totals from each bond sheet via cross-sheet refs.
    sheet_meta[bond] = {"data_first": int, "data_last": int, "header_row": int}
    """
    ws = wb.create_sheet(title="SUMMARY", index=0)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"

    n_cols = 10
    last_col = get_column_letter(n_cols)

    # ROW 1 — banner
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    cell = ws.cell(row=1, column=1, value="COMMITMENT TRACKER  —  K.S. Distillery  ·  All Bonds")
    cell.font = Font(name="Calibri", size=22, bold=True, color="FFFFFFFF")
    cell.fill = fill(NAVY_DEEP); cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 36

    # ROW 2 — sub-info
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    cell = ws.cell(row=2, column=1,
                   value=(f"15 bonds · 283 KSBC shops · Generated {date.today():%d %b %Y}   "
                          f"·   Open any bond tab to fill commitments; this SUMMARY rolls up automatically."))
    cell.font = Font(name="Calibri", size=10, italic=True, color=GOLD_DIM)
    cell.fill = fill(NAVY_DEEP); cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18

    # ROW 3 — spacer
    ws.row_dimensions[3].height = 6

    # ROW 4 — column header
    headers = ["#", "Bond", "Field Staff", "Meeting Date", "Next Meeting",
               "Shops × SKUs", "Target Cases", "Cases Sold", "% Achieved", "Status"]
    widths   = [5, 22, 24, 15, 15, 13, 14, 14, 12, 18]
    for c, (h, w) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
        cell.fill = fill(NAVY_MID)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=MED)
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[4].height = 28

    # Data rows — one per bond
    for idx, bond in enumerate(bonds_in_order, start=1):
        r = 4 + idx
        meta = sheet_meta[bond]
        sheet_ref = f"'{bond}'"
        first = meta["data_first"]
        last  = meta["data_last"]

        # # · Bond · Field Staff · Meeting Date · Next Meeting · Shops×SKUs · Target · Sold · % · Status
        ws.cell(row=r, column=1, value=idx).alignment = Alignment(horizontal="center")
        b_cell = ws.cell(row=r, column=2, value=bond)
        b_cell.font = Font(name="Calibri", size=11, bold=True)
        b_cell.alignment = Alignment(horizontal="left", indent=1)
        # Field staff — hardcode (cached at build time)
        ws.cell(row=r, column=3, value=meta["field_staff"] or "—").alignment = Alignment(horizontal="left", indent=1)
        # Meeting Date — pull from cell C3 of bond sheet
        md_cell = ws.cell(row=r, column=4, value=f"={sheet_ref}!C3")
        md_cell.number_format = 'dd-mmm-yyyy'
        md_cell.alignment = Alignment(horizontal="center")
        # Next Meeting — pull from cell G3 of bond sheet
        nm_cell = ws.cell(row=r, column=5, value=f"={sheet_ref}!G3")
        nm_cell.number_format = 'dd-mmm-yyyy'
        nm_cell.alignment = Alignment(horizontal="center")
        # Shops × SKUs (informational)
        ws.cell(row=r, column=6, value=f"{meta['shops']} × {meta['skus']}").alignment = Alignment(horizontal="center")
        # Target Cases — sum of col K on bond sheet (was J before Latest Closing added)
        target = ws.cell(row=r, column=7, value=f"=SUM({sheet_ref}!K{first}:K{last})")
        target.number_format = '#,##0.00;-#,##0.00;-'
        target.alignment = Alignment(horizontal="center")
        target.font = Font(name="Calibri", size=10, bold=True)
        # Cases Sold — sum of col L on bond sheet (was K)
        sold = ws.cell(row=r, column=8, value=f"=SUM({sheet_ref}!L{first}:L{last})")
        sold.number_format = '#,##0.00;-#,##0.00;-'
        sold.alignment = Alignment(horizontal="center")
        # % achieved (col 9)
        pct = ws.cell(row=r, column=9,
                      value=f"=IF(G{r}>0,H{r}/G{r},\"\")")
        pct.number_format = '0%'; pct.alignment = Alignment(horizontal="center")
        pct.font = Font(name="Calibri", size=10, bold=True)
        # Status (col 10) — references shifted by 1
        status = ws.cell(row=r, column=10,
                         value=(f'=IF(G{r}=0,"— No commit yet",'
                                f'IF(NOT(ISNUMBER(H{r})),"Awaiting refresh",'
                                f'IF(H{r}>=G{r},"✅ Achieved",'
                                f'IF(I{r}>=0.75,"🚀 On track",'
                                f'IF(I{r}>=0.40,"⚠ At risk","🚫 Behind")))))'))
        status.font = Font(name="Calibri", size=10, bold=True)
        status.alignment = Alignment(horizontal="center")
        # Borders on all cells
        for c in range(1, n_cols + 1):
            ws.cell(row=r, column=c).border = thin_border()
        ws.row_dimensions[r].height = 22

    # GRAND TOTAL row — label spans first 6 columns (#, Bond, Field Staff,
    # Meeting Date, Next Meeting, Shops × SKUs)
    total_r = 4 + len(bonds_in_order) + 1
    ws.merge_cells(start_row=total_r, start_column=1, end_row=total_r, end_column=6)
    lbl = ws.cell(row=total_r, column=1, value="TOTAL · 15 BONDS")
    lbl.font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
    lbl.fill = fill(NAVY_DEEP); lbl.alignment = Alignment(horizontal="right", vertical="center")
    # Total target (col 7=G) / sold (col 8=H) / % (col 9=I) / status (col 10=J)
    last_data_r = 4 + len(bonds_in_order)
    t_target = ws.cell(row=total_r, column=7, value=f"=SUM(G5:G{last_data_r})")
    t_sold   = ws.cell(row=total_r, column=8, value=f"=SUM(H5:H{last_data_r})")
    t_pct    = ws.cell(row=total_r, column=9, value=f"=IF(G{total_r}>0,H{total_r}/G{total_r},\"\")")
    t_status = ws.cell(row=total_r, column=10,
                       value=(f'=IF(G{total_r}=0,"— No commit yet",'
                              f'IF(NOT(ISNUMBER(H{total_r})),"Awaiting refresh",'
                              f'IF(H{total_r}>=G{total_r},"✅ Achieved",'
                              f'IF(I{total_r}>=0.75,"🚀 On track",'
                              f'IF(I{total_r}>=0.40,"⚠ At risk","🚫 Behind")))))'))
    for cell in (t_target, t_sold, t_pct, t_status):
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
        cell.fill = fill(NAVY_DEEP); cell.alignment = Alignment(horizontal="center")
    t_target.number_format = '#,##0.00;-#,##0.00;-'
    t_sold.number_format   = '#,##0.00;-#,##0.00;-'
    t_pct.number_format    = '0%'
    ws.row_dimensions[total_r].height = 26

    # Conditional formatting on Status column (now col J = 10)
    status_range = f"J5:J{total_r}"
    for rule_text, font_color, bg_color in [
        ("✅ Achieved",       GREEN_DARK,  GREEN_LIGHT),
        ("🚀 On track",       "FF1565C0",  "FFBBDEFB"),
        ("⚠ At risk",         "FFE65100",  "FFFFE0B2"),
        ("🚫 Behind",         "FFC62828",  "FFFFCDD2"),
        ("Awaiting refresh",  "FF4B5563",  "FFECEFF1"),
        ("— No commit yet",   "FF4B5563",  "FFECEFF1"),
    ]:
        cf = CellIsRule(operator="equal", formula=[f'"{rule_text}"'],
                        fill=PatternFill("solid", fgColor=bg_color),
                        font=Font(name="Calibri", size=10, bold=True, color=font_color))
        ws.conditional_formatting.add(status_range, cf)

    # Data bar on % achieved (col I = 9)
    from openpyxl.formatting.rule import DataBarRule
    bar = DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                      color="FF1565C0", showValue=True)
    ws.conditional_formatting.add(f"I5:I{last_data_r}", bar)


# ---------------------- top-level builder ----------------------
def build_workbook(xlsx_in, bonds, out_path):
    """Build the one combined workbook: SUMMARY + 15 bond sheets."""
    wb = Workbook()
    # Drop the default blank sheet
    wb.remove(wb.active)

    sheet_meta = {}
    for bond in bonds:
        data = parse_bond(xlsx_in, bond)
        n = add_bond_sheet(wb, data, bond)
        # Record metadata for the SUMMARY tab — header_row in bond sheets is 6;
        # data starts at row 7 and ends at last data row (skipping shop banners is fine
        # because banner rows have blank J/K so SUM ignores them)
        # First data row index: 7. Last data row index = 7 + n - 1.
        # n includes BOTH banner rows and data rows, so last = 6 + n.
        sheet_meta[bond] = {
            "data_first": 7,
            "data_last":  6 + n,
            "shops": len(data["shops"]),
            "skus":  sum(len(sh["rows"]) for sh in data["shops"]),
            "field_staff": data["field_staff"],
        }

    build_summary(wb, bonds, sheet_meta)

    Path(os.path.dirname(out_path)).mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


# ---------------------- cli ----------------------
def main():
    p = argparse.ArgumentParser()
    p.add_argument("--workbook", required=True)
    p.add_argument("--bonds", required=True,
                   help="Comma-separated bond names, or 'ALL' for the standard 15")
    p.add_argument("--out", required=True)
    args = p.parse_args()

    # Note: Bond tabs have Meeting Date in C3 (merged C3:D3) and Next Meeting Date
    # in G3 (merged G3:H3). Both are yellow input cells. The refresh script reads
    # both; if G3 is blank it falls back to meeting_date + 30 days.
    if args.bonds.upper() == "ALL":
        bonds = ["KOTTAYAM", "THODUPUZHA", "TRIPUNITHURA", "PATHANAMTHITTA", "ALUVA",
                 "ATTINGAL", "THRISSUR", "KOTTARAKARA", "NEDUMANGAD", "PALAKKAD",
                 "KANNUR", "KOZHIKODE", "PERINTHALMANNA", "KOLLAM", "ALAPPUZHA"]
    else:
        bonds = [b.strip() for b in args.bonds.split(",")]

    build_workbook(args.workbook, bonds, args.out)
    print(f"OK: {args.out}")
    print(f"  bonds: {len(bonds)}  ·  SUMMARY tab + {len(bonds)} bond tabs")


if __name__ == "__main__":
    main()
