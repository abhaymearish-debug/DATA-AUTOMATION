#!/usr/bin/env python3
"""
BUILD LOAD DISPATCH TRACKER  —  K.S. Distillery factory dispatch register.

Rebuilds `Load dispatch/LOAD DISPATCH TRACKER.xlsx` from its own DISPATCH LOG
data, applying the KSD dark navy + gold house design and every fix from the
3-agent audit of 25 Jul 2026.

USAGE
    python3 build_load_dispatch_tracker.py --out <scratch.xlsx> [--base "<Claude folder>"]
    python3 build_load_dispatch_tracker.py --out <scratch.xlsx> --source <existing.xlsx>

DESIGN RULES THAT MUST NOT BE BROKEN
  1. DISPATCH LOG data starts at ROW 5, and columns B/C/D/E/F stay
     Dispatch Date / Warehouse / Cases / Unload Date / Status — in that order.
     `.claude/scripts/refresh_warehouse_live_artifact.py::read_factory_loads()`
     reads this file with iter_rows(min_row=5) and indexes r[1..5].
     Columns G+ are free (Remarks moved to I on 25 Jul 2026).
  2. MERGE FIRST, THEN STYLE EVERY CELL of every merged range. openpyxl 3.1.5
     drops style-only merge interiors on load+save, so a merged band styled
     before merging renders as a single coloured column.
  3. This script must be the LAST writer. Any later openpyxl pass re-drops
     merged-range interior fills.
  4. Period comparison is NUMERIC (SUMIFS/COUNTIFS against real date bounds in
     LISTS!E2:F2), never TEXT(date,"YYYYMM") string compare — the string method
     is volatile, ~170k string conversions per recalc, and returns different
     results in Excel vs LibreOffice for blank cells.
"""
from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# --------------------------------------------------------------- palette --
NAVY_DEEP = "FF0D1B4A"   # canvas
PANEL     = "FF14224F"   # panel body
PANEL_ALT = "FF182A5A"   # panel zebra / tile body
HEADER    = "FF223466"   # table header band
BANNER    = "FF2E3F70"   # cluster banner
GOLD      = "FFFFB300"
GOLD_DIM  = "FFFFD54F"
STEEL     = "FFB9C4E8"
STEEL_DIM = "FF8FA3C8"
WHITE     = "FFFFFFFF"
CYAN      = "FF4FC3F7"
GREEN     = "FF81C784"
AMBER     = "FFFFB74D"
CORAL     = "FFE57373"

# light palette - DISPATCH LOG body (a data-entry sheet reads better light)
L_BODY    = "FFFFFFFF"
L_ZEBRA   = "FFF1F5FA"
L_GRID    = "FFE3E6EB"
L_TEXT    = "FF1F2937"
T_GREEN   = "FF2E7D32"; T_GREEN_BG = "FFDCEDC8"
T_AMBER   = "FFE65100"; T_AMBER_BG = "FFFFE0B2"
T_RED     = "FFC62828"; T_RED_BG   = "FFFFCDD2"

FONT = "Aptos Narrow"
FULL_TRUCK = 720            # confirmed truck ceiling: 27 of 54 loads are exactly 720

# warehouse -> ASM cluster (geographic, mirrors WH_CLUSTER in
# .claude/scripts/warehouse_live_template.html). BATTATHUR is a best-guess.
CLUSTERS = [
    ("CLUSTER 1", "SOUTH", ["ALAPPUZHA", "ATTINGAL", "BALARAMAPURAM", "KARUNAGAPPALLY",
                            "KOLLAM", "KOTTARAKARA", "MENAMKULAM", "NEDUMANGAD",
                            "PATHANAMTHITA", "THIRUVALLA"]),
    ("CLUSTER 2", "CENTRAL", ["ALUVA", "AYARKKUNNAM", "CHALAKUDY", "KADAVANTHRA",
                              "KOTHAMANGALAM", "KOTTAYAM", "PERUMBAVOOR", "THODUPUZHA",
                              "THRISSUR", "TRIPUNITHURA"]),
    ("CLUSTER 3", "NORTH", ["BATTATHUR", "KALPETTA", "KANNUR", "KOZHIKODE", "MENONPARA",
                            "NADUVANNUR", "PALAKKAD", "PERINTHALMANNA"]),
]
WAREHOUSES = sorted(w for _, _, ws in CLUSTERS for w in ws)

# --------------------------------------------------------- number formats -
CS0   = '#,##0;-#,##0;"·"'
NUM0  = '0;-0;"·"'
NUM1  = '0.0;-0.0;"·"'
NUMZ  = '0;-0;0'          # day counts: zero is meaningful, never a dot
PCT0  = '0%;-0%;"·"'
PCT1  = '0.0%;-0.0%;"·"'
DATEF = 'dd-mmm-yy;;"·"'

N_PERIOD = 69           # period table rows (LISTS!B1:D69)
N_WH = len(WAREHOUSES)  # 28
LAST_ROW = 504


# ---------------------------------------------------------------- helpers -
def F(sz=10, b=False, i=False, color=WHITE):
    return Font(name=FONT, size=sz, bold=b, italic=i, color=color)


def fill(hex_):
    return PatternFill("solid", fgColor=hex_) if hex_ else PatternFill()


def side(color, style="thin"):
    return Side(style=style, color=color)


def paint(ws, ref, bg=None, font=None, align=None, border=None, numfmt=None):
    """Style every cell in a range. Safe on merged ranges IF called after merge."""
    got = ws[ref]
    rows = got if isinstance(got, tuple) else ((got,),)
    for row in rows:
        row = row if isinstance(row, tuple) else (row,)
        for c in row:
            if bg is not None:
                c.fill = fill(bg)
            if font is not None:
                c.font = font
            if align is not None:
                c.alignment = align
            if border is not None:
                c.border = border
            if numfmt is not None:
                c.number_format = numfmt


def band(ws, ref, value=None, bg=None, font=None, align=None, border=None, numfmt=None):
    """MERGE FIRST, THEN STYLE - the only safe order under openpyxl 3.1.5."""
    if ":" in ref:
        ws.merge_cells(ref)
    anchor = ref.split(":")[0]
    if value is not None:
        ws[anchor] = value
    paint(ws, ref, bg=bg, font=font, align=align, border=border, numfmt=numfmt)
    return ws[anchor]


LEFT   = Alignment(horizontal="left", vertical="center", indent=1)
CENTER = Alignment(horizontal="center", vertical="center")
RIGHT  = Alignment(horizontal="right", vertical="center", indent=1)
RIGHT0 = Alignment(horizontal="right", vertical="center")
WRAPC  = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAPL  = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)


def LOG(col, r1=5, r2=LAST_ROW):
    return "'DISPATCH LOG'!${0}${1}:${0}${2}".format(col, r1, r2)


def logo_png(base: Path):
    """KSD shield, black flood knocked out via corner-BFS. Non-fatal."""
    try:
        import os
        import tempfile
        from collections import deque

        from PIL import Image as PILImage
        src = base / "Internal Docs" / "KSD LOGO.png"
        if not src.exists():
            return None
        im = PILImage.open(src).convert("RGBA")
        im = im.resize((max(1, int(im.width * 260 / im.height)), 260), PILImage.LANCZOS)
        px, (w, h) = im.load(), im.size
        seen, dq = set(), deque([(0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1)])
        while dq:
            x, y = dq.popleft()
            if (x, y) in seen or not (0 <= x < w and 0 <= y < h):
                continue
            seen.add((x, y))
            p = px[x, y]
            if p[3] == 0 or p[0] > 46 or p[1] > 46 or p[2] > 46:
                continue
            px[x, y] = (0, 0, 0, 0)
            dq.extend([(x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)])
        im = im.resize((max(1, int(w * 46 / h)), 46), PILImage.LANCZOS)
        fd, out = tempfile.mkstemp(prefix="ksd_logo_ldt_", suffix=".png")
        os.close(fd)
        im.save(out)
        return out
    except Exception:
        return None


# ------------------------------------------------------------ read the log -
def read_log(path: Path):
    """Pull every data row out of the existing DISPATCH LOG, verbatim."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["DISPATCH LOG"]
    # Remarks moved G -> I on 26 Jul 2026; G now holds the computed "Lag (d)".
    # Only fall back to G on a genuine pre-26-Jul sheet, otherwise an empty
    # Remarks cell silently ingests the lag number.
    legacy = "LAG" not in str(ws.cell(4, 7).value or "").upper()
    out = []
    for r in range(5, ws.max_row + 1):
        d = ws.cell(r, 2).value
        if d is None:
            continue
        wh = ws.cell(r, 3).value
        cs = ws.cell(r, 4).value
        ud = ws.cell(r, 5).value
        # remarks lived in G in the pre-25-Jul-2026 layout, I after
        rm = ws.cell(r, 9).value or (ws.cell(r, 7).value if legacy else None)
        if isinstance(d, dt.datetime):
            d = d.date()
        if isinstance(ud, dt.datetime):
            ud = ud.date()
        out.append({"d": d, "w": str(wh).strip().upper() if wh else "",
                    "cs": cs, "u": ud, "rm": rm})
    out.sort(key=lambda x: (x["d"], x["w"]))
    return out


# ================================================================== LISTS ==
def build_lists(ws, wh_rows, dash):
    """Hidden helper sheet.
      A1:A28   warehouse picklist (alphabetical)      -> named WAREHOUSES
      B1:D69   period table: label | start | end      -> B named MONTHS_LIST
      E1/F1    resolved YYYYMM strings ; E2/F2 real date bounds
      E4       concatenated integrity alerts
      G1:G28   cases + epsilon (LARGE tie-break)      H1:H28 warehouse name
      I1:I28   in-transit cases + eps                 J1:J28 avg lag + eps
      M1:M6    six-month trend totals (live)
      P1:Q28   warehouse -> cluster map
    """
    ws.sheet_state = "veryHidden"
    ws.sheet_properties.tabColor = "FF6B7280"
    ws.sheet_view.showGridLines = False

    for i, w in enumerate(WAREHOUSES, start=1):
        ws.cell(i, 1, w)

    # period table. Rows 2-4 are DYNAMIC so the dashboard never goes stale.
    rows = [("ALL TIME", '="190001"', '="999912"'),
            ("CURRENT MONTH", '=TEXT(TODAY(),"YYYYMM")', '=TEXT(TODAY(),"YYYYMM")'),
            ("LAST 3 MONTHS", '=TEXT(EDATE(TODAY(),-2),"YYYYMM")', '=TEXT(TODAY(),"YYYYMM")'),
            ("LAST 6 MONTHS", '=TEXT(EDATE(TODAY(),-5),"YYYYMM")', '=TEXT(TODAY(),"YYYYMM")')]
    for y in range(2026, 2031):                      # FY windows
        rows.append(("FY {0}-{1}".format(y, str(y + 1)[2:]),
                     '="{0}04"'.format(y), '="{0}03"'.format(y + 1)))
    y, m = 2026, 4                                   # 60 calendar months
    MON = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN",
           "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
    for _ in range(60):
        rows.append(("{0} {1}".format(MON[m - 1], y),
                     '="{0}{1:02d}"'.format(y, m), '="{0}{1:02d}"'.format(y, m)))
        m += 1
        if m == 13:
            m, y = 1, y + 1
    assert len(rows) == N_PERIOD, len(rows)
    for i, (lbl, s, e) in enumerate(rows, start=1):
        ws.cell(i, 2, lbl)
        ws.cell(i, 3, s)
        ws.cell(i, 4, e)

    sel = "DASHBOARD!" + dash["SEL"]
    ws["E1"] = '=IFERROR(VLOOKUP({0},LISTS!$B$1:$D${1},2,0),"190001")'.format(sel, N_PERIOD)
    ws["F1"] = '=IFERROR(VLOOKUP({0},LISTS!$B$1:$D${1},3,0),"999912")'.format(sel, N_PERIOD)
    ws["E2"] = ('=IFERROR(DATE(VALUE(LEFT(LISTS!$E$1,4)),'
                'VALUE(RIGHT(LISTS!$E$1,2)),1),DATE(1900,1,1))')
    ws["F2"] = ('=IFERROR(EOMONTH(DATE(VALUE(LEFT(LISTS!$F$1,4)),'
                'VALUE(RIGHT(LISTS!$F$1,2)),1),0),DATE(9999,12,31))')

    LB, LD, LG = LOG("B"), LOG("D"), LOG("G")
    ws["E4"] = (
        "="
        'IF(COUNTIF(LISTS!$B$1:$B$%d,%s)=0,'
        '"⚠  INVALID PERIOD — SHOWING ALL TIME        ","")' % (N_PERIOD, sel)
        + '&IF(ROUND(DASHBOARD!{0},2)<>ROUND(DASHBOARD!{1},2),'
          '"⚠  "&TEXT(DASHBOARD!{1}-DASHBOARD!{0},"#,##0")'
          '&" CS ON AN UNLISTED WAREHOUSE — CHECK LOG COL C        ","")'.format(
              dash["TOT_CS"], dash["KPI_CS"])
        + '&IF(SUMPRODUCT(({0}<>"")*(1-ISNUMBER({0})))>0,'
          '"⚠  TEXT DATE IN THE LOG — FIX BEFORE READING        ","")'.format(LB)
        + '&IF(COUNTIF({0},"<0")>0,'
          '"⚠  UNLOAD DATE BEFORE DISPATCH DATE        ","")'.format(LG)
        + '&IF(COUNTIF({0},">{1}")>0,'
          '"⚠  A LOAD EXCEEDS {1} CS — VERIFY AGAINST THE GATE LOG        ","")'.format(
              LD, FULL_TRUCK)
        + '&IF(COUNTA({0})>0,'
          '"⚠  LOG NEARLY FULL — EXTEND THE RANGES PAST ROW 504        ","")'.format(
              LOG("B", 485, LAST_ROW))
        + '&IF(AND(COUNTIF(LISTS!$B$1:$B${0},{1})>0,LISTS!$F$2<TODAY()),'
          '"◷  VIEWING A CLOSED PERIOD        ","")'.format(N_PERIOD, sel)
    )

    for i, r in enumerate(wh_rows, start=1):          # dashboard row of each WH
        ws.cell(i, 7, "=DASHBOARD!$D${0}+{1}*0.000001".format(r, i))
        ws.cell(i, 8, "=DASHBOARD!$B${0}".format(r))
        ws.cell(i, 9, "=DASHBOARD!$G${0}+{1}*0.000001".format(r, i))
        # avg-lag helper: DASHBOARD!J is "" when a warehouse has no unloaded
        # load in the period, and ""+number is #VALUE! — coerce to 0 first.
        ws.cell(i, 10, '=IF(ISNUMBER(DASHBOARD!$J${0}),DASHBOARD!$J${0},0)+{1}*0.0000001'.format(r, i))

    for i in range(6):                                # live 6-month trend
        off = -(5 - i)
        ws.cell(i + 1, 13,
                '=SUMIFS({0},{1},">="&EOMONTH(TODAY(),{2})+1,{1},"<="&EOMONTH(TODAY(),{3}))'
                .format(LD, LB, off - 1, off))

    r = 1
    for cid, _, whs in CLUSTERS:
        for w in whs:
            ws.cell(r, 16, w)
            ws.cell(r, 17, cid.replace("CLUSTER ", "C"))
            r += 1

    for col, w in (("A", 20), ("B", 17), ("C", 10), ("D", 10), ("E", 12), ("F", 12),
                   ("G", 12), ("H", 20), ("I", 12), ("J", 12), ("M", 12),
                   ("P", 20), ("Q", 8)):
        ws.column_dimensions[col].width = w
    paint(ws, "A1:Q{0}".format(max(N_PERIOD, N_WH)), font=F(9, color="FF000000"))


# =========================================================== DISPATCH LOG ==
LOG_COLS = [("#", 6.5), ("Dispatch Date", 14), ("Warehouse", 20), ("Cases", 10),
            ("Unload Date", 14), ("Status", 17), ("Lag (d)", 9.5), ("Fill %", 9.5),
            ("Remarks", 34)]


def build_log(ws, data):
    ws.sheet_properties.tabColor = NAVY_DEEP
    ws.sheet_view.showGridLines = False

    for i, (_, w) in enumerate(LOG_COLS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.column_dimensions["J"].width = 9
    ws.column_dimensions["J"].hidden = True          # cluster helper

    # hero (rows 1-3) - dark band, data stays at row 5 for the consumer
    for r, h in ((1, 32), (2, 16), (3, 6), (4, 26)):
        ws.row_dimensions[r].height = h
    band(ws, "A1:F1", "DISPATCH LOG", NAVY_DEEP, F(18, True, color=WHITE), LEFT)
    band(ws, "G1:I1",
         '=TEXT(COUNT($B$5:$B${0}),"#,##0")&" LOADS  ·  "'
         '&TEXT(SUM($D$5:$D${0}),"#,##0")&" CS"'.format(LAST_ROW),
         NAVY_DEEP, F(11, True, color=GOLD), RIGHT)
    band(ws, "A2:I2",
         "ONE ROW PER LOAD  ·  STATUS, LAG AND FILL % COMPUTE THEMSELVES  ·  "
         "THE DASHBOARD UPDATES ITSELF",
         NAVY_DEEP, F(9.5, i=True, color=GOLD_DIM), LEFT)
    band(ws, "A3:I3", None, NAVY_DEEP, border=Border(bottom=side(GOLD, "medium")))

    hdr_b = Border(bottom=side(GOLD, "thin"))
    for i, (t, _) in enumerate(LOG_COLS, start=1):
        c = ws.cell(4, i, t)
        c.fill, c.font, c.border = fill(HEADER), F(10.5, True, color=WHITE), hdr_b
        c.alignment = LEFT if i == 3 else CENTER
    ws.freeze_panes = "A5"

    grid = Border(bottom=side(L_GRID), right=side(L_GRID))
    for r in range(5, LAST_ROW + 1):
        idx = r - 5
        rec = data[idx] if idx < len(data) else None
        ws.row_dimensions[r].height = 17
        ws.cell(r, 1, '=IF($B{0}="","",COUNT($B$5:$B{0}))'.format(r))
        if rec:
            ws.cell(r, 2, rec["d"])
            ws.cell(r, 3, rec["w"])
            ws.cell(r, 4, rec["cs"])
            if rec["u"]:
                ws.cell(r, 5, rec["u"])
            if rec["rm"]:
                ws.cell(r, 9, rec["rm"])
        ws.cell(r, 6, '=IF($B{0}="","",IF($E{0}="","\U0001F69A IN TRANSIT",'
                      'IF($E{0}<$B{0},"⚠ CHECK DATES","✅ UNLOADED")))'.format(r))
        ws.cell(r, 7, '=IF(OR($B{0}="",$E{0}=""),"",IFERROR($E{0}-$B{0},""))'.format(r))
        ws.cell(r, 8, '=IF($D{0}="","",IFERROR($D{0}/{1},""))'.format(r, FULL_TRUCK))
        ws.cell(r, 10, '=IF($C{0}="","",IFERROR(VLOOKUP($C{0},LISTS!$P$1:$Q${1},2,0),"?"))'
                .format(r, N_WH))

        zebra = L_ZEBRA if (r - 5) % 2 else L_BODY
        for cc in range(1, 10):
            cell = ws.cell(r, cc)
            cell.fill, cell.border = fill(zebra), grid
            cell.font = F(10, color=L_TEXT)
            cell.alignment = LEFT if cc in (3, 9) else CENTER
        ws.cell(r, 9).alignment = WRAPL
        ws.cell(r, 2).number_format = "dd-mmm-yy"
        ws.cell(r, 5).number_format = 'dd-mmm-yy;;"·"'
        ws.cell(r, 4).number_format = CS0
        ws.cell(r, 7).number_format = NUM0
        ws.cell(r, 8).number_format = PCT0
        ws.cell(r, 1).font = F(9.5, color="FF6B7280")

    nxt = 5 + len(data)
    if ws.sheet_view.selection:
        ws.sheet_view.selection[-1].activeCell = "B{0}".format(nxt)
        ws.sheet_view.selection[-1].sqref = "B{0}".format(nxt)
    ws.auto_filter.ref = "A4:I{0}".format(4 + max(len(data), 1))   # no blank rows

    cf = ws.conditional_formatting
    rng = "F5:F{0}".format(LAST_ROW)
    cf.add(rng, FormulaRule(formula=['ISNUMBER(SEARCH("IN TRANSIT",$F5))'],
                            fill=fill(T_AMBER_BG), font=F(10, True, color=T_AMBER)))
    cf.add(rng, FormulaRule(formula=['ISNUMBER(SEARCH("UNLOADED",$F5))'],
                            fill=fill(T_GREEN_BG), font=F(10, True, color=T_GREEN)))
    cf.add(rng, FormulaRule(formula=['ISNUMBER(SEARCH("CHECK DATES",$F5))'],
                            fill=fill(T_RED_BG), font=F(10, True, color=T_RED)))
    cf.add("G5:G{0}".format(LAST_ROW),
           CellIsRule(operator="greaterThanOrEqual", formula=["4"],
                      fill=fill(T_RED_BG), font=F(10, True, color=T_RED)))
    cf.add("G5:G{0}".format(LAST_ROW),
           CellIsRule(operator="equal", formula=["3"],
                      fill=fill(T_AMBER_BG), font=F(10, True, color=T_AMBER)))
    cf.add("D5:D{0}".format(LAST_ROW),
           CellIsRule(operator="greaterThan", formula=[str(FULL_TRUCK)],
                      fill=fill(T_RED_BG), font=F(10, True, color=T_RED)))
    # duplicate load tripwire - same date + warehouse + cases logged twice
    cf.add("A5:I{0}".format(LAST_ROW), FormulaRule(
        formula=['AND($B5<>"",SUMPRODUCT(($B$5:$B${0}=$B5)*($C$5:$C${0}=$C5)*'
                 '($D$5:$D${0}=$D5))>1)'.format(LAST_ROW)],
        fill=fill("FFFFF3E0")))

    # entry guards (all three were missing before 25 Jul 2026)
    dv_w = DataValidation(type="list", formula1="WAREHOUSES", allow_blank=True,
                          showErrorMessage=True, errorStyle="stop",
                          errorTitle="Not a Bevco warehouse",
                          error="Pick one of the 28 warehouses from the list. A free-typed "
                                "name is invisible to every dashboard figure.",
                          showInputMessage=True, promptTitle="Warehouse",
                          prompt="Click the arrow and choose from the list.")
    dv_w.add("C5:C{0}".format(LAST_ROW))
    dv_c = DataValidation(type="whole", operator="between", formula1="1",
                          formula2=str(FULL_TRUCK), allow_blank=True,
                          showErrorMessage=True, errorStyle="warning",
                          errorTitle="Unusual load size",
                          error="Every logged truck so far is 190-{0} cases. Continue only "
                                "if the gate note really says this.".format(FULL_TRUCK))
    dv_c.add("D5:D{0}".format(LAST_ROW))
    dv_d = DataValidation(type="date", operator="between", formula1="DATE(2025,1,1)",
                          formula2="TODAY()+30", allow_blank=True,
                          showErrorMessage=True, errorStyle="warning",
                          errorTitle="Check the dispatch date",
                          error="That date is outside the expected window.")
    dv_d.add("B5:B{0}".format(LAST_ROW))
    dv_u = DataValidation(type="date", operator="greaterThanOrEqual", formula1="$B5",
                          allow_blank=True, showErrorMessage=True, errorStyle="stop",
                          errorTitle="Unload before dispatch",
                          error="The unload date cannot precede the dispatch date.")
    dv_u.add("E5:E{0}".format(LAST_ROW))
    for dv in (dv_w, dv_c, dv_d, dv_u):
        ws.add_data_validation(dv)

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:4"
    ws.print_area = "A1:I{0}".format(4 + max(len(data), 1))
    ws.oddFooter.left.text = "&A"
    ws.oddFooter.right.text = "Page &P of &N"


# ============================================================== DASHBOARD ==
COLW = [("A", 1.6), ("B", 20), ("C", 6.5), ("D", 14), ("E", 12), ("F", 12),
        ("G", 13.5), ("H", 12), ("I", 7), ("J", 7), ("K", 2.4), ("L", 14),
        ("M", 10), ("N", 14), ("O", 14), ("P", 1.6)]

TABLE_HDR = ["Warehouse", "Loads", "Cases", "Share", "Fill %", "In-Transit",
             "Last\nDispatch", "Days\nAgo", "Avg\nDays"]


def build_dashboard(ws, base: Path):
    ws.sheet_properties.tabColor = GOLD
    ws.sheet_view.showGridLines = False
    for col, w in COLW:
        ws.column_dimensions[col].width = w

    R_HERO, R_SUB, R_RULE = 2, 3, 4
    R_KL, R_KV, R_KC = 6, 7, 8              # KPI label / value / caption
    R_SEL = 11                               # period selector + alert strip
    R_SECT = 13                              # section headers
    R_HDR = 14                               # table column headers
    r = 15
    blocks = []
    for cid, region, whs in CLUSTERS:
        blocks.append({"banner": r, "first": r + 1, "last": r + len(whs),
                       "cid": cid, "region": region, "whs": whs})
        r += len(whs) + 1
    R_TOT = r
    R_FOOT = R_TOT + 2
    LAST = R_FOOT + 1

    heights = {1: 6, R_HERO: 38, R_SUB: 18, R_RULE: 5, 5: 10, R_KL: 16, R_KV: 34,
               R_KC: 17, 9: 6, 10: 10, R_SEL: 26, 12: 10, R_SECT: 26, R_HDR: 38,
               R_TOT: 26, R_TOT + 1: 8, R_FOOT: 24, LAST: 6}
    for rr in range(1, LAST + 1):
        ws.row_dimensions[rr].height = heights.get(rr, 17)

    paint(ws, "A1:P{0}".format(LAST), bg=NAVY_DEEP, font=F(10, color=STEEL), align=CENTER)

    # ---- hero
    band(ws, "B{0}:H{0}".format(R_HERO), "LOAD DISPATCH TRACKER",
         NAVY_DEEP, F(24, True, color=WHITE), LEFT)
    band(ws, "I{0}:N{0}".format(R_HERO),
         '="AS ON  "&UPPER(TEXT(TODAY(),"DD MMM YYYY"))',
         NAVY_DEEP, F(11, True, color=GOLD), RIGHT)
    band(ws, "B{0}:O{0}".format(R_SUB),
         '="K.S. DISTILLERY  →  BEVCO WAREHOUSES   ·   FACTORY DISPATCH REGISTER   ·   "'
         '&IF(COUNT({0})=0,"NO LOADS LOGGED YET",'
         '"REGISTER COVERS "&UPPER(TEXT(SMALL({0},1),"DD MMM YYYY"))'
         '&"  →  "&UPPER(TEXT(MAX({0}),"DD MMM YYYY")))'.format(LOG("B")),
         NAVY_DEEP, F(9.5, i=True, color=GOLD_DIM), LEFT)
    band(ws, "A{0}:P{0}".format(R_RULE), None, NAVY_DEEP,
         border=Border(bottom=side(GOLD, "medium")))

    lp = logo_png(base)
    if lp:
        try:
            from openpyxl.drawing.image import Image as XLImage
            img = XLImage(lp)
            img.anchor = "O{0}".format(R_HERO)
            ws.add_image(img)
        except Exception:
            pass

    # ---- KPI tiles
    LB, LC, LD, LE, LG, LJ = (LOG("B"), LOG("C"), LOG("D"), LOG("E"),
                              LOG("G"), LOG("J"))
    PER = '{0},">="&LISTS!$E$2,{0},"<="&LISTS!$F$2'.format(LB)

    tiles = [
        ("B", "C", "TOTAL LOADS", CYAN,
         "=COUNTIFS({0})".format(PER), NUM0,
         '="AVG "&TEXT(IFERROR($D${0}/$B${0},0),"#,##0")&" CS PER LOAD"'.format(R_KV)),
        ("D", "E", "TOTAL CASES", GOLD,
         "=SUMIFS({0},{1})".format(LD, PER), CS0,
         '="PERIOD ▸  "&UPPER($C${0})'.format(R_SEL)),
        ("F", "G", "AVG DAYS TO UNLOAD", GREEN,
         '=IFERROR(ROUND(AVERAGEIFS({0},{1}),1),"·")'.format(LG, PER), NUM1,
         '="UNLOADED ≤ 1 DAY:  "&TEXT(IFERROR(COUNTIFS({0},"<=1",{1})'
         '/MAX(1,COUNTIFS({0},">=0",{1})),0),"0%")'.format(LG, PER)),
        ("H", "J", "TRUCK FILL RATE", GOLD_DIM,
         '=IFERROR($D${0}/($B${0}*{1}),"·")'.format(R_KV, FULL_TRUCK), PCT0,
         '=TEXT(COUNTIFS({0},{1},{2}),"0")&" FULL  ·  "'
         '&TEXT($B${3}-COUNTIFS({0},{1},{2}),"0")&" PART LOADS"'.format(
             LD, FULL_TRUCK, PER, R_KV)),
        ("L", "M", "IN TRANSIT   ·   LIVE", AMBER,
         '=SUMPRODUCT(({0}<>"")*({1}="")*{2})'.format(LB, LE, LD), CS0,
         '=TEXT(SUMPRODUCT(({0}<>"")*({1}="")),"0")&" LOADS ON THE ROAD"'.format(LB, LE)),
        ("N", "O", "DAYS SINCE LAST LOAD   ·   LIVE", STEEL,
         '=IF(COUNT({0})=0,"·",IF(MAX({0})>TODAY(),"—",TODAY()-MAX({0})))'.format(LB),
         NUMZ,
         '=IFERROR("LAST:  "&UPPER(TEXT(MAX({0}),"DD MMM"))&"  ·  "'
         '&INDEX({1},MATCH(MAX({0}),{0},0))'
         '&IF(COUNTIF({0},MAX({0}))>1,"  +"&COUNTIF({0},MAX({0}))-1&" MORE",""),'
         '"NO LOADS YET")'.format(LB, LC)),
    ]
    for c1, c2, label, accent, val, nf, cap in tiles:
        live = "LIVE" in label
        band(ws, "{0}{2}:{1}{2}".format(c1, c2, R_KL), label, PANEL_ALT,
             F(8.5, True, color=GOLD_DIM if live else STEEL_DIM),
             Alignment(horizontal="left", vertical="bottom", indent=1),
             Border(left=side(accent, "medium")))
        band(ws, "{0}{2}:{1}{2}".format(c1, c2, R_KV), val, PANEL_ALT,
             F(22, True, color=accent),
             Alignment(horizontal="left", vertical="center", indent=1),
             Border(left=side(accent, "medium")), nf)
        band(ws, "{0}{2}:{1}{2}".format(c1, c2, R_KC), cap, PANEL_ALT,
             F(8.5, color=STEEL),
             Alignment(horizontal="left", vertical="top", indent=1),
             Border(left=side(accent, "medium"),
                    bottom=side(GOLD if live else PANEL_ALT)))

    # ---- period selector + integrity strip
    band(ws, "B{0}".format(R_SEL), "PERIOD ▸", NAVY_DEEP, F(10, True, color=GOLD), RIGHT)
    band(ws, "C{0}:D{0}".format(R_SEL), "CURRENT MONTH", GOLD_DIM,
         F(11, True, color=NAVY_DEEP), CENTER,
         Border(left=side(GOLD, "medium"), right=side(GOLD, "medium"),
                top=side(GOLD, "medium"), bottom=side(GOLD, "medium")))
    band(ws, "F{0}:O{0}".format(R_SEL),
         '=IF(LISTS!$E$4="","✓   ALL INTEGRITY CHECKS PASS  ·  THE LOG IS CLEAN",'
         'LISTS!$E$4)',
         PANEL, F(10, True, color=GREEN), LEFT, Border(left=side(GOLD, "medium")))
    ws.conditional_formatting.add(
        "F{0}:O{0}".format(R_SEL),
        FormulaRule(formula=['LEFT($F${0},1)="⚠"'.format(R_SEL)],
                    fill=fill("FF4A1A1A"), font=F(10, True, color=CORAL), stopIfTrue=True))
    ws.conditional_formatting.add(
        "F{0}:O{0}".format(R_SEL),
        FormulaRule(formula=['LEFT($F${0},1)="◷"'.format(R_SEL)],
                    fill=fill(PANEL_ALT), font=F(10, True, color=GOLD_DIM), stopIfTrue=True))

    dv_p = DataValidation(type="list", formula1="MONTHS_LIST", allow_blank=False,
                          showErrorMessage=True, errorStyle="stop",
                          errorTitle="Pick a period",
                          error="Choose a period from the list. A typed value silently "
                                "falls back to ALL TIME.",
                          showInputMessage=True, promptTitle="Period",
                          prompt="Click the arrow to change what the dashboard shows.")
    dv_p.add("C{0}".format(R_SEL))
    ws.add_data_validation(dv_p)

    gold_rule = Border(bottom=side(GOLD, "medium"))
    band(ws, "B{0}:J{0}".format(R_SECT), "DISPATCH BY WAREHOUSE  ·  CLUSTERWISE",
         NAVY_DEEP, F(11.5, True, color=GOLD), LEFT, gold_rule)
    band(ws, "L{0}:O{0}".format(R_SECT), "PERIOD SNAPSHOT",
         NAVY_DEEP, F(11.5, True, color=GOLD), LEFT, gold_rule)

    for i, h in enumerate(TABLE_HDR):
        c = ws.cell(R_HDR, 2 + i, h)
        c.fill, c.font = fill(HEADER), F(10, True, color=WHITE)
        c.alignment = (WRAPL if i == 0 else
                       WRAPC if "\n" in h else CENTER)
        c.border = Border(bottom=side(GOLD, "thin"))

    STRUCTURAL = {R_HERO, R_SUB, R_RULE, R_KL, R_KV, R_KC, R_SEL, R_SECT,
                  R_HDR, R_TOT, R_TOT + 1, R_FOOT, LAST}

    def rh(row, h):
        """Row heights on the right panel must never clobber a structural row —
        both panels live on the same rows."""
        if row not in STRUCTURAL:
            ws.row_dimensions[row].height = h

    NFMT = (("C", NUM0), ("D", CS0), ("E", PCT1), ("F", PCT0),
            ("G", CS0), ("H", DATEF), ("I", NUMZ), ("J", NUM1))
    wh_rows = []
    thin = Border(bottom=side("FF243766"))
    for blk in blocks:
        br, f_, l_ = blk["banner"], blk["first"], blk["last"]
        cid = blk["cid"].replace("CLUSTER ", "C")
        ws["B{0}".format(br)] = "{0} · {1}".format(blk["cid"], blk["region"])
        ws["C{0}".format(br)] = "=SUM(C{0}:C{1})".format(f_, l_)
        ws["D{0}".format(br)] = "=SUM(D{0}:D{1})".format(f_, l_)
        ws["E{0}".format(br)] = '=IFERROR(D{0}/$D${1},"")'.format(br, R_TOT)
        ws["F{0}".format(br)] = '=IFERROR(D{0}/(C{0}*{1}),"")'.format(br, FULL_TRUCK)
        ws["G{0}".format(br)] = "=SUM(G{0}:G{1})".format(f_, l_)
        ws["H{0}".format(br)] = "=MAX(H{0}:H{1})".format(f_, l_)
        ws["I{0}".format(br)] = '=IF(H{0}=0,"",MAX(0,MIN(TODAY(),LISTS!$F$2)-H{0}))'.format(br)
        ws["J{0}".format(br)] = ('=IFERROR(ROUND(AVERAGEIFS({0},{1},"{2}",{3}),1),"")'
                                 .format(LG, LJ, cid, PER))
        paint(ws, "B{0}:J{0}".format(br), bg=BANNER,
              font=F(10.5, True, color=GOLD_DIM), align=CENTER, border=thin)
        ws["B{0}".format(br)].alignment = LEFT
        for col, nf in NFMT:
            ws["{0}{1}".format(col, br)].number_format = nf

        for i, w in enumerate(blk["whs"]):
            rr = f_ + i
            wh_rows.append(rr)
            ws["B{0}".format(rr)] = w
            ws["C{0}".format(rr)] = "=COUNTIFS({0},$B{1},{2})".format(LC, rr, PER)
            ws["D{0}".format(rr)] = "=SUMIFS({0},{1},$B{2},{3})".format(LD, LC, rr, PER)
            ws["E{0}".format(rr)] = '=IFERROR(D{0}/$D${1},"")'.format(rr, R_TOT)
            ws["F{0}".format(rr)] = '=IFERROR(D{0}/(C{0}*{1}),"·")'.format(rr, FULL_TRUCK)
            ws["G{0}".format(rr)] = ('=SUMPRODUCT(({0}=$B{1})*({2}<>"")*({3}="")*{4})'
                                     .format(LC, rr, LB, LE, LD))
            ws["H{0}".format(rr)] = "=_xlfn.MAXIFS({0},{1},$B{2},{3})".format(LB, LC, rr, PER)
            ws["I{0}".format(rr)] = '=IF(H{0}=0,"·",MAX(0,MIN(TODAY(),LISTS!$F$2)-H{0}))'.format(rr)
            ws["J{0}".format(rr)] = ('=IFERROR(ROUND(AVERAGEIFS({0},{1},$B{2},{3}),1),"·")'
                                     .format(LG, LC, rr, PER))
            bg = PANEL if (i % 2 == 0) else PANEL_ALT
            paint(ws, "B{0}:J{0}".format(rr), bg=bg, font=F(10, color=STEEL),
                  align=CENTER, border=thin)
            ws["B{0}".format(rr)].alignment = LEFT
            ws["B{0}".format(rr)].font = F(10, color=WHITE)
            ws["D{0}".format(rr)].alignment = RIGHT
            for col, nf in NFMT:
                ws["{0}{1}".format(col, rr)].number_format = nf

    b1, b2, b3 = (b["banner"] for b in blocks)
    ws["B{0}".format(R_TOT)] = "TOTAL"
    ws["C{0}".format(R_TOT)] = "=C{0}+C{1}+C{2}".format(b1, b2, b3)
    ws["D{0}".format(R_TOT)] = "=D{0}+D{1}+D{2}".format(b1, b2, b3)
    ws["E{0}".format(R_TOT)] = '=IF(D{0}=0,"",1)'.format(R_TOT)
    ws["F{0}".format(R_TOT)] = '=IFERROR(D{0}/(C{0}*{1}),"")'.format(R_TOT, FULL_TRUCK)
    ws["G{0}".format(R_TOT)] = "=G{0}+G{1}+G{2}".format(b1, b2, b3)
    ws["H{0}".format(R_TOT)] = "=_xlfn.MAXIFS({0},{1})".format(LB, PER)
    ws["I{0}".format(R_TOT)] = '=IF(H{0}=0,"",MAX(0,MIN(TODAY(),LISTS!$F$2)-H{0}))'.format(R_TOT)
    ws["J{0}".format(R_TOT)] = '=IFERROR(ROUND(AVERAGEIFS({0},{1}),1),"")'.format(LG, PER)
    paint(ws, "B{0}:J{0}".format(R_TOT), bg="FF374151", font=F(12, True, color=WHITE),
          align=CENTER, border=Border(top=side(GOLD, "medium")))
    ws["B{0}".format(R_TOT)].alignment = LEFT
    ws["D{0}".format(R_TOT)].font = F(12, True, color=GOLD)
    for col, nf in NFMT:
        ws["{0}{1}".format(col, R_TOT)].number_format = nf
    ws["E{0}".format(R_TOT)].number_format = PCT0

    wh_only = " ".join("{0}{1}:{0}{2}".format("D", b["first"], b["last"]) for b in blocks)
    ws.conditional_formatting.add(
        wh_only,
        DataBarRule(start_type="num", start_value=0, end_type="max",
                    color=GOLD[2:], showValue=True, minLength=0, maxLength=45))
    for col, rule in (("G", CellIsRule(operator="greaterThan", formula=["0"],
                                       font=F(10, True, color=AMBER))),
                      ("J", CellIsRule(operator="greaterThan", formula=["2.5"],
                                       font=F(10, True, color=CORAL)))):
        ws.conditional_formatting.add(
            " ".join("{0}{1}:{0}{2}".format(col, b["first"], b["last"]) for b in blocks), rule)

    # ---- right panel
    G1 = "LISTS!$G$1:$G${0}".format(N_WH)
    H1 = "LISTS!$H$1:$H${0}".format(N_WH)
    I1 = "LISTS!$I$1:$I${0}".format(N_WH)
    J1 = "LISTS!$J$1:$J${0}".format(N_WH)

    def top_of(helper, unit, dec=0):
        numfmt = "#,##0" + (".0" if dec else "")
        return ('=IF(MAX({0})<0.5,"—",INDEX({1},MATCH(MAX({0}),{0},0))'
                '&"  ·  "&TEXT(ROUND(MAX({0}),{2}),"{3}")&" {4}"'
                '&IF(SUMPRODUCT((ROUND({0},{2})=ROUND(MAX({0}),{2}))*1)>1,"  (TIED)",""))'
                .format(helper, H1, dec, numfmt, unit))

    snap = [
        ("WAREHOUSES SERVED", '=COUNTIF({0},">=0.5")&"  OF  {1}"'.format(G1, N_WH)),
        ("TOP BY CASES", top_of(G1, "CS")),
        ("BIGGEST IN-TRANSIT",
         '=IF(MAX({0})<0.5,"NONE — ALL CLEAR",INDEX({1},MATCH(MAX({0}),{0},0))'
         '&"  ·  "&TEXT(ROUND(MAX({0}),0),"#,##0")&" CS"'
         '&IF(SUMPRODUCT((ROUND({0},0)=ROUND(MAX({0}),0))*1)>1,"  (TIED)",""))'.format(I1, H1)),
        ("SLOWEST UNLOAD", top_of(J1, "DAYS", 1)),
        ("LATEST DISPATCH",
         '=IF($H${0}=0,"—",UPPER(TEXT($H${0},"DD MMM YYYY"))&"  ·  "'
         '&IFERROR(INDEX({1},MATCH($H${0},{2},0)),"—"))'.format(R_TOT, LC, LB)),
    ]
    rp = R_HDR + 1
    for i, (lbl, val) in enumerate(snap):
        bg = PANEL if i % 2 == 0 else PANEL_ALT
        band(ws, "L{0}:M{0}".format(rp), lbl, bg, F(9, color=STEEL_DIM), LEFT, thin)
        band(ws, "N{0}:O{0}".format(rp), val, bg, F(9.5, True, color=WHITE), RIGHT0, thin)
        rh(rp, 17)
        rp += 1

    def section(row, title):
        band(ws, "L{0}:O{0}".format(row), title, NAVY_DEEP, F(10.5, True, color=GOLD),
             LEFT, Border(bottom=side(GOLD, "thin")))
        rh(row, 22)

    rp += 1
    section(rp, "TRUCK UTILISATION")
    rp += 1
    util = [
        ("FULL LOADS ({0} CS)".format(FULL_TRUCK),
         '=TEXT(COUNTIFS({0},{1},{2}),"0")&"  OF  "&TEXT($B${3},"0")'.format(
             LD, FULL_TRUCK, PER, R_KV)),
        ("PART LOADS",
         '=TEXT($B${3}-COUNTIFS({0},{1},{2}),"0")&"  ·  avg "'
         '&TEXT(IFERROR(SUMIFS({0},{0},"<{1}",{2})/MAX(1,COUNTIFS({0},"<{1}",{2})),0),'
         '"#,##0")&" CS"'.format(LD, FULL_TRUCK, PER, R_KV)),
        ("AVG FILL vs TRUCK",
         '=TEXT(IFERROR($D${0}/($B${0}*{1}),0),"0%")&"  ·  "'
         '&TEXT(IFERROR($B${0}*{1}-$D${0},0),"#,##0")&" SLACK"'.format(R_KV, FULL_TRUCK)),
        ("LARGEST  ·  SMALLEST",
         '=IF($B${0}=0,"—",TEXT(_xlfn.MAXIFS({1},{2}),"#,##0")&" CS  ·  "'
         '&TEXT(_xlfn.MINIFS({1},{2}),"#,##0")&" CS")'.format(R_KV, LD, PER)),
    ]
    for i, (lbl, val) in enumerate(util):
        bg = PANEL if i % 2 == 0 else PANEL_ALT
        band(ws, "L{0}:M{0}".format(rp), lbl, bg, F(9, color=STEEL_DIM), LEFT, thin)
        band(ws, "N{0}:O{0}".format(rp), val, bg, F(9.5, True, color=WHITE), RIGHT0, thin)
        rh(rp, 17)
        rp += 1

    rp += 1
    section(rp, "TOP 5  —  WHERE THE CASES WENT")
    rp += 1
    top5_first = rp
    for n in range(1, 6):
        bg = PANEL if n % 2 else PANEL_ALT
        band(ws, "L{0}:M{0}".format(rp),
             '=IF(LARGE({0},{1})<0.5,"·",{1}&".    "&INDEX({2},MATCH(LARGE({0},{1}),{0},0)))'
             .format(G1, n, H1), bg, F(10, color=WHITE), LEFT, thin)
        band(ws, "N{0}".format(rp),
             '=IF(LARGE({0},{1})<0.5,"",ROUND(LARGE({0},{1}),0))'.format(G1, n),
             bg, F(10, True, color=GOLD), RIGHT, thin, CS0)
        band(ws, "O{0}".format(rp),
             '=IFERROR(IF(LARGE({0},{1})<0.5,"",ROUND(LARGE({0},{1}),0)/$D${2}),"")'
             .format(G1, n, R_TOT), bg, F(9.5, color=STEEL_DIM), RIGHT, thin, PCT1)
        rh(rp, 17)
        rp += 1
    ws.conditional_formatting.add(
        "N{0}:N{1}".format(top5_first, rp - 1),
        DataBarRule(start_type="num", start_value=0, end_type="max",
                    color=GOLD[2:], showValue=True, minLength=0, maxLength=40))

    rp += 1
    section(rp, "UNLOAD SPEED  —  PERIOD")
    rp += 1
    spd_first = rp
    tot_unl = 'MAX(1,COUNTIFS({0},">=0",{1}))'.format(LG, PER)
    buckets = [("≤ 1 DAY", 'COUNTIFS({0},"<=1",{1})'.format(LG, PER), GREEN, "FAST"),
               ("2 DAYS", 'COUNTIFS({0},">1",{0},"<=2",{1})'.format(LG, PER), GOLD_DIM, "OK"),
               ("3+ DAYS", 'COUNTIFS({0},">2",{1})'.format(LG, PER), CORAL, "SLOW")]
    for i, (lbl, expr, col, tier) in enumerate(buckets):
        bg = PANEL if i % 2 == 0 else PANEL_ALT
        band(ws, "L{0}".format(rp), lbl, bg, F(10, color=WHITE), LEFT, thin)
        band(ws, "M{0}".format(rp), "=" + expr, bg, F(10, True, color=WHITE), RIGHT, thin, NUM0)
        band(ws, "N{0}".format(rp), '=IFERROR({0}/{1},"")'.format(expr, tot_unl),
             bg, F(10, True, color=col), RIGHT, thin, PCT0)
        band(ws, "O{0}".format(rp), tier, bg, F(9, True, color=col), RIGHT, thin)
        rh(rp, 17)
        rp += 1
    ws.conditional_formatting.add(
        "N{0}:N{1}".format(spd_first, rp - 1),
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                    color=STEEL_DIM[2:], showValue=True, minLength=0, maxLength=40))

    rp += 1
    section(rp, "DISPATCH TREND  ·  LAST 6 MONTHS")
    rp += 1
    tr_first = rp
    for i in range(6):
        off = -(5 - i)
        bg = PANEL if i % 2 == 0 else PANEL_ALT
        band(ws, "L{0}".format(rp), '=UPPER(TEXT(EDATE(TODAY(),{0}),"MMM YY"))'.format(off),
             bg, F(10, color=WHITE), LEFT, thin)
        band(ws, "M{0}".format(rp), "=LISTS!$M${0}".format(i + 1), bg,
             F(10, True, color=GOLD if i == 5 else WHITE), RIGHT, thin, CS0)
        band(ws, "N{0}".format(rp),
             '=IFERROR(LISTS!$M${0}/MAX(1,SUM(LISTS!$M$1:$M$6)),"")'.format(i + 1),
             bg, F(10, color=STEEL), RIGHT, thin, PCT0)
        if i == 0:
            band(ws, "O{0}".format(rp), "—", bg, F(9.5, color=STEEL_DIM), RIGHT, thin)
        else:
            band(ws, "O{0}".format(rp),
                 '=IF(OR(LISTS!$M${0}=0,LISTS!$M${1}=0),"—",'
                 'IF(LISTS!$M${1}>=LISTS!$M${0},"▲  ","▼  ")'
                 '&TEXT(ABS(LISTS!$M${1}/LISTS!$M${0}-1),"0%"))'.format(i, i + 1),
                 bg, F(9.5, color=STEEL_DIM), RIGHT, thin)
            ws.conditional_formatting.add("O{0}".format(rp), FormulaRule(
                formula=['LEFT($O${0},1)="▲"'.format(rp)], font=F(9.5, True, color=GREEN)))
            ws.conditional_formatting.add("O{0}".format(rp), FormulaRule(
                formula=['LEFT($O${0},1)="▼"'.format(rp)], font=F(9.5, True, color=CORAL)))
        rh(rp, 17)
        rp += 1
    ws.conditional_formatting.add(
        "M{0}:M{1}".format(tr_first, rp - 1),
        DataBarRule(start_type="num", start_value=0, end_type="max",
                    color=GOLD[2:], showValue=True, minLength=0, maxLength=40))
    band(ws, "L{0}:O{0}".format(rp), None, NAVY_DEEP,
         border=Border(top=side(GOLD, "medium")))

    band(ws, "B{0}:O{0}".format(R_FOOT),
         "Fill % = cases ÷ {0} (one full truck).   Avg Days = unload − dispatch, "
         "in-transit excluded.   Days Ago runs to the period close.   "
         "IN TRANSIT and DAYS SINCE LAST LOAD are always live — they ignore the period "
         "selector.   Clusters follow the ASM geography.".format(FULL_TRUCK),
         NAVY_DEEP, F(8.5, i=True, color=STEEL_DIM),
         Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True))

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.35
    ws.page_margins.top = ws.page_margins.bottom = 0.35
    ws.print_area = "A1:P{0}".format(LAST)
    ws.oddFooter.left.text = "&A"
    ws.oddFooter.right.text = "Page &P of &N"
    ws.sheet_view.tabSelected = True
    ws.sheet_view.selection[0].activeCell = "C{0}".format(R_SEL)
    ws.sheet_view.selection[0].sqref = "C{0}".format(R_SEL)

    return {"SEL": "$C${0}".format(R_SEL), "TOT_CS": "$D${0}".format(R_TOT),
            "KPI_CS": "$D${0}".format(R_KV), "wh_rows": wh_rows, "last": LAST}


def suppress_error_flags(path, sheet_ranges):
    """Inject <ignoredErrors> so Excel stops drawing green corner triangles.

    The dashboard's cluster-banner SUM() rows legitimately "omit adjacent
    cells", and several cells fall back to the text "·" in a numeric column —
    both trip Excel's error checker even though the workbook is correct.
    openpyxl exposes no API for this element, so it is patched into the sheet
    XML post-save. <ignoredErrors> must sit after <headerFooter> and before
    <drawing>/<extLst> per CT_Worksheet's sequence.
    """
    import re
    import shutil
    import tempfile
    import zipfile

    FLAGS = ('formula="1" formulaRange="1" numberStoredAsText="1" '
             'emptyCellReference="1" evalError="1" unlockedFormula="1" '
             'twoDigitTextYear="1" listDataValidation="1"')
    TAIL = ("<drawing", "<legacyDrawing", "<picture", "<oleObjects",
            "<controls", "<tableParts", "<extLst")

    zin = zipfile.ZipFile(path)
    wbxml = zin.read("xl/workbook.xml").decode("utf-8")
    order = re.findall(r'<sheet [^>]*name="([^"]+)"', wbxml)
    targets = {}
    for i, nm in enumerate(order, start=1):
        if nm in sheet_ranges:
            targets["xl/worksheets/sheet{0}.xml".format(i)] = sheet_ranges[nm]

    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    import os
    os.close(fd)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in targets and b"<ignoredErrors" not in data:
                x = data.decode("utf-8")
                blk = ('<ignoredErrors><ignoredError sqref="{0}" {1}/>'
                       "</ignoredErrors>".format(targets[item.filename], FLAGS))
                cut = min((x.index(t) for t in TAIL if t in x),
                          default=x.rindex("</worksheet>"))
                data = (x[:cut] + blk + x[cut:]).encode("utf-8")
            zout.writestr(item, data)
    zin.close()
    shutil.move(tmp, path)


# =================================================================== main ==
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default=str(Path.home() / "mnt" / "Claude"))
    ap.add_argument("--source", default=None)
    ap.add_argument("--out", required=True)
    ap.add_argument("--fix", action="append", default=[],
                    help='audit correction "YYYY-MM-DD|WAREHOUSE|field=value", '
                         'e.g. "2026-06-08|THRISSUR|cs=449"')
    args = ap.parse_args()

    base = Path(args.base)
    src = (Path(args.source) if args.source
           else base / "Load dispatch" / "LOAD DISPATCH TRACKER.xlsx")
    data = read_log(src)

    for spec in args.fix:
        d_s, w_s, assign = spec.split("|")
        field, _, val = assign.partition("=")
        hit = [x for x in data if str(x["d"]) == d_s and x["w"] == w_s]
        if len(hit) != 1:
            raise SystemExit("--fix {0!r} matched {1} rows, expected 1".format(spec, len(hit)))
        old = hit[0][field]
        if field == "cs":
            hit[0][field] = float(val)
        elif field in ("d", "u"):
            hit[0][field] = dt.date.fromisoformat(val) if val else None
        else:
            hit[0][field] = val
        print("  FIX  {0} {1}  {2}: {3} -> {4}".format(d_s, w_s, field, old, hit[0][field]))

    n_loads = len(data)
    n_cases = sum(x["cs"] for x in data if isinstance(x["cs"], (int, float)))
    unknown = sorted({x["w"] for x in data} - set(WAREHOUSES))
    if unknown:
        raise SystemExit("ABORT - warehouse not in the 28-name list: {0}".format(unknown))

    wb = openpyxl.Workbook()
    dash = wb.active
    dash.title = "DASHBOARD"
    log = wb.create_sheet("DISPATCH LOG")
    lists = wb.create_sheet("LISTS")

    meta = build_dashboard(dash, base)
    build_log(log, data)
    build_lists(lists, meta["wh_rows"], meta)

    wb.defined_names["WAREHOUSES"] = DefinedName(
        "WAREHOUSES", attr_text="LISTS!$A$1:$A${0}".format(N_WH))
    wb.defined_names["MONTHS_LIST"] = DefinedName(
        "MONTHS_LIST", attr_text="LISTS!$B$1:$B${0}".format(N_PERIOD))

    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(args.out)
    suppress_error_flags(args.out, {"DASHBOARD": "A1:P{0}".format(meta["last"]),
                                    "DISPATCH LOG": "A1:J{0}".format(LAST_ROW)})

    print("BUILT  {0}".format(args.out))
    print("  {0} loads · {1:,.0f} cases · {2} -> {3}".format(
        n_loads, n_cases, min(x["d"] for x in data), max(x["d"] for x in data)))
    print("  {0} of {1} warehouses served · {2} in transit".format(
        len({x["w"] for x in data}), N_WH, sum(1 for x in data if not x["u"])))


if __name__ == "__main__":
    main()
