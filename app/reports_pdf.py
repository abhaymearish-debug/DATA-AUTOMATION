"""The K.S. Distillery report PDF layout, reproduced exactly.

MEASURED, NOT APPROXIMATED
--------------------------
Every number below was read out of the content stream of the PDFs the office
already receives (produced by jsPDF), not eyeballed from a screenshot:

    page            453.114 x 315.4 for 4 brands + TOTAL over 10 rows
    navy            rgb(0.039, 0.161, 0.31)
    gold            rgb(1, 0.741, 0.188)
    zebra           rgb(0.96, 0.97, 0.99)   on even rows
    greyed zero     rgb(0.78, 0.8, 0.839)   zeros are dimmed, not black
    title band      34pt, "K.S DISTILLERY" Helvetica-Bold 15 gold, centred
    gold band       34pt, report title Bold 9.6 navy at x=8,
                    period Bold 8.6 navy right-aligned,
                    group line Bold 9.6 navy centred
    header band     38.4pt navy, column labels Bold 7 gold, 7.6pt line pitch
    data row        15pt, Helvetica 8.5, figures centred in their column
    total row       gold 1pt rule, 19pt navy band, gold 1pt rule,
                    Helvetica-Bold 8.8 gold
    column width    59.4pt per numeric column; the label column takes the rest

Both the cluster PDFs and the "current view" export render through here, so the
two can never drift apart in appearance.
"""

from __future__ import annotations

from pathlib import Path

from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas as pdfcanvas

# ---- measured constants -------------------------------------------------

NAVY = colors.Color(0.039, 0.161, 0.31)
GOLD = colors.Color(1.0, 0.741, 0.188)
ZEBRA = colors.Color(0.96, 0.97, 0.99)
ZERO_GREY = colors.Color(0.78, 0.80, 0.839)
RULE_BODY = colors.Color(0.886, 0.902, 0.937)   # column separators over the rows
RULE_HEAD = colors.Color(0.227, 0.306, 0.525)   # the same, over the navy band
GRID_W = 0.2

TITLE_BAND_H = 34.0
SUB_BAND_H = 34.0
HEAD_BAND_H = 38.4
ROW_H = 15.0
TOTAL_BAND_H = 19.0
RULE_H = 1.0
FOOTER_H = 23.0        # measured: 150.4 + 15*rows is the page height

COL_W = 59.4          # every numeric column
DOC_W = 453.114       # the page width the office's PDFs use, every document
LABEL_MIN_HARD = 90.0 # never let the label column fall below this
LABEL_MIN = 150.0
PAD_L = 6.0

F_TITLE = 15.0
F_SUB = 9.6
F_PERIOD = 8.6
F_HEAD = 7.0
F_ROW = 8.5
F_TOTAL = 8.8
HEAD_PITCH = 7.6


def _wrap(text: str, size: float, max_w: float) -> list[str]:
    """Wrap a column label on word boundaries, measured by real glyph width."""
    words, lines, cur = text.split(), [], ""
    for w in words:
        trial = f"{cur} {w}".strip()
        if pdfmetrics.stringWidth(trial, "Helvetica-Bold", size) <= max_w or not cur:
            cur = trial
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def _fmt(v: float, round_off: bool) -> str:
    if not v:
        return "0"
    if round_off or abs(v - round(v)) < 0.005:
        return f"{round(v):,.0f}"
    return f"{v:,.2f}"


def page_size(labels: list[str], columns: list[str], rows: int) -> tuple[float, float, float]:
    """Kept for callers with a single page shape."""
    return document_size([{"labels": labels, "columns": columns, "rows": rows}])


def columns_width(n_columns: int) -> float:
    """Width of one numeric column on a page carrying `n_columns` of them.

    Normally the measured 59.4. It only shrinks on a page wide enough to squeeze
    the label column below what a shop name needs — better a narrower figure
    column than a name clipped to nothing.
    """
    if n_columns <= 0:
        return COL_W
    return min(COL_W, (DOC_W - LABEL_MIN_HARD) / n_columns)


def label_width_for(width: float, n_columns: int) -> float:
    """The label column takes whatever the numeric columns leave."""
    return width - columns_width(n_columns + 1) * (n_columns + 1)


def document_size(pages: list[dict]) -> tuple[float, float, float]:
    """One page size for the WHOLE document.

    Width is a constant: every PDF the office receives is 453.114pt wide
    whatever the cluster or the brand count, which is what makes them readable
    on a phone. Height is 150.4pt of banding plus 15pt per row, sized for the
    busiest page so the document never changes shape as you scroll.
    """
    max_cols = max(len(p["columns"]) for p in pages)
    max_rows = max(p["rows"] for p in pages)
    height = (TITLE_BAND_H + SUB_BAND_H + HEAD_BAND_H
              + ROW_H * max_rows + RULE_H + TOTAL_BAND_H + RULE_H + FOOTER_H)
    return DOC_W, height, label_width_for(DOC_W, max_cols)


def draw_page(c, *, width: float, height: float, label_w: float,
              report_title: str, period: str, group_line: str,
              label_heading: str, columns: list[str],
              rows: list[dict], totals: dict, total_label: str = "TOTAL",
              page_no: int = 1, pages: int = 1, round_off: bool = False) -> None:
    """Draw one page. `rows` are {'name': str, 'cells': {col: value}, 'total': v}."""
    all_cols = list(columns) + ["TOTAL"]
    col_w = columns_width(len(all_cols))

    # ---- title band ----
    y = height
    c.setFillColor(NAVY)
    c.rect(0, y - TITLE_BAND_H, width, TITLE_BAND_H, stroke=0, fill=1)
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", F_TITLE)
    c.drawCentredString(width / 2, y - 22.0, "K.S DISTILLERY")
    y -= TITLE_BAND_H

    # ---- gold band ----
    c.setFillColor(GOLD)
    c.rect(0, y - SUB_BAND_H, width, SUB_BAND_H, stroke=0, fill=1)
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", F_SUB)
    c.drawString(8.0, y - 15.0, report_title)
    c.setFont("Helvetica-Bold", F_PERIOD)
    c.drawRightString(width - 8.0, y - 15.0, period)
    c.setFont("Helvetica-Bold", F_SUB)
    c.drawCentredString(width / 2, y - 28.0, group_line)
    y -= SUB_BAND_H

    # ---- column header band ----
    c.setFillColor(NAVY)
    c.rect(0, y - HEAD_BAND_H, width, HEAD_BAND_H, stroke=0, fill=1)
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", F_HEAD)
    c.drawString(PAD_L, y - 21.7, label_heading)

    wrapped = [_wrap(col, F_HEAD, col_w - 4) for col in all_cols]
    n_lines = max(len(w) for w in wrapped)
    first = y - (21.4 - (n_lines - 1) * (HEAD_PITCH / 2))
    for i, lines in enumerate(wrapped):
        cx = label_w + col_w * i + col_w / 2
        # Each label's own block stays vertically centred in the band.
        start = y - (21.4 - (len(lines) - 1) * (HEAD_PITCH / 2))
        for j, line in enumerate(lines):
            c.drawCentredString(cx, start - j * HEAD_PITCH, line)
    c.setStrokeColor(RULE_HEAD)
    c.setLineWidth(GRID_W)
    for i in range(len(all_cols) + 1):
        x = label_w + col_w * i
        c.line(x, y - HEAD_BAND_H, x, y)
    head_bottom = y - HEAD_BAND_H
    y -= HEAD_BAND_H

    # ---- data rows ----
    for i, r in enumerate(rows):
        if i % 2 == 0:
            c.setFillColor(ZEBRA)
            c.rect(0, y - ROW_H, width, ROW_H, stroke=0, fill=1)

        base = y - 10.5
        c.setFillColor(colors.black)
        c.setFont("Helvetica", F_ROW)
        name = str(r["name"])
        while (pdfmetrics.stringWidth(name, "Helvetica", F_ROW) > label_w - 2 * PAD_L
               and len(name) > 3):
            name = name[:-1]
        c.drawString(PAD_L, base, name)

        for j, col in enumerate(columns):
            v = r["cells"].get(col, 0) or 0
            # A zero is dimmed so the eye lands on the figures that matter.
            c.setFillColor(ZERO_GREY if not v else colors.black)
            c.drawCentredString(label_w + col_w * j + col_w / 2, base, _fmt(v, round_off))
        c.setFillColor(colors.black)
        c.drawCentredString(label_w + col_w * len(columns) + col_w / 2, base,
                            _fmt(r.get("total", 0), round_off))
        y -= ROW_H

    if rows:
        c.setStrokeColor(RULE_BODY)
        c.setLineWidth(GRID_W)
        for i in range(len(all_cols) + 1):
            x = label_w + col_w * i
            c.line(x, y, x, head_bottom)

    # ---- total band ----
    c.setFillColor(GOLD)
    c.rect(0, y - RULE_H, width, RULE_H, stroke=0, fill=1)
    y -= RULE_H
    c.setFillColor(NAVY)
    c.rect(0, y - TOTAL_BAND_H, width, TOTAL_BAND_H, stroke=0, fill=1)
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", F_TOTAL)
    c.drawString(PAD_L, y - 12.5, total_label)
    for j, col in enumerate(columns):
        c.drawCentredString(label_w + col_w * j + col_w / 2, y - 12.5,
                            _fmt(totals.get(col, 0), round_off))
    c.drawCentredString(label_w + col_w * len(columns) + col_w / 2, y - 12.5,
                        _fmt(totals.get("__total__", 0), round_off))
    y -= TOTAL_BAND_H
    c.setFillColor(GOLD)
    c.rect(0, y - RULE_H, width, RULE_H, stroke=0, fill=1)

    # ---- footer ----
    c.setFillColor(colors.Color(0.45, 0.48, 0.55))
    c.setFont("Helvetica", 7)
    c.drawRightString(width - PAD_L, 16.0, f"Page {page_no} of {pages}")
    c.showPage()


# ---------------------------------------------------------------------------
# "Current view" export
# ---------------------------------------------------------------------------

ROWS_PER_PAGE = 26


def build_view_pdf(data: dict, out_path: Path, *, round_off: bool = False,
                   filters: dict | None = None) -> Path:
    """Print exactly what the report page is showing."""
    rows = data["rows"]
    brands = data["brands"]

    chunks = [rows[i:i + ROWS_PER_PAGE] for i in range(0, len(rows), ROWS_PER_PAGE)] or [[]]
    width, height, label_w = document_size([
        {"labels": [r["name"] for r in rows], "columns": brands,
         "rows": max(len(ch) for ch in chunks)}])

    period = ""
    span = data.get("span") or {}
    if span.get("min"):
        period = f"{span['min']} to {span['max']}"
    if filters and filters.get("date_from"):
        period = f"{filters['date_from']} to {filters.get('date_to') or span.get('max','')}"

    scope = f"{data['label'].upper()} VIEW"
    if filters:
        extra = [v for v in (filters.get("bond"), filters.get("warehouse")) if v]
        if extra:
            scope += "  ·  " + "  ·  ".join(extra)

    totals = dict(data["grand"]["cells"])
    totals["__total__"] = data["grand"]["total"]

    c = pdfcanvas.Canvas(str(out_path), pagesize=(width, height))
    for i, chunk in enumerate(chunks, start=1):
        draw_page(
            c, width=width, height=height, label_w=label_w,
            report_title="SECONDARY SALES - CUMULATIVE",
            period=period, group_line=scope,
            label_heading=data["label"].upper(), columns=brands,
            rows=chunk, totals=totals, total_label="GRAND TOTAL",
            page_no=i, pages=len(chunks), round_off=round_off,
        )
    c.save()
    return out_path


# ---------------------------------------------------------------------------
# Warehouse Stock Report
# ---------------------------------------------------------------------------
#
# A different geometry from the brandwise report, and again measured from the
# PDF the office receives rather than guessed:
#
#     page          A4 portrait, 595.276 x 841.89
#     title band    45.4pt navy
#     gold band 1   22.7pt  report name left, snapshot date right
#     gold band 2   22.7pt  warehouse name centred
#     header band   31.4pt navy, labels Helvetica-Bold 10 gold
#     row           32.45pt, Helvetica 11, values centred, zeros dimmed
#     total row     32.5pt navy band
#     columns       259.2 / 79.4 / 83.8 / 94.8 / 78.0  (= 595.2)

A4_W, A4_H = 595.276, 841.89
ST_TITLE_H = 45.4
ST_BAND_H = 22.7
ST_HEAD_H = 31.4
ST_ROW_H = 32.45
ST_COLS = [259.2, 79.4, 83.8, 94.8, 78.0]
ST_HEADINGS = ["ITEM NAME", "PACK", "PHYSICAL", "ALLOTABLE", "PENDING"]
ST_INK = colors.Color(0.059, 0.098, 0.176)
ST_ZERO = colors.Color(0.784, 0.804, 0.843)
ST_ROWS_PER_PAGE = 18


def _col_x(i: int) -> float:
    return sum(ST_COLS[:i])


def draw_stock_page(c, *, warehouse: str, as_of: str, rows: list[dict],
                    totals: dict | None, page_no: int, pages: int) -> None:
    """One warehouse page of the Warehouse Stock Report."""
    y = A4_H

    c.setFillColor(NAVY)
    c.rect(0, y - ST_TITLE_H, A4_W, ST_TITLE_H, stroke=0, fill=1)
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", F_TITLE)
    c.drawCentredString(A4_W / 2, y - 29.0, "K.S DISTILLERY")
    y -= ST_TITLE_H

    c.setFillColor(GOLD)
    c.rect(0, y - ST_BAND_H, A4_W, ST_BAND_H, stroke=0, fill=1)
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(9.9, y - 15.0, "WAREHOUSE STOCK REPORT")
    c.drawRightString(A4_W - 9.9, y - 15.0, as_of)
    y -= ST_BAND_H

    c.setFillColor(GOLD)
    c.rect(0, y - ST_BAND_H, A4_W, ST_BAND_H, stroke=0, fill=1)
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(A4_W / 2, y - 15.5, warehouse)
    y -= ST_BAND_H

    c.setFillColor(NAVY)
    c.rect(0, y - ST_HEAD_H, A4_W, ST_HEAD_H, stroke=0, fill=1)
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(9.9, y - 19.0, ST_HEADINGS[0])
    for i in range(1, len(ST_COLS)):
        c.drawCentredString(_col_x(i) + ST_COLS[i] / 2, y - 19.0, ST_HEADINGS[i])
    y -= ST_HEAD_H

    for i, r in enumerate(rows):
        if i % 2 == 0:
            c.setFillColor(ZEBRA)
            c.rect(0, y - ST_ROW_H, A4_W, ST_ROW_H, stroke=0, fill=1)
        base = y - 19.2
        c.setFillColor(ST_INK)
        c.setFont("Helvetica", 11)
        name = str(r["brand"])
        while (pdfmetrics.stringWidth(name, "Helvetica", 11) > ST_COLS[0] - 20
               and len(name) > 3):
            name = name[:-1]
        c.drawString(9.9, base, name)
        c.drawCentredString(_col_x(1) + ST_COLS[1] / 2, base, str(r["pack"]))
        for j, key in enumerate(("physical", "allotable", "pending"), start=2):
            v = r.get(key, 0) or 0
            c.setFillColor(ST_ZERO if not v else ST_INK)
            c.drawCentredString(_col_x(j) + ST_COLS[j] / 2, base, _fmt(v, True))
        y -= ST_ROW_H

    if totals is not None:
        c.setFillColor(NAVY)
        c.rect(0, y - ST_ROW_H, A4_W, ST_ROW_H, stroke=0, fill=1)
        c.setFillColor(GOLD)
        c.setFont("Helvetica-Bold", 11)
        c.drawString(9.9, y - 19.2, "TOTAL")
        for j, key in enumerate(("physical", "allotable", "pending"), start=2):
            c.drawCentredString(_col_x(j) + ST_COLS[j] / 2, y - 19.2,
                                _fmt(totals.get(key, 0), True))

    c.setFillColor(colors.Color(0.45, 0.48, 0.55))
    c.setFont("Helvetica", 8)
    c.drawRightString(A4_W - 9.9, 24.0, f"Page {page_no} of {pages}")
    c.showPage()


def build_stock_pdf(data: dict, out_path: Path, *, title_suffix: str = "") -> Path:
    """Warehouse Stock Report — one page per warehouse, paginated if long."""
    as_of = data["date"]
    try:
        from datetime import date as _d
        as_of = _d.fromisoformat(data["date"]).strftime("%d %b %Y")
    except ValueError:
        pass

    # Work out the page count first so "Page n of m" is right from page one.
    plan = []
    for wh in data["warehouses"]:
        rows = wh["rows"]
        chunks = [rows[i:i + ST_ROWS_PER_PAGE]
                  for i in range(0, len(rows), ST_ROWS_PER_PAGE)] or [[]]
        for k, chunk in enumerate(chunks):
            plan.append((wh, chunk, k == len(chunks) - 1))

    c = pdfcanvas.Canvas(str(out_path), pagesize=(A4_W, A4_H))
    for i, (wh, chunk, is_last) in enumerate(plan, start=1):
        draw_stock_page(
            c, warehouse=wh["name"], as_of=as_of, rows=chunk,
            totals=wh["totals"] if is_last else None,
            page_no=i, pages=len(plan),
        )
    c.save()
    return out_path


# ---------------------------------------------------------------------------
# Daily grid PDF  (Shop Sales - Daily / Secondary Sales - Daily)
#
# Geometry measured off the office's own published PDFs rather than guessed:
# page 695.4 tall for 19 body rows, 45.4 title band, 26.0 sub band, a 54.0
# header split 30/24 between weekday and day number, and 30.0 rows. Colours and
# rule weights are the values read out of those files' content streams.
# ---------------------------------------------------------------------------

D_TITLE_H = 45.4
D_SUB_H   = 26.0
D_SUB_GAP  = 16.0           # the least air allowed between the two sub-band strings
D_SUB_LINE = 17.0           # the extra height when the period needs its own line
D_HEAD_H  = 54.0
D_DOW_H   = 30.0            # weekday half of the header; the rest is the date
D_ROW_H   = 30.0
D_INSET   = 12.0            # text inset on the sub band and on row labels
# Slack around the widest thing in a column, solved off the office's own files:
# their label column runs 17.42pt wider than its longest name, day columns 9.9,
# the total column 9.07. (Their Saturday columns come out 0.9 narrower than this
# reproduces, which is the only place the two differ — under 2pt on the page.)
D_PAD_LABEL = 17.42
D_PAD_DAY   = 9.9
D_PAD_TOTAL = 9.07

DF_TITLE = 18.0
DF_SUB   = 13.0
DF_SUB_MIN = 8.5
DF_HEAD  = 10.0
DF_BODY  = 11.0

D_NAVY   = colors.Color(0.04, 0.17, 0.32)
D_GOLD   = colors.Color(0.98, 0.69, 0.1)
D_GOLD_T = colors.Color(0.98, 0.686, 0.098)   # title and cluster-row text
D_NAVY_T = colors.Color(0.043, 0.173, 0.322)  # sub band and grand-row text
D_INK    = colors.Color(0.059, 0.098, 0.176)  # bond labels and figures
D_ZEBRA  = colors.Color(0.96, 0.97, 0.99)
D_ZERO   = colors.Color(0.784, 0.804, 0.843)  # a nil cell, on white or navy
D_ZERO_G = colors.Color(0.569, 0.482, 0.247)  # a nil cell on the gold total row
D_LW_HEAD = 1.4
D_LW_BODY = 0.4

BOLD = "Helvetica-Bold"
BOOK = "Helvetica"


def _w(text: str, font: str, size: float) -> float:
    return pdfmetrics.stringWidth(str(text), font, size)


def daily_geometry(grid: dict) -> tuple[float, float, float, list[float], float]:
    """(page_w, page_h, label_w, [day widths], total_w) for one grid."""
    rows = grid["rows"]
    days = grid["days"]

    label_w = max([_w(grid.get("group_label", "Bond").upper(), BOLD, DF_HEAD)]
                  + [_w(r["label"], BOLD if r["kind"] != "bond" else BOOK, DF_BODY)
                     for r in rows]) + D_PAD_LABEL

    day_w = []
    for i, d in enumerate(days):
        widest_value = max(_w(str(r['cells'][i]), BOLD, DF_BODY) for r in rows)
        day_w.append(max(_w(d["dow"], BOLD, DF_HEAD),
                         _w(d["dom"], BOLD, DF_HEAD),
                         widest_value) + D_PAD_DAY)

    total_w = max([_w("TOTAL", BOLD, DF_HEAD)]
                  + [_w(str(r['total']), BOLD, DF_BODY) for r in rows]) + D_PAD_TOTAL

    page_w = label_w + sum(day_w) + total_w
    _size, sub_h = _sub_band(grid, page_w)
    page_h = D_TITLE_H + sub_h + D_HEAD_H + D_ROW_H * len(rows)
    return page_w, page_h, label_w, day_w, total_w


def _sub_band(grid: dict, page_w: float) -> tuple[float, float]:
    """The sub band's type size, and how tall the band has to be.

    The band carries the report name on the left and the period on the right,
    and the page is only as wide as the grid needs. Four days of shop sales is
    a narrow page, and there the two strings ran straight through each other -
    'SHOP SALES DAILYugust 2026 - 4 August 2026'. So: shrink until they fit
    side by side, and when even the floor will not do it, give the period a
    line of its own.
    """
    title = grid["title"].upper()
    period = grid["period"]["label"]
    room = page_w - 2 * D_INSET - D_SUB_GAP
    size = DF_SUB
    while size > DF_SUB_MIN:
        if _w(title, BOLD, size) + _w(period, BOLD, size) <= room:
            return size, D_SUB_H
        size -= 0.5
    if _w(title, BOLD, size) + _w(period, BOLD, size) <= room:
        return size, D_SUB_H
    return size, D_SUB_H + D_SUB_LINE


def build_daily_pdf(grid: dict, path: Path) -> Path:
    """Render one bond x day grid to a single page."""
    page_w, page_h, label_w, day_w, total_w = daily_geometry(grid)
    days, rows = grid["days"], grid["rows"]

    c = pdfcanvas.Canvas(str(path), pagesize=(page_w, page_h))
    c.setTitle(f"{grid['title']} — {grid['period']['label']}")

    def box(x, y, w, h, fill, stroke, lw):
        if fill is not None:
            c.setFillColor(fill)
        c.setStrokeColor(stroke)
        c.setLineWidth(lw)
        c.rect(x, y - h, w, h, stroke=1, fill=1 if fill is not None else 0)

    def centred(text, x, w, baseline, font, size, colour):
        c.setFont(font, size)
        c.setFillColor(colour)
        c.drawString(x + (w - _w(text, font, size)) / 2.0, baseline, str(text))

    # --- title band --------------------------------------------------------
    c.setFillColor(D_NAVY)
    c.rect(0, page_h - D_TITLE_H, page_w, D_TITLE_H, stroke=0, fill=1)
    centred("K.S DISTILLERY", 0, page_w, page_h - 29.0, BOLD, DF_TITLE, D_GOLD_T)

    # --- sub band ----------------------------------------------------------
    size, sub_h = _sub_band(grid, page_w)
    title, period = grid["title"].upper(), grid["period"]["label"]
    sub_top = page_h - D_TITLE_H
    c.setFillColor(D_GOLD)
    c.rect(0, sub_top - sub_h, page_w, sub_h, stroke=0, fill=1)
    c.setFont(BOLD, size)
    c.setFillColor(D_NAVY_T)
    if sub_h == D_SUB_H:
        base = sub_top - 17.55
        c.drawString(D_INSET, base, title)
        c.drawString(page_w - D_INSET - _w(period, BOLD, size), base, period)
    else:
        centred(title, 0, page_w, sub_top - 16.2, BOLD, size, D_NAVY_T)
        centred(period, 0, page_w, sub_top - 16.2 - D_SUB_LINE, BOLD, size, D_NAVY_T)

    # --- header ------------------------------------------------------------
    head_top = sub_top - sub_h
    box(0, head_top, label_w, D_HEAD_H, D_NAVY, D_GOLD, D_LW_HEAD)
    centred(grid.get("group_label", "Bond").upper(), 0, label_w,
            head_top - 30.5, BOLD, DF_HEAD, D_GOLD_T)

    x = label_w
    for d, w in zip(days, day_w):
        box(x, head_top, w, D_DOW_H, D_NAVY, D_GOLD, D_LW_HEAD)
        centred(d["dow"], x, w, head_top - 18.5, BOLD, DF_HEAD, D_GOLD_T)
        box(x, head_top - D_DOW_H, w, D_HEAD_H - D_DOW_H, D_NAVY, D_GOLD, D_LW_HEAD)
        centred(d["dom"], x, w, head_top - D_DOW_H - 15.5, BOLD, DF_HEAD, D_GOLD_T)
        x += w
    box(x, head_top, total_w, D_HEAD_H, D_NAVY, D_GOLD, D_LW_HEAD)
    centred("TOTAL", x, total_w, head_top - 30.5, BOLD, DF_HEAD, D_GOLD_T)

    # --- body --------------------------------------------------------------
    y = head_top - D_HEAD_H
    stripe = 0                      # the zebra restarts inside every cluster
    for r in rows:
        if r["kind"] == "cluster":
            fill, ink, zero, font = D_NAVY, D_GOLD_T, D_ZERO, BOLD
            stripe = 0
        elif r["kind"] == "grand":
            fill, ink, zero, font = D_GOLD, D_NAVY_T, D_ZERO_G, BOLD
        else:
            fill = D_ZEBRA if stripe % 2 == 0 else colors.white
            ink, zero, font = D_INK, D_ZERO, BOOK
            stripe += 1

        widths = [label_w] + day_w + [total_w]
        x = 0.0
        for w in widths:
            box(x, y, w, D_ROW_H, fill, D_GOLD, D_LW_BODY)
            x += w

        base = y - 18.85
        c.setFont(font, DF_BODY)
        c.setFillColor(ink)
        c.drawString(D_INSET, base, r["label"])

        x = label_w
        for value, w in zip(r["cells"], day_w):
            centred(str(value), x, w, base, font, DF_BODY, ink if value else zero)
            x += w
        # The period total is always bold, even on a plain bond row.
        centred(str(r['total']), x, total_w, base, BOLD, DF_BODY,
                ink if r["total"] else zero)
        y -= D_ROW_H

    c.showPage()
    c.save()
    return path

# ---------------------------------------------------------------------------
# Shop Sales - Cumulative, current view
# ---------------------------------------------------------------------------
#
# The per-bond books are one page per shop, which is right for handing a bond
# its own file. This is the other thing people want: the screen, printed - the
# bonds you filtered to, each with its shops under it, on as few pages as it
# takes. Same A4 geometry as the stock report so the two sit together in a
# folder without looking like they came from different companies.

CV_ROW_H = 21.0
CV_BOND_H = 23.5
CV_COLS = [243.076, 82.0, 86.0, 96.2, 88.0]
CV_HEADINGS = ["SHOP", "OPENING", "RECEIPT", "SALES", "CLOSING"]
CV_FOOT = 17.7


def _cv_x(i: int) -> float:
    return sum(CV_COLS[:i])


def _cv_rows_per_page(first: bool) -> int:
    top = A4_H - (ST_TITLE_H + ST_BAND_H * 2 if first else ST_BAND_H) - ST_HEAD_H
    return max(1, int((top - (CV_FOOT + 14)) // CV_ROW_H))


def build_cumulative_view_pdf(data: dict, out_path: Path, *, scope: str = "All bonds") -> Path:
    """Print the cumulative screen: the bonds on it, with their shops."""
    period = data.get("period", {})
    flat: list[dict] = []
    for g in data.get("bonds", []):
        flat.append({"kind": "bond", "label": g["bond"].title(),
                     "note": f"{len(g['shops'])} shop{'' if len(g['shops']) == 1 else 's'}",
                     "values": g["totals"]})
        for shop in g["shops"]:
            flat.append({"kind": "shop", "label": shop["name"], "note": shop["code"],
                         "values": shop["values"]})
    if not flat:
        flat = [{"kind": "shop", "label": "Nothing matches those filters.",
                 "note": "", "values": [0, 0, 0, 0]}]

    pages: list[list[dict]] = []
    rest = flat
    while rest:
        take = _cv_rows_per_page(not pages)
        pages.append(rest[:take])
        rest = rest[take:]

    c = pdfcanvas.Canvas(str(out_path), pagesize=(A4_W, A4_H))
    c.setTitle("Shop Sales Cumulative")
    for n, chunk in enumerate(pages):
        y = A4_H
        if n == 0:
            c.setFillColor(NAVY)
            c.rect(0, y - ST_TITLE_H, A4_W, ST_TITLE_H, stroke=0, fill=1)
            c.setFillColor(GOLD)
            c.setFont("Helvetica-Bold", 18)
            c.drawCentredString(A4_W / 2, y - 28.0, "K.S DISTILLERY")
            y -= ST_TITLE_H

            c.setFillColor(GOLD)
            c.rect(0, y - ST_BAND_H, A4_W, ST_BAND_H, stroke=0, fill=1)
            c.setFillColor(NAVY)
            c.setFont("Helvetica-Bold", 11)
            c.drawString(15, y - 15.0, "SHOP SALES CUMULATIVE")
            c.drawRightString(A4_W - 15, y - 15.0, period.get("long", ""))
            y -= ST_BAND_H

        c.setFillColor(GOLD)
        c.rect(0, y - ST_BAND_H, A4_W, ST_BAND_H, stroke=0, fill=1)
        c.setFillColor(NAVY)
        c.setFont("Helvetica-Bold", 11)
        c.drawString(15, y - 15.0, scope)
        c.drawRightString(A4_W - 15, y - 15.0, period.get("short", ""))
        y -= ST_BAND_H

        c.setFillColor(NAVY)
        c.rect(0, y - ST_HEAD_H, A4_W, ST_HEAD_H, stroke=0, fill=1)
        c.setFillColor(GOLD)
        c.setFont("Helvetica-Bold", 9.5)
        for i, head in enumerate(CV_HEADINGS):
            if i == 0:
                c.drawString(15, y - 19.5, head)
            else:
                c.drawCentredString(_cv_x(i) + CV_COLS[i] / 2, y - 19.5, head)
        y -= ST_HEAD_H

        stripe = 0
        for row in chunk:
            bond = row["kind"] == "bond"
            h = CV_BOND_H if bond else CV_ROW_H
            if bond:
                fill, ink, font, size = NAVY, GOLD, "Helvetica-Bold", 10
                stripe = 0
            else:
                fill = colors.Color(0.96, 0.97, 0.99) if stripe % 2 else colors.white
                ink, font, size = ST_INK, "Helvetica", 9
                stripe += 1
            c.setFillColor(fill)
            c.rect(0, y - h, A4_W, h, stroke=0, fill=1)

            base = y - (h / 2 + size * 0.3)
            c.setFillColor(ink)
            c.setFont(font, size)
            label = row["label"]
            room = CV_COLS[0] - 30 - (_w(row["note"], font, size - 1.5) if row["note"] else 0)
            if _w(label, font, size) > room:
                while _w(label + "...", font, size) > room and len(label) > 4:
                    label = label[:-1]
                label += "..."
            c.drawString(15, base, label)
            if row["note"]:
                c.setFont(font, size - 1.5)
                c.setFillColor(ink if bond else ST_ZERO)
                c.drawRightString(CV_COLS[0] - 6, base, row["note"])

            for i, v in enumerate(row["values"], start=1):
                c.setFillColor(ST_ZERO if (not bond and not v) else ink)
                c.setFont(font, size)
                c.drawCentredString(_cv_x(i) + CV_COLS[i] / 2, base, f"{v:,.2f}")
            y -= h

        if n == len(pages) - 1:
            total = data.get("total", [0, 0, 0, 0])
            c.setFillColor(GOLD)
            c.rect(0, y - CV_BOND_H, A4_W, CV_BOND_H, stroke=0, fill=1)
            c.setFillColor(NAVY)
            c.setFont("Helvetica-Bold", 10)
            c.drawString(15, y - (CV_BOND_H / 2 + 3.0), "GRAND TOTAL")
            for i, v in enumerate(total, start=1):
                c.drawCentredString(_cv_x(i) + CV_COLS[i] / 2, y - (CV_BOND_H / 2 + 3.0), f"{v:,.2f}")

        c.setFont("Helvetica", 8)
        c.setFillColor(NAVY)
        c.drawCentredString(A4_W / 2, CV_FOOT, f"Page {n + 1} of {len(pages)}")
        c.showPage()
    c.save()
    return out_path


# ---------------------------------------------------------------------------
# Target vs Achievement
#
# MEASURED, NOT APPROXIMATED
# --------------------------
# Every number below was read out of the content streams of the two PDFs the
# office already circulates - the cluster summary and a single-cluster sheet -
# so this is a reproduction rather than a resemblance:
#
#     page          595.276 wide, 123.25 + 60 per bond (two 30pt rows)
#     title band    42pt navy, "K.S DISTILLERY" Bold 18 gold, centred
#     gold band     28pt, title Bold 13 navy at x=9, "AS ON ..." right at x=9
#     header band   53.25pt navy, labels Bold 7.25 WHITE, 11.75 line pitch,
#                   bracketed by 2pt gold rules and divided by 1.4pt gold
#     columns       bond 104, cat 28, brands share 377.276, total 44, pct 42
#     bond cell     navy, spans the pair; name Bold 10 white (gold on a total)
#     target row    grey 0.95 / gold 0.95,0.70,0.18 on a total row
#     achieved row  white / gold 1,0.74,0.19 on a total row
#     every cell    stroked 0.4pt in 0.78 grey; the pair closes on a 1.5pt navy
#     figures       Helvetica 9 centred; Bold on the total column and on totals
#     percentage    Bold 9.5 red (0.812,0.075,0.133), black on the grand total
# ---------------------------------------------------------------------------

TA_W        = 595.276
TA_TITLE_H  = 42.0
TA_SUB_H    = 28.0
TA_HEAD_H   = 53.25
TA_ROW_H    = 30.0
TA_BOND_W   = 104.0
TA_CAT_W    = 28.0
TA_TOT_W    = 44.0
TA_PCT_W    = 42.0
TA_INSET    = 9.0
TA_WRAP_PAD = 8.0
TA_PITCH    = 11.75

TA_NAVY   = colors.Color(0.04, 0.16, 0.31)
TA_NAVY_T = colors.Color(0.043, 0.161, 0.31)
TA_GOLD   = colors.Color(1.0, 0.74, 0.19)
TA_GOLD_T = colors.Color(1.0, 0.741, 0.192)
TA_GOLD_D = colors.Color(0.95, 0.70, 0.18)      # the target half of a total row
TA_SHADE  = colors.Color(0.95, 0.95, 0.95)      # the target half of a bond row
TA_GRID   = colors.Color(0.78, 0.78, 0.78)
TA_RED    = colors.Color(0.812, 0.075, 0.133)
TA_WHITE  = colors.Color(1.0, 1.0, 1.0)
TA_BLACK  = colors.Color(0.0, 0.0, 0.0)

_MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

TAF_TITLE = 18.0
TAF_SUB   = 13.0
TAF_HEAD  = 7.25
TAF_NAME  = 10.0
TAF_BODY  = 9.0
TAF_PCT   = 9.5

# The office writes these brands short. Their sheet, their spelling.
TA_LABELS = {
    "BCB": "BCB", "BLENDERS": "BLENDERS CHOICE", "CHAIRMANS": "CCB",
    "KS99": "KS.99", "MAGIC": "MAGIC BLEND", "MORNING": "MORNING WALKERS",
    "OLDPEARL": "OLD PEARL", "OLDFORT": "ROYAL OLD FORT", "OTHER": "OTHER",
}


def _ta_wrap(text: str, size: float, max_w: float) -> list[str]:
    """Break a header label on spaces, the way the office's own file breaks."""
    words, lines, cur = str(text).split(), [], ""
    for word in words:
        trial = f"{cur} {word}".strip()
        if cur and _w(trial, BOLD, size) > max_w:
            lines.append(cur)
            cur = word
        else:
            cur = trial
    if cur:
        lines.append(cur)
    return lines or [""]


def _ta_num(v) -> str:
    """Whole cases, no separators - the shape the office's sheet prints."""
    try:
        return f"{int(round(float(v or 0)))}"
    except (TypeError, ValueError):
        return "0"


def _ta_centre(c, x0: float, x1: float, y: float, text: str,
               font: str, size: float) -> None:
    c.setFont(font, size)
    c.drawString((x0 + x1) / 2 - _w(text, font, size) / 2, y, text)


def target_geometry(columns: list, pairs: int) -> tuple[float, float, list[float]]:
    """(page_w, page_h, [x boundaries]) for a sheet of this many bonds."""
    span = TA_W - TA_BOND_W - TA_CAT_W - TA_TOT_W - TA_PCT_W
    each = span / max(1, len(columns))
    xs = [0.0, TA_BOND_W, TA_BOND_W + TA_CAT_W]
    for i in range(len(columns)):
        xs.append(xs[2] + each * (i + 1))
    xs.append(xs[-1] + TA_TOT_W)
    xs.append(TA_W)
    height = TA_TITLE_H + TA_SUB_H + TA_HEAD_H + TA_ROW_H * 2 * max(1, pairs)
    return TA_W, height, xs


def build_target_pdf(data: dict, out_path: Path, *, rows: list | None = None,
                     scope: str = "") -> Path:
    """One Target vs Achievement sheet: a cluster, the summary, or the view."""
    cols = data["columns"]
    rows = data["rows"] if rows is None else rows
    page_w, page_h, xs = target_geometry(cols, len(rows))

    end = data["period"]["to"]
    as_on = f"AS ON {int(end[8:10])} {_MONTH_ABBR[int(end[5:7]) - 1]} {end[:4]}"
    title = "TARGET VS ACHIEVEMENT" + (f" - {scope}" if scope else "")

    c = pdfcanvas.Canvas(str(out_path), pagesize=(page_w, page_h))
    c.setTitle(out_path.stem)

    # ---- the two bands ----
    top = page_h
    c.setFillColor(TA_NAVY)
    c.rect(0, top - TA_TITLE_H, page_w, TA_TITLE_H, stroke=0, fill=1)
    c.setFillColor(TA_GOLD_T)
    _ta_centre(c, 0, page_w, top - 26.0, "K.S DISTILLERY", BOLD, TAF_TITLE)

    sub_top = top - TA_TITLE_H
    c.setFillColor(TA_GOLD)
    c.rect(0, sub_top - TA_SUB_H, page_w, TA_SUB_H, stroke=0, fill=1)
    c.setFillColor(TA_NAVY_T)
    c.setFont(BOLD, TAF_SUB)
    base = sub_top - 18.5
    c.drawString(TA_INSET, base, title)
    c.drawString(page_w - TA_INSET - _w(as_on, BOLD, TAF_SUB), base, as_on)

    # ---- the header band ----
    head_top = sub_top - TA_SUB_H
    head_bot = head_top - TA_HEAD_H
    c.setFillColor(TA_NAVY)
    c.rect(0, head_bot, page_w, TA_HEAD_H, stroke=0, fill=1)

    labels = ["BOND", "CAT"] + [TA_LABELS.get(col["key"], str(col["label"]).upper())
                                for col in cols] + ["GRAND TOTAL", "ACH %"]
    mid = head_bot + TA_HEAD_H / 2 - 2.538
    c.setFillColor(TA_WHITE)
    for i, label in enumerate(labels):
        x0, x1 = xs[i], xs[i + 1]
        if i == 0:                      # BOND is the one label set against the edge
            c.setFont(BOLD, TAF_HEAD)
            c.drawString(TA_INSET, mid, label)
            continue
        lines = _ta_wrap(label, TAF_HEAD, (x1 - x0) - TA_WRAP_PAD)
        first = mid + TA_PITCH * (len(lines) - 1) / 2
        for j, line in enumerate(lines):
            _ta_centre(c, x0, x1, first - TA_PITCH * j, line, BOLD, TAF_HEAD)

    c.setStrokeColor(TA_GOLD)
    c.setLineWidth(2.0)
    c.line(0, head_top - 1.0, page_w, head_top - 1.0)
    c.line(0, head_bot + 1.0, page_w, head_bot + 1.0)
    c.setLineWidth(1.4)
    for x in xs[1:-1]:
        c.line(x, head_top, x, head_bot)

    # ---- the bonds ----
    # Exactly one row wears the gold, and it is the one the sheet adds up to:
    # the cluster on a cluster sheet, the network on the summary. A cluster
    # sitting among its own bonds is only a subtotal, so it takes bold figures
    # and keeps the body's colours.
    has_bonds = any(r.get("kind") == "bond" for r in rows)
    y = head_bot
    for idx, row in enumerate(rows):
        total = idx == len(rows) - 1
        strong = total or (has_bonds and row.get("kind") == "cluster")
        pair_top, pair_bot = y, y - TA_ROW_H * 2
        tgt_fill = TA_GOLD_D if total else TA_SHADE
        ach_fill = TA_GOLD if total else TA_WHITE

        c.setStrokeColor(TA_GRID)
        c.setLineWidth(0.4)
        c.setFillColor(TA_NAVY)
        c.rect(xs[0], pair_bot, xs[1] - xs[0], TA_ROW_H * 2, stroke=1, fill=1)
        for i in range(1, len(xs) - 2):
            c.setFillColor(tgt_fill)
            c.rect(xs[i], pair_top - TA_ROW_H, xs[i + 1] - xs[i], TA_ROW_H, stroke=1, fill=1)
            c.setFillColor(ach_fill)
            c.rect(xs[i], pair_bot, xs[i + 1] - xs[i], TA_ROW_H, stroke=1, fill=1)
        c.setFillColor(ach_fill)
        c.rect(xs[-2], pair_bot, xs[-1] - xs[-2], TA_ROW_H * 2, stroke=1, fill=1)

        # the name, on its navy, gold once the row is a total
        c.setFillColor(TA_GOLD_T if total else TA_WHITE)
        c.setFont(BOLD, TAF_NAME)
        c.drawString(TA_INSET, pair_bot + TA_ROW_H - 3.5, str(row.get("label", "")))

        body = BOLD if strong else BOOK
        for which, tag, line_y in (("tgt", "TGT", pair_top - TA_ROW_H + 11.85),
                                   ("ach", "ACH", pair_bot + 11.85)):
            c.setFillColor(TA_BLACK)
            _ta_centre(c, xs[1], xs[2], line_y, tag, BOLD, TAF_BODY)
            cells = row.get(which) or {}
            for i, col in enumerate(cols):
                _ta_centre(c, xs[2 + i], xs[3 + i], line_y,
                           _ta_num(cells.get(col["key"], 0)), body, TAF_BODY)
            _ta_centre(c, xs[-3], xs[-2], line_y,
                       _ta_num(row.get(which + "_total", 0)), BOLD, TAF_BODY)

        pct = row.get("pct")
        c.setFillColor(TA_BLACK if total else TA_RED)
        _ta_centre(c, xs[-2], xs[-1], pair_bot + TA_ROW_H - 3.325,
                   "-" if pct is None else f"{pct:.2f}%", BOLD, TAF_PCT)

        c.setStrokeColor(TA_NAVY)
        c.setLineWidth(1.5)
        c.line(0, pair_bot, page_w, pair_bot)
        y = pair_bot

    c.showPage()
    c.save()
    return out_path
