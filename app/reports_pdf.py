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

from decimal import Decimal, ROUND_HALF_UP
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

COL_W = 59.4          # every numeric column, at its roomiest
COL_MIN = 44.0        # and at its narrowest: a wrapped brand head still fits
DOC_W = 453.114       # the page width the office's PDFs use, every document
LABEL_MIN_HARD = 90.0 # never let the label column fall below this
LABEL_MIN = 150.0
PAD_L = 6.0
F_ROW_MIN = 6.8       # a name may shrink this far before it is cut

F_TITLE = 15.0
F_SUB = 9.6
F_PERIOD = 8.6
F_HEAD = 7.0
F_ROW = 8.5
F_TOTAL = 8.8
HEAD_PITCH = 7.6


def _fit(text: str, font: str, size: float, width: float) -> str:
    """`text` shortened to `width`, with an ellipsis when anything was cut."""
    if pdfmetrics.stringWidth(text, font, size) <= width:
        return text
    ell = "\u2026"
    out = text
    while out and pdfmetrics.stringWidth(out + ell, font, size) > width:
        out = out[:-1]
    return (out + ell) if out else text[:1]


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


def trimmed(text: str) -> str:
    """Drop a decimal point that has nothing but zeros after it.

    "5.00" is 5 with noise on the end, and a column of them is noise all the
    way down. Right-to-left, so "1,000.00" loses its two zeros and then its
    point and stops at the nought it needs - it does not become "1,".
    """
    return text.rstrip("0").rstrip(".") if "." in text else text


def whole(v) -> int:
    """Half away from zero, the way the office rounds - never banker's.

    Python's round() and its "%.0f" both round a tie to the even number, so
    448.5 came out 448 where the office's own sheet says 449, and the browser
    and Excel - which both round half up - printed 449 for the very same cell.
    Figures land on a half often enough in these files for that to be a visible
    disagreement between a report and its own PDF.
    """
    return int(Decimal(str(float(v or 0))).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def cases(v, round_off: bool) -> str:
    """A case count as the reports print it.

    Off, up to two places - a shop can be issued half a case and the figure
    should say so - but only where there is something to say. On, whole
    cases, which is what somebody reading the sheet out loud wants.
    """
    n = float(v or 0)
    return f"{whole(n):,}" if round_off else trimmed(f"{n:,.2f}")


def _fmt(v: float, round_off: bool) -> str:
    if not v:
        return "0"
    if round_off or abs(v - whole(v)) < 0.005:
        return f"{whole(v):,}"
    return trimmed(f"{v:,.2f}")


def page_size(labels: list[str], columns: list[str], rows: int) -> tuple[float, float, float]:
    """Kept for callers with a single page shape."""
    return document_size([{"labels": labels, "columns": columns, "rows": rows}])


def label_need(labels: list[str]) -> float:
    """What the longest label would like: its own width, plus its padding."""
    widest = max((pdfmetrics.stringWidth(str(l), "Helvetica", F_ROW)
                  for l in labels), default=0.0)
    return widest + 2 * PAD_L


def columns_width(n_columns: int, wanted: float = 0.0) -> float:
    """Width of one numeric column on a page carrying `n_columns` of them.

    The figures in these columns are three or four digits; the 59.4 measured
    off the office's own PDF is there for the wrapped brand heading above
    them, and a brand heading wraps onto another line quite happily. A shop
    name cannot. So when the labels want more room than the leftovers, the
    numeric columns give some back - down to COL_MIN, which still carries
    MATURED, the longest word any brand puts in a heading.
    """
    if n_columns <= 0:
        return COL_W
    # What a column may be at most: the measured width, or less on a page with
    # so many brands that the label column would otherwise be starved.
    cap = min(COL_W, (DOC_W - LABEL_MIN_HARD) / n_columns)
    # What the labels would like it to be.
    want = (DOC_W - wanted) / n_columns
    return max(min(COL_MIN, cap), min(cap, want))


def label_width_for(width: float, n_columns: int, wanted: float = 0.0) -> float:
    """The label column takes whatever the numeric columns leave."""
    return width - columns_width(n_columns + 1, wanted) * (n_columns + 1)


def document_size(pages: list[dict]) -> tuple[float, float, float]:
    """One page size for the WHOLE document.

    Width is a constant: every PDF the office receives is 453.114pt wide
    whatever the cluster or the brand count, which is what makes them readable
    on a phone. Height is 150.4pt of banding plus 15pt per row, sized for the
    busiest page so the document never changes shape as you scroll.
    """
    max_cols = max(len(p["columns"]) for p in pages)
    max_rows = max(p["rows"] for p in pages)
    wanted = max((label_need(p["labels"]) for p in pages if p["labels"]), default=0.0)
    height = (TITLE_BAND_H + SUB_BAND_H + HEAD_BAND_H
              + ROW_H * max_rows + RULE_H + TOTAL_BAND_H + RULE_H + FOOTER_H)
    return DOC_W, height, label_width_for(DOC_W, max_cols, wanted)


def draw_page(c, *, width: float, height: float, label_w: float,
              report_title: str, period: str, group_line: str,
              label_heading: str, columns: list[str],
              rows: list[dict], totals: dict, total_label: str = "TOTAL",
              page_no: int = 1, pages: int = 1, round_off: bool = False) -> None:
    """Draw one page. `rows` are {'name': str, 'cells': {col: value}, 'total': v}."""
    all_cols = list(columns) + ["TOTAL"]
    # Taken from the label column the caller settled on, not worked out again
    # here: the two used to be computed separately and a page whose labels had
    # claimed extra room drew its headings out of line with its figures.
    col_w = (width - label_w) / len(all_cols)

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
        # Cut with an ellipsis, so a shortened label reads as shortened. Bare
        # truncation produced '2016-KARUNAGAPALLY SOUT' and
        # '2016-KARUNAGAPALLY SOUTH' as the same string, which is two
        # different shops printing under one name with no sign anything was
        # lost.
        # A shop name is the one thing on the page that cannot be looked up
        # anywhere else, so it shrinks before it is cut: down to F_ROW_MIN,
        # and only then to an ellipsis.
        room = label_w - 2 * PAD_L
        raw = str(r["name"])
        size = F_ROW
        while (size > F_ROW_MIN
               and pdfmetrics.stringWidth(raw, "Helvetica", size) > room):
            size -= 0.25
        c.setFont("Helvetica", size)
        c.drawString(PAD_L, base, _fit(raw, "Helvetica", size, room))
        c.setFont("Helvetica", F_ROW)

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

    # The window the screen is showing, said the way the cluster PDFs say it -
    # "1 September 2026 - 21 September 2026", not a pair of ISO stamps.
    def _said(iso: str) -> str:
        from datetime import date as _date
        try:
            d = _date.fromisoformat(str(iso))
        except (TypeError, ValueError):
            return str(iso or "")
        return f"{d.day} {d:%B %Y}"

    span = data.get("span") or {}
    opens = data.get("opens") or {}
    lo = (filters or {}).get("date_from") or opens.get("from") or span.get("min")
    hi = (filters or {}).get("date_to") or opens.get("to") or span.get("max")
    period = f"{_said(lo)} - {_said(hi)}" if lo else ""

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
ST_FOOT = 52.0                 # air under the table, and the page number in it


def stock_page_height(n_rows: int) -> float:
    """A4, or as much taller as this warehouse needs.

    A warehouse used to be cut at eighteen rows and continued overleaf, so
    the long ones arrived as two pages that had to be read together - and
    the first of them ended in a blank space exactly where its total should
    have been. One warehouse is one page now; the page is what stretches.
    """
    need = (ST_TITLE_H + ST_BAND_H * 2 + ST_HEAD_H
            + n_rows * ST_ROW_H + ST_ROW_H + ST_FOOT)
    return max(A4_H, need)
# The grid. A hairline between cells so a long name and its figures stay on
# one line for the eye, and a gold rule closing the two navy bands - the
# header and the total - the way the office's own sheets rule a block.
ST_HAIR = colors.Color(0.855, 0.871, 0.898)
ST_LW_GRID = 0.4
ST_LW_HEAD = 0.9      # the seams between the header's five labels
ST_LW_GOLD = 1.3


def _col_x(i: int) -> float:
    return sum(ST_COLS[:i])


def draw_stock_page(c, *, warehouse: str, as_of: str, rows: list[dict],
                    totals: dict | None, page_no: int, pages: int) -> None:
    """One warehouse, one page - as tall as that warehouse needs it to be."""
    page_h = stock_page_height(len(rows))
    c.setPageSize((A4_W, page_h))
    y = page_h

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

    head_top = y
    c.setFillColor(NAVY)
    c.rect(0, y - ST_HEAD_H, A4_W, ST_HEAD_H, stroke=0, fill=1)
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 10)
    # Every column centred, the first one included - it used to sit hard
    # against the left edge while the four beside it were centred, which read
    # as a label that had been left where it landed.
    c.drawCentredString(ST_COLS[0] / 2, y - 19.0, ST_HEADINGS[0])
    for i in range(1, len(ST_COLS)):
        c.drawCentredString(_col_x(i) + ST_COLS[i] / 2, y - 19.0, ST_HEADINGS[i])
    y -= ST_HEAD_H
    body_top = y

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
        c.drawCentredString(ST_COLS[0] / 2, base, name)
        c.drawCentredString(_col_x(1) + ST_COLS[1] / 2, base, str(r["pack"]))
        for j, key in enumerate(("physical", "allotable", "pending"), start=2):
            v = r.get(key, 0) or 0
            c.setFillColor(ST_ZERO if not v else ST_INK)
            c.drawCentredString(_col_x(j) + ST_COLS[j] / 2, base, _fmt(v, True))
        y -= ST_ROW_H

    body_bottom = y

    total_top = total_bottom = None
    if totals is not None:
        total_top = y
        c.setFillColor(NAVY)
        c.rect(0, y - ST_ROW_H, A4_W, ST_ROW_H, stroke=0, fill=1)
        c.setFillColor(GOLD)
        c.setFont("Helvetica-Bold", 11)
        c.drawCentredString(ST_COLS[0] / 2, y - 19.2, "TOTAL")
        for j, key in enumerate(("physical", "allotable", "pending"), start=2):
            c.drawCentredString(_col_x(j) + ST_COLS[j] / 2, y - 19.2,
                                _fmt(totals.get(key, 0), True))
        total_bottom = y - ST_ROW_H

    # The grid goes on last, over the bands and the zebra, so nothing paints
    # over it. Columns run the whole table; rows only cross the body, where
    # the figures are - a navy band is one piece, not five cells.
    edges = [_col_x(i) for i in range(1, len(ST_COLS))]
    bottom = total_bottom if total_bottom is not None else body_bottom

    c.setLineWidth(ST_LW_GRID)
    c.setStrokeColor(ST_HAIR)
    for x in edges:
        c.line(x, body_top, x, body_bottom)
    n = len(rows)
    for k in range(1, n):
        yy = body_top - k * ST_ROW_H
        c.line(0, yy, A4_W, yy)

    # Inside the header band the same columns carry on, in the gold the labels
    # themselves are set in and a shade heavier than the body grid, so the five
    # headings read as five cells rather than as one long navy strip with words
    # spaced along it. The total band gets none: the header is five labels and
    # the seams tell them apart, but the total is one statement about the
    # warehouse, and ruling it into cells only put stray lines across a band
    # that should read solid.
    c.setLineWidth(ST_LW_HEAD)
    c.setStrokeColor(GOLD)
    for x in edges:
        c.line(x, head_top, x, body_top)

    # And the gold lining: a rule above and below each navy band.
    c.setLineWidth(ST_LW_GOLD)
    c.setStrokeColor(GOLD)
    for yy in (head_top, body_top):
        c.line(0, yy, A4_W, yy)
    if total_top is not None:
        for yy in (total_top, total_bottom):
            c.line(0, yy, A4_W, yy)
    else:
        # A continuation page closes on a hairline; the total comes later.
        c.setLineWidth(ST_LW_GRID)
        c.setStrokeColor(ST_HAIR)
        c.line(0, bottom, A4_W, bottom)

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

    # One page per warehouse, so the count is simply how many there are.
    houses = list(data["warehouses"])

    # The screen and the workbook both end on a network total; the PDF stopped
    # at the last warehouse, so the one figure a reader most often wants was
    # the only one they had to add up by hand. A closing page carries each
    # warehouse's totals and the network total under them - only when there is
    # more than one warehouse, since for a single warehouse its own total page
    # already IS the network total.
    pages = len(houses) + (1 if len(houses) > 1 else 0)

    c = pdfcanvas.Canvas(str(out_path), pagesize=(A4_W, A4_H))
    for i, wh in enumerate(houses, start=1):
        draw_stock_page(
            c, warehouse=wh["name"], as_of=as_of, rows=wh["rows"],
            totals=wh["totals"], page_no=i, pages=pages,
        )
    if len(houses) > 1:
        summary = [{"brand": wh["name"], "pack": "",
                    "physical": wh["totals"]["physical"],
                    "allotable": wh["totals"]["allotable"],
                    "pending": wh["totals"]["pending"]}
                   for wh in houses]
        draw_stock_page(
            c, warehouse="ALL WAREHOUSES", as_of=as_of, rows=summary,
            totals=data.get("grand"), page_no=pages, pages=pages,
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


def build_daily_pdf(grid: dict, path: Path, *, round_off: bool = False) -> Path:
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
            centred(cases(value, round_off), x, w, base, font, DF_BODY,
                    ink if value else zero)
            x += w
        # The period total is always bold, even on a plain bond row.
        centred(cases(r["total"], round_off), x, total_w, base, BOLD, DF_BODY,
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


def build_cumulative_view_pdf(data: dict, out_path: Path, *, scope: str = "All bonds",
                              round_off: bool = False) -> Path:
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
                c.drawCentredString(_cv_x(i) + CV_COLS[i] / 2, base, cases(v, round_off))
            y -= h

        if n == len(pages) - 1:
            total = data.get("total", [0, 0, 0, 0])
            c.setFillColor(GOLD)
            c.rect(0, y - CV_BOND_H, A4_W, CV_BOND_H, stroke=0, fill=1)
            c.setFillColor(NAVY)
            c.setFont("Helvetica-Bold", 10)
            c.drawString(15, y - (CV_BOND_H / 2 + 3.0), "GRAND TOTAL")
            for i, v in enumerate(total, start=1):
                c.drawCentredString(_cv_x(i) + CV_COLS[i] / 2, y - (CV_BOND_H / 2 + 3.0),
                                    cases(v, round_off))

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
#     bond cell     navy, spans the pair; name Bold 10 white (gold on a total).
#                   The office sets BOND and the names against the left edge;
#                   here they are centred, which is the one place this sheet
#                   departs from theirs, and it was asked for.
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
TA_SEAM   = colors.Color(0.133, 0.251, 0.498)   # the grid inside a navy band
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


def _ta_num(v, round_off: bool = True) -> str:
    """Cases, no separators - the shape the office's sheet prints.

    Whole when the sheet is rounded off, two places when it is not: a target
    is always a whole case, but what a bond actually sold need not be.
    """
    try:
        n = float(v or 0)
    except (TypeError, ValueError):
        return "0"
    return f"{whole(n)}" if round_off else trimmed(f"{n:.2f}")


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
                     scope: str = "", round_off: bool = True) -> Path:
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
    # Three tiers, as every other sheet in the office reads: a bond on the
    # body's colours, a cluster subtotal on navy with gold figures, and the one
    # row the sheet adds up to in gold. A cluster used to differ from the bonds
    # above it by bold figures alone, which is no difference at all halfway
    # down a page of numbers.
    has_bonds = any(r.get("kind") == "bond" for r in rows)
    y = head_bot
    for idx, row in enumerate(rows):
        total = idx == len(rows) - 1
        cluster = not total and has_bonds and row.get("kind") == "cluster"
        strong = total or cluster
        pair_top, pair_bot = y, y - TA_ROW_H * 2
        tgt_fill = TA_GOLD_D if total else TA_NAVY if cluster else TA_SHADE
        ach_fill = TA_GOLD if total else TA_NAVY if cluster else TA_WHITE
        ink = TA_GOLD_T if cluster else TA_BLACK

        c.setStrokeColor(TA_SEAM if cluster else TA_GRID)
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

        # the name, centred on its navy, gold once the row is a total
        c.setFillColor(TA_GOLD_T if total or cluster else TA_WHITE)
        _ta_centre(c, xs[0], xs[1], pair_bot + TA_ROW_H - 3.5,
                   str(row.get("label", "")), BOLD, TAF_NAME)

        body = BOLD if strong else BOOK
        for which, tag, line_y in (("tgt", "TGT", pair_top - TA_ROW_H + 11.85),
                                   ("ach", "ACH", pair_bot + 11.85)):
            c.setFillColor(ink)
            _ta_centre(c, xs[1], xs[2], line_y, tag, BOLD, TAF_BODY)
            cells = row.get(which) or {}
            for i, col in enumerate(cols):
                _ta_centre(c, xs[2 + i], xs[3 + i], line_y,
                           _ta_num(cells.get(col["key"], 0), round_off),
                           body, TAF_BODY)
            _ta_centre(c, xs[-3], xs[-2], line_y,
                       _ta_num(row.get(which + "_total", 0), round_off),
                       BOLD, TAF_BODY)

        pct = row.get("pct")
        c.setFillColor(TA_BLACK if total else TA_GOLD_T if cluster else TA_RED)
        _ta_centre(c, xs[-2], xs[-1], pair_bot + TA_ROW_H - 3.325,
                   "-" if pct is None else f"{pct:.2f}%", BOLD, TAF_PCT)

        c.setStrokeColor(TA_NAVY)
        c.setLineWidth(1.5)
        c.line(0, pair_bot, page_w, pair_bot)
        y = pair_bot

    c.showPage()
    c.save()
    return out_path


# ---------------------------------------------------------------------------
# Purchase instruction - one sheet per bond
#
# MEASURED, NOT APPROXIMATED
# --------------------------
# Read out of the content stream of the Aluva sheet the office already uses:
#
#     page          A4 portrait, 595.28 x 841.89
#     title band    85pt navy; house name Bold 10 gold at (20, 819.89),
#                   bond Bold 22 white at 793.89, the line under it Book 9 in
#                   0.706,0.765,0.843 at 777.89, the count Bold 9 gold at
#                   765.89, closed by a 3pt gold rule
#     block bar     24pt navy from x=20, width 555.28, with a 4pt gold tab;
#                   title Bold 11 white at x=32, right label Bold 9 gold
#     caption       14pt on 0.94,0.95,0.97, Oblique 7.5 in 0.392,0.431,0.49
#     head row      18pt navy to x=485, gold from there; labels Bold 8.5
#     data row      18pt, white over 0.96,0.97,0.99, the MQ block on cream
#     total row     20pt on 0.9,0.93,0.97, Bold 9
#     columns       brand at x=30; figures right-aligned at 314.78, 395.1,
#                   475.1, 565.4
#     a nil         prints as a full stop, not a zero and not a dash
# ---------------------------------------------------------------------------

PB_W, PB_H = 595.28, 841.89
PB_L, PB_R = 20.0, 575.28
PB_TITLE_H = 85.0
PB_RULE_H = 3.0
PB_BAR_H = 24.0
PB_TAB_W = 4.0
PB_CAP_H = 14.0
PB_HEAD_H = 18.0
PB_ROW_H = 18.0
PB_TOT_H = 20.0
PB_MQ_X = 485.0
PB_BOT = 34.0
PB_CONT_H = 34.0

PB_NAVY = colors.Color(0.04, 0.17, 0.32)
PB_NAVY_T = colors.Color(0.043, 0.173, 0.322)
PB_GOLD = colors.Color(0.98, 0.69, 0.10)
PB_GOLD_T = colors.Color(0.98, 0.686, 0.098)
PB_SUB = colors.Color(0.706, 0.765, 0.843)
PB_CAP_BG = colors.Color(0.94, 0.95, 0.97)
PB_CAP_T = colors.Color(0.392, 0.431, 0.49)
PB_INK = colors.Color(0.118, 0.137, 0.176)
PB_ZEBRA = colors.Color(0.96, 0.97, 0.99)
PB_CREAM = colors.Color(1.0, 0.98, 0.90)
PB_TOTBG = colors.Color(0.90, 0.93, 0.97)
PB_TOTMQ = colors.Color(1.0, 0.92, 0.70)
PB_WHITE = colors.Color(1, 1, 1)

# right edges, measured
PB_COLS = [("l3ms", 314.78), ("rl", 395.10), ("rq", 475.10), ("mq", 565.40)]
PB_LABELS = {"l3ms": "L3MS", "rl": "RL", "rq": "RQ", "mq": "MQ"}


def _pb_num(value) -> str:
    """A nil is a full stop here. The office's sheet prints it that way, and a
    page of zeros reads as data when it is the absence of any."""
    n = int(value or 0)
    if not n:
        return "."
    # Indian grouping, the way every other figure in this app is written.
    s = str(abs(n))
    if len(s) > 3:
        head, tail = s[:-3], s[-3:]
        parts = []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        s = ",".join(parts) + "," + tail
    return ("-" if n < 0 else "") + s


def _pb_right(c, x: float, y: float, text: str, font: str, size: float) -> None:
    c.setFont(font, size)
    c.drawString(x - _w(text, font, size), y, text)


def _pb_block(c, y: float, title: str, right: str, rows: list, totals: dict,
              caption: str = "") -> float:
    """One bar, its table, its total. Returns the y it finished at."""
    c.setFillColor(PB_NAVY)
    c.rect(PB_L, y - PB_BAR_H, PB_R - PB_L, PB_BAR_H, stroke=0, fill=1)
    c.setFillColor(PB_GOLD)
    c.rect(PB_L, y - PB_BAR_H, PB_TAB_W, PB_BAR_H, stroke=0, fill=1)
    c.setFillColor(PB_WHITE)
    c.setFont(BOLD, 11)
    c.drawString(32.0, y - 16.0, title)
    if right:
        c.setFillColor(PB_GOLD_T)
        _pb_right(c, PB_R, y - 16.0, right, BOLD, 9)
    y -= PB_BAR_H

    if caption:
        c.setFillColor(PB_CAP_BG)
        c.rect(PB_L, y - PB_CAP_H, PB_R - PB_L, PB_CAP_H, stroke=0, fill=1)
        c.setFillColor(PB_CAP_T)
        c.setFont("Helvetica-Oblique", 7.5)
        c.drawString(32.0, y - 10.0, caption)
        y -= PB_CAP_H

    c.setFillColor(PB_NAVY)
    c.rect(PB_L, y - PB_HEAD_H, PB_MQ_X - PB_L, PB_HEAD_H, stroke=0, fill=1)
    c.setFillColor(PB_GOLD)
    c.rect(PB_MQ_X, y - PB_HEAD_H, PB_R - PB_MQ_X, PB_HEAD_H, stroke=0, fill=1)
    c.setFillColor(PB_GOLD_T)
    c.setFont(BOLD, 8.5)
    c.drawString(30.0, y - 12.0, "BRAND")
    for key, edge in PB_COLS:
        c.setFillColor(PB_NAVY_T if key == "mq" else PB_GOLD_T)
        _pb_right(c, edge, y - 12.0, PB_LABELS[key], BOLD, 8.5)
    y -= PB_HEAD_H

    for i, row in enumerate(rows):
        c.setFillColor(PB_WHITE if i % 2 == 0 else PB_ZEBRA)
        c.rect(PB_L, y - PB_ROW_H, PB_MQ_X - PB_L, PB_ROW_H, stroke=0, fill=1)
        c.setFillColor(PB_CREAM)
        c.rect(PB_MQ_X, y - PB_ROW_H, PB_R - PB_MQ_X, PB_ROW_H, stroke=0, fill=1)
        c.setFillColor(PB_INK)
        c.setFont(BOOK, 8.5)
        c.drawString(30.0, y - 12.0, str(row["label"]))
        for key, edge in PB_COLS:
            bold = key == "mq"
            c.setFillColor(PB_NAVY_T if bold else PB_INK)
            _pb_right(c, edge, y - 12.0, _pb_num(row.get(key)),
                      BOLD if bold else BOOK, 8.5)
        y -= PB_ROW_H

    c.setFillColor(PB_TOTBG)
    c.rect(PB_L, y - PB_TOT_H, PB_MQ_X - PB_L, PB_TOT_H, stroke=0, fill=1)
    c.setFillColor(PB_TOTMQ)
    c.rect(PB_MQ_X, y - PB_TOT_H, PB_R - PB_MQ_X, PB_TOT_H, stroke=0, fill=1)
    c.setFillColor(PB_NAVY_T)
    c.setFont(BOLD, 9)
    c.drawString(30.0, y - 13.0, "TOTAL")
    for key, edge in PB_COLS:
        _pb_right(c, edge, y - 13.0, _pb_num(totals.get(key)), BOLD, 9)
    return y - PB_TOT_H


def _pb_block_height(rows: int, caption: bool) -> float:
    return (PB_BAR_H + (PB_CAP_H if caption else 0.0)
            + PB_HEAD_H + rows * PB_ROW_H + PB_TOT_H)


def _pb_title(c, group: str, month_label: str, line: str, first: bool) -> float:
    """The band at the top. Full on page one, a thin strip after it."""
    if first:
        c.setFillColor(PB_NAVY)
        c.rect(0, PB_H - PB_TITLE_H, PB_W, PB_TITLE_H, stroke=0, fill=1)
        c.setFillColor(PB_GOLD_T)
        c.setFont(BOLD, 10)
        c.drawString(PB_L, PB_H - 22.0, "K.S. DISTILLERY")
        c.setFillColor(PB_WHITE)
        c.setFont(BOLD, 22)
        c.drawString(PB_L, PB_H - 48.0, group)
        c.setFillColor(PB_SUB)
        c.setFont(BOOK, 9)
        c.drawString(PB_L, PB_H - 64.0, f"Purchase Instruction  ·  {month_label.upper()}")
        c.setFillColor(PB_GOLD_T)
        c.setFont(BOLD, 9)
        c.drawString(PB_L, PB_H - 76.0, line)
        c.setFillColor(PB_GOLD)
        c.rect(0, PB_H - PB_TITLE_H - PB_RULE_H, PB_W, PB_RULE_H, stroke=0, fill=1)
        return PB_H - PB_TITLE_H - PB_RULE_H - 13.0

    c.setFillColor(PB_NAVY)
    c.rect(0, PB_H - PB_CONT_H, PB_W, PB_CONT_H, stroke=0, fill=1)
    c.setFillColor(PB_WHITE)
    c.setFont(BOLD, 11)
    c.drawString(PB_L, PB_H - 22.0, group)
    c.setFillColor(PB_GOLD_T)
    c.setFont(BOLD, 8.5)
    _pb_right(c, PB_R, PB_H - 22.0, f"Purchase Instruction · {month_label.upper()}",
              BOLD, 8.5)
    c.setFillColor(PB_GOLD)
    c.rect(0, PB_H - PB_CONT_H - 2.0, PB_W, 2.0, stroke=0, fill=1)
    return PB_H - PB_CONT_H - 2.0 - 13.0


# The bond sheet prints a brand by its trade name, not by the full SKU line:
# "OLD PEARL", not "OLD PEARL NO.1 MATURED XXX RUM". Eight rows, always, in
# this order - a shop that bought one brand still shows the other seven as
# nil, because the question the sheet answers is what was indented against
# everything on offer.
PB_BRAND: dict[str, str] = {
    "1139703": "BCB",
    "1139707": "BLENDER'S CHOICE",
    "1139715": "CHAIRMAN'S CHOICE",
    "1339710": "K.S 99 LIFE TIME MATURED",
    "1339718": "MAGIC BLEND RESERVED",
    "1139708": "MORNING WALKERS",
    "1339703": "OLD PEARL",
    "1339704": "ROYAL OLD FORT",
}


def build_pi_group_pdf(group: str, month_label: str, brands: list,
                       total_row: dict, shops: list, out_path: Path) -> Path:
    """One bond's (or warehouse's) instruction: its total, then every shop.

    `brands` is [{key, label}]; `total_row` and each shop carry a cells dict
    keyed by brand and a totals dict.
    """
    c = pdfcanvas.Canvas(str(out_path), pagesize=(PB_W, PB_H))
    c.setTitle(f"{group} - Purchase Instruction {month_label}")

    def on_indent(cells: dict) -> int:
        """Brands this sheet actually instructs a buy of.

        The instruction is RL, RQ and MQ; L3MS is what the shop sold over the
        last three months and is history, not an indent. So a brand with a
        year of sales behind it and nothing asked for this month does not
        count - which is the difference between 'eight brands' on every block
        and a number worth reading.
        """
        return sum(1 for b in brands
                   if any((cells.get(b["key"]) or {}).get(k)
                          for k, _ in PB_COLS if k != "l3ms"))

    def brand_chip(cells: dict) -> str:
        n = on_indent(cells)
        return f"{n} brand{'' if n == 1 else 's'}"

    line = (f"{len(shops)} shop{'' if len(shops) == 1 else 's'}  \u00b7  "
            f"{brand_chip(total_row.get('cells') or {})} on indent")
    y = _pb_title(c, group, month_label, line, first=True)

    def rows_for(cells):
        return [{"label": b["label"], **{k: (cells.get(b["key"]) or {}).get(k, 0)
                                         for k, _ in PB_COLS}} for b in brands]

    y = _pb_block(c, y, f"{group} TOTAL",
                  f"{len(shops)} shop{'' if len(shops) == 1 else 's'}",
                  rows_for(total_row.get("cells") or {}),
                  total_row.get("total") or {},
                  caption="Sum of every shop below. L3MS in bottles · "
                          "RL / RQ / MQ in cases.")
    y -= 12.0

    for shop in shops:
        need = _pb_block_height(len(brands), caption=False)
        if y - need < PB_BOT:
            c.showPage()
            y = _pb_title(c, group, month_label, line, first=False)
        y = _pb_block(c, y, str(shop.get("label", "")),
                      brand_chip(shop.get("cells") or {}),
                      rows_for(shop.get("cells") or {}),
                      shop.get("total") or {})
        y -= 12.0

    c.showPage()
    c.save()
    return out_path


# ---------------------------------------------------------------------------
# Secondary Sales - Analysis
#
# The on-screen sheet, one page: this pull against the one before it, the
# difference, and last month taken whole. House bands and tiers - navy title,
# gold period band, navy header, warehouses on the body with the comparison
# block on cream, a cluster on gold, the total on navy with gold figures - so
# it reads like every other sheet the office forwards. A4 landscape wide,
# as tall as the rows it has.
# ---------------------------------------------------------------------------

II_W = 841.89
II_INSET = 12.0
II_TITLE_H = 34.0
II_SUB_H = 26.0
II_GROUP_H = 20.0
II_HEAD_H = 18.0
II_ROW_H = 15.0
II_FOOT_GAP = 6.0
II_NAME_W = 118.0
II_TAIL = 24.0

II_CREAM = colors.Color(1.0, 0.988, 0.941)
II_ZEBRA = colors.Color(0.965, 0.972, 0.988)
II_LINE = colors.Color(0.80, 0.82, 0.86)
II_INK = colors.Color(0.106, 0.165, 0.29)
II_UP = colors.Color(0.106, 0.498, 0.231)
II_DOWN = colors.Color(0.706, 0.137, 0.094)
II_WHITE = colors.Color(1, 1, 1)
II_NOTE = colors.Color(0.40, 0.44, 0.52)


def _ii_day(iso: str, long: bool = False) -> str:
    if not iso:
        return ""
    y, m, d = int(iso[:4]), int(iso[5:7]), int(iso[8:10])
    if long:
        months = ["January", "February", "March", "April", "May", "June", "July",
                  "August", "September", "October", "November", "December"]
        return f"{d} {months[m - 1]} {y}"
    return f"{d}{_MONTH_ABBR[m - 1].upper()}"


def _ii_num(v, round_off: bool) -> str:
    if v is None:
        return "-"
    return _fmt(float(v), round_off) if v else "0"


def build_item_issue_pdf(data: dict, out_path: Path, *, round_off: bool = True,
                         scope: str = "") -> Path:
    rows = data["rows"]
    ds = data.get("day_sale") or {}
    ind = data.get("industry") or {}
    has_ind = bool(ind.get("cases") or ind.get("prior"))
    foot_n = 1 + (1 if has_ind else 0)

    # sixteen columns: name, two blocks of six, difference (cases, %), last month
    n_fig = 15
    # edge to edge, like the bands above it
    fig_w = (II_W - II_NAME_W) / n_fig
    xs = [0.0, II_NAME_W]
    for i in range(n_fig):
        xs.append(xs[-1] + fig_w)

    body_h = II_ROW_H * (len(rows) + foot_n) + II_FOOT_GAP
    page_h = II_TITLE_H + II_SUB_H + II_GROUP_H + II_HEAD_H + body_h + II_TAIL
    c = pdfcanvas.Canvas(str(out_path), pagesize=(II_W, page_h))
    c.setTitle(out_path.stem)

    def centre(x0, x1, y, text, font, size):
        c.setFont(font, size)
        c.drawString((x0 + x1) / 2 - _w(text, font, size) / 2, y, text)

    # ---- title and period bands ----
    top = page_h
    c.setFillColor(NAVY)
    c.rect(0, top - II_TITLE_H, II_W, II_TITLE_H, stroke=0, fill=1)
    c.setFillColor(GOLD)
    centre(0, II_W, top - 22.5, "K.S DISTILLERY", BOLD, F_TITLE)

    sub_top = top - II_TITLE_H
    c.setFillColor(GOLD)
    c.rect(0, sub_top - II_SUB_H, II_W, II_SUB_H, stroke=0, fill=1)
    c.setFillColor(NAVY)
    c.setFont(BOLD, 10.0)
    title = "SECONDARY SALES - ANALYSIS" + (f" - {scope.upper()}" if scope else "")
    base = sub_top - 16.5
    c.drawString(II_INSET, base, title)
    as_on = f"AS ON {_ii_day(data.get('as_on', ''), long=True).upper()}"
    c.drawString(II_W - II_INSET - _w(as_on, BOLD, 10.0), base, as_on)

    # ---- the two header rows ----
    g_top = sub_top - II_SUB_H
    g_bot = g_top - II_GROUP_H
    h_bot = g_bot - II_HEAD_H
    c.setFillColor(NAVY)
    c.rect(0, h_bot, II_W, II_GROUP_H + II_HEAD_H, stroke=0, fill=1)

    def month_of(iso):
        if not iso:
            return ""
        months = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY",
                  "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]
        return f"{months[int(iso[5:7]) - 1]} {iso[:4]}"

    cur_head = f"{month_of(data.get('as_on'))} - AS ON {_ii_day(data.get('as_on', ''), long=True).upper()}"
    pri_head = (f"{month_of(data.get('prior_as_on'))} - AS ON "
                f"{_ii_day(data.get('prior_as_on', ''), long=True).upper()}"
                if data.get("prior_as_on") else "NO PRIOR PULL")
    groups = [(1, 7, cur_head), (7, 13, pri_head), (13, 15, "DIFFERENCE")]
    c.setFillColor(GOLD)
    for a, b, text in groups:
        centre(xs[a], xs[b], g_bot + 6.5, _fit(text, BOLD, 7.6, xs[b] - xs[a] - 6), BOLD, 7.6)
    last_word = "LAST MONTH"
    last_sub = f"({_ii_day(data['last_key'].split('_')[1])})" if data.get("last_key") else ""
    # spans both header rows, so it is centred on the two of them together
    mid = h_bot + (II_GROUP_H + II_HEAD_H) / 2
    c.setFillColor(GOLD)
    if last_sub:
        centre(xs[15], xs[16], mid + 1.5, last_word, BOLD, 7.0)
        centre(xs[15], xs[16], mid - 7.5, last_sub, BOLD, 7.0)
    else:
        centre(xs[15], xs[16], mid - 2.5, last_word, BOLD, 7.0)
    centre(xs[0], xs[1], h_bot + (II_GROUP_H + II_HEAD_H) / 2 - 2.5, "WAREHOUSE", BOLD, 8.0)

    labels = (["STN", "GTN", "TOTAL", "C FED", "BAR", _ii_day(data.get("as_on", ""))]
              + ["STN", "GTN", "TOTAL", "C FED", "BAR", _ii_day(data.get("prior_as_on", "")) or "-"]
              + ["CASES", "%"])
    c.setFillColor(II_WHITE)
    for i, label in enumerate(labels):
        centre(xs[1 + i], xs[2 + i], h_bot + 6.0, label, BOLD, 7.0)

    c.setStrokeColor(RULE_HEAD)
    c.setLineWidth(0.5)
    for i in (1, 7, 13, 15):
        c.line(xs[i], g_top, xs[i], h_bot)
    c.line(xs[1], g_bot, xs[15], g_bot)
    c.setFillColor(GOLD)
    c.setStrokeColor(GOLD)
    c.setLineWidth(1.4)
    c.line(0, h_bot, II_W, h_bot)

    # ---- the rows ----
    y = h_bot
    zebra = False
    for row in rows:
        kind = row.get("kind")
        top_y, bot_y = y, y - II_ROW_H
        if kind == "grand":
            fill_main = fill_prior = NAVY
            ink = GOLD
        elif kind == "cluster":
            fill_main = fill_prior = GOLD
            ink = NAVY
        else:
            zebra = not zebra
            fill_main = II_ZEBRA if zebra else II_WHITE
            fill_prior = II_CREAM
            ink = II_INK
        c.setFillColor(fill_main)
        c.rect(xs[0], bot_y, xs[7] - xs[0], II_ROW_H, stroke=0, fill=1)
        c.rect(xs[13], bot_y, xs[15] - xs[13], II_ROW_H, stroke=0, fill=1)
        c.setFillColor(fill_prior)
        c.rect(xs[7], bot_y, xs[13] - xs[7], II_ROW_H, stroke=0, fill=1)
        c.rect(xs[15], bot_y, xs[16] - xs[15], II_ROW_H, stroke=0, fill=1)

        strong = kind in ("grand", "cluster")
        base_y = bot_y + 4.6
        c.setFillColor(ink)
        c.setFont(BOLD, 8.2)
        c.drawString(xs[0] + II_INSET, base_y,
                     _fit(str(row.get("label", "")), BOLD, 8.2, II_NAME_W - II_INSET - 4))

        cur, pri = row.get("cur") or {}, row.get("prior") or {}
        figs = ([cur.get(k) for k in ("stn", "gtn", "total", "cfed", "bar", "all")]
                + [pri.get(k) for k in ("stn", "gtn", "total", "cfed", "bar", "all")])
        for i, v in enumerate(figs):
            is_run = i in (5, 11)
            font = BOLD if (strong or is_run) else BOOK
            if not strong and not v:
                c.setFillColor(ZERO_GREY)
            else:
                c.setFillColor(ink)
            centre(xs[1 + i], xs[2 + i], base_y, _ii_num(v, round_off), font, 8.2)

        diff, pct = row.get("diff"), row.get("pct")
        tone = ink if strong else (II_UP if (diff or 0) > 0 else II_DOWN if (diff or 0) < 0 else II_INK)
        c.setFillColor(tone)
        centre(xs[13], xs[14], base_y, _ii_num(diff, round_off), BOLD, 8.2)
        centre(xs[14], xs[15], base_y,
               "-" if pct is None else f"{pct:.{1 if round_off else 2}f}%", BOLD, 8.2)
        c.setFillColor(ink)
        centre(xs[15], xs[16], base_y, _ii_num(row.get("last_month"), round_off),
               BOLD if strong else BOOK, 8.2)

        c.setStrokeColor(II_LINE if not strong else (GOLD if kind == "grand" else NAVY))
        c.setLineWidth(0.3 if not strong else 0.8)
        c.line(xs[0], bot_y, xs[16], bot_y)
        y = bot_y

    # separators down the body
    c.setStrokeColor(II_LINE)
    c.setLineWidth(0.4)
    for i in (1, 7, 13, 15):
        c.line(xs[i], h_bot, xs[i], y)

    # ---- day sale and industry: not in the export, so set apart under it ----
    y -= II_FOOT_GAP
    feet = [("DAY SALE", ds.get("cur"), ds.get("prior"), ds.get("diff"), ds.get("pct"))]
    if has_ind:
        a, b = ind.get("cases") or 0, ind.get("prior") or 0
        feet.append(("INDUSTRY TOTAL", a, b, a - b, ((a - b) / b * 100) if b else None))
    for label, a, b, dff, pc in feet:
        top_y, bot_y = y, y - II_ROW_H
        c.setFillColor(II_ZEBRA)
        c.rect(xs[0], bot_y, xs[16] - xs[0], II_ROW_H, stroke=0, fill=1)
        base_y = bot_y + 4.6
        c.setFillColor(II_INK)
        c.setFont(BOLD, 8.2)
        c.drawString(xs[0] + II_INSET, base_y, label)
        centre(xs[1], xs[7], base_y, _ii_num(a, round_off), BOLD, 8.2)
        centre(xs[7], xs[13], base_y, _ii_num(b, round_off), BOLD, 8.2)
        c.setFillColor(II_UP if (dff or 0) > 0 else II_DOWN if (dff or 0) < 0 else II_INK)
        centre(xs[13], xs[14], base_y, _ii_num(dff, round_off), BOLD, 8.2)
        centre(xs[14], xs[15], base_y, "-" if pc is None else f"{pc:.{1 if round_off else 2}f}%", BOLD, 8.2)
        c.setStrokeColor(II_LINE)
        c.setLineWidth(0.3)
        c.line(xs[0], bot_y, xs[16], bot_y)
        y = bot_y

    c.setFillColor(II_NOTE)
    c.setFont(BOOK, 6.8)
    period = str(data.get("period_label") or "")
    prior = str(data.get("prior_label") or "")
    note = f"Period {period}" + (f" against {prior}" if prior else "") + "  ·  cases"
    c.drawString(II_INSET, 9.0, note)

    c.showPage()
    c.save()
    return out_path
