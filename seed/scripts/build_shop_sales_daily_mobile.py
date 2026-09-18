#!/usr/bin/env python3
"""
SHOP SALES DAILY (bond × day) — beautified rebuild.

Stays LANDSCAPE: the report gains a column every day, and by month end 31 day
columns cannot fit a portrait page at any readable size. Orientation and column
order are unchanged.

What changes is presentation: a gold grid on the header, cell borders, alternating
row tint, an emphasised TOTAL row, and the page trimmed to the content so nothing
is wasted. Column widths are measured, so the sheet keeps working as days are
added through the month.

Input: daily.json (from extract_daily.py)
"""

import json
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color
from reportlab.pdfbase.pdfmetrics import stringWidth

# ── palette (sampled from the source PDF) ────────────────────────────────────
NAVY  = Color(0.043, 0.161, 0.310)
GOLD  = Color(1.000, 0.741, 0.192)
ZEBRA = Color(0.960, 0.970, 0.990)
WHITE = Color(1, 1, 1)
INK   = Color(0.059, 0.098, 0.176)   # figures
INK_SOFT = Color(0.157, 0.157, 0.157)  # bond names
DIM   = Color(0.784, 0.804, 0.843)   # a zero
RULE  = Color(0.780, 0.780, 0.780)   # cell hairline

TITLE_H = 45.4
BAND_H  = 26.0
HDR_TOP = 30.0       # weekday row
HDR_BOT = 24.0       # day-number row
ROW_H   = 30.0

F_TITLE = 18
F_BAND  = 13
F_HDR   = 10
F_DATA  = 11
PAD     = 5.0
LAB_PAD = 6.0
HAIR    = 0.4
HDR_RULE      = 1.4   # gold grid inside the header block
HDR_RULE_EDGE = 2.0   # gold rules closing the header block
TOT_RULE      = 1.6   # gold rules closing the TOTAL row

MIN_W = 595.276       # never narrower than A4 width


def widths(doc):
    """Measured column widths. The day count grows through the month, so this
    has to be computed every build — a hardcoded grid breaks on day 13."""
    rows, wk, dn = doc["rows"], doc["weekday"], doc["daynum"]
    n = len(wk)
    out = []
    # bond column, left aligned
    w = max(stringWidth(r["name"], "Helvetica", F_DATA) for r in rows)
    w = max(w, stringWidth(wk[0], "Helvetica-Bold", F_HDR))
    out.append(w + LAB_PAD * 3)
    # day columns + TOTAL
    for i in range(1, n):
        w = max(stringWidth(r["values"][i - 1], "Helvetica-Bold", F_DATA) for r in rows)
        w = max(w, stringWidth(wk[i], "Helvetica-Bold", F_HDR),
                   stringWidth(dn[i], "Helvetica-Bold", F_HDR))
        out.append(w + PAD * 2)
    return out


def centred(c, text, font, size, colour, x0, x1, ymid):
    c.setFont(font, size); c.setFillColor(colour)
    c.drawString((x0 + x1) / 2.0 - stringWidth(text, font, size) / 2.0,
                 ymid - size * 0.35, text)


def cell(c, x0, x1, y, h, fill=None, edge=RULE, lw=HAIR):
    c.setStrokeColor(edge); c.setLineWidth(lw)
    if fill is not None:
        c.setFillColor(fill)
    c.rect(x0, y, x1 - x0, h, stroke=1, fill=1 if fill is not None else 0)


def build(doc, path):
    cw = widths(doc)
    PW = max(MIN_W, sum(cw))
    if PW > sum(cw):                       # spread the slack so the grid meets the edge
        extra = (PW - sum(cw)) / len(cw)
        cw = [w + extra for w in cw]
    xs = [0.0]
    for w in cw:
        xs.append(xs[-1] + w)

    rows = doc["rows"]
    HDR_H = HDR_TOP + HDR_BOT
    PH = TITLE_H + BAND_H + HDR_H + ROW_H * len(rows)

    c = canvas.Canvas(path, pagesize=(PW, PH))
    c.setTitle(doc["band_left"])

    # ── masthead ─────────────────────────────────────────────────────────────
    y = PH - TITLE_H
    c.setFillColor(NAVY); c.rect(0, y, PW, TITLE_H, stroke=0, fill=1)
    centred(c, doc["title"], "Helvetica-Bold", F_TITLE, GOLD, 0, PW, y + TITLE_H / 2)

    # ── gold band ────────────────────────────────────────────────────────────
    y -= BAND_H
    c.setFillColor(GOLD); c.rect(0, y, PW, BAND_H, stroke=0, fill=1)
    c.setFont("Helvetica-Bold", F_BAND); c.setFillColor(NAVY)
    c.drawString(12, y + BAND_H / 2 - F_BAND * 0.35, doc["band_left"])
    c.drawRightString(PW - 12, y + BAND_H / 2 - F_BAND * 0.35, doc["band_right"])

    # ── header: weekday over day number; BOND and TOTAL span both ────────────
    y -= HDR_H
    top_y = y + HDR_BOT
    for i, (wk, dn) in enumerate(zip(doc["weekday"], doc["daynum"])):
        if dn:                                    # a day column
            cell(c, xs[i], xs[i + 1], top_y, HDR_TOP, NAVY, GOLD, HDR_RULE)
            cell(c, xs[i], xs[i + 1], y, HDR_BOT, NAVY, GOLD, HDR_RULE)
            centred(c, wk, "Helvetica-Bold", F_HDR, GOLD, xs[i], xs[i + 1], top_y + HDR_TOP / 2)
            centred(c, dn, "Helvetica-Bold", F_HDR, GOLD, xs[i], xs[i + 1], y + HDR_BOT / 2)
        else:                                     # BOND / TOTAL span the block
            cell(c, xs[i], xs[i + 1], y, HDR_H, NAVY, GOLD, HDR_RULE)
            centred(c, wk, "Helvetica-Bold", F_HDR, GOLD, xs[i], xs[i + 1], y + HDR_H / 2)
    c.setStrokeColor(GOLD); c.setLineWidth(HDR_RULE_EDGE); c.setLineCap(0)
    c.line(0, y + HDR_H - HDR_RULE_EDGE / 2, PW, y + HDR_H - HDR_RULE_EDGE / 2)
    c.line(0, y + HDR_RULE_EDGE / 2, PW, y + HDR_RULE_EDGE / 2)

    # ── body ─────────────────────────────────────────────────────────────────
    for r in rows:
        y -= ROW_H
        is_tot = r["total"]
        bg = NAVY if is_tot else (ZEBRA if r["zebra"] else WHITE)

        cell(c, xs[0], xs[1], y, ROW_H, bg)
        c.setFont("Helvetica-Bold" if is_tot else "Helvetica", F_DATA)
        c.setFillColor(GOLD if is_tot else INK_SOFT)
        c.drawString(xs[0] + LAB_PAD * 2, y + ROW_H / 2 - F_DATA * 0.35, r["name"])

        for i, v in enumerate(r["values"]):
            col = i + 1
            cell(c, xs[col], xs[col + 1], y, ROW_H, bg)
            # a zero stays dimmed even on the navy total band — the source does
            # this, and it keeps "no sales" reading the same way on every row
            if r["dim"][i]:
                ink = DIM
            elif is_tot:
                ink = GOLD
            else:
                ink = INK
            last = col == len(cw) - 1          # the TOTAL column
            centred(c, v, "Helvetica-Bold" if (is_tot or last) else "Helvetica",
                    F_DATA, ink, xs[col], xs[col + 1], y + ROW_H / 2)

        if is_tot:                              # gold rules close the total band
            c.setStrokeColor(GOLD); c.setLineWidth(TOT_RULE)
            c.line(0, y + ROW_H - TOT_RULE / 2, PW, y + ROW_H - TOT_RULE / 2)
            c.line(0, y + TOT_RULE / 2, PW, y + TOT_RULE / 2)

    c.showPage(); c.save()
    return PW, PH, len(doc["weekday"])


if __name__ == "__main__":
    doc = json.load(open(sys.argv[1] if len(sys.argv) > 1 else "daily.json"))
    out = sys.argv[2] if len(sys.argv) > 2 else "SHOP SALES DAILY - MOBILE.pdf"
    w, h, n = build(doc, out)
    print(f"{n} columns, {len(doc['rows'])} rows -> page {w:.0f} x {h:.0f}pt")
    print("wrote", out)
