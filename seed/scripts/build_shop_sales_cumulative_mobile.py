#!/usr/bin/env python3
"""
SHOP SALES CUMULATIVE (per shop, brand × pack) — mobile-friendly rebuild.

The original is already A4 portrait, so orientation is not the problem here.
What makes it hard on a phone is that 5 columns are spread across the full 595pt
width — the BRAND/PACK column alone is ~90pt wider than its longest label — so at
fit-width the type renders far smaller than it needs to.

This rebuild keeps EVERY row, every value, the same fonts, colours, zebra and row
heights, and simply narrows the page to the width the content actually needs.
Same A4 height, so nothing reflows vertically and the pagination is unchanged.

Input: cumulative.json (from extract_cumulative.py)
"""

import json
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color
from reportlab.pdfbase.pdfmetrics import stringWidth

# ── palette (sampled from the original PDF) ──────────────────────────────────
NAVY   = Color(0.043, 0.161, 0.310)
GOLD   = Color(1.000, 0.741, 0.192)
ZEBRA  = Color(0.960, 0.970, 0.990)
WHITE  = Color(1, 1, 1)
INK    = Color(0.314, 0.314, 0.314)   # live figures + pack labels
DIM    = Color(0.784, 0.804, 0.843)   # zero figures
GREY   = Color(0.549, 0.549, 0.549)   # page footer

# ── type + metrics (unchanged from the original) ─────────────────────────────
F_TITLE = 18
F_BAND  = 13
F_SHOP  = 11
F_HDR   = 9.5
F_DATA  = 9
F_TOTAL = 10.5     # the TOTAL row reads a step up from the body
F_FOOT  = 8

TOTAL_RULE = 1.6   # gold rules closing the TOTAL row top and bottom
HDR_RULE      = 1.6  # gold column separators inside the header row
HDR_RULE_EDGE = 2.2  # gold rule closing the bottom of the header row

PH        = 841.890      # A4 height retained — pagination must not change
TITLE_H   = 45.4
BAND_H    = 22.7
HDR_H     = 23.4
ROW_H_MAX = 22.8         # nominal; shrunk only if a shop would not otherwise fit
FOOT_GAP  = 9.7          # footer baseline offset from the page bottom
FOOT_RES  = 26.0         # space kept clear for the page-number strip

SIDE_M    = 14.0         # band text inset
BAND_GAP  = 24.0         # minimum gap between the two band captions
PAD       = 8.0          # per side inside a numeric column
LABEL_PAD = 6.2          # brand label inset (matches the original)
PACK_IND  = 12.0         # extra indent on pack rows

HEADERS = ["BRAND/PACK", "OPENING", "RECEIPT", "SALES", "CLOSING"]


def page_width(pages):
    """Narrowest page that still holds the content, the headers and the band."""
    # numeric columns: widest figure vs widest header
    num = []
    for i in range(4):
        w = max(stringWidth(r["values"][i], "Helvetica-Bold", F_DATA)
                for p in pages for r in p["rows"])
        w = max(w, stringWidth(HEADERS[i + 1], "Helvetica-Bold", F_HDR))
        num.append(w + PAD * 2)

    # label column: widest brand, widest indented pack, and the header
    lab = 0.0
    for p in pages:
        for r in p["rows"]:
            font = "Helvetica" if r["kind"] == "pack" else "Helvetica-Bold"
            ind = PACK_IND if r["kind"] == "pack" else 0.0
            lab = max(lab, ind + stringWidth(r["label"], font, F_DATA))
    lab = max(lab, stringWidth(HEADERS[0], "Helvetica-Bold", F_HDR))
    lab += LABEL_PAD * 2

    table = lab + sum(num)

    # the period band must not collide
    band = 0.0
    for p in pages:
        h = p["head"]
        if "band13" in h:
            band = max(band, stringWidth(h["band13"][0], "Helvetica-Bold", F_BAND)
                       + stringWidth(h["band13"][1], "Helvetica-Bold", F_BAND))
        band = max(band, stringWidth(h.get("shop", ""), "Helvetica-Bold", F_SHOP)
                   + stringWidth(h.get("bond", ""), "Helvetica-Bold", F_SHOP))
    band += SIDE_M * 2 + BAND_GAP

    pw = max(table, band)
    if pw > table:                       # spread the slack across the columns
        extra = (pw - table) / 5.0
        lab += extra
        num = [w + extra for w in num]
    return pw, lab, num


def row_height(pages):
    """One row height for the whole document, sized so that EVERY shop fits on
    its own page.

    Page 1 is the tight one: it carries the masthead and the period band on top
    of the shop band and column header, so it has ~68pt less room than a
    continuation page. A tall shop landing on page 1 is what splits a shop
    across two pages — the failure is invisible until the shop order or the bond
    changes, because it depends entirely on which shop happens to come first.
    """
    worst = ROW_H_MAX
    for i, pg in enumerate(pages):
        overhead = (TITLE_H + BAND_H if "title" in pg["head"] else 0) \
                   + BAND_H + HDR_H + FOOT_RES
        worst = min(worst, (PH - overhead) / len(pg["rows"]))
    return worst


def centred(c, text, font, size, colour, x0, x1, y_mid):
    c.setFont(font, size); c.setFillColor(colour)
    c.drawString((x0 + x1) / 2.0 - stringWidth(text, font, size) / 2.0,
                 y_mid - size * 0.35, text)


def build(pages, path):
    ROW_H = row_height(pages)
    PW, LABW, NUMW = page_width(pages)
    xs = [0.0, LABW]
    for w in NUMW:
        xs.append(xs[-1] + w)

    c = canvas.Canvas(path, pagesize=(PW, PH))
    c.setTitle("SHOP SALES CUMULATIVE")
    n = len(pages)

    for pi, p in enumerate(pages):
        head = p["head"]
        y = PH

        if "title" in head:                       # masthead, page 1 only
            y -= TITLE_H
            c.setFillColor(NAVY); c.rect(0, y, PW, TITLE_H, stroke=0, fill=1)
            centred(c, head["title"], "Helvetica-Bold", F_TITLE, GOLD, 0, PW, y + TITLE_H / 2)

        if "band13" in head:                      # report name + period
            y -= BAND_H
            c.setFillColor(GOLD); c.rect(0, y, PW, BAND_H, stroke=0, fill=1)
            c.setFont("Helvetica-Bold", F_BAND); c.setFillColor(NAVY)
            c.drawString(SIDE_M, y + BAND_H / 2 - F_BAND * 0.35, head["band13"][0])
            c.drawRightString(PW - SIDE_M, y + BAND_H / 2 - F_BAND * 0.35, head["band13"][1])

        y -= BAND_H                               # shop + bond
        c.setFillColor(GOLD); c.rect(0, y, PW, BAND_H, stroke=0, fill=1)
        c.setFont("Helvetica-Bold", F_SHOP); c.setFillColor(NAVY)
        c.drawString(SIDE_M, y + BAND_H / 2 - F_SHOP * 0.35, head["shop"])
        c.drawRightString(PW - SIDE_M, y + BAND_H / 2 - F_SHOP * 0.35, head["bond"])

        y -= HDR_H                                # column headers
        # painted as five discrete cells, exactly as the source does, so the
        # output is structurally diffable against it
        c.setFillColor(NAVY)
        for i in range(5):
            c.rect(xs[i], y, xs[i + 1] - xs[i], HDR_H, stroke=0, fill=1)
        c.setFont("Helvetica-Bold", F_HDR); c.setFillColor(GOLD)
        c.drawString(LABEL_PAD, y + HDR_H / 2 - F_HDR * 0.35, HEADERS[0])
        for i in range(4):
            centred(c, HEADERS[i + 1], "Helvetica-Bold", F_HDR, GOLD,
                    xs[i + 1], xs[i + 2], y + HDR_H / 2)

        # gold grid on the header row. Drawn inside the row so its height — and
        # therefore the pagination — is unchanged. No rule along the top: the
        # gold shop/bond band already sits directly above it.
        c.setStrokeColor(GOLD); c.setLineCap(0)
        c.setLineWidth(HDR_RULE)
        for i in range(1, 5):
            c.line(xs[i], y, xs[i], y + HDR_H)
        c.setLineWidth(HDR_RULE_EDGE)
        c.line(0, y + HDR_RULE_EDGE / 2, PW, y + HDR_RULE_EDGE / 2)

        for r in p["rows"]:                       # body
            y -= ROW_H
            mid = y + ROW_H / 2
            kind = r["kind"]

            is_total = kind == "total"
            if kind in ("brand", "total"):
                bg, ink = NAVY, (GOLD if is_total else WHITE)
                font = "Helvetica-Bold"
            else:
                bg = ZEBRA if r["zebra"] else WHITE
                ink, font = INK, "Helvetica"

            size = F_TOTAL if is_total else F_DATA

            c.setFillColor(bg)
            for i in range(5):
                c.rect(xs[i], y, xs[i + 1] - xs[i], ROW_H, stroke=0, fill=1)

            c.setFont(font, size); c.setFillColor(ink)
            c.drawString(LABEL_PAD + (PACK_IND if kind == "pack" else 0),
                         mid - size * 0.35, r["label"])

            for i, v in enumerate(r["values"]):
                col = ink if kind != "pack" else (DIM if r["dim"][i] else INK)
                centred(c, v, "Helvetica-Bold", size, col, xs[i + 1], xs[i + 2], mid)

            # gold rules top and bottom — drawn inside the row so the row height,
            # and therefore the pagination, is unchanged
            if is_total:
                c.setStrokeColor(GOLD); c.setLineWidth(TOTAL_RULE)
                inset = TOTAL_RULE / 2.0
                c.line(0, y + ROW_H - inset, PW, y + ROW_H - inset)
                c.line(0, y + inset, PW, y + inset)

        centred(c, f"Page {pi + 1} of {n}", "Helvetica", F_FOOT, GREY,
                0, PW, FOOT_GAP + F_FOOT * 0.35)
        c.showPage()

    c.save()
    return PW, xs


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "cumulative.json"
    out = sys.argv[2] if len(sys.argv) > 2 else "SHOP SALES CUMULATIVE - MOBILE.pdf"
    pages = json.load(open(src))
    pw, xs = build(pages, out)
    tall = max(len(p["rows"]) for p in pages)
    rh = row_height(pages)
    used = BAND_H + HDR_H + tall * rh
    print(f"page {pw:.1f} × {PH:.0f}pt (was 595.3) | tallest page {tall} rows = "
          f"{used:.0f}pt of {PH - 20:.0f}pt usable | row {rh:.2f}pt")
    print("wrote", out)
