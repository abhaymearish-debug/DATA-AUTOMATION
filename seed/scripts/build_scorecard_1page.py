#!/usr/bin/env python3
"""BOND LIQUIDATION SCORECARD — one page, same format, phone-legible.

Identical layout to the source report (4 metric blocks side by side on a
single landscape page) but with the dead whitespace squeezed out so the
page is ~1/3 narrower and every figure is set larger and bolder.  Net
effect: text renders roughly 1.6x bigger at fit-to-width on a phone.
"""
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor

from build_scorecard_mobile import ROWS, PERIOD  # single source of data

# ---------------------------------------------------------------- palette
NAVY     = HexColor("#0B2C52")
NAVY_DK  = HexColor("#081F3B")
GOLD     = HexColor("#FAAF19")
RED      = HexColor("#CF1322")
GREEN    = HexColor("#2E7D0E")
RED_LT   = HexColor("#FF7875")
GREEN_LT = HexColor("#52C41A")
INK      = HexColor("#1F2430")
SUB      = HexColor("#5A6472")
ZEBRA    = HexColor("#EFF3FA")
WHITE    = HexColor("#FFFFFF")
TOTAL_BG = HexColor("#2B353F")
AVG_BG   = HexColor("#3D4A57")
GRID     = HexColor("#CBD4E3")
MUTED    = HexColor("#9AA3B2")

BLOCKS = ["SHOP LIQUIDATION (KSBC)", "SECONDARY SALES",
          "FED / BAR INVOICE", "TOTAL LIQUIDATION"]

# ---------------------------------------------------------------- geometry
M        = 5                      # page margin
BOND_W   = 97
CW       = [31, 31, 39, 39]       # AUG, JUL, d CS, d %
BLOCK_W  = sum(CW)
GUT      = 5                      # gutter between blocks
PW       = M * 2 + BOND_W + 4 * BLOCK_W + 3 * GUT

H_HERO, H_BAND, H_BLK, H_COL, H_ROW, H_FOOT = 28, 19, 15, 15, 17.0, 11
PH = H_HERO + H_BAND + H_BLK + H_COL + len(ROWS) * H_ROW + H_FOOT + 3

# x origin of each block
BX = [M + BOND_W + i * (BLOCK_W + GUT) for i in range(4)]


def is_neg(t):
    return t.startswith("-") and t != "-"


def tri(c, cx, cy, up, col, s=4.0):
    c.setFillColor(col)
    p = c.beginPath()
    if up:
        p.moveTo(cx, cy + s * 0.60); p.lineTo(cx - s * 0.60, cy - s * 0.45); p.lineTo(cx + s * 0.60, cy - s * 0.45)
    else:
        p.moveTo(cx, cy - s * 0.60); p.lineTo(cx - s * 0.60, cy + s * 0.45); p.lineTo(cx + s * 0.60, cy + s * 0.45)
    p.close()
    c.drawPath(p, fill=1, stroke=0)


def delta(c, x, w, yb, txt, size, pos, neg_c, flat, ref=None):
    if txt in ("-", "", None):
        c.setFont("Helvetica", size); c.setFillColor(flat)
        c.drawCentredString(x + w / 2, yb, "–"); return
    src = ref if ref not in (None, "-", "") else txt
    neg = is_neg(src)
    zero = src.lstrip("-").rstrip("%") in ("0", "0.0")
    col = flat if zero else (neg_c if neg else pos)
    f = "Helvetica-Bold"
    c.setFont(f, size); c.setFillColor(col)
    tw = c.stringWidth(txt, f, size)
    gap, mk = 2.2, 4.8
    sx = x + (w - (mk + gap + tw)) / 2
    if not zero:
        tri(c, sx + mk / 2, yb + 2.6, not neg, col)
    c.drawString(sx + mk + gap, yb, txt)


def plain(c, x, w, yb, txt, font, size, col, flat):
    c.setFont(font, size)
    if txt in ("-", "", None):
        c.setFillColor(flat); c.drawCentredString(x + w / 2, yb, "–")
    else:
        c.setFillColor(col); c.drawCentredString(x + w / 2, yb, txt)


def fit_label(c, label, avail, base):
    """largest size <= base at which the label still fits the bond column"""
    size = base
    while size > 5.4 and c.stringWidth(label, "Helvetica-Bold", size) > avail:
        size -= 0.2
    return size


def build(out):
    c = canvas.Canvas(out, pagesize=(PW, PH))
    c.setTitle("BOND LIQUIDATION SCORECARD - AUG vs JUL 2026")

    y = PH
    # ---- hero
    y -= H_HERO
    c.setFillColor(NAVY); c.rect(0, y, PW, H_HERO, fill=1, stroke=0)
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(PW / 2, y + 8.5, "K.S DISTILLERY")

    # ---- gold band
    y -= H_BAND
    c.setFillColor(GOLD); c.rect(0, y, PW, H_BAND, fill=1, stroke=0)
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(M + 2, y + 5.5, "BOND LIQUIDATION SCORECARD")
    c.setFont("Helvetica-Bold", 9.5)
    c.drawRightString(PW - M - 2, y + 5.5, PERIOD)

    # ---- block titles
    y -= H_BLK
    top_tbl = y + H_BLK
    c.setFillColor(NAVY); c.rect(M, y, BOND_W, H_BLK + 0, fill=1, stroke=0)
    for i, t in enumerate(BLOCKS):
        c.setFillColor(NAVY); c.rect(BX[i], y, BLOCK_W, H_BLK, fill=1, stroke=0)
        c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 8.6)
        c.drawCentredString(BX[i] + BLOCK_W / 2, y + 4.6, t)

    # ---- column header
    y -= H_COL
    c.setFillColor(NAVY); c.rect(M, y, PW - 2 * M, H_COL, fill=1, stroke=0)
    c.setFillColor(WHITE); c.setFont("Helvetica-Bold", 9)
    c.drawString(M + 5, y + 4.4, "Bond")
    c.setFillColor(GOLD); c.setFont("Helvetica-Bold", 8.2)
    for i in range(4):
        x = BX[i]
        for j, h in enumerate(["AUG", "JUL", "Δ CS", "Δ %"]):
            c.drawCentredString(x + CW[j] / 2, y + 4.4, h)
            x += CW[j]

    body_top = y
    # ---- rows
    zi = 0
    for label, kind, cells in ROWS:
        y -= H_ROW
        yb = y + 5.6
        if kind == "bond":
            bg = ZEBRA if zi % 2 else WHITE
            zi += 1
            c.setFillColor(bg); c.rect(M, y, PW - 2 * M, H_ROW, fill=1, stroke=0)
            c.setStrokeColor(GRID); c.setLineWidth(0.35); c.line(M, y, PW - M, y)
            c.setFillColor(NAVY)
            c.setFont("Helvetica-Bold", fit_label(c, label, BOND_W - 9, 8.6))
            c.drawString(M + 5, yb, label)
            for i in range(4):
                v = cells[i * 4:i * 4 + 4]; x = BX[i]
                plain(c, x, CW[0], yb, v[0], "Helvetica-Bold", 9.4, INK, MUTED); x += CW[0]
                plain(c, x, CW[1], yb, v[1], "Helvetica", 9.0, SUB, MUTED); x += CW[1]
                delta(c, x, CW[2], yb, v[2], 8.6, GREEN, RED, MUTED); x += CW[2]
                delta(c, x, CW[3], yb, v[3], 8.6, GREEN, RED, MUTED, ref=v[2])
        else:
            fill = NAVY if kind == "cluster" else (TOTAL_BG if kind == "total" else AVG_BG)
            lab = GOLD if kind == "cluster" else WHITE
            prior = HexColor("#C9D4E4") if kind == "cluster" else HexColor("#C9CFD6")
            zi = 0
            c.setFillColor(fill); c.rect(M, y, PW - 2 * M, H_ROW, fill=1, stroke=0)
            c.setFillColor(lab)
            c.setFont("Helvetica-Bold", fit_label(c, label, BOND_W - 9, 8.6))
            c.drawString(M + 5, yb, label)
            for i in range(4):
                v = cells[i * 4:i * 4 + 4]; x = BX[i]
                plain(c, x, CW[0], yb, v[0], "Helvetica-Bold", 9.4, lab, HexColor("#8FA0B8")); x += CW[0]
                plain(c, x, CW[1], yb, v[1], "Helvetica-Bold", 9.0, prior, HexColor("#8FA0B8")); x += CW[1]
                delta(c, x, CW[2], yb, v[2], 8.6, GREEN_LT, RED_LT, HexColor("#8FA0B8")); x += CW[2]
                delta(c, x, CW[3], yb, v[3], 8.6, GREEN_LT, RED_LT, HexColor("#8FA0B8"), ref=v[2])

    # ---- gold gutters between blocks + outer frame
    c.setStrokeColor(GOLD); c.setLineWidth(1.4)
    for i in range(4):
        for gx in (BX[i] - GUT / 2 - 0.5, BX[i] + BLOCK_W + GUT / 2 - 0.5):
            if M < gx < PW - M:
                c.line(gx, y, gx, top_tbl)
    c.line(BX[0] - GUT / 2 - 0.5, y, BX[0] - GUT / 2 - 0.5, top_tbl)
    c.setStrokeColor(NAVY); c.setLineWidth(1.0)
    c.rect(M, y, PW - 2 * M, top_tbl - y, fill=0, stroke=1)

    # ---- footer
    c.setFillColor(HexColor("#6B7684")); c.setFont("Helvetica", 6.4)
    c.drawString(M + 2, 4,
                 "KSBC tertiary + Consumer Fed + BAR   ·   Total Liquidation = Shop Liquidation + Fed / BAR   "
                 "·   average daily sale ÷ 22 days")
    c.drawRightString(PW - M - 2, 4, "K.S DISTILLERY")

    c.showPage(); c.save()
    print("wrote %s  (%.0f x %.0f pt)" % (out, PW, PH))


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "scorecard_1page.pdf")
