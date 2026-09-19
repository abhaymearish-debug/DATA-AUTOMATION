#!/usr/bin/env python3
"""Liquidation Summary: the bond liquidation scorecard, this window vs last month.

Four blocks across, each SEP / AUG / change in cases / change per cent:

    SHOP LIQUIDATION   cases sold out of KSBC shops
    SECONDARY SALES    every dispatch line, whoever it went to
    FED / BAR INVOICE  the part of that dispatch invoiced straight to Consumer
                       Fed outlets and bars - liquidation, because nobody sells
                       it on, which is why it is broken out
    TOTAL LIQUIDATION  shop + fed/bar. Not shop + secondary: that would count
                       every case a shop both received and sold, twice.

Geometry read out of the office's own PDF with a content-stream parse:

    page       960 wide; height is 106 + 20 per row, so 20 rows gives 506
    title      44 navy, 'K.S DISTILLERY' 18pt gold
    subtitle   26 gold, 12pt navy
    header     36 navy in two tiers, gold 9pt group titles and 8.5pt columns
    row        20, zebra white / #F5F7FC, every cell hairlined in grey
    gutter     6pt navy between blocks, gold 1pt rules down both edges
    delta      a 5 x 5.5 triangle then the figure, the pair centred in the cell

Usage:
    build_liquidation_pdf.py --data <scorecard.json> --out <file.pdf>
                             [--title "SEPTEMBER vs AUGUST 2026  ·  DAYS 1-15"]
"""

from __future__ import annotations

import argparse
import json
import sys
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas as pdfcanvas

# ---------------------------------------------------------------- geometry --
PAGE_W = 960.0
TITLE_H, SUB_H, HEAD_H, ROW_H = 44.0, 26.0, 36.0, 20.0
BOND_W, GROUP_W, GUTTER = 110.0, 208.0, 6.0
COL_W = GROUP_W / 4
INSET, BOND_X = 12.0, 6.0
BASE_TITLE, BASE_SUB, BASE_ROW = 28.0, 17.0, 13.0
BASE_GROUP, BASE_BOND, BASE_COL = 12.0, 22.0, 30.0
TRI_W, TRI_H, TRI_GAP, TRI_LIFT = 5.0, 5.5, 1.75, 0.5
HAIRLINE = 0.3

F_TITLE, F_SUB, F_GROUP, F_BOND_HEAD, F_COL, F_ROW = 18.0, 12.0, 9.0, 10.0, 8.5, 8.0
BOLD, BOOK = "Helvetica-Bold", "Helvetica"

NAVY    = colors.Color(0.04, 0.17, 0.32)
NAVY_T  = colors.Color(0.043, 0.173, 0.322)
GOLD    = colors.Color(0.98, 0.69, 0.1)
GOLD_T  = colors.Color(0.98, 0.686, 0.098)
PAPER   = colors.Color(0.96, 0.97, 0.99)
INK     = colors.Color(0.157, 0.157, 0.157)
HAIR    = colors.Color(0.78, 0.78, 0.78)
SLATE   = colors.Color(0.17, 0.21, 0.25)
SLATE_2 = colors.Color(0.24, 0.29, 0.34)
UP      = colors.Color(0.247, 0.525, 0.0)
UP_MARK = colors.Color(0.25, 0.53, 0.0)
DOWN    = colors.Color(0.812, 0.075, 0.133)
UP_DARK = colors.Color(0.322, 0.769, 0.102)     # on a navy or slate row
DOWN_DK = colors.Color(1.0, 0.471, 0.459)

GROUPS = ["SHOP LIQUIDATION (KSBC)", "SECONDARY SALES",
          "FED / BAR INVOICE", "TOTAL LIQUIDATION"]


def w(text, font, size):
    return pdfmetrics.stringWidth(str(text), font, size)


def whole(v: float) -> int:
    return int(Decimal(str(v)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


# Round off: whole cases and whole per cent, the way somebody reads the sheet
# out. Off, the two places the office's own sheet carries.
CASES = {False: "0.01", True: "1"}
PER_CENT = {False: "0.1", True: "1"}


def money(v: float, places: str = "0.01") -> str:
    """Cases to two places, per cent to one - as the office's sheet prints.

    Rounded half-up through Decimal rather than by format string, so a figure
    reads the same here as it does in the workbook and on the screen.
    """
    q = Decimal(str(v)).quantize(Decimal(places), rounding=ROUND_HALF_UP)
    # A figure that lands whole prints whole: "5.00" is 5 with noise after it.
    out = f"{q:f}"
    return out.rstrip("0").rstrip(".") if "." in out else out


def group_x(i: int) -> float:
    return BOND_W + i * (GROUP_W + GUTTER)


def cell_x(g: int, c: int) -> float:
    return group_x(g) + c * COL_W


# ------------------------------------------------------------------ render --
def draw_cell(c, x, y, width, fill):
    c.setFillColor(fill)
    c.setStrokeColor(HAIR)
    c.setLineWidth(HAIRLINE)
    c.rect(x, y - ROW_H, width, ROW_H, stroke=1, fill=1)


def draw_delta(c, x, y, text, rising, tint, font, size):
    """A triangle and a figure, centred together in the cell."""
    base = y - BASE_ROW
    span = TRI_W + TRI_GAP + w(text, font, size)
    left = x + COL_W / 2 - span / 2

    c.setFillColor(tint)
    c.setStrokeColor(tint)
    p = c.beginPath()
    if rising:
        p.moveTo(left + TRI_W / 2, base + TRI_LIFT + TRI_H)
        p.lineTo(left, base + TRI_LIFT)
        p.lineTo(left + TRI_W, base + TRI_LIFT)
    else:
        p.moveTo(left + TRI_W / 2, base + TRI_LIFT)
        p.lineTo(left, base + TRI_LIFT + TRI_H)
        p.lineTo(left + TRI_W, base + TRI_LIFT + TRI_H)
    p.close()
    c.drawPath(p, stroke=1, fill=1)

    c.setFont(font, size)
    c.drawString(left + TRI_W + TRI_GAP, base, text)


def draw_row(c, y, row, stripe, round_off=False):
    kind = row["kind"]
    if kind == "cluster":
        fill, ink, up, down, font = NAVY, GOLD_T, UP_DARK, DOWN_DK, BOLD
    elif kind == "total":
        fill, ink, up, down, font = SLATE, colors.white, UP_DARK, DOWN_DK, BOLD
    elif kind == "average":
        fill, ink, up, down, font = SLATE_2, colors.white, UP_DARK, DOWN_DK, BOLD
    else:
        fill = PAPER if stripe % 2 else colors.white
        ink, up, down, font = INK, UP, DOWN, BOOK
    base = y - BASE_ROW

    draw_cell(c, 0, y, BOND_W, fill)
    c.setFillColor(ink)
    c.setFont(font, F_ROW)
    c.drawString(BOND_X, base, row["label"])

    for g, (now, was) in enumerate(row["blocks"]):
        x0 = group_x(g)
        if g:
            # the navy channel between blocks, gold-ruled on both edges
            c.setFillColor(NAVY)
            c.rect(x0 - GUTTER, y - ROW_H, GUTTER, ROW_H, stroke=0, fill=1)
            c.setStrokeColor(GOLD)
            c.setLineWidth(1)
            c.line(x0 - GUTTER, y, x0 - GUTTER, y - ROW_H)
            c.line(x0, y, x0, y - ROW_H)

        for col in range(4):
            draw_cell(c, cell_x(g, col), y, COL_W, fill)

        for col, v in ((0, now), (1, was)):
            c.setFillColor(ink)
            c.setFont(font, F_ROW)
            text = "-" if not v else money(v, CASES[round_off])
            c.drawCentredString(cell_x(g, col) + COL_W / 2, base, text)

        # Change in cases. Two empty months have no story, so they read '-'
        # rather than a confident zero.
        if not now and not was:
            c.setFillColor(ink)
            c.setFont(font, F_ROW)
            c.drawCentredString(cell_x(g, 2) + COL_W / 2, base, "-")
        else:
            d = now - was
            rising = d >= 0
            figure = money(abs(d), CASES[round_off])
            draw_delta(c, cell_x(g, 2), y,
                       figure if rising else f"-{figure}",
                       rising, (up if rising else down), font, F_ROW)

        # Per cent needs something to divide by.
        if not was:
            c.setFillColor(ink)
            c.setFont(font, F_ROW)
            c.drawCentredString(cell_x(g, 3) + COL_W / 2, base, "-")
        else:
            pct = (now - was) / was * 100
            rising = pct >= 0
            shown = money(abs(pct), PER_CENT[round_off])
            label = f"{shown}%" if rising else f"-{shown}%"
            draw_delta(c, cell_x(g, 3), y, label, rising, (up if rising else down), font, F_ROW)


def build(data: dict, subtitle: str, out: Path, round_off: bool = False) -> int:
    rows = data["rows"]
    page_h = TITLE_H + SUB_H + HEAD_H + ROW_H * len(rows)

    c = pdfcanvas.Canvas(str(out), pagesize=(PAGE_W, page_h))
    c.setTitle("Liquidation Summary")

    y = page_h
    c.setFillColor(NAVY)
    c.rect(0, y - TITLE_H, PAGE_W, TITLE_H, stroke=0, fill=1)
    c.setFillColor(GOLD_T)
    c.setFont(BOLD, F_TITLE)
    c.drawCentredString(PAGE_W / 2, y - BASE_TITLE, "K.S DISTILLERY")
    y -= TITLE_H

    c.setFillColor(GOLD)
    c.rect(0, y - SUB_H, PAGE_W, SUB_H, stroke=0, fill=1)
    c.setFillColor(NAVY_T)
    c.setFont(BOLD, F_SUB)
    # The report is called Liquidation Summary everywhere else - on the nav, in
    # the filename, in the chat about it - so the band says that too.
    c.drawString(INSET, y - BASE_SUB, "LIQUIDATION SUMMARY")
    c.drawRightString(PAGE_W - INSET, y - BASE_SUB, subtitle)
    y -= SUB_H

    c.setFillColor(NAVY)
    c.rect(0, y - HEAD_H, PAGE_W, HEAD_H, stroke=0, fill=1)
    c.setFillColor(GOLD_T)
    c.setFont(BOLD, F_BOND_HEAD)
    c.drawString(8, y - BASE_BOND, "Bond")
    for g, name in enumerate(GROUPS):
        x0 = group_x(g)
        c.setFont(BOLD, F_GROUP)
        c.drawCentredString(x0 + GROUP_W / 2, y - BASE_GROUP, name)
        c.setFont(BOLD, F_COL)
        for col, label in enumerate(data.get("headings", ["NOW", "WAS", "CS", "%"])):
            if col < 2:
                c.drawCentredString(cell_x(g, col) + COL_W / 2, y - BASE_COL, label)
            else:
                draw_delta(c, cell_x(g, col), y - (BASE_COL - BASE_ROW), label,
                           True, GOLD_T, BOLD, F_COL)
        if g:
            c.setStrokeColor(GOLD)
            c.setLineWidth(1)
            c.line(x0 - GUTTER, y, x0 - GUTTER, y - HEAD_H)
            c.line(x0, y, x0, y - HEAD_H)
    y -= HEAD_H

    stripe = 0
    for row in rows:
        draw_row(c, y, row, stripe, round_off)
        stripe = stripe + 1 if row["kind"] == "bond" else 0
        y -= ROW_H

    c.showPage()
    c.save()
    return len(rows)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True, help="the scorecard, as JSON")
    ap.add_argument("--out", required=True)
    ap.add_argument("--title", default="")
    ap.add_argument("--round", action="store_true",
                    help="whole cases and whole per cent")
    a = ap.parse_args()

    blob = json.loads(Path(a.data).read_text())
    if not blob.get("rows"):
        print("ERROR: nothing to draw for that window.")
        return 3
    out = Path(a.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    n = build(blob, a.title or blob.get("subtitle", ""), out, a.round)
    print(f"{n} rows  ->  {out.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
