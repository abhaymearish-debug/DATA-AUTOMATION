#!/usr/bin/env python3
"""Shop Sales - Cumulative: one PDF per bond, one page per shop.

Nothing here is guessed. Every band height, column width, font size, colour,
baseline and pagination rule was read out of the office's own PDFs with a
content-stream parse, so this prints the same document rather than a lookalike:

    page          400.865 x 841.89     (tall and narrow - reads on a phone)
    title band    45.4  navy, 'K.S DISTILLERY' 18pt gold      (page 1 only)
    period band   22.7  gold, 11pt navy                       (page 1 only)
    shop band     22.7  gold, 11pt navy                       (every page)
    header        23.4  navy, 9.5pt gold, gold rules 1.6 / 2.2
    body row      22.8 at full size, squeezed to 18.35 when a shop needs it
    baseline      row_top - (row_height / 2 + font_size * 0.275)
    columns       177.55 / 59.70 / 56.945 / 47.445 / 59.225

The squeeze is the part that matters: the office picks ONE row height for the
whole bond, the largest height at which every shop still lands on a single
page, and only splits a shop when even the 18.35 floor will not hold it.  That
is why their files read one shop per page.

Usage:
    build_shop_cumulative_pdfs.py --raw <cumulative.xlsx> --outdir <dir>
                                  [--period "1 August 2026 - 17 August 2026"]
                                  [--verify]
"""

from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas as pdfcanvas

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app import reports_api  # noqa: E402  (path set above)

# ---------------------------------------------------------------- geometry --
PAGE_H = 841.89
COLS = [177.55, 59.70, 56.945, 47.445, 59.225]
TITLE_H, BAND_H, HEAD_H = 45.4, 22.7, 23.4
ROW_MAX, ROW_MIN, TOTAL_MIN = 22.8, 18.35, 20.075
BOTTOM, FOOT_BASE = 26.0, 17.7
INSET, BRAND_X, PACK_X, HEAD_X = 15.0, 6.2, 18.2, 5.0
COL0_PAD, BASE_BAND, BASE_TITLE = 4.841, 15.0, 28.0
# The gold rules that close a band, and the hairline grid inside the body -
# both the weights the Warehouse Stock Report uses, so the two reports sit in
# a folder together without one looking heavier-handed than the other. The
# rule under the header used to be 2.2, which read as a bar rather than a line.
RULE_V, RULE_H = 0.9, 1.3
GRID_LW = 0.4

F_TITLE, F_BAND, F_HEAD, F_ROW, F_TOTAL, F_FOOT = 18.0, 11.0, 9.5, 9.0, 10.5, 8.0
BOLD, BOOK = "Helvetica-Bold", "Helvetica"

NAVY   = colors.Color(0.04, 0.16, 0.31)
NAVY_T = colors.Color(0.043, 0.161, 0.31)
GOLD   = colors.Color(1.0, 0.74, 0.19)
GOLD_T = colors.Color(1.0, 0.741, 0.192)
PAPER  = colors.Color(0.96, 0.97, 0.99)
INK    = colors.Color(0.314, 0.314, 0.314)
ZERO   = colors.Color(0.784, 0.804, 0.843)
HAIR   = colors.Color(0.855, 0.871, 0.898)

HEADINGS = ["BRAND/PACK", "OPENING", "RECEIPT", "SALES", "CLOSING"]
MEASURES = [("Shop Opening Cases", "Shop Opening Bottles"),
            ("Shop In Cases", "Shop In Bottles"),
            ("Shop Out Cases", "Shop Out Bottles"),
            ("Shop Closing Cases", "Shop Closing Bottles")]

BODY_TOP_FIRST = PAGE_H - TITLE_H - BAND_H - BAND_H - HEAD_H   # 727.69
BODY_TOP_REST = PAGE_H - BAND_H - HEAD_H                       # 795.79


def w(text, font, size):
    return pdfmetrics.stringWidth(str(text), font, size)


ROUND_OFF = False


def money(value: float, round_off: bool | None = None) -> str:
    """Rounded half up - the way the office's clerk rounds, never banker's.

    Whole cases when the sheet is rounded off, up to two decimals otherwise.
    A figure that lands on a whole case prints as one either way: "5.00" is 5
    with noise on the end, and a column of them is noise all the way down.
    """
    whole = ROUND_OFF if round_off is None else round_off
    step = Decimal("1") if whole else Decimal("0.01")
    got = Decimal(str(value or 0)).quantize(step, rounding=ROUND_HALF_UP)
    out = f"{got:,.0f}" if whole else f"{got:,.2f}"
    return out.rstrip("0").rstrip(".") if "." in out else out


# -------------------------------------------------------------------- data --
def read_lines(path: Path):
    """A window rebuilt from the daily exports, handed over as JSON.

    Same shape as reading a cumulative file, so everything downstream - the
    squeeze, the pagination, the layout - is identical whichever way the
    window was assembled.
    """
    import json
    blob = json.loads(Path(path).read_text())
    master = reports_api.load_master()
    data: dict = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(
        lambda: [0.0, 0.0, 0.0, 0.0]))))
    names = {k: v for k, v in blob["names"].items()}
    for key, values in blob["lines"].items():
        code, brand, pack = key.split("|", 2)
        bond = str(master.get(code, {}).get("bond", "")).strip().upper()
        if not bond:
            continue
        cell = data[bond][code][brand][pack]
        for k in range(4):
            cell[k] += values[k]
    return data, names


def read_cumulative(path: Path):
    """bond -> shop code -> brand -> pack -> [opening, receipt, sales, closing]."""
    master = reports_api.load_master()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    header = [str(h).strip() if h else "" for h in next(rows)]
    ix = {h: i for i, h in enumerate(header)}

    need = ["Shop Code", "Shop Name", "Brand Name", "Packing", "Bottle Per Case"]
    for col in need + [c for pair in MEASURES for c in pair]:
        if col not in ix:
            raise SystemExit(f"ERROR: the raw file has no '{col}' column.")

    data: dict = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(
        lambda: [0.0, 0.0, 0.0, 0.0]))))
    names: dict[str, str] = {}

    for r in rows:
        if r is None or r[ix["Shop Code"]] is None:
            continue
        code = str(r[ix["Shop Code"]]).strip()
        bond = str(master.get(code, {}).get("bond", "")).strip().upper()
        if not bond:
            continue
        names[code] = str(r[ix["Shop Name"]] or "").strip().upper()
        brand = str(r[ix["Brand Name"]] or "").strip().upper()
        pack = str(r[ix["Packing"]] or "").strip().upper()
        bpc = r[ix["Bottle Per Case"]] or 0

        cell = data[bond][code][brand][pack]
        for k, (c_col, b_col) in enumerate(MEASURES):
            cases = r[ix[c_col]] or 0
            bottles = r[ix[b_col]] or 0
            cell[k] += float(cases) + (float(bottles) / bpc if bpc else 0.0)

    wb.close()
    return data, names


def shop_lines(brands: dict):
    """Flatten one shop into printable rows plus its four totals.

    Brands read alphabetically; packs read as the office sorts them, on the
    label rather than the millilitres, so 1000 ML files before 500 ML.
    """
    out: list[dict] = []
    totals = [0.0, 0.0, 0.0, 0.0]
    for brand in sorted(brands):
        packs = brands[brand]
        sub = [0.0, 0.0, 0.0, 0.0]
        for pack in packs:
            for k in range(4):
                sub[k] += packs[pack][k]
        out.append({"kind": "brand", "label": brand, "values": sub})
        for pack in sorted(packs):
            out.append({"kind": "pack", "label": pack, "values": packs[pack]})
        for k in range(4):
            totals[k] += sub[k]
    out.append({"kind": "total", "label": "TOTAL", "values": totals})
    return out, totals


def shop_order(codes, master, names):
    """Master-data order first - that list is the office's own sequence."""
    rank = {code: i for i, code in enumerate(master)}
    return sorted(codes, key=lambda c: (rank.get(c, 10_000), names.get(c, c)))


# ------------------------------------------------------- page measurements --
def row_height(shops: list[list[dict]]) -> float:
    """One height for the whole bond: the largest that still fits every shop."""
    best = ROW_MAX
    for i, lines in enumerate(shops):
        avail = (BODY_TOP_FIRST if i == 0 else BODY_TOP_REST) - BOTTOM
        best = min(best, avail / max(1, len(lines)))
    return max(ROW_MIN, min(ROW_MAX, best))


def page_width(shops: list[list[dict]]) -> tuple[float, list[float]]:
    """Office width unless a label would run into the next column."""
    cols = list(COLS)
    widest = 0.0
    for lines in shops:
        for row in lines:
            if row["kind"] == "pack":
                widest = max(widest, PACK_X + w("  " + row["label"], BOOK, F_ROW))
            else:
                size = F_TOTAL if row["kind"] == "total" else F_ROW
                widest = max(widest, BRAND_X + w(row["label"], BOLD, size))
    cols[0] = max(cols[0], widest + COL0_PAD)
    return sum(cols), cols


# ------------------------------------------------------------------ render --
def draw_bands(c, cols, width, page_one, shop, bond, period) -> float:
    y = PAGE_H
    if page_one:
        c.setFillColor(NAVY)
        c.rect(0, y - TITLE_H, width, TITLE_H, stroke=0, fill=1)
        c.setFont(BOLD, F_TITLE)
        c.setFillColor(GOLD_T)
        c.drawCentredString(width / 2, y - BASE_TITLE, "K.S DISTILLERY")
        y -= TITLE_H

        c.setFillColor(GOLD)
        c.rect(0, y - BAND_H, width, BAND_H, stroke=0, fill=1)
        c.setFont(BOLD, F_BAND)
        c.setFillColor(NAVY_T)
        c.drawString(INSET, y - BASE_BAND, "SHOP SALES CUMULATIVE")
        c.drawRightString(width - INSET, y - BASE_BAND, period)
        y -= BAND_H

    c.setFillColor(GOLD)
    c.rect(0, y - BAND_H, width, BAND_H, stroke=0, fill=1)
    c.setFont(BOLD, F_BAND)
    c.setFillColor(NAVY_T)
    c.drawString(INSET, y - BASE_BAND, shop)
    c.drawRightString(width - INSET, y - BASE_BAND, bond)
    y -= BAND_H

    base = y - (HEAD_H / 2 + F_HEAD * 0.275)
    x = 0.0
    for label, cw in zip(HEADINGS, cols):
        c.setFillColor(NAVY)
        c.rect(x, y - HEAD_H, cw, HEAD_H, stroke=0, fill=1)
        c.setFillColor(GOLD_T)
        c.setFont(BOLD, F_HEAD)
        c.drawString(x + HEAD_X, base, label)
        c.setStrokeColor(GOLD)
        if cw is not cols[-1]:
            c.setLineWidth(RULE_V)
            c.line(x + cw, y, x + cw, y - HEAD_H)
        c.setLineWidth(RULE_H)
        c.line(x, y - HEAD_H + RULE_H / 2, x + cw, y - HEAD_H + RULE_H / 2)
        x += cw
    return y - HEAD_H


def draw_row(c, cols, width, y, row, h, stripe) -> None:
    brand, total = row["kind"] == "brand", row["kind"] == "total"
    size = F_TOTAL if total else F_ROW
    font = BOLD if (brand or total) else BOOK
    base = y - (h / 2 + size * 0.275)

    if brand or total:
        fill, ink = NAVY, (GOLD_T if total else colors.white)
    else:
        fill, ink = (PAPER if stripe % 2 else colors.white), INK

    c.setFillColor(fill)
    c.rect(0, y - h, width, h, stroke=0, fill=1)
    c.setFillColor(ink)
    c.setFont(font, size)
    if brand or total:
        c.drawString(BRAND_X, base, row["label"])
    else:
        c.drawString(PACK_X, base, "  " + row["label"])

    x = cols[0]
    for k, cw in enumerate(cols[1:]):
        text = money(row["values"][k])
        c.setFillColor(ZERO if (not brand and not total and row["values"][k] == 0) else ink)
        c.drawCentredString(x + cw / 2, base, text)
        x += cw


def draw_grid(c, cols, width, placed) -> None:
    """The grid, laid over the rows so no fill can paint across it.

    A pack row is ruled into cells the way the stock report rules its body:
    hairline columns, and a hairline closing each row. A brand row is a navy
    band and stays solid - it is the heading for the packs under it, not a
    row of cells. The shop's total is lined in gold above and below, which is
    how the stock report closes its own total and is what tells you, at a
    glance down a page of shops, where one shop ends and the next begins.
    """
    edges = [sum(cols[:i]) for i in range(1, len(cols))]

    c.setLineWidth(GRID_LW)
    c.setStrokeColor(HAIR)
    for row, y, rh in placed:
        if row["kind"] in ("brand", "total"):
            continue
        for x in edges:
            c.line(x, y - rh, x, y)
        c.line(0, y - rh, width, y - rh)

    c.setLineWidth(RULE_H)
    c.setStrokeColor(GOLD)
    for row, y, rh in placed:
        if row["kind"] == "total":
            c.line(0, y, width, y)
            c.line(0, y - rh, width, y - rh)


def paginate(shops, h):
    """Greedy fill, exactly as the office's own files break."""
    pages = []
    for i, lines in enumerate(shops):
        top = BODY_TOP_FIRST if i == 0 else BODY_TOP_REST
        y, current = top, []
        for row in lines:
            need = max(h, TOTAL_MIN) if row["kind"] == "total" else h
            if y - need < BOTTOM and current:
                pages.append((i, current))
                y, current = BODY_TOP_REST, []
            current.append(row)
            y -= need
        pages.append((i, current))
    return pages


def build_bond_pdf(bond, shops, names, master, period, out) -> tuple[int, float]:
    codes = shop_order(list(shops), master, names)
    laid = [shop_lines(shops[code])[0] for code in codes]
    sold = sum(lines[-1]["values"][2] for lines in laid)

    h = row_height(laid)
    width, cols = page_width(laid)
    pages = paginate(laid, h)

    c = pdfcanvas.Canvas(str(out), pagesize=(width, PAGE_H))
    c.setTitle(f"Shop Sales Cumulative - {bond.title()}")
    for n, (shop_ix, rows) in enumerate(pages):
        y = draw_bands(c, cols, width, n == 0, names.get(codes[shop_ix], codes[shop_ix]),
                       bond, period)
        placed = []
        for i, row in enumerate(rows):
            rh = max(h, TOTAL_MIN) if row["kind"] == "total" else h
            draw_row(c, cols, width, y, row, rh, i)
            placed.append((row, y, rh))
            y -= rh
        draw_grid(c, cols, width, placed)
        c.setFont(BOOK, F_FOOT)
        c.setFillColor(NAVY_T)
        c.drawCentredString(width / 2, FOOT_BASE, f"Page {n + 1} of {len(pages)}")
        c.showPage()
    c.save()
    return len(pages), sold


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--raw", default="")
    ap.add_argument("--lines", default="", help="a window rebuilt from the daily exports")
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--period", default="")
    ap.add_argument("--verify", action="store_true")
    ap.add_argument("--round-off", action="store_true",
                    help="whole cases, the way the screen shows them with the switch on")
    a = ap.parse_args()

    # money() reads this, so every figure on every page follows the switch
    # without threading a flag through the layout.
    global ROUND_OFF
    ROUND_OFF = bool(a.round_off)

    if not a.raw and not a.lines:
        print("ERROR: give either --raw (a cumulative file) or --lines (a rebuilt window).")
        return 2

    outdir = Path(a.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    if a.lines:
        raw = Path(a.lines)
        if not raw.is_file():
            print(f"ERROR: window file not found: {raw}")
            return 2
        data, names = read_lines(raw)
    else:
        raw = Path(a.raw)
        if not raw.is_file():
            print(f"ERROR: cumulative raw not found: {raw}")
            return 2
        data, names = read_cumulative(raw)
    if not data:
        print("ERROR: no rows in this file map to a bond - check Bond Mapping in Settings.")
        return 3

    master = list(reports_api.load_master())
    period = a.period or raw.stem.replace("_", " ").strip()
    grand = 0.0
    for bond in sorted(data):
        out = outdir / f"Shop Sales Cumulative - {bond.title()}.pdf"
        pages, sold = build_bond_pdf(bond, data[bond], names, master, period, out)
        grand += sold
        print(f"{bond}: {len(data[bond])} shop(s), {pages} page(s), "
              f"{money(sold)} cs sold  ->  {out.name}")

    if a.verify:
        print(f"\n  total sales across all bonds : {money(grand)} cs")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
