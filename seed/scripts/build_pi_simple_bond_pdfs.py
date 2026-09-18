#!/usr/bin/env python3
"""
SIMPLE per-bond PI PDFs — shop by shop, RL / RQ / MQ for the current month
plus the change in MQ vs the prior month. Nothing else.

Built 12 Aug 2026 at Abhay's request. Deliberately NOT the same thing as
build_pi_bond_pdfs.py (that one exports the full locked drill-down sheets via
LibreOffice). This one is a clean 5-column table generated with reportlab.

    Brand | RL | RQ | MQ | Chg vs <PRIOR>

Source: the RAW PI DATA sheet of the newest PI INSIGHTS workbook — every PI
line for both months, parsed verbatim from the KSBC ERP exports. Nothing is
recomputed here, so the PDFs cannot drift from the workbook.

Usage:
    python3 build_pi_simple_bond_pdfs.py [--base DIR] [--workbook PATH]
                                         [--bonds "KOLLAM,ALUVA"] [--outdir DIR]
"""
import argparse
import collections
import os
import re
from collections import defaultdict

import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas

# ---------------------------------------------------------------- house style
NAVY = (0x0D / 255, 0x1B / 255, 0x4A / 255)
NAVY_SOFT = (0x26 / 255, 0x3F / 255, 0x80 / 255)
GOLD = (0xFF / 255, 0xB3 / 255, 0x00 / 255)
WHITE = (1, 1, 1)
INK = (0x1F / 255, 0x29 / 255, 0x37 / 255)
GREY = (0x6B / 255, 0x72 / 255, 0x80 / 255)
ZEBRA = (0xF3 / 255, 0xF4 / 255, 0xF6 / 255)
GREEN = (0x2E / 255, 0x7D / 255, 0x32 / 255)
RED = (0xC6 / 255, 0x28 / 255, 0x28 / 255)

PAGE_W, PAGE_H = A4
L_M = 16 * mm
R_M = 16 * mm
BOT_M = 14 * mm

HDR_H = 34 * mm          # first-page header band
CONT_H = 16 * mm         # continuation-page band
SHOP_H = 8.6 * mm        # shop banner
COLH_H = 6.4 * mm        # column header
ROW_H = 5.4 * mm         # brand row
TOT_H = 6.0 * mm         # per-shop total row
GAP = 4.2 * mm           # gap between shop blocks

FONT = "DejaVu"
FONT_B = "DejaVu-Bold"


def register_fonts():
    """DejaVu carries the ·, +/- and en-dash glyphs the standard 14 fonts lack."""
    for name, path in (
        (FONT, "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        (FONT_B, "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ):
        if not os.path.exists(path):
            raise SystemExit(f"font not found: {path}")
        pdfmetrics.registerFont(TTFont(name, path))


# ------------------------------------------------------------------ data load
def find_workbook(base):
    d = os.path.join(base, "PURCHASE INSTRUCTION")
    cands = [
        os.path.join(d, f)
        for f in os.listdir(d)
        if f.startswith("PI INSIGHTS - ") and f.endswith(".xlsx") and not f.startswith("~$")
    ]
    if not cands:
        raise SystemExit("no PI INSIGHTS workbook found in PURCHASE INSTRUCTION/")
    return max(cands, key=os.path.getmtime)


def months_from_name(path):
    """'PI INSIGHTS - JULY & AUGUST 2026.xlsx' -> ('JULY', 'AUGUST', '2026')"""
    m = re.search(r"PI INSIGHTS - (\w+) & (\w+) (\d{4})", os.path.basename(path))
    if not m:
        raise SystemExit(f"cannot read months from filename: {path}")
    return m.group(1).upper(), m.group(2).upper(), m.group(3)


def load_lines(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["RAW PI DATA"]
    rows = ws.iter_rows(values_only=True)

    ix = None
    for r in rows:
        vals = [str(v).strip() if v is not None else "" for v in r]
        if "MQ (cs)" in vals and "Shop Code" in vals:
            ix = {h: i for i, h in enumerate(vals)}
            break
    if ix is None:
        raise SystemExit("could not locate the header row on RAW PI DATA")

    out = []
    for r in rows:
        if r[ix["Month"]] in (None, ""):
            continue

        def num(key):
            v = r[ix[key]]
            try:
                return float(v)
            except (TypeError, ValueError):
                return 0.0

        out.append(
            {
                "month": str(r[ix["Month"]]).upper(),
                "shop": str(r[ix["Shop Code"]]).strip(),
                "shop_name": str(r[ix["Shop Name"]] or "").strip(),
                "bond": str(r[ix["Bond"]] or "").strip().upper(),
                "brand": str(r[ix["Brand"]] or "").strip(),
                "full": str(r[ix["Product Brand (KSBC)"]] or "").strip(),
                "rl": num("RL (cs)"),
                "rq": num("RQ (cs)"),
                "mq": num("MQ (cs)"),
            }
        )
    wb.close()
    return out


def shape(lines, cur, prior):
    """bond -> [shop dicts], each with brand rows carrying cur RL/RQ/MQ + delta MQ."""
    cur_l = [l for l in lines if l["month"] == cur]
    pri_l = [l for l in lines if l["month"] == prior]

    prior_mq = defaultdict(float)                  # (shop, brand) -> MQ
    prior_shop_mq = defaultdict(float)             # shop -> ΣMQ
    prior_shop_bond, prior_shop_name = {}, {}
    full_name = {}                                 # short brand code -> full KSBC name
    for l in cur_l + pri_l:
        if l["full"]:
            full_name.setdefault(l["brand"], l["full"])
    for l in pri_l:
        prior_mq[(l["shop"], l["brand"])] += l["mq"]
        prior_shop_mq[l["shop"]] += l["mq"]
        prior_shop_bond[l["shop"]] = l["bond"]
        prior_shop_name[l["shop"]] = l["shop_name"]

    # current month, aggregated by shop x brand (never assume the key is unique)
    cur_agg = defaultdict(lambda: {"rl": 0.0, "rq": 0.0, "mq": 0.0})
    shop_meta, shops_by_bond = {}, defaultdict(set)
    for l in cur_l:
        a = cur_agg[(l["shop"], l["brand"])]
        a["rl"] += l["rl"]
        a["rq"] += l["rq"]
        a["mq"] += l["mq"]
        shop_meta[l["shop"]] = (l["shop_name"], l["bond"])
        shops_by_bond[l["bond"]].add(l["shop"])

    bonds = defaultdict(list)
    for bond, shops in shops_by_bond.items():
        for sc in shops:
            name, _ = shop_meta[sc]
            brands = {b for (s, b) in cur_agg if s == sc}
            # a brand that carried a ceiling last month but has no line this
            # month has lost it — show it at 0 so the shop total reconciles
            brands |= {b for (s, b) in prior_mq if s == sc and prior_mq[(s, b)] > 0}

            rows = []
            for b in brands:
                a = cur_agg.get((sc, b), {"rl": 0.0, "rq": 0.0, "mq": 0.0})
                pm = prior_mq.get((sc, b), 0.0)
                rows.append(
                    {
                        "brand": full_name.get(b, b),
                        "rl": a["rl"],
                        "rq": a["rq"],
                        "mq": a["mq"],
                        "d": a["mq"] - pm,
                        "dropped": a["mq"] == 0 and pm > 0,
                    }
                )
            rows.sort(key=lambda r: (-r["mq"], r["brand"]))
            bonds[bond].append(
                {
                    "code": sc,
                    "name": name,
                    "rows": rows,
                    "rl": sum(r["rl"] for r in rows),
                    "rq": sum(r["rq"] for r in rows),
                    "mq": sum(r["mq"] for r in rows),
                    "d": sum(r["mq"] for r in rows) - prior_shop_mq.get(sc, 0.0),
                    "new": sc not in prior_shop_mq,
                }
            )

    # shops that had a PI last month but none this month — flagged, not dropped
    missing = defaultdict(list)
    for sc, pmq in prior_shop_mq.items():
        if sc not in shop_meta:
            missing[prior_shop_bond[sc]].append((sc, prior_shop_name.get(sc, ""), pmq))

    for b in bonds:
        bonds[b].sort(key=lambda s: (-s["mq"], s["name"]))
    return bonds, missing


def load_active_shops(base):
    """{shop code: (name, bond)} for every ACTIVE KSBC shop in the master."""
    path = os.path.join(base, "MASTER DATA CONFIRMED.xlsx")
    if not os.path.exists(path):
        print("  !! master not found — blank-PI shops will not be listed")
        return {}
    ws = openpyxl.load_workbook(path, data_only=True, read_only=True)["16-4-25"]
    out = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) < 9 or r[4] is None:
            continue
        name = str(r[4]).strip()
        # master writes both "2023 OYOOR" and "10001-PONNANI"
        m = re.match(r"\s*(\d+)\s*[-\s]", name)
        if not m:
            continue
        if str(r[5] or "").strip().upper() != "KSBC":
            continue
        if str(r[8] or "").strip().upper() != "ACTIVE":
            continue
        # master prefixes the name with the code ("8010-MANORAMA JUNCTION")
        label = re.sub(r"^\s*\d+\s*[-\s]\s*", "", name).strip() or name
        out[m.group(1)] = (label, str(r[7] or "").strip().upper())
    return out


# -------------------------------------------------------------------- drawing
def fmt(v):
    return "·" if abs(v) < 0.005 else f"{v:,.0f}"


def fmt_d(v):
    if abs(v) < 0.005:
        return "·"
    return f"+{v:,.0f}" if v > 0 else f"{v:,.0f}"


class BondPDF:
    def __init__(self, path, bond, cur, prior, year):
        self.c = rl_canvas.Canvas(path, pagesize=A4)
        self.c.setTitle(f"{bond} — Purchase Instruction {cur.title()} {year}")
        self.bond, self.cur, self.prior, self.year = bond, cur, prior, year
        self.page = 0
        self.y = 0
        self.x_brand = L_M + 3 * mm
        right = PAGE_W - R_M
        self.x_d = right - 3 * mm
        self.x_mq = self.x_d - 24 * mm
        self.x_rq = self.x_mq - 18 * mm
        self.x_rl = self.x_rq - 18 * mm

    # ---- page furniture
    def header(self, first):
        c = self.c
        h = HDR_H if first else CONT_H
        c.setFillColorRGB(*NAVY)
        c.rect(0, PAGE_H - h, PAGE_W, h, stroke=0, fill=1)
        c.setFillColorRGB(*GOLD)
        c.rect(0, PAGE_H - h - 1.6, PAGE_W, 1.6, stroke=0, fill=1)

        if first:
            c.setFillColorRGB(*GOLD)
            c.setFont(FONT_B, 8)
            c.drawString(L_M, PAGE_H - 12 * mm, "K.S. DISTILLERY")
            c.setFillColorRGB(*WHITE)
            c.setFont(FONT_B, 21)
            c.drawString(L_M, PAGE_H - 21 * mm, self.bond)
            c.setFillColorRGB(0.72, 0.77, 0.91)
            c.setFont(FONT, 9.5)
            c.drawString(
                L_M,
                PAGE_H - 27.5 * mm,
                f"Purchase Instruction  ·  {self.cur.title()} {self.year}"
                f"  ·  change shown against {self.prior.title()}",
            )
            c.setFillColorRGB(*GOLD)
            c.setFont(FONT, 8.5)
            c.drawRightString(PAGE_W - R_M, PAGE_H - 21 * mm, "RL · RQ · MQ  (cases)")
        else:
            c.setFillColorRGB(*WHITE)
            c.setFont(FONT_B, 11)
            c.drawString(L_M, PAGE_H - 10.5 * mm, self.bond)
            c.setFillColorRGB(0.72, 0.77, 0.91)
            c.setFont(FONT, 8.5)
            c.drawRightString(
                PAGE_W - R_M,
                PAGE_H - 10.5 * mm,
                f"{self.cur.title()} {self.year}  ·  continued",
            )
        self.y = PAGE_H - h - 8 * mm

    def footer(self):
        c = self.c
        c.setFillColorRGB(*GREY)
        c.setFont(FONT, 7.5)
        c.drawString(L_M, BOT_M - 4 * mm, "MQ = RL + RQ  ·  source: KSBC Purchase Instruction")
        c.drawRightString(PAGE_W - R_M, BOT_M - 4 * mm, f"Page {self.page}")

    def new_page(self, first=False):
        if self.page:
            self.footer()
            self.c.showPage()
        self.page += 1
        self.header(first)

    def need(self, h):
        if self.y - h < BOT_M:
            self.new_page()

    # ---- table pieces
    def col_header(self):
        c = self.c
        y = self.y - COLH_H
        c.setFillColorRGB(*NAVY_SOFT)
        c.rect(L_M, y, PAGE_W - L_M - R_M, COLH_H, stroke=0, fill=1)
        c.setFillColorRGB(*WHITE)
        c.setFont(FONT_B, 8)
        c.drawString(self.x_brand, y + 2.1 * mm, "BRAND")
        for x, t in ((self.x_rl, "RL"), (self.x_rq, "RQ"), (self.x_mq, "MQ")):
            c.drawRightString(x, y + 2.1 * mm, t)
        c.setFillColorRGB(*GOLD)
        c.drawRightString(self.x_d, y + 2.1 * mm, f"CHG vs {self.prior[:3]}")
        self.y = y

    def shop_block(self, s):
        c = self.c
        block_h = SHOP_H + COLH_H + len(s["rows"]) * ROW_H + TOT_H
        # keep a shop's banner with at least its header + 2 rows
        self.need(min(block_h, SHOP_H + COLH_H + 2 * ROW_H + TOT_H))

        y = self.y - SHOP_H
        c.setFillColorRGB(*NAVY)
        c.rect(L_M, y, PAGE_W - L_M - R_M, SHOP_H, stroke=0, fill=1)
        c.setFillColorRGB(*WHITE)
        c.setFont(FONT_B, 10)
        label = f"{s['name']}  ({s['code']})"
        if s["new"]:
            label += "   — new this month"
        c.drawString(self.x_brand, y + 2.7 * mm, label)
        self.y = y

        self.col_header()

        for i, r in enumerate(s["rows"]):
            self.need(ROW_H + TOT_H)
            y = self.y - ROW_H
            if i % 2 == 1:
                c.setFillColorRGB(*ZEBRA)
                c.rect(L_M, y, PAGE_W - L_M - R_M, ROW_H, stroke=0, fill=1)
            c.setFillColorRGB(*(GREY if r["dropped"] else INK))
            c.setFont(FONT, 8.5)
            name = r["brand"]
            avail = self.x_rl - self.x_brand - 6 * mm
            if c.stringWidth(name, FONT, 8.5) > avail:
                raise SystemExit(
                    f"brand name too wide for the column: {name!r} "
                    f"({c.stringWidth(name, FONT, 8.5):.1f}pt > {avail:.1f}pt)"
                )
            c.drawString(self.x_brand, y + 1.7 * mm, name)
            for x, v in ((self.x_rl, r["rl"]), (self.x_rq, r["rq"]), (self.x_mq, r["mq"])):
                c.drawRightString(x, y + 1.7 * mm, fmt(v))
            d = r["d"]
            c.setFillColorRGB(*(GREEN if d > 0.005 else RED if d < -0.005 else GREY))
            c.setFont(FONT_B if abs(d) > 0.005 else FONT, 8.5)
            c.drawRightString(self.x_d, y + 1.7 * mm, fmt_d(d))
            self.y = y

        y = self.y - TOT_H
        c.setFillColorRGB(0.85, 0.87, 0.92)
        c.rect(L_M, y, PAGE_W - L_M - R_M, TOT_H, stroke=0, fill=1)
        c.setFillColorRGB(*NAVY)
        c.setFont(FONT_B, 8.5)
        c.drawString(self.x_brand, y + 1.9 * mm, "TOTAL")
        for x, v in ((self.x_rl, s["rl"]), (self.x_rq, s["rq"]), (self.x_mq, s["mq"])):
            c.drawRightString(x, y + 1.9 * mm, fmt(v))
        d = s["d"]
        c.setFillColorRGB(*(GREEN if d > 0.005 else RED if d < -0.005 else GREY))
        c.drawRightString(self.x_d, y + 1.9 * mm, fmt_d(d))
        self.y = y - GAP

    def bond_total(self, shops, missing=()):
        self.need(11 * mm)
        c = self.c
        y = self.y - 8.6 * mm
        c.setFillColorRGB(*NAVY)
        c.rect(L_M, y, PAGE_W - L_M - R_M, 8.6 * mm, stroke=0, fill=1)
        c.setFillColorRGB(*GOLD)
        c.setFont(FONT_B, 9.5)
        c.drawString(self.x_brand, y + 2.7 * mm, f"{self.bond} — ALL SHOPS")
        c.setFillColorRGB(*WHITE)
        for x, v in (
            (self.x_rl, sum(s["rl"] for s in shops)),
            (self.x_rq, sum(s["rq"] for s in shops)),
            (self.x_mq, sum(s["mq"] for s in shops)),
        ):
            c.drawRightString(x, y + 2.7 * mm, fmt(v))
        # a shop that got no PI at all this month lost its whole ceiling —
        # that loss is part of the bond's movement, so it belongs in this total
        d = sum(s["d"] for s in shops) - sum(m[2] for m in missing)
        c.setFillColorRGB(*(GOLD if abs(d) > 0.005 else WHITE))
        c.drawRightString(self.x_d, y + 2.7 * mm, fmt_d(d))
        self.y = y - GAP

    def note(self, text):
        self.need(6 * mm)
        self.c.setFillColorRGB(*GREY)
        self.c.setFont(FONT, 8)
        self.c.drawString(self.x_brand, self.y - 4 * mm, text)
        self.y -= 6 * mm

    def save(self):
        self.footer()
        self.c.save()


def build(bond, shops, missing, blank, cur, prior, year, outdir):
    path = os.path.join(outdir, f"{bond} — PI {cur.title()} {year}.pdf")
    pdf = BondPDF(path, bond, cur, prior, year)
    pdf.new_page(first=True)
    pdf.bond_total(shops, missing)
    for s in shops:
        pdf.shop_block(s)
    for code, name, pmq in sorted(missing, key=lambda m: -m[2]):
        pdf.note(
            f"No {cur.title()} PI received for {name} ({code}) — lost the "
            f"{pmq:,.0f} cs it carried in {prior.title()} (included in the total above)."
        )
    for code, name in sorted(blank, key=lambda b: b[1]):
        pdf.note(
            f"KSBC issued a blank {cur.title()} PI for {name} ({code}) — "
            f"no brands on indent, nothing to show."
        )
    pdf.save()
    return path, pdf.page


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default="/sessions/modest-lucid-cray/mnt/Claude")
    ap.add_argument("--workbook")
    ap.add_argument("--bonds")
    ap.add_argument("--outdir")
    a = ap.parse_args()

    register_fonts()
    wbpath = a.workbook or find_workbook(a.base)
    prior, cur, year = months_from_name(wbpath)
    print(f"workbook: {os.path.basename(wbpath)}")
    print(f"current: {cur}   prior: {prior}   year: {year}")

    lines = load_lines(wbpath)
    bonds, missing = shape(lines, cur, prior)

    master = load_active_shops(a.base)
    placed = {sc for v in bonds.values() for sc in (s["code"] for s in v)}
    placed |= {m[0] for v in missing.values() for m in v}
    blank = collections.defaultdict(list)
    for sc, (nm, bd) in master.items():
        if sc not in placed:
            blank[bd].append((sc, nm))
    n_blank = sum(len(v) for v in blank.values())
    print(f"active KSBC shops in master: {len(master)}   "
          f"blank-PI shops listed by note: {n_blank}")

    outdir = a.outdir or os.path.join(
        a.base, "PURCHASE INSTRUCTION", f"BOND PI SHEETS - {cur}"
    )
    os.makedirs(outdir, exist_ok=True)

    want = [b.strip().upper() for b in a.bonds.split(",")] if a.bonds else sorted(bonds)
    tot_mq = tot_d = 0
    for b in want:
        if b not in bonds:
            print(f"  !! no shops for bond {b}")
            continue
        shops = bonds[b]
        path, pages = build(b, shops, missing.get(b, []), blank.get(b, []),
                            cur, prior, year, outdir)
        _ = path
        mq = sum(s["mq"] for s in shops)
        d = sum(s["d"] for s in shops) - sum(m[2] for m in missing.get(b, []))
        tot_mq += mq
        tot_d += d
        print(
            f"  {b:<16} {len(shops):>3} shops  {pages:>2}pp  "
            f"MQ {mq:>6,.0f}  chg {d:>+6,.0f}"
        )
    print(f"\n  {'TOTAL':<16}                MQ {tot_mq:>6,.0f}  chg {tot_d:>+6,.0f}")
    covered = len(placed | {s for v in blank.values() for s, _ in v})
    print(f"\ncoverage: {len(placed & set(master))} of {len(master)} active shops with a table; "
          f"{n_blank} named by note; {len(set(master) - (placed | {s for v in blank.values() for s, _ in v}))} unaccounted")
    print(f"\noutput: {outdir}")


if __name__ == "__main__":
    main()
