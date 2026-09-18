#!/usr/bin/env python3
"""Bondwise x brandwise TARGET PDFs, one per cluster.
Reproduces the locked JULY 2026 design from TARGET -<MONTH>.xlsx.
Usage: build_target_cluster_pdfs.py --month AUGUST [--year 2026] [--base <Claude folder>] [--out DIR]
"""
import argparse, os, re, sys
import openpyxl
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape

NAVY   = (0.086, 0.137, 0.247)   # #16233F header / total bands
GOLD   = (0.906, 0.725, 0.235)   # #E7B93C rules + eyebrow + cluster tag
GOLD_D = (0.722, 0.525, 0.043)   # #B8860B target values on body rows
BONDNV = (0.122, 0.227, 0.420)   # #1F3A6B bond names
INK    = (0.169, 0.184, 0.220)   # #2B2F38 data
ZEBRA  = (0.961, 0.969, 0.984)   # #F5F7FB
TBAND  = (0.984, 0.949, 0.839)   # #FBF2D6 target column wash
HAIR   = (0.863, 0.890, 0.937)   # #DCE3EF row rules
WHITE  = (1, 1, 1)

X0, X1 = 26.0, 815.9
TARGET_W, BRAND_W = 70.0, 74.5
HDR_T, HDR_B = 30.0, 91.0
CH_T, CH_B   = 105.0, 138.0
ROW_H, TOT_H = 26.0, 28.0
PAD = 10.0


def _default_base():
    p = os.path.abspath(__file__)
    while p != "/":
        p = os.path.dirname(p)
        if os.path.isdir(os.path.join(p, "Targets")) and os.path.isfile(os.path.join(p, "CLAUDE.md")):
            return p
    return os.getcwd()


def read_workbook(path):
    """-> (brands, [(cluster_no, [(bond, [vals], total)])])"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    hrow = next(r for r in range(1, 12) if str(ws.cell(r, 1).value or "").strip().upper() == "BOND")
    cols, brands = [], []
    for c in range(2, ws.max_column + 1):
        v = ws.cell(hrow, c).value
        if not v:
            continue
        if str(v).strip().upper().startswith("TARGET"):
            tgt_col, tgt_hdr = c, str(v).strip()
            break
        cols.append(c); brands.append(str(v).strip())
    clusters, cur = [], None
    for r in range(hrow + 1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "").strip()
        if not label:
            continue
        m = re.fullmatch(r"CLUSTER\s+(\d+)", label, re.I)
        if m:
            cur = (int(m.group(1)), [])
            clusters.append(cur); continue
        if label.upper().startswith("GRAND"):
            break
        if cur is None:
            continue
        vals = [float(ws.cell(r, c).value or 0) for c in cols]
        tot  = float(ws.cell(r, tgt_col).value or 0)
        if abs(sum(vals) - tot) > 0.001:
            raise SystemExit(f"ABORT: {label} brands sum to {sum(vals):g} but TARGET reads {tot:g}")
        cur[1].append((label, vals, tot))
    return brands, tgt_hdr, clusters


def fmt(v):
    return "·" if not v else f"{v:,.0f}"


def draw(c, H, month, year, cno, rows, brands, tgt_hdr):
    def y(v):        return H - v
    def band(t, b, col):
        c.setFillColorRGB(*col); c.rect(X0, y(b), X1 - X0, b - t, stroke=0, fill=1)
    def rule(v, w, col, x0=X0, x1=X1):
        c.setStrokeColorRGB(*col); c.setLineWidth(w); c.line(x0, y(v), x1, y(v))

    bond_x1 = X1 - TARGET_W - BRAND_W * len(brands)
    centers = [bond_x1 + BRAND_W * (i + 0.5) for i in range(len(brands))] + [X1 - TARGET_W / 2]

    # ---- header band
    band(HDR_T, HDR_B, NAVY)
    c.setFillColorRGB(*GOLD); c.setFont("Helvetica-Bold", 9)
    c.drawString(X0 + 16, y(50.6), "K.S. DISTILLERY")
    c.setFillColorRGB(*WHITE); c.setFont("Helvetica-Bold", 16)
    c.drawString(X0 + 16, y(70.2), f"{month.upper()} {year} · TARGET")
    c.setFillColorRGB(*GOLD); c.setFont("Helvetica-Bold", 18)
    c.drawRightString(X1 - 16, y(70.5), f"CLUSTER {cno}")
    rule(HDR_B, 2.4, GOLD)

    # ---- column header
    band(CH_T, CH_B, NAVY)
    c.setFillColorRGB(*WHITE); c.setFont("Helvetica-Bold", 7.6)
    c.drawString(X0 + PAD, y((CH_T + CH_B) / 2 + 2.7), "BOND")
    for cx, name in zip(centers, brands + [tgt_hdr]):
        words = name.split()
        top = (CH_T + CH_B) / 2 - 9.0 * (len(words) - 1) / 2 + 2.7
        for i, w in enumerate(words):
            c.drawCentredString(cx, y(top + 9.0 * i), w)
    rule(CH_B, 1.4, GOLD)

    # ---- target column wash
    body_b = CH_B + ROW_H * len(rows)
    c.setFillColorRGB(*TBAND)
    c.rect(X1 - TARGET_W, y(body_b), TARGET_W, body_b - CH_B, stroke=0, fill=1)

    # ---- body rows
    for i, (bond, vals, tot) in enumerate(rows):
        t = CH_B + ROW_H * i
        if i % 2:
            c.setFillColorRGB(*ZEBRA)
            c.rect(X0, y(t + ROW_H), (X1 - TARGET_W) - X0, ROW_H, stroke=0, fill=1)
        base = y(t + ROW_H / 2 + 3.3)
        c.setFillColorRGB(*BONDNV); c.setFont("Helvetica-Bold", 9.2)
        c.drawString(X0 + PAD, base, bond)
        c.setFillColorRGB(*INK); c.setFont("Helvetica", 9.2)
        for cx, v in zip(centers, vals):
            c.drawCentredString(cx, base, fmt(v))
        c.setFillColorRGB(*GOLD_D); c.setFont("Helvetica-Bold", 9.2)
        c.drawCentredString(centers[-1], base, fmt(tot))
        rule(t, 0.4, HAIR)
    rule(body_b, 0.4, HAIR)
    rule(body_b, 1.4, GOLD)

    # ---- cluster total
    band(body_b, body_b + TOT_H, NAVY)
    base = y(body_b + TOT_H / 2 + 3.4)
    tot_vals = [sum(r[1][i] for r in rows) for i in range(len(brands))]
    c.setFillColorRGB(*WHITE); c.setFont("Helvetica-Bold", 9.4)
    c.drawString(X0 + PAD, base, "CLUSTER TOTAL")
    for cx, v in zip(centers, tot_vals):
        c.drawCentredString(cx, base, fmt(v))
    c.setFillColorRGB(*GOLD); c.setFont("Helvetica-Bold", 9.4)
    c.drawCentredString(centers[-1], base, fmt(sum(r[2] for r in rows)))
    c.showPage()


def draw_all(c, H, month, year, clusters, brands, tgt_hdr):
    """Cluster-rollup page for the ASM group — one row per cluster + grand total."""
    CH_T2, CH_B2 = 106.0, 139.6
    ROW_H2, TOT_H2, TARGET_W2 = 34.0, 36.0, 74.0

    def y(v):     return H - v
    def band(t, b, col):
        c.setFillColorRGB(*col); c.rect(X0, y(b), X1 - X0, b - t, stroke=0, fill=1)
    def rule(v, w, col):
        c.setStrokeColorRGB(*col); c.setLineWidth(w); c.line(X0, y(v), X1, y(v))

    bw = (X1 - TARGET_W2 - (X0 + 155.0)) / len(brands)
    centers = [X0 + 155.0 + bw * (i + 0.5) for i in range(len(brands))] + [X1 - TARGET_W2 / 2]

    band(HDR_T, HDR_B, NAVY)
    c.setFillColorRGB(*GOLD); c.setFont("Helvetica-Bold", 9)
    c.drawString(X0 + 16, y(50.6), "K.S. DISTILLERY")
    c.setFillColorRGB(*WHITE); c.setFont("Helvetica-Bold", 16)
    c.drawString(X0 + 16, y(70.2), f"{month.upper()} {year} · TARGET")
    rule(HDR_B, 2.4, GOLD)

    band(CH_T2, CH_B2, NAVY)
    c.setFillColorRGB(*WHITE); c.setFont("Helvetica-Bold", 7.8)
    c.drawString(X0 + 12, y((CH_T2 + CH_B2) / 2 + 2.8), "CLUSTER")
    for cx, name in zip(centers, brands + [tgt_hdr]):
        words = name.split()
        top = (CH_T2 + CH_B2) / 2 - 9.0 * (len(words) - 1) / 2 + 2.8
        for i, w in enumerate(words):
            c.drawCentredString(cx, y(top + 9.0 * i), w)
    rule(CH_B2, 1.4, GOLD)

    body_b = CH_B2 + ROW_H2 * len(clusters)
    c.setFillColorRGB(*TBAND)
    c.rect(X1 - TARGET_W2, y(body_b), TARGET_W2, body_b - CH_B2, stroke=0, fill=1)

    grand = [0.0] * len(brands)
    for i, (cno, rows) in enumerate(clusters):
        t = CH_B2 + ROW_H2 * i
        if i % 2:
            c.setFillColorRGB(*ZEBRA)
            c.rect(X0, y(t + ROW_H2), (X1 - TARGET_W2) - X0, ROW_H2, stroke=0, fill=1)
        base = y(t + ROW_H2 / 2 + 3.5)
        vals = [sum(r[1][k] for r in rows) for k in range(len(brands))]
        grand = [g + v for g, v in zip(grand, vals)]
        c.setFillColorRGB(*BONDNV); c.setFont("Helvetica-Bold", 10)
        c.drawString(X0 + 12, base, f"CLUSTER {cno}")
        c.setFillColorRGB(*INK); c.setFont("Helvetica", 9.6)
        for cx, v in zip(centers, vals):
            c.drawCentredString(cx, base, fmt(v))
        c.setFillColorRGB(*GOLD_D); c.setFont("Helvetica-Bold", 9.6)
        c.drawCentredString(centers[-1], base, fmt(sum(r[2] for r in rows)))
        rule(t, 0.5, HAIR)
    rule(body_b, 0.5, HAIR); rule(body_b, 1.4, GOLD)

    band(body_b, body_b + TOT_H2, NAVY)
    base = y(body_b + TOT_H2 / 2 + 3.6)
    c.setFillColorRGB(*WHITE); c.setFont("Helvetica-Bold", 10.2)
    c.drawString(X0 + 12, base, "GRAND TOTAL")
    for cx, v in zip(centers, grand):
        c.drawCentredString(cx, base, fmt(v))
    c.setFillColorRGB(*GOLD); c.setFont("Helvetica-Bold", 10.2)
    c.drawCentredString(centers[-1], base, fmt(sum(r[2] for _, rs in clusters for r in rs)))
    c.showPage()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", required=True)
    ap.add_argument("--year", default="2026")
    ap.add_argument("--base", default=None)
    ap.add_argument("--out", default=None)
    ap.add_argument("--clusters", default=None, help="comma list, e.g. 1,3")
    ap.add_argument("--all-only", action="store_true", help="build only the ALL CLUSTERS rollup")
    ap.add_argument("--no-all", action="store_true", help="skip the ALL CLUSTERS rollup")
    a = ap.parse_args()
    base = a.base or _default_base()
    mu = a.month.upper()
    folder = os.path.join(base, "Targets", f"{mu} TARGETS")
    src = os.path.join(folder, f"TARGET -{mu}.xlsx")
    if not os.path.isfile(src):
        raise SystemExit(f"ABORT: no target workbook at {src}")
    out = a.out or folder
    os.makedirs(out, exist_ok=True)
    brands, tgt_hdr, clusters = read_workbook(src)
    want = {int(x) for x in a.clusters.split(",")} if a.clusters else None
    W, H = landscape(A4)
    made = []
    for cno, rows in clusters:
        if a.all_only or (want and cno not in want):
            continue
        p = os.path.join(out, f"TARGET {mu} {a.year} - CLUSTER {cno}.pdf")
        c = canvas.Canvas(p, pagesize=(W, H))
        c.setTitle(f"K.S. Distillery — {mu} {a.year} Target — Cluster {cno}")
        draw(c, H, mu, a.year, cno, rows, brands, tgt_hdr)
        c.save()
        made.append((p, rows))
        print(f"  CLUSTER {cno}: {len(rows)} bonds, {sum(r[2] for r in rows):,.0f} cs -> {os.path.basename(p)}")
    if not a.no_all:
        p = os.path.join(out, f"TARGET {mu} {a.year} - ALL CLUSTERS.pdf")
        c = canvas.Canvas(p, pagesize=(W, H))
        c.setTitle(f"K.S. Distillery — {mu} {a.year} Target — All Clusters")
        draw_all(c, H, mu, a.year, clusters, brands, tgt_hdr)
        c.save()
        made.append((p, None))
        g = sum(r[2] for _, rs in clusters for r in rs)
        print(f"  ALL CLUSTERS: {len(clusters)} clusters, {g:,.0f} cs -> {os.path.basename(p)}")
    if not made:
        raise SystemExit("ABORT: nothing to build")
    print(f"\n  {len(made)} PDF(s) in {out}")


if __name__ == "__main__":
    main()
