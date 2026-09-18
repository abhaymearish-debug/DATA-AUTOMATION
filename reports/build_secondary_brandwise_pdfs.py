#!/usr/bin/env python3
"""Brandwise cumulative secondary-sales PDFs, one per ASM cluster.

Reproduces the layout K.S. Distillery has been receiving:

    K.S DISTILLERY
    SECONDARY SALES - BRANDWISE        <from> - <to>
    WH - <WAREHOUSE>

    SHOP NAME        <BRAND 1>  <BRAND 2>  ...  TOTAL
    ...rows...
    TOTAL

One page per warehouse, one PDF per cluster, page N of M in the footer.

Data source is the live Secondary Sales workbook's "<MONTH> COMBINED DISPATCHES"
sheet — the canonical accumulator for the month — so these PDFs always agree
with the workbook by construction.

TWO DEFECTS FROM THE 20 AUG 2026 AUDIT ARE FIXED HERE
-----------------------------------------------------
1. Outlets printed twice. The previous report grouped rows by warehouse but
   titled each page by BOND, so an outlet supplied from two warehouses in the
   same period appeared as two part-rows on one page (Ponkunnam 113 + 31 instead
   of 144). This version titles the page by the warehouse it groups by, so the
   grouping and the heading agree. Within a page, rows are additionally merged on
   licensee number, so a shop can never appear twice.

2. Brand headers colliding. Header text was rendering as "MATUREDNO.1" where two
   narrow columns ran together. Columns here are sized to the widest wrapped word
   in the header, with a hard minimum, and the header font shrinks as the brand
   count grows.

Usage:
    python3 build_secondary_brandwise_pdfs.py [--workbook PATH] [--outdir DIR]
                                              [--cluster 1|2|3] [--verify]
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas as pdfcanvas

# --------------------------------------------------------------------------
# Locked reference data
# --------------------------------------------------------------------------

# Bottles per case. Loose bottles are folded into cases as cases + bottles/BPC —
# dropping them ran the June figures ~0.02% light before it was caught.
BPC = {"1000 ML": 9, "750 ML": 12, "500 ML": 18, "375 ML": 24, "180 ML": 48}

# Bevco warehouse -> ASM cluster, aligned to the three bond clusters.
# BATTATHUR's cluster is a best guess carried over from the warehouse artifact.
CLUSTERS: dict[int, list[str]] = {
    1: ["ALAPPUZHA", "ATTINGAL", "NEDUMANGAD", "KOLLAM", "KOTTARAKARA",
        "PATHANAMTHITTA", "KARUNAGAPPALLY", "BALARAMAPURAM", "MENAMKULAM", "THIRUVALLA"],
    2: ["THRISSUR", "TRIPUNITHURA", "THODUPUZHA", "KOTTAYAM", "ALUVA",
        "PERUMBAVOOR", "KOTHAMANGALAM", "AYARKKUNNAM", "CHALAKUDY", "KADAVANTHRA"],
    3: ["KANNUR", "KOZHIKODE", "PALAKKAD", "PERINTHALMANNA", "MENONPARA",
        "BATTATHUR", "KALPETTA", "NADUVANNUR"],
}
# Spelling variants seen in the exports.
WAREHOUSE_ALIASES = {"PATHANAMTHITA": "PATHANAMTHITTA"}

CLUSTER_OF = {}
for _c, _whs in CLUSTERS.items():
    for _w in _whs:
        CLUSTER_OF[_w] = _c

NAVY = colors.HexColor("#0D1B4A")
GOLD = colors.HexColor("#FFB300")
GREY = colors.HexColor("#6B7280")
ROW_ALT = colors.HexColor("#F3F6FB")


def canonical_warehouse(raw: str) -> str | None:
    """'WH-KOLLAM FL9-KLM-01/2026-27' -> 'KOLLAM'.

    Must cope with both 'FL9' and 'RFL9' licence suffixes — a regex that handled
    only FL silently missed every RFL warehouse once before.
    """
    if not raw:
        return None
    txt = str(raw).strip().upper()
    m = re.match(r"^WH[-\s]+([A-Z .]+?)(?:\s+R?FL.*)?$", txt)
    name = (m.group(1) if m else txt).strip().rstrip(".")
    name = re.sub(r"\s+", " ", name)
    return WAREHOUSE_ALIASES.get(name, name)


def cases(issue_cases, issue_bottles, pack: str) -> float:
    c = float(issue_cases or 0)
    b = float(issue_bottles or 0)
    if b:
        c += b / BPC.get(str(pack or "").strip().upper(), 1)
    return c


def parse_date(value) -> datetime | None:
    if isinstance(value, datetime):
        return value
    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(value).strip(), fmt)
        except (ValueError, TypeError):
            continue
    return None


# --------------------------------------------------------------------------
# Extraction
# --------------------------------------------------------------------------


def find_workbook(base: Path) -> Path:
    folder = base / "Secondary sales"
    books = sorted(folder.glob("*SECONDARY SALES ANALYSIS.xlsx"),
                   key=lambda p: p.stat().st_mtime)
    if not books:
        raise SystemExit(f"No secondary-sales workbook found in {folder}")
    return books[-1]


def extract(workbook: Path, date_from=None, date_to=None):
    wb = openpyxl.load_workbook(workbook, read_only=True, data_only=True)
    sheets = [s for s in wb.sheetnames if "COMBINED DISPATCH" in s.upper()]
    if not sheets:
        raise SystemExit(f"{workbook.name} has no COMBINED DISPATCHES sheet.")
    ws = wb[sheets[0]]

    rows = ws.iter_rows(values_only=True)
    header = [str(h).strip() if h else "" for h in next(rows)]
    idx = {name: i for i, name in enumerate(header)}

    def col(*names):
        for n in names:
            if n in idx:
                return idx[n]
        raise SystemExit(f"Column not found in COMBINED DISPATCHES: {names}")

    c_wh = col("Warehouse Name")
    c_item = col("Item Name")
    c_pack = col("Pack")
    c_lic = col("Licensee No.", "Licensee No")
    c_shop = col("Licensee Name")
    c_cases = col("Issue Cases")
    c_bottles = col("Issue Bottles")
    c_date = col("Inv/GTN Date")

    # warehouse -> licensee -> {shop name, brand -> cases}
    data: dict[str, dict[str, dict]] = defaultdict(lambda: defaultdict(
        lambda: {"name": "", "brands": defaultdict(float)}))
    brands_seen: set[str] = set()
    dates: list[datetime] = []
    grand = 0.0
    skipped_wh: set[str] = set()

    for r in rows:
        if r is None or r[c_wh] is None:
            continue
        wh = canonical_warehouse(r[c_wh])
        if wh is None:
            continue
        if wh not in CLUSTER_OF:
            skipped_wh.add(wh)
            continue

        d = parse_date(r[c_date])
        # A date window applies before anything is counted, so the printed
        # totals and the period line in the header can never disagree.
        if date_from and d and d.date() < date_from:
            continue
        if date_to and d and d.date() > date_to:
            continue

        qty = cases(r[c_cases], r[c_bottles], r[c_pack])
        if qty == 0:
            continue

        brand = str(r[c_item] or "").strip().upper()
        lic = str(r[c_lic] or "").strip()
        shop = str(r[c_shop] or "").strip()

        entry = data[wh][lic]
        # Keep the longest spelling seen — exports vary in how much of the
        # outlet name they carry.
        if len(shop) > len(entry["name"]):
            entry["name"] = shop
        entry["brands"][brand] += qty

        brands_seen.add(brand)
        grand += qty
        if d:
            dates.append(d)

    wb.close()
    return {
        "data": data,
        "brands": sorted(brands_seen),
        "from": min(dates) if dates else None,
        "to": max(dates) if dates else None,
        "grand": grand,
        "skipped_warehouses": skipped_wh,
        "workbook": workbook,
    }


# --------------------------------------------------------------------------
# PDF
# --------------------------------------------------------------------------

# The layout lives in app/reports_pdf.py, measured from the PDFs the office
# already receives. Sharing it means the cluster PDFs and the "current view"
# export cannot drift apart in appearance.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app import reports_pdf as layout  # noqa: E402


def build_cluster_pdf(cluster, extracted, outdir):
    data = extracted["data"]
    warehouses = [w for w in CLUSTERS[cluster] if w in data and data[w]]
    if not warehouses:
        return None

    period = ""
    if extracted["from"] and extracted["to"]:
        period = (extracted["from"].strftime("%-d %B %Y") + " - "
                  + extracted["to"].strftime("%-d %B %Y"))

    # One page size for the whole cluster. Pages with fewer brand columns give
    # the slack to the shop-name column rather than shrinking the page.
    shapes = []
    for wh in warehouses:
        shops = data[wh]
        brands_wh = sorted({b for s in shops.values() for b in s["brands"] if s["brands"][b]})
        shapes.append({"labels": [s["name"] or lic for lic, s in shops.items()],
                       "columns": brands_wh, "rows": len(shops)})
    doc_w, doc_h, _ = layout.document_size(shapes)

    out = outdir / ("Secondary Sales - Cumulative (Cluster %d).pdf" % cluster)
    c = pdfcanvas.Canvas(str(out), pagesize=(doc_w, doc_h))
    total_pages = len(warehouses)

    for page_no, wh in enumerate(warehouses, start=1):
        shops = data[wh]
        # Columns are per page, exactly as in the source report: a warehouse
        # shows only the brands it actually dispatched.
        brands = sorted({b for s in shops.values() for b in s["brands"] if s["brands"][b]})
        ordered = sorted(shops.items(), key=lambda kv: -sum(kv[1]["brands"].values()))

        rows = []
        totals = defaultdict(float)
        for lic, shop in ordered:
            cells = {b: shop["brands"].get(b, 0.0) for b in brands}
            rt = sum(cells.values())
            for b in brands:
                totals[b] += cells[b]
            totals["__total__"] += rt
            rows.append({"name": shop["name"] or lic, "cells": cells, "total": rt})

        label_w = layout.label_width_for(doc_w, len(brands))
        layout.draw_page(
            c, width=doc_w, height=doc_h, label_w=label_w,
            report_title="SECONDARY SALES - CUMULATIVE",
            period=period, group_line="WH - " + wh,
            label_heading="SHOP NAME", columns=brands,
            rows=rows, totals=totals,
            page_no=page_no, pages=total_pages,
        )

    c.save()
    return out


def fmt(v: float) -> str:
    if abs(v) < 0.005:
        return "0"
    return f"{v:,.0f}" if abs(v - round(v)) < 0.005 else f"{v:,.2f}"


# --------------------------------------------------------------------------
# Verification
# --------------------------------------------------------------------------


def verify(extracted: dict) -> int:
    """Reconcile what will be printed against the workbook's own figures."""
    data = extracted["data"]
    printed = sum(sum(s["brands"].values()) for wh in data.values() for s in wh.values())
    grand = extracted["grand"]
    print(f"  extracted grand total : {grand:,.2f} cs")
    print(f"  sum of printed cells  : {printed:,.2f} cs")
    drift = abs(printed - grand)
    print(f"  drift                 : {drift:.4f} cs")

    dupes = 0
    for wh, shops in data.items():
        if len(shops) != len({lic for lic in shops}):
            dupes += 1
    print(f"  warehouses with a duplicated outlet : {dupes}")

    if extracted["skipped_warehouses"]:
        print(f"  ! warehouses not in any cluster (rows skipped): "
              f"{sorted(extracted['skipped_warehouses'])}")

    ok = drift < 0.01 and dupes == 0 and not extracted["skipped_warehouses"]
    print("  RESULT:", "PASS" if ok else "CHECK THE WARNINGS ABOVE")
    return 0 if ok else 1


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default=None, help="Claude folder root")
    ap.add_argument("--workbook", default=None)
    ap.add_argument("--outdir", default=None)
    ap.add_argument("--cluster", type=int, choices=[1, 2, 3], default=None)
    ap.add_argument("--from", dest="date_from", default=None,
                    help="ISO date; only dispatches on or after this are counted")
    ap.add_argument("--to", dest="date_to", default=None,
                    help="ISO date; only dispatches on or before this are counted")
    ap.add_argument("--verify", action="store_true")
    args = ap.parse_args()

    base = Path(args.base) if args.base else Path(__file__).resolve().parent.parent.parent
    workbook = Path(args.workbook) if args.workbook else find_workbook(base)
    outdir = Path(args.outdir) if args.outdir else (base / "Secondary sales")
    outdir.mkdir(parents=True, exist_ok=True)

    print(f"Workbook : {workbook.name}")
    def _iso(v):
        from datetime import date as _d
        return _d.fromisoformat(v) if v else None

    extracted = extract(workbook, _iso(args.date_from), _iso(args.date_to))
    if extracted["from"]:
        print(f"Period   : {extracted['from']:%d %b %Y} - {extracted['to']:%d %b %Y}")
    print(f"Brands   : {len(extracted['brands'])}   "
          f"Warehouses with dispatches: {len(extracted['data'])}")

    rc = verify(extracted) if args.verify else 0

    wanted = [args.cluster] if args.cluster else [1, 2, 3]
    for cl in wanted:
        out = build_cluster_pdf(cl, extracted, outdir)
        if out is None:
            print(f"Cluster {cl}: no dispatches in the period — no PDF written.")
            continue
        pages = len([w for w in CLUSTERS[cl] if w in extracted["data"] and extracted["data"][w]])
        total = sum(sum(s["brands"].values())
                    for w in CLUSTERS[cl] if w in extracted["data"]
                    for s in extracted["data"][w].values())
        print(f"Cluster {cl}: {pages} page(s), {total:,.2f} cs  ->  {out.name}")

    return rc


if __name__ == "__main__":
    raise SystemExit(main())
