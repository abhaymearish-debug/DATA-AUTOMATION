#!/usr/bin/env python3
"""
build_mobile_summary.py — KSD MOBILE SUMMARY.xlsx (created 8 Jul 2026, Abhay-approved)

Purpose: a compact, values-only workbook at the Claude folder root that Google Drive
for Desktop mirrors to Abhay's Drive, so phone/remote Cowork sessions (which cannot
reach the Mac's disk) can answer day-to-day sales questions via the Drive connector's
read_file_content. The big analysis workbooks are too large / formula-heavy to read
through the connector; this file is the phone-query surface.

Sources (auto-discovered, read-only):
  - newest `KSBC shop sales/*ANALYSIS.xlsx`      (tertiary — bond/shop/daily/brand-pack)
  - newest `Secondary sales/*SECONDARY SALES ANALYSIS.xlsx` (dispatches — bond/brand/pack/daily)
  - `Warehouse stock/_history/stock_history.csv` (latest snapshot per warehouse)
  - `MASTER DATA CONFIRMED.xlsx` sheet 16-4-25   (shop -> bond mapping)

Output: `<base>/KSD MOBILE SUMMARY.xlsx` — 10 sheets, all static values, no formulas.
Non-fatal by design: each section builds independently; a missing source produces a
stub sheet with a notice instead of aborting. Exit 0 on save, 1 only if nothing saved.

Wired into scheduled tasks ksbc-shop-sales / secondary-sales / warehouse-stock as a
post-build step (failure of this step must never fail the build itself).
"""
import argparse, csv, os, re, sys, tempfile
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BONDS = ["KOLLAM","KOZHIKODE","ATTINGAL","PALAKKAD","KOTTARAKARA","KANNUR","ALAPPUZHA",
         "NEDUMANGAD","ALUVA","PERINTHALMANNA","THODUPUZHA","PATHANAMTHITTA","KOTTAYAM",
         "THRISSUR","TRIPUNITHURA"]
MONTHS = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER",
          "OCTOBER","NOVEMBER","DECEMBER"]

NAVY = "FF1A237E"; GOLD = "FFFFB300"; GREY = "FF6B7280"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(name="Aptos Narrow", size=11, bold=True, color="FFFFFFFF")
TITLE_FONT = Font(name="Aptos Narrow", size=14, bold=True, color=NAVY)
SUB_FONT = Font(name="Aptos Narrow", size=10, italic=True, color=GREY)
BASE_FONT = Font(name="Aptos Narrow", size=10)
BOLD_FONT = Font(name="Aptos Narrow", size=10, bold=True)

def find_base(cli):
    if cli: return Path(cli)
    p = Path(__file__).resolve()
    for parent in p.parents:
        if (parent / "KSBC shop sales").is_dir(): return parent
    sys.exit("Could not locate Claude base folder (no 'KSBC shop sales' in any parent).")

def newest(folder, pattern):
    cands = [f for f in folder.glob(pattern) if not f.name.startswith("~$")]
    return max(cands, key=lambda f: f.stat().st_mtime) if cands else None

def num(v):
    try:
        return float(v) if v is not None else 0.0
    except (TypeError, ValueError):
        return 0.0

def rating(sell):
    if sell is None: return "— No activity"
    if sell >= 80: return "🚀 High Performance"
    if sell >= 60: return "✅ Balanced"
    if sell >= 40: return "⚠️ Inventory Heavy"
    return "🚫 Critical Overstock"

def shop_key(v):
    """Normalise a shop code to a plain digit string."""
    if v is None: return None
    s = str(v).strip()
    if s.endswith(".0"): s = s[:-2]
    return s if s.isdigit() else None

# ------------------------------------------------------------------ writers
def style_sheet(ws, title, subtitle, headers, widths, freeze="A4"):
    ws["A1"] = title; ws["A1"].font = TITLE_FONT
    ws["A2"] = subtitle; ws["A2"].font = SUB_FONT
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = 26
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False
    return 4  # first data row

def write_rows(ws, r0, rows, bold_rows=()):
    for i, row in enumerate(rows):
        r = r0 + i
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = BOLD_FONT if i in bold_rows else BASE_FONT
            if isinstance(v, float): cell.number_format = "#,##0.00"

def stub(ws, title, msg):
    ws["A1"] = title; ws["A1"].font = TITLE_FONT
    ws["A3"] = msg; ws["A3"].font = SUB_FONT

# ------------------------------------------------------------------ sections
def load_master(base, warnings):
    m = {}
    try:
        wb = openpyxl.load_workbook(base / "MASTER DATA CONFIRMED.xlsx", read_only=True, data_only=True)
        ws = wb["16-4-25"]
        for row in ws.iter_rows(min_row=3, values_only=True):
            code = shop_key(row[3])
            if code:
                m[code] = {"name": str(row[4] or "").strip(), "cat": str(row[5] or "").strip(),
                           "staff": str(row[6] or "").strip(), "bond": str(row[7] or "").strip(),
                           "status": str(row[8] or "").strip()}
        wb.close()
        if "11014" in m: m.setdefault("111014", m["11014"])  # MUKKAM legacy alias, defensive
    except Exception as e:
        warnings.append(f"master load failed: {e}")
    return m

def build_ksbc(out, base, master, warnings, summary):
    src = newest(base / "KSBC shop sales", "*ANALYSIS.xlsx")
    if not src:
        for nm in ("KSBC BONDS","KSBC DAILY BY BOND","KSBC SHOPS","KSBC BRAND-PACK"):
            stub(out.create_sheet(nm), nm, "No KSBC analysis workbook found.")
        warnings.append("KSBC workbook missing"); return
    month = next((mn for mn in MONTHS if src.name.upper().startswith(mn)), None)
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    period = str(wb["BOND PERFORMANCE"].cell(row=1, column=1).value or "").split("—")[-1].strip()
    summary["ksbc_src"] = f"{src.name} ({period})"

    # --- KSBC BONDS (from BOND PERFORMANCE; %/rating recomputed since formulas aren't cached)
    bonds = []
    for row in wb["BOND PERFORMANCE"].iter_rows(min_row=3, max_col=5, values_only=True):
        name = str(row[0] or "").strip().upper()
        if name in BONDS:
            o, rcv, s, c = (num(x) for x in row[1:5])
            st = round(s / (o + rcv) * 100, 1) if (o + rcv) > 0 else None
            bonds.append([name, round(o,2), round(rcv,2), round(s,2), round(c,2),
                          st if st is not None else "—", rating(st)])
    bonds.sort(key=lambda r: -r[3])
    tot = ["TOTAL"] + [round(sum(b[i] for b in bonds),2) for i in (1,2,3,4)]
    st = round(tot[3]/(tot[1]+tot[2])*100,1) if (tot[1]+tot[2])>0 else None
    tot += [st if st is not None else "—", rating(st)]
    ws = out.create_sheet("KSBC BONDS")
    r0 = style_sheet(ws, f"KSBC TERTIARY SALES BY BOND — {period}",
                     f"Source: {src.name} · cases · sorted by sales desc",
                     ["Bond","Opening","Receipts","Sales","Closing","Sell-Through %","Rating"],
                     [18,11,11,11,11,14,22])
    write_rows(ws, r0, bonds + [tot], bold_rows={len(bonds)})
    if len(bonds) != 15: warnings.append(f"KSBC BONDS: {len(bonds)} bonds (expected 15)")
    summary["ksbc_total_sales"] = tot[3]

    # --- KSBC DAILY BY BOND (parse daily raw sheets, fold bottles/BPC, map via master)
    daily = {}; unmatched = set()
    day_sheets = sorted((s for s in wb.sheetnames
                         if month and re.fullmatch(rf"{month} (\d+)", s)),
                        key=lambda s: int(s.split()[-1]))
    for sn in day_sheets:
        d = int(sn.split()[-1]); per_bond = {}
        for row in wb[sn].iter_rows(min_row=2, max_col=13, values_only=True):
            code = shop_key(row[1])
            if not code: continue
            bpc = num(row[6]); sold = num(row[11]) + (num(row[12])/bpc if bpc else 0.0)
            if not sold: continue
            info = master.get(code)
            bond = info["bond"] if info else "UNMATCHED"
            if not info: unmatched.add(code)
            per_bond[bond] = per_bond.get(bond, 0.0) + sold
        daily[d] = per_bond
    cols = BONDS + (["UNMATCHED"] if any("UNMATCHED" in v for v in daily.values()) else [])
    ws = out.create_sheet("KSBC DAILY BY BOND")
    r0 = style_sheet(ws, f"KSBC DAILY SALES BY BOND — {month or ''} 2026",
                     "cases sold per day (from daily raw sheets; cumulative-period corrections not reflected)",
                     ["Date"] + cols + ["TOTAL"], [12] + [13]*len(cols) + [12])
    rows = []
    for d in sorted(daily):
        vals = [round(daily[d].get(b, 0.0), 2) for b in cols]
        rows.append([f"{d} {month.title()}"] + vals + [round(sum(vals), 2)])
    tot_row = ["MTD TOTAL"] + [round(sum(daily[d].get(b,0.0) for d in daily),2) for b in cols]
    tot_row += [round(sum(tot_row[1:]),2)]
    write_rows(ws, r0, rows + [tot_row], bold_rows={len(rows)})
    if unmatched: warnings.append(f"daily sheets: {len(unmatched)} unmatched shop codes {sorted(unmatched)[:5]}")

    # --- KSBC SHOPS (from the 15 region sheets)
    shops = []
    for b in BONDS:
        if b not in wb.sheetnames: continue
        for row in wb[b].iter_rows(min_row=5, max_col=7, values_only=True):
            code = shop_key(row[0])
            if not code: continue
            o, rcv, s, c = (num(x) for x in row[3:7])
            st = round(s/(o+rcv)*100,1) if (o+rcv)>0 else None
            shops.append([code, str(row[1] or "").strip(), b, str(row[2] or "").strip(),
                          round(o,2), round(rcv,2), round(s,2), round(c,2),
                          st if st is not None else "—", rating(st)])
    shops.sort(key=lambda r: -r[6])
    ws = out.create_sheet("KSBC SHOPS")
    r0 = style_sheet(ws, f"KSBC SHOP-WISE (MTD) — {period}",
                     f"{len(shops)} active shops · cases · sorted by sales desc",
                     ["Shop Code","Shop Name","Bond","Field Staff","Opening","Receipts",
                      "Sales","Closing","Sell-Through %","Rating"],
                     [10,30,15,16,10,10,10,10,13,22])
    write_rows(ws, r0, shops)
    summary["ksbc_shops"] = len(shops)

    # --- KSBC BRAND-PACK (aggregate the COMBINED roll-up)
    comb = next((s for s in wb.sheetnames if "COMBINED" in s and "CUMULATIVE" not in s), None)
    ws = out.create_sheet("KSBC BRAND-PACK")
    if comb:
        agg = {}
        for row in wb[comb].iter_rows(min_row=2, max_col=11, values_only=True):
            if not row[4]: continue
            key = (str(row[4]).strip(), str(row[5] or "").strip())
            a = agg.setdefault(key, [0.0]*4)
            for i, col in enumerate((7,8,9,10)): a[i] += num(row[col])
        rows = [[k[0], k[1]] + [round(x,2) for x in v] for k, v in agg.items()]
        rows.sort(key=lambda r: -r[4])
        tot = ["TOTAL",""] + [round(sum(r[i] for r in rows),2) for i in (2,3,4,5)]
        r0 = style_sheet(ws, f"KSBC BRAND × PACK (MTD) — {period}",
                         f"Source sheet: {comb} · cases",
                         ["Brand","Pack","Opening","Receipts","Sales","Closing"],
                         [34,10,11,11,11,11])
        write_rows(ws, r0, rows + [tot], bold_rows={len(rows)})
    else:
        stub(ws, "KSBC BRAND × PACK", "No COMBINED sheet found in the KSBC workbook.")
    wb.close()

def build_secondary(out, base, warnings, summary):
    src = newest(base / "Secondary sales", "*SECONDARY SALES ANALYSIS.xlsx")
    if not src:
        for nm in ("SECONDARY BONDS","SECONDARY BRANDS","SECONDARY PACKS","SECONDARY DAILY"):
            stub(out.create_sheet(nm), nm, "No Secondary analysis workbook found.")
        warnings.append("Secondary workbook missing"); return
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    period = str(wb["BOND PERFORMANCE"].cell(row=1, column=1).value or "").split("—")[-1].strip()
    summary["sec_src"] = f"{src.name} ({period})"

    def table(sheet, title, headers, widths):
        ws_out = out.create_sheet(title)
        if sheet not in wb.sheetnames:
            stub(ws_out, title, f"Sheet '{sheet}' not found."); return None
        rows = []
        for row in wb[sheet].iter_rows(min_row=4, max_col=6, values_only=True):
            label = str(row[0] or "").strip()
            if not label: continue
            if "TOTAL" in label.upper(): continue
            vals = [label] + [round(num(x),2) for x in row[1:5]]
            if len(headers) > 5:  # share col: render as %
                vals.append(round(num(row[5])*100,1))
            rows.append(vals)
        tot = ["TOTAL"] + [round(sum(r[i] for r in rows),2) for i in range(1,5)]
        if len(headers) > 5: tot.append(100.0)
        r0 = style_sheet(ws_out, f"{title} — {period}", f"Source: {src.name} · cases",
                         headers, widths)
        write_rows(ws_out, r0, rows + [tot], bold_rows={len(rows)})
        return rows
    b = table("BOND PERFORMANCE", "SECONDARY BONDS",
              ["Bond","KSBC (cs)","Consumer fed (cs)","BAR (cs)","Total (cs)"], [18,12,15,10,12])
    table("BRAND PERFORMANCE", "SECONDARY BRANDS",
          ["Brand","KSBC","Consumer fed","BAR","Total","% of Grand Total"], [36,10,13,10,11,15])
    table("PACK PERFORMANCE", "SECONDARY PACKS",
          ["Pack","KSBC","Consumer fed","BAR","Total","% of Grand Total"], [14,10,13,10,11,15])
    if b: summary["sec_total"] = round(sum(r[4] for r in b),2)

    ws = out.create_sheet("SECONDARY DAILY")
    if "DAILY TREND" in wb.sheetnames:
        rows = [[str(r[0]), str(r[1] or ""), round(num(r[2]),2)]
                for r in wb["DAILY TREND"].iter_rows(min_row=4, max_col=3, values_only=True)
                if r[0]]
        r0 = style_sheet(ws, f"SECONDARY DAILY DISPATCH TREND — {period}",
                         f"Source: {src.name} · cases invoiced per day",
                         ["Date","Day","Cases Invoiced"], [13,12,14])
        write_rows(ws, r0, rows)
    else:
        stub(ws, "SECONDARY DAILY", "No DAILY TREND sheet found.")
    wb.close()

def build_warehouse(out, base, warnings, summary):
    ws = out.create_sheet("WAREHOUSE STOCK")
    hist = base / "Warehouse stock" / "_history" / "stock_history.csv"
    if not hist.exists():
        stub(ws, "WAREHOUSE STOCK", "stock_history.csv not found."); warnings.append("stock history missing"); return
    latest = {}
    with open(hist, newline="") as f:
        for row in csv.DictReader(f):
            w = row["warehouse"].strip()
            if w not in latest or row["date"] > latest[w]["date"]: latest[w] = row
    rows = [[w, d["date"], round(num(d["physical"]),1), round(num(d["allotable"]),1),
             round(num(d["pending"]),1)] for w, d in latest.items()]
    rows.sort(key=lambda r: -r[2])
    tot = ["NETWORK TOTAL","", *[round(sum(r[i] for r in rows),1) for i in (2,3,4)]]
    as_of = max(r[1] for r in rows) if rows else "?"
    r0 = style_sheet(ws, f"BEVCO WAREHOUSE STOCK — latest snapshot per warehouse (as of {as_of})",
                     "cases · per-warehouse own most-recent snapshot · Physical = on-hand, Allotable = free to allot, Pending = allotted not lifted",
                     ["Warehouse","Snapshot Date","Physical","Allotable","Pending"], [26,14,11,11,11])
    write_rows(ws, r0, rows + [tot], bold_rows={len(rows)})
    summary["wh"] = f"{len(rows)} warehouses, network physical {tot[2]:,.0f} cs (as of {as_of})"

def build_readme(out, summary, warnings):
    ws = out.create_sheet("README", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 118
    lines = [
        ("K.S. DISTILLERY — MOBILE SUMMARY", TITLE_FONT),
        (f"Generated {datetime.now().strftime('%d %b %Y, %H:%M')} by build_mobile_summary.py. "
         "All figures in CASES. Values only (no formulas).", SUB_FONT),
        ("", None),
        ("WHAT THIS FILE IS", BOLD_FONT),
        ("A compact snapshot of the three daily K.S. Distillery data streams, rebuilt after every "
         "pipeline run and mirrored to Google Drive so questions can be answered from any device "
         "without the Mac. For deeper drill-downs (DETAIL sheets, aging, PI, incentives) the full "
         "workbooks in the Claude folder are the source of truth.", BASE_FONT),
        ("", None),
        ("SHEETS", BOLD_FONT),
        ("KSBC BONDS / KSBC SHOPS — month-to-date tertiary sales (KSBC shop → consumer) by bond and by shop.", BASE_FONT),
        ("KSBC DAILY BY BOND — per-day tertiary sales matrix; the last date row = the latest day's sales ('today').", BASE_FONT),
        ("KSBC BRAND-PACK — month-to-date tertiary sales by brand × pack size.", BASE_FONT),
        ("SECONDARY BONDS / BRANDS / PACKS / DAILY — warehouse → outlet dispatches (KSBC + Consumer fed + BAR).", BASE_FONT),
        ("WAREHOUSE STOCK — latest Bevco warehouse position (physical / allotable / pending).", BASE_FONT),
        ("", None),
        ("TERMINOLOGY", BOLD_FONT),
        ("Tertiary sale = KSBC shop → consumer. Secondary sale = warehouse → any outlet. "
         "Invoice sale = warehouse → Consumer fed or BAR only. "
         "Total liquidation = tertiary + Consumer fed + BAR (never add the KSBC dispatch leg on top of tertiary).", BASE_FONT),
        ("Ratings: ≥80% 🚀 High Performance · ≥60% ✅ Balanced · ≥40% ⚠️ Inventory Heavy · else 🚫 Critical Overstock. "
         "Sell-Through % = Sales ÷ (Opening + Receipts).", BASE_FONT),
        ("", None),
        ("DATA AS OF", BOLD_FONT),
        (f"KSBC: {summary.get('ksbc_src','—')}", BASE_FONT),
        (f"Secondary: {summary.get('sec_src','—')}", BASE_FONT),
        (f"Warehouse: {summary.get('wh','—')}", BASE_FONT),
    ]
    if warnings:
        lines += [("", None), ("BUILD WARNINGS", BOLD_FONT)] + [(f"• {w}", BASE_FONT) for w in warnings]
    for i, (txt, font) in enumerate(lines, 1):
        c = ws.cell(row=i, column=1, value=txt)
        if font: c.font = font
        c.alignment = Alignment(wrap_text=True, vertical="top")

# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base"); ap.add_argument("--out")
    a = ap.parse_args()
    base = find_base(a.base)
    out_path = Path(a.out) if a.out else base / "KSD MOBILE SUMMARY.xlsx"

    warnings, summary = [], {}
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    master = load_master(base, warnings)
    build_ksbc(wb, base, master, warnings, summary)
    build_secondary(wb, base, warnings, summary)
    build_warehouse(wb, base, warnings, summary)
    build_readme(wb, summary, warnings)

    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False,
                                      dir=str(out_path.parent)); tmp.close()
    try:
        wb.save(tmp.name); os.replace(tmp.name, out_path)
    except PermissionError:
        os.unlink(tmp.name)
        sys.exit(f"BLOCKED: {out_path.name} is open/locked (Excel ~$ lock?). Close it and re-run.")
    print(f"SAVED {out_path}")
    print(f"  KSBC:      {summary.get('ksbc_src','—')} · MTD sales {summary.get('ksbc_total_sales','?')} cs · {summary.get('ksbc_shops','?')} shops")
    print(f"  Secondary: {summary.get('sec_src','—')} · MTD dispatches {summary.get('sec_total','?')} cs")
    print(f"  Warehouse: {summary.get('wh','—')}")
    for w in warnings: print(f"  WARN: {w}")

if __name__ == "__main__":
    main()
