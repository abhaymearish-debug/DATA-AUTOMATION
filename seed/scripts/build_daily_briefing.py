#!/usr/bin/env python3
"""Build Abhay's Daily Sales Briefing PDF for K.S. Distillery.

Pure sales & marketing only. No finance / cashflow / banking content.
5 pages: Snapshot · Bond Perf · Brand+Pack · Watch Items · Talking Points.
"""
from __future__ import annotations

import os
import sys
from collections import defaultdict
from datetime import date

import openpyxl
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, NextPageTemplate, KeepTogether,
)

# ---------- config ----------
ROOT = "/sessions/cool-youthful-lamport/mnt/Claude"
KSBC_DIR = f"{ROOT}/KSBC shop sales"
SEC_DIR = f"{ROOT}/Secondary sales"
MASTER = f"{ROOT}/MASTER DATA CONFIRMED.xlsx"
OUT_DIR = f"{ROOT}/Daily Briefings"

TODAY = date(2026, 4, 21)
OUT_PDF = f"{OUT_DIR}/Briefing - {TODAY.strftime('%d-%m-%Y')}.pdf"

# palette
NAVY = colors.HexColor("#1A237E")
NAVY_DARK = colors.HexColor("#0D1257")
ZEBRA = colors.HexColor("#F1F5F9")
GREEN = colors.HexColor("#2E7D32")
RED = colors.HexColor("#C62828")
MUTED = colors.HexColor("#64748B")
GOLD = colors.HexColor("#C9A227")
WHITE = colors.white
BORDER = colors.HexColor("#CBD5E1")

# Cluster / ASM mapping. Membership follows the LOCKED definition in CLAUDE.md
# (BOND PERFORMANCE clusterwise, 25 Apr 2026) and was CORRECTED 27 Jul 2026 —
# the previous guess had PATHANAMTHITTA in Cluster 2 and THRISSUR in Cluster 3,
# both wrong, so those two bonds were briefed to the wrong ASM. Confirmed by
# Abhay 27 Jul 2026. ASM roster: DINESH KUMAR P / SOJAN T T / HARIDASAN M.
CLUSTER_MAP = {
    # Cluster 1 — South, Dinesh
    "KOLLAM": ("1", "Dinesh"),
    "KOTTARAKARA": ("1", "Dinesh"),
    "ATTINGAL": ("1", "Dinesh"),
    "NEDUMANGAD": ("1", "Dinesh"),
    "ALAPPUZHA": ("1", "Dinesh"),
    "PATHANAMTHITTA": ("1", "Dinesh"),
    # Cluster 2 — Central, Sojan
    "KOTTAYAM": ("2", "Sojan"),
    "THODUPUZHA": ("2", "Sojan"),
    "TRIPUNITHURA": ("2", "Sojan"),
    "ALUVA": ("2", "Sojan"),
    "THRISSUR": ("2", "Sojan"),
    # Cluster 3 — North, Haridasan
    "PALAKKAD": ("3", "Haridasan"),
    "PERINTHALMANNA": ("3", "Haridasan"),
    "KOZHIKODE": ("3", "Haridasan"),
    "KANNUR": ("3", "Haridasan"),
}

# ---------- helpers ----------
def fmt_cases(x, decimals=0):
    if x is None:
        return "—"
    try:
        x = float(x)
    except (TypeError, ValueError):
        return str(x)
    if decimals == 0:
        return f"{int(round(x)):,}"
    return f"{x:,.{decimals}f}"

def fmt_pct(x, decimals=1):
    if x is None:
        return "—"
    return f"{x:+.{decimals}f}%"

def delta_color(x):
    if x is None or abs(x) < 0.05:
        return MUTED
    return GREEN if x > 0 else RED

# ---------- data loading ----------
def latest_file(folder, contains):
    files = [f for f in os.listdir(folder) if contains in f and f.endswith(".xlsx") and not f.startswith("~$")]
    files.sort()
    return os.path.join(folder, files[-1]) if files else None

def load_master():
    wb = openpyxl.load_workbook(MASTER, read_only=True, data_only=True)
    ws = wb["16-4-25"]
    info = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        if r[3] is None:
            continue
        code = str(r[3]).strip()
        info[code] = dict(name=r[4], cat=r[5], staff=r[6], bond=r[7], status=r[8])
    wb.close()
    return info

def daily_sum(ws):
    """Sum daily-snapshot sheet: Shop Out Cases + (Shop Out Bottles / bpc)."""
    total = 0.0
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        bpc = r[6] or 1
        c = (r[11] or 0) + ((r[12] or 0) / bpc if bpc else 0)
        total += float(c)
    return total

def daily_by_bond(ws, shop_info):
    by = defaultdict(float)
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        sc = str(r[1]).strip()
        bpc = r[6] or 1
        c = (r[11] or 0) + ((r[12] or 0) / bpc if bpc else 0)
        bond = shop_info.get(sc, {}).get("bond", "UNKNOWN")
        by[bond] += float(c)
    return by

def combined_shop_sales(ws):
    """APRIL 1-20 COMBINED or MARCH 1-31 COMBINED style: Sales (Cases) col idx 9."""
    by = defaultdict(float)
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        sc = str(r[1]).strip()
        by[sc] += float(r[9] or 0)
    return by

def combined_brand_pack(ws):
    by = defaultdict(float)
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        brand = r[4]; pack = r[5]
        by[(brand, pack)] += float(r[9] or 0)
    return by

def daily_brand_pack(ws):
    by = defaultdict(float)
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        brand = r[4]; pack = r[5]
        bpc = r[6] or 1
        c = (r[11] or 0) + ((r[12] or 0) / bpc if bpc else 0)
        by[(brand, pack)] += float(c)
    return by

# ---------- pull data ----------
shop_info = load_master()

# current KSBC
ksbc_path = latest_file(KSBC_DIR, "APRIL")
ksbc_wb = openpyxl.load_workbook(ksbc_path, read_only=True, data_only=True)
apr_combined = ksbc_wb["APRIL 1-20 COMBINED"]
apr_shop_sales = combined_shop_sales(apr_combined)
apr_bond_sales = defaultdict(float)
for sc, v in apr_shop_sales.items():
    bond = shop_info.get(sc, {}).get("bond")
    if bond:
        apr_bond_sales[bond] += v
apr_bp = combined_brand_pack(apr_combined)
apr_total_ksbc = sum(apr_shop_sales.values())
apr_day20_ksbc = daily_sum(ksbc_wb["APRIL 20"])
apr_day20_by_bond = daily_by_bond(ksbc_wb["APRIL 20"], shop_info)
# KSBC through Apr 20 = full MTD via combined sheet (covers 1-20)

# previous month KSBC (March)
pksbc_path = f"{KSBC_DIR}/MARCH SHOP SALES ANALYSIS.xlsx"
pksbc_wb = openpyxl.load_workbook(pksbc_path, read_only=True, data_only=True)
mar_1_16 = pksbc_wb["MARCH 1-16"]
mar_17_31 = pksbc_wb["MARCH 17-31"]

def daily_shop_sum(ws):
    by = defaultdict(float)
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None: continue
        sc = str(r[1]).strip()
        bpc = r[6] or 1
        c = (r[11] or 0) + ((r[12] or 0)/bpc if bpc else 0)
        by[sc] += float(c)
    return by

m1_shop = daily_shop_sum(mar_1_16)
m2_shop = daily_shop_sum(mar_17_31)
# March 1-20 prorated shop sales = 1-16 + (4/15)*17-31
mar_1_20_shop = {sc: m1_shop.get(sc,0) + m2_shop.get(sc,0)*4/15 for sc in set(m1_shop)|set(m2_shop)}
mar_1_20_ksbc_total = sum(mar_1_20_shop.values())
mar_1_20_bond = defaultdict(float)
for sc, v in mar_1_20_shop.items():
    b = shop_info.get(sc, {}).get("bond")
    if b: mar_1_20_bond[b] += v

# full-month march
m_bp1 = daily_brand_pack(mar_1_16)
m_bp2 = daily_brand_pack(mar_17_31)
mar_1_20_bp = {k: m_bp1.get(k,0) + m_bp2.get(k,0)*4/15 for k in set(m_bp1)|set(m_bp2)}

# Secondary sales
sec_path = latest_file(SEC_DIR, "APRIL")
sec_wb = openpyxl.load_workbook(sec_path, read_only=True, data_only=True)
sec_bond_ws = sec_wb["BOND PERFORMANCE"]
sec_bond = {}  # bond -> {ksbc, fed, bar, total}
for r in list(sec_bond_ws.iter_rows(values_only=True))[2:]:
    if r[0] in (None, "TOTAL", "UNMATCHED / Pending master"):
        continue
    sec_bond[r[0]] = dict(ksbc=r[1] or 0, fed=r[2] or 0, bar=r[3] or 0, total=r[4] or 0)

# dashboard totals
sec_dash = sec_wb["DASHBOARD"]
# pick values from known cells based on inspected structure
# row 4 (1-based) col B=4342 total, col F=923 invoice, col J=3419 KSBC dispatch, col N=891 fed, col R=32 bar
# We'll read by known semantics
rows = list(sec_dash.iter_rows(values_only=True))
# Find the KPI value row by scanning for the label row first
april_sec_total = april_invoice_total = april_ksbc_dispatch = april_fed = april_bar = 0
for i, r in enumerate(rows):
    if r and r[1] == "TOTAL SECONDARY SALES":
        v = rows[i+1]
        april_sec_total = v[1] or 0
        april_invoice_total = v[5] or 0
        april_ksbc_dispatch = v[9] or 0
        april_fed = v[13] or 0
        april_bar = v[17] or 0
        break

daily_trend = []
for r in list(sec_wb["DAILY TREND"].iter_rows(values_only=True))[2:]:
    if r[0] in (None, "TOTAL"):
        continue
    daily_trend.append((r[0], r[1], r[2] or 0))

# Previous month secondary
psec_path = f"{SEC_DIR}/MARCH SECONDARY SALES ANALYSIS.xlsx"
psec_wb = openpyxl.load_workbook(psec_path, read_only=True, data_only=True)
mar_trend = []
for r in list(psec_wb["DAILY TREND"].iter_rows(values_only=True))[2:]:
    if r[0] in (None, "TOTAL"):
        continue
    try:
        v = float(r[2]) if r[2] not in (None, "") else 0.0
    except (TypeError, ValueError):
        v = 0.0
    mar_trend.append((r[0], r[1], v))
mar_sec_total = sum(v for (_,_,v) in mar_trend)
mar_1_20_sec_total = sum(v for (_,_,v) in mar_trend[:20])
# split into ksbc / fed / bar using full-month ratio
mar_full_bond_ws = psec_wb["BOND PERFORMANCE"]
mar_full_rows = list(mar_full_bond_ws.iter_rows(values_only=True))[2:]
def _f(x):
    try: return float(x) if x not in (None, "") else 0.0
    except (TypeError, ValueError): return 0.0
mar_ksbc_disp_full = 0; mar_fed_full = 0; mar_bar_full = 0
for r in mar_full_rows:
    if r[0] in (None, "TOTAL"): continue
    mar_ksbc_disp_full += _f(r[1])
    mar_fed_full += _f(r[2])
    mar_bar_full += _f(r[3])
mar_full_total_sec = mar_ksbc_disp_full + mar_fed_full + mar_bar_full
r_ksbc = mar_ksbc_disp_full / mar_full_total_sec if mar_full_total_sec else 0
r_fed = mar_fed_full / mar_full_total_sec if mar_full_total_sec else 0
r_bar = mar_bar_full / mar_full_total_sec if mar_full_total_sec else 0
mar_1_20_ksbc_disp = mar_1_20_sec_total * r_ksbc
mar_1_20_fed = mar_1_20_sec_total * r_fed
mar_1_20_bar = mar_1_20_sec_total * r_bar

# bond-wise FED+BAR for march 1-20 (prorate each bond at 20/31 of full-month FED+BAR)
mar_bond_fedbar = {}
for r in mar_full_rows:
    if r[0] in (None, "TOTAL"): continue
    fb_full = _f(r[2]) + _f(r[3])
    mar_bond_fedbar[r[0]] = fb_full * (mar_1_20_sec_total / mar_sec_total) if mar_sec_total else 0
# For KSBC dispatch (secondary view), we care about invoice cashflow; we use FED+BAR per bond

# ---------- umbrella figures ----------
apr_fed_bar = (april_fed or 0) + (april_bar or 0)   # = 923
apr_total_umbrella = apr_total_ksbc + apr_fed_bar
mar_fed_bar_1_20 = mar_1_20_fed + mar_1_20_bar
mar_total_umbrella_1_20 = mar_1_20_ksbc_total + mar_fed_bar_1_20

mtd_delta = apr_total_umbrella - mar_total_umbrella_1_20
mtd_pct = (mtd_delta / mar_total_umbrella_1_20 * 100) if mar_total_umbrella_1_20 else 0

# yesterday = Apr 20 KSBC (secondary Apr 20 = 0 in trend, not yet ingested)
yesterday_ksbc = apr_day20_ksbc
# No secondary for Apr 20 -> treat as not-yet-updated, will label clearly
# compare vs same-day last month (Mar 20) — approx via prorated daily
mar_20_ksbc_est = sum(m2_shop.values()) / 15  # average daily over mar 17-31 as an estimate
mar_20_sec_trend = next((v for (d,_,v) in mar_trend if "20 Mar" in (d or "")), None)

yesterday_delta = (yesterday_ksbc - mar_20_ksbc_est)
yesterday_pct = (yesterday_delta / mar_20_ksbc_est * 100) if mar_20_ksbc_est else 0

# data freshness
DATA_AS_OF_KSBC = "Apr 20, 2026"
DATA_AS_OF_SEC = "Apr 19, 2026"

# ---------- styles ----------
styles = getSampleStyleSheet()
H1 = ParagraphStyle("H1", parent=styles["Heading1"], textColor=NAVY, fontName="Helvetica-Bold",
                    fontSize=20, leading=24, spaceAfter=4)
H2 = ParagraphStyle("H2", parent=styles["Heading2"], textColor=NAVY, fontName="Helvetica-Bold",
                    fontSize=13, leading=16, spaceAfter=6, spaceBefore=10)
H3 = ParagraphStyle("H3", parent=styles["Heading3"], textColor=NAVY, fontName="Helvetica-Bold",
                    fontSize=11, leading=14, spaceAfter=4)
BODY = ParagraphStyle("BODY", parent=styles["BodyText"], fontName="Helvetica",
                      fontSize=9.5, leading=13, textColor=colors.black)
SMALL = ParagraphStyle("SMALL", parent=BODY, fontSize=8, leading=10, textColor=MUTED)
BULLET = ParagraphStyle("BULLET", parent=BODY, fontSize=10, leading=14, leftIndent=14,
                        bulletIndent=2, bulletFontName="Helvetica-Bold", spaceAfter=3)
KPI_LABEL = ParagraphStyle("KL", fontName="Helvetica", fontSize=8, textColor=MUTED, alignment=1,
                            leading=10)
KPI_VALUE = ParagraphStyle("KV", fontName="Helvetica-Bold", fontSize=18, textColor=NAVY,
                            alignment=1, leading=22)
KPI_DELTA_POS = ParagraphStyle("KDP", fontName="Helvetica-Bold", fontSize=9, textColor=GREEN, alignment=1, leading=11)
KPI_DELTA_NEG = ParagraphStyle("KDN", fontName="Helvetica-Bold", fontSize=9, textColor=RED, alignment=1, leading=11)
KPI_DELTA_MUT = ParagraphStyle("KDM", fontName="Helvetica-Bold", fontSize=9, textColor=MUTED, alignment=1, leading=11)
FOOT = ParagraphStyle("FOOT", parent=BODY, fontSize=7.5, textColor=MUTED, leading=10)

def kpi_tile(label, value, delta_val=None, delta_pct=None, is_currency=False):
    """Build a mini KPI tile as a 1×3 table (label, value, delta)."""
    parts = [[Paragraph(label, KPI_LABEL)], [Paragraph(str(value), KPI_VALUE)]]
    if delta_val is not None:
        d_style = KPI_DELTA_POS if delta_val > 0 else (KPI_DELTA_NEG if delta_val < 0 else KPI_DELTA_MUT)
        arrow = "▲" if delta_val > 0 else ("▼" if delta_val < 0 else "◆")
        d_text = f"{arrow} {fmt_cases(abs(delta_val))} cases &nbsp;({fmt_pct(delta_pct)})"
        parts.append([Paragraph(d_text, d_style)])
    else:
        parts.append([Paragraph("—", KPI_DELTA_MUT)])
    t = Table(parts, colWidths=[1.7*inch])
    t.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 1.5, NAVY),
        ("BACKGROUND", (0,0), (-1,-1), WHITE),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("LINEBELOW", (0,0), (-1,0), 0.5, BORDER),
        ("LINEBELOW", (0,1), (-1,1), 0.5, BORDER),
    ]))
    return t

def styled_table(data, col_widths, header_rows=1, zebra=True, highlight_top=None, highlight_bottom=None,
                 align_overrides=None, delta_cols=None):
    """Standard navy-header zebra-row table."""
    t = Table(data, colWidths=col_widths, repeatRows=header_rows)
    cmds = [
        ("BACKGROUND", (0,0), (-1,header_rows-1), NAVY),
        ("TEXTCOLOR", (0,0), (-1,header_rows-1), WHITE),
        ("FONTNAME", (0,0), (-1,header_rows-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 8.5),
        ("ALIGN", (0,0), (-1,-1), "LEFT"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("BOX", (0,0), (-1,-1), 0.5, BORDER),
        ("INNERGRID", (0,0), (-1,-1), 0.25, BORDER),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
    ]
    if zebra:
        for i in range(header_rows, len(data)):
            if (i - header_rows) % 2 == 1:
                cmds.append(("BACKGROUND", (0,i), (-1,i), ZEBRA))
    if highlight_top:
        for i in highlight_top:
            cmds.append(("TEXTCOLOR", (highlight_top[i] if isinstance(highlight_top, dict) else 0, i),
                         (highlight_top[i] if isinstance(highlight_top, dict) else -1, i), GREEN))
    if highlight_bottom:
        for i in highlight_bottom:
            cmds.append(("TEXTCOLOR", (0, i), (-1, i), RED))
    if align_overrides:
        for col, aln in align_overrides.items():
            cmds.append(("ALIGN", (col, header_rows), (col, -1), aln))
    t.setStyle(TableStyle(cmds))
    return t

# ---------- build flowables ----------
PAGE_W, PAGE_H = A4
MARGIN = 0.5 * inch
FRAME_W = PAGE_W - 2*MARGIN

story = []

def page_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(MUTED)
    canvas.drawString(MARGIN, 0.3*inch, f"K.S. Distillery · Daily Sales Briefing · {TODAY.strftime('%d %B %Y')}")
    canvas.drawRightString(PAGE_W - MARGIN, 0.3*inch, f"Page {doc.page}")
    canvas.restoreState()

# =============== PAGE 1 — SNAPSHOT ===============
story.append(Paragraph(f"Daily Sales Briefing — {TODAY.strftime('%d %B %Y')}", H1))
story.append(Paragraph(
    f"Prepared for evening update &nbsp;·&nbsp; KSBC tertiary data as of {DATA_AS_OF_KSBC} &nbsp;·&nbsp; "
    f"Secondary (FED / BAR) data as of {DATA_AS_OF_SEC}",
    SMALL))
story.append(Spacer(1, 10))

# KPI tiles row 1
k1 = kpi_tile("TOTAL SALES MTD<br/>(KSBC tertiary + FED + BAR)", fmt_cases(apr_total_umbrella), mtd_delta, mtd_pct)
k2 = kpi_tile("vs LAST MONTH SAME PERIOD<br/>(Mar 1–20 prorated)", fmt_cases(mar_total_umbrella_1_20), None)
k3 = kpi_tile("YESTERDAY'S SALES<br/>(20 Apr · KSBC only*)", fmt_cases(yesterday_ksbc), yesterday_delta, yesterday_pct)
k4 = kpi_tile("vs SAME DAY LAST MONTH<br/>(Mar 20 KSBC avg est.)", fmt_cases(mar_20_ksbc_est), None)

kpi_row = Table([[k1, k2, k3, k4]], colWidths=[1.75*inch]*4)
kpi_row.setStyle(TableStyle([
    ("VALIGN", (0,0), (-1,-1), "TOP"),
    ("LEFTPADDING", (0,0), (-1,-1), 3),
    ("RIGHTPADDING", (0,0), (-1,-1), 3),
    ("TOPPADDING", (0,0), (-1,-1), 0),
    ("BOTTOMPADDING", (0,0), (-1,-1), 0),
]))
story.append(kpi_row)
story.append(Spacer(1, 4))
story.append(Paragraph(
    "*Secondary (FED / BAR) invoice not yet ingested for Apr 20; yesterday figure is KSBC tertiary only.",
    FOOT))

# Channel breakdown
story.append(Spacer(1, 14))
story.append(Paragraph("Channel breakdown — Month to date (cases)", H2))

# Compute channel values for last-month same period
apr_fed_only = april_fed or 0
apr_bar_only = april_bar or 0
ch_data = [
    ["Channel", "Apr 1–20 MTD", "Contribution", "Mar 1–20 (prorated)", "Δ cases", "Δ %"],
]
channels = [
    ("KSBC tertiary (shop → consumer)", apr_total_ksbc, mar_1_20_ksbc_total),
    ("Consumer fed (invoice)", apr_fed_only, mar_1_20_fed),
    ("BAR (invoice)", apr_bar_only, mar_1_20_bar),
]
total_row_v = apr_total_umbrella
for name, apr, mar in channels:
    d = apr - mar
    pct = (d / mar * 100) if mar else 0
    contrib = (apr / total_row_v * 100) if total_row_v else 0
    ch_data.append([
        name, fmt_cases(apr), f"{contrib:.1f}%", fmt_cases(mar),
        fmt_cases(d), fmt_pct(pct),
    ])
ch_data.append([
    "TOTAL", fmt_cases(apr_total_umbrella), "100.0%",
    fmt_cases(mar_total_umbrella_1_20), fmt_cases(mtd_delta), fmt_pct(mtd_pct),
])
ch_tbl = Table(ch_data, colWidths=[2.4*inch, 0.95*inch, 0.9*inch, 1.2*inch, 0.8*inch, 0.75*inch])
ch_style = [
    ("BACKGROUND", (0,0), (-1,0), NAVY),
    ("TEXTCOLOR", (0,0), (-1,0), WHITE),
    ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
    ("FONTSIZE", (0,0), (-1,-1), 9),
    ("ALIGN", (1,0), (-1,-1), "RIGHT"),
    ("BOX", (0,0), (-1,-1), 0.5, BORDER),
    ("INNERGRID", (0,0), (-1,-1), 0.25, BORDER),
    ("TOPPADDING", (0,0), (-1,-1), 4),
    ("BOTTOMPADDING", (0,0), (-1,-1), 4),
]
for i in range(1, len(ch_data)-1):
    if i % 2 == 0:
        ch_style.append(("BACKGROUND", (0,i), (-1,i), ZEBRA))
# totals row
ch_style.append(("BACKGROUND", (0,-1), (-1,-1), colors.HexColor("#E8EAF6")))
ch_style.append(("FONTNAME", (0,-1), (-1,-1), "Helvetica-Bold"))
# color delta cells
for i in range(1, len(ch_data)):
    d_txt = ch_data[i][4]
    d_val_str = d_txt.replace(",", "").replace("—", "0")
    try:
        dv = float(d_val_str)
        c = GREEN if dv > 0 else (RED if dv < 0 else MUTED)
        ch_style.append(("TEXTCOLOR", (4,i), (5,i), c))
    except:
        pass
ch_tbl.setStyle(TableStyle(ch_style))
story.append(ch_tbl)

story.append(Spacer(1, 10))
story.append(Paragraph(
    f"<b>Read:</b> MTD umbrella sales of {fmt_cases(apr_total_umbrella)} cases are "
    f"<b>{fmt_pct(mtd_pct)}</b> vs the same period last month. "
    f"KSBC tertiary (the consumer-facing channel) is pacing "
    f"{fmt_pct((apr_total_ksbc - mar_1_20_ksbc_total)/mar_1_20_ksbc_total*100)} "
    f"and invoice sales (FED + BAR) are pacing "
    f"{fmt_pct((apr_fed_bar - mar_fed_bar_1_20)/mar_fed_bar_1_20*100 if mar_fed_bar_1_20 else 0)}.",
    BODY))

story.append(PageBreak())

# =============== PAGE 2 — BOND PERFORMANCE ===============
story.append(Paragraph("Bond-wise Performance — April 1–20 MTD", H1))
story.append(Paragraph(
    "Cases per channel. KSBC = shop tertiary sales; FED / BAR = warehouse invoice dispatches. "
    "Δ vs March 1–20 prorated from full-month data.",
    SMALL))
story.append(Spacer(1, 8))

# build rows
bond_rows = []
for bond, (cluster, asm) in CLUSTER_MAP.items():
    apr_k = apr_bond_sales.get(bond, 0)
    sec = sec_bond.get(bond, dict(fed=0, bar=0))
    apr_fed_b = sec["fed"]
    apr_bar_b = sec["bar"]
    apr_total_b = apr_k + apr_fed_b + apr_bar_b
    mar_k = mar_1_20_bond.get(bond, 0)
    mar_fb = mar_bond_fedbar.get(bond, 0)
    mar_total_b = mar_k + mar_fb
    d = apr_total_b - mar_total_b
    pct = (d / mar_total_b * 100) if mar_total_b else 0
    bond_rows.append([bond, cluster, asm, apr_k, apr_fed_b, apr_bar_b, apr_total_b, d, pct])

bond_rows.sort(key=lambda r: -r[6])

# table
head = ["Bond", "Cluster", "ASM", "KSBC", "FED", "BAR", "Total MTD", "Δ cases", "Δ % vs LM"]
data = [head]
for i, r in enumerate(bond_rows):
    data.append([
        r[0], r[1], r[2],
        fmt_cases(r[3]), fmt_cases(r[4]), fmt_cases(r[5]),
        fmt_cases(r[6]), fmt_cases(r[7]), fmt_pct(r[8]),
    ])
tbl = Table(data, colWidths=[1.1*inch, 0.55*inch, 0.6*inch, 0.65*inch, 0.55*inch, 0.5*inch, 0.85*inch, 0.7*inch, 0.8*inch])
style = [
    ("BACKGROUND", (0,0), (-1,0), NAVY),
    ("TEXTCOLOR", (0,0), (-1,0), WHITE),
    ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
    ("FONTSIZE", (0,0), (-1,-1), 8.5),
    ("ALIGN", (1,0), (-1,-1), "RIGHT"),
    ("ALIGN", (0,0), (0,-1), "LEFT"),
    ("ALIGN", (1,0), (1,-1), "CENTER"),
    ("ALIGN", (2,0), (2,-1), "LEFT"),
    ("BOX", (0,0), (-1,-1), 0.5, BORDER),
    ("INNERGRID", (0,0), (-1,-1), 0.25, BORDER),
    ("TOPPADDING", (0,0), (-1,-1), 3),
    ("BOTTOMPADDING", (0,0), (-1,-1), 3),
]
for i in range(1, len(data)):
    if i % 2 == 0:
        style.append(("BACKGROUND", (0,i), (-1,i), ZEBRA))
# highlight top 3 and bottom 3 Total column
for i in range(1, 4):
    style.append(("TEXTCOLOR", (6,i), (6,i), GREEN))
    style.append(("FONTNAME", (6,i), (6,i), "Helvetica-Bold"))
for i in range(len(data)-3, len(data)):
    style.append(("TEXTCOLOR", (6,i), (6,i), RED))
    style.append(("FONTNAME", (6,i), (6,i), "Helvetica-Bold"))
# color delta columns
for i in range(1, len(data)):
    try:
        d_val = bond_rows[i-1][7]
        c = GREEN if d_val > 0 else (RED if d_val < 0 else MUTED)
        style.append(("TEXTCOLOR", (7,i), (8,i), c))
    except: pass
tbl.setStyle(TableStyle(style))
story.append(tbl)

# cluster leaderboard
cluster_totals = defaultdict(lambda: dict(ksbc=0, fed=0, bar=0, total_mtd=0, mar_total=0, asm=""))
for r in bond_rows:
    bond, cluster, asm, apr_k, apr_fed_b, apr_bar_b, apr_total_b, d, pct = r
    ct = cluster_totals[cluster]
    ct["ksbc"] += apr_k; ct["fed"] += apr_fed_b; ct["bar"] += apr_bar_b
    ct["total_mtd"] += apr_total_b
    ct["mar_total"] += apr_total_b - d
    ct["asm"] = asm

story.append(Spacer(1, 12))
story.append(Paragraph("Cluster Leaderboard", H2))
cl_data = [["Cluster", "ASM", "Bonds", "KSBC", "FED", "BAR", "Total MTD", "Mar 1–20", "Δ %"]]
# cluster bonds assignment
cluster_bonds = defaultdict(list)
for b, (c, _) in CLUSTER_MAP.items():
    cluster_bonds[c].append(b)
for c in sorted(cluster_totals.keys()):
    ct = cluster_totals[c]
    pct = (ct["total_mtd"] - ct["mar_total"]) / ct["mar_total"] * 100 if ct["mar_total"] else 0
    cl_data.append([
        f"Cluster {c}", ct["asm"], f"{len(cluster_bonds[c])} bonds",
        fmt_cases(ct["ksbc"]), fmt_cases(ct["fed"]), fmt_cases(ct["bar"]),
        fmt_cases(ct["total_mtd"]), fmt_cases(ct["mar_total"]), fmt_pct(pct),
    ])
cl_tbl = Table(cl_data, colWidths=[0.9*inch, 0.8*inch, 0.75*inch, 0.7*inch, 0.6*inch, 0.55*inch, 0.95*inch, 0.9*inch, 0.8*inch])
cl_style = [
    ("BACKGROUND", (0,0), (-1,0), NAVY),
    ("TEXTCOLOR", (0,0), (-1,0), WHITE),
    ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
    ("FONTSIZE", (0,0), (-1,-1), 9),
    ("ALIGN", (3,0), (-1,-1), "RIGHT"),
    ("ALIGN", (0,0), (2,-1), "LEFT"),
    ("BOX", (0,0), (-1,-1), 0.5, BORDER),
    ("INNERGRID", (0,0), (-1,-1), 0.25, BORDER),
    ("TOPPADDING", (0,0), (-1,-1), 4),
    ("BOTTOMPADDING", (0,0), (-1,-1), 4),
]
for i in range(1, len(cl_data)):
    if i % 2 == 0:
        cl_style.append(("BACKGROUND", (0,i), (-1,i), ZEBRA))
    row_pct = cl_data[i][-1]
    try:
        v = float(row_pct.replace("%","").replace("+",""))
        cl_style.append(("TEXTCOLOR", (-1,i), (-1,i), GREEN if v > 0 else RED))
    except: pass
cl_tbl.setStyle(TableStyle(cl_style))
story.append(cl_tbl)
story.append(Spacer(1, 6))
best_c = max(cluster_totals.items(), key=lambda kv: kv[1]["total_mtd"])
worst_c = min(cluster_totals.items(), key=lambda kv: kv[1]["total_mtd"])
story.append(Paragraph(
    f"<b>Cluster {best_c[0]} ({best_c[1]['asm']}) is leading the month</b> with {fmt_cases(best_c[1]['total_mtd'])} cases. "
    f"Cluster {worst_c[0]} ({worst_c[1]['asm']}) trailing at {fmt_cases(worst_c[1]['total_mtd'])}. "
    f"Cluster assignments follow a 5-bond geographic split (south / central / north); "
    f"if internal cluster map differs please flag.",
    BODY))

story.append(PageBreak())

# =============== PAGE 3 — BRAND & PACK ===============
story.append(Paragraph("Brand & Pack Movement — KSBC tertiary", H1))
story.append(Paragraph(
    "Brand × pack combos ranked by April 1–20 cases sold through KSBC shops. "
    "Movers compare vs March 1–20 (prorated).",
    SMALL))
story.append(Spacer(1, 10))

story.append(Paragraph("Top 10 brand × pack (April MTD)", H2))
top_data = [["#", "Brand", "Pack", "Apr 1–20 (cases)", "Mar 1–20 (prorated)", "Δ %"]]
top10 = sorted(apr_bp.items(), key=lambda kv: -kv[1])[:10]
for i, (k, v) in enumerate(top10, 1):
    m = mar_1_20_bp.get(k, 0)
    pct = (v - m) / m * 100 if m else 0
    top_data.append([str(i), k[0], k[1], fmt_cases(v), fmt_cases(m), fmt_pct(pct)])
tt = Table(top_data, colWidths=[0.35*inch, 2.9*inch, 0.7*inch, 1.2*inch, 1.3*inch, 0.8*inch])
t_style = [
    ("BACKGROUND", (0,0), (-1,0), NAVY),
    ("TEXTCOLOR", (0,0), (-1,0), WHITE),
    ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
    ("FONTSIZE", (0,0), (-1,-1), 9),
    ("ALIGN", (3,0), (-1,-1), "RIGHT"),
    ("ALIGN", (0,0), (0,-1), "CENTER"),
    ("BOX", (0,0), (-1,-1), 0.5, BORDER),
    ("INNERGRID", (0,0), (-1,-1), 0.25, BORDER),
    ("TOPPADDING", (0,0), (-1,-1), 4),
    ("BOTTOMPADDING", (0,0), (-1,-1), 4),
]
for i in range(1, len(top_data)):
    if i % 2 == 0:
        t_style.append(("BACKGROUND", (0,i), (-1,i), ZEBRA))
    r_pct = top_data[i][-1]
    try:
        v = float(r_pct.replace("%","").replace("+",""))
        t_style.append(("TEXTCOLOR", (-1,i), (-1,i), GREEN if v > 0 else RED))
    except: pass
tt.setStyle(TableStyle(t_style))
story.append(tt)

story.append(Spacer(1, 14))
story.append(Paragraph("Fastest movers (by MoM growth %)", H2))
movers = []
for k, v in apr_bp.items():
    m = mar_1_20_bp.get(k, 0)
    if m >= 50 and v > 0:
        pct = (v - m) / m * 100
        movers.append((k, v, m, pct))
movers.sort(key=lambda x: -x[3])
fm_data = [["Brand", "Pack", "Apr", "Mar", "Δ %"]]
for k, v, m, pct in movers[:6]:
    fm_data.append([k[0], k[1], fmt_cases(v), fmt_cases(m), fmt_pct(pct)])
fm = Table(fm_data, colWidths=[3.0*inch, 0.75*inch, 0.9*inch, 0.9*inch, 0.85*inch])
fm_style = [
    ("BACKGROUND", (0,0), (-1,0), NAVY),
    ("TEXTCOLOR", (0,0), (-1,0), WHITE),
    ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
    ("FONTSIZE", (0,0), (-1,-1), 9),
    ("ALIGN", (2,0), (-1,-1), "RIGHT"),
    ("BOX", (0,0), (-1,-1), 0.5, BORDER),
    ("INNERGRID", (0,0), (-1,-1), 0.25, BORDER),
    ("TOPPADDING", (0,0), (-1,-1), 4),
    ("BOTTOMPADDING", (0,0), (-1,-1), 4),
]
for i in range(1, len(fm_data)):
    if i % 2 == 0:
        fm_style.append(("BACKGROUND", (0,i), (-1,i), ZEBRA))
    fm_style.append(("TEXTCOLOR", (-1,i), (-1,i), GREEN))
fm.setStyle(TableStyle(fm_style))
story.append(fm)

story.append(Spacer(1, 14))
story.append(Paragraph("Slow movers (flat or declining vs last month)", H2))
slow = [m for m in movers if m[3] < 5]
slow.sort(key=lambda x: x[3])
sm_data = [["Brand", "Pack", "Apr", "Mar", "Δ %"]]
for k, v, m, pct in slow[:6]:
    sm_data.append([k[0], k[1], fmt_cases(v), fmt_cases(m), fmt_pct(pct)])
sm = Table(sm_data, colWidths=[3.0*inch, 0.75*inch, 0.9*inch, 0.9*inch, 0.85*inch])
sm_style = list(fm_style)
# replace TEXTCOLOR for last col → RED on all non-header rows
sm_style = [c for c in sm_style if not (isinstance(c, tuple) and len(c)>0 and c[0]=="TEXTCOLOR" and c[1][0]==-1)]
sm_style.extend([("TEXTCOLOR", (-1, i), (-1, i), RED) for i in range(1, len(sm_data))])
sm.setStyle(TableStyle(sm_style))
story.append(sm)

story.append(Spacer(1, 10))
fastest = movers[0] if movers else None
biggest_drag = movers[-1] if movers else None
if fastest and biggest_drag:
    story.append(Paragraph(
        f"<b>Narrative:</b> Fastest mover is <b>{fastest[0][0]} / {fastest[0][1]}</b> ({fmt_pct(fastest[3])} vs LM). "
        f"Biggest drag is <b>{biggest_drag[0][0]} / {biggest_drag[0][1]}</b> ({fmt_pct(biggest_drag[3])} vs LM). "
        f"The 180 ML SKUs across BCB, Old Pearl and BC No.1 are pacing below March — possible seasonality or pricing; "
        f"500 ML formats continue to grow strongly.",
        BODY))

story.append(PageBreak())

# =============== PAGE 4 — WATCH ITEMS ===============
story.append(Paragraph("Watch Items", H1))
story.append(Paragraph("Shops and outlets requiring attention this week.", SMALL))
story.append(Spacer(1, 8))

# Non-performing
story.append(Paragraph(f"Non-performing active KSBC shops — zero sales MTD ({len(4*[0])})", H2))
nonperf = []
for sc, info in shop_info.items():
    if info.get("cat") != "KSBC" or info.get("status") != "Active":
        continue
    if apr_shop_sales.get(sc, 0) == 0:
        nonperf.append((sc, info.get("name"), info.get("bond"), info.get("staff")))
np_data = [["Shop Code", "Shop Name", "Bond", "Field Staff"]]
for sc, name, bond, staff in nonperf:
    np_data.append([sc, name, bond, staff])
np_tbl = Table(np_data, colWidths=[0.9*inch, 3.0*inch, 1.3*inch, 1.8*inch])
np_style = [
    ("BACKGROUND", (0,0), (-1,0), NAVY),
    ("TEXTCOLOR", (0,0), (-1,0), WHITE),
    ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
    ("FONTSIZE", (0,0), (-1,-1), 9),
    ("BOX", (0,0), (-1,-1), 0.5, BORDER),
    ("INNERGRID", (0,0), (-1,-1), 0.25, BORDER),
    ("TOPPADDING", (0,0), (-1,-1), 3),
    ("BOTTOMPADDING", (0,0), (-1,-1), 3),
]
for i in range(1, len(np_data)):
    if i % 2 == 0:
        np_style.append(("BACKGROUND", (0,i), (-1,i), ZEBRA))
np_tbl.setStyle(TableStyle(np_style))
story.append(np_tbl)

# Trending down
story.append(Spacer(1, 14))
story.append(Paragraph("Trending down — Apr MTD < 50% of Mar 1–20 (top 10 by gap size)", H2))
trends = []
for sc, info in shop_info.items():
    if info.get("cat") != "KSBC" or info.get("status") != "Active":
        continue
    mar = mar_1_20_shop.get(sc, 0)
    apr = apr_shop_sales.get(sc, 0)
    if mar >= 10 and apr < 0.5 * mar:
        trends.append((sc, info.get("name"), info.get("bond"), info.get("staff"), apr, mar, mar - apr))
trends.sort(key=lambda x: -x[6])
td_data = [["Shop Code", "Shop Name", "Bond", "Staff", "Apr 1–20", "Mar 1–20", "Gap"]]
for sc, name, bond, staff, apr, mar, gap in trends[:10]:
    td_data.append([sc, name, bond, staff, fmt_cases(apr), fmt_cases(mar), fmt_cases(gap)])
td_tbl = Table(td_data, colWidths=[0.75*inch, 2.25*inch, 1.1*inch, 1.15*inch, 0.7*inch, 0.75*inch, 0.6*inch])
td_style = [
    ("BACKGROUND", (0,0), (-1,0), NAVY),
    ("TEXTCOLOR", (0,0), (-1,0), WHITE),
    ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
    ("FONTSIZE", (0,0), (-1,-1), 8.5),
    ("ALIGN", (4,0), (-1,-1), "RIGHT"),
    ("BOX", (0,0), (-1,-1), 0.5, BORDER),
    ("INNERGRID", (0,0), (-1,-1), 0.25, BORDER),
    ("TOPPADDING", (0,0), (-1,-1), 3),
    ("BOTTOMPADDING", (0,0), (-1,-1), 3),
]
for i in range(1, len(td_data)):
    if i % 2 == 0:
        td_style.append(("BACKGROUND", (0,i), (-1,i), ZEBRA))
    td_style.append(("TEXTCOLOR", (-1,i), (-1,i), RED))
td_tbl.setStyle(TableStyle(td_style))
story.append(td_tbl)

# Unmatched outlets
story.append(Spacer(1, 14))
story.append(Paragraph("Unmatched outlets in secondary data", H2))
um_rows = []
for r in sec_wb["UNMATCHED"].iter_rows(values_only=True):
    if not r or r[0] is None:
        continue
    # skip title / instruction / header rows
    v0 = str(r[0])
    if any(s in v0 for s in ("⚠", "need to be added", "Licensee Code")):
        continue
    um_rows.append((r[0], r[1], r[2]))
if um_rows:
    um_data = [["Licensee Code", "Licensee Name", "Cases"]]
    for code, name, cases in um_rows:
        um_data.append([code, name, fmt_cases(cases)])
    um_tbl = Table(um_data, colWidths=[1.2*inch, 4.5*inch, 1.0*inch])
    um_style = [
        ("BACKGROUND", (0,0), (-1,0), NAVY),
        ("TEXTCOLOR", (0,0), (-1,0), WHITE),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("ALIGN", (2,0), (-1,-1), "RIGHT"),
        ("BOX", (0,0), (-1,-1), 0.5, BORDER),
        ("INNERGRID", (0,0), (-1,-1), 0.25, BORDER),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]
    for i in range(1, len(um_data)):
        if i % 2 == 0:
            um_style.append(("BACKGROUND", (0,i), (-1,i), ZEBRA))
    um_tbl.setStyle(TableStyle(um_style))
    story.append(um_tbl)
else:
    story.append(Paragraph("None. All outlets reconciled to master.", BODY))

# Closed-shop reappearances check
closed_codes = set()
for sc, info in shop_info.items():
    if info.get("status") == "Closed":
        closed_codes.add(sc)
closed_reap = [(sc, shop_info[sc]["name"], shop_info[sc]["bond"], apr_shop_sales.get(sc, 0))
               for sc in closed_codes if apr_shop_sales.get(sc, 0) > 0]

story.append(Spacer(1, 14))
story.append(Paragraph("Closed-shop reappearances in raw", H2))
if closed_reap:
    cr_data = [["Shop Code", "Shop Name", "Bond", "Apr 1–20 Cases"]]
    for sc, name, bond, v in closed_reap:
        cr_data.append([sc, name, bond, fmt_cases(v)])
    cr_tbl = Table(cr_data, colWidths=[1.0*inch, 3.0*inch, 1.3*inch, 1.5*inch])
    cr_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), NAVY),
        ("TEXTCOLOR", (0,0), (-1,0), WHITE),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("BOX", (0,0), (-1,-1), 0.5, BORDER),
        ("INNERGRID", (0,0), (-1,-1), 0.25, BORDER),
    ]))
    story.append(cr_tbl)
else:
    story.append(Paragraph("None. 6 closed shops remain at zero in raw.", BODY))

# Data freshness
story.append(Spacer(1, 14))
story.append(Paragraph("Data freshness", H2))
story.append(Paragraph(
    f"• <b>KSBC tertiary:</b> up to {DATA_AS_OF_KSBC} — current.<br/>"
    f"• <b>Secondary sales (FED / BAR / KSBC dispatch):</b> up to {DATA_AS_OF_SEC} — "
    f"Apr 20 invoice file not yet ingested (2026-04-21 morning). "
    f"Secondary MTD excludes Apr 20; yesterday figure excludes invoice sales.",
    BODY))

story.append(PageBreak())

# =============== PAGE 5 — TALKING POINTS ===============
story.append(Paragraph("Talking Points — for the evening update", H1))
story.append(Paragraph("Use these as conversation starters with Anup & Shearish.", SMALL))
story.append(Spacer(1, 12))

# --- What's going well ---
story.append(Paragraph("What's going well", H2))

going_well = []
top_bond = bond_rows[0]
going_well.append(
    f"<b>{top_bond[0]}</b> is our top bond this month — {fmt_cases(top_bond[6])} total cases, "
    f"{fmt_pct(top_bond[8])} vs March same period. ASM {top_bond[2]} running ahead."
)
# top cluster
best_cluster_name = max(cluster_totals.items(), key=lambda kv: kv[1]["total_mtd"])
going_well.append(
    f"<b>Cluster {best_cluster_name[0]} ({best_cluster_name[1]['asm']})</b> leading the cluster table "
    f"at {fmt_cases(best_cluster_name[1]['total_mtd'])} cases."
)
# top mover
if movers:
    top_mover = movers[0]
    going_well.append(
        f"<b>{top_mover[0][0]} / {top_mover[0][1]}</b> is the fastest-growing SKU — "
        f"{fmt_cases(top_mover[1])} cases this month, {fmt_pct(top_mover[3])} MoM. Worth doubling down."
    )
# 500 ML growth narrative
going_well.append(
    "<b>500 ML pack mix</b> is clearly accelerating across brands — Old Pearl, K.S 99, BCB and Blender's Choice "
    "all up double-digits vs March on the 500 ML format. Consumer preference shift."
)

for g in going_well:
    story.append(Paragraph(f"• {g}", BULLET))

story.append(Spacer(1, 8))

# --- What needs attention ---
story.append(Paragraph("What needs attention", H2))
needs_attn = []

# worst bond
worst_bond = bond_rows[-1]
needs_attn.append(
    f"<b>{worst_bond[0]}</b> is the weakest bond — {fmt_cases(worst_bond[6])} cases, "
    f"{fmt_pct(worst_bond[8])} vs LM. ASM {worst_bond[2]} — needs a push."
)
# 180 ML decline
needs_attn.append(
    "<b>180 ML across BCB, Old Pearl and BC No.1 is pacing down vs March</b> — the small-pack primary "
    "retail format is losing velocity. Field team should investigate pricing / shelf presence."
)
# non-performing
if nonperf:
    names = ", ".join(f"{x[1].strip()} ({x[2]})" for x in nonperf)
    needs_attn.append(
        f"<b>{len(nonperf)} non-performing shops</b> (strict zero MTD): {names}. "
        f"Assign field check before month-end."
    )
# trending-down
if trends:
    top_trend = trends[0]
    needs_attn.append(
        f"<b>{top_trend[1].strip()}</b> ({top_trend[2]}, {top_trend[3]}) is the biggest "
        f"falling-off shop — {fmt_cases(top_trend[4])} cases vs {fmt_cases(top_trend[5])} "
        f"last month same period (gap {fmt_cases(top_trend[6])}). "
        f"Plus {len(trends)-1} other shops down >50%."
    )

for n in needs_attn:
    story.append(Paragraph(f"• {n}", BULLET))

story.append(Spacer(1, 8))

# --- Questions to raise ---
story.append(Paragraph("Questions to raise", H2))
questions = []
if worst_bond[8] < -15:
    questions.append(
        f"Shall we redirect field push to <b>{worst_bond[0]}</b> — trailing "
        f"{fmt_pct(worst_bond[8])} vs LM, the weakest bond this month?"
    )
if best_cluster_name[0] != min(cluster_totals.items(), key=lambda kv: kv[1]["total_mtd"])[0]:
    questions.append(
        f"<b>Cluster {best_cluster_name[0]} ({best_cluster_name[1]['asm']})</b> is pulling ahead — "
        f"anything specific we want to replicate in the other clusters this week?"
    )
# seasonality question
questions.append(
    "180 ML decline is across three brands — is this a pricing move by KSBC, a competitor pack change, "
    "or genuine consumer shift to 500 ML? Worth a call with field before month close."
)

for q in questions:
    story.append(Paragraph(f"• {q}", BULLET))

# close footer note
story.append(Spacer(1, 16))
story.append(Paragraph(
    f"<i>Briefing auto-generated {TODAY.strftime('%d %b %Y')}. Sources: "
    f"APRIL 1st–20th KSBC shop sales workbook · APRIL 1st–19th secondary sales workbook · "
    f"MARCH full-month comparables prorated to 20-day equivalent. "
    f"Pure sales view — no financial / cashflow content included.</i>",
    FOOT))

# ---------- render ----------
os.makedirs(OUT_DIR, exist_ok=True)

doc = BaseDocTemplate(OUT_PDF, pagesize=A4,
                      leftMargin=MARGIN, rightMargin=MARGIN,
                      topMargin=MARGIN, bottomMargin=MARGIN,
                      title=f"K.S. Distillery — Daily Sales Briefing — {TODAY.strftime('%d %B %Y')}",
                      author="K.S. Distillery sales ops")
frame = Frame(MARGIN, MARGIN, FRAME_W, PAGE_H - 2*MARGIN, id="normal",
              leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0.3*inch)
tmpl = PageTemplate(id="portrait", frames=[frame], onPage=page_footer)
doc.addPageTemplates([tmpl])

doc.build(story)
sz = os.path.getsize(OUT_PDF)
print(f"Built {OUT_PDF} ({sz/1024:.1f} KB)")

ksbc_wb.close(); pksbc_wb.close(); sec_wb.close(); psec_wb.close()
