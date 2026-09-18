"""
Daily pace — DEMO of the revised chart. Not wired into the live build.

Two changes under discussion (27 Jul 2026):

  1. NO per-day green/red verdict. 72% of green days in July were green only
     because a FED/BAR invoice landed that day, and 83% of red days had no
     invoice at all -- so the daily colour was mostly measuring dispatch
     timing, not selling. The bars now just report what sold.

  2. The judgement moves to the CUMULATIVE position, where lumpiness washes
     out: a required line against an actual line, with the gap shaded. Bars
     are stacked by channel so the executive can see for himself which of his
     days were shop flow and which were a truck.

Chart form matches the liquidation-live artifact already in use: channel
stacked bars + cumulative pace lines on a second axis.
"""

from __future__ import annotations

import datetime as _dt
import os
import sys
from collections import defaultdict

from reportlab.lib import colors

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import pace_data as PD
import pace_engine as PE
import pace_shops as PS
import pace_pdf as PP
from pace_pdf import (M, W, H, NAVY, NAVY_MID, NAVY_SOFT, GOLD, INK, MUTED,
                      LINE, ZEBRA, WHITE, GREEN, GREEN_BG, RED, RED_BG,
                      AMBER, AMBER_BG, _fmt)

SHOP_C = colors.HexColor("#93B0DE")      # shop liquidation — the daily flow
SHOP_E = colors.HexColor("#4A6FA5")
INV_C = colors.HexColor("#F0B323")       # FED/BAR invoice — the lumpy events
REQ_C = colors.HexColor("#8A93A6")
GAP_C = colors.HexColor("#FBE3E3")


def channel_split(hist_rows, plan) -> dict:
    """bond -> {day: (tertiary, invoice)}"""
    out = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0]))
    for r in hist_rows:
        d = _dt.date.fromisoformat(str(r["date"])[:10])
        if d.year != plan["year"] or d.month != plan["mno"]:
            continue
        v = out[r["bond"]][d.day]
        v[0] += float(r["tertiary_cs"])
        v[1] += float(r["invoice_cs"])
    return out


CLIMB_C = colors.HexColor("#E65100")     # the climb now required
PROJ_C = colors.HexColor("#C62828")      # where this rate actually lands
WEDGE_C = colors.HexColor("#FFF4E0")


def chart_v2(c, x, y, w, h, b, split, plan):
    """
    Stacked daily bars + the fork in the road.

    Up to today: where he should have been (grey dashed) against where he is
    (red), gap shaded and labelled in DAYS behind — people feel time more
    sharply than case volume.

    From today: two lines fan out to month end. The climb still required
    (orange, ends on target) and where his current rate actually lands (red
    dotted). The wedge between them is the month he is about to lose, and the
    labelled distance at the right edge is the shortfall.
    """
    rows = b["rows"]
    as_of = plan["as_of"].day
    ndays = plan["ndays"]
    per_day = b["per_day"]
    target = b["target_month"]
    projected = b["mtd_actual"] + b["run_rate"] * b["days_left"]
    short = target - projected

    c.setFillColor(colors.HexColor("#FBFCFE"))
    c.roundRect(x, y - h, w, h, 4, stroke=0, fill=1)
    c.setFillColor(NAVY_MID)
    c.setFont("Helvetica-Bold", 8.4)
    c.drawString(x + 12, y - 14, f"DAY BY DAY — {plan['month'].title()}")

    lx = x + 128
    for lab, col in (("shop sales", SHOP_C), ("invoice", INV_C), ("sold", RED),
                     ("still required", CLIMB_C), ("on this rate", PROJ_C)):
        c.setFillColor(col)
        c.rect(lx, y - 17.5, 6, 6, stroke=0, fill=1)
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 6.4)
        c.drawString(lx + 9, y - 16.5, lab)
        lx += 9 + c.stringWidth(lab, "Helvetica", 6.4) + 16

    pad_l, pad_r, pad_b, pad_t = 30, 62, 21, 34
    px, py = x + pad_l, y - h + pad_b
    pw, ph = w - pad_l - pad_r, h - pad_b - pad_t
    n = len(rows)
    bw = pw / max(n, 1)

    tot = sorted(v for v in ((r["actual"] or 0) for r in rows
                             if not r["dry"] and r["day"] <= as_of) if v > 0.001)
    ref = max(tot[min(int(len(tot) * 0.85), len(tot) - 1)] if tot else 0.0,
              per_day, 0.001)
    cap = ref * 1.45
    dp = 0 if ref >= 10 else 1

    cmax = max(target, projected, b["mtd_actual"]) * 1.02 or 1.0

    def dx(d):
        return px + (d - 0.5) * bw

    def cy(v):
        return py + ph * min(max(v, 0) / cmax, 1.0)

    c.setStrokeColor(colors.HexColor("#E6EAF2"))
    c.setLineWidth(0.4)
    for f in (0.5, 1.0):
        c.line(px, py + ph * f, px + pw, py + ph * f)

    # --- the wedge: target versus where this rate lands ------------------
    tx, ty0 = dx(as_of), cy(b["mtd_actual"])
    ex = dx(ndays)
    p = c.beginPath()
    p.moveTo(tx, ty0)
    p.lineTo(ex, cy(target))
    p.lineTo(ex, cy(projected))
    p.close()
    c.setFillColor(WEDGE_C if short > 0.5 else colors.HexColor("#E8F5E9"))
    c.drawPath(p, stroke=0, fill=1)

    # --- shaded gap to date ----------------------------------------------
    pts_req = [(dx(r["day"]), cy(r["target_cum"])) for r in rows
               if r["day"] <= as_of]
    pts_act = [(dx(r["day"]), cy(r["actual_cum"] or 0)) for r in rows
               if r["day"] <= as_of]
    if len(pts_act) > 1:
        p = c.beginPath()
        p.moveTo(*pts_act[0])
        for pt in pts_act[1:]:
            p.lineTo(*pt)
        for pt in reversed(pts_req):
            p.lineTo(*pt)
        p.close()
        c.setFillColor(GAP_C if b['gap'] < 0 else colors.HexColor('#E4F1E4'))
        c.drawPath(p, stroke=0, fill=1)

    # --- stacked bars ----------------------------------------------------
    for i, r in enumerate(rows):
        bx = px + i * bw
        if r["dry"]:
            c.setFillColor(colors.HexColor("#EDEFF4"))
            c.rect(bx + bw * .18, py, bw * .64, ph * .04, stroke=0, fill=1)
            continue
        if r["actual"] is None:
            continue
        ter, inv = split[b["bond"]].get(r["day"], [0.0, 0.0])
        sc = (r["actual"] / (ter + inv)) if (ter + inv) > 0.001 else 1.0
        ter, inv = ter * sc, inv * sc
        h_t = ph * (min(ter, cap) / cap)
        h_i = ph * (min(ter + inv, cap) / cap) - h_t
        c.setFillColor(SHOP_C)
        c.rect(bx + bw * .18, py, bw * .64, max(h_t, 0.5), stroke=0, fill=1)
        if inv > 0.005:
            c.setFillColor(INV_C)
            c.rect(bx + bw * .18, py + h_t, bw * .64, max(h_i, 0.8),
                   stroke=0, fill=1)
        c.setFillColor(SHOP_E)
        c.setFont("Helvetica-Bold", 5.6)
        lbl = f"{r['actual']:,.{dp}f}"
        if r["actual"] > cap:
            lbl = "▲ " + lbl
        c.drawCentredString(bx + bw * .5, py + h_t + max(h_i, 0) + 3.0, lbl)

    # --- lines -----------------------------------------------------------
    c.setStrokeColor(REQ_C)
    c.setLineWidth(1.2)
    c.setDash(4, 3)
    c.lines([(pts_req[i][0], pts_req[i][1], pts_req[i + 1][0], pts_req[i + 1][1])
             for i in range(len(pts_req) - 1)])

    c.setStrokeColor(PROJ_C)                       # projection
    c.setLineWidth(1.6)
    c.setDash(2.5, 2.5)
    c.line(tx, ty0, ex, cy(projected))
    c.setDash()

    c.setStrokeColor(CLIMB_C)                      # the climb still required
    c.setLineWidth(2.0)
    c.line(tx, ty0, ex, cy(target))

    if len(pts_act) > 1:                           # actual
        c.setStrokeColor(RED)
        c.setLineWidth(2.2)
        c.lines([(pts_act[i][0], pts_act[i][1], pts_act[i + 1][0],
                  pts_act[i + 1][1]) for i in range(len(pts_act) - 1)])
    c.setFillColor(RED)
    c.circle(tx, ty0, 2.6, stroke=0, fill=1)

    # --- the shortfall, drawn as a distance ------------------------------
    yt, yp = cy(target), cy(projected)
    acc = PROJ_C if short > 0.5 else GREEN
    word = "SHORT" if short > 0.5 else "CLEAR"
    if abs(yt - yp) >= 26:
        c.setStrokeColor(acc)
        c.setLineWidth(0.9)
        c.line(ex + 6, yt, ex + 6, yp)
        for yy in (yt, yp):
            c.line(ex + 3.5, yy, ex + 8.5, yy)
        c.setFillColor(acc)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(ex + 11, (yt + yp) / 2 - 5, f"{abs(short):,.0f}")
        c.setFont("Helvetica-Bold", 6)
        c.drawString(ex + 11, (yt + yp) / 2 - 12.5, word)
        c.setFillColor(CLIMB_C)
        c.setFont("Helvetica-Bold", 6.4)
        c.drawString(ex + 4, yt + 6, f"{target:,.0f}")
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 5.6)
        c.drawString(ex + 4, yt + 12.5, "target")
        c.setFillColor(PROJ_C)
        c.setFont("Helvetica-Bold", 6.4)
        c.drawString(ex + 4, yp - 11, f"{projected:,.0f}")
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 5.6)
        c.drawString(ex + 4, yp - 17.5, "on this rate")
    else:
        # target and projection nearly coincide -- one stacked block instead
        mid = (yt + yp) / 2
        c.setFillColor(CLIMB_C)
        c.setFont("Helvetica-Bold", 6.4)
        c.drawString(ex + 6, mid + 9, f"{target:,.0f}")
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 5.6)
        c.drawString(ex + 6 + c.stringWidth(f"{target:,.0f}",
                                            "Helvetica-Bold", 6.4) + 3,
                     mid + 9, "target")
        c.setFillColor(PROJ_C)
        c.setFont("Helvetica-Bold", 6.4)
        c.drawString(ex + 6, mid + 1, f"{projected:,.0f}")
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 5.6)
        c.drawString(ex + 6 + c.stringWidth(f"{projected:,.0f}",
                                            "Helvetica-Bold", 6.4) + 3,
                     mid + 1, "on this rate")
        c.setFillColor(acc)
        c.setFont("Helvetica-Bold", 7.4)
        c.drawString(ex + 6, mid - 9.5, f"{abs(short):,.0f} {word}")

    # days behind, written into the gap band itself
    if b["gap"] < -0.5 and per_day > 0:
        dbehind = abs(b["gap"]) / per_day
        mid = int(len(pts_act) * 0.62)
        cxm = (pts_act[mid][0] + pts_req[mid][0]) / 2
        cym = (pts_act[mid][1] + pts_req[mid][1]) / 2
        txt = f"{dbehind:,.1f} DAYS BEHIND"
        tw = c.stringWidth(txt, "Helvetica-Bold", 7)
        c.setFillColor(WHITE)                       # pill, so it reads over bars
        c.roundRect(cxm - tw / 2 - 5, cym - 6, tw + 10, 12, 3,
                    stroke=0, fill=1)
        c.setStrokeColor(colors.HexColor("#E8B4B4"))
        c.setLineWidth(0.5)
        c.roundRect(cxm - tw / 2 - 5, cym - 6, tw + 10, 12, 3,
                    stroke=1, fill=0)
        c.setFillColor(colors.HexColor("#B0464A"))
        c.setFont("Helvetica-Bold", 7)
        c.drawCentredString(cxm, cym - 2.5, txt)

    c.setFillColor(MUTED)
    c.setFont("Helvetica", 6.2)
    for i, r in enumerate(rows):
        if r["day"] % 5 == 0 or r["day"] == 1:
            c.drawCentredString(px + i * bw + bw / 2, py - 10, str(r["day"]))
    c.setFont("Helvetica", 6.4)
    c.drawString(px, py - 20,
                 "bars = cases sold that day, split by channel   ·   "
                 "pink band = how far behind you already are   ·   "
                 "from today the orange line is the climb still required, "
                 "the dotted line is where this rate lands")
    return y - h


def bond_page_v2(pdf, plan, bond, shops, split, staff_name):
    b = plan["bonds"][bond]
    p = plan
    behind = b["gap"]
    c = pdf.c
    who = staff_name.title() if staff_name else "VACANT — NO EXECUTIVE"
    y = pdf.header(bond, f"{b['pct_done']:.0%}",
                   f"{who}   ·   Cluster {b['cluster']}   ·   "
                   f"sales through {p['as_of']:%d %B}")
    y -= 8

    gap_w = 12
    pw = (W - 2 * M - 2 * gap_w) / 3
    ph = 50
    panels = [
        ("THE BASE", NAVY_SOFT, colors.HexColor("#F7F9FC"), [
            ("Monthly target", _fmt(b["target_month"]), INK, False),
            ("Selling days", str(p["selling_days"]), INK, False),
            ("Target / day", _fmt(b["per_day"], 1), NAVY_MID, True)]),
        ("WHERE YOU STAND", GREEN if behind >= 0 else AMBER,
         GREEN_BG if behind >= 0 else AMBER_BG, [
            ("Sold so far", _fmt(b["mtd_actual"]), INK, False),
            ("Should be at", _fmt(b["mtd_expected"]), INK, False),
            ("Running at", _fmt(b["run_rate"], 1), INK, True)]),
        ("THE PRESSURE", RED if behind < 0 else GREEN,
         RED_BG if behind < 0 else GREEN_BG, [
            ("Behind by" if behind < 0 else "Ahead by", _fmt(abs(behind)),
             RED if behind < 0 else GREEN, False),
            ("Days left", str(b["days_left"]), INK, False),
            ("Needed / day", _fmt(b["needed_per_day"], 1),
             RED if behind < 0 else GREEN, True)]),
    ]
    for i, (title, accent, bg, stats) in enumerate(panels):
        x = M + i * (pw + gap_w)
        c.setFillColor(bg)
        c.roundRect(x, y - ph, pw, ph, 4, stroke=0, fill=1)
        c.setFillColor(accent)
        c.roundRect(x, y - ph, 3.4, ph, 1.7, stroke=0, fill=1)
        c.setFillColor(accent)
        c.setFont("Helvetica-Bold", 7.4)
        c.drawString(x + 13, y - 12, title)
        sw = (pw - 22) / 3
        for j, (lab, val, col, big) in enumerate(stats):
            sx = x + 13 + j * sw
            c.setFillColor(MUTED)
            c.setFont("Helvetica", 6.4)
            c.drawString(sx, y - 25, lab.upper())
            c.setFillColor(col)
            c.setFont("Helvetica-Bold", 15 if big else 11.5)
            c.drawString(sx, y - 42, val + (" cs" if big else ""))
    y -= ph + 10

    chart_v2(c, M, y, W - 2 * M, 190, b, split, plan)
    y -= 190 + 12

    rows_s = shops.get(bond, [])
    c.setFillColor(NAVY_MID)
    c.setFont("Helvetica-Bold", 8.4)
    c.drawString(M, y - 10, "AVERAGE SALES PER SHOP, PER DAY")
    c.setFillColor(MUTED)
    c.setFont("Helvetica-Oblique", 6.8)
    elapsed = sum(1 for d in range(1, p["as_of"].day + 1)
                  if d not in p["dry_days"])
    c.drawString(M, y - 19, f"cases per selling day · this month "
                            f"({p['month'].title()}, {elapsed} days so far) "
                            f"against last month · KSBC shops")
    if rows_s:
        c.setFillColor(NAVY_MID)
        c.setFont("Helvetica-Bold", 7.4)
        c.drawRightString(W - M, y - 10,
                          f"KSBC SHOPS: {sum(r['now'] for r in rows_s):,.1f} "
                          f"cs/day   (last month "
                          f"{sum(r['last'] for r in rows_s):,.1f})")
    y -= 26
    ncol = 3
    third = (W - 2 * M - 12 * (ncol - 1)) / ncol
    cols = [("#", 20, "c"), ("SHOP", third - 134, "l"),
            ("NOW/DAY", 44, "c"), ("LAST", 36, "c"), ("+/−", 34, "c")]
    per_col = -(-len(rows_s) // ncol) if rows_s else 0
    for side in range(ncol):
        chunk = rows_s[side * per_col:(side + 1) * per_col]
        if chunk:
            PP._shop_daily_table(c, M + side * (third + 12), y, third, cols,
                                 chunk, side * per_col)
    pdf.footer(f"{bond} · {who if staff_name else 'Vacant'} · daily pace "
               f"through {plan['as_of']:%d %b %Y}")
    c.showPage()


if __name__ == "__main__":
    base = PD._default_base()
    hist = PE._load_history(base)
    plan = PE.build_plan(base, hist_rows=hist)
    plan["base"] = base
    master = PD.load_master(base)
    shops = PS.build_shop_daily(base, plan, hist, master)
    split = channel_split(hist, plan)
    staff = {}
    for info in master.values():
        s = (info["staff"] or "").strip()
        if info["status"].upper() != "CLOSED" and info["bond"] and s \
                and s.upper() != "VACANT":
            staff.setdefault(info["bond"], s)

    out = sys.argv[1] if len(sys.argv) > 1 else "/tmp/demo.pdf"
    # four bonds chosen to show the range: almost no invoice, very lumpy,
    # invoice-dominated, and one in the middle
    picks = ["KOTTARAKARA", "KOZHIKODE", "ALUVA", "KOLLAM"]
    pdf = PP.PacePDF(out, base)
    for bnd in picks:
        bond_page_v2(pdf, plan, bnd, shops, split, staff.get(bnd))
    pdf.c.save()
    print("wrote", out)
