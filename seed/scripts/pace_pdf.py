"""
Daily pace — per-cluster PDFs.

One PDF per cluster, handed to the ASM. Design language matches the existing
`TARGET <MONTH> 2026 - CLUSTER <N>.pdf` files: A4 landscape, navy header band
with a gold K.S. DISTILLERY eyebrow, gold rule, navy table headers, zebra
rows, gold accent on the column that matters.

Structure:
    page 1      cluster summary  — KPI strip + bond league table, ranked
    pages 2..N  one page per bond — THE BASE / WHERE YOU STAND / THE PRESSURE,
                a day-by-day chart of target vs actual, and the shops to visit

The three panels are the argument: here is the number you agreed to, here is
where you actually are, and here is what that shortfall has done to today's
number. The chart makes the drift visible; the shop table makes it fixable.
"""

from __future__ import annotations

import datetime as _dt
import os

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as _canvas

# --- palette (matches the existing cluster target PDFs) --------------------
NAVY = colors.HexColor("#0D1B4A")
NAVY_MID = colors.HexColor("#1A237E")
NAVY_SOFT = colors.HexColor("#263F80")
GOLD = colors.HexColor("#F0B323")
GOLD_SOFT = colors.HexColor("#FDF3DC")
INK = colors.HexColor("#1F2937")
MUTED = colors.HexColor("#6B7280")
LINE = colors.HexColor("#CBD5E1")
ZEBRA = colors.HexColor("#F4F6FA")
WHITE = colors.white
GREEN = colors.HexColor("#2E7D32")
GREEN_BG = colors.HexColor("#E6F4E6")
RED = colors.HexColor("#C62828")
RED_BG = colors.HexColor("#FDECEC")
RED_DEEP = colors.HexColor("#8E1B1B")
SHOP_C = colors.HexColor("#93B0DE")      # shop liquidation — the daily flow
SHOP_E = colors.HexColor("#4A6FA5")
INV_C = colors.HexColor("#F0B323")       # FED/BAR invoice — the lumpy events
REQ_C = colors.HexColor("#8A93A6")
GAP_R = colors.HexColor("#FBE3E3")
GAP_G = colors.HexColor("#E4F1E4")
GREEN_DEEP = colors.HexColor("#1B5E20")
AMBER = colors.HexColor("#E65100")
AMBER_BG = colors.HexColor("#FFF1E0")

W, H = landscape(A4)          # 841.89 x 595.28
M = 40                        # side margin
TOP_M = 22                    # top margin — tightened 40->22 (Abhay,
                              # 27 Jul 2026, "push this part more upside")
HDR_H = 64


def _fmt(v, dp=0):
    if v is None:
        return "·"
    if abs(v) < 0.05 and dp == 0:
        return "·"
    return f"{v:,.{dp}f}"


def _ord(n: int) -> str:
    if 10 <= n % 100 <= 20:
        sfx = "th"
    else:
        sfx = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{sfx}"


def _logo_path(base):
    p = os.path.join(base, "Internal Docs", "KSD LOGO.png")
    return p if os.path.exists(p) else None


class PacePDF:
    def __init__(self, path, base):
        self.c = _canvas.Canvas(path, pagesize=landscape(A4))
        self.c.setTitle(os.path.basename(path).replace(".pdf", ""))
        self.base = base
        self.logo = _logo_path(base)

    # -- chrome ------------------------------------------------------------
    def header(self, title, right, sub=None):
        c = self.c
        top = H - TOP_M
        c.setFillColor(NAVY)
        c.rect(M, top - HDR_H, W - 2 * M, HDR_H, stroke=0, fill=1)
        c.setFillColor(GOLD)
        c.rect(M, top - HDR_H - 3, W - 2 * M, 3, stroke=0, fill=1)

        c.setFillColor(GOLD)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(M + 18, top - 19, "K.S. DISTILLERY")
        c.setFillColor(WHITE)
        c.setFont("Helvetica-Bold", 19)
        c.drawString(M + 18, top - 42, title)
        if sub:
            c.setFillColor(colors.HexColor("#B9C4E8"))
            c.setFont("Helvetica", 9)
            c.drawString(M + 18, top - 55, sub)

        c.setFillColor(GOLD)
        c.setFont("Helvetica-Bold", 21)
        c.drawRightString(W - M - 18, top - 41, right)
        return top - HDR_H - 3

    def footer(self, note):
        """Removed 27 Jul 2026 (Abhay) — kept as a no-op so call sites stay put."""
        return

    # -- primitives --------------------------------------------------------
    def kpi_strip(self, y, tiles, h=58):
        """tiles = [(label, value, sublabel, accent_color, bg)]"""
        c = self.c
        n = len(tiles)
        gap = 10
        w = (W - 2 * M - gap * (n - 1)) / n
        for i, (lab, val, sub, accent, bg) in enumerate(tiles):
            x = M + i * (w + gap)
            c.setFillColor(bg or colors.HexColor("#F7F9FC"))
            c.roundRect(x, y - h, w, h, 4, stroke=0, fill=1)
            c.setFillColor(accent)
            c.roundRect(x, y - h, 3.2, h, 1.6, stroke=0, fill=1)
            c.setFillColor(MUTED)
            c.setFont("Helvetica-Bold", 7.2)
            c.drawString(x + 12, y - 16, lab.upper())
            c.setFillColor(accent if accent != NAVY_SOFT else INK)
            c.setFont("Helvetica-Bold", 20)
            c.drawString(x + 11, y - 39, val)
            if sub:
                c.setFillColor(MUTED)
                c.setFont("Helvetica", 7.4)
                c.drawString(x + 12, y - 51, sub)
        return y - h

    def table(self, y, cols, rows, row_h=21, hdr_h=28, accent_col=None,
              bold_col=0):
        """cols = [(header, width, align)] ; rows = list of (cells, meta)"""
        c = self.c
        total_w = sum(w for _, w, _ in cols)
        x0 = M + (W - 2 * M - total_w) / 2

        c.setFillColor(NAVY)
        c.rect(x0, y - hdr_h, total_w, hdr_h, stroke=0, fill=1)
        if accent_col is not None:
            ax = x0 + sum(w for _, w, _ in cols[:accent_col])
            c.setFillColor(NAVY_MID)
            c.rect(ax, y - hdr_h, cols[accent_col][1], hdr_h, stroke=0, fill=1)
        x = x0
        for i, (head, w, al) in enumerate(cols):
            c.setFillColor(GOLD if i == accent_col else WHITE)
            c.setFont("Helvetica-Bold", 7.6)
            parts = head.split("\n")
            for j, part in enumerate(parts):
                yy = y - hdr_h / 2 - 3 + (len(parts) - 1) * 4.5 - j * 9
                if al == "l":
                    c.drawString(x + 9, yy, part)
                elif al == "r":
                    c.drawRightString(x + w - 9, yy, part)
                else:
                    c.drawCentredString(x + w / 2, yy, part)
            x += w
        yy = y - hdr_h

        for idx, (cells, meta) in enumerate(rows):
            is_total = meta.get("total")
            h = row_h + (4 if is_total else 0)
            if is_total:
                c.setFillColor(NAVY)
            else:
                c.setFillColor(ZEBRA if idx % 2 else WHITE)
            c.rect(x0, yy - h, total_w, h, stroke=0, fill=1)
            if accent_col is not None and not is_total:
                ax = x0 + sum(w for _, w, _ in cols[:accent_col])
                c.setFillColor(meta.get("accent_bg", GOLD_SOFT))
                c.rect(ax, yy - h, cols[accent_col][1], h, stroke=0, fill=1)
            if not is_total:
                c.setStrokeColor(LINE)
                c.setLineWidth(0.4)
                c.line(x0, yy - h, x0 + total_w, yy - h)

            x = x0
            for i, (val, (head, w, al)) in enumerate(zip(cells, cols)):
                if is_total:
                    c.setFillColor(GOLD if i == accent_col else WHITE)
                    c.setFont("Helvetica-Bold", 9)
                elif i == accent_col:
                    c.setFillColor(meta.get("accent_fg", colors.HexColor("#8A6100")))
                    c.setFont("Helvetica-Bold", 10)
                elif i == bold_col:
                    c.setFillColor(NAVY_MID)
                    c.setFont("Helvetica-Bold", 8.6)
                else:
                    c.setFillColor(meta.get(f"fg{i}", INK))
                    c.setFont(meta.get(f"font{i}", "Helvetica"), 8.6)
                ty = yy - h / 2 - 3
                if al == "l":
                    c.drawString(x + 9, ty, str(val))
                elif al == "r":
                    c.drawRightString(x + w - 9, ty, str(val))
                else:
                    c.drawCentredString(x + w / 2, ty, str(val))
                x += w
            yy -= h
        return yy

    # -- the day-by-day chart ---------------------------------------------
    def pace_chart(self, x, y, w, h, b, split, plan):
        """
        Daily shop sales against the day's number.

        Green at or above the gold daily-target line, red below, each bar
        carrying its own figure. The cumulative required-vs-sold pair was
        tried alongside this and dropped (Abhay, 27 Jul 2026) -- the month
        position is already stated in the strip above, and one horizontal
        reference line is what makes a red bar self-explanatory.

        The per-day verdict is honest on this basis: the target is SHOP
        liquidation only, and shop sales accrue every day. It was removed
        while FED/BAR invoice was in the target, because 72% of green days
        were green only because a truck was invoiced.
        """
        c = self.c
        rows = b["rows"]
        as_of = plan["as_of"].day
        per_day = b["per_day"]

        c.setFillColor(colors.HexColor("#FBFCFE"))
        c.roundRect(x, y - h, w, h, 4, stroke=0, fill=1)
        c.setFillColor(NAVY_MID)
        c.setFont("Helvetica-Bold", 8.4)
        c.drawString(x + 12, y - 14, f"DAY BY DAY — {plan['month'].title()}")

        lx = x + 128
        for lab, col in (("daily target", GOLD), ("hit it", GREEN),
                         ("fell short", colors.HexColor("#E9A0A0"))):
            c.setFillColor(col)
            c.rect(lx, y - 17.5, 6, 6, stroke=0, fill=1)
            c.setFillColor(MUTED)
            c.setFont("Helvetica", 6.4)
            c.drawString(lx + 9, y - 16.5, lab)
            lx += 9 + c.stringWidth(lab, "Helvetica", 6.4) + 16

        pad_l, pad_r, pad_b, pad_t = 30, 16, 14, 30
        px, py = x + pad_l, y - h + pad_b
        pw, ph = w - pad_l - pad_r, h - pad_b - pad_t
        n = len(rows)
        bw = pw / max(n, 1)

        # Axis capped at the 85th percentile (or the target, whichever is
        # larger), NOT the maximum -- one outsized day otherwise flattens the
        # month. Over-cap bars print their true figure with a marker.
        tot = sorted(v for v in ((r["actual"] or 0) for r in rows
                                 if not r["dry"] and r["day"] <= as_of)
                     if v > 0.001)
        ref = max(tot[min(int(len(tot) * 0.85), len(tot) - 1)] if tot else 0.0,
                  per_day, 0.001)
        cap = ref * 1.45
        dp = 0 if ref >= 10 else 1

        c.setStrokeColor(colors.HexColor("#E6EAF2"))
        c.setLineWidth(0.4)
        for f in (0.5, 1.0):
            c.line(px, py + ph * f, px + pw, py + ph * f)

        for i, r in enumerate(rows):
            bx = px + i * bw
            if r["dry"]:
                c.setFillColor(colors.HexColor("#EDEFF4"))
                c.rect(bx + bw * .18, py, bw * .64, ph * .04, stroke=0, fill=1)
                continue
            if r["actual"] is None:
                continue
            hit = r["actual"] >= r["target_day"] - 1e-9
            bh = ph * (min(r["actual"], cap) / cap)
            c.setFillColor(GREEN if hit else colors.HexColor("#E9A0A0"))
            c.rect(bx + bw * .18, py, bw * .64, max(bh, 0.5), stroke=0, fill=1)
            c.setFillColor(GREEN if hit else RED)
            c.setFont("Helvetica-Bold", 5.6)
            lbl = f"{r['actual']:,.{dp}f}"
            if r["actual"] > cap:
                lbl = "\u25b2 " + lbl
            c.drawCentredString(bx + bw * .5, py + bh + 3.0, lbl)

        ty = py + ph * min(per_day / cap, 1.0)
        c.setStrokeColor(GOLD)
        c.setLineWidth(1.5)
        c.setDash(4, 3)
        c.line(px, ty, px + pw, ty)
        c.setDash()
        c.setFillColor(colors.HexColor("#8A6100"))
        c.setFont("Helvetica-Bold", 6.6)
        c.drawRightString(px - 4, ty - 2.4, f"{per_day:,.{dp}f}")

        c.setFillColor(MUTED)
        c.setFont("Helvetica", 6.2)
        for i, r in enumerate(rows):
            if r["day"] % 5 == 0 or r["day"] == 1:
                c.drawCentredString(px + i * bw + bw / 2, py - 10, str(r["day"]))
        return y - h


def _shop_daily_table(c, x, y, w, cols, rows, offset):
    """Half of the per-shop daily-average list."""
    c.setFillColor(NAVY)
    c.rect(x, y - 15, w, 15, stroke=0, fill=1)
    hx = x
    for i, (head, cw, al) in enumerate(cols):
        c.setFillColor(GOLD if i == 2 else WHITE)
        c.setFont("Helvetica-Bold", 6.2)
        if al == "l":
            c.drawString(hx + 6, ty_c(y, 15), head)
        else:
            c.drawCentredString(hx + cw / 2, ty_c(y, 15), head)
        hx += cw
    ty = y - 15

    def clip(t, font, size, maxw):
        if c.stringWidth(t, font, size) <= maxw:
            return t
        while t and c.stringWidth(t + "…", font, size) > maxw:
            t = t[:-1]
        return t + "…"

    for i, r in enumerate(rows):
        rh = 14.2
        c.setFillColor(ZEBRA if i % 2 else WHITE)
        c.rect(x, ty - rh, w, rh, stroke=0, fill=1)
        hx = x + sum(cw for _, cw, _ in cols[:2])
        c.setFillColor(colors.HexColor("#FFF9EC"))
        c.rect(hx, ty - rh, cols[2][1], rh, stroke=0, fill=1)

        hx = x
        vals = [str(offset + i + 1), r["shop"], f"{r['now']:,.1f}",
                f"{r['last']:,.1f}",
                ("·" if abs(r["delta"]) < 0.05 else
                 ("+" if r["delta"] > 0 else "") + f"{r['delta']:,.1f}")]
        for j, (val, (head, cw, al)) in enumerate(zip(vals, cols)):
            if j == 0:
                c.setFillColor(MUTED); c.setFont("Helvetica", 6.2)
            elif j == 1:
                c.setFillColor(NAVY_MID); c.setFont("Helvetica-Bold", 7.4)
            elif j == 2:
                c.setFillColor(colors.HexColor("#8A6100"))
                c.setFont("Helvetica-Bold", 8)
            elif j == 3:
                c.setFillColor(MUTED); c.setFont("Helvetica", 7.2)
            else:
                c.setFillColor(MUTED if abs(r["delta"]) < 0.05 else
                               (GREEN if r["delta"] > 0 else RED))
                c.setFont("Helvetica-Bold", 7.2)
            txt = clip(str(val), c._fontname, c._fontsize,
                       cw - (4 if j == 0 else 9))
            if al == "l":
                c.drawString(hx + 6, ty - 10, txt)
            else:
                c.drawCentredString(hx + cw / 2, ty - 10, txt)
            hx += cw
        ty -= rh
    return ty


def ty_c(y, h):
    return y - h / 2 - 2.2


# ---------------------------------------------------------------------------

def cluster_summary_page(pdf, plan, cluster, bonds, asm, staff):
    p = plan
    y = pdf.header(f"{p['month'].title()} {p['year']} · DAILY PACE",
                   f"CLUSTER {cluster}",
                   f"{asm or ''}   ·   sales through {p['as_of']:%d %B}   ·   "
                   f"all figures in cases")
    y -= 20

    tgt = sum(plan["bonds"][b]["target_month"] for b in bonds)
    sold = sum(plan["bonds"][b]["mtd_actual"] for b in bonds)
    exp = sum(plan["bonds"][b]["mtd_expected"] for b in bonds)
    need = sum(plan["bonds"][b]["needed_per_day"] for b in bonds)
    perday = sum(plan["bonds"][b]["per_day"] for b in bonds)
    dleft = plan["bonds"][bonds[0]]["days_left"]
    gap = sold - exp

    y = pdf.kpi_strip(y, [
        ("Cluster target", _fmt(tgt), f"{_fmt(perday)} cs a day", NAVY_SOFT, None),
        ("Sold so far", _fmt(sold),
         f"{sold / exp:.0%} of where you should be" if exp else "",
         GREEN, GREEN_BG),
        ("Should be at", _fmt(exp), f"by {p['as_of']:%d %b}", NAVY_SOFT, None),
        ("Behind by", _fmt(abs(gap)) if gap < 0 else "on pace",
         "shortfall to date" if gap < 0 else "keep going",
         RED if gap < 0 else GREEN, RED_BG if gap < 0 else GREEN_BG),
        ("Needed per day", _fmt(need), f"for the last {dleft} days", GOLD,
         GOLD_SOFT),
    ])
    y -= 22

    cols = [("BOND", 118, "l"), ("EXECUTIVE", 118, "l"),
            ("TARGET", 68, "c"), ("PER\nDAY", 55, "c"), ("SOLD", 68, "c"),
            ("SHOULD BE\nAT", 76, "c"), ("BEHIND\nBY", 68, "c"),
            ("PACE", 52, "c"), ("NEEDED\nPER DAY", 82, "c"),
            ("DAYS\nLEFT", 52, "c")]
    rows = []
    ranked = sorted(bonds, key=lambda b: -plan["bonds"][b]["pct_done"])
    for b in ranked:
        d = plan["bonds"][b]
        behind = d["gap"]
        rows.append(([
            b, (staff.get(b) or "— vacant —").title(),
            _fmt(d["target_month"]), _fmt(d["per_day"], 1),
            _fmt(d["mtd_actual"]), _fmt(d["mtd_expected"]),
            _fmt(behind) if behind < 0 else "on pace",
            f"{(d['mtd_actual'] / d['mtd_expected']) if d['mtd_expected'] else 0:.0%}",
            _fmt(d["needed_per_day"], 1),
            str(d["days_left"]),
        ], {
            "fg6": RED if behind < 0 else GREEN,
            "font6": "Helvetica-Bold",
            "fg1": MUTED if not staff.get(b) else INK,
            "fg7": RED if (d["mtd_actual"] / max(d["mtd_expected"], 1e-9)) < 0.7
                   else (AMBER if (d["mtd_actual"] / max(d["mtd_expected"], 1e-9)) < 1.0
                         else GREEN),
            "font7": "Helvetica-Bold",
        }))
    rows.append(([
        "CLUSTER TOTAL", "", _fmt(tgt), _fmt(perday, 1), _fmt(sold),
        _fmt(exp), _fmt(gap) if gap < 0 else "on pace",
        f"{sold / exp:.0%}" if exp else "·", _fmt(need, 1), str(dleft),
    ], {"total": True}))
    y = pdf.table(y, cols, rows, accent_col=8)

    y -= 26
    c = pdf.c
    c.setFillColor(colors.HexColor("#FBFCFE"))
    c.roundRect(M, y - 44, W - 2 * M, 44, 4, stroke=0, fill=1)
    c.setFillColor(GOLD)
    c.roundRect(M, y - 44, 3.2, 44, 1.6, stroke=0, fill=1)
    c.setFillColor(NAVY_MID)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(M + 14, y - 16, "HOW THE DAILY NUMBER WORKS")
    c.setFillColor(INK)
    c.setFont("Helvetica", 8.2)
    c.drawString(M + 14, y - 30,
                 f"Every bond's daily target is its monthly target divided by "
                 f"{plan['selling_days']} selling days. NEEDED PER DAY is what is "
                 f"left to sell divided by the days that remain —")
    c.drawString(M + 14, y - 40,
                 "so every day a bond falls short, its own number goes up the "
                 "next morning. Dry days carry no target and are spread over the "
                 "remaining days.")
    pdf.footer(f"Cluster {cluster} · daily pace · generated "
               f"{_dt.date.today():%d %b %Y}")
    pdf.c.showPage()


def bond_page(pdf, plan, bond, targets, staff_name):
    b = plan["bonds"][bond]
    p = plan
    behind = b["gap"]
    who = (staff_name.title() if staff_name
           else "VACANT — NO EXECUTIVE")
    pace_hdr = (b["mtd_actual"] / b["mtd_expected"]) if b["mtd_expected"] > 0 else 0
    y = pdf.header(bond, f"{pace_hdr:.0%}",
                   f"{who}   ·   Cluster {b['cluster']}   ·   "
                   f"shop sales through {p['as_of']:%d %B}")
    y -= 8

    c = pdf.c
    # --- TODAY | THE MONTH -----------------------------------------------
    # Redesigned 27 Jul 2026. The previous strip compared "you should sell"
    # against "you are selling" as two bars -- but the chart below already
    # makes that comparison once per day, 26 times, against the gold line. So
    # the strip now does what the chart cannot: state today's instruction,
    # and place the month.
    should = b["per_day"]
    actual = b["run_rate"]
    need = b["needed_per_day"]
    met = bool(b.get("target_met"))
    ok = b["gap"] >= 0
    mult = (need / actual) if actual > 0.01 else 0.0
    proj = b["mtd_actual"] + actual * b["days_left"]
    short = b["target_month"] - proj

    sh = 80
    lw = (W - 2 * M) * 0.365

    # LEFT — the instruction
    c.setFillColor(colors.HexColor("#1B5E20") if met else RED_DEEP)
    c.roundRect(M, y - sh, lw, sh, 4, stroke=0, fill=1)
    c.setFillColor(GOLD)
    c.roundRect(M, y - sh, 4, sh, 2, stroke=0, fill=1)
    pale = colors.HexColor("#C8E6C9") if met else colors.HexColor("#F0CFCF")
    c.setFillColor(pale)
    c.setFont("Helvetica-Bold", 7.2)
    c.drawString(M + 17, y - 16, "TARGET ALREADY MET" if met else "TODAY YOU NEED")
    big = f"+{b.get('surplus', 0):,.0f}" if met else f"{need:,.1f}"
    c.setFillColor(WHITE)
    c.setFont("Helvetica-Bold", 34)
    c.drawString(M + 16, y - 50, big)
    c.setFillColor(pale)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(M + 19 + c.stringWidth(big, "Helvetica-Bold", 34), y - 50,
                 "cs over" if met else "cs")
    c.setFont("Helvetica-Bold", 7.6)
    c.drawString(M + 17, y - 65,
                 "every case from here is upside" if met else
                 (f"{mult:,.1f}× the {actual:,.1f} a day you have been doing"
                  if mult > 0 else f"your normal day is {should:,.1f} cs"))
    c.setFont("Helvetica", 6.8)
    c.drawString(M + 17, y - 75.5,
                 f"{b['days_left']} days left  ·  your normal day is "
                 f"{should:,.1f} cs")

    # RIGHT — where the month stands.
    # The monthly SHOP total is deliberately NOT shown (Abhay, 27 Jul 2026).
    # It is a derived number — 80% of the bond target — and printing "136 of
    # 280 cs sold" invents a target nobody was given: the exec knows his bond
    # target is 350 and reads 280 as a different, quieter one. So the block is
    # framed entirely against PACE — where he should be by today — which is a
    # real position and needs no monthly total on the page.
    rx = M + lw + 10
    rw = (W - M) - rx
    c.setFillColor(colors.HexColor("#F5F7FB"))
    c.roundRect(rx, y - sh, rw, sh, 4, stroke=0, fill=1)
    c.setFillColor(NAVY_MID)
    c.setFont("Helvetica-Bold", 7.2)
    c.drawString(rx + 16, y - 16, "THE MONTH  ·  SHOP LIQUIDATION")

    c.setFillColor(GREEN if ok else RED)
    c.setFont("Helvetica-Bold", 26)
    c.drawString(rx + 15, y - 45, f"{b['mtd_actual']:,.0f}")
    vw = c.stringWidth(f"{b['mtd_actual']:,.0f}", "Helvetica-Bold", 26)
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 9)
    c.drawString(rx + 19 + vw, y - 45, "cs sold so far")

    pace = (b["mtd_actual"] / b["mtd_expected"]) if b["mtd_expected"] > 0 else 0
    c.setFillColor(GREEN if ok else RED)
    c.setFont("Helvetica-Bold", 20)
    c.drawRightString(W - M - 16, y - 42, f"{pace:.0%}")
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 6.2)
    c.drawRightString(W - M - 16, y - 51, "of where you should be by today")

    # bar: the whole track IS where he should be by today, so a full bar = on pace
    bx, bw_, byy = rx + 16, rw - 32, y - 62
    c.setFillColor(colors.HexColor("#E4E9F2"))
    c.roundRect(bx, byy, bw_, 8, 4, stroke=0, fill=1)
    c.setFillColor(GREEN if ok else RED)
    c.roundRect(bx, byy, max(bw_ * min(pace, 1.0), 3), 8, 4, stroke=0, fill=1)

    c.setFillColor(GREEN if ok else RED)
    c.setFont("Helvetica-Bold", 7.6)
    c.drawString(rx + 16, y - 75.5,
                 (f"ahead by {abs(b['gap']):,.0f} cs" if ok else
                  f"behind by {abs(b['gap']):,.0f} cs") +
                 (f"  ·  at this rate you finish {abs(short):,.0f} cs "
                  f"{'short' if short > 0 else 'clear'}"))
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 6.8)
    c.drawRightString(W - M - 16, y - 75.5,
                      f"you should be at {b['mtd_expected']:,.0f} cs by today")
    y -= sh + 6

    # --- day-by-day chart, sized to whatever the shop table does not need -
    # The shop table is BOTTOM-ANCHORED so it always finishes on the page
    # margin, and the chart expands to meet it. Bonds with few outlets used to
    # leave a block of dead white space under the table; now that height goes
    # to the chart instead. KOTTAYAM (25 shops, 9 rows a column) is the
    # binding case and gets the smallest chart.
    shops = targets.get(bond, [])
    ncol = 3
    per_col = -(-len(shops) // ncol) if shops else 1
    # Bottom margin trimmed from 40pt to 26pt (Abhay, 27 Jul 2026 —
    # "push this part more down"). ~9mm, still inside any printer's
    # safe area, and the 18pt it frees goes to the chart.
    tbl_top = 26 + 31 + per_col * 14.2
    ch_h = max(min(y - 12 - tbl_top, 250), 150)

    pdf.pace_chart(M, y, W - 2 * M, ch_h, b, p.get("split", {}), p)
    y = min(y - ch_h - 12, tbl_top)

    # --- sales per shop, month to date -----------------------------------
    c.setFillColor(NAVY_MID)
    c.setFont("Helvetica-Bold", 8.4)
    c.drawString(M, y - 10, "SALES PER SHOP — MONTH TO DATE")
    if shops:
        tot_now = sum(r["now"] for r in shops)
        tot_last = sum(r["last"] for r in shops)
        c.setFillColor(NAVY_MID)
        c.setFont("Helvetica-Bold", 7.4)
        c.drawRightString(W - M, y - 10,
                          f"KSBC SHOPS: {tot_now:,.0f} cs   "
                          f"(same days last month {tot_last:,.0f})")
    y -= 16

    third = (W - 2 * M - 12 * (ncol - 1)) / ncol
    cols = [("#", 20, "c"), ("SHOP", third - 144, "l"),
            ("THIS MO", 44, "c"), ("LAST MO", 40, "c"), ("+/−", 40, "c")]
    for side in range(ncol):
        x = M + side * (third + 12)
        chunk = shops[side * per_col:(side + 1) * per_col]
        if chunk:
            _shop_daily_table(c, x, y, third, cols, chunk, side * per_col)
    if not shops:
        c.setFillColor(MUTED)
        c.setFont("Helvetica-Oblique", 8)
        c.drawString(M, y - 16, "No KSBC shops mapped to this bond.")

    pdf.footer(f"{bond} · {(staff_name or 'Vacant').title()} · daily pace "
               f"through {plan['as_of']:%d %b %Y}")
    pdf.c.showPage()


def build_cluster_pdfs(plan, targets, master, clusters, asms, outdir) -> list:
    staff = {}
    for info in master.values():
        if info["status"].upper() == "CLOSED":
            continue
        s = (info["staff"] or "").strip()
        if info["bond"] and s and s.upper() != "VACANT":
            staff.setdefault(info["bond"], s)

    os.makedirs(outdir, exist_ok=True)
    made = []
    for cl in sorted(clusters):
        bonds = [b for b in clusters[cl] if b in plan["bonds"]]
        bonds.sort(key=lambda b: -plan["bonds"][b]["pct_done"])
        path = os.path.join(outdir, f"DAILY PACE - CLUSTER {cl}.pdf")
        pdf = PacePDF(path, plan.get("base", ""))
        cluster_summary_page(pdf, plan, cl, bonds, asms.get(cl), staff)
        for b in bonds:
            bond_page(pdf, plan, b, targets, staff.get(b))
        pdf.c.save()
        made.append(path)
    return made
