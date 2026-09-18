#!/usr/bin/env python3
"""
SHOPSALES COMPARATIVE — mobile-friendly (portrait) rebuild.

Same report, same design language as the landscape original
(comparative_shopsales_current-7.pdf); re-laid out on A4 PORTRAIT so it fills a
phone screen at fit-width instead of rendering as a tiny landscape strip.

Everything is data-driven: change DATA / TITLE / PERIOD and rebuild.
Column widths auto-fit the page, so longer bond names or bigger numbers
re-flow rather than clip.
"""

from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color
from reportlab.pdfbase.pdfmetrics import stringWidth

# ── palette (sampled from the original PDF) ──────────────────────────────────
NAVY       = Color(0.043, 0.161, 0.310)   # header bands, cluster/total rows
GOLD       = Color(1.000, 0.741, 0.192)   # accent band, cluster text, rules
ZEBRA      = Color(0.960, 0.970, 0.990)   # alternating row tint
PINK       = Color(1.000, 0.800, 0.820)   # sell-through column wash
WHITE      = Color(1, 1, 1)
INK        = Color(0.059, 0.098, 0.176)   # numeric text
INK_SOFT   = Color(0.157, 0.157, 0.157)   # bond names, stock-net %
RED        = Color(0.776, 0.157, 0.157)   # sell-through %
RED_TREND  = Color(0.812, 0.075, 0.133)   # negative trend on light rows
GREEN      = Color(0.247, 0.525, 0.000)   # positive trend on light rows
GREEN_LT   = Color(0.565, 0.933, 0.565)   # positive trend on navy rows
PINK_LT    = Color(1.000, 0.714, 0.757)   # negative trend on navy rows

GRID       = Color(0.839, 0.867, 0.914)   # column separators through the body
GRID_W     = 0.4

# ── sell-through rating tiers ────────────────────────────────────────────────
# Identical rule and palette to the KSBC SHOP SALES ANALYSIS workbook:
#   >=80% High Performance · >=60% Balanced · >=40% Inventory Heavy · else
#   Critical Overstock.  (fill, dark text, bright text for dark bands)
TIERS = [
    (80.0, Color(0.733, 0.871, 0.984), Color(0.082, 0.396, 0.753), Color(0.310, 0.765, 0.969)),
    (60.0, Color(0.863, 0.929, 0.784), Color(0.180, 0.490, 0.196), Color(0.506, 0.780, 0.518)),
    (40.0, Color(1.000, 0.878, 0.698), Color(0.902, 0.318, 0.000), Color(1.000, 0.718, 0.302)),
    (-1e9, Color(1.000, 0.804, 0.824), Color(0.776, 0.157, 0.157), Color(0.898, 0.451, 0.451)),
]


def tier_of(pct):
    """pct is the printed string, e.g. '51.68%'."""
    v = float(str(pct).replace("%", "").replace(",", ""))
    for t in TIERS:
        if v >= t[0]:
            return t[1:]


# ── report content ───────────────────────────────────────────────────────────
TITLE   = "K.S DISTILLERY"
SUBTITLE = "SHOP SALES -ANALYSIS"
PERIOD  = "1 Aug 2026 to 16 Aug 2026"

HEADERS = ["BOND", "OPENING", "RECEIPT", "SALES", "CLOSING",
           "STOCK NET", "STOCK NET %", "SELL-THROUGH %"]
HEADERS_WRAPPED = [["BOND"], ["OPENING"], ["RECEIPT"], ["SALES"], ["CLOSING"],
                   ["STOCK", "NET"], ["STOCK", "NET %"], ["SELL-", "THROUGH %"]]
GROUP_HEADER = "AVERAGE SALES / DAY"
SUB_HEADERS = ["CM", "LM", "TREND"]

# kind: "bond" = normal row, "sum" = cluster / overall band
# cols: opening, receipt, sales, closing, stock_net, stock_net_pct,
#       sell_through, cm, lm, trend
DATA = [
    ("bond", "ALAPPUZHA", "414.28", "108.00", "269.94", "252.34", "-161.94", "-39.09%", "51.68%", "16.87", "24.01", "-7.14"),
    ("bond", "ATTINGAL", "523.92", "366.00", "265.84", "624.08", "100.16", "19.12%", "29.87%", "16.61", "15.04", "+1.58"),
    ("bond", "KOLLAM", "476.81", "266.00", "401.28", "341.53", "-135.28", "-28.37%", "54.02%", "25.08", "24.44", "+0.64"),
    ("bond", "KOTTARAKARA", "694.35", "123.89", "297.98", "520.26", "-174.09", "-25.07%", "36.42%", "18.62", "14.09", "+4.53"),
    ("bond", "NEDUMANGAD", "714.33", "139.00", "260.66", "592.67", "-121.66", "-17.03%", "30.55%", "16.29", "14.51", "+1.78"),
    ("bond", "PATHANAMTHITTA", "305.42", "27.00", "43.96", "288.47", "-16.95", "-5.55%", "13.22%", "2.75", "4.77", "-2.02"),
    ("sum", "CLUSTER - 1", "3129.11", "1029.89", "1539.66", "2619.35", "-509.76", "-16.29%", "37.02%", "96.23", "96.86", "-0.63"),
    ("bond", "ALUVA", "316.82", "28.00", "23.83", "320.99", "4.17", "1.32%", "6.91%", "1.49", "1.85", "-0.37"),
    ("bond", "KOTTAYAM", "863.17", "406.23", "316.45", "952.95", "89.78", "10.40%", "24.93%", "19.78", "14.44", "+5.33"),
    ("bond", "THODUPUZHA", "486.26", "65.00", "67.86", "483.40", "-2.86", "-0.59%", "12.31%", "4.24", "5.09", "-0.85"),
    ("bond", "THRISSUR", "340.66", "169.00", "145.07", "364.59", "23.93", "7.02%", "28.46%", "9.07", "14.36", "-5.29"),
    ("bond", "TRIPUNITHURA", "519.51", "270.92", "238.28", "552.14", "32.63", "6.28%", "30.15%", "14.89", "6.93", "+7.96"),
    ("sum", "CLUSTER - 2", "2526.42", "939.15", "791.49", "2674.07", "147.65", "5.84%", "22.84%", "49.47", "42.67", "+6.79"),
    ("bond", "KANNUR", "208.93", "3.00", "64.27", "147.66", "-61.27", "-29.33%", "30.33%", "4.02", "15.10", "-11.09"),
    ("bond", "KOZHIKODE", "361.16", "315.00", "292.76", "383.40", "22.24", "6.16%", "43.30%", "18.30", "19.62", "-1.33"),
    ("bond", "PALAKKAD", "366.85", "767.08", "334.93", "799.01", "432.16", "117.80%", "29.54%", "20.93", "23.14", "-2.21"),
    ("bond", "PERINTHALMANNA", "98.81", "604.00", "129.43", "573.38", "474.57", "480.29%", "18.42%", "8.09", "10.28", "-2.19"),
    ("sum", "CLUSTER - 3", "1035.75", "1689.08", "821.39", "1903.45", "867.70", "83.78%", "30.14%", "51.34", "68.15", "-16.81"),
    ("grand", "TOTAL", "6691.28", "3658.12", "3152.54", "7196.87", "505.59", "7.56%", "30.46%", "197.03", "207.68", "-10.65"),
]

# ── page geometry (A4 portrait) ──────────────────────────────────────────────
PW, PH = 595.276, 841.890

TITLE_H  = 46.0      # navy masthead
BAND_H   = 34.0      # gold subtitle band
HDR_TOP  = 30.0      # header: group row
HDR_BOT  = 32.0      # header: column labels
HDR_H    = HDR_TOP + HDR_BOT

F_TITLE  = 20
F_BAND   = 13
F_DATA   = 11        # starting point; auto-shrunk only if the table won't fit
F_HDR    = 8.5
PAD      = 4.0       # per side, inside every column
BOTTOM_M = 0.0       # none — the TOTAL band finishes flush with the page edge
HDR_RULE      = 1.6  # gold grid inside the header block
HDR_RULE_EDGE = 2.2  # gold rules closing the top and bottom of the header


def fit_columns(f_data, f_hdr, pad):
    """Widths that hold the widest datum and the widest header line."""
    need = []

    # BOND column: left-aligned, so it carries an extra indent
    w = max(stringWidth(r[1], "Helvetica-Bold" if r[0] in ("sum", "grand") else "Helvetica", f_data)
            for r in DATA)
    w = max(w, stringWidth("BOND", "Helvetica-Bold", f_hdr))
    need.append(w + pad * 3)

    # seven numeric columns under single headers
    for i in range(7):
        w = max(stringWidth(r[2 + i], "Helvetica-Bold" if r[0] in ("sum", "grand") else "Helvetica", f_data)
                for r in DATA)
        w = max(w, max(stringWidth(l, "Helvetica-Bold", f_hdr) for l in HEADERS_WRAPPED[i + 1]))
        need.append(w + pad * 2)

    # CM / LM
    for i in (9, 10):
        w = max(stringWidth(r[i], "Helvetica-Bold" if r[0] in ("sum", "grand") else "Helvetica", f_data)
                for r in DATA)
        w = max(w, stringWidth(SUB_HEADERS[0 if i == 9 else 1], "Helvetica-Bold", f_hdr))
        need.append(w + pad * 2)

    # TREND: arrow + gap + value
    arrow = f_data * 0.52
    w = max(stringWidth(r[11], "Helvetica-Bold", f_data) for r in DATA) + arrow + f_data * 0.30
    w = max(w, stringWidth("TREND", "Helvetica-Bold", f_hdr))
    need.append(w + pad * 2)

    # the CM/LM/TREND group must also hold its banner
    grp = sum(need[8:])
    banner = stringWidth(GROUP_HEADER, "Helvetica-Bold", f_hdr) + pad * 2
    if grp < banner:
        extra = (banner - grp) / 3.0
        for i in range(8, 11):
            need[i] += extra

    return need


def solve():
    """Largest data font that still fits the page width."""
    f = F_DATA
    while f > 6:
        f_hdr = max(7.0, min(F_HDR, f - 2.0))
        need = fit_columns(f, f_hdr, PAD)
        if sum(need) <= PW:
            slack = PW - sum(need)
            need = [w + slack * (w / sum(need)) for w in need]   # spread evenly
            return f, f_hdr, need
        f -= 0.25
    raise SystemExit("table will not fit")


F_DATA, F_HDR, COLW = solve()
XS = [0.0]
for w in COLW:
    XS.append(XS[-1] + w)

ROW_H = (PH - TITLE_H - BAND_H - HDR_H - BOTTOM_M) / len(DATA)
ROW_H = min(ROW_H, 38.0)


def centred(c, text, font, size, colour, x0, x1, y_mid):
    c.setFont(font, size)
    c.setFillColor(colour)
    w = stringWidth(text, font, size)
    c.drawString((x0 + x1) / 2.0 - w / 2.0, y_mid - size * 0.35, text)


def build(path):
    c = canvas.Canvas(path, pagesize=(PW, PH))
    c.setTitle(f"{SUBTITLE} {PERIOD}")

    # ── masthead ─────────────────────────────────────────────────────────────
    y = PH - TITLE_H
    c.setFillColor(NAVY)
    c.rect(0, y, PW, TITLE_H, stroke=0, fill=1)
    centred(c, TITLE, "Helvetica-Bold", F_TITLE, GOLD, 0, PW, y + TITLE_H / 2)

    # ── gold band: report name + period ──────────────────────────────────────
    y -= BAND_H
    c.setFillColor(GOLD)
    c.rect(0, y, PW, BAND_H, stroke=0, fill=1)
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", F_BAND)
    c.drawString(12, y + BAND_H / 2 - F_BAND * 0.35, SUBTITLE)
    c.drawRightString(PW - 12, y + BAND_H / 2 - F_BAND * 0.35, PERIOD)

    # ── table header ─────────────────────────────────────────────────────────
    hdr_top_y = y - HDR_TOP
    hdr_bot_y = hdr_top_y - HDR_BOT
    c.setFillColor(NAVY)
    c.rect(0, hdr_bot_y, PW, HDR_H, stroke=0, fill=1)

    # single-span headers (rows merged vertically)
    for i in range(8):
        lines = HEADERS_WRAPPED[i]
        mid = hdr_bot_y + HDR_H / 2
        start = mid + (len(lines) - 1) * (F_HDR + 1.5) / 2.0
        for j, line in enumerate(lines):
            centred(c, line, "Helvetica-Bold", F_HDR, WHITE,
                    XS[i], XS[i + 1], start - j * (F_HDR + 1.5))

    # AVERAGE SALES / DAY group banner + its three sub-columns
    centred(c, GROUP_HEADER, "Helvetica-Bold", F_HDR, GOLD,
            XS[8], XS[11], hdr_top_y + HDR_TOP / 2)
    for j, lab in enumerate(SUB_HEADERS):
        centred(c, lab, "Helvetica-Bold", F_HDR, WHITE,
                XS[8 + j], XS[9 + j], hdr_bot_y + HDR_BOT / 2)

    # gold rules framing the header block
    c.setStrokeColor(GOLD)
    c.setLineCap(0)
    # column separators — full height, except inside the AVERAGE SALES / DAY
    # group where only the sub-columns divide
    c.setLineWidth(HDR_RULE)
    for i in range(1, 11):
        top_of = HDR_H if i <= 8 else HDR_BOT
        c.line(XS[i], hdr_bot_y, XS[i], hdr_bot_y + top_of)
    # under the group banner, and under the whole header block
    c.line(XS[8], hdr_bot_y + HDR_BOT, XS[11], hdr_bot_y + HDR_BOT)
    c.setLineWidth(HDR_RULE_EDGE)
    c.line(0, hdr_bot_y + HDR_H, PW, hdr_bot_y + HDR_H)
    c.line(0, hdr_bot_y, PW, hdr_bot_y)

    # ── body ─────────────────────────────────────────────────────────────────
    ry = hdr_bot_y
    for idx, (kind, name, *vals) in enumerate(DATA):
        ry -= ROW_H
        is_grand = kind == "grand"                 # the one final line
        is_sum = kind in ("sum", "grand")
        # cluster bands are navy-on-gold-text; the grand total inverts to a gold
        # band with navy text so it can't be mistaken for another cluster
        band = GOLD if is_grand else NAVY
        ink = NAVY if is_grand else GOLD

        if is_sum:
            c.setFillColor(band)
            c.rect(0, ry, PW, ROW_H, stroke=0, fill=1)
        else:
            # zebra alternates across ALL table rows (summary bands included in
            # the count, then painted over navy) — matches the original exactly
            if idx % 2 == 0:
                c.setFillColor(ZEBRA)
                c.rect(0, ry, PW, ROW_H, stroke=0, fill=1)
            # sell-through cell — KSBC rating tier, not a flat wash
            c.setFillColor(tier_of(vals[6])[0])
            c.rect(XS[7], ry, XS[8] - XS[7], ROW_H, stroke=0, fill=1)
            # column separators — carried down from the header, light so they
            # guide the eye without boxing every figure in. Drawn per row, which
            # leaves the navy / gold bands solid and joins seamlessly elsewhere.
            c.setStrokeColor(GRID)
            c.setLineWidth(GRID_W)
            # inset the verticals wherever the row abuts a rule, so a grey tick
            # never lands on the gold (the rule is stroke-centred, so it covers
            # half its width either side of the boundary)
            above_is_rule = idx == 0 or DATA[idx - 1][0] in ("sum", "grand")
            below_is_rule = idx + 1 < len(DATA) and DATA[idx + 1][0] in ("sum", "grand")
            top_in = (HDR_RULE_EDGE if idx == 0 else 1.6) / 2.0 if above_is_rule else 0.0
            bot_in = 1.6 / 2.0 if below_is_rule else 0.0
            for x in XS[1:-1]:
                c.line(x, ry + bot_in, x, ry + ROW_H - top_in)
            # row separator on the top edge — skipped where a band sits above,
            # since that edge already carries its gold rule
            if not above_is_rule:
                c.line(0, ry + ROW_H, PW, ry + ROW_H)

        mid = ry + ROW_H / 2

        # bond / cluster label — centred like every other column
        centred(c, name, "Helvetica-Bold" if is_sum else "Helvetica", F_DATA,
                ink if is_sum else INK_SOFT, XS[0], XS[1], mid)

        # opening → closing
        for i in range(4):
            centred(c, vals[i], "Helvetica-Bold" if is_sum else "Helvetica", F_DATA,
                    ink if is_sum else INK, XS[1 + i], XS[2 + i], mid)

        # stock net, stock net %
        centred(c, vals[4], "Helvetica-Bold" if is_sum else "Helvetica", F_DATA,
                ink if is_sum else INK, XS[5], XS[6], mid)
        centred(c, vals[5], "Helvetica-Bold" if is_sum else "Helvetica", F_DATA,
                ink if is_sum else INK_SOFT, XS[6], XS[7], mid)

        # sell-through % — tier text: dark on the light cells and on the gold
        # TOTAL band, bright variant on the navy cluster bands
        t = tier_of(vals[6])
        centred(c, vals[6], "Helvetica-Bold", F_DATA,
                t[2] if (is_sum and not is_grand) else t[1], XS[7], XS[8], mid)

        # CM / LM
        for j in (0, 1):
            centred(c, vals[7 + j], "Helvetica-Bold" if is_sum else "Helvetica", F_DATA,
                    ink if is_sum else INK, XS[8 + j], XS[9 + j], mid)

        # trend: arrow + magnitude — the arrow carries the sign, so the +/-
        # is stripped from the printed value (never from the data)
        tv = vals[9]
        up = not tv.startswith("-")
        tv = tv.lstrip("+-")
        # a zero move is flat — no arrow, neutral ink. An up arrow on 0.00 would
        # read as growth.
        flat = abs(float(tv.replace(",", "") or 0)) < 0.005
        if flat:
            tcol = ink if is_sum else INK_SOFT
        elif is_sum and not is_grand:
            tcol = GREEN_LT if up else PINK_LT     # pastels for navy bands
        else:
            tcol = GREEN if up else RED_TREND      # dark pair on light/gold
        a = 0.0 if flat else F_DATA * 0.52
        gap = 0.0 if flat else F_DATA * 0.30
        tw = stringWidth(tv, "Helvetica-Bold", F_DATA)
        total = a + gap + tw
        ax = (XS[10] + XS[11]) / 2.0 - total / 2.0
        ay = mid - a / 2.0
        if not flat:
            c.setFillColor(tcol)
            p = c.beginPath()
            if up:
                p.moveTo(ax + a / 2, ay + a); p.lineTo(ax, ay); p.lineTo(ax + a, ay)
            else:
                p.moveTo(ax + a / 2, ay); p.lineTo(ax, ay + a); p.lineTo(ax + a, ay + a)
            p.close()
            c.drawPath(p, stroke=0, fill=1)
        c.setFillColor(tcol)
        c.setFont("Helvetica-Bold", F_DATA)
        c.drawString(ax + a + gap, mid - F_DATA * 0.35, tv)

        # gold rules above/below every summary band. The bottom rule is inset by
        # half its width where the band sits on the page edge — a stroke centred
        # on y=0 loses half its ink off the page.
        if is_sum:
            c.setStrokeColor(NAVY if is_grand else GOLD)
            c.setLineWidth(1.6)
            c.line(0, ry + ROW_H, PW, ry + ROW_H)
            # skip the bottom rule when the next row is itself a band — its own
            # top rule sits on the same line and would bury this one
            if not (idx + 1 < len(DATA) and DATA[idx + 1][0] in ("sum", "grand")):
                c.line(0, max(ry, 0.8), PW, max(ry, 0.8))

    c.showPage()
    c.save()
    return path


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "SHOPSALES COMPARATIVE - MOBILE.pdf"
    build(out)
    print(f"data font {F_DATA:.2f}pt | row {ROW_H:.1f}pt | table width {sum(COLW):.1f}/{PW:.0f}pt")
    print("wrote", out)
