#!/usr/bin/env python3
"""BOND LIQUIDATION SCORECARD — mobile-friendly rebuild.

Same report, same format, same numbers as the source landscape PDF
(AUGUST vs JULY 2026 - DAYS 1-22), re-flowed to a phone-width portrait
page: one metric block per page so every figure is legible at
fit-to-width on a phone.
"""
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor

# ---------------------------------------------------------------- palette
NAVY      = HexColor("#0B2C52")
NAVY_DK   = HexColor("#081F3B")
GOLD      = HexColor("#FAAF19")
RED       = HexColor("#CF1322")
GREEN     = HexColor("#3F8600")
RED_LT    = HexColor("#FF7875")
GREEN_LT  = HexColor("#52C41A")
INK       = HexColor("#282828")
ZEBRA     = HexColor("#F5F7FC")
WHITE     = HexColor("#FFFFFF")
TOTAL_BG  = HexColor("#2B353F")
AVG_BG    = HexColor("#3D4A57")
GRID      = HexColor("#D6DCE8")
MUTED     = HexColor("#8A93A2")

# ---------------------------------------------------------------- data
# 16 cells per row: 4 blocks x (AUG, JUL, dCS, d%)
ROWS = [
    ("ALAPPUZHA",              "bond",    ["322","561","-240","-43%",  "117","877","-760","-87%",   "-","91","-91","-100%",   "322","652","-331","-51%"]),
    ("ATTINGAL",               "bond",    ["395","371","24","6%",      "499","411","88","22%",     "10","31","-21","-68%",   "405","402","3","1%"]),
    ("KOLLAM",                 "bond",    ["510","586","-76","-13%",   "428","1030","-602","-58%", "70","200","-130","-65%", "580","786","-206","-26%"]),
    ("KOTTARAKARA",            "bond",    ["387","309","78","25%",     "367","439","-72","-16%",   "-","30","-30","-100%",   "387","339","48","14%"]),
    ("NEDUMANGAD",             "bond",    ["326","312","15","5%",      "399","425","-26","-6%",    "221","186","35","19%",   "547","498","50","10%"]),
    ("PATHANAMTHITTA",         "bond",    ["55","108","-52","-48%",    "52","165","-113","-68%",   "11","13","-2","-15%",    "66","121","-54","-45%"]),
    ("CLUSTER - 1 TOTAL",      "cluster", ["1996","2247","-252","-11%","1862","3347","-1485","-44%","312","551","-239","-43%","2308","2798","-491","-18%"]),
    ("ALUVA",                  "bond",    ["30","39","-10","-25%",     "52","70","-18","-26%",     "20","40","-20","-50%",   "50","79","-30","-38%"]),
    ("KOTTAYAM",               "bond",    ["433","383","50","13%",     "684","1043","-359","-34%", "232","172","60","35%",   "665","555","110","20%"]),
    ("THODUPUZHA",             "bond",    ["87","127","-40","-31%",    "125","161","-36","-22%",   "35","14","21","150%",    "122","141","-19","-13%"]),
    ("THRISSUR",               "bond",    ["213","301","-88","-29%",   "587","314","273","87%",    "117","99","18","18%",    "330","400","-70","-17%"]),
    ("TRIPUNITHURA",           "bond",    ["333","211","122","58%",    "609","431","178","41%",    "255","126","129","102%", "588","337","251","74%"]),
    ("CLUSTER - 2 TOTAL",      "cluster", ["1096","1062","35","3%",    "2057","2019","38","2%",    "659","451","208","46%",  "1755","1513","243","16%"]),
    ("KANNUR",                 "bond",    ["75","350","-275","-79%",   "23","491","-468","-95%",   "12","107","-95","-89%",  "87","457","-370","-81%"]),
    ("KOZHIKODE",              "bond",    ["395","491","-96","-20%",   "826","1011","-185","-18%", "388","311","77","25%",   "783","802","-19","-2%"]),
    ("PALAKKAD",               "bond",    ["490","491","-1","0%",      "1286","649","637","98%",   "280","177","103","58%",  "770","668","102","15%"]),
    ("PERINTHALMANNA",         "bond",    ["259","190","69","37%",     "658","109","549","504%",   "15","-","15","-",        "274","190","84","44%"]),
    ("CLUSTER - 3 TOTAL",      "cluster", ["1219","1522","-302","-20%","2793","2260","533","24%",  "695","595","100","17%",  "1914","2117","-202","-10%"]),
    ("UNMAPPED",               "bond",    ["-","-","-","-",            "-","7","-7","-100%",       "-","7","-7","-100%",     "-","7","-7","-100%"]),
    ("UNMAPPED CLUSTER TOTAL", "cluster", ["-","-","-","-",            "-","7","-7","-100%",       "-","7","-7","-100%",     "-","7","-7","-100%"]),
    ("TOTAL",                  "total",   ["4311","4831","-519","-11%","6713","7633","-920","-12%","1666","1604","62","4%",  "5977","6435","-457","-7%"]),
    ("AVERAGE DAILY SALE",     "avg",     ["196","220","-24","-11%",   "305","347","-42","-12%",   "76","73","3","4%",       "272","292","-21","-7%"]),
]

BLOCKS = [
    ("SHOP LIQUIDATION (KSBC)", "KSBC shop sales - tertiary"),
    ("SECONDARY SALES",         "Warehouse to outlets - all channels"),
    ("FED / BAR INVOICE",       "Consumer Fed + BAR invoice sales"),
    ("TOTAL LIQUIDATION",       "Shop liquidation + Fed / BAR invoice"),
]

PERIOD = "AUGUST vs JULY 2026  ·  DAYS 1–22  ·  cases"

# ---------------------------------------------------------------- geometry
PW, PH = 400, 612
M      = 8
X0     = M
COLS   = [128, 60, 60, 74, 62]          # Bond, AUG, JUL, d CS, d %
XS     = [X0]
for w in COLS:
    XS.append(XS[-1] + w)
TW     = sum(COLS)

H_HERO, H_BAND, H_HEAD, H_ROW, H_FOOT = 44, 26, 22, 22.0, 20


def is_neg(t):
    return t.startswith("-") and t not in ("-",)


def tri(c, cx, cy, up, col, s=4.6):
    """small solid triangle marker"""
    c.setFillColor(col)
    p = c.beginPath()
    if up:
        p.moveTo(cx, cy + s * 0.62)
        p.lineTo(cx - s * 0.62, cy - s * 0.48)
        p.lineTo(cx + s * 0.62, cy - s * 0.48)
    else:
        p.moveTo(cx, cy - s * 0.62)
        p.lineTo(cx - s * 0.62, cy + s * 0.48)
        p.lineTo(cx + s * 0.62, cy + s * 0.48)
    p.close()
    c.drawPath(p, fill=1, stroke=0)


def delta_cell(c, x, w, ybase, txt, font, size, pos_col, neg_col, flat_col, ref=None):
    """triangle + value, centred as one group.

    Direction follows the block's Δ CS value (`ref`) so that a rounded
    "0%" beside a negative case move still reads as a fall, exactly as
    the source report renders it.
    """
    if txt in ("-", "", None):
        c.setFont("Helvetica", size)
        c.setFillColor(flat_col)
        c.drawCentredString(x + w / 2, ybase, "–")
        return
    src = ref if ref not in (None, "-", "") else txt
    neg = is_neg(src)
    zero = src.lstrip("-").rstrip("%") in ("0", "0.0")
    col = flat_col if zero else (neg_col if neg else pos_col)
    c.setFont(font, size)
    tw = c.stringWidth(txt, font, size)
    gap = 3.0
    total = 5.7 + gap + tw
    sx = x + (w - total) / 2
    if not zero:
        tri(c, sx + 2.9, ybase + 2.9, not neg, col)
    c.setFillColor(col)
    c.drawString(sx + 5.7 + gap, ybase, txt)


def plain_cell(c, x, w, ybase, txt, font, size, col, flat_col):
    c.setFont(font, size)
    if txt in ("-", "", None):
        c.setFillColor(flat_col)
        c.drawCentredString(x + w / 2, ybase, "–")
    else:
        c.setFillColor(col)
        c.drawCentredString(x + w / 2, ybase, txt)


def draw_page(c, bi):
    title, sub = BLOCKS[bi]
    y = PH

    # ---- hero band
    y -= H_HERO
    c.setFillColor(NAVY)
    c.rect(0, y, PW, H_HERO, fill=1, stroke=0)
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(PW / 2, y + 24, "K.S DISTILLERY")
    c.setFillColor(WHITE)
    c.setFont("Helvetica-Bold", 9)
    c.drawCentredString(PW / 2, y + 10, "BOND LIQUIDATION SCORECARD")

    # ---- gold section band
    y -= H_BAND
    c.setFillColor(GOLD)
    c.rect(0, y, PW, H_BAND, fill=1, stroke=0)
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 11.5)
    c.drawString(M, y + 8.5, title)
    c.setFont("Helvetica-Bold", 8)
    c.drawRightString(PW - M, y + 9, "%d of %d" % (bi + 1, len(BLOCKS)))

    # ---- period strip
    y -= 15
    c.setFillColor(NAVY_DK)
    c.rect(0, y, PW, 15, fill=1, stroke=0)
    c.setFillColor(GOLD)
    c.setFont("Helvetica-Bold", 7.6)
    c.drawString(M, y + 4.5, PERIOD)
    c.setFillColor(HexColor("#9FB0C7"))
    c.setFont("Helvetica", 7.2)
    c.drawRightString(PW - M, y + 4.5, sub)

    # ---- column header
    y -= H_HEAD
    c.setFillColor(NAVY)
    c.rect(X0, y, TW, H_HEAD, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(XS[0] + 6, y + 7, "Bond")
    c.setFillColor(GOLD)
    for i, h in enumerate(["AUG", "JUL", "Δ CS", "Δ %"], start=1):
        c.drawCentredString(XS[i] + COLS[i] / 2, y + 7, h)

    # ---- rows
    zi = 0
    for label, kind, cells in ROWS:
        vals = cells[bi * 4:bi * 4 + 4]
        y -= H_ROW
        yb = y + 7.0

        if kind == "bond":
            bg = ZEBRA if zi % 2 else WHITE
            zi += 1
            c.setFillColor(bg)
            c.rect(X0, y, TW, H_ROW, fill=1, stroke=0)
            c.setStrokeColor(GRID)
            c.setLineWidth(0.4)
            c.line(X0, y, X0 + TW, y)
            c.setFillColor(NAVY)
            c.setFont("Helvetica-Bold", 8.6)
            c.drawString(XS[0] + 6, yb, label)
            plain_cell(c, XS[1], COLS[1], yb, vals[0], "Helvetica-Bold", 10, INK, MUTED)
            plain_cell(c, XS[2], COLS[2], yb, vals[1], "Helvetica", 10, HexColor("#5A6472"), MUTED)
            delta_cell(c, XS[3], COLS[3], yb, vals[2], "Helvetica-Bold", 9.6, GREEN, RED, MUTED)
            delta_cell(c, XS[4], COLS[4], yb, vals[3], "Helvetica-Bold", 9.6, GREEN, RED, MUTED, ref=vals[2])
        else:
            fill = NAVY if kind == "cluster" else (TOTAL_BG if kind == "total" else AVG_BG)
            lab_col = GOLD if kind == "cluster" else WHITE
            zi = 0
            c.setFillColor(fill)
            c.rect(X0, y, TW, H_ROW, fill=1, stroke=0)
            c.setFillColor(lab_col)
            size = 8.6 if len(label) < 20 else 7.6
            c.setFont("Helvetica-Bold", size)
            c.drawString(XS[0] + 6, yb, label)
            plain_cell(c, XS[1], COLS[1], yb, vals[0], "Helvetica-Bold", 10, lab_col, HexColor("#8FA0B8"))
            plain_cell(c, XS[2], COLS[2], yb, vals[1], "Helvetica-Bold", 10,
                       HexColor("#C9D4E4") if kind == "cluster" else HexColor("#C9CFD6"), HexColor("#8FA0B8"))
            delta_cell(c, XS[3], COLS[3], yb, vals[2], "Helvetica-Bold", 9.6, GREEN_LT, RED_LT, HexColor("#8FA0B8"))
            delta_cell(c, XS[4], COLS[4], yb, vals[3], "Helvetica-Bold", 9.6, GREEN_LT, RED_LT, HexColor("#8FA0B8"), ref=vals[2])

    # frame
    c.setStrokeColor(NAVY)
    c.setLineWidth(0.9)
    c.rect(X0, y, TW, PH - H_HERO - H_BAND - 15 - H_HEAD - y + (y - y), fill=0, stroke=0)

    # ---- footer
    c.setFillColor(HexColor("#6B7684"))
    c.setFont("Helvetica", 6.6)
    c.drawCentredString(PW / 2, 9,
                        "KSBC tertiary + Consumer Fed + BAR  ·  Total Liquidation = Shop + Fed/BAR  "
                        "·  average daily sale ÷ 22 days")


def main(out):
    c = canvas.Canvas(out, pagesize=(PW, PH))
    c.setTitle("BOND LIQUIDATION SCORECARD - AUG vs JUL 2026 (mobile)")
    for bi in range(len(BLOCKS)):
        draw_page(c, bi)
        c.showPage()
    c.save()
    print("wrote", out)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "scorecard_mobile.pdf")
