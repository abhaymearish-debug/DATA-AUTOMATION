#!/usr/bin/env python3
"""
ITEM ISSUE CONSOLIDATION (secondary sales, month vs month) — mobile-friendly rebuild.

Same report, same design as the landscape original
(item_issue_consolidation_2026-08-3.pdf); re-laid out on A4 PORTRAIT so it fills
a phone screen at fit-width instead of rendering as a tiny landscape strip.

Data-driven: edit the DATA / header constants and rebuild. Column widths are
measured, not hardcoded, so longer warehouse names or bigger numbers re-flow.
"""

from datetime import date

from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color
from reportlab.pdfbase.pdfmetrics import stringWidth

NB = " "          # non-breaking space: keeps a header label atomic

# ── palette (sampled from the original PDF) ──────────────────────────────────
NAVY     = Color(0.043, 0.161, 0.310)   # bands, header, TOTAL row
GOLD     = Color(1.000, 0.741, 0.192)   # title + header label text
AMBER    = Color(1.000, 0.750, 0.000)   # cluster row background
CREAM    = Color(1.000, 0.980, 0.900)   # prior-month column block
WHITE    = Color(1, 1, 1)
BLACK    = Color(0, 0, 0)
RED      = Color(0.753, 0.000, 0.000)   # negative difference
GREEN    = Color(0.216, 0.337, 0.137)   # positive difference
GREY     = Color(0.549, 0.549, 0.549)   # page footer
HAIRLINE = Color(0, 0, 0)

# ── report content ───────────────────────────────────────────────────────────
TITLE    = "K.S DISTILLERY"
FOOTER   = "Page 1 of 1"

# The two as-on dates drive every date string on the page — change these two
# lines for next month and the captions, banners and title all follow.
CUR_ASON   = date(2026, 8, 10)
PRIOR_ASON = date(2026, 7, 10)
PRIOR_END  = date(2026, 7, 31)      # the LAST MONTH reference column


def longdate(d):
    return f"{d.day} {d:%B} {d.year}"          # 10 August 2026


def mon(d):
    return f"{d:%b}".upper()                   # AUG


SUBLEFT  = (f"SECONDARY SALES · {mon(CUR_ASON)} {CUR_ASON:%y}"
            f" vs {mon(PRIOR_ASON)} {PRIOR_ASON:%y}")
SUBRIGHT = f"AS ON {longdate(CUR_ASON)}"

CUR_BANNER   = f"{CUR_ASON:%B}".upper()   + f" {CUR_ASON.year} - as on {longdate(CUR_ASON)}"
PRIOR_BANNER = f"{PRIOR_ASON:%B}".upper() + f" {PRIOR_ASON.year} - as on {longdate(PRIOR_ASON)}"
DIFF_BANNER  = "DIFFERENCE"

# per column: (header lines, group)  group: None = spans both header rows
COLS = [
    (["WAREHOUSE"],                   None),
    (["STN"],                         "cur"),
    (["GTN"],                         "cur"),
    (["TOTAL"],                       "cur"),
    ([f"C{NB}FED"],                   "cur"),
    (["BAR"],                         "cur"),
    (["10Aug"],                       "cur"),
    (["STN"],                         "prior"),
    (["GTN"],                         "prior"),
    (["TOTAL"],                       "prior"),
    ([f"C{NB}FED"],                   "prior"),
    (["BAR"],                         "prior"),
    (["10Jul"],                       "prior"),
    (["Cases"],                       "diff"),
    (["%"],                           "diff"),
    (["LAST MONTH", f"({PRIOR_END.day}{NB}{mon(PRIOR_END)})"], None),  # NB keeps it atomic
]

CUR_COLS   = [i for i, c in enumerate(COLS) if c[1] == "cur"]
PRIOR_COLS = [i for i, c in enumerate(COLS) if c[1] == "prior"]
DIFF_COLS  = [i for i, c in enumerate(COLS) if c[1] == "diff"]

LABEL_COL   = 0                 # WAREHOUSE / cluster / TOTAL label
FIRST_VAL   = 1                 # first column carrying a value
LAST_COL    = len(COLS) - 1     # LAST MONTH
CUR_TOTAL   = CUR_COLS[-1]      # the 10Aug column, printed navy bold
PRIOR_TOTAL = PRIOR_COLS[-1]    # the 10Jul column

# rows: ("wh", sl, name, [15 values]) | ("cluster"/"total", label, [15 values])
#       ("foot", label, cur_value, prior_value, cases, pct)
# value order: STN GTN TOTAL CFED BAR 10Aug | STN GTN TOTAL CFED BAR 10Jul | Cases % LastMonth
DATA = [
    ("wh", "1",  "BALARAMAPURAM",  ["0","10","10","0","0","10",       "0","85","85","0","0","85",         "-75","-88.24%","415"]),
    ("wh", "2",  "NEDUMANGAD",     ["105","43","148","50","0","198",  "0","55","55","128","0","183",      "15","8.2%","432"]),
    ("wh", "3",  "ATTINGAL",       ["216","76","292","0","0","292",   "20","41","61","20","3","84",       "208","247.62%","303"]),
    ("wh", "4",  "MENAMKULAM",     ["14","12","26","40","0","66",     "0","18","18","1","0","19",         "47","247.37%","300"]),
    ("wh", "5",  "KOLLAM",         ["3","241","244","40","0","284",   "17","449","466","40","0","506",    "-222","-43.87%","1,477"]),
    ("wh", "6",  "KARUNAGAPPALLY", ["0","5","5","0","0","5",          "17","45","62","0","0","62",        "-57","-91.94%","102"]),
    ("wh", "7",  "KOTTARAKARA",    ["19","45","64","0","0","64",      "0","121","121","30","0","151",     "-87","-57.62%","589"]),
    ("wh", "8",  "PATHANAMTHITA",  ["0","3","3","0","3","6",          "0","19","19","8","0","27",         "-21","-77.78%","104"]),
    ("wh", "9",  "THIRUVALLA",     ["0","19","19","0","0","19",       "22","19","41","0","0","41",        "-22","-53.66%","87"]),
    ("wh", "10", "ALAPPUZHA",      ["50","41","91","0","0","91",      "140","200","340","46","0","386",   "-295","-76.42%","1,059"]),
    ("cluster", "CLUSTER - 1",     ["407","495","902","130","3","1,035", "216","1,052","1,268","273","3","1,544", "-509","-33%","4,868"]),
    ("wh", "11", "KOTTAYAM",       ["89","49","138","143","0","281",  "256","64","320","90","0","410",    "-129","-31.46%","656"]),
    ("wh", "12", "AYARKKUNNAM",    ["65","7","72","41","0","113",     "115","55","170","2","0","172",     "-59","-34.3%","477"]),
    ("wh", "13", "THODUPUZHA",     ["0","6","6","0","0","6",          "29","10","39","3","0","42",        "-36","-85.71%","54"]),
    ("wh", "14", "TRIPUNITHURA",   ["74","49","123","90","0","213",   "35","18","53","56","0","109",      "104","95.41%","392"]),
    ("wh", "15", "KADAVANTHRA",    ["59","18","77","0","0","77",      "0","7","7","0","0","7",            "70","1000%","312"]),
    ("wh", "16", "PERUMBAVOOR",    ["0","0","0","0","0","0",          "0","7","7","0","0","7",            "-7","-100%","8"]),
    ("wh", "17", "KOTHAMANGALAM",  ["36","24","60","0","0","60",      "0","30","30","10","0","40",        "20","50%","134"]),
    ("wh", "18", "ALUVA",          ["0","1","1","0","0","1",          "0","3","3","0","0","3",            "-2","-66.67%","5"]),
    ("wh", "19", "CHALAKUDY",      ["0","85","85","0","0","85",       "23","41","64","15","0","79",       "6","7.59%","147"]),
    ("wh", "20", "THRISSUR",       ["0","46","46","70","0","116",     "4","9","13","50","0","63",         "53","84.13%","293"]),
    ("cluster", "CLUSTER - 2",     ["323","285","608","344","0","952", "462","244","706","226","0","932", "20","2%","2,478"]),
    ("wh", "21", "PALAKKAD",       ["0","29","29","10","0","39",      "107","105","212","115","0","327",  "-288","-88.07%","678"]),
    ("wh", "22", "MENONPARA",      ["365","120","485","73","0","558", "20","21","41","0","0","41",        "517","1260.98%","102"]),
    ("wh", "23", "PERINTHALMANNA", ["85","59","144","0","0","144",    "30","57","87","0","0","87",        "57","65.52%","139"]),
    ("wh", "24", "KOZHIKODE",      ["58","21","79","0","0","79",      "0","36","36","26","7","69",        "10","14.49%","246"]),
    ("wh", "25", "NADUVANNUR",     ["81","16","97","62","0","159",    "30","61","91","109","0","200",     "-41","-20.5%","680"]),
    ("wh", "26", "KALPETTA",       ["0","10","10","25","0","35",      "30","21","51","33","0","84",       "-49","-58.33%","353"]),
    ("wh", "27", "KANNUR",         ["0","1","1","2","0","3",          "0","17","17","45","0","62",        "-59","-95.16%","163"]),
    ("wh", "28", "BATTATHUR",      ["0","0","0","10","0","10",        "75","78","153","15","0","168",     "-158","-94.05%","339"]),
    ("cluster", "CLUSTER - 3",     ["589","256","845","182","0","1,027", "292","396","688","343","7","1,038", "-11","-1%","2,700"]),
    ("total",   "TOTAL",           ["1,319","1,036","2,355","656","3","3,014", "970","1,692","2,662","842","10","3,514", "-500","-14%","10,046"]),
    ("foot", "Day Sale",       "657", "475", "182",  "38%"),
    ("foot", "Industry Total", "0",   "100", "-100", "-100%"),
]

# ── page geometry (A4 portrait) ──────────────────────────────────────────────
PW, PH = 595.276, 841.890

TITLE_H  = 40.0
SUB_H    = 26.0
HDR_TOP  = 24.0      # group-banner row
HDR_BOT  = 26.0      # column-label row
HDR_H    = HDR_TOP + HDR_BOT
FOOT_H   = 26.0      # page-number strip
BOTTOM_M = 10.0

F_TITLE  = 16
F_SUB    = 11       # report name + as-on date
F_DATA   = 10.5      # starting point; auto-shrunk only if the table won't fit
F_FOOT   = 8
PAD      = 2.75      # per side, inside every column. Mirroring the Aug/Jul
                     # columns costs ~9pt, which pushed the 10.5pt layout 1.9pt
                     # over the page; trimming a quarter-point of padding buys
                     # it back, and the equal-slack pass returns the padding.
HAIR     = 0.43
HDR_RULE      = 1.4  # gold grid inside the header block
HDR_RULE_EDGE = 2.0  # gold rules closing the header block top and bottom


def hdr_font(fd):
    return max(7.0, min(9.5, fd - 1.75))


def row_values(r):
    """The 15 body values a row contributes to each column, or None where merged."""
    if r[0] in ("wh", "cluster", "total"):
        return r[3] if r[0] == "wh" else r[2]
    return None


def fit_columns(fd, fh, pad):
    need = []
    for i, (lines, _) in enumerate(COLS):
        # widest datum in this column
        w = 0.0
        for r in DATA:
            vals = row_values(r)
            if vals is None:                      # merged footer row
                if r[0] == "foot" and i == DIFF_COLS[0]:
                    w = max(w, stringWidth(r[4], "Helvetica-Bold", fd))
                if r[0] == "foot" and i == DIFF_COLS[1]:
                    w = max(w, stringWidth(r[5], "Helvetica-Bold", fd))
                continue
            if i == LABEL_COL:
                txt = r[2] if r[0] == "wh" else r[1]
                w = max(w, stringWidth(txt, "Helvetica-Bold" if r[0] != "wh" else "Helvetica", fd))
            else:
                w = max(w, stringWidth(vals[i - FIRST_VAL],
                                       "Helvetica-Bold" if r[0] != "wh" else "Helvetica", fd))
        # widest unbreakable header word
        for line in lines:
            for word in line.split(" "):
                w = max(w, stringWidth(word.replace(NB, " "), "Helvetica-Bold", fh))
        # the label column is left-aligned with a 2×pad indent
        need.append(w + (pad * 4 if i == LABEL_COL else pad * 2))

    # the footer labels sit in the label column too
    lab = max(stringWidth(r[1], "Helvetica-Bold", fd) for r in DATA if r[0] == "foot")
    need[LABEL_COL] = max(need[LABEL_COL], lab + pad * 2)

    # AUGUST and JULY carry the same six measures and are read ACROSS, so each
    # pair is mirrored to the wider of the two. Sized independently, Aug STN
    # ("1,319") came out 9pt wider than Jul STN ("970") and the two banners --
    # which should be identical blocks -- rendered at different widths.
    for a, pr in zip(CUR_COLS, PRIOR_COLS):
        need[a] = need[pr] = max(need[a], need[pr])

    # The two period banners differ in length, so the LARGER requirement is
    # applied to both groups; sizing each to its own banner would undo the
    # mirroring above. Same for the merged Day Sale / Industry Total cells.
    foot = [r for r in DATA if r[0] == "foot"]
    want = max(stringWidth(CUR_BANNER, "Helvetica-Bold", fh),
               stringWidth(PRIOR_BANNER, "Helvetica-Bold", fh),
               max(stringWidth(r[2], "Helvetica-Bold", fd) for r in foot),
               max(stringWidth(r[3], "Helvetica-Bold", fd) for r in foot)) + pad * 2
    have = sum(need[i] for i in CUR_COLS)
    if have < want:
        extra = (want - have) / len(CUR_COLS)
        for i in CUR_COLS + PRIOR_COLS:
            need[i] += extra

    # DIFFERENCE spans only its own two columns
    have = sum(need[i] for i in DIFF_COLS)
    want = stringWidth(DIFF_BANNER, "Helvetica-Bold", fh) + pad * 2
    if have < want:
        extra = (want - have) / len(DIFF_COLS)
        for i in DIFF_COLS:
            need[i] += extra
    return need


def solve():
    fd = F_DATA
    while fd > 5.5:
        fh = hdr_font(fd)
        need = fit_columns(fd, fh, PAD)
        if sum(need) <= PW:
            # leftover width is shared EQUALLY, not in proportion. Proportional
            # sharing widens the already-wide columns and makes the rhythm more
            # uneven; an equal share also keeps the Aug/Jul mirroring intact.
            share = (PW - sum(need)) / len(need)
            return fd, fh, [w + share for w in need]
        fd -= 0.25
    raise SystemExit("table will not fit")


F_DATA, F_HDR, COLW = solve()
XS = [0.0]
for w in COLW:
    XS.append(XS[-1] + w)

ROW_H = (PH - TITLE_H - SUB_H - HDR_H - FOOT_H - BOTTOM_M) / len(DATA)


def wrap(text, font, size, maxw):
    """Greedy word wrap; NB-spaces never break."""
    words, lines, cur = text.split(" "), [], ""
    for wd in words:
        trial = wd if not cur else cur + " " + wd
        if stringWidth(trial.replace(NB, " "), font, size) <= maxw or not cur:
            cur = trial
        else:
            lines.append(cur); cur = wd
    if cur:
        lines.append(cur)
    return lines


def centred(c, text, font, size, colour, x0, x1, y_mid, dy=0.0):
    c.setFont(font, size); c.setFillColor(colour)
    t = text.replace(NB, " ")
    c.drawString((x0 + x1) / 2.0 - stringWidth(t, font, size) / 2.0,
                 y_mid - size * 0.35 + dy, t)


def cell(c, x0, x1, y, h, fill=None, edge=None, lw=None):
    """One table cell. `edge`/`lw` override the default hairline border --
    the header block uses a thick gold grid."""
    # ONE rect that both fills and strokes — the source does the same, and
    # drawing fill and border separately would double every object on the page
    c.setStrokeColor(edge or HAIRLINE); c.setLineWidth(lw or HAIR)
    if fill is not None:
        c.setFillColor(fill)
    c.rect(x0, y, x1 - x0, h, stroke=1, fill=1 if fill is not None else 0)


def diff_colour(v, on_navy):
    if on_navy:
        return WHITE
    return RED if v.strip().startswith("-") else GREEN


def build(path):
    c = canvas.Canvas(path, pagesize=(PW, PH))
    c.setTitle(f"{SUBLEFT} {SUBRIGHT}")

    # ── masthead + sub-band (one navy block, two text rows) ──────────────────
    y = PH - TITLE_H
    c.setFillColor(NAVY); c.rect(0, y, PW, TITLE_H, stroke=0, fill=1)
    centred(c, TITLE, "Helvetica-Bold", F_TITLE, GOLD, 0, PW, y + TITLE_H / 2)

    y -= SUB_H
    c.setFillColor(NAVY); c.rect(0, y, PW, SUB_H, stroke=0, fill=1)
    c.setFont("Helvetica-Bold", F_SUB); c.setFillColor(GOLD)
    c.drawString(12, y + SUB_H / 2 - F_SUB * 0.35, SUBLEFT)
    c.drawRightString(PW - 12, y + SUB_H / 2 - F_SUB * 0.35, SUBRIGHT)

    # ── header ──────────────────────────────────────────────────────────────
    top_y = y - HDR_TOP
    bot_y = top_y - HDR_BOT

    for i, (lines, group) in enumerate(COLS):
        if group is None:                                   # spans both rows
            cell(c, XS[i], XS[i + 1], bot_y, HDR_H, NAVY, GOLD, HDR_RULE)
            flat = []
            for ln in lines:
                flat += wrap(ln, "Helvetica-Bold", F_HDR, COLW[i] - PAD * 2)
            start = bot_y + HDR_H / 2 + (len(flat) - 1) * (F_HDR + 1.2) / 2.0
            for j, ln in enumerate(flat):
                centred(c, ln, "Helvetica-Bold", F_HDR, GOLD,
                        XS[i], XS[i + 1], start - j * (F_HDR + 1.2))
        else:                                               # label row only
            cell(c, XS[i], XS[i + 1], bot_y, HDR_BOT, NAVY, GOLD, HDR_RULE)
            centred(c, lines[0], "Helvetica-Bold", F_HDR, GOLD,
                    XS[i], XS[i + 1], bot_y + HDR_BOT / 2)

    for banner, idxs in ((CUR_BANNER, CUR_COLS), (PRIOR_BANNER, PRIOR_COLS), (DIFF_BANNER, DIFF_COLS)):
        x0, x1 = XS[idxs[0]], XS[idxs[-1] + 1]
        cell(c, x0, x1, top_y, HDR_TOP, NAVY, GOLD, HDR_RULE)
        centred(c, banner, "Helvetica-Bold", F_HDR, GOLD, x0, x1, top_y + HDR_TOP / 2)

    c.setStrokeColor(GOLD); c.setLineWidth(HDR_RULE_EDGE); c.setLineCap(0)
    inset = HDR_RULE_EDGE / 2.0
    c.line(0, top_y + HDR_TOP - inset, PW, top_y + HDR_TOP - inset)
    c.line(0, bot_y + inset, PW, bot_y + inset)

    # ── body ────────────────────────────────────────────────────────────────
    ry = bot_y
    for r in DATA:
        ry -= ROW_H
        mid = ry + ROW_H / 2
        kind = r[0]

        if kind == "foot":
            _, label, curv, priorv, cases, pct = r
            cell(c, XS[LABEL_COL], XS[LABEL_COL + 1], ry, ROW_H, WHITE)
            centred(c, label, "Helvetica-Bold", F_DATA, BLACK,
                    XS[LABEL_COL], XS[LABEL_COL + 1], mid)
            cell(c, XS[CUR_COLS[0]], XS[CUR_COLS[-1] + 1], ry, ROW_H, WHITE)
            centred(c, curv, "Helvetica-Bold", F_DATA, NAVY,
                    XS[CUR_COLS[0]], XS[CUR_COLS[-1] + 1], mid)
            cell(c, XS[PRIOR_COLS[0]], XS[PRIOR_COLS[-1] + 1], ry, ROW_H, CREAM)
            centred(c, priorv, "Helvetica-Bold", F_DATA, NAVY,
                    XS[PRIOR_COLS[0]], XS[PRIOR_COLS[-1] + 1], mid)
            for idx, v in zip(DIFF_COLS, (cases, pct)):
                cell(c, XS[idx], XS[idx + 1], ry, ROW_H, WHITE)
                centred(c, v, "Helvetica-Bold", F_DATA, diff_colour(v, False),
                        XS[idx], XS[idx + 1], mid)
            cell(c, XS[LAST_COL], XS[LAST_COL + 1], ry, ROW_H, WHITE)
            continue

        is_cluster, is_total = kind == "cluster", kind == "total"
        bg   = AMBER if is_cluster else (NAVY if is_total else None)
        font = "Helvetica" if kind == "wh" else "Helvetica-Bold"
        ink  = WHITE if is_total else BLACK
        vals = r[3] if kind == "wh" else r[2]

        if kind == "wh":                       # warehouse name, left-aligned
            cell(c, XS[LABEL_COL], XS[LABEL_COL + 1], ry, ROW_H, WHITE)
            c.setFont("Helvetica", F_DATA); c.setFillColor(BLACK)
            c.drawString(XS[LABEL_COL] + PAD * 2, mid - F_DATA * 0.35, r[2])
        else:                                  # cluster / total label, centred
            cell(c, XS[LABEL_COL], XS[LABEL_COL + 1], ry, ROW_H, bg)
            centred(c, r[1], "Helvetica-Bold", F_DATA, ink,
                    XS[LABEL_COL], XS[LABEL_COL + 1], mid)

        for i in range(FIRST_VAL, len(COLS)):
            v = vals[i - FIRST_VAL]
            wash = bg if bg is not None else (CREAM if i in PRIOR_COLS else WHITE)
            cell(c, XS[i], XS[i + 1], ry, ROW_H, wash)
            if i in (CUR_TOTAL, PRIOR_TOTAL):  # period totals
                col = WHITE if is_total else NAVY
                f   = "Helvetica-Bold"
            elif i in DIFF_COLS:
                col = diff_colour(v, is_total)
                f   = "Helvetica-Bold"
            else:
                col, f = ink, font
            centred(c, v, f, F_DATA, col, XS[i], XS[i + 1], mid)

    # ── page footer ─────────────────────────────────────────────────────────
    centred(c, FOOTER, "Helvetica", F_FOOT, GREY, 0, PW, ry - FOOT_H / 2)

    c.showPage(); c.save()
    return path


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "ITEM ISSUE CONSOLIDATION - MOBILE.pdf"
    build(out)
    print(f"data {F_DATA:.2f}pt | header {F_HDR:.2f}pt | row {ROW_H:.1f}pt | "
          f"table {sum(COLW):.1f}/{PW:.0f}pt")
    print("wrote", out)
