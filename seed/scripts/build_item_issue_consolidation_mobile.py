#!/usr/bin/env python3
"""
ITEM ISSUE CONSOLIDATION (secondary sales, month vs month) — mobile A4 portrait.

Data-driven: replace DATA and set CUR_ASON / PRIOR_ASON / PRIOR_END for the next
month. Column widths are measured, not hardcoded, so longer warehouse names or
bigger figures re-flow; the data font steps down in 0.25pt increments rather
than clipping.

FORMATTING PASS — 19 Sep 2026. Same 16 columns, same figures, same periods; only
the grid, banding, colour and type treatment changed. What changed and why:

  * The boxed grid is gone. Every cell used to carry a full black hairline box,
    which put ~540 boxes on the page and made a 16-column table read as noise.
    Now: a light rule between rows, a light vertical INSIDE each read-group, and
    a navy vertical only where a group ENDS (GROUP_EDGES) — so WAREHOUSE | AUG |
    JUL | DIFFERENCE | LAST MONTH read as five blocks instead of sixteen columns.
  * Zebra banding on the warehouse rows, tinted per block (ZEBRA in the white
    zones, CREAM_Z inside the prior-month block) so a row can be tracked across
    the full width without losing which period you are in.
  * A nil value is drawn in ZERO grey. Roughly a third of the cells are "0"; at
    full ink they competed with the figures that matter. 115 cells on the 5 Aug
    build — every one verified to be a literal "0".
  * Cluster bands were full-chroma amber across all 16 columns, which shouted
    louder than the TOTAL row beneath them. Now a soft gold tint with navy bold
    text and gold rules closing it, so the hierarchy runs warehouse → cluster →
    TOTAL in ascending weight rather than the middle term winning.
  * LAST MONTH is a reference column, so it sits on a neutral grey with softer
    ink (INK_SOFT) instead of competing with the live figures.
  * Row heights are weighted (WEIGHT), so a sub-total and the TOTAL row are
    physically taller than a warehouse line.
  * The header block is one navy field with gold group separators only at
    GROUP_EDGES; the old 1.4pt gold grid between every header cell has gone.
  * Period-total headers are uppercase ("5 AUG", "31 JUL") to match STN / GTN /
    TOTAL / C FED / BAR — they were mixed-case.
  * The merged Day Sale / Industry Total rows carry no internal grid (the light
    verticals stop at grid_bot) and sit on their own grey strip.

MASTHEAD / HEADER PASS — 19 Sep 2026 (second round, figures still untouched):

  * The stacked title band + strapline band became ONE masthead: a gold accent
    bar at the paper edge, a gold rule anchoring a left lockup (letterspaced
    wordmark over a muted-steel report line), and a gold AS ON pill on the
    right with the unit note beneath it. It is 8pt SHORTER than the two bands
    it replaced, so the table gained the height.
  * The period banners are chips, inset so navy gutters separate them. The
    LIVE period is a gold chip with navy text; the prior period and DIFFERENCE
    sit back on a lighter navy. Which month you are reading is now visible at a
    glance instead of being two lines of identical gold text.
  * NO LOGO — Abhay's call, 19 Sep 2026. The gold bar stands in for it. Do not
    reinstate the KSD shield here.

TRAP, hit once and worth keeping in mind: letterspacing is set with `Tc` on a
text object, and `Tc` is part of the PDF TEXT STATE, which SURVIVES the BT/ET
block. Leaving it set letterspaced every later drawString on the page -- it
silently widened the whole table and broke every centred cell. `tracked()`
resets it to 0 before the text object closes.

Verified against the source PDF by `verify_item_issue_consolidation.py`:
488 figures identical across 34 rows, every red/green flag still matching the
sign of its value.
"""
from datetime import date
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color
from reportlab.pdfbase.pdfmetrics import stringWidth

NB = " "

def hexc(h, a=1):
    h = h.lstrip("#")
    return Color(int(h[0:2],16)/255, int(h[2:4],16)/255, int(h[4:6],16)/255, a)

# ── palette ──────────────────────────────────────────────────────────────────
NAVY       = hexc("0B2950")   # masthead, header block, TOTAL row
NAVY_SOFT  = hexc("1B3C68")   # group dividers in the body
NAVY_CHIP  = hexc("17355F")   # prior-period / difference chips in the header
STEEL      = hexc("A9B8D6")   # masthead eyebrow, secondary header ink
GOLD       = hexc("FFBD31")   # header text, rules
GOLD_DIM   = hexc("8A6B28")   # intra-group separators inside the header block
CLUSTER_BG = hexc("FFEFC4")   # sub-total band (was full-chroma amber)
CLUSTER_RL = hexc("E8A317")   # rules closing the cluster band
CREAM      = hexc("FFFAE9")   # prior-month block
CREAM_Z    = hexc("FAF1D6")   # prior-month block, zebra row
WHITE      = hexc("FFFFFF")
ZEBRA      = hexc("EFF3F9")
REF_BG     = hexc("F1F3F6")   # LAST MONTH reference column
REF_Z      = hexc("E4E8EE")
FOOT_BG    = hexc("EDEFF3")   # Day Sale / Industry Total strip
INK        = hexc("14181F")
INK_SOFT   = hexc("55606E")   # reference-column figures
ZERO       = hexc("AEB5C0")   # a nil value recedes
RULE       = hexc("D9DEE6")   # light row / intra-group rules
RED        = hexc("B3261E")
GREEN      = hexc("1B6B3A")
GREY       = hexc("9AA1AC")

# ── report content ───────────────────────────────────────────────────────────
TITLE  = "K.S DISTILLERY"
FOOTER = "Page 1 of 1"

CUR_ASON   = date(2026, 8, 5)
PRIOR_ASON = date(2026, 7, 31)
PRIOR_END  = date(2026, 7, 31)

longdate = lambda d: f"{d.day} {d:%B} {d.year}"
mon      = lambda d: f"{d:%b}".upper()

SUBLEFT  = f"SECONDARY SALES  ·  {mon(CUR_ASON)} {CUR_ASON:%y} vs {mon(PRIOR_ASON)} {PRIOR_ASON:%y}"
SUBRIGHT = f"AS ON {longdate(CUR_ASON)}"
CUR_BANNER   = f"{CUR_ASON:%B}".upper()   + f" {CUR_ASON.year}  ·  as on {longdate(CUR_ASON)}"
PRIOR_BANNER = f"{PRIOR_ASON:%B}".upper() + f" {PRIOR_ASON.year}  ·  as on {longdate(PRIOR_ASON)}"
DIFF_BANNER  = "DIFFERENCE"

COLS = [
    (["WAREHOUSE"], None),
    (["STN"], "cur"), (["GTN"], "cur"), (["TOTAL"], "cur"),
    ([f"C{NB}FED"], "cur"), (["BAR"], "cur"),
    ([f"{CUR_ASON.day}{NB}{mon(CUR_ASON)}"], "cur"),
    (["STN"], "prior"), (["GTN"], "prior"), (["TOTAL"], "prior"),
    ([f"C{NB}FED"], "prior"), (["BAR"], "prior"),
    ([f"{PRIOR_ASON.day}{NB}{mon(PRIOR_ASON)}"], "prior"),
    (["Cases"], "diff"), (["%"], "diff"),
    (["LAST MONTH", f"({PRIOR_END.day}{NB}{mon(PRIOR_END)})"], None),
]
CUR_COLS   = [i for i,c in enumerate(COLS) if c[1]=="cur"]
PRIOR_COLS = [i for i,c in enumerate(COLS) if c[1]=="prior"]
DIFF_COLS  = [i for i,c in enumerate(COLS) if c[1]=="diff"]
LABEL_COL, FIRST_VAL = 0, 1
LAST_COL    = len(COLS)-1
CUR_TOTAL   = CUR_COLS[-1]
PRIOR_TOTAL = PRIOR_COLS[-1]
# vertical navy dividers close each read-group
GROUP_EDGES = [1, CUR_COLS[-1]+1, PRIOR_COLS[-1]+1, DIFF_COLS[-1]+1]

# rows: ('wh', sl, name, [15 values]) | ('cluster'/'total', label, [15 values])
#       ('foot', label, cur_value, prior_value, cases, pct)
# value order: STN GTN TOTAL CFED BAR <cur> | STN GTN TOTAL CFED BAR <prior> | Cases % LastMonth
DATA = [
    ("wh", "1",  "BALARAMAPURAM",   ["0","10","10","0","0","10",  "193","222","415","0","0","415",  "-405","-97.59%","415"]),
    ("wh", "2",  "NEDUMANGAD",      ["71","0","71","20","0","91",  "31","191","222","208","2","432",  "-341","-78.94%","432"]),
    ("wh", "3",  "ATTINGAL",        ["30","48","78","0","0","78",  "82","198","280","20","3","303",  "-225","-74.26%","303"]),
    ("wh", "4",  "MENAMKULAM",      ["0","2","2","10","0","12",  "135","131","266","34","0","300",  "-288","-96%","300"]),
    ("wh", "5",  "KOLLAM",          ["0","12","12","15","0","27",  "195","1,057","1,252","225","0","1,477",  "-1,450","-98.17%","1,477"]),
    ("wh", "6",  "KARUNAGAPPALLY",  ["0","2","2","0","0","2",  "23","79","102","0","0","102",  "-100","-98.04%","102"]),
    ("wh", "7",  "KOTTARAKARA",     ["0","32","32","0","0","32",  "267","292","559","30","0","589",  "-557","-94.57%","589"]),
    ("wh", "8",  "PATHANAMTHITA",   ["0","0","0","0","0","0",  "51","35","86","18","0","104",  "-104","-100%","104"]),
    ("wh", "9",  "THIRUVALLA",      ["0","0","0","0","0","0",  "51","36","87","0","0","87",  "-87","-100%","87"]),
    ("wh", "10", "ALAPPUZHA",       ["0","18","18","0","0","18",  "513","455","968","91","0","1,059",  "-1,041","-98.3%","1,059"]),
    ("cluster", "CLUSTER - 1",     ["101","124","225","45","0","270",  "1,541","2,696","4,237","626","5","4,868",  "-4,598","-94%","4,868"]),
    ("wh", "11", "KOTTAYAM",        ["0","7","7","61","0","68",  "351","84","435","221","0","656",  "-588","-89.63%","656"]),
    ("wh", "12", "AYARKKUNNAM",     ["25","3","28","0","0","28",  "285","172","457","20","0","477",  "-449","-94.13%","477"]),
    ("wh", "13", "THODUPUZHA",      ["0","1","1","0","0","1",  "29","21","50","4","0","54",  "-53","-98.15%","54"]),
    ("wh", "14", "TRIPUNITHURA",    ["10","15","25","40","0","65",  "154","67","221","171","0","392",  "-327","-83.42%","392"]),
    ("wh", "15", "KADAVANTHRA",     ["14","7","21","0","0","21",  "231","41","272","40","0","312",  "-291","-93.27%","312"]),
    ("wh", "16", "PERUMBAVOOR",     ["0","0","0","0","0","0",  "0","8","8","0","0","8",  "-8","-100%","8"]),
    ("wh", "17", "KOTHAMANGALAM",   ["36","9","45","0","0","45",  "0","124","124","10","0","134",  "-89","-66.42%","134"]),
    ("wh", "18", "ALUVA",           ["0","0","0","0","0","0",  "0","5","5","0","0","5",  "-5","-100%","5"]),
    ("wh", "19", "CHALAKUDY",       ["0","18","18","0","0","18",  "35","88","123","24","0","147",  "-129","-87.76%","147"]),
    ("wh", "20", "THRISSUR",        ["0","25","25","5","0","30",  "58","122","180","113","0","293",  "-263","-89.76%","293"]),
    ("cluster", "CLUSTER - 2",     ["85","85","170","106","0","276",  "1,143","732","1,875","603","0","2,478",  "-2,202","-89%","2,478"]),
    ("wh", "21", "PALAKKAD",        ["0","26","26","0","0","26",  "225","251","476","202","0","678",  "-652","-96.17%","678"]),
    ("wh", "22", "MENONPARA",       ["365","87","452","33","0","485",  "20","62","82","20","0","102",  "383","375.49%","102"]),
    ("wh", "23", "PERINTHALMANNA",  ["0","0","0","0","0","0",  "68","69","137","2","0","139",  "-139","-100%","139"]),
    ("wh", "24", "KOZHIKODE",       ["0","0","0","0","0","0",  "125","48","173","66","7","246",  "-246","-100%","246"]),
    ("wh", "25", "NADUVANNUR",      ["13","4","17","2","0","19",  "233","146","379","301","0","680",  "-661","-97.21%","680"]),
    ("wh", "26", "KALPETTA",        ["0","3","3","15","0","18",  "217","65","282","71","0","353",  "-335","-94.9%","353"]),
    ("wh", "27", "KANNUR",          ["0","1","1","2","0","3",  "0","66","66","97","0","163",  "-160","-98.16%","163"]),
    ("wh", "28", "BATTATHUR",       ["0","0","0","0","0","0",  "178","146","324","15","0","339",  "-339","-100%","339"]),
    ("cluster", "CLUSTER - 3",     ["378","121","499","52","0","551",  "1,066","853","1,919","774","7","2,700",  "-2,149","-80%","2,700"]),
    ("total",   "TOTAL",           ["564","330","894","203","0","1,097",  "3,750","4,281","8,031","2,003","12","10,046",  "-8,949","-89%","10,046"]),
    ("foot", "Day Sale",        "328",   "207",   "121",    "58%"),
    ("foot", "Industry Total",  "0",     "100",   "-100",   "-100%"),
]

# ── geometry ─────────────────────────────────────────────────────────────────
PW, PH   = 595.276, 841.890
TOPBAR_H = 3.0        # gold accent at the paper edge
TITLE_H  = 56.0       # masthead: wordmark lockup + as-on pill
SUB_H    = 0.0        # folded into the masthead
HDR_TOP  = 24.0
HDR_BOT  = 25.0
HDR_H    = HDR_TOP + HDR_BOT
FOOT_H   = 20.0
BOTTOM_M = 10.0
F_TITLE, F_SUB, F_DATA, F_FOOT = 17, 8.5, 10.5, 7.5
F_PILL   = 9.0
TRACK    = 1.5        # letterspacing on the wordmark
PAD      = 2.75
HAIR     = 0.4
GRP_RULE = 0.9        # navy group divider in the body
HDR_SEP  = 0.6        # intra-group separator in the header
HDR_EDGE = 1.8        # gold rules closing the header block
# row-height weights: a sub-total reads heavier than a warehouse line
WEIGHT = {"wh": 1.0, "cluster": 1.14, "total": 1.24, "foot": 1.06}

hdr_font = lambda fd: max(7.0, min(9.5, fd - 1.75))

def row_values(r):
    if r[0] in ("wh","cluster","total"):
        return r[3] if r[0]=="wh" else r[2]
    return None

def fit_columns(fd, fh, pad):
    need = []
    for i,(lines,_) in enumerate(COLS):
        w = 0.0
        for r in DATA:
            vals = row_values(r)
            if vals is None:
                if r[0]=="foot" and i==DIFF_COLS[0]: w = max(w, stringWidth(r[4],"Helvetica-Bold",fd))
                if r[0]=="foot" and i==DIFF_COLS[1]: w = max(w, stringWidth(r[5],"Helvetica-Bold",fd))
                continue
            if i==LABEL_COL:
                txt = r[2] if r[0]=="wh" else r[1]
                w = max(w, stringWidth(txt, "Helvetica-Bold" if r[0]!="wh" else "Helvetica", fd))
            else:
                w = max(w, stringWidth(vals[i-FIRST_VAL], "Helvetica-Bold" if r[0]!="wh" else "Helvetica", fd))
        for line in lines:
            for word in line.split(" "):
                w = max(w, stringWidth(word.replace(NB," "), "Helvetica-Bold", fh))
        need.append(w + (pad*4 if i==LABEL_COL else pad*2))

    lab = max(stringWidth(r[1],"Helvetica-Bold",fd) for r in DATA if r[0]=="foot")
    need[LABEL_COL] = max(need[LABEL_COL], lab + pad*2)

    # AUG and JUL carry the same measures and are read ACROSS — mirror each pair
    for a, pr in zip(CUR_COLS, PRIOR_COLS):
        need[a] = need[pr] = max(need[a], need[pr])

    foot = [r for r in DATA if r[0]=="foot"]
    want = max(stringWidth(CUR_BANNER,"Helvetica-Bold",fh),
               stringWidth(PRIOR_BANNER,"Helvetica-Bold",fh),
               max(stringWidth(r[2],"Helvetica-Bold",fd) for r in foot),
               max(stringWidth(r[3],"Helvetica-Bold",fd) for r in foot)) + pad*2
    have = sum(need[i] for i in CUR_COLS)
    if have < want:
        extra = (want-have)/len(CUR_COLS)
        for i in CUR_COLS + PRIOR_COLS: need[i] += extra

    have = sum(need[i] for i in DIFF_COLS)
    want = stringWidth(DIFF_BANNER,"Helvetica-Bold",fh) + pad*2
    if have < want:
        extra = (want-have)/len(DIFF_COLS)
        for i in DIFF_COLS: need[i] += extra
    return need

def solve():
    fd = F_DATA
    while fd > 5.5:
        fh = hdr_font(fd)
        need = fit_columns(fd, fh, PAD)
        if sum(need) <= PW:
            share = (PW - sum(need))/len(need)
            return fd, fh, [w+share for w in need]
        fd -= 0.25
    raise SystemExit("table will not fit")

F_DATA, F_HDR, COLW = solve()
XS = [0.0]
for w in COLW: XS.append(XS[-1]+w)

_avail  = PH - TOPBAR_H - TITLE_H - SUB_H - HDR_H - FOOT_H - BOTTOM_M
_units  = sum(WEIGHT[r[0]] for r in DATA)
UNIT_H  = _avail / _units
row_h   = lambda kind: UNIT_H * WEIGHT[kind]

def wrap(text, font, size, maxw):
    words, lines, cur = text.split(" "), [], ""
    for wd in words:
        trial = wd if not cur else cur+" "+wd
        if stringWidth(trial.replace(NB," "), font, size) <= maxw or not cur: cur = trial
        else: lines.append(cur); cur = wd
    if cur: lines.append(cur)
    return lines

def centred(c, text, font, size, colour, x0, x1, y_mid, dy=0.0):
    c.setFont(font, size); c.setFillColor(colour)
    t = text.replace(NB, " ")
    c.drawString((x0+x1)/2.0 - stringWidth(t,font,size)/2.0, y_mid - size*0.35 + dy, t)

def tracked_width(text, font, size, track):
    return stringWidth(text, font, size) + track * len(text)

def tracked(c, text, font, size, colour, x, y, track):
    # Letterspacing lives on the text object, not the canvas -- and Tc is part
    # of the PDF TEXT STATE, which SURVIVES the BT/ET block. Leaving it set
    # letterspaces every later drawString on the page (it silently widened the
    # whole table and broke centring the first time round), so reset it here.
    t = c.beginText(x, y)
    t.setFont(font, size); t.setFillColor(colour); t.setCharSpace(track)
    t.textOut(text)
    t.setCharSpace(0)
    c.drawText(t)

def pill(c, text, font, size, x_right, y_mid, bg, ink, padx=9.0, h=18.0):
    w = stringWidth(text, font, size) + padx * 2
    x = x_right - w
    c.setFillColor(bg); c.setStrokeColor(bg); c.setLineWidth(0.6)
    c.roundRect(x, y_mid - h/2, w, h, h/2, stroke=1, fill=1)
    c.setFont(font, size); c.setFillColor(ink)
    c.drawString(x + padx, y_mid - size*0.35, text)
    return x

def fill(c, x0, x1, y, h, colour):
    c.setFillColor(colour); c.rect(x0, y, x1-x0, h, stroke=0, fill=1)

def hline(c, x0, x1, y, colour, lw):
    c.setStrokeColor(colour); c.setLineWidth(lw); c.setLineCap(0); c.line(x0, y, x1, y)

def vline(c, x, y0, y1, colour, lw):
    c.setStrokeColor(colour); c.setLineWidth(lw); c.setLineCap(0); c.line(x, y0, x, y1)

def diff_colour(v, on_navy):
    if on_navy: return WHITE
    return RED if v.strip().startswith("-") else GREEN

def is_nil(v):
    return v.strip() in ("0", "0%", "-", "")

def build(path):
    c = canvas.Canvas(path, pagesize=(PW, PH))
    c.setTitle(f"{SUBLEFT} {SUBRIGHT}")

    # ── masthead ────────────────────────────────────────────────────────────
    # One band: wordmark lockup on the left, as-on pill on the right.
    # Replaces the old stacked title band + strapline band, and costs no height.
    fill(c, 0, PW, PH - TOPBAR_H, TOPBAR_H, GOLD)          # accent at the paper edge
    y = PH - TOPBAR_H - TITLE_H
    fill(c, 0, PW, y, TITLE_H, NAVY)

    # a gold bar anchors the lockup where a logo would otherwise sit
    BAR_W, BAR_H = 3.2, 30.0
    fill(c, 16.0, 16.0 + BAR_W, y + (TITLE_H - BAR_H)/2, BAR_H, GOLD)
    tx = 16.0 + BAR_W + 12.0

    tracked(c, TITLE, "Helvetica-Bold", F_TITLE, GOLD, tx, y + 30.0, TRACK)
    tracked(c, SUBLEFT, "Helvetica-Bold", F_SUB, STEEL, tx, y + 15.0, 0.9)

    pill(c, SUBRIGHT.upper(), "Helvetica-Bold", F_PILL, PW - 16.0, y + 32.0, GOLD, NAVY)
    c.setFont("Helvetica", 7.0); c.setFillColor(STEEL)
    c.drawRightString(PW - 16.0, y + 12.0, "ALL FIGURES IN CASES")

    # ── header block ────────────────────────────────────────────────────────
    top_y = y - HDR_TOP
    bot_y = top_y - HDR_BOT
    fill(c, 0, PW, bot_y, HDR_H, NAVY)

    for i,(lines,group) in enumerate(COLS):
        if group is None:
            flat = []
            for ln in lines: flat += wrap(ln, "Helvetica-Bold", F_HDR, COLW[i]-PAD*2)
            start = bot_y + HDR_H/2 + (len(flat)-1)*(F_HDR+1.2)/2.0
            for j,ln in enumerate(flat):
                centred(c, ln, "Helvetica-Bold", F_HDR, GOLD, XS[i], XS[i+1], start - j*(F_HDR+1.2))
        else:
            centred(c, lines[0], "Helvetica-Bold", F_HDR, GOLD, XS[i], XS[i+1], bot_y + HDR_BOT/2)

    # The live period is a gold chip; the prior period and the difference sit
    # back on a lighter navy. Chips are inset so navy gutters separate them.
    IN = 1.4
    for banner, idxs, bg, ink in ((CUR_BANNER,   CUR_COLS,   GOLD,      NAVY),
                                  (PRIOR_BANNER, PRIOR_COLS, NAVY_CHIP, GOLD),
                                  (DIFF_BANNER,  DIFF_COLS,  NAVY_CHIP, GOLD)):
        x0, x1 = XS[idxs[0]] + IN, XS[idxs[-1]+1] - IN
        c.setFillColor(bg); c.setStrokeColor(bg); c.setLineWidth(0.5)
        c.roundRect(x0, top_y + IN, x1-x0, HDR_TOP - IN*2, 2.4, stroke=1, fill=1)
        centred(c, banner, "Helvetica-Bold", F_HDR, ink, x0, x1, top_y + HDR_TOP/2)

    # dim separators inside a group, gold only where a group ends
    for i in range(1, len(COLS)):
        if i in GROUP_EDGES: continue
        vline(c, XS[i], bot_y, bot_y+HDR_BOT, GOLD_DIM, HDR_SEP)
    for i in GROUP_EDGES:
        vline(c, XS[i], bot_y, top_y, GOLD, 1.3)
        vline(c, XS[i], top_y, top_y + HDR_TOP, NAVY, 1.3)
    hline(c, 0, PW, top_y + HDR_TOP - HDR_EDGE/2, GOLD, HDR_EDGE)
    hline(c, 0, PW, bot_y + HDR_EDGE/2, GOLD, HDR_EDGE)

    # ── body ────────────────────────────────────────────────────────────────
    ry = bot_y
    body_top = bot_y
    grid_bot = None
    for idx, r in enumerate(DATA):
        kind = r[0]
        h = row_h(kind); ry -= h; mid = ry + h/2
        if kind in ("total", "foot") and grid_bot is None:
            grid_bot = ry + h   # merged / summary rows carry no internal grid

        if kind == "foot":
            _, label, curv, priorv, cases, pct = r
            fill(c, 0, PW, ry, h, FOOT_BG)
            centred(c, label, "Helvetica-Bold", F_DATA, NAVY, XS[LABEL_COL], XS[LABEL_COL+1], mid)
            centred(c, curv, "Helvetica-Bold", F_DATA, NAVY, XS[CUR_COLS[0]], XS[CUR_COLS[-1]+1], mid)
            centred(c, priorv, "Helvetica-Bold", F_DATA, NAVY, XS[PRIOR_COLS[0]], XS[PRIOR_COLS[-1]+1], mid)
            for i, v in zip(DIFF_COLS, (cases, pct)):
                centred(c, v, "Helvetica-Bold", F_DATA, diff_colour(v, False), XS[i], XS[i+1], mid)
            hline(c, 0, PW, ry+h, RULE, HAIR)
            continue

        zeb = (idx % 2 == 1)
        if kind == "cluster":
            fill(c, 0, PW, ry, h, CLUSTER_BG)
            hline(c, 0, PW, ry+h, CLUSTER_RL, 0.8)
            hline(c, 0, PW, ry,   CLUSTER_RL, 0.8)
        elif kind == "total":
            fill(c, 0, PW, ry, h, NAVY)
        else:
            fill(c, 0, XS[PRIOR_COLS[0]], ry, h, ZEBRA if zeb else WHITE)
            fill(c, XS[PRIOR_COLS[0]], XS[PRIOR_COLS[-1]+1], ry, h, CREAM_Z if zeb else CREAM)
            fill(c, XS[DIFF_COLS[0]], XS[LAST_COL], ry, h, ZEBRA if zeb else WHITE)
            fill(c, XS[LAST_COL], PW, ry, h, REF_Z if zeb else REF_BG)
            hline(c, 0, PW, ry, RULE, HAIR)

        vals = r[3] if kind=="wh" else r[2]
        if kind == "wh":
            c.setFont("Helvetica", F_DATA); c.setFillColor(INK)
            c.drawString(XS[LABEL_COL]+PAD*2, mid - F_DATA*0.35, r[2])
        else:
            centred(c, r[1], "Helvetica-Bold", F_DATA,
                    WHITE if kind=="total" else NAVY,
                    XS[LABEL_COL], XS[LABEL_COL+1], mid)

        for i in range(FIRST_VAL, len(COLS)):
            v = vals[i-FIRST_VAL]
            if kind == "total":
                col = WHITE if i not in DIFF_COLS else WHITE
                f = "Helvetica-Bold"
            elif kind == "cluster":
                col = diff_colour(v, False) if i in DIFF_COLS else NAVY
                f = "Helvetica-Bold"
            elif i in (CUR_TOTAL, PRIOR_TOTAL):
                col, f = NAVY, "Helvetica-Bold"
            elif i in DIFF_COLS:
                col, f = diff_colour(v, False), "Helvetica-Bold"
            elif i == LAST_COL:
                col, f = (ZERO if is_nil(v) else INK_SOFT), "Helvetica"
            else:
                col, f = (ZERO if is_nil(v) else INK), "Helvetica"
            centred(c, v, f, F_DATA, col, XS[i], XS[i+1], mid)

    body_bot = ry
    # intra-group column separators, drawn once over the whole body
    for i in range(1, len(COLS)):
        if i in GROUP_EDGES: continue
        vline(c, XS[i], grid_bot if grid_bot is not None else body_bot, body_top, RULE, HAIR)
    for i in GROUP_EDGES:
        vline(c, XS[i], body_bot, body_top, NAVY_SOFT, GRP_RULE)
    hline(c, 0, PW, body_bot, NAVY, 1.2)

    centred(c, FOOTER, "Helvetica", F_FOOT, GREY, 0, PW, body_bot - FOOT_H/2)
    c.showPage(); c.save()
    return path

if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv)>1 else "v2.pdf"
    build(out)
    print(f"data {F_DATA:.2f}pt | header {F_HDR:.2f}pt | unit row {UNIT_H:.1f}pt | table {sum(COLW):.1f}/{PW:.0f}pt")
