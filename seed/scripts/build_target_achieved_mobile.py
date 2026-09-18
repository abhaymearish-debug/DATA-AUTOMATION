#!/usr/bin/env python3
"""
TARGET V/S ACHIEVED — mobile-friendly rebuild.

Same report, same columns, same TGT/ACH block structure. Landscape -> portrait A4
so it fills a phone screen, with the type sized up into the space that frees.

Format is unchanged. The only additions are presentational: a gold grid on the
header, alternating tint on the bond blocks, and a gold rule closing the cluster
total.

Input: target.json (from extract_target.py)
"""

import json
import sys
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color
from reportlab.pdfbase.pdfmetrics import stringWidth

# ── palette (sampled from the source PDFs) ───────────────────────────────────
NAVY  = Color(0.043, 0.161, 0.310)
GOLD  = Color(1.000, 0.741, 0.192)
AMBER = Color(1.000, 0.741, 0.192)   # cluster-total band
ZEBRA = Color(0.960, 0.970, 0.990)   # alternating bond block
WHITE = Color(1, 1, 1)
BLACK = Color(0, 0, 0)
RED   = Color(0.812, 0.075, 0.133)   # ACH % on a light row
RED_LT= Color(1.000, 0.392, 0.392)   # ACH % on the navy grand-total band
RULE  = Color(0.780, 0.780, 0.780)   # cell hairline


def shade(col, f):
    """Tint a band colour for the TGT row. Darkens a light band and lightens a
    dark one, so the same rule works on white, gold and navy alike."""
    lum = 0.299 * col.red + 0.587 * col.green + 0.114 * col.blue
    if lum > 0.5:                       # light band -> darken
        return Color(col.red * f, col.green * f, col.blue * f)
    return Color(col.red + (1 - col.red) * (1 - f),      # dark band -> lighten
                 col.green + (1 - col.green) * (1 - f),
                 col.blue + (1 - col.blue) * (1 - f))


TGT_SHADE = 0.945                    # how far the TGT row departs from its band

PW, PH = 595.276, 841.890

TITLE_H = 42.0
BAND_H  = 28.0
FOOT_H  = 0.0        # this report has no page footer — reserving space
                     # for one just left dead white below the total band

F_TITLE = 18
F_BAND  = 13
F_HDR   = 9.0        # header labels; smaller than the data on purpose —
                     # they are short known labels, and the width they free
                     # goes into padding and into the data font
F_DATA  = 10.5
PAD     = 3.5        # per side inside a column. Trimmed from 4.5 to buy
                     # header font — with equal brand columns the padding is
                     # multiplied by 8, so 1pt here is worth ~0.5pt of type
ROW_MAX = 30.0       # comfortable row; the page is then trimmed to fit
LAB_PAD = 3.0        # label-column indent; trimmed to buy width for
                     # the equal brand columns
HAIR    = 0.4
HDR_RULE      = 1.4  # gold grid inside the header
HDR_RULE_EDGE = 2.0  # gold rules closing the header block
SEP_RULE      = 1.5  # navy rule segregating each bond / cluster block
HDR_LEAD      = 4.5  # leading between wrapped header lines
HDR_VPAD      = 18.0 # air above and below the header text


def fit(doc, fh, fd, pad):
    """Column widths measured from the content and the wrapped header lines."""
    H = doc["headers"]
    blocks = doc["blocks"]
    need = []

    # STAFF - BOND, left aligned
    w = max(stringWidth(l, "Helvetica-Bold", fd)
            for b in blocks for l in b["name"])
    w = max(w, stringWidth(header_text(H[0]), "Helvetica-Bold", fh))
    need.append(w + LAB_PAD * 3)

    # CAT + the eight brands + GRAND TOTAL  (indices 1..10 of the value lists)
    for i in range(1, 11):
        w = max(stringWidth(b[k][i - 1], "Helvetica-Bold", fd)
                for b in blocks for k in ("tgt", "ach"))
        w = max(w, max(stringWidth(word, "Helvetica-Bold", fh)
                       for word in header_text(H[i]).split(" ")))
        need.append(w + pad * 2)

    # ACH %
    w = max(stringWidth(b["pct"], "Helvetica-Bold", fd) for b in blocks)
    w = max(w, stringWidth(header_text(H[11]), "Helvetica-Bold", fh))
    need.append(w + pad * 2)
    return need


BRAND_COLS = range(2, 10)            # the eight brand columns

# Short brand names for the column headers. The source labels carry the full
# registered names ("OLD PEARL NO.1 MATURED XXX"), most of which is boilerplate
# that forced 4-line wraps and squeezed the header font down to 6.25pt.
SHORT = {
    "BCB NO.1 CLASSIC":             "BCB",
    "BLENDERS CHOICE NO.1":         "BLENDERS CHOICE",
    "CHAIRMANS CHOICE XO":          "CCB",   # field shorthand — "CHAIRMAN'S"
                                            # is a 10-char unbreakable word and
                                            # with equal columns it caps the
                                            # header font for ALL eight brands
    "K.S 99 LIFE TIME MATURED XXX": "KS.99",
    "MAGIC BLEND RESERVED XXX":     "MAGIC BLEND",
    "MORNING WALKERS XO":           "MORNING WALKERS",
    "OLD PEARL NO.1 MATURED XXX":   "OLD PEARL",
    "ROYAL OLD FORT NO.1 XXX":      "ROYAL OLD FORT",
}


SHORT["STAFF - BOND"] = "BOND"


def band_title(text):
    """House wording for the report name. "TARGET vs ACHIEVEMENT" is the term
    used on the target workbooks, so the whole set follows it."""
    return (text.replace("TARGET V/S ACHIEVED", "TARGET VS ACHIEVEMENT")
                .replace("CLUSTERS SUMMARY", "CLUSTER SUMMARY")
                .replace("CLUSTER - ", "CLUSTER "))


def header_text(lines):
    """Full header label -> short brand name, where one is defined."""
    full = " ".join(lines)
    return SHORT.get(full, full)


def allocate(doc, fh, fd, pad, pool=None):
    """All eight BRAND columns share one width — they hold the same kind of
    measure, so an even rhythm reads far better than sizing each to its own
    label. The four other columns (bond, CAT, grand total, ACH %) keep their
    measured widths, and whatever is left over is split eight ways.

    Sizing each brand column to its own longest WORD got it backwards:
    "CHAIRMANS" is a long word in a short label, so that column came out widest
    and wrapped to 2 lines, while "ROYAL OLD FORT NO.1 XXX" — more text but all
    short words — was squeezed narrowest and wrapped to 5.
    """
    # Take each fixed column at its widest across the WHOLE set, so all four
    # sheets come out with identical geometry — otherwise the brand columns land
    # at 47.3 on one sheet and 52.0 on another and the set looks mismatched.
    docs_pool = pool or [doc]
    per = [fit(d, fh, fd, pad) for d in docs_pool]
    floors = [max(f[i] for f in per) for i in range(len(per[0]))]
    fixed = [i for i in range(len(floors)) if i not in BRAND_COLS]
    brand_w = (PW - sum(floors[i] for i in fixed)) / len(BRAND_COLS)
    if brand_w < max(floors[i] for i in BRAND_COLS):
        return None                  # equal columns will not fit at this size
    out = list(floors)
    for i in BRAND_COLS:
        out[i] = brand_w
    return out


def solve(doc, docs=None):
    """Size the type against EVERY sheet in the set, not just this one, so the
    four files come out at the same font instead of 9.75 / 10.5 / 9.75 / 10.5."""
    pool = docs if docs else [doc]
    fd = F_DATA
    while fd > 8:
        # search the header size independently — with equal brand columns the
        # binding constraint is the single longest header WORD, so the header
        # font has to come down further than a fixed offset from the data font
        fh = min(F_HDR, fd - 1.0)
        while fh >= 5.5:
            got = allocate(doc, fh, fd, PAD, pool)
            if got is not None:
                return fd, fh, got
            fh -= 0.25
        fd -= 0.25
    raise SystemExit("table will not fit")


def wrap(text, font, size, maxw):
    """Greedy word wrap. The width solver only guarantees the widest WORD fits,
    so the label must be re-wrapped to whatever width the column ended up with —
    drawing the source's own line breaks overflows the narrower portrait column."""
    words, lines, cur = text.split(" "), [], ""
    for w in words:
        trial = w if not cur else cur + " " + w
        if stringWidth(trial, font, size) <= maxw or not cur:
            cur = trial
        else:
            lines.append(cur); cur = w
    if cur:
        lines.append(cur)
    return lines


def centred(c, text, font, size, colour, x0, x1, ymid, dy=0.0):
    c.setFont(font, size); c.setFillColor(colour)
    c.drawString((x0 + x1) / 2.0 - stringWidth(text, font, size) / 2.0,
                 ymid - size * 0.35 + dy, text)


def cell(c, x0, x1, y, h, fill=None, edge=RULE, lw=HAIR):
    c.setStrokeColor(edge); c.setLineWidth(lw)
    if fill is not None:
        c.setFillColor(fill)
    c.rect(x0, y, x1 - x0, h, stroke=1, fill=1 if fill is not None else 0)


def build(doc, path, docs=None):
    F_D, F_H, COLW = solve(doc, docs)
    xs = [0.0]
    for w in COLW:
        xs.append(xs[-1] + w)

    blocks = doc["blocks"]
    # re-wrap every header to the width its column actually ended up with
    hdr = [wrap(header_text(l), "Helvetica-Bold", F_H, COLW[i] - PAD * 2)
           for i, l in enumerate(doc["headers"])]
    hdr_lines = max(len(l) for l in hdr)
    HDR_H = hdr_lines * (F_H + HDR_LEAD) + HDR_VPAD

    # These tables are short — 4 to 7 bond blocks. On A4 that leaves ~30% of the
    # page empty, which on a phone is just dead scroll. So the rows take a
    # comfortable fixed height and the PAGE is trimmed to the content instead.
    # Width is unchanged at A4 width, so it still prints sensibly.
    ROW_H = ROW_MAX
    BLOCK_H = ROW_H * 2
    page_h = TITLE_H + BAND_H + HDR_H + BLOCK_H * len(blocks) + FOOT_H
    if page_h > PH:                       # never taller than A4
        ROW_H = (PH - TITLE_H - BAND_H - HDR_H - FOOT_H) / (len(blocks) * 2)
        BLOCK_H = ROW_H * 2
        page_h = PH

    c = canvas.Canvas(path, pagesize=(PW, page_h))
    c.setTitle(band_title(doc["band_left"]))

    # ── masthead ─────────────────────────────────────────────────────────────
    y = page_h - TITLE_H
    c.setFillColor(NAVY); c.rect(0, y, PW, TITLE_H, stroke=0, fill=1)
    centred(c, doc["title"], "Helvetica-Bold", F_TITLE, GOLD, 0, PW, y + TITLE_H / 2)

    # ── gold band: report name + as-on ───────────────────────────────────────
    y -= BAND_H
    c.setFillColor(GOLD); c.rect(0, y, PW, BAND_H, stroke=0, fill=1)
    c.setFont("Helvetica-Bold", F_BAND); c.setFillColor(NAVY)
    left = band_title(doc["band_left"])
    c.drawString(10, y + BAND_H / 2 - F_BAND * 0.35, left)
    gap = PW - 20 - stringWidth(left, "Helvetica-Bold", F_BAND) \
                  - stringWidth(doc["band_right"], "Helvetica-Bold", F_BAND)
    assert gap > 12, f"band captions collide (gap {gap:.1f}pt) — reduce F_BAND"
    c.drawRightString(PW - 10, y + BAND_H / 2 - F_BAND * 0.35, doc["band_right"])

    # ── header ───────────────────────────────────────────────────────────────
    y -= HDR_H
    for i, lines in enumerate(hdr):
        cell(c, xs[i], xs[i + 1], y, HDR_H, NAVY, GOLD, HDR_RULE)
        start = y + HDR_H / 2 + (len(lines) - 1) * (F_H + HDR_LEAD) / 2.0
        for j, ln in enumerate(lines):
            centred(c, ln, "Helvetica-Bold", F_H, GOLD,
                    xs[i], xs[i + 1], start - j * (F_H + HDR_LEAD))
    c.setStrokeColor(GOLD); c.setLineWidth(HDR_RULE_EDGE); c.setLineCap(0)
    c.line(0, y + HDR_H - HDR_RULE_EDGE / 2, PW, y + HDR_H - HDR_RULE_EDGE / 2)
    c.line(0, y + HDR_RULE_EDGE / 2, PW, y + HDR_RULE_EDGE / 2)

    # ── bond blocks ──────────────────────────────────────────────────────────
    # The summary sheet has the same SHAPE as a cluster sheet one level up: its
    # cluster rows are the detail rows and GRAND TOTAL is the total. Remap so it
    # uses the identical palette — white detail rows, one gold total band —
    # instead of a page of gold with a navy band at the bottom.
    summary = any(x.get("kind") == "grand" for x in blocks)
    remap = {"total": "bond", "grand": "total"} if summary else {}

    for b in blocks:
        top = y
        y -= BLOCK_H
        kind = b.get("kind", "total" if b["total"] else "bond")
        kind = remap.get(kind, kind)
        is_tot = kind != "bond"

        body = {"total": AMBER, "grand": NAVY}.get(kind, WHITE)
        # ink follows the band: black on white and on gold, white on navy
        ink = WHITE if kind == "grand" else BLACK
        pct_col = {"bond": RED, "total": BLACK, "grand": RED_LT}[kind]

        # merged label cell, spanning both rows
        cell(c, xs[0], xs[1], y, BLOCK_H, NAVY)
        # On the summary the cluster rows are DETAIL rows, so "CLUSTER - 1 TOTAL"
        # reads wrong — they are just CLUSTER - 1. GRAND TOTAL keeps its wording.
        label = " ".join(b["name"]).replace("CLUSTER - ", "CLUSTER ")
        if kind == "bond" and label.upper().endswith(" TOTAL"):
            label = label[:-6].rstrip()
        # wrap to the column rather than reusing the source's own line breaks,
        # which split awkwardly as "CLUSTER -" / "1 TOTAL"
        lines = wrap(label, "Helvetica-Bold", F_D, COLW[0] - LAB_PAD * 2)
        start = y + BLOCK_H / 2 + (len(lines) - 1) * (F_D + 1.5) / 2.0
        for j, ln in enumerate(lines):
            centred(c, ln, "Helvetica-Bold", F_D, GOLD if is_tot else WHITE,
                    xs[0], xs[1], start - j * (F_D + 1.5))

        # TGT row then ACH row
        for k, key in enumerate(("tgt", "ach")):
            ry = top - ROW_H * (k + 1)
            vals = b[key]
            row_bg = shade(body, TGT_SHADE) if key == "tgt" else body
            for i, v in enumerate(vals):
                col = i + 1
                cell(c, xs[col], xs[col + 1], ry, ROW_H, row_bg)
                last = i == len(vals) - 1              # the GRAND TOTAL column
                bold = is_tot or i == 0 or last
                col_ink = GOLD if (kind == "grand" and last) else ink
                centred(c, v, "Helvetica-Bold" if bold else "Helvetica", F_D,
                        col_ink, xs[col], xs[col + 1], ry + ROW_H / 2)

        # merged ACH % cell
        cell(c, xs[11], xs[12], y, BLOCK_H, body)
        centred(c, b["pct"], "Helvetica-Bold", F_D, pct_col, xs[11], xs[12], y + BLOCK_H / 2)

        # Segregate every block with a full-width rule. It must CONTRAST with
        # the band it sits on — gold rules on the gold cluster-total bands were
        # invisible, which is what made the summary read as one solid mass.
        c.setStrokeColor(NAVY); c.setLineWidth(SEP_RULE)
        c.line(0, y + SEP_RULE / 2, PW, y + SEP_RULE / 2)

    c.showPage(); c.save()
    return F_D, F_H, ROW_H, sum(COLW), page_h


if __name__ == "__main__":
    docs = json.load(open(sys.argv[1]))
    outdir = sys.argv[2] if len(sys.argv) > 2 else "."
    names = {
        "achieved_target_cluster_-_1-4.pdf":      "TARGET vs ACHIEVED - CLUSTER 1 - MOBILE.pdf",
        "achieved_target_cluster_-_2-4.pdf":      "TARGET vs ACHIEVED - CLUSTER 2 - MOBILE.pdf",
        "achieved_target_cluster_-_3-4.pdf":      "TARGET vs ACHIEVED - CLUSTER 3 - MOBILE.pdf",
        "achieved_target_clusters_summary-4.pdf": "TARGET vs ACHIEVED - SUMMARY - MOBILE.pdf",
    }
    for src, doc in docs.items():
        out = f"{outdir}/{names.get(src, src)}"
        fd, fh, rh, tw, ph = build(doc, out, list(docs.values()))
        print(f"{names.get(src,src):<48} data {fd:.2f}pt  hdr {fh:.2f}pt  "
              f"row {rh:.1f}pt  page {tw:.0f} x {ph:.0f}pt")
