#!/usr/bin/env python3
"""Parse SHOP SALES DAILY (bond × day matrix) into structured JSON."""

import json
import sys
from collections import defaultdict

import pdfplumber

NAVY = (0.04, 0.16, 0.31)
ZEBRA = (0.96, 0.97, 0.99)
DIM = (0.78, 0.8, 0.84)


def norm(c):
    if c is None:
        return None
    if isinstance(c, (int, float)):
        return (round(c, 2),) * 3
    return tuple(round(x, 2) for x in c)


def parse(path):
    p = pdfplumber.open(path).pages[0]

    rows = defaultdict(list)
    for r in p.rects:
        rows[round(r["top"], 1)].append(r)
    # The header is TWO rect rows: a full-width row carrying BOND / weekday /
    # TOTAL, then a narrower row holding just the day numbers under the day
    # columns. Only full-width rows are data — treating the day-number row as
    # data reads "1 2 3 ..." as a bond.
    wide = max(len(v) for v in rows.values())
    full = sorted(t for t, v in rows.items() if len(v) == wide)

    xs = [round(c["x0"], 1) for c in sorted(rows[full[0]], key=lambda r: r["x0"])]
    xs.append(round(max(c["x1"] for c in rows[full[0]]), 1))
    ncol = len(xs) - 1

    def cell(x0, x1, top, bot):
        cs = [ch for ch in p.chars
              if x0 - 0.6 <= ch["x0"] < x1 - 0.6 and top - 1 <= ch["top"] <= bot - 2]
        cs.sort(key=lambda c: (round(c["top"]), c["x0"]))
        return "".join(c["text"] for c in cs).strip(), cs

    by_size = defaultdict(list)
    for ch in p.chars:
        by_size[round(ch["size"], 1)].append(ch)

    def band(size, left=True):
        cs = sorted(by_size.get(size, []), key=lambda c: c["x0"])
        if not cs:
            return ""
        split, best = len(cs), 0.0
        for i in range(1, len(cs)):
            gap = cs[i]["x0"] - cs[i - 1]["x1"]
            if gap > best:
                best, split = gap, i
        sel = cs[:split] if left else cs[split:]
        return "".join(c["text"] for c in sel).strip()

    title = "".join(c["text"] for c in sorted(by_size.get(18.0, []), key=lambda c: c["x0"])).strip()

    # header block: weekday row on top, day number under it; BOND and TOTAL
    # span both, so read them across the whole block
    hdr_top = full[0]
    hdr_bot = rows[hdr_top][0]["bottom"]
    # the day-number strip sits directly under the weekday row
    sub = [t for t in rows if hdr_top < t < full[1] and len(rows[t]) < wide]
    if sub:
        hdr_bot = max(rows[t][0]["bottom"] for t in sub)
    mid = full[0] + (rows[full[0]][0]["bottom"] - full[0]) / 2
    weekday, daynum, spans = [], [], []
    for i in range(ncol):
        wk, _ = cell(xs[i], xs[i + 1], hdr_top, mid)
        dn, _ = cell(xs[i], xs[i + 1], mid, hdr_bot)
        whole, _ = cell(xs[i], xs[i + 1], hdr_top, hdr_bot)
        spans.append(whole)
        weekday.append(wk)
        daynum.append(dn)

    body = []
    for t in full[1:]:
        cs = sorted(rows[t], key=lambda r: r["x0"])
        bot = cs[0]["bottom"]
        fill = norm(cs[0].get("non_stroking_color"))
        name, _ = cell(xs[0], xs[1], t, bot)
        vals, dim = [], []
        for i in range(1, ncol):
            v, chars = cell(xs[i], xs[i + 1], t, bot)
            vals.append(v)
            dim.append(bool(chars) and norm(chars[0].get("non_stroking_color")) == DIM)
        body.append({"name": name, "values": vals, "dim": dim,
                     "total": fill == NAVY, "zebra": fill == ZEBRA})

    return {"title": title, "band_left": band(13.0, True), "band_right": band(13.0, False),
            "weekday": weekday, "daynum": daynum, "spans": spans, "rows": body,
            "size": [round(p.width, 1), round(p.height, 1)]}


if __name__ == "__main__":
    d = parse(sys.argv[1])
    json.dump(d, open(sys.argv[2], "w"), indent=1)
    print(f"{len(d['rows'])} rows, {len(d['weekday'])} columns")
    print("band :", d["band_left"], "|", d["band_right"])
    print("wkday:", d["weekday"])
    print("daynum:", d["daynum"])
    print("spans:", d["spans"][0], "...", d["spans"][-1])
    print("first:", json.dumps(d["rows"][0]))
    print("last :", json.dumps(d["rows"][-1]))
