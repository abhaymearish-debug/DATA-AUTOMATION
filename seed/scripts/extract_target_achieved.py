#!/usr/bin/env python3
"""Parse a TARGET V/S ACHIEVED cluster PDF into structured JSON.

Each bond is one block of two rows (TGT / ACH) with the bond name and the
ACH % merged across both. The cluster total is the same shape, styled gold.
"""

import json
import sys
from collections import defaultdict

import pdfplumber


def norm(c):
    if c is None:
        return None
    if isinstance(c, (int, float)):
        return (round(c, 2),) * 3
    return tuple(round(x, 2) for x in c)


NAVY = (0.04, 0.16, 0.31)
GOLD = (1.0, 0.74, 0.19)


def parse(path):
    p = pdfplumber.open(path).pages[0]

    rows = defaultdict(list)
    for r in p.rects:
        rows[round(r["top"], 1)].append(r)

    # the header row is the widest full band of cells; every later 12-cell row
    # opens a bond block and the 10-cell row under it is its ACH line
    full = sorted(t for t, v in rows.items() if len(v) == 12)
    xs = [round(c["x0"], 1) for c in sorted(rows[full[0]], key=lambda r: r["x0"])]
    xs.append(round(max(c["x1"] for c in rows[full[0]]), 1))

    def lines_in(x0, x1, top, bot):
        """Text in a cell, kept as separate LINES. A wrapped header must not be
        flattened -- "BCB NO.1 / CLASSIC" concatenates to "BCB NO.1CLASSIC"."""
        by_line = defaultdict(list)
        for ch in p.chars:
            if x0 - 0.6 <= ch["x0"] < x1 - 0.6 and top - 1 <= ch["top"] <= bot - 2:
                by_line[round(ch["top"], 0)].append(ch)
        out = []
        for k in sorted(by_line):
            out.append("".join(c["text"] for c in sorted(by_line[k], key=lambda c: c["x0"])).strip())
        return [o for o in out if o]

    def text_in(x0, x1, top, bot):
        return " ".join(lines_in(x0, x1, top, bot))

    head = {}
    for ch in p.chars:
        s = round(ch["size"], 1)
        head.setdefault(s, [])
        head[s].append(ch)

    def line(size, left=True, tol=1.0):
        """The band carries a left caption and a right one. Split on the actual
        GAP between the two runs, not on the page midpoint — a long left caption
        crosses the midpoint and loses its last letter to the right-hand side."""
        cs = sorted([c for sz, v in head.items() if abs(sz - size) <= tol for c in v],
                    key=lambda c: c["x0"])
        if not cs:
            return ""
        split = len(cs)
        best = 0.0
        for i in range(1, len(cs)):
            gap = cs[i]["x0"] - cs[i - 1]["x1"]
            if gap > best:
                best, split = gap, i
        sel = cs[:split] if left else cs[split:]
        return "".join(c["text"] for c in sel).strip()

    title = "".join(c["text"] for c in sorted(head.get(18.0, []), key=lambda c: c["x0"])).strip()
    band_l = line(13.0, True)
    band_r = line(13.0, False)

    hdr_top = full[0]
    hdr_bot = rows[hdr_top][0]["bottom"]
    headers = [lines_in(xs[i], xs[i + 1], hdr_top, hdr_bot) for i in range(len(xs) - 1)]

    blocks = []
    for t in full[1:]:
        cs = sorted(rows[t], key=lambda r: r["x0"])
        bot = cs[0]["bottom"]                       # the merged label spans both rows
        tgt_bot = t + (bot - t) / 2
        ach_top = tgt_bot
        name = lines_in(xs[0], xs[1], t, bot)
        pct = text_in(xs[-2], xs[-1], t, bot)
        fill = norm(cs[1].get("non_stroking_color"))
        # Three row styles: bond (white), cluster total (gold), grand total
        # (navy). Match the NEAREST band colour rather than testing equality —
        # the TGT row carries a slight shade of its band, so an exact test
        # misreads a shaded gold row as an ordinary white one.
        bands = {"bond": (1.0, 1.0, 1.0), "total": GOLD, "grand": NAVY}
        kind = min(bands, key=lambda k: sum((a - b) ** 2
                                            for a, b in zip(fill, bands[k])))
        tgt = [text_in(xs[i], xs[i + 1], t, tgt_bot) for i in range(1, len(xs) - 2)]
        ach = [text_in(xs[i], xs[i + 1], ach_top, bot) for i in range(1, len(xs) - 2)]
        blocks.append({
            "name": name, "pct": pct, "tgt": tgt, "ach": ach,
            "total": kind == "total", "kind": kind,
        })

    return {"title": title, "band_left": band_l, "band_right": band_r,
            "headers": headers, "blocks": blocks,
            "size": [round(p.width, 1), round(p.height, 1)]}


if __name__ == "__main__":
    out = {}
    for path in sys.argv[1:-1]:
        out[path.split("/")[-1]] = parse(path)
    json.dump(out, open(sys.argv[-1], "w"), indent=1)
    for k, v in out.items():
        from collections import Counter
        kinds = Counter(b["kind"] for b in v["blocks"])
        print(f"{k:<42} {dict(kinds)}   {len(v['headers'])} cols")
    d = list(out.values())[0]
    print("\nheaders:", d["headers"])
    print("first block:", json.dumps(d["blocks"][0]))
    print("last block :", json.dumps(d["blocks"][-1]))
