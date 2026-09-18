#!/usr/bin/env python3
"""Parse the shop-sales cumulative PDF into structured JSON (one entry per shop page)."""

import json
import sys
import pdfplumber

NAVY  = (0.04, 0.16, 0.31)
ZEBRA = (0.96, 0.97, 0.99)
WHITE = (1.0, 1.0, 1.0)
GOLD  = (1.0, 0.74, 0.19)


def norm(c):
    if c is None:
        return None
    if isinstance(c, (int, float)):
        return (round(c, 2),) * 3
    return tuple(round(x, 2) for x in c)


def parse(path):
    doc = pdfplumber.open(path)
    pages = []
    for p in doc.pages:
        # column boundaries come from the widest set of same-top cells
        by_top = {}
        for r in p.rects:
            by_top.setdefault(round(r["top"], 1), []).append(r)

        cells = {t: sorted(v, key=lambda r: r["x0"]) for t, v in by_top.items() if len(v) == 5}
        xs = None
        for t in sorted(cells):
            xs = [round(c["x0"], 1) for c in cells[t]] + [round(cells[t][-1]["x1"], 1)]
            break

        chars_by_top = {}
        for ch in p.chars:
            chars_by_top.setdefault(round(ch["top"], 1), []).append(ch)

        def text_in(row_top, row_bot, x0, x1):
            out = []
            for t, cs in chars_by_top.items():
                if row_top - 1 <= t <= row_bot - 4:
                    for ch in cs:
                        if x0 - 0.5 <= ch["x0"] < x1 - 0.5:
                            out.append(ch)
            return "".join(c["text"] for c in sorted(out, key=lambda c: c["x0"]))

        # header text (page 1 has the masthead + period band)
        head = {}
        for t, cs in sorted(chars_by_top.items()):
            s = "".join(c["text"] for c in sorted(cs, key=lambda c: c["x0"]))
            size = round(cs[0]["size"], 1)
            if size == 18.0:
                head["title"] = s
            elif size == 13.0:
                head["band13"] = [
                    "".join(c["text"] for c in sorted(cs, key=lambda c: c["x0"]) if c["x0"] < p.width / 2),
                    "".join(c["text"] for c in sorted(cs, key=lambda c: c["x0"]) if c["x0"] >= p.width / 2),
                ]
            elif size == 11.0:
                left = [c for c in sorted(cs, key=lambda c: c["x0"]) if c["x0"] < p.width / 2]
                right = [c for c in sorted(cs, key=lambda c: c["x0"]) if c["x0"] >= p.width / 2]
                head["shop"] = "".join(c["text"] for c in left)
                head["bond"] = "".join(c["text"] for c in right)

        rows = []
        for t in sorted(cells):
            cs = cells[t]
            bot = cs[0]["bottom"]
            fill = norm(cs[0].get("non_stroking_color"))
            label = text_in(t, bot, xs[0], xs[1]).strip()
            if label.upper().startswith("BRAND/PACK"):
                continue                                  # column header row
            vals, val_dim = [], []
            for i in range(1, 5):
                seg = []
                for tt, cc in chars_by_top.items():
                    if t - 1 <= tt <= bot - 4:
                        for ch in cc:
                            if xs[i] - 0.5 <= ch["x0"] < xs[i + 1] - 0.5:
                                seg.append(ch)
                seg = sorted(seg, key=lambda c: c["x0"])
                vals.append("".join(c["text"] for c in seg))
                val_dim.append(bool(seg) and norm(seg[0].get("non_stroking_color")) == (0.78, 0.8, 0.84))
            if not label and not any(vals):
                continue
            kind = "pack"
            if fill == NAVY:
                kind = "total" if label.upper().startswith("TOTAL") else "brand"
            rows.append({"kind": kind, "label": label, "values": vals,
                         "dim": val_dim, "zebra": fill == ZEBRA})
        pages.append({"head": head, "rows": rows,
                      "cols": xs, "height": round(p.height, 1), "width": round(p.width, 1)})
    return pages


if __name__ == "__main__":
    src = sys.argv[1]
    out = sys.argv[2]
    data = parse(src)
    json.dump(data, open(out, "w"), indent=1)
    tot = sum(len(p["rows"]) for p in data)
    print(f"{len(data)} pages, {tot} rows -> {out}")
    print("page 1 head:", data[0]["head"])
    print("page 2 head:", data[1]["head"])
    print("first rows:", json.dumps(data[0]["rows"][:3], indent=1))
    print("total row :", json.dumps(data[0]["rows"][-1]))
