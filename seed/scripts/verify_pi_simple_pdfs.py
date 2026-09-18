#!/usr/bin/env python3
"""
End-to-end accuracy check for the simple per-bond PI PDFs.

This deliberately does NOT trust the PI INSIGHTS workbook. It re-parses the
original KSBC ERP .xls exports with an independent parser and compares those
numbers against what is actually rendered in the finished PDFs. If the builder,
the workbook, or the PDF layer had corrupted a value, this would catch it —
re-deriving from the same intermediate would only prove the code agrees with
itself.

Checks, per shop x brand:
    RL, RQ, MQ  == raw August export
    CHG         == August MQ - July MQ, from the raw exports
Plus per-shop TOTAL rows, per-bond ALL SHOPS rows, page/coverage checks, and a
scan for any truncated brand label.
"""
import collections
import glob
import html
import os
import re
import sys

import pymupdf

BASE = "/sessions/modest-lucid-cray/mnt/Claude"
PDFDIR = os.path.join(BASE, "PURCHASE INSTRUCTION", "BOND PI SHEETS - AUGUST")
CUR, PRIOR = "AUGUST", "JULY"

CELL = re.compile(r"<t[dh][^>]*>(.*?)</t[dh]>", re.S | re.I)
TAG = re.compile(r"<[^>]+>")
SHOP = re.compile(r"Shop\s*:\s*([0-9]+)\s*[- ]\s*([^,]*)", re.I)
MONTH = re.compile(r"Report\s*Month\s*:\s*([A-Za-z]+)", re.I)


def parse_raw(path):
    """-> (month, shop_code, {full_brand: (rl, rq, mq)}) straight from the export."""
    raw = open(path, "rb").read().decode("utf-8", "ignore")
    cells = [html.unescape(TAG.sub("", c)).strip() for c in CELL.findall(raw)]
    blob = " ".join(cells[:10])

    ms, mm_ = SHOP.search(blob), MONTH.search(blob)
    if not ms or not mm_:
        return None
    shop, month = ms.group(1).strip(), mm_.group(1).upper()

    try:
        hdr = cells.index("MQ(c/s)")
    except ValueError:
        return month, shop, {}

    out = {}
    i = hdr + 1
    while i + 6 < len(cells) + 1 and i + 6 <= len(cells):
        chunk = cells[i : i + 7]
        if len(chunk) < 7 or not chunk[0].isdigit():
            break
        _sr, _code, brand, _prev, rl, rq, mq = chunk

        def n(v):
            try:
                return float(str(v).replace(",", ""))
            except ValueError:
                return 0.0

        b = brand.strip()
        if b:
            p = out.get(b, (0.0, 0.0, 0.0))
            out[b] = (p[0] + n(rl), p[1] + n(rq), p[2] + n(mq))
        i += 7
    return month, shop, out


def load_all_raws():
    aug, jul = collections.defaultdict(dict), collections.defaultdict(dict)
    seen = collections.Counter()
    files = glob.glob(os.path.join(BASE, "PURCHASE INSTRUCTION", "**", "*.xls"), recursive=True)
    for f in files:
        r = parse_raw(f)
        if not r:
            print(f"  !! unparseable: {os.path.basename(f)}")
            continue
        month, shop, brands = r
        tgt = aug if month == CUR else jul if month == PRIOR else None
        if tgt is None:
            continue
        seen[(month, shop)] += 1
        for b, v in brands.items():
            p = tgt[shop].get(b, (0.0, 0.0, 0.0))
            tgt[shop][b] = (p[0] + v[0], p[1] + v[1], p[2] + v[2])
    dups = {k: c for k, c in seen.items() if c > 1}
    return aug, jul, len(files), dups


NUM = re.compile(r"^[+\-]?[\d,]+$|^·$")


def num(t):
    return 0.0 if t == "·" else float(t.replace(",", "").replace("+", ""))


def read_pdfs():
    """-> bonds{bond: {shop: {brand: (rl,rq,rq,chg)}}}, totals, allshops, meta"""
    data = collections.defaultdict(lambda: collections.defaultdict(dict))
    totals, allshops, pages, notes = {}, {}, {}, collections.defaultdict(list)
    for f in sorted(os.listdir(PDFDIR)):
        if not f.endswith(".pdf"):
            continue
        bond = f.split(" — ")[0]
        doc = pymupdf.open(os.path.join(PDFDIR, f))
        pages[bond] = len(doc)
        cur = None
        for page in doc:
            rows = collections.defaultdict(list)
            for blk in page.get_text("dict")["blocks"]:
                for ln in blk.get("lines", []):
                    for sp in ln["spans"]:
                        t = sp["text"].strip()
                        if t:
                            rows[round(sp["origin"][1], 1)].append(
                                (sp["origin"][0], t, round(sp["size"], 1))
                            )
            for y in sorted(rows):
                items = sorted(rows[y])
                sizes = {s for _, _, s in items}
                txt = " ".join(t for _, t, _ in items)
                if 10.0 in sizes:                                  # shop banner
                    m = re.search(r"\((\d{4,6})\)", txt)
                    if m:
                        cur = m.group(1)
                        data[bond].setdefault(cur, {})
                    continue
                if "No August PI received" in txt:                 # lost-PI footnote
                    notes[bond].append(txt)
                    continue
                if 8.0 in sizes:                                   # column header
                    continue
                toks = [t for _, t, _ in items]
                nums = [t for t in toks if NUM.match(t)]
                if len(nums) < 4:
                    continue
                label = " ".join(t for t in toks if not NUM.match(t)).strip()
                vals = tuple(num(t) for t in nums[-4:])
                if "ALL SHOPS" in label:
                    allshops[bond] = vals
                elif label == "TOTAL":
                    totals[(bond, cur)] = vals
                elif cur:
                    data[bond][cur][label] = vals
        doc.close()
    return data, totals, allshops, pages, notes


def main():
    print("parsing raw KSBC exports (independent of the workbook)...")
    aug, jul, nfiles, dups = load_all_raws()
    print(f"  files scanned: {nfiles}   August shops: {len(aug)}   July shops: {len(jul)}")
    if dups:
        print(f"  !! duplicate shop files: {dups}")

    print("reading finished PDFs...")
    pdf, totals, allshops, pages, notes = read_pdfs()
    npdf_shops = sum(len(v) for v in pdf.values())
    print(f"  bonds: {len(pdf)}   shop blocks: {npdf_shops}   pages: {sum(pages.values())}")

    errs, cells, checked_shops = [], 0, 0
    truncated = []

    for bond, shops in pdf.items():
        for shop, brands in shops.items():
            checked_shops += 1
            a, j = aug.get(shop, {}), jul.get(shop, {})
            if shop not in aug:
                errs.append(f"{bond} {shop}: rendered but absent from the August raws")
                continue
            # every brand the raws carry (either month) must be on the page
            expect = set(a) | {b for b, v in j.items() if v[2] > 0}
            for b in expect:
                arl, arq, amq = a.get(b, (0.0, 0.0, 0.0))
                jmq = j.get(b, (0.0, 0.0, 0.0))[2]
                if b not in brands:
                    if amq == 0 and jmq == 0:
                        continue          # nothing to show either month
                    errs.append(f"{bond} {shop}: brand row missing '{b}'")
                    continue
                got = brands[b]
                cells += 4
                for lbl, g, e in (
                    ("RL", got[0], arl),
                    ("RQ", got[1], arq),
                    ("MQ", got[2], amq),
                    ("CHG", got[3], amq - jmq),
                ):
                    if abs(g - e) > 0.005:
                        errs.append(f"{bond} {shop} '{b}' {lbl}: pdf {g:g} vs raw {e:g}")
            for b in brands:
                if b not in expect:
                    errs.append(f"{bond} {shop}: extra brand row '{b}' not in raws")
                if b.endswith(("…", "..")) or (b and not b[-1].isalnum() and b[-1] not in ")'"):
                    truncated.append(f"{bond} {shop}: '{b}'")

            # per-shop TOTAL row
            t = totals.get((bond, shop))
            if t is None:
                errs.append(f"{bond} {shop}: no TOTAL row")
            else:
                e = (
                    sum(v[0] for v in a.values()),
                    sum(v[1] for v in a.values()),
                    sum(v[2] for v in a.values()),
                    sum(v[2] for v in a.values()) - sum(v[2] for v in j.values()),
                )
                cells += 4
                for i, lbl in enumerate(("RL", "RQ", "MQ", "CHG")):
                    if abs(t[i] - e[i]) > 0.005:
                        errs.append(f"{bond} {shop} TOTAL {lbl}: pdf {t[i]:g} vs raw {e[i]:g}")

    # per-bond ALL SHOPS row, incl. shops that lost their PI this month
    for bond, got in allshops.items():
        mine = set(pdf[bond])
        e_rl = sum(sum(v[0] for v in aug[s].values()) for s in mine if s in aug)
        e_rq = sum(sum(v[1] for v in aug[s].values()) for s in mine if s in aug)
        e_mq = sum(sum(v[2] for v in aug[s].values()) for s in mine if s in aug)
        e_ch = e_mq - sum(sum(v[2] for v in jul.get(s, {}).values()) for s in mine)
        # subtract the ceiling of any shop noted as having no August PI
        for n in notes.get(bond, []):
            m = re.search(r"lost the ([\d,]+) cs", n)
            if m:
                e_ch -= float(m.group(1).replace(",", ""))
        cells += 4
        for i, (lbl, e) in enumerate((("RL", e_rl), ("RQ", e_rq), ("MQ", e_mq), ("CHG", e_ch))):
            if abs(got[i] - e) > 0.005:
                errs.append(f"{bond} ALL SHOPS {lbl}: pdf {got[i]:g} vs raw {e:g}")

    # coverage: which raw August shops never made it into a PDF
    rendered = {s for v in pdf.values() for s in v}
    unrendered = sorted(set(aug) - rendered)
    lost_pi = sorted(set(jul) - set(aug))

    print()
    print(f"shop blocks verified : {checked_shops}")
    print(f"cells compared       : {cells:,}")
    print(f"truncated labels     : {len(truncated)}")
    print(f"ERRORS               : {len(errs)}")
    for e in errs[:25]:
        print("   ", e)
    if len(errs) > 25:
        print(f"    ... {len(errs)-25} more")

    print(f"\nAugust raw shops not rendered ({len(unrendered)}):")
    for s in unrendered:
        n = sum(v[2] for v in aug[s].values())
        print(f"    {s}  (August ΣMQ {n:g})")
    print(f"\nShops with a July PI but none in August ({len(lost_pi)}):")
    for s in lost_pi:
        print(f"    {s}  (July ΣMQ {sum(v[2] for v in jul[s].values()):g})")

    return 1 if errs else 0


if __name__ == "__main__":
    sys.exit(main())
