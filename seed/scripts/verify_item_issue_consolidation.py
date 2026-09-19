"""Diff a rebuilt ITEM ISSUE CONSOLIDATION against its source, cell by cell.

    python3 verify_item_issue_consolidation.py <source>.pdf <rebuild>.pdf

Re-reads BOTH PDFs independently and diffs by x-position.

Deliberately does NOT trust data.json or the build script — it parses the
rendered output the same way it parses the original.
"""
import pdfplumber, sys

def rowsof(page):
    rows = {}
    for w in page.extract_words():
        rows.setdefault(round(w['top'], 0), []).append(w)
    out = []
    for k in sorted(rows):
        if out and k - out[-1][0] <= 2.5: out[-1][1].extend(rows[k])
        else: out.append([k, list(rows[k])])
    for y, ws in out: ws.sort(key=lambda w: w['x0'])
    return out

def parse(path):
    page = pdfplumber.open(path).pages[0]
    rows = rowsof(page)
    # the column-label row is the one carrying STN twice
    lab_row = next(r for r in rows if [w['text'] for w in r[1]].count('STN') == 2)
    # merge tokens that belong to one label ("C FED", "5 AUG"): an intra-label
    # gap is a few points, a between-column gap is ~20pt
    hdr, cents, i = lab_row[1], [], 0
    while i < len(hdr):
        x0, x1 = hdr[i]['x0'], hdr[i]['x1']
        while i+1 < len(hdr) and hdr[i+1]['x0'] - x1 < 4.0:
            i += 1; x1 = hdr[i]['x1']
        cents.append((x0 + x1)/2); i += 1
    # LAST MONTH's second line may share the label row's y band or sit above it
    if len(cents) == 14:
        lm = next(r for r in rows
                  if r is not lab_row and any(t['text'].startswith('(31') for t in r[1]))
        cents.append((lm[1][0]['x0'] + lm[1][-1]['x1'])/2)
    assert len(cents) == 15, (path, len(cents), [round(x) for x in cents])
    bounds = [(cents[i]+cents[i+1])/2 for i in range(14)]
    label_max = cents[0] - (bounds[0] - cents[0])

    def colof(x):
        if x < label_max: return -1
        for i, b in enumerate(bounds):
            if x < b: return i
        return 14

    cells = {}
    for y, ws in rows:
        if y <= lab_row[0] + 5: continue
        if any(w['text'] == 'Page' for w in ws): continue
        lab, vals = [], [[] for _ in range(15)]
        for w in ws:
            c = colof((w['x0'] + w['x1'])/2)
            if c == -1: lab.append(w['text'])
            else: vals[c].append(w['text'])
        key = " ".join(lab)
        cells[key] = ["".join(v) for v in vals]
    return cells

a, b = parse(sys.argv[1]), parse(sys.argv[2])
fails, checked = [], 0
if set(a) != set(b):
    fails.append(f"ROW SET differs: only-source={set(a)-set(b)} only-rebuild={set(b)-set(a)}")
MERGED = {"Day Sale", "Industry Total"}   # values span a whole period block
for k in a:
    if k not in b: continue
    if k in MERGED:
        # a merged value is centred over its block, so its x-centre can fall in
        # either neighbouring bucket -- compare the ordered value sequence
        xs = [v for v in a[k] if v]
        ys = [v for v in b[k] if v]
        checked += len(xs)
        if xs != ys: fails.append(f"{k!r} merged row: source={xs} rebuild={ys}")
        continue
    for i, (x, y) in enumerate(zip(a[k], b[k])):
        checked += 1
        if x != y: fails.append(f"{k!r} col{i}: source={x!r} rebuild={y!r}")
print(f"rows: source {len(a)} | rebuild {len(b)}")
print(f"cells compared: {checked}")
if fails:
    print(f"\n*** {len(fails)} MISMATCH(ES) ***")
    for f in fails[:40]: print("  ", f)
    sys.exit(1)
print("\nALL CELLS IDENTICAL")
