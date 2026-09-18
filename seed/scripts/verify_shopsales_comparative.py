"""Per-cell diff of the comparative report vs its source.

Tokenising a concatenated char stream is unsafe here: the row LABEL contains a
digit ("CLUSTER - 1") which glues onto the first numeric cell. Columns are read
by x-position instead, so labels and figures can never bleed into each other.
"""
import pdfplumber

SRC = "/sessions/gifted-adoring-hypatia/mnt/uploads/comparative_shopsales_current-7.pdf"
NEW = "SHOPSALES COMPARATIVE - MOBILE.pdf"

def cells(path, label_x):
    p = pdfplumber.open(path).pages[0]
    rows = {}
    for ch in p.chars:
        if round(ch["size"], 1) != 11.0:      # body rows only
            continue
        rows.setdefault(round(ch["top"] / 6), []).append(ch)
    out = []
    for k in sorted(rows):
        cs = sorted(rows[k], key=lambda c: c["x0"])
        label = "".join(c["text"] for c in cs if c["x0"] < label_x).strip()
        vals, cur, prev = [], "", None
        for c in [c for c in cs if c["x0"] >= label_x]:
            if prev is not None and c["x0"] - prev["x1"] > 1.6:
                vals.append(cur); cur = ""
            cur += c["text"]; prev = c
        if cur: vals.append(cur)
        out.append((label, [v.strip() for v in vals if v.strip()]))
    return out

a = cells(SRC, 125)      # source BOND column ends at x=125
b = cells(NEW, 104)       # rebuilt BOND column is narrower
print(f"rows: {len(a)} vs {len(b)}")

fails, n = [], 0
for i, ((la, va), (lb, vb)) in enumerate(zip(a, b)):
    if va != vb:
        fails.append(f"row {i}: {la!r} {va} vs {lb!r} {vb}")
    else:
        n += len(va)
print(f"numeric cells matched: {n}")

print("\nrow labels (source -> rebuilt):")
for (la, _), (lb, _) in zip(a, b):
    mark = "  " if la == lb else "->"
    print(f" {mark} {la:<20} {lb}")

print("\n" + ("ALL FIGURES MATCH" if not fails else f"{len(fails)} MISMATCH"))
for f in fails: print("  ", f)
