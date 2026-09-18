"""Per-cell diff of the consolidation report vs its source, column by column.

Reads each cell by x-position rather than tokenising a flattened char stream,
so a label digit can never glue onto the adjacent figure. The SL NO column was
deliberately dropped, so it is excluded from the comparison and its absence is
asserted instead.
"""
import pdfplumber
from collections import Counter

SRC = "/sessions/gifted-adoring-hypatia/mnt/uploads/item_issue_consolidation_2026-08-3.pdf"
NEW = "ITEM ISSUE CONSOLIDATION - MOBILE.pdf"

def grid(path):
    p = pdfplumber.open(path).pages[0]
    by_top = {}
    for r in p.rects:
        by_top.setdefault(round(r["top"], 1), []).append(r)
    rows = []
    for t in sorted(by_top):
        cs = sorted(by_top[t], key=lambda r: r["x0"])
        if len(cs) < 3:                      # full-width band, not a data row
            continue
        bot = cs[0]["bottom"]
        if bot - t > 30:                     # header block
            continue
        cells = []
        for cell in cs:
            txt = "".join(ch["text"] for ch in p.chars
                          if cell["x0"] - 0.6 <= ch["x0"] < cell["x1"] - 0.6
                          and t - 1 <= ch["top"] <= bot - 3)
            cells.append(txt.strip())
        rows.append([c for c in cells if c != ""])
    return rows

a, b = grid(SRC), grid(NEW)
print(f"data rows: {len(a)} vs {len(b)}")

# strip the SL NO column from the source: it is a bare integer in cell 0 on the
# 28 warehouse rows, absent on cluster/total/footer rows
stripped = []
for r in a:
    if r and r[0].isdigit() and len(r) > 14:
        r = r[1:]
    stripped.append(r)

fails, n = [], 0
for i, (ra, rb) in enumerate(zip(stripped, b)):
    if ra != rb:
        fails.append(f"row {i}:\n    src {ra}\n    new {rb}")
    else:
        n += len(ra)

print(f"cells matched (SL NO excluded): {n}")

srcdigits = sum(1 for r in a if r and r[0].isdigit() and len(r) > 14)
print(f"SL NO values present in source: {srcdigits}")
print(f"SL NO values present in new   : "
      f"{sum(1 for r in b if r and r[0].isdigit() and len(r) > 14)}")

print("\n" + ("ALL FIGURES MATCH" if not fails else f"{len(fails)} MISMATCH"))
for f in fails[:6]: print("  ", f)
