"""Re-derive the rebuilt PDF from scratch and diff it against the SOURCE pdf --
deliberately not against cumulative.json, which would only prove the code agrees
with itself."""
import sys
from collections import Counter
sys.path.insert(0, ".")
from extract_cumulative import parse

SRC = "/sessions/gifted-adoring-hypatia/mnt/uploads/shop_sales_cumulative_bond_alappuzha-6.pdf"
NEW = "SHOP SALES CUMULATIVE - MOBILE.pdf"

a, b = parse(SRC), parse(NEW)
fails = []

def chk(cond, msg):
    if not cond: fails.append(msg)

chk(len(a) == len(b), f"page count {len(a)} vs {len(b)}")
print(f"pages: {len(a)} vs {len(b)}")

rows_a = sum(len(p["rows"]) for p in a)
rows_b = sum(len(p["rows"]) for p in b)
chk(rows_a == rows_b, f"row count {rows_a} vs {rows_b}")
print(f"rows : {rows_a} vs {rows_b}")

vals = dims = zeb = labs = 0
for i, (pa, pb) in enumerate(zip(a, b), 1):
    chk(pa["head"].get("shop") == pb["head"].get("shop"),
        f"pg{i} shop {pa['head'].get('shop')!r} vs {pb['head'].get('shop')!r}")
    chk(pa["head"].get("bond") == pb["head"].get("bond"), f"pg{i} bond")
    chk(pa["head"].get("band13") == pb["head"].get("band13"), f"pg{i} band")
    chk(pa["head"].get("title") == pb["head"].get("title"), f"pg{i} title")
    chk(len(pa["rows"]) == len(pb["rows"]),
        f"pg{i} rows {len(pa['rows'])} vs {len(pb['rows'])}")
    for j, (ra, rb) in enumerate(zip(pa["rows"], pb["rows"])):
        if ra["label"].strip() != rb["label"].strip():
            fails.append(f"pg{i} r{j} label {ra['label']!r} vs {rb['label']!r}")
        else: labs += 1
        if ra["values"] != rb["values"]:
            fails.append(f"pg{i} r{j} values {ra['values']} vs {rb['values']}")
        else: vals += 4
        if ra["dim"] != rb["dim"]:
            fails.append(f"pg{i} r{j} dim {ra['dim']} vs {rb['dim']}")
        else: dims += 4
        if ra["zebra"] != rb["zebra"]:
            fails.append(f"pg{i} r{j} zebra {ra['zebra']} vs {rb['zebra']}")
        else: zeb += 1
        chk(ra["kind"] == rb["kind"], f"pg{i} r{j} kind")

print(f"labels matched      : {labs}")
print(f"values matched      : {vals}")
print(f"dim (zero) flags    : {dims}")
print(f"zebra flags         : {zeb}")

# independent numeric sweep straight off the raw char stream
import pdfplumber, re
def toks(path):
    out = []
    for p in pdfplumber.open(path).pages:
        out += re.findall(r'\d+\.\d\d', "".join(
            ch['text'] for ch in sorted(p.chars, key=lambda c: (round(c['top']/5), c['x0']))))
    return out
ta, tb = toks(SRC), toks(NEW)
chk(ta == tb, "raw numeric stream differs")
print(f"raw numeric tokens  : {len(ta)} vs {len(tb)} identical={ta==tb}")

def fills(path):
    c = Counter()
    for p in pdfplumber.open(path).pages:
        for r in p.rects:
            v = r.get('non_stroking_color')
            v = (round(v,2),)*3 if isinstance(v,(int,float)) else tuple(round(x,2) for x in v)
            c[v] += 1
    return c
fa, fb = fills(SRC), fills(NEW)
print(f"\nsource fills: {dict(fa)}")
print(f"new    fills: {dict(fb)}")
# source paints 5 cells per row, the rebuild paints one band per row
scale_ok = all(fb[k] * 5 == v or k == (1.0,0.74,0.19) or fb[k] == v for k, v in fa.items())
print(f"fill classes present in both: {set(fa) == set(fb)}")

w = pdfplumber.open(NEW).pages[0].width
print(f"\npage width: {pdfplumber.open(SRC).pages[0].width:.1f} -> {w:.1f}pt "
      f"({pdfplumber.open(SRC).pages[0].width/w:.2f}x larger type at fit-width)")

print("\n" + ("ALL CHECKS PASS" if not fails else f"{len(fails)} FAILURES"))
for f in fails[:20]: print("  ", f)
