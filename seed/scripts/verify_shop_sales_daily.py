import sys, json
sys.path.insert(0,".")
from extract_daily import parse
SRC="/sessions/gifted-adoring-hypatia/mnt/uploads/shop_sales_daily_bond_current-9.pdf"
NEW="SHOP SALES DAILY - MOBILE.pdf"
a,b=parse(SRC),parse(NEW)
f=[]; cells=0
def chk(c,m):
    if not c: f.append(m)
chk(a["title"]==b["title"], f"title {a['title']!r} vs {b['title']!r}")
chk(a["band_left"]==b["band_left"], "band left")
chk(a["band_right"]==b["band_right"], "band right")
chk(a["weekday"]==b["weekday"], f"weekday {a['weekday']} vs {b['weekday']}")
chk(a["daynum"]==b["daynum"], f"daynum {a['daynum']} vs {b['daynum']}")
chk(len(a["rows"])==len(b["rows"]), f"rows {len(a['rows'])} vs {len(b['rows'])}")
for x,y in zip(a["rows"],b["rows"]):
    if x["name"]!=y["name"]: f.append(f"name {x['name']!r} vs {y['name']!r}")
    if x["values"]!=y["values"]: f.append(f"{x['name']} values differ")
    else: cells+=len(x["values"])
    if x["dim"]!=y["dim"]: f.append(f"{x['name']} dim flags differ")
    if x["zebra"]!=y["zebra"]: f.append(f"{x['name']} zebra differs")
    if x["total"]!=y["total"]: f.append(f"{x['name']} total flag differs")
print(f"rows {len(a['rows'])}  columns {len(a['weekday'])}  cells matched {cells}")
import pdfplumber
print("markers (curves) in source:", len(pdfplumber.open(SRC).pages[0].curves),
      " in rebuild:", len(pdfplumber.open(NEW).pages[0].curves))
print("\n"+("ALL MATCH" if not f else f"{len(f)} FAILURES"))
for x in f[:8]: print("   ",x)
