"""Build warehouse stock dashboard + append history."""
import re, csv, sys, os, json
from pathlib import Path
from datetime import datetime
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FOLDER = Path(__file__).resolve().parent.parent.parent / "Warehouse stock"
HIST_DIR = FOLDER / "_history"
HIST_CSV = HIST_DIR / "stock_history.csv"
BRAND_PACK_CSV = HIST_DIR / "brand_pack_history.csv"

# ---------- Palette (locked) ----------
BG_DARK = "FF0B1929"
BG_MID = "FF122240"
BG_ACCENT = "FF1E3A5F"
FG_CYAN = "FF00D4FF"
FG_GREEN = "FF00F5A0"
FG_WHITE = "FFFFFFFF"
FG_GOLD = "FF00D5FF"
# ---------- helpers ----------
def clean(s):
    if s is None: return ""
    s = re.sub(r'<[^>]+>', '', s)
    s = s.replace('&amp;', '&').replace('&nbsp;', ' ').strip()
    return s

_num_warnings = []
def _num(s):
    try: return float(s.replace(',',''))
    except (AttributeError, ValueError):
        # Flag genuinely unexpected non-numerics (not blanks/dashes/NIL) instead of
        # silently zeroing them - a silent 0 here would undercount stock.
        if s not in (None, '', '-', '–', 'NIL', 'nil', 'N/A', 'NA'):
            _num_warnings.append(s)
        return 0.0

def extract_warehouse_name(txt):
    m = re.search(r'Warehouse\s*:\s*<b>([^<]+)</b>', txt)
    if not m: return None
    raw = m.group(1).strip()
    m2 = re.match(r'WH-([A-Z ]+?)(?:\s+(?:FL|RFL).*)?$', raw)
    if m2:
        return m2.group(1).strip()
    # Fallback (name has digits/&/odd licence code): strip "WH-" then keep tokens
    # until the first licence-code-looking token. Mirrors canon_wh() in
    # maintain_inbound_history.py so both scripts canonicalise identically, and
    # does NOT truncate multi-word names to the first word (the old fallback bug).
    s = raw.upper()
    if s.startswith('WH-'): s = s[3:].lstrip()
    keep = []
    for tok in s.split():
        if any(c.isdigit() for c in tok) or '/' in tok or tok.startswith('FL') or tok.startswith('RFL'):
            break
        keep.append(tok)
    return ' '.join(keep).strip() or s

# A Bevco stock export carries the date in THREE places, and only one of them
# answers "which day is this stock?":
#
#   <b>Report Date & Time : </b>19-Sep-2026            <- page header, no time
#   Report Date &amp; Time : 19-Sep-2026 09:09 AM      <- when it was printed
#   Warehouse : <b>WH-KOLLAM ...</b>,Report Period : <b>19-Sep-2026</b>
#
# The first two are the ERP's print stamp. The third sits inside the report,
# beside the warehouse it belongs to, and is the period the report was asked
# for. Pull a past day - which Bevco allows - and the print stamp is today
# while the period is the day you chose.
#
# This used to read the print stamp, so every back-dated export was filed
# under the day it was DOWNLOADED. September 4th, 6th, 13th, 20th and 21st
# were pulled correctly, on the 23rd, and all five were written into the 23rd.
# The day asked for is the day it belongs to, so Report Period is read first
# and the print stamp is only a fallback for an export that has no period.
_RE_PERIOD = re.compile(
    r'Report\s+Period\s*:\s*(?:<b>\s*)?([0-9]{1,2}-[A-Za-z]{3}-[0-9]{4})', re.I)
_RE_PRINTED = re.compile(
    r'Report\s+Date\s*(?:&amp;|&)\s*Time\s*:\s*(?:</b>)?\s*([0-9]{1,2}-[A-Za-z]{3}-[0-9]{4})', re.I)


def extract_report_date(txt):
    m = _RE_PERIOD.search(txt) or _RE_PRINTED.search(txt)
    if not m:
        return None
    # '04-Sep-2026' and '4-Sep-2026' are the same day and the two fields do not
    # agree on the padding, so it is normalised here rather than leaving two
    # spellings to be compared downstream.
    dt = datetime.strptime(m.group(1), "%d-%b-%Y")
    return f"{dt.day}-{dt.strftime('%b')}-{dt.year}"


def extract_printed_date(txt):
    """When the ERP printed it - not what it is for. Reported, never filed."""
    m = _RE_PRINTED.search(txt)
    return m.group(1) if m else None

def parse_rows(txt):
    cells = [clean(m.group(1)) for m in re.finditer(r'<td[^>]*>(.*?)</td>', txt, flags=re.DOTALL)]
    rows, n, idx = [], len(cells), 0
    while idx < n:
        c = cells[idx]
        if c.isdigit() and idx + 19 <= n:
            item = cells[idx+1]; pack = cells[idx+3]
            if item and any(ch.isalpha() for ch in item) and pack and ('ML' in pack.upper() or 'LTR' in pack.upper()):
                try:
                    rows.append({
                        'item': item,
                        'code': cells[idx+2],
                        'pack': pack,
                        'phys_case': _num(cells[idx+4]),
                        'allot_case': _num(cells[idx+6]),
                        'pend_case': _num(cells[idx+12]),
                    })
                    idx += 19; continue
                except Exception: pass
        idx += 1
    return rows

# ---------- Parse all raws ----------
# Raw-drop filename guard (widened 11 Aug 2026). Bevco/browser downloads arrive
# in TWO shapes for the same export: the historic underscore form
# `Report_11-Aug-2026__7_.xls` and a space form `Report 11-Aug-2026 (7).xls`
# (Chrome's default duplicate-suffix naming). The old `startswith("Report_")`
# test silently matched ZERO files on a space-named drop, so the build reported
# `STATUS: NO_RAWS` on a folder holding a full 28-file day and the dashboard
# went stale without an error. Accept either separator; inert on the old form.
# Extension (widened 24 Sep 2026). The app's warehouse_stock stream accepts
# .xls AND .xlsx, and a portal handing back 'Report_21-Sep-2026.XLS' is the
# same export in a different case - but `p.suffix == '.xls'` saw none of
# those. A raw this glob misses is a raw nothing ever reads: the build
# printed NO_RAWS, exited 0, the job went green, and the day was absent from
# the history, the reports and the status calendar with a successful run
# behind it. The gate that accepts the upload and the glob that reads it have
# to agree, so this takes whatever validate_warehouse_name() let through.
raw_files = sorted(p for p in FOLDER.iterdir()
                   if re.match(r'^Report[ _]', p.name)
                   and p.suffix.lower() in ('.xls', '.xlsx'))

# No-raws guard (added 30 Jul 2026). Documented in warehouse-stock-workbook.md
# step 9 and consumed by the `warehouse-stock` scheduled task, but never
# implemented — an empty folder crashed with IndexError on the report_date
# filename fallback instead of exiting cleanly. A clean folder is the NORMAL
# steady state (raws are deleted after every successful build), so a re-fire of
# the task on an already-built day must be a quiet no-op, not a traceback.
if not raw_files:
    print("STATUS: NO_RAWS")
    print(f"    No Report_* files in '{FOLDER}'. Nothing to build.")
    # A re-fire of the scheduled task on an already-built day is a quiet
    # no-op, and stays one. An upload is not: KSD_EXPECT_RAWS says the app
    # has just written this day's files into the folder, so finding none
    # means they are not files this script reads - and exiting 0 there filed
    # the day as built when nothing had been built.
    sys.exit(3 if os.environ.get("KSD_EXPECT_RAWS") else 0)

warehouse_rows = {}   # wh_name -> list of dict rows
report_date = None
report_dates_seen = {}   # date-string -> count of files carrying it
unreadable = []
printed_seen = set()
periods_seen = set()   # Report Period ONLY - never the print stamp
for p in raw_files:
    txt = p.read_text(errors='replace')
    if 'Warehouse :' not in txt:
        # The frameset root of a Bevco export legitimately lands here. So does
        # a workbook re-saved out of Excel, a download that was cut short, and
        # an error page saved under the export's name - none of which is a
        # stock file, and all of which used to pass through in silence.
        unreadable.append(p.name)
        continue
    wh = extract_warehouse_name(txt)
    d = extract_report_date(txt)
    if d:
        report_dates_seen[d] = report_dates_seen.get(d, 0) + 1
        if not report_date: report_date = d
    pulled = extract_printed_date(txt)
    if pulled:
        printed_seen.add(pulled)
    _per = _RE_PERIOD.search(txt)
    if _per:
        _pd = datetime.strptime(_per.group(1), "%d-%b-%Y")
        periods_seen.add(f"{_pd.day}-{_pd.strftime('%b')}-{_pd.year}")
    rows = parse_rows(txt)
    if wh and rows:
        warehouse_rows[wh] = rows

# Date-consistency guard: every Bevco file in one drop must carry the SAME
# Report Date. A mixed-day batch would otherwise be silently summed under a
# single date (wrong aggregation). Abort so the operator removes the stray file.
if len(report_dates_seen) > 1:
    pairs = ', '.join(f"{k} ({v} file/s)" for k, v in sorted(report_dates_seen.items()))
    print(f"ERROR_MIXED_DATES: raw files span multiple report dates: {pairs}")
    print("    A day's export is built on its own; two days in one folder would "
          "be summed under a single date.")
    print("    ACTION: upload one report date at a time. Uploading the newer "
          "date on its own moves the older day's leftover files aside "
          "automatically, so this clears itself on the next upload.")
    sys.exit(2)

# Nothing-parsed guard (added 24 Sep 2026). Zero warehouses used to run
# straight on: it saved an EMPTY day workbook, wrote a 0-row history entry -
# which, on a date that already had good rows, DELETED them - and only then
# fell over on an empty list. The day was gone from the reports and the
# calendar, and the history was worse than before the run. Nothing is written
# until at least one warehouse has been read.
if not warehouse_rows:
    shown = ', '.join(p.name for p in raw_files[:6])
    print(f"ERROR_NO_ROWS_PARSED: not one warehouse could be read out of "
          f"{len(raw_files)} file(s): {shown}" + (" ..." if len(raw_files) > 6 else ""))
    print("    A Bevco stock export is an HTML table saved as .xls. A file "
          "re-saved from Excel, a download that stopped short, or an error "
          "page saved under the export's name all read as empty here.")
    print("    ACTION: download the day's export from Bevco again and upload it "
          "exactly as it arrives, without opening it in Excel first.")
    sys.exit(4)

if unreadable:
    print(f"UNREADABLE_RAWS: {len(unreadable)} file(s) held no warehouse and were "
          f"skipped: {', '.join(unreadable[:6])}" + (" ..." if len(unreadable) > 6 else ""))

if _num_warnings:
    sample = ', '.join(sorted({str(x) for x in _num_warnings})[:5])
    print(f"⚠️  NUMERIC_COERCE_WARN: {len(_num_warnings)} non-numeric stock cell(s) read as 0 (e.g. {sample})")

# THE DATE PICKED ON UPLOAD IS THE DATE (24 Sep 2026). The office chooses the
# day in the upload dialog and pulls that day's export for it; the app passes
# that choice here as KSD_REPORT_DATE and the stock is filed under it. The
# files are only asked to agree: their Report Period must be the same day, so
# picking the 20th and dropping in the 4th's folder is stopped here rather than
# filing one day's stock under another. What this script used to do instead -
# read a date of its own out of the export, and the wrong one, the print stamp
# - is how five back-dated September days were written into the 23rd.
_picked = os.environ.get("KSD_REPORT_DATE", "").strip()
if _picked:
    _pdt = datetime.strptime(_picked, "%Y-%m-%d")
    _picked_txt = f"{_pdt.day}-{_pdt.strftime('%b')}-{_pdt.year}"
    # Checked against the Report Period alone. The print stamp is the day the
    # file came down and has no say in which day it is for.
    _inside = sorted(periods_seen)
    if _inside and _inside != [_picked_txt]:
        print(f"ERROR_DATE_MISMATCH: the upload was for {_picked_txt}, but the files' "
              f"Report Period says {', '.join(_inside)}.")
        print(f"    ACTION: upload the {_picked_txt} export for {_picked_txt}, or pick "
              f"{', '.join(_inside)} as the date.")
        sys.exit(6)
    report_date = _picked_txt

if not report_date:
    # fallback from filename
    m = re.search(r'(\d{2}-[A-Za-z]{3}-\d{4})', raw_files[0].name)
    report_date = m.group(1) if m else datetime.now().strftime("%d-%b-%Y")

dt = datetime.strptime(report_date, "%d-%b-%Y")
report_iso = dt.strftime("%Y-%m-%d")
day_ord_lookup = {1:'st',2:'nd',3:'rd'}
day = dt.day
suf = day_ord_lookup.get(day if day<20 else day%10, 'th')
month_name = dt.strftime("%B").upper()
display_date = dt.strftime("%d %B %Y").lstrip('0')
output_filename = f"{month_name} {day}{suf} WAREHOUSE STOCK.xlsx"

# ---------- Scope-change guard (added 22 Apr 2026) ----------
# Compare today's warehouse set vs. the most recent prior date in history and
# flag anything missing or newly-added. Catches the BALARAMAPURAM-shaped failure
# where Bevco's export silently drops a warehouse — previously only visible
# post-hoc by reading the dashboard and counting rows.
today_wh_set = set(warehouse_rows.keys())
prior_wh_set = set()
prior_date_for_scope = None
if HIST_CSV.exists():
    with HIST_CSV.open() as f:
        r = csv.reader(f)
        next(r, None)  # header
        rows_by_date = defaultdict(set)
        for row in r:
            if row and row[0] != report_iso:
                rows_by_date[row[0]].add(row[1])
        if rows_by_date:
            prior_date_for_scope = max(rows_by_date.keys())
            prior_wh_set = rows_by_date[prior_date_for_scope]

missing = sorted(prior_wh_set - today_wh_set)
newly_added = sorted(today_wh_set - prior_wh_set)

if prior_date_for_scope:
    print(f"SCOPE_COMPARE_WITH: {prior_date_for_scope} ({len(prior_wh_set)} warehouses) | TODAY: {len(today_wh_set)} warehouses")
else:
    print(f"SCOPE_COMPARE_WITH: NONE | TODAY: {len(today_wh_set)} warehouses")

if missing:
    print(f"⚠️  MISSING_WAREHOUSES: {len(missing)} warehouse(s) present on {prior_date_for_scope} but absent today: {', '.join(missing)}")
    print(f"    ACTION: re-export the missing warehouse(s) from Bevco, drop into 'Warehouse stock/', and re-run.")

# Partial-drop guard (added 24 Sep 2026). A drop holding a handful of the
# network's warehouses used to be filed as the day: the history took it, and
# every figure drawn off that date - network stock, cover days, the stock
# trend - read a fraction of the real position as if it were the whole of it.
# A warning nobody reads is not enough protection for a figure a director
# acts on. The raws stay where they are, so exporting the rest and uploading
# them for the SAME date merges with what is already here and this runs
# through on the full set.
if prior_wh_set and len(today_wh_set) < 0.8 * len(prior_wh_set):
    print(f"ERROR_PARTIAL_DAY: this drop holds {len(today_wh_set)} of the "
          f"{len(prior_wh_set)} warehouses that reported on {prior_date_for_scope}. "
          "Filing it would understate network stock for the day.")
    print(f"    MISSING: {', '.join(missing)}")
    print("    ACTION: export the missing warehouse(s) and upload them for the "
          "same date - they join the ones already uploaded - then this runs "
          "again on the full set.")
    sys.exit(5)
if newly_added:
    print(f"NEW_WAREHOUSES: {len(newly_added)} warehouse(s) appearing today that weren't on {prior_date_for_scope}: {', '.join(newly_added)}")
if not missing and not newly_added and prior_date_for_scope:
    print(f"SCOPE_OK: warehouse set matches {prior_date_for_scope} exactly.")

# ---------- Aggregate ----------
wh_totals = {}
brand_totals = defaultdict(lambda: {'phys':0, 'allot':0, 'pend':0})
pack_totals = defaultdict(lambda: {'phys':0, 'allot':0, 'pend':0})
unique_brands = set()
sku_rows = 0
sku_low_allot = 0
sku_zero_allot = 0

for wh, rows in warehouse_rows.items():
    p = sum(r['phys_case'] for r in rows)
    a = sum(r['allot_case'] for r in rows)
    pe = sum(r['pend_case'] for r in rows)
    wh_totals[wh] = {'phys':p, 'allot':a, 'pend':pe}
    for r in rows:
        brand_totals[r['item']]['phys'] += r['phys_case']
        brand_totals[r['item']]['allot'] += r['allot_case']
        brand_totals[r['item']]['pend'] += r['pend_case']
        pack_totals[r['pack']]['phys'] += r['phys_case']
        pack_totals[r['pack']]['allot'] += r['allot_case']
        pack_totals[r['pack']]['pend'] += r['pend_case']
        unique_brands.add(r['item'])
        sku_rows += 1
        if r['allot_case'] <= 5: sku_low_allot += 1
        if r['allot_case'] == 0: sku_zero_allot += 1

total_phys = sum(v['phys'] for v in wh_totals.values())
total_allot = sum(v['allot'] for v in wh_totals.values())
total_pend = sum(v['pend'] for v in wh_totals.values())
n_wh = len(wh_totals)

allot_pcts = [(v['allot']/v['phys']*100) for v in wh_totals.values() if v['phys']>0]
# Network-weighted ratio (total allotable / total physical) - matches the KPI
# tiles exactly. The old simple mean of per-WH ratios understated this by ~2.4pp
# and read inconsistently against the headline tiles above it.
avg_allot_pct = (total_allot/total_phys*100) if total_phys else 0
strong_count = sum(1 for p in allot_pcts if p >= 80)
weak_count = sum(1 for p in allot_pcts if p < 50)

print(f"Parsed {n_wh} warehouses; total_phys={total_phys}, total_allot={total_allot}, total_pend={total_pend}")
print(f"avg_allot%={avg_allot_pct:.1f}, strong={strong_count}, weak={weak_count}")
print(f"brands={len(unique_brands)}, sku_rows={sku_rows}, low_allot={sku_low_allot}, zero_allot={sku_zero_allot}")

# ---------- Build workbook ----------
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "DASHBOARD"

# Column widths
widths = {'A':2,'B':5,'C':32,'D':12,'E':12,'F':12,'G':12,'H':3,'I':4,'J':30,'K':12,'L':12,'M':12,'N':12,'O':2}
for col,w in widths.items():
    ws.column_dimensions[col].width = w

thin_border = Side(border_style="thin", color="FF1E3A5F")
def style(cell, fill=None, fg=None, bold=False, size=10, align='left', fmt=None, border=False):
    if fill: cell.fill = PatternFill("solid", fgColor=fill)
    cell.font = Font(color=fg or FG_WHITE, bold=bold, size=size, name="Calibri")
    if align=='center': cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    elif align=='right': cell.alignment = Alignment(horizontal='right', vertical='center')
    else: cell.alignment = Alignment(horizontal='left', vertical='center')
    if fmt: cell.number_format = fmt

# Title
ws.merge_cells("B2:N2"); ws["B2"] = "📦 K.S. DISTILLERY — WAREHOUSE STOCK DASHBOARD"
style(ws["B2"], fill=BG_DARK, fg=FG_CYAN, bold=True, size=26, align='center')
ws.row_dimensions[2].height = 40

ws.merge_cells("B3:N3"); ws["B3"] = f"{display_date}  |  Scope: {n_wh} Bevco Warehouses across Kerala"
style(ws["B3"], fill=BG_DARK, fg=FG_WHITE, bold=False, size=11, align='center')

# Fill black background for A1:O4
for r in range(1,5):
    for col in range(1,16):
        c = ws.cell(row=r, column=col)
        if c.fill.patternType is None:
            c.fill = PatternFill("solid", fgColor=BG_DARK)

# KPIs row 7/8/9
ws.merge_cells("B7:G7"); ws["B7"] = "PHYSICAL STOCK"
style(ws["B7"], fill=BG_ACCENT, fg=FG_CYAN, bold=True, size=10, align='center')
ws.merge_cells("H7:J7"); ws["H7"] = "ALLOTABLE"
style(ws["H7"], fill=BG_ACCENT, fg=FG_CYAN, bold=True, size=10, align='center')
ws.merge_cells("K7:N7"); ws["K7"] = "PENDING"
style(ws["K7"], fill=BG_ACCENT, fg=FG_CYAN, bold=True, size=10, align='center')
ws.row_dimensions[7].height = 22

ws.merge_cells("B8:G8"); ws["B8"] = total_phys
style(ws["B8"], fill=BG_ACCENT, fg=FG_WHITE, bold=True, size=24, align='center', fmt='#,##0')
ws.merge_cells("H8:J8"); ws["H8"] = total_allot
style(ws["H8"], fill=BG_ACCENT, fg=FG_WHITE, bold=True, size=24, align='center', fmt='#,##0')
ws.merge_cells("K8:N8"); ws["K8"] = total_pend
style(ws["K8"], fill=BG_ACCENT, fg=FG_WHITE, bold=True, size=24, align='center', fmt='#,##0')
ws.row_dimensions[8].height = 42

ws.merge_cells("B9:G9"); ws["B9"] = "total cases at warehouses"
style(ws["B9"], fill=BG_ACCENT, fg=FG_WHITE, bold=False, size=10, align='center')
ws.merge_cells("H9:J9"); ws["H9"] = "ready for dispatch"
style(ws["H9"], fill=BG_ACCENT, fg=FG_WHITE, bold=False, size=10, align='center')
ws.merge_cells("K9:N9"); ws["K9"] = "committed, not shipped"
style(ws["K9"], fill=BG_ACCENT, fg=FG_WHITE, bold=False, size=10, align='center')
ws.row_dimensions[9].height = 18

# Section headers row 13
ws.merge_cells("B13:G13"); ws["B13"] = "📊 WAREHOUSE STOCK BREAKDOWN"
style(ws["B13"], fill=BG_DARK, fg=FG_CYAN, bold=True, size=14, align='left')
ws.merge_cells("I13:N13"); ws["I13"] = "💡 KEY INSIGHTS"
style(ws["I13"], fill=BG_DARK, fg=FG_CYAN, bold=True, size=14, align='left')

# Subsection labels row 14
ws.merge_cells("B14:G14"); ws["B14"] = "STOCK BY WAREHOUSE (Cases)"
style(ws["B14"], fill=BG_MID, fg=FG_GREEN, bold=True, size=11, align='left')
ws.merge_cells("I14:N14"); ws["I14"] = "TOP 10 BRANDS BY PHYSICAL STOCK"
style(ws["I14"], fill=BG_MID, fg=FG_GREEN, bold=True, size=11, align='left')

# Table headers row 15
headers_left = [('B15','#'),('C15','Warehouse'),('D15','Physical'),('E15','Allotable'),('F15','Pending'),('G15','% of Total')]
headers_right = [('I15','#'),('J15','Brand'),('K15','Physical'),('L15','Allotable'),('M15','Pending'),('N15','% of Total')]
for coord,val in headers_left + headers_right:
    ws[coord] = val
    style(ws[coord], fill=BG_ACCENT, fg=FG_WHITE, bold=True, size=10, align='center')
ws.row_dimensions[15].height = 22

# Warehouse table rows 16..
wh_sorted = sorted(wh_totals.items(), key=lambda kv: -kv[1]['phys'])
start_row = 16
for i, (wh, v) in enumerate(wh_sorted):
    r = start_row + i
    ws.cell(row=r, column=2, value=i+1)
    ws.cell(row=r, column=3, value=wh)
    ws.cell(row=r, column=4, value=int(v['phys']))
    ws.cell(row=r, column=5, value=int(v['allot']))
    ws.cell(row=r, column=6, value=int(v['pend']))
    ws.cell(row=r, column=7, value=(v['phys']/total_phys) if total_phys else 0)
    style(ws.cell(row=r, column=2), fill=BG_MID, fg=FG_WHITE, size=10, align='center')
    style(ws.cell(row=r, column=3), fill=BG_MID, fg=FG_WHITE, size=10, align='left')
    for col in (4,5,6):
        style(ws.cell(row=r, column=col), fill=BG_MID, fg=FG_WHITE, size=10, align='right', fmt='#,##0')
    style(ws.cell(row=r, column=7), fill=BG_MID, fg=FG_WHITE, size=10, align='right', fmt='0.0%')

total_row = start_row + len(wh_sorted)
ws.cell(row=total_row, column=3, value="TOTAL")
ws.cell(row=total_row, column=4, value=int(total_phys))
ws.cell(row=total_row, column=5, value=int(total_allot))
ws.cell(row=total_row, column=6, value=int(total_pend))
ws.cell(row=total_row, column=7, value=1)
for col in range(2,8):
    style(ws.cell(row=total_row, column=col), fill=FG_GOLD, fg=BG_DARK, bold=True, size=10,
          align='center' if col==3 else ('right' if col>=4 else 'center'),
          fmt='#,##0' if col in (4,5,6) else ('0.0%' if col==7 else None))

# Top 10 Brands (right panel)
brands_sorted = sorted(brand_totals.items(), key=lambda kv: -kv[1]['phys'])[:10]
for i,(b, v) in enumerate(brands_sorted):
    r = 16 + i
    ws.cell(row=r, column=9, value=i+1)
    ws.cell(row=r, column=10, value=b)
    ws.cell(row=r, column=11, value=int(v['phys']))
    ws.cell(row=r, column=12, value=int(v['allot']))
    ws.cell(row=r, column=13, value=int(v['pend']))
    ws.cell(row=r, column=14, value=(v['phys']/total_phys) if total_phys else 0)
    style(ws.cell(row=r, column=9), fill=BG_MID, fg=FG_WHITE, size=10, align='center')
    style(ws.cell(row=r, column=10), fill=BG_MID, fg=FG_WHITE, size=10, align='left')
    for col in (11,12,13):
        style(ws.cell(row=r, column=col), fill=BG_MID, fg=FG_WHITE, size=10, align='right', fmt='#,##0')
    style(ws.cell(row=r, column=14), fill=BG_MID, fg=FG_WHITE, size=10, align='right', fmt='0.0%')

# STOCK BY PACK SIZE subsection starts row 27 in template
ws.merge_cells("I27:N27"); ws["I27"] = "STOCK BY PACK SIZE"
style(ws["I27"], fill=BG_MID, fg=FG_GREEN, bold=True, size=11, align='left')

headers_pack = [('I28','#'),('J28','Pack'),('K28','Physical'),('L28','Allotable'),('M28','Pending'),('N28','% of Total')]
for coord,val in headers_pack:
    ws[coord] = val
    style(ws[coord], fill=BG_ACCENT, fg=FG_WHITE, bold=True, size=10, align='center')
ws.row_dimensions[28].height = 22

pack_sorted = sorted(pack_totals.items(), key=lambda kv: -kv[1]['phys'])
for i,(pk, v) in enumerate(pack_sorted):
    r = 29 + i
    ws.cell(row=r, column=9, value=i+1)
    ws.cell(row=r, column=10, value=pk)
    ws.cell(row=r, column=11, value=int(v['phys']))
    ws.cell(row=r, column=12, value=int(v['allot']))
    ws.cell(row=r, column=13, value=int(v['pend']))
    ws.cell(row=r, column=14, value=(v['phys']/total_phys) if total_phys else 0)
    style(ws.cell(row=r, column=9), fill=BG_MID, fg=FG_WHITE, size=10, align='center')
    style(ws.cell(row=r, column=10), fill=BG_MID, fg=FG_WHITE, size=10, align='left')
    for col in (11,12,13):
        style(ws.cell(row=r, column=col), fill=BG_MID, fg=FG_WHITE, size=10, align='right', fmt='#,##0')
    style(ws.cell(row=r, column=14), fill=BG_MID, fg=FG_WHITE, size=10, align='right', fmt='0.0%')

# STOCK HEALTH INDICATORS row 35 (leave 34 blank if fewer packs)
ws.merge_cells("I35:N35"); ws["I35"] = "STOCK HEALTH INDICATORS"
style(ws["I35"], fill=BG_MID, fg=FG_GREEN, bold=True, size=11, align='left')

indicators = [
    ("Network Allotable %", f"{avg_allot_pct:.1f}%", "allotable / physical (all WH)"),
    (f"Warehouses ≥ 80% allotable", strong_count, f"of {n_wh} warehouses"),
    (f"Warehouses < 50% allotable", weak_count, "needs attention"),
    ("Total SKU × Warehouse rows", sku_rows, f"{len(unique_brands)} unique brands"),
    ("SKUs with Allotable ≤ 5", sku_low_allot, "low dispatch readiness"),
    ("SKUs with zero Allotable", sku_zero_allot, "physical stock blocked"),
]
for i,(lbl, val, note) in enumerate(indicators):
    r = 36 + i
    ws.cell(row=r, column=9, value=lbl)
    style(ws.cell(row=r, column=9), fill=BG_MID, fg=FG_WHITE, size=10, align='left')
    # merge I:K for label
    ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=11)
    # L = value
    ws.cell(row=r, column=12, value=val)
    style(ws.cell(row=r, column=12), fill=BG_MID, fg=FG_CYAN, bold=True, size=11, align='center')
    # M:N = note
    ws.merge_cells(start_row=r, start_column=13, end_row=r, end_column=14)
    ws.cell(row=r, column=13, value=note)
    style(ws.cell(row=r, column=13), fill=BG_MID, fg=FG_WHITE, size=10, align='left')
    ws.row_dimensions[r].height = 20

# Fill dark background everywhere within A1:O70
for r in range(1, 71):
    for col in range(1, 16):
        c = ws.cell(row=r, column=col)
        if c.fill.patternType is None or c.fill.start_color.rgb in (None,):
            c.fill = PatternFill("solid", fgColor=BG_DARK)

# Freeze + hide gridlines
ws.sheet_view.showGridLines = False
ws.sheet_view.zoomScale = 100

# Save
out = FOLDER / output_filename
# If workbook file exists and is locked, fail
lockfile = FOLDER / f"~${output_filename}"
if lockfile.exists():
    print(f"ERROR_LOCK: {lockfile}")
    sys.exit(2)

# Save to scratch first, then atomic move
import os as _os, tempfile as _tempfile
_scratch_fd, _scratch_path = _tempfile.mkstemp(prefix="wh_scratch_", suffix=".xlsx")
_os.close(_scratch_fd)
scratch = Path(_scratch_path)
wb.save(scratch)

# Overwrite target (overwrite yesterday is expected per CLAUDE.md rules — snapshot workbook)
import shutil
shutil.copy2(scratch, out)
# Clean up scratch file so /tmp doesn't accumulate orphaned wh_scratch_*.xlsx
# across runs (the original PID-6 PermissionError was a stale scratch left over
# from a prior run owned by nobody:nogroup; the mkstemp() switch made paths
# unique, but unlinking is what actually prevents accumulation long-term).
try:
    scratch.unlink()
except OSError:
    pass
print(f"SAVED: {out}")

# ---------- Append to history ----------
HIST_DIR.mkdir(exist_ok=True)
# Two history files: per-day totals + per-warehouse per-day
# We keep 1 CSV with wh-level rows
new_rows = []
for wh, v in wh_totals.items():
    new_rows.append([report_iso, wh, int(v['phys']), int(v['allot']), int(v['pend'])])

# Re-run-safe write: drop any existing rows for report_iso, then write all prior
# dates + today. A same-day re-run (e.g. after re-exporting a warehouse that was
# missing from the first drop) now UPDATES the date instead of skipping it - the
# old append-only dedup froze the first, possibly-incomplete, set in history.
prior_hist_rows = []
had_date = False
if HIST_CSV.exists():
    with HIST_CSV.open() as f:
        r = csv.reader(f)
        for i, row in enumerate(r):
            if i == 0: continue
            if not row: continue
            if row[0] == report_iso:
                had_date = True
                continue
            prior_hist_rows.append(row)
# Belt and braces on top of the nothing-parsed guard above: an empty parse
# must never reach this rewrite, because the rewrite drops the date's existing
# rows before writing the new ones. Zero new rows there is not "a day with no
# stock" - it is a day erased.
if not new_rows:
    print(f"HIST_SKIPPED: nothing parsed for {report_iso}; history left untouched")
    sys.exit(4)

_tmp = HIST_CSV.with_suffix('.csv.tmp')
with _tmp.open('w', newline='') as f:
    w = csv.writer(f)
    w.writerow(['date','warehouse','physical','allotable','pending'])
    for row in prior_hist_rows:
        w.writerow(row)
    for row in new_rows:
        w.writerow(row)
os.replace(_tmp, HIST_CSV)
print(f"HIST_{'UPDATED' if had_date else 'APPENDED'}: {report_iso} ({len(new_rows)} warehouse rows)")

# ---------- Append brand × pack × warehouse history (added 10 Jun 2026, Abhay-requested) ----------
# Persists the full SKU-level detail (warehouse, brand, pack, product code,
# physical/allotable/pending) before the raws are deleted, so brand-by-warehouse
# questions are answerable any day, historically. Dedupe by date, same as
# stock_history.csv.
# Re-run-safe (same policy as stock_history): replace any existing rows for report_iso.
bp_prior_rows = []
bp_had_date = False
if BRAND_PACK_CSV.exists():
    with BRAND_PACK_CSV.open() as f:
        r = csv.reader(f)
        for i, row in enumerate(r):
            if i == 0: continue
            if not row: continue
            if row[0] == report_iso:
                bp_had_date = True
                continue
            bp_prior_rows.append(row)
bp_count = 0
_bp_tmp = BRAND_PACK_CSV.with_suffix('.csv.tmp')
with _bp_tmp.open('w', newline='') as f:
    w = csv.writer(f)
    w.writerow(['date','warehouse','brand','pack','product_code','physical','allotable','pending'])
    for row in bp_prior_rows:
        w.writerow(row)
    for wh, rows in warehouse_rows.items():
        for r_ in rows:
            w.writerow([report_iso, wh, r_['item'], r_['pack'], r_['code'],
                        r_['phys_case'], r_['allot_case'], r_['pend_case']])
            bp_count += 1
os.replace(_bp_tmp, BRAND_PACK_CSV)
print(f"BRAND_PACK_HIST_{'UPDATED' if bp_had_date else 'APPENDED'}: {report_iso} ({bp_count} SKU rows)")

# ---------- Maintain inbound_history.csv -----------
# Pulls Secondary sales COMBINED DISPATCHES, computes inbound = ΔΣphys + Σdispatched
# across the (prior_snapshot, today] window, appends per-WH rows. Idempotent.
# Failures are NON-FATAL so the dashboard build never aborts because of a
# Secondary-sales dependency.
try:
    import importlib.util as _ilu
    _spec = _ilu.spec_from_file_location(
        "maintain_inbound_history",
        Path(__file__).resolve().parent / "maintain_inbound_history.py",
    )
    _mod = _ilu.module_from_spec(_spec); _spec.loader.exec_module(_mod)
    _root = FOLDER.parent  # ".../Claude"
    _w, _s = _mod.maintain_inbound_history(_root, verbose=True)
    print(f"INBOUND_HIST: appended {_w} row(s), skipped {_s} date(s) with missing Secondary coverage")
except Exception as _e:
    print(f"INBOUND_HIST_ERROR (non-fatal): {_e}")

# ---------- Build change summary vs prior date ----------
all_rows = []
with HIST_CSV.open() as f:
    r = csv.reader(f)
    next(r, None)  # header
    for row in r:
        all_rows.append(row)

dates_in_hist = sorted(set(row[0] for row in all_rows))
prior_date = None
for d in reversed(dates_in_hist):
    if d < report_iso:
        prior_date = d; break

if prior_date:
    today_map = {row[1]: (int(row[2]), int(row[3]), int(row[4])) for row in all_rows if row[0]==report_iso}
    prior_map = {row[1]: (int(row[2]), int(row[3]), int(row[4])) for row in all_rows if row[0]==prior_date}
    # Like-for-like comparison: warehouses present in BOTH snapshots.
    # Warehouses missing from one side are coverage artefacts (raw not uploaded
    # that day), NOT real day-over-day stock movement — surface them separately.
    common_wh = set(today_map) & set(prior_map)
    today_only = sorted(set(today_map) - set(prior_map))
    prior_only = sorted(set(prior_map) - set(today_map))
    diffs = []
    for wh in common_wh:
        t = today_map[wh]
        p = prior_map[wh]
        diffs.append((wh, t[0]-p[0], t[0], p[0]))
    drained = sorted([d for d in diffs if d[1] < 0], key=lambda x: x[1])[:3]
    replenished = sorted([d for d in diffs if d[1] > 0], key=lambda x: -x[1])[:3]
    print(f"COMPARE_WITH: {prior_date}")
    print("CHANGE_SUMMARY:")
    print("  DRAINED:")
    for wh, d, t, p in drained:
        print(f"    - {wh}: {p}→{t} ({d:+d})")
    print("  REPLENISHED:")
    for wh, d, t, p in replenished:
        print(f"    + {wh}: {p}→{t} (+{d})")
    if today_only:
        print(f"  NEWLY_COVERED (in today only — not real stock movement): {', '.join(today_only)}")
    if prior_only:
        print(f"  DROPPED_COVERAGE (in {prior_date} only — raw missing today): {', '.join(prior_only)}")
else:
    print("COMPARE_WITH: NONE")

# ---------- Top warehouse and brand ----------
top_wh = wh_sorted[0]
top_brand = brands_sorted[0]
print(f"TOP_WAREHOUSE: {top_wh[0]} ({int(top_wh[1]['phys'])} cases)")
print(f"TOP_BRAND: {top_brand[0]} ({int(top_brand[1]['phys'])} cases)")

# A back-dated pull is normal and worth saying out loud, so a day filed
# under one date from an export printed on another is visible in the log
# rather than being something you have to know to look for.
if printed_seen and printed_seen != {report_date}:
    print(f"PULLED_ON: {', '.join(sorted(printed_seen))} (back-dated export; "
          f"filed under its Report Period, {report_date})")
print(f"REPORT_DATE: {report_date}")
print(f"REPORT_ISO: {report_iso}")
print(f"N_WH: {n_wh}")
print(f"AVG_ALLOT: {avg_allot_pct:.1f}")
print(f"STRONG: {strong_count}")
print(f"WEAK: {weak_count}")
print(f"ZERO_ALLOT_SKU: {sku_zero_allot}")
print(f"OUTPUT: {output_filename}")

# ---------- Delete raw files (workbook + history already persisted above) ----------
# Per Abhay's rule (22 Apr 2026): once the dashboard is rebuilt for the latest
# stock position, always delete the raw Bevco Report_*.xls drops. The history
# CSV is already appended at this point, so no data is lost.
# Interactive runs: sandbox can delete in-place after the user grants file-delete
# permission via mcp__cowork__allow_cowork_file_delete. Scheduled runs: delete
# is blocked — the launchd cleanup agent handles it at 23:00. Either way, try
# quietly; log what happened.
import os
deleted_count = 0
blocked_count = 0
for p in raw_files:
    try:
        os.remove(p)
        deleted_count += 1
    except PermissionError:
        blocked_count += 1
    except Exception as e:
        blocked_count += 1
        print(f"RAW_DELETE_ERROR: {p.name}: {e}")
if deleted_count:
    print(f"RAW_DELETED: {deleted_count} file(s)")
if blocked_count:
    # Expected branch in the normal SKILL.md flow: agent calls
    # allow_cowork_file_delete + rm AFTER the approval gate in step 5, and the
    # launchd safety-net runs at 23:00 for scheduled (autonomous) runs. Not a
    # script failure — log neutrally so it doesn't read as one.
    print(f"RAW_DELETE_DEFERRED: {blocked_count} file(s) (sandbox locked — agent will rm after approval, or launchd at 23:00 for scheduled runs)")
