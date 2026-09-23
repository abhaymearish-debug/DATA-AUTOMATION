"""TOTAL LIQUIDATION figures for the liquidation dashboard.

Total liquidation = KSBC tertiary shop sales + CFD invoice sales + BAR invoice
sales (all three legs, in cases), per Abhay's definition.

This is the liquidation-live artifact's own builder, brought into the app: the
refresh script used to bake the numbers into the artifact's HTML every evening
because an artifact iframe could not call back for them. Here the page asks the
app, so the same functions run per request instead - the arithmetic below is
the script's, unchanged, and that is the point. Headline totals come from the
AUTHORITATIVE roll-up sheets (KSBC '... COMBINED', Secondary 'COMBINED
DISPATCHES') via the shared liq_extract extractors - the same source of truth
as TOTAL LIQUIDATION -<MONTH>.xlsx. The daily trend is computed from the KSBC
daily raw sheets (Shop Out) + Secondary dispatches by Inv/GTN date; it is the
SHAPE of liquidation over the month and may differ from the COMBINED headline
by a fraction of a percent (KSBC case-conversion rounding).

The month the dashboard opens on is auto-detected from the newest KSBC analysis
workbook, so it rolls over on its own.
"""
from __future__ import annotations

import calendar
import json
import logging
import re
import threading
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from . import config
from . import liq_extract as B

log = logging.getLogger("ksd.liquidation")


class Unavailable(RuntimeError):
    """A month cannot be built - a missing workbook, usually. The endpoint
    turns this into a 404 the page can show, where the script it came from
    would have exited."""


def print(*args):                                   # noqa: A001
    """The script narrates its progress on stdout. In a server that belongs in
    the log, and keeping the name means the body below stays the script's."""
    log.info(" ".join(str(a) for a in args))


MONTHS_UP = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE',
             'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']
MONTH_TITLE = ['January', 'February', 'March', 'April', 'May', 'June',
               'July', 'August', 'September', 'October', 'November', 'December']
# Short brand labels for charts (keyed by the full BRAND_COLS name).
SHORT = {
    'BCB NO.1 CLASSIC BRANDY': 'BCB',
    "BLENDER'S CHOICE NO.1 BRANDY": "Blender's Choice",
    "CHAIRMAN'S CHOICE XO BRANDY": "Chairman's Choice",
    'K.S 99 LIFE TIME MATURED XXX RUM': 'K.S 99',
    'MAGIC BLEND RESERVED XXX RUM': 'Magic Blend',
    "MORNING WALKER'S XO BRANDY": 'Morning Walker',
    'OLD PEARL NO.1 MATURED XXX RUM': 'Old Pearl',
    'ROYAL OLD FORT NO.1 XXX RUM': 'Royal Old Fort',
}


PACK_LABEL = {180: '180ml', 375: '375ml', 500: '500ml',
              750: '750ml', 1000: '1000ml'}


def norm_pack(v):
    m = re.search(r'(\d+)', str(v or ''))
    if not m:
        return None
    return PACK_LABEL.get(int(m.group(1)), m.group(1) + 'ml')


def claude_root() -> Path:
    """The workspace the app serves from, not a path relative to this file."""
    return config.CLAUDE_ROOT


# --- month / file discovery --------------------------------------------------
def detect_current_month(root: Path) -> str:
    """Newest KSBC analysis workbook by mtime -> its month (auto-roll)."""
    folder = root / "KSBC shop sales"
    best = None
    for f in folder.glob("*ANALYSIS.xlsx"):
        if f.name.startswith("~$"):
            continue
        up = f.name.upper()
        mon = next((m for m in MONTHS_UP if up.startswith(m + " ")), None)
        if not mon:
            continue
        mt = f.stat().st_mtime
        if best is None or mt > best[0]:
            best = (mt, mon)
    if best is None:
        raise Unavailable("No KSBC analysis workbook found to detect the month.")
    return best[1]


def window_end_day(root: Path, month: str, ksbc_path=None):
    """Parse the '... 1st - <N>th ...' window end from the KSBC mid-month
    filename; None for a full-month workbook (use max daily date instead).

    MUST be derived from the workbook that find_analysis_file() actually
    returned. This previously globbed independently and took the FIRST match,
    so a superseded mid-month workbook left on disk set days_elapsed and the
    period label while the data came from a different file -- silently
    inflating run-rate and projection (a stale 'JUNE 1st - 10th' file, even
    zero-byte, produced 841 cs/day against a true 280).
    """
    folder = root / "KSBC shop sales"
    mids = [f for f in folder.glob(f"{month} 1st - *ANALYSIS.xlsx")
            if not f.name.startswith("~$")]
    if len(mids) > 1:
        print(f"WARNING: {len(mids)} mid-month {month} workbooks on disk "
              f"({', '.join(sorted(f.name for f in mids))}). "
              f"Superseded files should be deleted -- see CLAUDE.md cleanup rule.")
    if ksbc_path:
        m = re.search(r"1st\s*-\s*(\d+)", Path(ksbc_path).name)
        return int(m.group(1)) if m else None
    return None


# --- per-shop rows via the locked extractors ---------------------------------
def build_rows(root: Path, month: str, master, active_fed, active_bar):
    """Returns (rows, ksbc_path, sec_path) or (None, ...) if a source is absent."""
    ksbc_path, _ = B.find_analysis_file(
        str(root / "KSBC shop sales"), month,
        "SHOP SALES ANALYSIS.xlsx", "ANALYSIS.xlsx")
    sec_path, _ = B.find_analysis_file(
        str(root / "Secondary sales"), month,
        "SECONDARY SALES ANALYSIS.xlsx", "SECONDARY SALES ANALYSIS.xlsx")
    if not ksbc_path or not sec_path:
        return None, ksbc_path, sec_path
    ksbc_sales, _ = B.load_ksbc_brand_sales(ksbc_path)
    sec_sales, _ = B.load_secondary_brand_sales(sec_path, active_fed, active_bar)
    rows = []
    for code, m in master.items():
        if not m['active'] or m['type'] not in {'KSBC', 'CFD', 'BAR'}:
            continue
        brands = dict(ksbc_sales.get(code, {})) if m['type'] == 'KSBC' else dict(sec_sales.get(code, {}))
        brands = {k: round(float(v or 0), 2) for k, v in brands.items()}
        rows.append({
            'code': code, 'name': m['name'], 'staff': m['staff'],
            'bond': m['bond'], 'type': m['type'], 'brands': brands,
            'total': round(sum(brands.values()), 2),
        })
    return rows, ksbc_path, sec_path


def aggregate(rows):
    by_type = defaultdict(float)
    by_bond = defaultdict(float)
    by_bond_shops = defaultdict(int)
    by_brand = defaultdict(float)
    by_staff = defaultdict(float)
    staff_bond = {}
    grand = 0.0
    active_outlets = 0
    for r in rows:
        by_type[r['type']] += r['total']
        by_bond[r['bond']] += r['total']
        by_staff[r['staff']] += r['total']
        staff_bond.setdefault(r['staff'], r['bond'])
        if r['total'] > 0:
            by_bond_shops[r['bond']] += 1
            active_outlets += 1
        for b, c in r['brands'].items():
            by_brand[b] += c
        grand += r['total']
    return {
        'byType': {k: round(v, 2) for k, v in by_type.items()},
        'byBond': by_bond, 'byBondShops': by_bond_shops,
        'byBrand': by_brand, 'byStaff': by_staff, 'staffBond': staff_bond,
        'grand': round(grand, 2), 'activeOutlets': active_outlets,
    }


# --- daily trend -------------------------------------------------------------
def ksbc_daily(root: Path, month: str, master=None):
    """{day:int -> cases} KSBC tertiary Shop Out (cases + bottles/bpc) per day.

    With `master`, also returns a per-bond split of the SAME rows as
    `.byBond` -> {bond: {day: cases}} (26 Jul 2026, Abhay: daily-trend
    drill-downs). Built from the identical row scan, so every bond series
    sums back to the network series exactly."""
    ksbc_path, _ = B.find_analysis_file(
        str(root / "KSBC shop sales"), month,
        "SHOP SALES ANALYSIS.xlsx", "ANALYSIS.xlsx")
    if not ksbc_path:
        return {}
    from openpyxl import load_workbook
    wb = load_workbook(ksbc_path, read_only=True, data_only=True)
    pat = re.compile(rf"^{month}\s+(\d+)$", re.I)
    out = {}
    by_bond = defaultdict(lambda: defaultdict(float))
    for sn in wb.sheetnames:
        m = pat.match(sn.strip())
        if not m:
            continue
        day = int(m.group(1))
        ws = wb[sn]
        tot = 0.0
        for row in ws.iter_rows(min_row=2, values_only=True):
            try:
                bpc = float(row[6] or 0)
                oc = float(row[11] or 0)
                ob = float(row[12] or 0) if len(row) > 12 and row[12] is not None else 0.0
            except (TypeError, ValueError, IndexError):
                continue
            cs = oc + (ob / bpc if bpc else 0.0)
            tot += cs
            if master is not None and cs:
                try:
                    code = int(row[1])
                except (TypeError, ValueError, IndexError):
                    continue
                bond = (master.get(code, {}) or {}).get('bond') or 'UNKNOWN'
                by_bond[bond][day] += cs
        out[day] = round(tot, 2)
    wb.close()
    out = dict(out)
    if master is not None:
        out['byBond'] = {b: {d: round(v, 2) for d, v in dd.items()}
                         for b, dd in by_bond.items()}
    return out


def secondary_daily(root: Path, month: str, fed, bar, master=None):
    """{day:int -> cases} CFD+BAR Issue Cases per Inv/GTN date.

    With `master`, also returns `.byBond` -> {bond: {day: cases}} from the
    same rows (26 Jul 2026) for the daily-trend drill-downs."""
    sec_path, _ = B.find_analysis_file(
        str(root / "Secondary sales"), month,
        "SECONDARY SALES ANALYSIS.xlsx", "SECONDARY SALES ANALYSIS.xlsx")
    if not sec_path:
        return {}
    from openpyxl import load_workbook
    wb = load_workbook(sec_path, read_only=True, data_only=True)
    cands = [s for s in wb.sheetnames if s.upper().rstrip().endswith('COMBINED DISPATCHES')]
    if not cands:
        wb.close()
        return {}
    ws = max((wb[s] for s in cands), key=lambda w: w.max_row)
    target = fed | bar
    out = defaultdict(float)
    by_bond = defaultdict(lambda: defaultdict(float))
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 13 or row[6] is None:
            continue
        try:
            code = int(row[6])
        except (TypeError, ValueError):
            continue
        if code not in target:
            continue
        day = parse_day(row[11])
        if day is None:
            continue
        try:
            cases = float(row[12]) if row[12] is not None else 0.0
        except (TypeError, ValueError):
            cases = 0.0
        out[day] += cases
        if master is not None and cases:
            bond = (master.get(code, {}) or {}).get('bond') or 'UNKNOWN'
            by_bond[bond][day] += cases
    wb.close()
    res = {d: round(v, 2) for d, v in out.items()}
    if master is not None:
        res['byBond'] = {b: {d: round(v, 2) for d, v in dd.items()}
                         for b, dd in by_bond.items()}
    return res


def parse_day(val):
    """Day-of-month int from a date cell that may be a datetime or a string
    like '05-06-2026' / '2026-06-05' / '5/6/2026'."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.day
    s = str(val).strip()
    m = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$", s)
    if m:
        return int(m.group(1))  # DD-MM-YYYY (Kerala export convention)
    m = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$", s)
    if m:
        return int(m.group(3))  # YYYY-MM-DD
    return None


# --- same-window breakdown (for month-over-month to the same date) -----------
def window_breakdown(root: Path, month: str, master, fed, bar, day_max: int):
    """Liquidation restricted to days 1..day_max.

    AGGREGATES are period-true: where a `<MONTH> 1-16 CUMULATIVE` sheet exists
    it supersedes the sum of daily sheets for that block (KSBC rounding drift),
    exactly as the COMBINED roll-up does. Applied identically to BOTH months,
    so current vs prior stays apples-to-apples.
    PER-DAY series (dailyK/dailyI/dailyBrand/dailyPack) remain daily-sheet
    based -- the cumulative gives a block total, not a per-day shape."""
    from openpyxl import load_workbook
    res = {'grand': 0.0, 'byType': defaultdict(float), 'byBrand': defaultdict(float),
           'byBond': defaultdict(float), 'dailyK': defaultdict(float), 'dailyI': defaultdict(float),
           'byStaff': defaultdict(float), 'byShop': defaultdict(float),
           'byPack': defaultdict(float),
           'packType': defaultdict(lambda: [0.0, 0.0, 0.0]),
           'dailyPack': defaultdict(lambda: defaultdict(float)),
           'dailyBond': defaultdict(lambda: defaultdict(lambda: [0.0, 0.0])),
           'dailyBrand': defaultdict(lambda: defaultdict(float))}
    # KSBC tertiary (daily sheets)
    ksbc_path, _ = B.find_analysis_file(
        str(root / "KSBC shop sales"), month, "SHOP SALES ANALYSIS.xlsx", "ANALYSIS.xlsx")
    if ksbc_path:
        wb = load_workbook(ksbc_path, read_only=True, data_only=True)
        pat = re.compile(rf"^{month}\s+(\d+)$", re.I)
        # PERIOD TRUTH: KSBC's portal exports in fixed blocks (1-16, 17-EOM) and
        # summing the daily exports drifts from the block total by case-conversion
        # rounding. ALWAYS prefer a cumulative block sheet over the dailies it
        # covers -- for every block present, not just 1-16.
        #
        # A cumulative is a BLOCK TOTAL and cannot be sliced, so it is only
        # usable when the window fully covers its range. Comparing July 1-17 vs
        # June 1-17, June's 17-30 sheet is unusable (window ends at 17) and day
        # 17 correctly falls back to the daily sheet.
        cum_re = re.compile(rf"^{re.escape(month)}\s*(\d+)\s*-\s*(\d+)\s+CUMULATIVE$", re.I)
        cum_blocks = []
        for sn in wb.sheetnames:
            cm = cum_re.match(sn.strip())
            if cm:
                a, b = int(cm.group(1)), int(cm.group(2))
                if b <= day_max:                      # window fully covers the block
                    cum_blocks.append((a, b, sn))
        cum_days = set()
        for a, b, _ in cum_blocks:
            cum_days.update(range(a, b + 1))
        if cum_blocks:
            print(f"  {month} period-true blocks: "
                  + ", ".join(f"{a}-{b}" for a, b, _ in sorted(cum_blocks))
                  + (f" + dailies {max(cum_days)+1}-{day_max}"
                     if day_max > max(cum_days) else ""))

        def _row_vals(row):
            """-> (cases, brand, code, pack) or None if the row is unusable."""
            try:
                code = int(row[1])
                bpc = float(row[6] or 0)
                oc = float(row[11] or 0)
                ob = float(row[12] or 0) if len(row) > 12 and row[12] is not None else 0.0
            except (TypeError, ValueError, IndexError):
                return None
            cs = oc + (ob / bpc if bpc else 0.0)
            if cs == 0:
                return None
            bn = B.BRAND_LOOKUP.get(B.norm_brand(row[4]))
            if not bn:
                return None
            return cs, bn, code, norm_pack(row[5])

        def _add_agg(cs, bn, code, pk):
            mrec = master.get(code, {}) or {}
            res['byType']['KSBC'] += cs
            res['byBrand'][bn] += cs
            res['byBond'][mrec.get('bond') or 'UNKNOWN'] += cs
            res['byStaff'][mrec.get('staff') or '—'] += cs
            res['byShop'][code] += cs
            if pk:
                res['byPack'][pk] += cs
                res['packType'][pk][0] += cs
            res['grand'] += cs

        for sn in wb.sheetnames:
            m = pat.match(sn.strip())
            if not m:
                continue
            day = int(m.group(1))
            if day < 1 or day > day_max:
                continue
            # a covering cumulative block owns the aggregates for its days
            to_agg = day not in cum_days
            for row in wb[sn].iter_rows(min_row=2, values_only=True):
                v = _row_vals(row)
                if not v:
                    continue
                cs, bn, code, pk = v
                if to_agg:
                    _add_agg(cs, bn, code, pk)
                if pk:
                    res['dailyPack'][pk][day] += cs
                res['dailyBrand'][bn][day] += cs
                res['dailyK'][day] += cs
                res['dailyBond'][(master.get(code, {}) or {}).get('bond')
                                 or 'UNKNOWN'][day][0] += cs

        for _a, _b, _sn in cum_blocks:
            for row in wb[_sn].iter_rows(min_row=2, values_only=True):
                v = _row_vals(row)
                if not v:
                    continue
                _add_agg(*v)
        wb.close()
    # CFD + BAR invoice (dispatches by date)
    sec_path, _ = B.find_analysis_file(
        str(root / "Secondary sales"), month,
        "SECONDARY SALES ANALYSIS.xlsx", "SECONDARY SALES ANALYSIS.xlsx")
    if sec_path:
        wb = load_workbook(sec_path, read_only=True, data_only=True)
        cands = [s for s in wb.sheetnames if s.upper().rstrip().endswith('COMBINED DISPATCHES')]
        if cands:
            ws = max((wb[s] for s in cands), key=lambda w: w.max_row)
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or len(row) < 13 or row[6] is None:
                    continue
                try:
                    code = int(row[6])
                except (TypeError, ValueError):
                    continue
                typ = 'CFD' if code in fed else ('BAR' if code in bar else None)
                if typ is None:
                    continue
                day = parse_day(row[11])
                if day is None or day < 1 or day > day_max:
                    continue
                bn = B.BRAND_LOOKUP.get(B.norm_brand(row[4]))
                if not bn:
                    continue
                try:
                    cs = float(row[12] or 0)
                except (TypeError, ValueError):
                    cs = 0.0
                mrec = master.get(code, {}) or {}
                bond = mrec.get('bond') or 'UNKNOWN'
                res['byType'][typ] += cs
                res['byBrand'][bn] += cs
                res['byBond'][bond] += cs
                res['byStaff'][mrec.get('staff') or '—'] += cs
                res['byShop'][code] += cs
                pk = norm_pack(row[5])
                if pk:
                    res['byPack'][pk] += cs
                    res['packType'][pk][1 if typ == 'CFD' else 2] += cs
                    res['dailyPack'][pk][day] += cs
                res['dailyBrand'][bn][day] += cs
                res['dailyI'][day] += cs
                res['dailyBond'][bond][day][1] += cs
                res['grand'] += cs
        wb.close()
    return res


def pack_window(w, brand_cols, short, day_max):
    return {
        'grand': round(w['grand'], 2),
        'byType': {k: round(w['byType'].get(k, 0), 2) for k in ('KSBC', 'CFD', 'BAR')},
        'byBrand': {short[b]: round(w['byBrand'].get(b, 0), 2) for b in brand_cols},
        'byBond': {k: round(v, 2) for k, v in w['byBond'].items()},
        'byStaff': {k: round(v, 2) for k, v in w['byStaff'].items()},
        'byPack': {k: round(v, 2) for k, v in w['byPack'].items()},
        'daily': [{'day': d, 'ksbc': round(w['dailyK'].get(d, 0), 2),
                   'inv': round(w['dailyI'].get(d, 0), 2)} for d in range(1, day_max + 1)],
        # per-bond daily series for the trend drill-downs (26 Jul 2026) —
        # [ksbc, invoice] per day, same rows as the network series above
        'dailyBond': {b: [[d, round(dd.get(d, [0, 0])[0], 2),
                           round(dd.get(d, [0, 0])[1], 2)]
                          for d in range(1, day_max + 1)
                          if dd.get(d) and (dd[d][0] or dd[d][1])]
                      for b, dd in w['dailyBond'].items()},
    }


# --- payload -----------------------------------------------------------------
# --- stock in the shops ------------------------------------------------------
_SHORT_N = {re.sub(r"[^A-Z0-9]", "", k.upper()): v for k, v in SHORT.items()}


def shop_stock(root: Path, month: str, master):
    """Diagnostic view of the stock sitting in KSBC shops.

    Reads the <MONTH> COMBINED roll-up (shop x SKU) and, for the network and
    for every bond, computes: closing + sales by brand and pack, the brand x
    pack matrix, how many shops hold each brand, and -- the key diagnostic --
    NOT-MOVING stock: closing cases on shop x SKU lines that sold nothing at
    all this period. KSBC shops only (CFD/BAR are invoice channels).
    """
    ksbc_path, _ = B.find_analysis_file(
        str(root / "KSBC shop sales"), month,
        "SHOP SALES ANALYSIS.xlsx", "ANALYSIS.xlsx")
    if not ksbc_path:
        return None
    from openpyxl import load_workbook
    wb = load_workbook(ksbc_path, read_only=True, data_only=True)
    comb = next((sn for sn in wb.sheetnames if "COMBINED" in sn.upper()), None)
    if not comb:
        wb.close()
        return None
    it = wb[comb].iter_rows(min_row=1, values_only=True)
    hdr = next(it, None)
    if not hdr:
        wb.close()
        return None
    ix = {str(h).strip().lower(): i for i, h in enumerate(hdr) if h is not None}
    try:
        c_shop = ix["shop code"]; c_brand = ix["brand name"]
        c_pack = ix["packing"]; c_close = ix["closing (cases)"]
        c_sales = ix["sales (cases)"]
    except KeyError:
        wb.close()
        return None

    EPS = 0.005

    def blank():
        return {"closing": 0.0, "sales": 0.0, "stag": 0.0, "stagLines": 0,
                "shops": set(),
                "brand": defaultdict(lambda: [0.0, 0.0, 0.0, set()]),
                "pack": defaultdict(lambda: [0.0, 0.0]),
                "matrix": defaultdict(lambda: defaultdict(float))}

    net, by_bond, pack_ml = blank(), {}, {}
    for r in it:
        if not r or r[c_shop] is None:
            continue
        try:
            code = int(r[c_shop])
        except (TypeError, ValueError):
            continue
        m = master.get(code)
        if not m or m["type"] != "KSBC" or not m["active"]:
            continue
        try:
            cl = float(r[c_close] or 0); sa = float(r[c_sales] or 0)
        except (TypeError, ValueError):
            continue
        if cl <= EPS and sa <= EPS:
            continue
        raw = str(r[c_brand] or "").strip()
        brand = _SHORT_N.get(re.sub(r"[^A-Z0-9]", "", raw.upper()), raw or "-")
        mm = re.search(r"(\d+)", str(r[c_pack] or ""))
        ml = int(mm.group(1)) if mm else 0
        pack = PACK_LABEL.get(ml, f"{ml}ml" if ml else "-")
        pack_ml[pack] = ml
        bond = (m.get("bond") or "-").strip() or "-"
        holds = cl > EPS
        dead = holds and sa <= EPS          # stock present, nothing sold at all
        for d in (net, by_bond.setdefault(bond, blank())):
            d["closing"] += cl; d["sales"] += sa
            if holds:
                d["shops"].add(code)
            if dead:
                d["stag"] += cl; d["stagLines"] += 1
            br = d["brand"][brand]
            br[0] += cl; br[1] += sa
            if dead:
                br[2] += cl
            if holds:
                br[3].add(code)
            pk = d["pack"][pack]
            pk[0] += cl; pk[1] += sa
            d["matrix"][brand][pack] += cl
    wb.close()
    if net["closing"] <= 0:
        return None

    def out(d):
        return {"closing": round(d["closing"], 2), "sales": round(d["sales"], 2),
                "stag": round(d["stag"], 2), "stagLines": d["stagLines"],
                "shops": len(d["shops"]),
                "brand": {k: [round(v[0], 2), round(v[1], 2), round(v[2], 2), len(v[3])]
                          for k, v in d["brand"].items()},
                "pack": {k: [round(v[0], 2), round(v[1], 2)] for k, v in d["pack"].items()},
                "matrix": {b: {p: round(v, 2) for p, v in ps.items() if v > EPS}
                           for b, ps in d["matrix"].items()}}

    return {"brandOrder": sorted(net["brand"], key=lambda b: -net["brand"][b][0]),
            "packOrder": sorted(net["pack"], key=lambda p: pack_ml.get(p, 0)),
            "bondOrder": sorted(by_bond, key=lambda b: -by_bond[b]["closing"]),
            "net": out(net),
            "byBond": {b: out(v) for b, v in by_bond.items()}}


def build_payload(root: Path, month=None):
    master = B.load_master(str(root / "MASTER DATA CONFIRMED.xlsx"))
    active_fed = {c for c, m in master.items() if m['type'] == 'CFD' and m['active']}
    active_bar = {c for c, m in master.items() if m['type'] == 'BAR' and m['active']}

    cur = (month or detect_current_month(root)).strip().upper()
    cur_idx = MONTHS_UP.index(cur)
    year = datetime.now().year  # current working year

    rows, ksbc_path, sec_path = build_rows(root, cur, master, active_fed, active_bar)
    if rows is None:
        raise Unavailable(f"Missing analysis source for {cur}: "
                         f"KSBC={ksbc_path} Secondary={sec_path}")
    agg = aggregate(rows)

    # daily trend
    kd = ksbc_daily(root, cur, master)
    sd = secondary_daily(root, cur, active_fed, active_bar, master)
    kd_bond = kd.pop('byBond', {})
    sd_bond = sd.pop('byBond', {})
    all_days = sorted(set(kd) | set(sd))
    days_in_month = calendar.monthrange(year, cur_idx + 1)[1]
    win_end = window_end_day(root, cur, ksbc_path)
    days_elapsed = win_end or (max(all_days) if all_days else 0)
    daily = [{'day': d, 'ksbc': kd.get(d, 0.0), 'inv': sd.get(d, 0.0)}
             for d in range(1, days_elapsed + 1)]
    # per-bond daily series for the trend drill-downs (26 Jul 2026, Abhay) —
    # [day, ksbc, invoice], same rows that built `daily`, so every bond series
    # sums back to the network series exactly.
    daily_bond = {}
    for b in set(kd_bond) | set(sd_bond):
        kb, sb = kd_bond.get(b, {}), sd_bond.get(b, {})
        ser = [[d, round(kb.get(d, 0.0), 2), round(sb.get(d, 0.0), 2)]
               for d in range(1, days_elapsed + 1)
               if kb.get(d) or sb.get(d)]
        if ser:
            daily_bond[b] = ser

    # prior month (for pace / share momentum)
    prior = None
    if cur_idx - 1 >= 0:
        pmon = MONTHS_UP[cur_idx - 1]
        # A prior month that cannot be READ must not take the current month
        # down with it -- MARCH's Secondary workbook predates the COMBINED
        # DISPATCHES sheet, which used to make APRIL unbuildable even though
        # April's own two workbooks are perfectly good. Degrade to no prior,
        # exactly as a missing file already does. (31 Aug 2026)
        try:
            prows, pk, ps = build_rows(root, pmon, master, active_fed, active_bar)
        except Exception as e:                                  # noqa: BLE001
            print(f"WARNING: {pmon} unreadable ({type(e).__name__}: {e}) -- "
                  f"prior-month comparison suppressed for {cur}.")
            prows = None
        if prows is not None:
            pagg = aggregate(prows)
            pdays = calendar.monthrange(year, cur_idx)[1]
            prior = {
                'month': MONTH_TITLE[cur_idx - 1],
                'grand': pagg['grand'],
                'byType': pagg['byType'],
                'byBrand': {SHORT[b]: round(pagg['byBrand'].get(b, 0), 2) for b in B.BRAND_COLS},
                'byBond': {k: round(v, 2) for k, v in pagg['byBond'].items()},
                'daysInMonth': pdays,
                'runRate': round(pagg['grand'] / pdays, 2) if pdays else 0,
            }

    # same-window month-over-month comparison (current 1..N vs prior 1..N) ----
    cur_win = window_breakdown(root, cur, master, active_fed, active_bar, days_elapsed)
    compare = None
    movers = None
    if cur_idx - 1 >= 0:
        pmon = MONTHS_UP[cur_idx - 1]
        try:
            prior_win = window_breakdown(root, pmon, master, active_fed, active_bar, days_elapsed)
        except Exception as e:                                  # noqa: BLE001
            print(f"WARNING: {pmon} window unreadable ({type(e).__name__}: {e}) -- "
                  f"same-window comparison and movers suppressed for {cur}.")
            prior_win = {'grand': 0, 'byShop': {}}
        # BOTH prior legs must be present. window_breakdown treats each as
        # optional, so a missing prior KSBC or Secondary workbook used to yield
        # a one-legged 'prior' that still passed grand>0 -- rendering e.g.
        # MoM +264% instead of hiding the panel.
        _pk, _ps = B.find_analysis_file(
            str(root / "KSBC shop sales"), pmon,
            "SHOP SALES ANALYSIS.xlsx", "ANALYSIS.xlsx")[0], \
            B.find_analysis_file(
            str(root / "Secondary sales"), pmon,
            "SECONDARY SALES ANALYSIS.xlsx", "SECONDARY SALES ANALYSIS.xlsx")[0]
        if not (_pk and _ps):
            print(f"WARNING: {pmon} is incomplete (KSBC={bool(_pk)} "
                  f"Secondary={bool(_ps)}) -- month-over-month and movers suppressed.")
        elif prior_win['grand'] > 0:
            compare = {
                'window': days_elapsed,
                'curMonth': MONTH_TITLE[cur_idx],
                'priorMonth': MONTH_TITLE[cur_idx - 1],
                'cur': pack_window(cur_win, B.BRAND_COLS, SHORT, days_elapsed),
                'prior': pack_window(prior_win, B.BRAND_COLS, SHORT, days_elapsed),
            }
            # top outlet movers vs prior month, same window (daily basis)
            mv = []
            for code in set(cur_win['byShop']) | set(prior_win['byShop']):
                cu = round(cur_win['byShop'].get(code, 0.0), 2)
                pr = round(prior_win['byShop'].get(code, 0.0), 2)
                d = round(cu - pr, 2)
                if abs(d) < 0.005:
                    continue
                m = master.get(code, {}) or {}
                mv.append([code, m.get('name') or str(code), m.get('type') or '—',
                           m.get('bond') or '—', cu, pr, d])
            movers = {
                'up': sorted([r for r in mv if r[6] > 0], key=lambda x: -x[6])[:8],
                'down': sorted([r for r in mv if r[6] < 0], key=lambda x: x[6])[:8],
            }

    # ranked structures
    brand_sorted = sorted(((SHORT[b], round(agg['byBrand'].get(b, 0), 2))
                           for b in B.BRAND_COLS), key=lambda x: -x[1])
    bond_sorted = sorted(((bd, round(v, 2), agg['byBondShops'].get(bd, 0))
                          for bd, v in agg['byBond'].items()), key=lambda x: -x[1])
    staff_sorted = sorted(((st, round(v, 2), agg['staffBond'].get(st, ''))
                           for st, v in agg['byStaff'].items()), key=lambda x: -x[1])
    top_shops = sorted(rows, key=lambda r: -r['total'])[:15]
    top_shops_out = [[r['code'], r['name'], r['type'], r['bond'], r['total']]
                     for r in top_shops if r['total'] > 0]

    # bond x brand matrix (heatmap): rows=bonds (ranked), cols=brands (ranked)
    brand_order = [b for b, _ in brand_sorted]
    bond_order = [bd for bd, _, _ in bond_sorted]
    bb = defaultdict(lambda: defaultdict(float))
    for r in rows:
        for b, c in r['brands'].items():
            bb[r['bond']][SHORT[b]] += c
    matrix = [[round(bb[bd].get(b, 0), 2) for b in brand_order] for bd in bond_order]

    # coverage (active vs total outlets) by bond / type / staff -------------
    bond_cov = defaultdict(lambda: [0, 0])
    bond_cov_kc = defaultdict(lambda: [0, 0])   # KSBC + CFD only (excludes BAR)
    silent_outlets = defaultdict(list)          # bond -> silent KSBC/CFD outlets
    type_cov = {t: [0, 0] for t in ('KSBC', 'CFD', 'BAR')}
    staff_cov = defaultdict(lambda: [0, 0])
    for r in rows:
        bond_cov[r['bond']][1] += 1
        staff_cov[r['staff']][1] += 1
        is_kc = r['type'] in ('KSBC', 'CFD')
        if is_kc:
            bond_cov_kc[r['bond']][1] += 1
        if r['type'] in type_cov:
            type_cov[r['type']][1] += 1
        if r['total'] > 0:
            bond_cov[r['bond']][0] += 1
            staff_cov[r['staff']][0] += 1
            if is_kc:
                bond_cov_kc[r['bond']][0] += 1
            if r['type'] in type_cov:
                type_cov[r['type']][0] += 1
        elif is_kc:
            silent_outlets[r['bond']].append({'name': r['name'], 'type': r['type'], 'staff': r['staff']})

    # bond -> field staff (for the merged Bond Command Center column)
    bond_staff = defaultdict(set)
    for r in rows:
        if r['staff']:
            bond_staff[r['bond']].add(str(r['staff']))

    # Pareto: descending outlet totals (liquidating outlets only)
    shop_totals_desc = sorted((round(r['total'], 2) for r in rows if r['total'] > 0),
                              reverse=True)

    # Per-bond OUTLET breakdown for the bond-table drill-down (30 Jul 2026,
    # Abhay: "when I click on a bond I want the breakup of the outlets sales").
    # Month total per outlet is the same basis bondSorted sums, and the prior
    # figure is the prior month's SAME WINDOW — so the drill-down's two columns
    # add up to exactly the CASES and prior columns of the row it expands.
    # [code, name, type, month cases, prior-window cases, staff]
    try:
        prior_shop = prior_win['byShop']          # absent when there is no prior month
    except NameError:
        prior_shop = {}
    outlets_by_bond = defaultdict(list)
    for r in rows:
        outlets_by_bond[r['bond']].append([
            r['code'], r['name'], r['type'], round(r['total'], 2),
            round(prior_shop.get(r['code'], 0.0), 2), r['staff'] or '—'])
    for b in outlets_by_bond:
        outlets_by_bond[b].sort(key=lambda x: -x[3])
    outlets_by_bond = dict(outlets_by_bond)

    # brand x channel split (COMBINED basis, cases)
    bch = {SHORT[b]: [0.0, 0.0, 0.0] for b in B.BRAND_COLS}
    for r in rows:
        idx = 0 if r['type'] == 'KSBC' else (1 if r['type'] == 'CFD' else 2)
        for b, c in r['brands'].items():
            if b in SHORT:
                bch[SHORT[b]][idx] += c
    bch = {k: [round(x, 2) for x in v] for k, v in bch.items()}

    # per-brand daily shape for sparklines (daily-sheet basis)
    brand_daily = {SHORT[b]: [round(cur_win['dailyBrand'].get(b, {}).get(d, 0.0), 2)
                              for d in range(1, days_elapsed + 1)]
                   for b in B.BRAND_COLS}

    run_rate = round(agg['grand'] / days_elapsed, 2) if days_elapsed else 0
    projected = round(run_rate * days_in_month, 2)

    return {
        'month': MONTH_TITLE[cur_idx],
        'monthUpper': cur,
        'year': year,
        'periodLabel': f"{MONTH_TITLE[cur_idx]} 1–{days_elapsed}, {year}",
        'asOf': f"{year}-{cur_idx+1:02d}-{days_elapsed:02d}",
        'daysElapsed': days_elapsed,
        'daysInMonth': days_in_month,
        'grand': agg['grand'],
        'runRate': run_rate,
        'projected': projected,
        'byType': agg['byType'],
        'brandSorted': brand_sorted,
        'bondSorted': bond_sorted,
        'outletsByBond': outlets_by_bond,
        'staffSorted': staff_sorted,
        'topShops': top_shops_out,
        'daily': daily,
        'dailyBond': daily_bond,
        'matrix': {'bonds': bond_order, 'brands': brand_order, 'data': matrix},
        'totalOutlets': len(rows),
        'activeOutlets': agg['activeOutlets'],
        'prior': prior,
        'compare': compare,
        'monthIndex': cur_idx,
        'bondCoverage': dict(bond_cov),
        'bondCoverageKC': dict(bond_cov_kc),
        'silentOutlets': {b: sorted(v, key=lambda o: (o['type'], o['name'])) for b, v in silent_outlets.items()},
        'bondStaff': {b: ' / '.join(sorted(v)) for b, v in bond_staff.items()},
        'typeCoverage': type_cov,
        'staffCoverage': dict(staff_cov),
        'shopTotalsDesc': shop_totals_desc,
        'brandChannel': bch,
        'dailyBrand': brand_daily,
        'packMix': {k: round(v, 2) for k, v in cur_win['byPack'].items()},
        'packChannel': {k: [round(x, 2) for x in v] for k, v in cur_win['packType'].items()},
        'dailyPack': {k: [round(cur_win['dailyPack'][k].get(d, 0.0), 2)
                          for d in range(1, days_elapsed + 1)]
                      for k in cur_win['dailyPack']},
        'movers': movers,
        'targetDefault': prior['grand'] if prior else None,
        'shopStock': shop_stock(root, cur, master),
    }


# --- multi-month bundle (month filter, 31 Aug 2026, Abhay) -------------------
# The artifact ships EVERY month it has complete sources for, and the header
# month chips switch between them client-side. One payload per month is built
# by the same locked build_payload(), so a month the filter shows is exactly
# what a single-month build of that month would have produced.
#
# A month costs ~30 s to extract, so payloads are CACHED against the size+mtime
# of every workbook that feeds them (that month's KSBC + Secondary, the prior
# month's two -- which drive `compare`/`movers` -- and master). A daily refresh
# therefore rebuilds only the live month and reuses the closed ones; touch any
# source and its month rebuilds on its own.
CACHE_PATH = "Monthly statement-liquidation/.artifact/liq_month_cache.json"


def month_sources(root: Path, month: str):
    """The two analysis workbooks that feed `month` (either may be None)."""
    k, _ = B.find_analysis_file(str(root / "KSBC shop sales"), month,
                                "SHOP SALES ANALYSIS.xlsx", "ANALYSIS.xlsx")
    s, _ = B.find_analysis_file(str(root / "Secondary sales"), month,
                                "SECONDARY SALES ANALYSIS.xlsx",
                                "SECONDARY SALES ANALYSIS.xlsx")
    return k, s


def available_months(root: Path):
    """Months with BOTH legs on disk, in calendar order."""
    out = []
    for m in MONTHS_UP:
        k, s = month_sources(root, m)
        if k and s:
            out.append(m)
    return out


def month_signature(root: Path, month: str) -> str:
    """size+mtime fingerprint of everything build_payload() reads for `month`."""
    idx = MONTHS_UP.index(month)
    paths = list(month_sources(root, month))
    if idx > 0:
        paths += list(month_sources(root, MONTHS_UP[idx - 1]))
    paths.append(str(root / "MASTER DATA CONFIRMED.xlsx"))
    parts = []
    for p in paths:
        if not p:
            parts.append("-")
            continue
        try:
            st = Path(p).stat()
            parts.append(f"{Path(p).name}:{st.st_size}:{int(st.st_mtime)}")
        except OSError:
            parts.append("-")
    return "|".join(parts)


def build_bundle(root: Path, months=None, current=None, use_cache=True):
    """{'current': KEY, 'months': [...], 'data': {KEY: payload}}"""
    months = months or available_months(root)
    cur = (current or detect_current_month(root)).strip().upper()
    if cur not in months:
        months = sorted(set(months) | {cur}, key=MONTHS_UP.index)

    cache_file = root / CACHE_PATH
    cache = {}
    if use_cache and cache_file.exists():
        try:
            cache = json.loads(cache_file.read_text())
        except Exception as e:                                  # noqa: BLE001
            print(f"  cache unreadable ({e}) -- rebuilding all months.")
            cache = {}

    data, meta, fresh = {}, [], {}
    for m in months:
        sig = month_signature(root, m)
        hit = cache.get(m)
        if use_cache and hit and hit.get("sig") == sig and hit.get("payload"):
            data[m] = hit["payload"]
            print(f"  {m:<10} cached")
        else:
            try:
                data[m] = build_payload(root, m)
                print(f"  {m:<10} built   {data[m]['grand']:>10,.2f} cs")
            except Exception as e:                              # noqa: BLE001
                # An older month can be unbuildable for reasons that say nothing
                # about today's data -- e.g. MARCH's Secondary workbook predates
                # the COMBINED DISPATCHES sheet. Drop it from the filter and
                # carry on; only the month the artifact OPENS on is fatal.
                if m == cur:
                    raise
                print(f"  {m:<10} SKIPPED — {type(e).__name__}: {e}")
                continue
        fresh[m] = {"sig": sig, "payload": data[m]}

    if not data:
        raise Unavailable("No month could be built.")
    if cur not in data:
        cur = max(data, key=MONTHS_UP.index)

    if use_cache:
        try:
            cache_file.parent.mkdir(parents=True, exist_ok=True)
            cache_file.write_text(json.dumps(fresh, separators=(",", ":")))
        except OSError as e:                                    # noqa: BLE001
            print(f"  WARNING: cache not written ({e})")

    for m in sorted(data, key=MONTHS_UP.index):
        p = data[m]
        meta.append({
            "key": m,
            "name": p["month"],
            "short": p["month"][:3],
            "year": p["year"],
            "label": f"{p['month'][:3]} {p['year']}",
            "grand": p["grand"],
            "days": p["daysElapsed"],
            "daysInMonth": p["daysInMonth"],
            "partial": p["daysElapsed"] < p["daysInMonth"],
            "live": m == cur,
        })
    return {"current": cur, "months": meta, "data": data}



# ---------------------------------------------------------------------------
# What the dashboard asks for
# ---------------------------------------------------------------------------
# build_bundle() reads a JSON cache beside the workbooks, so a month whose
# sources have not changed is not rebuilt. That still means opening and parsing
# that cache on every request, and it is megabytes; this keeps the built bundle
# in memory and re-reads only when a source workbook's size or mtime moves.
_MEMO: dict = {}
# One build at a time. Two requests arriving on a cold cache would otherwise
# read the same workbooks twice over, each slowing the other down.
_LOCK = threading.Lock()


def dashboard_bundle(root: Path | None = None) -> dict:
    """{'current': KEY, 'months': [...], 'data': {KEY: payload}} for the page."""
    root = Path(root) if root else claude_root()
    cur = detect_current_month(root)
    months = sorted(set(available_months(root)) | {cur}, key=MONTHS_UP.index)
    stamp = "|".join(month_signature(root, m) for m in months)

    hit = _MEMO.get(str(root))
    if hit and hit[0] == stamp:
        return hit[1]

    with _LOCK:
        # Whoever waited here usually finds the bundle the request ahead of
        # them just built.
        hit = _MEMO.get(str(root))
        if hit and hit[0] == stamp:
            return hit[1]
        bundle = build_bundle(root, months=months, current=cur)
        _MEMO[str(root)] = (stamp, bundle)
    return bundle


def sources_note(root: Path | None = None) -> str:
    """What the builder can actually see, for when it cannot build.

    A dashboard that says only "no workbook" leaves you opening folders to
    find out which one. This names both legs and what is in them, because the
    answer is nearly always that one month's pair is half there.
    """
    root = Path(root) if root else claude_root()
    bits = []
    for label, pat in (("KSBC shop sales", "*ANALYSIS.xlsx"),
                       ("Secondary sales", "*SECONDARY SALES ANALYSIS.xlsx")):
        folder = root / label
        if not folder.is_dir():
            bits.append(f"{label}: no such folder")
            continue
        got = sorted(f.name for f in folder.glob(pat)
                     if not f.name.startswith("~$"))
        if got:
            bits.append(f"{label}: {', '.join(got[:4])}"
                        + (f" (+{len(got) - 4} more)" if len(got) > 4 else ""))
        else:
            other = len([f for f in folder.glob("*.xlsx")
                         if not f.name.startswith("~$")])
            bits.append(f"{label}: no month workbook"
                        + (f", {other} other file(s)" if other else ", empty"))
    return " · ".join(bits)


def warm() -> None:
    """Read the workbooks while the server is coming up.

    A month costs half a minute of openpyxl the first time and nothing
    afterwards, and the person who opens the dashboard after a deploy should
    not be the one paying for all of them.
    """
    try:
        dashboard_bundle()
    except Exception as exc:                                    # noqa: BLE001
        log.info("liquidation warm-up skipped: %s: %s", type(exc).__name__, exc)
