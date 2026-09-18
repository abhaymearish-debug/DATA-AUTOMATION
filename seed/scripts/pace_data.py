"""
Daily pace — data layer.

Builds the canonical per-day, per-bond, per-shop TOTAL LIQUIDATION history that
the daily-target system runs on.

    total liquidation = KSBC tertiary (shop -> consumer)
                      + CFD invoice sale (warehouse -> Consumer Fed outlet)
                      + BAR invoice sale (warehouse -> BAR outlet)

Same basis as `Targets/<MONTH> TARGETS/TARGET vs ACHIEVEMENT - <MONTH> <YEAR>.xlsx`
(verified 27 Jul 2026 against the July workbook), so daily and monthly speak the
same language.

Sources
  KSBC tertiary  : KSBC shop sales/<MONTH>*ANALYSIS.xlsx  daily raw sheets
                   ("JULY 1", "JULY 3-5", ...). CUMULATIVE / COMBINED sheets are
                   NEVER scanned -- they are period roll-ups, not days.
  CFD + BAR      : Secondary sales/<MONTH>*SECONDARY SALES ANALYSIS.xlsx
                   sheet "<MONTH> COMBINED DISPATCHES", rows whose licensee is
                   CAT=FED or CAT=BAR in master, keyed on Inv/GTN Date.

Output cache: Targets/.pace/daily_liquidation.csv
    date, bond, shop_code, cat, tertiary_cs, invoice_cs
"""

from __future__ import annotations

import csv
import datetime as _dt
import os
import re
import sys
from collections import defaultdict

import openpyxl

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

try:
    from ksbc_alias_guard import canonical_code as _canon
except Exception:                                            # pragma: no cover
    def _canon(code):
        return code


# --------------------------------------------------------------------------
# constants
# --------------------------------------------------------------------------

MONTHS = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
          "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]
MONTH_NUM = {m: i + 1 for i, m in enumerate(MONTHS)}

# bottles per case, by pack size (ml)
BPC = {180: 48, 375: 24, 500: 18, 750: 12, 1000: 9, 2000: 6}

MASTER_SHEET = "16-4-25"

_DAY_SHEET_RE = re.compile(r"^([A-Z]+)\s+(\d+)(?:\s*[-–]\s*(\d+))?$")

# A raw drop occasionally covers 2-3 consecutive days (a weekend batch). Those
# are genuine daily exports and spreading them evenly is defensible. Anything
# wider ("MARCH 1-16", "MARCH 17-31") is a PERIOD roll-up carrying no day shape
# -- including it would manufacture a flat, fake day curve.
MAX_RANGE_DAYS = 3
_ORD_RE = re.compile(r"^\s*(\d+)\s*(?:st|nd|rd|th)?\s*$", re.IGNORECASE)


def _default_base() -> str:
    """Walk up from this script to find the Claude root."""
    p = os.path.abspath(_HERE)
    for _ in range(6):
        p = os.path.dirname(p)
        if os.path.isfile(os.path.join(p, "MASTER DATA CONFIRMED.xlsx")):
            return p
    return os.path.expanduser("~/Downloads/Claude")


# --------------------------------------------------------------------------
# master data
# --------------------------------------------------------------------------

def load_master(base: str) -> dict:
    """shop_code(int) -> {name, cat, bond, staff, status, warehouse}"""
    path = os.path.join(base, "MASTER DATA CONFIRMED.xlsx")
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[MASTER_SHEET]
    out = {}
    header_seen = False
    for row in ws.iter_rows(values_only=True):
        if not header_seen:
            if row and str(row[0] or "").strip().lower().startswith("sl"):
                header_seen = True
            continue
        if not row or row[3] in (None, ""):
            continue
        try:
            code = _canon(int(str(row[3]).strip()))
        except (TypeError, ValueError):
            continue
        out[code] = {
            "warehouse": str(row[2] or "").strip(),
            "name": str(row[4] or "").strip(),
            "cat": str(row[5] or "").strip().upper(),
            "staff": str(row[6] or "").strip(),
            "bond": str(row[7] or "").strip().upper(),
            "status": str(row[8] or "").strip().title(),
        }
    wb.close()
    if not out:
        raise RuntimeError(f"master data empty or unreadable: {path}")
    return out


# --------------------------------------------------------------------------
# workbook discovery
# --------------------------------------------------------------------------

def _find_month_workbook(folder: str, month: str, suffix: str) -> str | None:
    """Prefer the clean full-month file, else the newest mid-month variant."""
    if not os.path.isdir(folder):
        return None
    mu = month.upper()
    full, mids = None, []
    for fn in os.listdir(folder):
        if not fn.lower().endswith(".xlsx") or fn.startswith("~$"):
            continue
        up = fn.upper()
        if not up.startswith(mu + " ") or suffix not in up:
            continue
        if up == f"{mu} {suffix}.XLSX".upper():
            full = fn
        else:
            mids.append(fn)
    if full:
        return os.path.join(folder, full)
    if mids:
        mids.sort(key=lambda f: os.path.getmtime(os.path.join(folder, f)))
        return os.path.join(folder, mids[-1])
    return None


def available_months(base: str) -> list[str]:
    """Months that have a KSBC analysis workbook on disk, in calendar order."""
    folder = os.path.join(base, "KSBC shop sales")
    found = set()
    if os.path.isdir(folder):
        for fn in os.listdir(folder):
            if fn.startswith("~$") or not fn.lower().endswith(".xlsx"):
                continue
            head = fn.split(" ")[0].upper()
            if head in MONTH_NUM:
                found.add(head)
    return [m for m in MONTHS if m in found]


# --------------------------------------------------------------------------
# KSBC tertiary — daily raw sheets
# --------------------------------------------------------------------------

def _day_sheets(sheetnames, month: str):
    """Yield (sheet_name, [day, ...]) for real daily sheets only."""
    mu = month.upper()
    for sn in sheetnames:
        up = sn.strip().upper()
        if "CUMULATIVE" in up or "COMBINED" in up:
            continue
        m = _DAY_SHEET_RE.match(up)
        if not m or m.group(1) != mu:
            continue
        d0 = int(m.group(2))
        d1 = int(m.group(3)) if m.group(3) else d0
        if d1 < d0:
            d0, d1 = d1, d0
        if d1 - d0 + 1 > MAX_RANGE_DAYS:
            continue                    # period roll-up, not a day
        yield sn, list(range(d0, d1 + 1))


def read_tertiary(base: str, month: str, year: int, master: dict) -> dict:
    """(date, shop_code) -> tertiary cases sold."""
    path = _find_month_workbook(os.path.join(base, "KSBC shop sales"),
                                month, "ANALYSIS")
    if not path:
        return {}
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    mno = MONTH_NUM[month.upper()]
    out = defaultdict(float)
    for sn, days in _day_sheets(wb.sheetnames, month):
        ws = wb[sn]
        per_shop = defaultdict(float)
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r or r[1] in (None, ""):
                continue
            try:
                code = _canon(int(str(r[1]).strip()))
            except (TypeError, ValueError):
                continue
            try:
                bpc = float(r[6] or 0)
                cs = float(r[11] or 0)
                btl = float(r[12] or 0)
            except (TypeError, ValueError):
                continue
            qty = cs + (btl / bpc if bpc else 0.0)
            if qty:
                per_shop[code] += qty
        # a range sheet ("JULY 3-5") covers N days -- spread it evenly, which is
        # the only defensible split when the export doesn't break it out.
        n = len(days)
        for d in days:
            try:
                date = _dt.date(year, mno, d)
            except ValueError:
                continue
            for code, qty in per_shop.items():
                out[(date, code)] += qty / n
    wb.close()
    return dict(out)


# --------------------------------------------------------------------------
# CFD + BAR invoice sales — secondary combined dispatches
# --------------------------------------------------------------------------

def _pack_bpc(pack) -> float:
    m = re.search(r"(\d+)", str(pack or ""))
    if not m:
        return 0.0
    return float(BPC.get(int(m.group(1)), 0))


def _parse_date(v):
    if isinstance(v, _dt.datetime):
        return v.date()
    if isinstance(v, _dt.date):
        return v
    s = str(v or "").strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y"):
        try:
            return _dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def read_invoice(base: str, month: str, master: dict) -> dict:
    """(date, shop_code) -> FED/BAR invoice cases dispatched."""
    path = _find_month_workbook(os.path.join(base, "Secondary sales"),
                                month, "SECONDARY SALES ANALYSIS")
    if not path:
        return {}
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    target = None
    for sn in wb.sheetnames:
        if "COMBINED DISPATCHES" in sn.upper():
            target = sn
            break
    if target is None:
        wb.close()
        return {}
    ws = wb[target]

    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    if not header:
        wb.close()
        return {}
    idx = {str(h or "").strip().lower(): i for i, h in enumerate(header)}

    def col(*names):
        for n in names:
            if n in idx:
                return idx[n]
        return None

    c_lic = col("licensee no.", "licensee no", "licensee")
    c_date = col("inv/gtn date", "inv/gtn  date", "invoice date")
    c_cs = col("issue cases", "issue  cases")
    c_btl = col("issue bottles", "issue  bottles")
    c_pack = col("pack", "packing")
    if None in (c_lic, c_date, c_cs):
        wb.close()
        raise RuntimeError(f"{os.path.basename(path)}: COMBINED DISPATCHES "
                           f"missing required columns {header}")

    out = defaultdict(float)
    for r in rows:
        if not r or r[c_lic] in (None, ""):
            continue
        try:
            code = _canon(int(str(r[c_lic]).strip()))
        except (TypeError, ValueError):
            continue
        info = master.get(code)
        if not info or info["cat"] not in ("FED", "BAR"):
            continue                    # KSBC leg is the dispatch pipeline, not a sale
        date = _parse_date(r[c_date])
        if date is None:
            continue
        try:
            qty = float(r[c_cs] or 0)
        except (TypeError, ValueError):
            qty = 0.0
        if c_btl is not None and c_pack is not None:
            try:
                btl = float(r[c_btl] or 0)
            except (TypeError, ValueError):
                btl = 0.0
            bpc = _pack_bpc(r[c_pack])
            if btl and bpc:
                qty += btl / bpc
        if qty:
            out[(date, code)] += qty
    wb.close()
    return dict(out)


# --------------------------------------------------------------------------
# authoritative month actuals (BOND PERFORMANCE, cumulative-corrected)
# --------------------------------------------------------------------------

def month_actuals(base: str, month: str) -> dict:
    """
    bond -> {"tertiary", "invoice", "total"} for the month, read from the
    BOND PERFORMANCE sheets.

    This -- not the sum of daily sheets -- is the source of truth for a bond's
    month-to-date achievement. KSBC's portal only exports in fixed 1-16 and
    17-end windows, and case-conversion rounding inside those exports means
    summing dailies drifts a few tenths of a percent from the period total
    (~38 cs on July's 5,800). The daily series carries the SHAPE; this carries
    the LEVEL. See CLAUDE.md "Cumulative-period rule".
    """
    out = defaultdict(lambda: {"tertiary": 0.0, "invoice": 0.0, "total": 0.0})

    kp = _find_month_workbook(os.path.join(base, "KSBC shop sales"),
                              month, "ANALYSIS")
    if kp:
        wb = openpyxl.load_workbook(kp, data_only=True, read_only=True)
        if "BOND PERFORMANCE" in wb.sheetnames:
            for r in wb["BOND PERFORMANCE"].iter_rows(values_only=True):
                name = str(r[0] or "").strip().upper()
                if (not name or name.startswith("CLUSTER") or name == "TOTAL"
                        or name == "BOND" or "BOND-WISE" in name):
                    continue
                try:
                    out[name]["tertiary"] = float(r[3] or 0)
                except (TypeError, ValueError):
                    pass
        wb.close()

    sp = _find_month_workbook(os.path.join(base, "Secondary sales"),
                              month, "SECONDARY SALES ANALYSIS")
    if sp:
        wb = openpyxl.load_workbook(sp, data_only=True, read_only=True)
        if "BOND PERFORMANCE" in wb.sheetnames:
            for r in wb["BOND PERFORMANCE"].iter_rows(values_only=True):
                name = str(r[0] or "").strip().upper()
                if (not name or name == "TOTAL" or name == "BOND"
                        or "BOND-WISE" in name):
                    continue
                try:                    # cols: KSBC | Consumer fed | BAR | Total
                    out[name]["invoice"] = float(r[2] or 0) + float(r[3] or 0)
                except (TypeError, ValueError):
                    pass
        wb.close()

    for v in out.values():
        v["total"] = v["tertiary"] + v["invoice"]
    return dict(out)


# --------------------------------------------------------------------------
# history build
# --------------------------------------------------------------------------

def build_history(base: str | None = None, year: int = 2026,
                  months: list[str] | None = None, verbose: bool = True) -> list[dict]:
    base = base or _default_base()
    master = load_master(base)
    months = months or available_months(base)

    recs: dict[tuple, dict] = {}
    for month in months:
        ter = read_tertiary(base, month, year, master)
        inv = read_invoice(base, month, master)
        for (date, code), qty in ter.items():
            k = (date, code)
            recs.setdefault(k, {"tertiary": 0.0, "invoice": 0.0})["tertiary"] += qty
        for (date, code), qty in inv.items():
            if date.month != MONTH_NUM[month]:
                continue                # stray out-of-month row in the sheet
            k = (date, code)
            recs.setdefault(k, {"tertiary": 0.0, "invoice": 0.0})["invoice"] += qty
        if verbose:
            t = sum(v for (d, _), v in ter.items())
            i = sum(v for (d, _), v in inv.items() if d.month == MONTH_NUM[month])
            print(f"  {month:<10} tertiary {t:9.1f} cs   invoice {i:8.1f} cs   "
                  f"total {t + i:9.1f} cs")

    rows = []
    unknown = defaultdict(float)
    for (date, code), v in sorted(recs.items()):
        info = master.get(code)
        if not info:
            unknown[code] += v["tertiary"] + v["invoice"]
            continue
        rows.append({
            "date": date.isoformat(),
            "bond": info["bond"],
            "shop_code": code,
            "cat": info["cat"],
            "status": info["status"],
            "tertiary_cs": round(v["tertiary"], 4),
            "invoice_cs": round(v["invoice"], 4),
        })
    if unknown and verbose:
        print(f"  ! {len(unknown)} shop code(s) not in master "
              f"({sum(unknown.values()):.1f} cs) -- excluded, investigate: "
              f"{sorted(unknown)[:8]}")
    return rows


def write_history(rows: list[dict], base: str | None = None) -> str:
    base = base or _default_base()
    outdir = os.path.join(base, "Targets", ".pace")
    os.makedirs(outdir, exist_ok=True)
    path = os.path.join(outdir, "daily_liquidation.csv")
    cols = ["date", "bond", "shop_code", "cat", "status", "tertiary_cs", "invoice_cs"]
    with open(path, "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)
    return path


if __name__ == "__main__":
    base = _default_base()
    print(f"base: {base}")
    print(f"months on disk: {', '.join(available_months(base))}")
    rows = build_history(base)
    p = write_history(rows, base)
    print(f"wrote {len(rows):,} rows -> {p}")
