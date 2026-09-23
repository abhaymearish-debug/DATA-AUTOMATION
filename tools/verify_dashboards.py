#!/usr/bin/env python3
"""Check both dashboards against the uploads they claim to be built from.

Run it after an upload, after a deploy, or whenever a figure looks wrong:

    python3 tools/verify_dashboards.py            # the app's own workspace
    python3 tools/verify_dashboards.py --root /data/workspace/mnt/Claude

It re-reads the raw exports with its OWN code - a second implementation, not a
call into the app's - and compares what it gets against what the dashboards
serve. That is the whole point: two implementations agreeing is evidence, one
implementation agreeing with itself is not.

Exit status is 0 when every check passes and 1 when any of them does not, so it
can be run unattended and its output read only when it complains.
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(HERE))

PASS, FAIL = "  ok  ", " FAIL "
_fails: list = []


def check(name: str, got, want, tol: float = 0.0, note: str = "") -> None:
    """One line of evidence: what the dashboard says, what the files say."""
    if isinstance(got, (int, float)) and isinstance(want, (int, float)):
        ok = abs(float(got) - float(want)) <= tol
        shown = f"{got:,.2f} vs {want:,.2f}"
    else:
        ok = got == want
        shown = f"{got!r} vs {want!r}"
    print(f"[{PASS if ok else FAIL}] {name:<46} {shown}{'  — ' + note if note else ''}")
    if not ok:
        _fails.append(name)


# ---------------------------------------------------------------------------
# The files, read independently
# ---------------------------------------------------------------------------
def read_history(path: Path) -> list:
    if not path.is_file():
        return []
    with path.open(newline="", encoding="utf-8-sig") as fh:
        return list(csv.DictReader(fh))


def num(v) -> float:
    try:
        return float(str(v).replace(",", "").strip() or 0)
    except ValueError:
        return 0.0


_DAY = re.compile(r"^([a-z]+)\s+(\d{1,2})(?:st|nd|rd|th)\.xlsx$", re.I)
_CUM = re.compile(r"([a-z]+)\s+(\d{1,2})\s*-\s*(\d{1,2})\s+CUMULATIVE\.xlsx$", re.I)
_SEC = re.compile(r"^RAW DATA\s*-\s*([a-z]+)\s+(\d{1,2})(?:st|nd|rd|th)\s+SECONDARY", re.I)


def flat(name) -> str:
    """A brand name with its punctuation and spacing taken out, so 'BLENDER'S
    CHOICE NO.1' and 'BLENDERS CHOICE NO 1' are the same brand. The exports do
    not write them the same way twice."""
    return re.sub(r"[^A-Z0-9]", "", str(name or "").upper())


def ksbc_cases(root: Path, month: str, brands: set) -> tuple:
    """(cases, last day covered) for a month's KSBC exports, period-true.

    Period-true here is written from the rule, not copied from the app: a
    block export supersedes the day files it covers, and blocks that contain
    one another are not added together.
    """
    from openpyxl import load_workbook
    folder = root / "KSBC shop sales"
    days, blocks = {}, []
    for f in folder.glob("*.xlsx"):
        m = _DAY.match(f.name)
        if m and m.group(1).upper() == month:
            days[int(m.group(2))] = f
    cum = folder / "_cumulative"
    if cum.is_dir():
        for f in cum.glob("*.xlsx"):
            m = _CUM.search(f.name)
            if m and m.group(1).upper() == month:
                blocks.append((int(m.group(2)), int(m.group(3)), f))

    # widest first; keep a block only if it covers a day nothing else does
    kept, covered = [], set()
    for a, b, f in sorted(blocks, key=lambda t: (t[0] - t[1], t[0])):
        span = set(range(a, b + 1))
        if span & covered:
            continue
        kept.append((a, b, f))
        covered |= span

    def sheet_cases(path: Path) -> float:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            tot = 0.0
            for r in ws.iter_rows(min_row=2, values_only=True):
                if not r or len(r) <= 12 or r[1] is None:
                    continue
                if flat(r[4]) not in brands:
                    continue
                try:
                    bpc = float(r[6] or 0)
                    tot += float(r[11] or 0) + (float(r[12] or 0) / bpc if bpc else 0)
                except (TypeError, ValueError):
                    continue
            return tot
        finally:
            wb.close()

    total = sum(sheet_cases(f) for _, _, f in kept)
    total += sum(sheet_cases(f) for d, f in sorted(days.items()) if d not in covered)
    last = max([0] + list(days) + [b for _, b, _ in kept])
    return total, last


# Bottles in a case, by pack. A physical fact about the bottling line, not a
# rule of the app's - written out again here so the verifier folds loose
# bottles into cases the same way without importing the code it is checking.
_BPC = {"1000 ML": 9, "750 ML": 12, "500 ML": 18, "375 ML": 24, "180 ML": 48}

_WH_ALIAS = {"PATHANAMTHITA": "PATHANAMTHITTA"}


def wh_name(raw) -> str:
    """'WH-KOLLAM FL9-KLM-01/2026-27' -> 'KOLLAM'.

    Written out here rather than imported, so a bug in the app's canonicaliser
    cannot hide behind itself. The licence code is spelt a dozen ways -
    'FL9-KLM-03', 'RFL9/PTA-02', 'FL09 NO 4/2026-27' - so rather than trying to
    match each of them, stop at the first token that has a digit or a slash in
    it or begins FL/RFL, which no warehouse name does.
    """
    txt = re.sub(r"\s+", " ", str(raw or "").strip().upper())
    if txt.startswith("WH-") or txt.startswith("WH "):
        txt = txt[3:].lstrip()
    keep = []
    for tok in txt.split():
        if any(c.isdigit() for c in tok) or "/" in tok \
                or tok.startswith("FL") or tok.startswith("RFL"):
            break
        keep.append(tok)
    nm = " ".join(keep).strip().rstrip(".") or txt
    return _WH_ALIAS.get(nm, nm)


def warehouse_invoices(root: Path, month: str, month_no: int) -> tuple:
    """{warehouse: cases} for a month, invoice-dated, straight off the exports.

    The month's ANALYSIS workbook where there is one, and the days it does not
    answer for filled from the raws - the same precedence the app uses, written
    again here so the two can disagree.
    """
    from openpyxl import load_workbook
    folder = root / "Secondary sales"
    if not folder.is_dir():
        return {}, 0
    files, claimed = [], set()
    full = sorted(folder.glob(f"{month} *SECONDARY SALES ANALYSIS.xlsx")) \
        + sorted(folder.glob(f"{month} SECONDARY SALES ANALYSIS.xlsx"))
    files += [(f, True) for f in full[-1:]]
    files += [(f, False) for f in sorted(folder.glob("RAW DATA*SECONDARY SALES.xlsx"))
              if _SEC.match(f.name) and _SEC.match(f.name).group(1).upper() == month]

    out: dict = defaultdict(float)
    last = 0
    for path, _is_book in files:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = None
            for nm in wb.sheetnames:
                if "COMBINED DISPATCH" in nm.upper():
                    ws = wb[nm]
                    break
            ws = ws or wb[wb.sheetnames[0]]
            got: dict = defaultdict(float)
            days = set()
            for r in ws.iter_rows(min_row=2, values_only=True):
                if not r or len(r) <= 12 or r[1] is None:
                    continue
                v = r[11]
                day = mon = None
                if hasattr(v, "day"):
                    day, mon = v.day, v.month
                else:
                    mm = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$", str(v or "").strip())
                    if mm:
                        day, mon = int(mm.group(1)), int(mm.group(2))
                if day is None or (mon and mon != month_no):
                    continue
                if day in claimed:
                    continue
                days.add(day)
                pack = re.sub(r"\s+", " ", str(r[5] or "").strip().upper())
                cases = num(r[12])
                bottles = num(r[13]) if len(r) > 13 else 0.0
                if bottles:
                    cases += bottles / _BPC.get(pack, 1)
                got[wh_name(r[1])] += cases
            for w, v in got.items():
                out[w] += v
            claimed |= days
            last = max([last] + list(days))
        finally:
            wb.close()
    return dict(out), last


def invoice_cases(root: Path, month: str, month_no: int, codes: set) -> tuple:
    """(cases, last invoice day) for a month's CFD/BAR exports."""
    from openpyxl import load_workbook
    folder = root / "Secondary sales"
    total, last = 0.0, 0
    # Whole cases are the whole story today. If an invoice ever carries loose
    # bottles they are NOT counted - the KSBC leg folds bottles in at the case
    # rate and this one does not - so the check below says so rather than
    # letting the difference ride.
    loose = [0.0]
    for f in sorted(folder.glob("*.xlsx")):
        m = _SEC.match(f.name)
        if not m or m.group(1).upper() != month:
            continue
        wb = load_workbook(f, read_only=True, data_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            for r in ws.iter_rows(min_row=2, values_only=True):
                if not r or len(r) <= 12 or r[6] is None:
                    continue
                try:
                    code = int(r[6])
                except (TypeError, ValueError):
                    continue
                if code not in codes:
                    continue
                day = mon = None
                v = r[11]
                if hasattr(v, "day"):
                    day, mon = v.day, v.month
                else:
                    mm = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$", str(v or "").strip())
                    if mm:
                        day, mon = int(mm.group(1)), int(mm.group(2))
                if day is None or (mon and mon != month_no):
                    continue
                total += num(r[12])
                loose[0] += num(r[13]) if len(r) > 13 else 0.0
                last = max(last, day)
        finally:
            wb.close()
    invoice_cases.loose = loose[0]
    return total, last


# ---------------------------------------------------------------------------
# The checks
# ---------------------------------------------------------------------------
def warehouse(root: Path) -> None:
    from app import reports_api

    print("\nWAREHOUSE DASHBOARD — what it serves vs what the uploads say")
    data = reports_api.warehouse_dashboard()
    if not data["stock"]:
        print("  no stock history in this workspace - nothing to check")
        return

    stock = read_history(root / "Warehouse stock" / "_history" / "stock_history.csv")
    bp = read_history(root / "Warehouse stock" / "_history" / "brand_pack_history.csv")

    latest = max(r["date"] for r in stock)
    check("as-of date is the newest stock report", data["asOf"], latest)

    rows = [r for r in stock if r["date"] == latest]
    served_whs = data["warehouses"]
    check("warehouses on that date",
          len(served_whs) if isinstance(served_whs, list) else served_whs, len(rows))
    for key, col in (("p", "physical"), ("a", "allotable"), ("n", "pending")):
        served = sum(r[key] for r in data["stock"] if r["d"] == latest)
        check(f"network {col} on {latest}", served, sum(num(r[col]) for r in rows), tol=len(rows))

    if bp:
        bp_latest = max(r["date"] for r in bp)
        if bp_latest == latest:
            check("brand/pack physical reconciles to stock",
                  sum(num(r["physical"]) for r in bp if r["date"] == latest),
                  sum(num(r["physical"]) for r in rows), tol=len(rows),
                  note="two files, same total")

    # A warehouse under two spellings is a whole warehouse missing from one
    # half of the page. Every key the page joins on has to exist on both sides.
    stock_whs = {r["w"] for r in data["stock"] if r["d"] == latest}
    sales_whs = set(data["sales"]["byWarehouse"])
    orphans = sorted(sales_whs - stock_whs)
    check("every warehouse with sales has a stock row", len(orphans), 0,
          note=("no stock row for: " + ", ".join(orphans[:4])) if orphans else
               f"{len(sales_whs)} of {len(stock_whs)} warehouses sold this window")

    clusters = data.get("clusters") or {}
    unmapped = sorted(w for w in stock_whs if w not in clusters)
    check("every warehouse belongs to a cluster", len(unmapped), 0,
          note=("unmapped: " + ", ".join(unmapped[:4])) if unmapped else "")
    if clusters:
        cl_phys = defaultdict(float)
        for r in data["stock"]:
            if r["d"] == latest:
                cl_phys[clusters.get(r["w"], 0)] += r["p"]
        check("the three clusters add to the network",
              sum(v for k, v in cl_phys.items() if k in (1, 2, 3)),
              sum(r["p"] for r in data["stock"] if r["d"] == latest), tol=0.5,
              note="a warehouse outside all three is a hole in every cluster card")

    # ---- dispatch, read back out of the invoices themselves ----------------
    # The dashboard's sales feed is the Secondary uploads, dated by Inv/GTN
    # date. inbound_history.csv is NOT checked here and must not be: the page
    # stopped reading it on 23 Sep 2026, and a check against a file nothing
    # serves can only ever produce a red line about a figure nobody sees.
    name = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY",
            "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]
    sales = data["sales"]
    for mon in sales["monthsUsed"]:
        month_no = name.index(mon) + 1
        mine, _last = warehouse_invoices(root, mon, month_no)
        served = sum(v["months"].get(mon, 0) for v in sales["byWarehouse"].values())
        check(f"{mon.title()} dispatch across the network", served, sum(mine.values()),
              tol=1.0, note="invoice-dated, read independently")
        drift = sorted(
            (abs(sales["byWarehouse"].get(w, {}).get("months", {}).get(mon, 0) - v), w)
            for w, v in mine.items())
        check(f"{mon.title()} dispatch warehouse by warehouse",
              drift[-1][0] if drift else 0, 0, tol=1.0,
              note=(f"worst: {drift[-1][1]}") if drift else "")

    # The demand each months-of-cover figure divides by: cases per uploaded
    # day, over a month of average length. A window with every day uploaded
    # must give back exactly the mean monthly total.
    cov = sales.get("coverage") or {}
    days_cov = sum(c["days"] for c in cov.values())
    days_win = sum(c["inMonth"] for c in cov.values())
    check("demand covers the window it names", days_cov, sales.get("daysCovered", 0), tol=0)
    if days_cov:
        mean_len = days_win / max(1, len(cov))
        bad = [w for w, v in sales["byWarehouse"].items()
               if abs(v["avg"] - round(v["total"] / days_cov * mean_len)) > 1]
        check("monthly demand is the uploaded-day rate over an average month",
              len(bad), 0,
              note=("offenders: " + ", ".join(bad[:4])) if bad else
                   f"{days_cov} of {days_win} days uploaded")

    # Three panels print a month-to-date sales figure. They are one number.
    dd = data.get("dailyDispatch") or {}
    cur = dd.get("cur")
    if cur:
        by_day = sum(cur["days"].values())
        by_wh = sum(sum(v.values()) for v in cur["byWh"].values())
        check("per-warehouse dispatch adds to the network", by_wh, by_day, tol=0.5,
              note="the header tile, the cluster cards and the daily panel read these")
        ch = sum(cur["ksbc"].values()) + sum(cur["inv"].values())
        check("the KSBC / invoice split adds to the month", ch, by_day, tol=0.5)
        if clusters:
            per_cl = defaultdict(float)
            for w, days in cur["byWh"].items():
                per_cl[clusters.get(w, 0)] += sum(days.values())
            check("the three clusters add to month-to-date sales",
                  sum(v for k, v in per_cl.items() if k in (1, 2, 3)), by_day, tol=0.5)
        gap = sorted(set(range(1, max(cur["covered"] or [0]) + 1)) - set(cur["covered"] or []))
        if gap:
            print(f"[ note ] {cur['m'].title()} has no upload for day(s) "
                  f"{', '.join(map(str, gap[:8]))} - the page says so on its face")
    else:
        print("[ note ] no Secondary Sales upload for the current month; the page "
              "says 'not uploaded' rather than showing a zero")

    cover_den = [w for w, v in sales["byWarehouse"].items() if v["avg"] <= 0]
    if cover_den:
        print(f"[ note ] {len(cover_den)} warehouse(s) have no dispatch in the window, so they "
              f"carry no months-of-cover: {', '.join(sorted(cover_den)[:4])}")


def liquidation(root: Path) -> None:
    from app import liq_live, liq_extract

    print("\nLIQUIDATION DASHBOARD — what it serves vs what the uploads say")
    try:
        bundle = liq_live.dashboard_bundle(root)
    except Exception as exc:                                    # noqa: BLE001
        print(f"  cannot build: {type(exc).__name__}: {exc}")
        _fails.append("liquidation bundle")
        return

    master = liq_extract.load_master(str(root / "MASTER DATA CONFIRMED.xlsx"))
    brands = {flat(b) for b in liq_extract.BRAND_COLS}
    inv_codes = {c for c, m in master.items()
                 if m["active"] and m["type"] in ("CFD", "BAR")}
    months = liq_live.MONTHS_UP

    for key in bundle["data"]:
        p = bundle["data"][key]
        month_no = months.index(key) + 1
        print(f"\n  {key.title()}")
        ks, ks_last = ksbc_cases(root, key, brands)
        inv, inv_last = invoice_cases(root, key, month_no, inv_codes)

        check(f"{key.title()}: KSBC tertiary", p["byType"].get("KSBC", 0), ks, tol=1.0,
              note="period-true, bottles folded at BPC")
        check(f"{key.title()}: CFD + BAR invoice",
              p["byType"].get("CFD", 0) + p["byType"].get("BAR", 0), inv, tol=1.0)
        check(f"{key.title()}: no loose bottles going uncounted",
              getattr(invoice_cases, "loose", 0), 0, tol=0.5,
              note="invoices are whole cases; bottles would be dropped")
        check(f"{key.title()}: total liquidation", p["grand"], ks + inv, tol=2.0)

        covered = max(ks_last, inv_last)
        check(f"{key.title()}: day N of the month covers the exports",
              p["daysElapsed"] >= covered, True,
              note=f"says day {p['daysElapsed']}, exports reach day {covered}")
        # Each leg over the days that leg answers for, then added. The two do
        # not cover the same window: KSBC arrives in period exports and stops
        # at the last block's end, invoices at the last invoice raised. One
        # divisor over both spreads a 23-day leg across 29 days.
        bs = p.get("basis") or {}
        want = 0.0
        if bs.get("ksbcTo"):
            want += bs["ksbcCases"] / bs["ksbcTo"]
        if bs.get("invTo"):
            want += bs["invCases"] / bs["invTo"]
        check(f"{key.title()}: run-rate is each leg over its own days",
              p["runRate"], round(want, 2), tol=0.02,
              note=f"KSBC to {bs.get('ksbcTo')}, invoices to {bs.get('invTo')}")
        check(f"{key.title()}: projection is that rate over the month",
              p["projected"], round(p["runRate"] * p["daysInMonth"], 2), tol=0.02)
        if not bs.get("ksbcDaily", True):
            print(f"[ note ] {key.title()}: the KSBC leg came in period exports "
                  f"({', '.join(f'{a}-{b}' for a, b in bs.get('ksbcBlocks', []))}), so it has "
                  f"no daily shape - the trend chart says so and plots the invoice leg alone")
        daily = sum(d["ksbc"] + d["inv"] for d in p["daily"])
        if daily > p["grand"] + 1:
            check(f"{key.title()}: daily trend does not exceed the total", daily, p["grand"], tol=1.0)
        else:
            print(f"[{PASS}] {key.title() + ': daily trend within the total':<46} "
                  f"{daily:,.2f} of {p['grand']:,.2f}"
                  + ("  — days inside a period export carry no daily shape"
                     if daily < p["grand"] - 1 else ""))

        shares = p.get("shopStock")
        if shares:
            net = shares["net"]
            by_bond = sum(v["closing"] for v in shares["byBond"].values())
            check(f"{key.title()}: shop stock bonds sum to the network",
                  by_bond, net["closing"], tol=1.0)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--root", default=None,
                    help="workspace to check (default: the app's own)")
    ap.add_argument("--skip", choices=["warehouse", "liquidation"], action="append", default=[])
    args = ap.parse_args()

    from app import config
    root = Path(args.root) if args.root else config.CLAUDE_ROOT
    print(f"Checking {root}")

    if "warehouse" not in args.skip:
        warehouse(root)
    if "liquidation" not in args.skip:
        liquidation(root)

    print()
    if _fails:
        print(f"{len(_fails)} check(s) failed: " + "; ".join(_fails))
        return 1
    print("Every check passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
