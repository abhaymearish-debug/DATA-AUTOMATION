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
    inbound = read_history(root / "Warehouse stock" / "_history" / "inbound_history.csv")
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

    # The flow columns have to account for the stock itself: warehouse by
    # warehouse, everything that came in less everything that went out is the
    # movement in physical stock. Exact on sound data; a gap means one of those
    # columns is not the window's figure, and the sales and inbound tiles are
    # reading it.
    per: dict = defaultdict(lambda: defaultdict(
        lambda: {"in": 0.0, "out": 0.0, "first": None, "start": 0.0, "last": None, "end": 0.0}))
    for r in inbound:
        mon = r["date"][:7]
        w = per[mon][r["warehouse"].strip().upper()]
        w["in"] += num(r["inbound_cases"])
        w["out"] += num(r["dispatched_cases"])
        if w["first"] is None or r["date"] < w["first"]:
            w["first"], w["start"] = r["date"], num(r["phys_start"])
        if w["last"] is None or r["date"] > w["last"]:
            w["last"], w["end"] = r["date"], num(r["phys_end"])
    for mon in sorted(per)[-4:]:
        whs = per[mon]
        net = sum(w["in"] - w["out"] for w in whs.values())
        moved = sum(w["end"] - w["start"] for w in whs.values())
        check(f"{mon}: flows account for the stock movement", net, moved,
              tol=max(2.0 * len(whs), 1.0),
              note=f"{sum(w['in'] for w in whs.values()):,.0f} in, "
                   f"{sum(w['out'] for w in whs.values()):,.0f} out")

    # the same row twice under two spellings of one warehouse doubles a month
    seen = defaultdict(set)
    clashes = []
    for r in inbound:
        canon = re.sub(r"[^A-Z]", "", r["warehouse"].upper())
        key = (r["date"], canon)
        if r["warehouse"] in seen[key]:
            continue
        seen[key].add(r["warehouse"])
        if len(seen[key]) > 1:
            clashes.append(f"{r['date']} {sorted(seen[key])}")
    check("one row per warehouse per day", len(clashes), 0,
          note=("; ".join(clashes[:3])) if clashes else "no repeated days")

    # monthly dispatch, and the average behind every months-of-cover figure
    sales = data["sales"]
    months = {m: defaultdict(float) for m in sales["monthsUsed"]}
    name = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY",
            "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]
    for r in inbound:
        mon = name[int(r["date"][5:7]) - 1]
        if mon in months:
            months[mon][r["warehouse"].strip().upper()] += num(r["dispatched_cases"])
    for mon in sales["monthsUsed"]:
        served = sum(v["months"].get(mon, 0) for v in sales["byWarehouse"].values())
        check(f"{mon.title()} dispatch across the network", served,
              sum(months[mon].values()), tol=len(sales["byWarehouse"]))

    bad = [w for w, v in sales["byWarehouse"].items()
           if v["avg"] != round(v["total"] / max(1, sum(
               1 for mm in sales["monthsUsed"] if v["months"].get(mm, 0) > 0) or len(months)))]
    check("monthly average divides by months with sales", len(bad), 0,
          note=("offenders: " + ", ".join(bad[:4])) if bad else "")

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
        check(f"{key.title()}: run-rate is the total over those days",
              p["runRate"], round(p["grand"] / p["daysElapsed"], 2) if p["daysElapsed"] else 0,
              tol=0.02)
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
