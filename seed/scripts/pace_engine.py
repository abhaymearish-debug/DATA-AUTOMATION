"""
Daily pace — engine.

One target line, kept deliberately simple, and it is SHOP LIQUIDATION ONLY:

    monthly shop target = official monthly bond target x SHOP_FACTOR (0.80)
    daily target        = monthly shop target / selling days

FED/BAR invoice was dropped from the target on 27 Jul 2026 (Abhay). It still
counts as liquidation and the team will be told so, but executives were
leaning on invoice instead of working the shops -- and because invoice arrives
as an event rather than a flow, it made the daily number unfair anyway (72% of
green days in July were green only because a truck was invoiced that day).
Shop sales accrue every day, so a flat daily number is honest on this basis
and the per-day verdict means something again.

The official target is set on TOTAL liquidation, so it is scaled to a shop
equivalent by a flat SHOP_FACTOR. Abhay chose flat over per-bond (27 Jul 2026)
because it is trivially explainable to the team -- "your shop target is 80% of
your bond target" needs no arithmetic.

The trade-off, recorded so it is not rediscovered later: shop share of
liquidation actually runs from 61% (TRIPUNITHURA, KOZHIKODE) to 99%
(PERINTHALMANNA), so a flat factor asks the invoice-heavy bonds for far more
on the shop leg than the shop-heavy ones. Scaling by each bond's own share was
built and measured -- it tightens the ask from -70%..+150% to +17%..+96% -- and
is one line away in shop_targets() if the meeting says the spread is unfair.

The exec is told one number ("sell 27 cases a day"). Every day he lands under
it, the shortfall rolls into what is left, so the CATCH-UP number climbs:

    needed per day now = (monthly target - sold so far) / days left

That climbing number is the whole point -- it puts the pressure on from day 1
instead of letting it surface on day 25.

The only thing that is NOT flat is dry days. The 1st of every month is dry in
Kerala, and extra dry days happen (4 May, 26 Jun 2026 were both network-zero).
A day the shops were shut carries no target, and its share is spread over the
days that remain. Telling someone to sell 27 cases on a closed day would get
the whole tracker dismissed on day one.
"""

from __future__ import annotations

import calendar
import datetime as _dt
import os
import sys
from collections import defaultdict

import openpyxl

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import pace_data as PD


# --- tunables --------------------------------------------------------------

DRY_DAY_OF_MONTH = 1        # Kerala: the 1st is dry every month
SHOP_FACTOR = 0.80          # share of the official bond target treated as the
                            # SHOP target (Abhay, 27 Jul 2026). Flat by choice.
EXTRA_UPLIFT = 0.00         # further stretch on top; left at zero because the
                            # stretch is already inside the official targets.
BASE_MONTHS = 3             # completed months used for the shop-share window

BONDS_ORDER = [
    "KOLLAM", "ALAPPUZHA", "NEDUMANGAD", "ATTINGAL", "KOTTARAKARA",
    "PATHANAMTHITTA", "KOTTAYAM", "THODUPUZHA", "TRIPUNITHURA", "THRISSUR",
    "ALUVA", "PALAKKAD", "KOZHIKODE", "KANNUR", "PERINTHALMANNA",
]

CLUSTERS = {
    1: ["ALAPPUZHA", "ATTINGAL", "NEDUMANGAD", "KOLLAM", "KOTTARAKARA",
        "PATHANAMTHITTA"],
    2: ["THRISSUR", "TRIPUNITHURA", "THODUPUZHA", "KOTTAYAM", "ALUVA"],
    3: ["KANNUR", "KOZHIKODE", "PALAKKAD", "PERINTHALMANNA"],
}
BOND_CLUSTER = {b: c for c, bs in CLUSTERS.items() for b in bs}

# Area Sales Managers, one per cluster. Roster from ASM_ROSTER in the aging /
# field-visit streams: DINESH KUMAR P, SOJAN T T, HARIDASAN M. Cluster
# membership follows the LOCKED definition in CLAUDE.md (BOND PERFORMANCE
# clusterwise), which is why PATHANAMTHITTA sits in Cluster 1 here -- the older
# guess in build_daily_briefing.py puts it in Cluster 2 and is self-declared as
# "no explicit mapping found". If master ever gains a real ASM column, read it.
ASM_BY_CLUSTER = {1: "Dinesh", 2: "Sojan", 3: "Haridasan"}


# --- monthly targets -------------------------------------------------------

def read_official_targets(base: str, month: str) -> dict:
    """bond -> official monthly bond target (cases, TOTAL liquidation)."""
    folder = os.path.join(base, "Targets", f"{month.upper()} TARGETS")
    if not os.path.isdir(folder):
        raise FileNotFoundError(
            f"no target folder for {month.upper()} -- expected {folder}. "
            f"Drop the month's target workbook there first.")
    cands = [f for f in os.listdir(folder)
             if f.lower().endswith(".xlsx") and not f.startswith("~$")
             and "ACHIEVEMENT" not in f.upper() and "PACE" not in f.upper()]
    if not cands:
        raise FileNotFoundError(f"no target workbook in {folder}")
    cands.sort(key=lambda f: (("TARGET" not in f.upper()),
                              -os.path.getmtime(os.path.join(folder, f))))
    path = os.path.join(folder, cands[0])

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    bondset = set(BONDS_ORDER)
    best: dict = {}
    for sn in wb.sheetnames:
        rows = [list(r) for r in wb[sn].iter_rows(values_only=True)]
        # Any header cell mentioning TARGET is a candidate, ignoring col 0 (the
        # sheet's hero title says "TARGET · JULY 2026") and % columns. Score
        # each by how many real bonds it yields -- that survives both the June
        # layout ("JUNE TARGET (cs)" col 4) and the July one ("TARGET (cs)" col 9).
        cands_c = set()
        for r in rows[:10]:
            for i, c in enumerate(r):
                if i == 0 or not c:
                    continue
                t = str(c).upper()
                if "TARGET" in t and "%" not in t and "ACH" not in t:
                    cands_c.add(i)
        for i in sorted(cands_c):
            found = {}
            for r in rows:
                name = str(r[0] or "").strip().upper()
                if name not in bondset or i >= len(r):
                    continue
                try:
                    v = float(r[i])
                except (TypeError, ValueError):
                    continue
                if v > 0:
                    found[name] = v
            if len(found) > len(best):
                best = found
    wb.close()
    if not best:
        raise RuntimeError(f"could not read bond targets from {path}")
    return {"targets": best, "file": os.path.basename(path)}


def shop_targets(base: str, month: str, year: int) -> dict:
    """
    bond -> monthly SHOP target = official monthly bond target x that bond's
    own 3-month shop share of liquidation, x (1 + EXTRA_UPLIFT).

    Shares come from the KSBC and Secondary BOND PERFORMANCE sheets, so they
    are authoritative period figures rather than sums of daily exports.
    """
    mno = PD.MONTH_NUM[month.upper()]
    months, y = [], year
    m = mno
    for _ in range(BASE_MONTHS):
        m -= 1
        if m == 0:
            m, y = 12, y - 1
        months.append(PD.MONTHS[m - 1])

    per_month = {}
    for name in months:
        a = PD.month_actuals(base, name)
        if any(v["tertiary"] > 0 for v in a.values()):
            per_month[name] = a
    if not per_month:
        raise RuntimeError(
            f"no completed months with KSBC data before {month.upper()} — "
            f"cannot compute shop share. Build the prior months' workbooks first.")

    off = read_official_targets(base, month)
    shares, shop_avg, out, fell_back = {}, {}, {}, []
    for b in BONDS_ORDER:
        ter = sum(a.get(b, {}).get("tertiary", 0.0) for a in per_month.values())
        tot = sum(a.get(b, {}).get("total", 0.0) for a in per_month.values())
        shares[b] = (ter / tot) if tot > 0 else 1.0
        shop_avg[b] = ter / len(per_month)
        o = off["targets"].get(b)
        if o:
            out[b] = o * SHOP_FACTOR * (1 + EXTRA_UPLIFT)
        else:                       # no official target -> hold its own average
            out[b] = shop_avg[b]
            fell_back.append(b)

    return {"targets": out, "shares": shares, "base": shop_avg,
            "official": off["targets"], "fell_back": fell_back,
            "base_months": list(per_month),
            "source": f"{off['file']} × {SHOP_FACTOR:.0%}"}


def prior_month(month: str, year: int) -> tuple[str, int]:
    i = PD.MONTH_NUM[month.upper()] - 1
    return (PD.MONTHS[11], year - 1) if i == 0 else (PD.MONTHS[i - 1], year)


# --- the plan --------------------------------------------------------------

def build_plan(base: str | None = None, month: str | None = None,
               year: int = 2026, as_of: _dt.date | None = None,
               hist_rows=None) -> dict:
    base = base or PD._default_base()
    if hist_rows is None:
        hist_rows = _load_history(base)

    dates = sorted({_dt.date.fromisoformat(str(r["date"])[:10]) for r in hist_rows})
    if month is None:
        month = PD.MONTHS[dates[-1].month - 1]
        year = dates[-1].year
    mno = PD.MONTH_NUM[month.upper()]
    in_month = [d for d in dates if d.year == year and d.month == mno]
    if as_of is None:
        as_of = in_month[-1] if in_month else _dt.date(year, mno, 1)
    ndays = calendar.monthrange(year, mno)[1]

    # dry days: the 1st always, plus any elapsed day that came in network-zero
    net_by_day = defaultdict(float)
    for r in hist_rows:
        d = _dt.date.fromisoformat(str(r["date"])[:10])
        if d.year == year and d.month == mno:
            net_by_day[d.day] += float(r["tertiary_cs"]) + float(r["invoice_cs"])
    dry = {DRY_DAY_OF_MONTH}
    for d in range(1, as_of.day + 1):
        if net_by_day.get(d, 0.0) < 1.0:
            dry.add(d)
    live_days = [d for d in range(1, ndays + 1) if d not in dry]
    n_live = len(live_days)

    # prior month's live days, for per-day comparisons downstream
    pm_probe, py_probe = prior_month(month, year)
    pmno = PD.MONTH_NUM[pm_probe]
    pnet = defaultdict(float)
    for r in hist_rows:
        d = _dt.date.fromisoformat(str(r["date"])[:10])
        if d.year == py_probe and d.month == pmno:
            pnet[d.day] += float(r["tertiary_cs"]) + float(r["invoice_cs"])
    pdays = calendar.monthrange(py_probe, pmno)[1]
    prior_selling = sum(1 for d in range(1, pdays + 1)
                        if d != DRY_DAY_OF_MONTH and pnet.get(d, 0.0) >= 1.0) or 1

    tg = shop_targets(base, month, year)
    targets_m = tg["targets"]
    if tg["fell_back"]:
        print(f"  ! no official target for {', '.join(tg['fell_back'])} — "
              f"held at their own 3-month shop average")
    pm, _py = prior_month(month, year)
    prior_act = PD.month_actuals(base, pm)

    # actuals by bond x day -- SHOP SALES ONLY
    act = defaultdict(lambda: defaultdict(float))
    for r in hist_rows:
        d = _dt.date.fromisoformat(str(r["date"])[:10])
        if d.year != year or d.month != mno:
            continue
        act[r["bond"]][d.day] += float(r["tertiary_cs"])

    # reconcile the daily series up to the authoritative BOND PERFORMANCE level
    # (summing KSBC's daily exports runs ~0.5-0.8% light vs its own period totals)
    auth = PD.month_actuals(base, month)
    bonds = {}
    for bond in BONDS_ORDER:
        dsum = sum(act[bond].values())
        a_tot = auth.get(bond, {}).get("tertiary", 0.0)
        k = (a_tot / dsum) if dsum > 0.01 and a_tot > 0 else 1.0

        tgt = targets_m.get(bond, 0.0)
        per_day = tgt / n_live if n_live else 0.0

        rows, cum_a, cum_t = [], 0.0, 0.0
        elapsed_live = 0
        for d in range(1, ndays + 1):
            dt = _dt.date(year, mno, d)
            isdry = d in dry
            t_day = 0.0 if isdry else per_day
            cum_t += t_day
            future = d > as_of.day
            aday = None if future else act[bond].get(d, 0.0) * k
            if aday is not None:
                cum_a += aday
            if not isdry and not future:
                elapsed_live += 1
            rows.append({
                "day": d, "date": dt, "dow": dt.strftime("%a"),
                "dry": isdry, "future": future,
                "target_day": t_day, "target_cum": cum_t,
                "actual": aday, "actual_cum": None if future else cum_a,
                "status": ("dry" if isdry else "—" if future
                           else "hit" if aday >= t_day - 1e-9 else "miss"),
            })

        mtd_a = cum_a
        mtd_t = next(r["target_cum"] for r in rows if r["day"] == as_of.day)
        left = [r for r in rows if r["day"] > as_of.day and not r["dry"]]
        remaining = tgt - mtd_a
        met = remaining <= 0.01
        needed = (max(remaining, 0.0) / len(left)) if left else 0.0

        bonds[bond] = {
            "bond": bond, "cluster": BOND_CLUSTER.get(bond),
            "target_month": tgt, "per_day": per_day, "rows": rows,
            "mtd_actual": mtd_a, "mtd_expected": mtd_t,
            "gap": mtd_a - mtd_t,
            "pct_done": (mtd_a / tgt) if tgt > 0 else 0.0,
            "pct_pace": (mtd_a / mtd_t) if mtd_t > 0 else 0.0,
            "days_left": len(left), "needed_per_day": needed,
            "target_met": met, "surplus": max(-remaining, 0.0),
            "next_date": left[0]["date"] if left else None,
            "run_rate": mtd_a / max(elapsed_live, 1),
            "elapsed_live": elapsed_live,
            "prior_actual": prior_act.get(bond, {}).get("tertiary", 0.0),
            "base_shop": tg["base"].get(bond, 0.0),
            "shop_share": tg["shares"].get(bond, 1.0),
            "official_target": tg["official"].get(bond, 0.0),
            "streak": _streak(rows, as_of.day),
        }

    ranked = sorted(bonds.values(), key=lambda b: -b["pct_done"])
    for i, b in enumerate(ranked, 1):
        b["rank"] = i
    for cl, members in CLUSTERS.items():
        peers = sorted((bonds[m] for m in members if m in bonds),
                       key=lambda x: -x["pct_done"])
        for i, b in enumerate(peers, 1):
            b["cluster_rank"] = i
            b["cluster_size"] = len(peers)
    # Where the month actually lands if nothing changes. A rate is a chore;
    # a projected shortfall is a verdict, and that is what lands.
    for b in bonds.values():
        b["projected"] = b["mtd_actual"] + b["run_rate"] * b["days_left"]
        b["shortfall"] = b["target_month"] - b["projected"]

    return {
        "month": month.upper(), "year": year, "mno": mno, "ndays": ndays,
        "as_of": as_of, "dry_days": sorted(dry), "selling_days": n_live,
        "prior_month": pm, "prior_selling_days": prior_selling,
        "target_source": tg["source"], "base_months": tg["base_months"],
        "uplift": EXTRA_UPLIFT, "basis": "KSBC shop liquidation",
        "bonds": bonds, "order": [b["bond"] for b in ranked],
        "network": _network(bonds, ndays, as_of),
    }


def _streak(rows, as_of_day) -> int:
    n = 0
    for r in reversed([r for r in rows if r["day"] <= as_of_day and not r["dry"]]):
        if r["actual"] is not None and r["actual"] >= r["target_day"] - 1e-9:
            n += 1
        else:
            break
    return n


def _network(bonds: dict, ndays: int, as_of) -> dict:
    out = {"rows": []}
    for d in range(1, ndays + 1):
        agg = {"day": d, "target_day": 0.0, "actual": 0.0, "dry": False,
               "future": False, "date": None, "dow": None}
        for b in bonds.values():
            r = b["rows"][d - 1]
            agg["date"], agg["dow"] = r["date"], r["dow"]
            agg["dry"], agg["future"] = r["dry"], r["future"]
            agg["target_day"] += r["target_day"]
            if r["actual"] is not None:
                agg["actual"] += r["actual"]
        if agg["future"]:
            agg["actual"] = None
        out["rows"].append(agg)
    for k in ("target_month", "mtd_actual", "mtd_expected", "needed_per_day",
              "per_day", "prior_actual"):
        out[k] = sum(b[k] for b in bonds.values())
    out["gap"] = out["mtd_actual"] - out["mtd_expected"]
    out["pct_done"] = (out["mtd_actual"] / out["target_month"]
                       if out["target_month"] else 0.0)
    out["pct_pace"] = (out["mtd_actual"] / out["mtd_expected"]
                       if out["mtd_expected"] else 0.0)
    out["days_left"] = max(b["days_left"] for b in bonds.values())
    out["run_rate"] = sum(b["run_rate"] for b in bonds.values())
    return out


def _load_history(base: str):
    import csv as _csv
    p = os.path.join(base, "Targets", ".pace", "daily_liquidation.csv")
    if not os.path.exists(p):
        rows = PD.build_history(base, verbose=False)
        PD.write_history(rows, base)
        return rows
    with open(p) as fh:
        return [{**r, "tertiary_cs": float(r["tertiary_cs"]),
                 "invoice_cs": float(r["invoice_cs"]),
                 "shop_code": int(r["shop_code"])}
                for r in _csv.DictReader(fh)]


if __name__ == "__main__":
    base = PD._default_base()
    plan = build_plan(base)
    n = plan["network"]
    print(f"{plan['month']} {plan['year']} · as on {plan['as_of']:%d %b} · "
          f"{plan['selling_days']} selling days (dry {plan['dry_days']})")
    print(f"base: {plan['target_source']}")
    print(f"\nNETWORK  target {n['target_month']:,.0f}  "
          f"= {n['per_day']:,.0f} cs/day   ·   sold {n['mtd_actual']:,.0f} "
          f"({n['pct_done']:.0%})   ·   {n['days_left']} days left   "
          f"·   NEEDS {n['needed_per_day']:,.0f} cs/day")
    print(f"\n{'#':>2}  {'BOND':<15} {'TARGET':>7} {'/day':>6} {'SOLD':>7} "
          f"{'SHOULD BE':>10} {'BEHIND':>8} {'DONE':>5} {'NEED/DAY':>9}")
    for b in [plan["bonds"][x] for x in plan["order"]]:
        print(f"{b['rank']:>2}  {b['bond']:<15} {b['target_month']:>7,.0f} "
              f"{b['per_day']:>6,.0f} {b['mtd_actual']:>7,.0f} "
              f"{b['mtd_expected']:>10,.0f} {b['gap']:>+8,.0f} "
              f"{b['pct_done']:>5.0%} {b['needed_per_day']:>9,.0f}")
