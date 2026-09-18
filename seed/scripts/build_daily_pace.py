#!/usr/bin/env python3
"""
Daily pace — driver.

    python3 .claude/scripts/build_daily_pace.py [--month JULY] [--year 2026]
                                                [--refresh] [--out DIR]

Produces, for the month in play:

  Targets/<MONTH> TARGETS/DAILY PACE TRACKER - <MONTH>.xlsx
  Targets/<MONTH> TARGETS/DAILY PACE PACK - <MONTH> <DD>.txt

--refresh rebuilds the daily-liquidation history cache from the source
workbooks (~30 s). Without it a cached CSV is reused when present.

Nothing is written over the live folder by this script -- it writes to the
scratch path given by --out (default: the outputs dir) so the caller can show
a summary and take approval first, per the folder's single-gate rule.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)

import pace_data as PD
import pace_engine as PE
import pace_shops as PS
import pace_pdf as PP
import pace_whatsapp as PW
import pace_workbook as WB


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default=None)
    ap.add_argument("--month", default=None)
    ap.add_argument("--year", type=int, default=2026)
    ap.add_argument("--as-of", default=None, help="YYYY-MM-DD")
    ap.add_argument("--refresh", action="store_true",
                    help="rebuild the daily history cache from source workbooks")
    ap.add_argument("--out", default=None, help="output directory")
    a = ap.parse_args()

    base = a.base or PD._default_base()
    print(f"base: {base}")

    if a.refresh or not os.path.exists(
            os.path.join(base, "Targets", ".pace", "daily_liquidation.csv")):
        print("rebuilding daily liquidation history…")
        rows = PD.build_history(base, year=a.year)
        PD.write_history(rows, base)
    hist = PE._load_history(base)

    as_of = _dt.date.fromisoformat(a.as_of) if a.as_of else None
    plan = PE.build_plan(base, month=a.month, year=a.year, as_of=as_of,
                         hist_rows=hist)
    master = PD.load_master(base)
    targets, prior_matched = PS.build_shop_daily(base, plan, hist, master)

    outdir = a.out or os.path.join(base, "Targets", f"{plan['month']} TARGETS")
    os.makedirs(outdir, exist_ok=True)

    xlsx = os.path.join(outdir, f"DAILY PACE TRACKER - {plan['month']}.xlsx")
    WB.build(plan, targets, xlsx)

    plan["base"] = base
    plan["split"] = PS.channel_split(hist, plan)
    plan["prior_matched"] = prior_matched
    pdfs = PP.build_cluster_pdfs(plan, targets, master, PE.CLUSTERS,
                                 PE.ASM_BY_CLUSTER, outdir)

    pack = PW.build_pack(plan, targets, master, PE.CLUSTERS, PE.ASM_BY_CLUSTER)
    txt = os.path.join(outdir, f"DAILY PACE PACK - {plan['month']}.txt")
    with open(txt, "w") as fh:
        fh.write(pack)

    n = plan["network"]
    onpace = sum(1 for b in plan["bonds"].values() if b["gap"] >= 0)
    print(f"\n{plan['month']} {plan['year']} · data through "
          f"{plan['as_of']:%d %b} · {plan['selling_days']} selling days "
          f"(dry: {plan['dry_days']})")
    print(f"targets read from: {plan['target_source']}")
    print(f"TARGET {n['target_month']:,.0f}  = {n['per_day']:,.0f} cs/day  ·  "
          f"SOLD {n['mtd_actual']:,.0f} ({n['pct_done']:.0%})  ·  "
          f"behind by {-n['gap']:,.0f}")
    print(f"{onpace} of 15 bonds on or ahead of pace · needs "
          f"{n['needed_per_day']:,.0f} cs/day for the last {n['days_left']} days")
    print(f"\nworkbook: {xlsx}")
    print(f"pack:     {txt}")
    for q in pdfs:
        print(f"pdf:      {q}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
