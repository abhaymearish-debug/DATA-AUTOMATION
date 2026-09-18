#!/usr/bin/env python3
"""
Verify the `ksd-aging-live` artifact — tier KPIs, bond direction, brand/pack
movement, insights and the newly-stuck / cleared lists.

Every figure on the page is a DIFF between two snapshots, so it cannot be
checked against a single workbook. Each diff is re-derived here by reopening
the two source workbooks it names, and compared to what the page renders.

That is deliberately the long way round: checking the page against the
extractor's own diff would only prove the code agrees with itself. It also
carries an INDEPENDENT reconciliation — the per-bond, per-brand and per-pack
movements must each sum to the network delta — which is what caught the
double-subtraction bug in the first cut of build_diff().

Requires: node + jsdom, and verify_aging_artifact_dump.js beside this file.

    python3 .claude/scripts/verify_aging_artifact.py [--base DIR]
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import math
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
DUMP_JS = HERE / "verify_aging_artifact_dump.js"

PASS, FAIL = [], []
TOL = 0.06
# The page treats anything under half a case as FLAT; the payload keeps the
# raw 0.05 cs epsilon. That is a display decision, not a data one, so the
# harness applies the page's rule when checking the page.
FLAT = 0.5


def check(name, got, want, tol=0.0):
    if isinstance(want, (int, float)) and isinstance(got, (int, float)):
        ok = abs(got - want) <= tol
    else:
        ok = got == want
    (PASS if ok else FAIL).append((name, got, want))
    return ok


def num(s):
    """First number in a rendered cell; typographic minus becomes a real one."""
    if s is None:
        return None
    m = re.search(r"-?[\d,]*\.?\d+", str(s).replace("−", "-"))
    return float(m.group(0).replace(",", "")) if m else None


def nums(s):
    return [float(x.replace(",", ""))
            for x in re.findall(r"-?[\d,]*\.?\d+", str(s).replace("−", "-"))]


def r1(x):
    """Round to 1dp the way JS does — half UP, not half-to-even.

    Aging cases are twelfths and forty-eighths, so exact halves are common and
    Python's banker's rounding invents 0.1-case disagreements against a page
    that renders correctly.
    """
    return math.floor(x * 10 + 0.5) / 10 if x >= 0 else -(math.floor(-x * 10 + 0.5) / 10)


def r0(x):
    return math.floor(x + 0.5) if x >= 0 else -math.floor(-x + 0.5)


def load_refresher():
    spec = importlib.util.spec_from_file_location(
        "rfr", HERE / "refresh_aging_live_artifact.py")
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def workbook_for_date(rfr, root: Path, date: str):
    y, mo, dy = (int(x) for x in date.split("-"))
    folder = root / "Aging stock"
    for p in list(folder.glob("AGING STOCK ANALYSIS - *.xlsx")) + \
             list((folder / "_archive").glob("AGING STOCK ANALYSIS - *.xlsx")):
        if p.name.startswith("~$"):
            continue
        parsed = rfr.parse_wb_name(p)
        if not parsed:
            continue
        _, to_m, day = parsed
        if rfr.MONTHS[to_m] == mo and day == dy:
            return p
    return None


# Labels are Title Case in the markup; the CSS uppercases them for display,
# so compare case-insensitively.
TIER_LABEL = {"nm": "non-moving", "cr": "critical", "sl": "slow moving",
              "new": "new stock", "ok": "healthy"}
TIER_NAME = {"nm": "Non-Moving", "cr": "Critical", "sl": "Slow Moving",
             "new": "New Stock", "ok": "Healthy",
             "absent": "Not on shelf", "gone": "Off the report"}
AGING = ("nm", "cr", "sl")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default=None)
    ap.add_argument("--html", default=None)
    args = ap.parse_args()

    rfr = load_refresher()
    root = Path(args.base) if args.base else rfr._default_root()

    tmp = Path(tempfile.mkdtemp())
    html = Path(args.html) if args.html else tmp / "aging_live.html"
    payload = rfr.build(root, None, html, backfill=True)

    env = dict(os.environ)
    jsd = Path("/tmp/node_modules/jsdom")
    if jsd.exists():
        env.setdefault("JSDOM_PATH", str(jsd))
    res = subprocess.run(["node", str(DUMP_JS), str(html)],
                         capture_output=True, text=True, env=env)
    if res.returncode != 0:
        print(res.stderr[:4000], file=sys.stderr)
        raise SystemExit("artifact failed to render in jsdom")
    if res.stderr.strip():
        print("  ! jsdom console output:\n" + res.stderr[:2000])
    dom = json.loads(res.stdout)

    diffs = payload["diffs"]
    T = payload["totals"]
    print(f"workbook : {payload['workbook']}")
    print(f"windows  : {len(diffs)}  ({', '.join(d['from'] for d in diffs)})\n")
    check("window button count", len(dom["windowLabels"]), len(diffs))

    cache: dict = {}

    def lines_for(date):
        if date not in cache:
            p = workbook_for_date(rfr, root, date)
            cache[date] = rfr.read_workbook(p)["all"] if p else None
        return cache[date]

    for wi, d in enumerate(diffs):
        w = dom["windows"].get(str(wi))
        tag = f"[{d['from']}→{d['to']}]"
        if w is None:
            check(f"{tag} window rendered", False, True)
            continue

        # ---- 1. re-derive the whole diff from the source workbooks -------
        prev_rows, cur_rows = lines_for(d["from"]), lines_for(d["to"])
        if prev_rows and cur_rows:
            truth = rfr.build_diff(prev_rows, cur_rows, d["from"], d["to"])
            for k in ("d_aging", "aging_from", "aging_to", "cs_newly",
                      "cs_cleared", "share_from", "share_to"):
                check(f"{tag} rebuilt · {k}", d[k], truth[k], TOL)
            for k in ("n_newly", "n_cleared", "lines_from", "lines_to"):
                check(f"{tag} rebuilt · {k}", d[k], truth[k])
            check(f"{tag} rebuilt · bond movement",
                  {r["bond"]: r1(r["d"]) for r in d["bond_mv"]},
                  {r["bond"]: r1(r["d"]) for r in truth["bond_mv"]})
            check(f"{tag} rebuilt · tier movement",
                  {k: (r1(v["from"]), r1(v["to"])) for k, v in d["tiers"].items()},
                  {k: (r1(v["from"]), r1(v["to"])) for k, v in truth["tiers"].items()})
        else:
            print(f"  ! {tag} source workbook missing — payload-only check")

        # ---- 2. independent reconciliation ------------------------------
        # each breakdown must sum to the same network delta
        for lbl, rows in (("bond", d["bond_mv"]), ("brand", d["brand_mv"]),
                          ("pack", d["pack_mv"])):
            check(f"{tag} {lbl} movement sums to the network delta",
                  round(sum(r["d"] for r in rows), 2), d["d_aging"], 0.15)
            check(f"{tag} {lbl} 'to' sums to aging_to",
                  round(sum(r["to"] for r in rows), 2), d["aging_to"], 0.15)
            check(f"{tag} {lbl} 'from' sums to aging_from",
                  round(sum(r["from"] for r in rows), 2), d["aging_from"], 0.15)
        check(f"{tag} aging tiers sum to the total",
              round(sum(d["tiers"][k]["to"] for k in AGING), 2), d["aging_to"], 0.15)
        check(f"{tag} tier deltas sum to the network delta",
              round(sum(d["tiers"][k]["to"] - d["tiers"][k]["from"] for k in AGING), 2),
              d["d_aging"], 0.15)
        recon = d["cs_newly"] - d["cs_cleared"] + d["cs_up"] + d["cs_down"]
        check(f"{tag} net = newly − cleared ± drift", round(recon, 2), d["d_aging"], 0.2)
        check(f"{tag} current window matches the live total",
              d["aging_to"], T["aging"], TOL)

        # ---- 3. tier KPI tiles ------------------------------------------
        tiles = w["tiers"]
        check(f"{tag} six tier tiles", len(tiles), 6)
        def growth(cur, prior):
            return (cur / prior - 1) * 100 if prior > 0 else (100 if cur > 0 else 0)

        hero = tiles[0]
        check(f"{tag} hero label", "total aging" in hero["label"].lower(), True)
        check(f"{tag} hero value", num(hero["val"]), r0(d["aging_to"]), 1)
        # pill = growth %, footer-right = absolute cases (liquidation anatomy)
        gh = growth(d["aging_to"], d["aging_from"])
        if abs(gh) >= 0.5:
            check(f"{tag} hero growth %", num(hero["delta"]), round(gh), 1)
        check(f"{tag} hero absolute change", num(hero["dd"]), r0(d["d_aging"]), 1)
        check(f"{tag} hero share", nums(hero["foot"])[0], d["share_to"], 0.06)
        check(f"{tag} hero footer names the comparison date",
              str(int(d["from"][8:10])) in hero["foot"], True)
        check(f"{tag} tiles carry no sparkline", any(t["spark"] for t in tiles), False)
        for i, k in enumerate(["nm", "cr", "sl", "new", "ok"]):
            t, tl = d["tiers"][k], tiles[i + 1]
            dd = t["to"] - t["from"]
            check(f"{tag} tile {k} label", TIER_LABEL[k] in tl["label"].lower(), True)
            check(f"{tag} tile {k} value", num(tl["val"]), r0(t["to"]), 1)
            check(f"{tag} tile {k} lines", nums(tl["foot"])[0], t["n_to"])
            # A snapshot taken before the NEW STOCK tier existed reports 0 for
            # it, which is "not measured". The tile must refuse to show a
            # delta rather than invent growth from a phantom zero.
            if d.get("new_tier_missing") and k == "new":
                check(f"{tag} tile new suppresses the phantom delta",
                      tl["delta"], "n/a")
                check(f"{tag} tile new says why", "new tier" in tl["foot"], True)
                continue
            gk = growth(t["to"], t["from"])
            if abs(gk) >= 0.5:
                check(f"{tag} tile {k} growth %", num(tl["delta"]), round(gk), 1)
            else:
                check(f"{tag} tile {k} tiny move reads as 0%, not −0%",
                      tl["delta"], "0%")
            check(f"{tag} tile {k} absolute change", num(tl["dd"]), r0(dd), 1)
            check(f"{tag} tile {k} footer names the prior value",
                  f"{r0(t['from']):,}" in tl["foot"], True)
            # Colour convention, and the single most important assertion on
            # this page: on an AGING tier a RISE must read red, but on Healthy
            # and New Stock a rise must read green. The markup uses
            # liquidation's .dpill pos|neg vocabulary, so red == "neg".
            # Gate on the PERCENTAGE, because that is what the pill shows:
            # Non-Moving moving -1 cs on a base of 698 is -0.1%, which renders
            # as a flat "0%" pill by design. Colour is only asserted where the
            # pill is actually showing a direction.
            if abs(gk) >= 0.5 and not (d.get("new_tier_missing") and k == "new"):
                bad = (dd > 0) if k in AGING else (dd < 0)
                check(f"{tag} tile {k} delta colour",
                      tl["dcls"].split()[-1], "neg" if bad else "pos")

        # ---- 4. bond direction table ------------------------------------
        rows = sorted(d["bond_mv"], key=lambda r: -r["d"])
        check(f"{tag} bond row count", len(w["bonds"]), len(rows))
        for got, src in zip(w["bonds"], rows):
            b = src["bond"]
            check(f"{tag} bond {b} name", got["bond"], b)
            check(f"{tag} bond {b} change", num(got["d"]), r1(src["d"]), 0.06)
            ft = nums(got["fromto"])
            check(f"{tag} bond {b} from", ft[0], r1(src["from"]), 0.06)
            check(f"{tag} bond {b} to", ft[1], r1(src["to"]), 0.06)
            check(f"{tag} bond {b} lines", nums(got["lines"])[0], src["n_to"])
            # aging as a share of THIS bond's own shelf -- the figure that
            # reorders the league table
            bd = next((x for x in payload["bonds"] if x["bond"] == b), None)
            if bd and bd.get("closing"):
                check(f"{tag} bond {b} intensity", num(got["intensity"]),
                      round(bd["aging"] / bd["closing"] * 100, 1), 0.15)
            want_dir = "▲ WORSE" if src["d"] > FLAT else (
                "▼ BETTER" if src["d"] < -FLAT else "— FLAT")
            check(f"{tag} bond {b} direction", got["dir"], want_dir)
        check(f"{tag} bond table sorted worst-first",
              [num(r["d"]) for r in w["bonds"]],
              sorted([num(r["d"]) for r in w["bonds"]], reverse=True))
        foot = w["bondFoot"]
        check(f"{tag} bond foot count", num(foot[1]), len(rows))
        check(f"{tag} bond foot worse", nums(foot[2])[0],
              sum(1 for r in rows if r["d"] > FLAT))
        check(f"{tag} bond foot better", nums(foot[2])[1],
              sum(1 for r in rows if r["d"] < -FLAT))
        check(f"{tag} bond foot delta", num(foot[6]), r1(d["d_aging"]), 0.15)
        check(f"{tag} bond foot lines", num(foot[9]), d["lines_to"])
        # tier legend must name all five tiers
        check(f"{tag} tier legend complete", len(w["legend"]), 5)

        # ---- 5. brand / pack tables -------------------------------------
        for key, dom_rows, src_rows, kname in (
                ("brand", w["brand"], sorted(d["brand_mv"], key=lambda r: -r["d"]), "brand"),
                ("pack",  w["pack"],  sorted(d["pack_mv"],  key=lambda r: -r["d"]), "pack")):
            check(f"{tag} {key} row count", len(dom_rows), len(src_rows))
            for got, src in zip(dom_rows, src_rows):
                check(f"{tag} {key} {src[kname]} name", got[0], src[kname])
                ft = nums(got[1])
                check(f"{tag} {key} {src[kname]} from", ft[0], r1(src["from"]), 0.06)
                check(f"{tag} {key} {src[kname]} to", ft[1], r1(src["to"]), 0.06)
                check(f"{tag} {key} {src[kname]} change", num(got[2]), r1(src["d"]), 0.06)
                check(f"{tag} {key} {src[kname]} lines", num(got[4]), src["n_to"])
            shares = [num(r[3]) for r in dom_rows]
            check(f"{tag} {key} shares sum to 100", round(sum(shares)), 100, 1)

        # ---- 6. newly / cleared (liquidation's .mv movers, top 8) --------
        for key, src, csf in (("newly", d["newly"], "cl"),
                              ("cleared", d["cleared"], "was_cl")):
            got_rows = w[key]
            check(f"{tag} {key} row count", len(got_rows), min(8, len(src)))
            # SIGNED: newly-stuck ADDS cases to the exposure, cleared REMOVES
            # them, and the page renders that as +x / −x. Comparing magnitudes
            # would pass on a list that showed a clearance as an addition.
            sign = 1 if key == "newly" else -1
            for got, l in zip(got_rows, src):
                check(f"{tag} {key} {l['shop']} cases",
                      num(got["d"]), sign * r1(l[csf]), 0.06)
                check(f"{tag} {key} {l['shop']} named", l["shop"] in got["nm"], True)
                check(f"{tag} {key} {l['shop']} bond named", l["bond"] in got["sm"], True)
                check(f"{tag} {key} {l['shop']} sku named",
                      l["br"] in got["sm"] and l["p"] in got["sm"], True)
            vals = [abs(num(r["d"])) for r in got_rows]
            check(f"{tag} {key} sorted biggest first",
                  all(vals[i] >= vals[i+1] - 0.06 for i in range(len(vals)-1)), True)
        # newly-stuck reads as a decline (red, .mv.d); cleared as a gain
        check(f"{tag} newly styled as decline",
              all("mv d" in r["cls"] for r in w["newly"]), True)
        check(f"{tag} cleared styled as gain",
              all("mv u" in r["cls"] for r in w["cleared"]), True)
        for got, l in zip(w["newly"], d["newly"]):
            want = ("arrived and stalled" if l["was"] == "absent"
                    else "was " + {"ok": "healthy", "new": "new stock"}.get(l["was"], l["was"]))
            check(f"{tag} newly {l['shop']} note", got["s"], want)
        for got, l in zip(w["cleared"], d["cleared"]):
            if l["how"] == "soldout":
                check(f"{tag} cleared {l['shop']} note", got["s"], "shelf is empty")
            elif l["how"] == "gone":
                check(f"{tag} cleared {l['shop']} note",
                      got["s"].startswith("off the report"), True)
            else:
                check(f"{tag} cleared {l['shop']} note",
                      got["s"].endswith("cs left, and moving"), True)
        check(f"{tag} newly header count", nums(w["newlyHint"])[0], d["n_newly"])
        check(f"{tag} cleared header count", nums(w["clearedHint"])[0], d["n_cleared"])

        # ---- 7b. charts --------------------------------------------------
        # trend / waterfall / flow, in the order they are built
        ch = w["charts"]
        check(f"{tag} three charts drawn", len(ch), 3)
        trend, wf, flow = ch
        check(f"{tag} trend labels = snapshot count",
              len(trend["labels"]), len(payload["hist_dates"]))
        check(f"{tag} trend aging series",
              [r1(v) for v in trend["sets"][0]["data"]],
              [r1(h["aging"]) for h in payload["history"]])
        # waterfall: start, +newly, −cleared, drift, end — as floating bars
        check(f"{tag} waterfall steps", len(wf["sets"][0]["data"]), 5)
        bars = wf["sets"][0]["data"]
        check(f"{tag} waterfall opens at the from-total", r1(bars[0][1]), r1(d["aging_from"]), 0.06)
        check(f"{tag} waterfall closes at the to-total", r1(bars[4][1]), r1(d["aging_to"]), 0.06)
        check(f"{tag} waterfall newly step",
              r1(bars[1][1] - bars[1][0]), r1(d["cs_newly"]), 0.06)
        check(f"{tag} waterfall cleared step",
              r1(bars[2][1] - bars[2][0]), r1(-d["cs_cleared"]), 0.06)
        # 0.15 not 0.06: this step is read back as the difference of two
        # RUNNING totals, so it carries the float error of every step before
        # it. The bar itself is drawn from the same running values, so what is
        # on screen is exact — only this re-derivation accumulates.
        check(f"{tag} waterfall drift step",
              r1(bars[3][1] - bars[3][0]), r1(d["cs_up"] + d["cs_down"]), 0.15)
        check(f"{tag} waterfall running total lands on the end",
              abs(bars[3][1] - d["aging_to"]) < 0.25, True)
        # flow chart: only transitions touching the aging population
        want_flow = [f for f in d["flow"]
                     if f["from"] in AGING or f["to"] in AGING][:8]
        check(f"{tag} flow bar count", len(flow["sets"][0]["data"]), len(want_flow))
        check(f"{tag} flow values",
              [r1(v) for v in flow["sets"][0]["data"]],
              [r1(f["cs"]) for f in want_flow])
        check(f"{tag} flow excludes non-aging transitions",
              all(" → " in l for l in flow["labels"]), True)
        for lab, f in zip(flow["labels"], want_flow):
            check(f"{tag} flow label {f['from']}→{f['to']}",
                  TIER_NAME[f["from"]] in lab and TIER_NAME[f["to"]] in lab, True)

        # ---- 7c. cluster cards -------------------------------------------
        mv_by = {r["bond"]: r for r in d["bond_mv"]}
        by_cl = {}
        for b in payload["bonds"]:
            g = by_cl.setdefault(b["cluster"], {"to": 0.0, "from": 0.0})
            m = mv_by.get(b["bond"])
            if m:
                g["to"] += m["to"]
                g["from"] += m["from"]
        check(f"{tag} cluster card count", len(w["clusters"]), len(payload["clusters"]))
        for cd in w["clusters"]:
            g = by_cl[int(cd["n"])]
            check(f"{tag} cluster {cd['n']} cases", num(cd["cases"]), r0(g["to"]), 1)
            check(f"{tag} cluster {cd['n']} movement",
                  num(cd["move"]), r0(g["to"] - g["from"]), 1)
        check(f"{tag} cluster cases sum to the network",
              round(sum(v["to"] for v in by_cl.values()), 1), round(d["aging_to"], 1), 0.2)

        # ---- 8. drill-down ----------------------------------------------
        if w.get("drill"):
            b = w["drill"]["bond"]
            mine = [l for l in payload["lines"] if l["b"] == b]
            shops = {}
            for l in mine:
                shops[l["c"]] = shops.get(l["c"], 0) + l["cl"]
            check(f"{tag} drill {b} shop count", len(w["drill"]["shops"]), len(shops))
            for sb in w["drill"]["shops"]:
                check(f"{tag} drill {b}/{sb['code']} cases",
                      num(sb["cs"]), r1(shops.get(sb["code"], 0)), 0.06)

    # ---- longitudinal: chronic + churn -----------------------------------
    # Independently recompute "aging in every snapshot" from the line history
    # rather than trusting longitudinal() -- this is the headline claim on the
    # page ("70% has never shifted") and it must not rest on one function.
    lg = payload["longi"]
    w0 = dom["windows"]["0"]
    hist_dates = sorted(payload["snap_dates"])
    if len(hist_dates) >= 2 and all(lines_for(x) for x in hist_dates):
        states = {}
        for dt in hist_dates:
            for r in lines_for(dt):
                k = rfr._line_key(r)
                states.setdefault(k, []).append(r["t"] in set(AGING))
        latest = {rfr._line_key(r) for r in lines_for(hist_dates[-1])
                  if r["t"] in set(AGING)}
        chronic = [k for k, h in states.items()
                   if k in latest and len(h) == len(hist_dates) and all(h)]
        cur_by_key = {rfr._line_key(r): r for r in lines_for(hist_dates[-1])}
        chronic_cs = sum(cur_by_key[k]["cl"] for k in chronic)
        check("chronic · position count", lg["chronic_n"], len(chronic))
        check("chronic · cases", lg["chronic_cs"], round(chronic_cs, 2), 0.15)
        churn = [k for k, h in states.items()
                 if k in latest and True in h and False in h
                 and any(not x for x in h[h.index(True):]) and h[-1]]
        check("churn · position count", lg["churn_n"], len(churn))
        check("chronic and churn are disjoint",
              set(chronic) & set(churn), set())
        check("chronic cannot exceed the exposure",
              lg["chronic_cs"] <= T["aging"] + 0.1, True)
    check("chronic table rows", len(w0["chronic"]), len(lg["chronic"]))
    for row, r in zip(w0["chronic"], lg["chronic"]):
        check(f"chronic row {r['shop']}", row[0], r["shop"])
        check(f"chronic cases {r['shop']}", num(row[3]), r1(r["cl"]), 0.06)
    check("chronic hint count", nums(w0["chronicHint"])[0], lg["chronic_n"])
    check("churn table rows", len(w0["churn"]), len(lg["churn"]))
    check("churn hint count", nums(w0["churnHint"])[0], lg["churn_n"])

    # ---- the four levers -------------------------------------------------
    iq = payload.get("iq") or {}
    if iq:
        lv = {x["lv"]: x for x in w0["levers"]}
        check("four lever cards", len(w0["levers"]), 4)
        check("lever · stop", num(lv["stop"]["val"]),
              r0(iq["redispatch"]["cs_stuck"]), 1)
        check("lever · refill", num(lv["refill"]["val"]),
              r0(iq["stockouts"]["rate"]), 1)
        check("lever · route", num(lv["route"]["val"]), r0(iq["routing"]["cs"]), 1)
        check("lever · incentive", num(lv["incent"]["val"]),
              r0(iq["incentives"]["still_cs"]), 1)
        check("lever · refill is a rate not a stock", lv["refill"]["unit"], "cs/mo")
        # each lever's detail table
        check("lever rows · stop", len(w0["leverRows"]["stop"]),
              len(iq["redispatch"]["events"]) + 1)      # +1 total row
        check("lever rows · refill", len(w0["leverRows"]["refill"]),
              len(iq["stockouts"]["rows"]) + 1)
        check("lever rows · route", len(w0["leverRows"]["route"]),
              len(iq["routing"]["rows"]) + 1)
        check("lever rows · incentive", len(w0["leverRows"]["incent"]),
              len(iq["incentives"]["rows"]) + 1)
        for row, e in zip(w0["leverRows"]["stop"], iq["redispatch"]["events"]):
            check(f"stop · {e['shop']} sent", num(row[4]), r1(e["cs"]), 0.06)
            check(f"stop · {e['shop']} stuck", num(row[6]), r1(e["stuck"]), 0.06)
        for row, r in zip(w0["leverRows"]["refill"], iq["stockouts"]["rows"]):
            check(f"refill · {r['shop']} rate", num(row[3]), r1(r["rate"]), 0.06)
        for row, r in zip(w0["leverRows"]["route"], iq["routing"]["rows"]):
            check(f"route · {r['bond']} {r['br']} stuck", num(row[2]), r1(r["stuck"]), 0.06)
            check(f"route · {r['bond']} {r['br']} clears in", num(row[4]),
                  r1(r["clear_mo"]), 0.06)
        # cover buckets must partition the whole exposure
        cov = iq["cover"]
        check("cover buckets partition the exposure",
              round(sum(v["cs"] for v in cov.values()), 1), round(T["aging"], 1), 0.15)
        check("cover bucket lines partition",
              sum(v["n"] for v in cov.values()), T["aging_n"])
        check("lever hint quotes the unclearable bucket",
              f"{cov['dead']['cs']:,.0f}".replace(",", ",") in w0["lvHint"]
              or f"{r0(cov['dead']['cs']):,}" in w0["lvHint"], True)
        # incentive "sold" must be sales SINCE the incentive, never the window
        for row, r in zip(w0["leverRows"]["incent"], iq["incentives"]["rows"]):
            if r["sold"] is not None:
                check(f"incentive · {r['shop']} sold-since", num(row[6]),
                      r1(r["sold"]), 0.06)
                cur = next((l for l in payload["lines"]
                            if l["c"] == r["code"] and l["br"] == r["br"]
                            and l["p"] == r["p"]), None)
                if cur:
                    check(f"incentive · {r['shop']} not the window total",
                          r["sold"] <= cur["sa"] + 0.06, True)

    # ---- heatmap ---------------------------------------------------------
    cell = {}
    for l in payload["lines"]:
        cell[(l["b"], l["br"])] = cell.get((l["b"], l["br"]), 0) + l["cl"]
    bonds_o = [b["bond"] for b in payload["bonds"]]
    brands_o = [b["brand"] for b in payload["brands"]]
    check("heat row count", len(w0["heat"]["rows"]), len(bonds_o))
    check("heat column count", len(w0["heat"]["cols"]), len(brands_o) + 2)
    for row, bond in zip(w0["heat"]["rows"], bonds_o):
        check(f"heat row {bond} label", row["bond"], bond)
        want_tot = sum(cell.get((bond, br), 0) for br in brands_o)
        check(f"heat row {bond} total", num(row["tot"]), r0(want_tot), 1)
    grand = num(w0["heat"]["allrow"][-1])
    check("heat grand total", grand, r0(T["aging"]), 2)

    # ---- global ----------------------------------------------------------
    check("header shows the shop scope", str(T["shops"]) in dom["header"]["scope"], True)
    print(f"{len(PASS)} passed, {len(FAIL)} failed")
    if FAIL:
        print("\nFAILURES")
        for n, got, want in FAIL[:50]:
            print(f"  ✗ {n}\n      got  {got!r}\n      want {want!r}")
        return 1
    print("ALL CHECKS PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
