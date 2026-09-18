#!/usr/bin/env python3
"""
Deep analysis over the aging workbook — the layer that turns the aging report
from a post-mortem into a set of sized levers.

Everything here was specified from an adversarial audit of the workbook
(9 Aug 2026) that looked for analytical value the dashboard was not using.
Each function returns figures that were hand-verified against the workbook at
the time; the docstrings record the definition, because the definition is the
part that is easy to get subtly wrong.

FOUR LEVERS the business actually has, and what sizes each:
  STOP        re-ordering into lines that are already dead  -> redispatch()
  REFILL      shops that sell it and are standing empty     -> stockouts()
  REDISTRIBUTE / re-route to a proven buyer in the bond     -> routing()
  INCENTIVE   pay per case to push it                       -> incentives()

Plus leading indicators: newstock() (aging already in the pipe) and
stall_profile() (when each line actually died).

TRAPS, all found the hard way — do not undo these:
  * "Months of Cover" is a TRAILING-3-MONTH figure. 32 lines exceed 100 months
    and the worst reads 2,056 months on 18.9 cs. Never plot it raw; bucket it.
  * cover == "NO SALES" does NOT mean "never sold". 52 lines read NO SALES yet
    sold 380 cs earlier in the window. "Never sold" is the NON-MOVING tier.
  * NEW STOCK lines that arrived days before the cut-off cannot have sold.
    Only judge the >= NEW_MIN_DAYS cohort.
"""

from __future__ import annotations

import re
from pathlib import Path

import openpyxl

# --- thresholds, all in one place -----------------------------------------
STOCKOUT_MIN_RATE = 1.0    # cs/month of proven demand before an empty shelf counts
STOCKOUT_MAX_COVER = 0.25  # months of stock left ("about a week")
BUYER_MIN_RATE = 0.3       # cs/month before a shop counts as a proven buyer
BUYER_MAX_COVER = 3.0      # months; a buyer already overstocked is not a home
ROUTE_MAX_MONTHS = 2.0     # buyers must absorb the stuck stock this fast
DEAD_MONTHS = 2            # consecutive zero-sale months before a drop is "into a dead line"
NEW_MIN_DAYS = 14          # days on shelf before a new placement can be judged
PUSHABLE_COVER = 6.0       # months; above this, pushing will not clear it

TIER_TOKENS = {"NON-MOVING": "nm", "CRITICAL": "cr", "SLOW": "sl",
               "NEW STOCK": "new", "HEALTHY": "ok"}


def _tier(raw):
    txt = re.sub(r"[^A-Z \-]", "", str(raw or "").upper()).strip()
    for name, key in TIER_TOKENS.items():
        if name in txt:
            return key
    return None


def _f(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return 0.0


def _cover(v):
    """-> float, or None for the NO SALES badge (no trailing sales at all)."""
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or "").strip().upper()
    if not s or "NO SALES" in s or s in {"-", "—"}:
        return None
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return None


# --------------------------------------------------------------------------
# Bond sheets: the only place with a per-month series
# --------------------------------------------------------------------------

def read_bond_sheets(wb, bonds: list[str], short=None) -> list[dict]:
    """Per shop x SKU monthly Sales and Closing series.

    FULL DETAIL carries only window totals, so anything about WHEN a line died
    -- or when stock was pushed into it -- has to come from here.
    """
    out = []
    for bond in bonds:
        if bond not in wb.sheetnames:
            continue
        ws = wb[bond]
        rows = list(ws.iter_rows(values_only=True))
        hdr_i = next((i for i, r in enumerate(rows)
                      if r and str(r[0]).strip() == "Severity"), None)
        if hdr_i is None:
            continue
        hdr = [str(c).strip() if c is not None else "" for c in rows[hdr_i]]
        sales_cols = [(i, h.split()[0]) for i, h in enumerate(hdr)
                      if h.endswith(" Sales") and not h.startswith("Total")]
        close_cols = [(i, h.split()[0]) for i, h in enumerate(hdr)
                      if h.endswith(" Closing")]
        try:
            c_brand = hdr.index("Brand"); c_pack = hdr.index("Pack")
            c_cover = hdr.index("Months of Cover"); c_zero = hdr.index("Zero Months")
        except ValueError:
            continue

        code = shop = None
        for r in rows[hdr_i + 1:]:
            if not r or r[0] is None:
                continue
            first = str(r[0]).strip()
            tier = _tier(first)
            if tier is None:
                # shop banner: "105029  ·  5029-Kollappally"
                m = re.match(r"^(\d{4,6})\s*·\s*(.+)$", first)
                if m:
                    code = m.group(1)
                    shop = re.sub(r"^\d{4,6}\s*[-–]\s*", "", m.group(2)).strip()
                continue
            if code is None:
                continue
            raw_brand = str(r[c_brand] or "").strip()
            out.append({
                "bond": bond, "code": code, "shop": shop,
                "brand": raw_brand,
                # display name, so the lever tables read "K.S 99" rather than
                # "K.S 99 LIFE TIME MATURED XXX RUM" and stay one line
                "br": short(raw_brand) if short else raw_brand,
                "pack": str(r[c_pack] or "").strip(),
                "tier": tier,
                "sales": {mn: _f(r[i]) for i, mn in sales_cols},
                "close": {mn: _f(r[i]) for i, mn in close_cols},
                "months": [mn for _, mn in sales_cols],
                "cover": _cover(r[c_cover]),
                "zero": int(_f(r[c_zero])),
            })
    return out


def stall_profile(bond_rows: list[dict]) -> dict:
    """When did each aging line last actually sell?

    The chronic measure only reaches back to the first snapshot on record;
    this reaches back to the start of the analysis window and dates the cause.
    """
    aging = [r for r in bond_rows if r["tier"] in ("nm", "cr", "sl")]
    buckets, detail = {}, []
    for r in aging:
        last = None
        for mn in r["months"]:
            if r["sales"].get(mn, 0) > 0.005:
                last = mn
        key = last or "never"
        cs = r["close"].get(r["months"][-1], 0)
        g = buckets.setdefault(key, {"cs": 0.0, "lines": 0})
        g["cs"] += cs
        g["lines"] += 1
        detail.append({"last": key, "cs": cs})
    order = (bond_rows[0]["months"] if bond_rows else [])
    seq = ["never"] + order
    return {"buckets": [{"month": k, "cs": round(buckets[k]["cs"], 2),
                         "lines": buckets[k]["lines"]}
                        for k in seq if k in buckets],
            "total_cs": round(sum(b["cs"] for b in buckets.values()), 2)}


def redispatch(bond_rows: list[dict], from_month: str | None = None) -> dict:
    """Stock pushed INTO lines that had already stopped selling.

    Implied receipts = closing_m - closing_(m-1) + sales_m. A drop counts only
    if the line sold nothing for DEAD_MONTHS consecutive months beforehand --
    deliberately conservative, so a line that sold recently then stalled is
    NOT counted. This sizes the "stop the re-order" lever, and it is the only
    figure in the workbook that does.
    """
    events, tot_in, tot_sold = [], 0.0, 0.0
    for r in bond_rows:
        mo = r["months"]
        start = mo.index(from_month) if from_month in mo else DEAD_MONTHS
        for i in range(max(start, DEAD_MONTHS), len(mo)):
            m, prev = mo[i], mo[i - 1]
            recv = r["close"].get(m, 0) - r["close"].get(prev, 0) + r["sales"].get(m, 0)
            if recv <= 0.5:
                continue
            if any(r["sales"].get(mo[j], 0) > 0.005
                   for j in range(i - DEAD_MONTHS, i)):
                continue          # it was still selling -- a fair re-order
            sold_after = sum(r["sales"].get(mo[j], 0) for j in range(i, len(mo)))
            events.append({"bond": r["bond"], "code": r["code"], "shop": r["shop"],
                           "br": r.get("br", r["brand"]), "p": r["pack"], "month": m,
                           "cs": round(recv, 2), "sold": round(sold_after, 2),
                           "stuck": round(max(0.0, recv - sold_after), 2)})
            tot_in += recv
            tot_sold += sold_after
    events.sort(key=lambda e: -e["stuck"])
    by_bond = {}
    for e in events:
        g = by_bond.setdefault(e["bond"], {"drops": 0, "cs": 0.0, "stuck": 0.0})
        g["drops"] += 1
        g["cs"] += e["cs"]
        g["stuck"] += e["stuck"]
    return {
        "events": events[:20], "n": len(events),
        "cs_in": round(tot_in, 2), "cs_sold": round(tot_sold, 2),
        "cs_stuck": round(tot_in - tot_sold, 2),
        "sell_through": round(tot_sold / tot_in * 100, 1) if tot_in else 0.0,
        "by_bond": sorted(({"bond": b, **{k: round(v, 2) if isinstance(v, float) else v
                                          for k, v in g.items()}}
                           for b, g in by_bond.items()),
                          key=lambda r: -r["stuck"])[:8],
        "from_month": from_month,
    }


# --------------------------------------------------------------------------
# FULL DETAIL: the healthy side of the ledger
# --------------------------------------------------------------------------

def stockouts(all_rows: list[dict], months: float) -> dict:
    """Shops with PROVEN demand and an empty shelf.

    The counterparty to aging, and invisible on an aging-only page: a bond can
    look clean simply because it is running dry. Rate is the line's own
    window sales spread over the window; "empty" is under a quarter month of
    its own consumption left.
    """
    hits = []
    for r in all_rows:
        if r["t"] != "ok":
            continue
        rate = r["sa"] / months if months else 0
        if rate < STOCKOUT_MIN_RATE:
            continue
        if r["cl"] > rate * STOCKOUT_MAX_COVER:
            continue
        hits.append({"bond": r["bond"], "code": r["code"], "shop": r["shop"],
                     "br": r["br"], "p": r["p"], "rate": round(rate, 2),
                     "cl": round(r["cl"], 2)})
    hits.sort(key=lambda r: -r["rate"])
    by_bond = {}
    for h in hits:
        g = by_bond.setdefault(h["bond"], {"n": 0, "rate": 0.0})
        g["n"] += 1
        g["rate"] += h["rate"]
    return {"rows": hits[:20], "n": len(hits),
            "rate": round(sum(h["rate"] for h in hits), 1),
            "by_bond": sorted(({"bond": b, "n": g["n"], "rate": round(g["rate"], 1)}
                               for b, g in by_bond.items()),
                              key=lambda r: -r["rate"])[:8]}


def routing(all_rows: list[dict], months: float) -> dict:
    """Stuck stock whose SKU has a proven buyer in the same bond.

    NOT a shop-to-shop transfer instruction -- STN moves warehouse->shop, not
    shop->shop, so that may not even be legal. It is a RE-ORDER ROUTING
    instruction: stop sending this SKU to this shop, the demand is over here.
    Framed that way it needs no permit.
    """
    groups = {}
    for r in all_rows:
        g = groups.setdefault((r["bond"], r["br"], r["p"]),
                              {"stuck": [], "buyers": []})
        rate = r["sa"] / months if months else 0
        if r["t"] in ("nm", "cr", "sl"):
            g["stuck"].append({"shop": r["shop"], "code": r["code"],
                               "cl": r["cl"], "rate": rate})
        elif r["t"] == "ok" and rate >= BUYER_MIN_RATE:
            cov = r["cl"] / rate if rate else 999
            if cov <= BUYER_MAX_COVER:
                g["buyers"].append({"shop": r["shop"], "code": r["code"],
                                    "rate": rate, "cl": r["cl"]})
    out, tot = [], 0.0
    for (bond, br, p), g in groups.items():
        stuck_cs = sum(s["cl"] for s in g["stuck"])
        buy_rate = sum(b["rate"] for b in g["buyers"])
        if stuck_cs < 1 or buy_rate <= 0:
            continue
        if stuck_cs / buy_rate > ROUTE_MAX_MONTHS:
            continue
        tot += stuck_cs
        out.append({
            "bond": bond, "br": br, "p": p,
            "stuck": round(stuck_cs, 2), "lines": len(g["stuck"]),
            "buy_rate": round(buy_rate, 1),
            "clear_mo": round(stuck_cs / buy_rate, 2),
            "buyers": sorted(g["buyers"], key=lambda b: -b["rate"])[:3],
            "worst": sorted(g["stuck"], key=lambda s: -s["cl"])[0],
        })
    out.sort(key=lambda r: -r["stuck"])
    return {"rows": out[:15], "n": len(out), "cs": round(tot, 2)}


def cover_buckets(all_rows: list[dict]) -> dict:
    """Split the exposure by whether pushing could ever clear it.

    A different conversation from the tier split: tiers say how dead a line
    is, this says whether selling harder is even the right instrument.
    """
    b = {"push": {"cs": 0.0, "n": 0}, "hard": {"cs": 0.0, "n": 0},
         "dead": {"cs": 0.0, "n": 0}}
    for r in all_rows:
        if r["t"] not in ("nm", "cr", "sl"):
            continue
        cov = r.get("cv")
        k = "dead" if cov is None or cov > 24 else ("push" if cov <= PUSHABLE_COVER else "hard")
        b[k]["cs"] += r["cl"]
        b[k]["n"] += 1
    return {k: {"cs": round(v["cs"], 2), "n": v["n"]} for k, v in b.items()}


# --------------------------------------------------------------------------
# NEW STOCK sheet: the only leading indicator in the file
# --------------------------------------------------------------------------

def newstock(wb) -> dict:
    """Placements still inside the 21-day grace, and how they are landing.

    Only the >= NEW_MIN_DAYS cohort is judged: a line that arrived two days
    before the cut-off cannot have sold, and counting it as a failure would
    overstate the risk by roughly a quarter.
    """
    if "NEW STOCK" not in wb.sheetnames:
        return {"n": 0, "cs": 0.0, "risk": [], "risk_n": 0, "risk_cs": 0.0,
                "prior_n": 0, "prior_cs": 0.0, "cohort": 0}
    ws = wb["NEW STOCK"]
    rows = list(ws.iter_rows(values_only=True))
    hdr_i = next((i for i, r in enumerate(rows)
                  if r and str(r[0]).strip() == "Bond"), None)
    if hdr_i is None:
        return {"n": 0, "cs": 0.0, "risk": [], "risk_n": 0, "risk_cs": 0.0,
                "prior_n": 0, "prior_cs": 0.0, "cohort": 0}
    hdr = [str(c).strip() if c is not None else "" for c in rows[hdr_i]]
    idx = {h: i for i, h in enumerate(hdr)}
    close_cols = [i for i, h in enumerate(hdr) if h.endswith(" Closing")]
    sales_col = next((i for i, h in enumerate(hdr) if h.endswith(" Sales")), None)
    prior_col = idx.get("Prior order (pre-window)")
    all_rows, risk = [], []
    for r in rows[hdr_i + 1:]:
        if not r or not r[0]:
            continue
        days = int(_f(r[idx.get("Days on Shelf", 6)]))
        cs = _f(r[close_cols[-1]]) if close_cols else 0.0
        sold = _f(r[sales_col]) if sales_col is not None else 0.0
        prior = str(r[prior_col] or "").strip() if prior_col is not None else "-"
        rec = {"bond": str(r[0]).strip(),
               "code": str(r[idx.get("Shop Code", 1)] or "").strip(),
               "shop": re.sub(r"^\d{4,6}\s*[-–]\s*", "",
                              str(r[idx.get("Shop Name", 2)] or "").strip()),
               "br": str(r[idx.get("Brand", 3)] or "").strip(),
               "p": str(r[idx.get("Pack", 4)] or "").strip(),
               "days": days, "cs": round(cs, 2), "sold": round(sold, 2),
               "prior": prior not in ("-", "", "None")}
        all_rows.append(rec)
        if days >= NEW_MIN_DAYS and sold <= 0.005 and cs > 0.05:
            risk.append(rec)
    risk.sort(key=lambda r: -r["cs"])
    prior = [r for r in all_rows if r["prior"]]
    cohort = [r for r in all_rows if r["days"] >= NEW_MIN_DAYS]
    return {
        "n": len(all_rows), "cs": round(sum(r["cs"] for r in all_rows), 2),
        "cohort": len(cohort),
        "risk": risk[:12], "risk_n": len(risk),
        "risk_cs": round(sum(r["cs"] for r in risk), 2),
        "prior_n": len(prior), "prior_cs": round(sum(r["cs"] for r in prior), 2),
    }


# --------------------------------------------------------------------------
# Incentive tracker join
# --------------------------------------------------------------------------

MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def incentives(root: Path, all_rows: list[dict], brand_short,
               bond_rows: list[dict] | None = None) -> dict:
    """Is the money working?

    Joins the live incentive lines to their current aging position. The
    dashboard has never seen this file, so a rate can sit on a line for five
    months having moved nothing and nobody notices.

    `sold` is deliberately sales SINCE THE INCENTIVE WENT LIVE, taken from the
    bond sheets' monthly series -- not the window total. A line incented from
    9 June that sold 45 cs across Feb-Aug did not sell 45 cs BECAUSE of the
    incentive, and reporting the window total next to the rate would credit
    the money for sales that predate it.
    """
    p = root / "Incentive tracking" / "INCENTIVE TRACKER.xlsx"
    if not p.exists():
        return {"n": 0, "rows": [], "still_n": 0, "still_cs": 0.0, "sold": 0.0}
    try:
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    except Exception:
        return {"n": 0, "rows": [], "still_n": 0, "still_cs": 0.0, "sold": 0.0}
    try:
        if "INCENTIVE LOG" not in wb.sheetnames:
            return {"n": 0, "rows": [], "still_n": 0, "still_cs": 0.0, "sold": 0.0}
        ws = wb["INCENTIVE LOG"]
        rows = list(ws.iter_rows(values_only=True))
        hdr = [str(c).strip() if c is not None else "" for c in rows[0]]
        idx = {h: i for i, h in enumerate(hdr)}
        by_key = {}
        for r in all_rows:
            by_key[(str(r["code"]), r["br"], r["p"])] = r
        series = {}
        for r in (bond_rows or []):
            series[(str(r["code"]), brand_short(r["brand"]), r["pack"])] = r

        def sold_since(key, live_from):
            """Sales in the months at or after the incentive's live-from."""
            s = series.get(key)
            if not s or not live_from:
                return None
            try:
                mon = MONTH_ABBR[int(str(live_from)[5:7]) - 1]
            except (ValueError, IndexError):
                return None
            if mon not in s["months"]:
                return None
            i = s["months"].index(mon)
            return round(sum(s["sales"].get(m, 0) for m in s["months"][i:]), 2)

        out = []
        for r in rows[1:]:
            if not r or not r[idx.get("Shop Code", 2)]:
                continue
            code = str(r[idx["Shop Code"]]).strip()
            br = brand_short(str(r[idx.get("Brand", 6)] or ""))
            pk = str(r[idx.get("Pack", 8)] or "").strip()
            live = str(r[idx.get("Live From", 1)] or "")[:10]
            cur = by_key.get((code, br, pk))
            out.append({
                "code": code,
                "shop": str(r[idx.get("Shop Name", 3)] or "").strip(),
                "bond": str(r[idx.get("Bond", 4)] or "").strip(),
                "br": br, "p": pk,
                "rate": _f(r[idx.get("₹/cs", 10)]),
                "from": live,
                "tier": cur["t"] if cur else "cleared",
                "cl": round(cur["cl"], 2) if cur else 0.0,
                "sold": sold_since((code, br, pk), live),
            })
        still = [o for o in out if o["tier"] in ("nm", "cr", "sl")]
        out.sort(key=lambda o: -o["cl"])
        return {"n": len(out), "rows": out[:14],
                "still_n": len(still),
                "still_cs": round(sum(o["cl"] for o in still), 2),
                "sold": round(sum(o["sold"] or 0 for o in still), 2)}
    finally:
        wb.close()
