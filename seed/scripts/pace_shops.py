"""
Daily pace — named-shop targeting.

A daily case number on its own is a scoreboard. This turns it into a route by
naming the outlets costing the bond cases today, split into the TWO problems
an executive can actually act on:

  SITTING ON STOCK   the shop is holding real stock and it is not selling.
                     Nothing to order — this is a push: work the counter,
                     the display, the staff. Ranked by cases lying idle.

  RUNNING DRY        the shop has little or nothing left of what it normally
                     sells. It cannot sell what it does not have, so this is
                     an indent or an STN, not a pep talk. Ranked by what the
                     shop normally moves in a month, because an empty shelf
                     at a fast outlet costs the most.

A shop can only appear once. RUNNING DRY wins the tie — an empty shelf is
both more urgent and more fixable than a slow one.

Every shop is judged against ITS OWN prior-month rate, never a bond average:
a 12 cs/month outlet should not be measured by a 200 cs/month outlet's
standard.
"""

from __future__ import annotations

import datetime as _dt
import os
import re
import sys
from collections import defaultdict

import openpyxl

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import pace_data as PD


# --- tunables --------------------------------------------------------------

SITTING_MIN_STOCK = 10.0     # cs on hand before idle stock is worth a visit
SITTING_MAX_SELLTHRU = 0.70  # selling under this share of its own pace = slow
DRY_MAX_COVER_DAYS = 10.0    # less than this many days of stock left = dry
DRY_MIN_RATE = 5.0           # cs/month the shop must normally do to matter
TOP_N = 4                    # rows per table per bond


def _combined_sheet(wb, month: str):
    mu = month.upper()
    for sn in wb.sheetnames:
        up = sn.upper()
        if up.startswith(mu) and "COMBINED" in up and "DISPATCH" not in up:
            return wb[sn]
    return None


def read_combined(base: str, month: str) -> dict:
    """(shop_code, brand, pack) -> {opening, receipts, sales, closing} in cases."""
    path = PD._find_month_workbook(os.path.join(base, "KSBC shop sales"),
                                   month, "ANALYSIS")
    if not path:
        return {}
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = _combined_sheet(wb, month)
    if ws is None:
        wb.close()
        return {}
    out = {}
    rows = ws.iter_rows(values_only=True)
    next(rows, None)
    for r in rows:
        if not r or r[1] in (None, ""):
            continue
        try:
            code = PD._canon(int(str(r[1]).strip()))
        except (TypeError, ValueError):
            continue
        brand = str(r[4] or "").strip()
        pack = str(r[5] or "").strip()
        try:
            o, rc, s, c = (float(r[7] or 0), float(r[8] or 0),
                           float(r[9] or 0), float(r[10] or 0))
        except (TypeError, ValueError):
            continue
        k = (code, brand, pack)
        if k in out:
            v = out[k]
            v["opening"] += o; v["receipts"] += rc
            v["sales"] += s;   v["closing"] += c
        else:
            out[k] = {"opening": o, "receipts": rc, "sales": s, "closing": c}
    wb.close()
    return out


def _short_brand(b: str) -> str:
    b = (b or "").upper()
    for key, short in (("OLD PEARL", "OLD PEARL"), ("CHAIRMAN", "CHAIRMAN'S"),
                       ("BLENDER", "BLENDER'S"), ("BCB", "BCB"),
                       ("K.S 99", "K.S 99"), ("KS 99", "K.S 99"),
                       ("MAGIC", "MAGIC BLEND"), ("MORNING", "MORNING WALKER"),
                       ("ROYAL OLD FORT", "ROYAL OLD FORT")):
        if key in b:
            return short
    return b.title()[:16]


def _short_pack(p: str) -> str:
    m = re.search(r"(\d+)", str(p or ""))
    return f"{m.group(1)}ml" if m else str(p or "")


def _short_shop(name: str) -> str:
    n = str(name or "").strip()
    n = re.sub(r"^\s*FL-?0?1\s+", "", n, flags=re.I)
    n = re.sub(r"^\d+\s*[-–]?\s*", "", n)
    n = n.title()
    if len(n) > 26:
        n = n[:26].rstrip()
        # a hard slice can strand an opening bracket -- "Mullackal (Pichu Iyer"
        if n.count("(") > n.count(")"):
            n = n[:n.rfind("(")].rstrip()
    return n


def _shop_rollups(cur, prev):
    """shop_code -> stock / sold / rate, plus per-SKU detail for the notes."""
    cs = defaultdict(lambda: {"stock": 0.0, "sold": 0.0, "receipts": 0.0})
    sku_stock, sku_sold = defaultdict(float), defaultdict(float)
    for (code, b, pk), v in cur.items():
        d = cs[code]
        d["stock"] += v["closing"]
        d["sold"] += v["sales"]
        d["receipts"] += v["receipts"]
        sku_stock[(code, b, pk)] += v["closing"]
        sku_sold[(code, b, pk)] += v["sales"]
    rate = defaultdict(float)
    sku_rate = defaultdict(float)
    for (code, b, pk), v in prev.items():
        rate[code] += v["sales"]
        sku_rate[(code, b, pk)] += v["sales"]
    return cs, rate, sku_stock, sku_sold, sku_rate


def build_shop_signals(base, plan, hist_rows=None, master=None) -> dict:
    """bond -> {"sitting": [...], "dry": [...]}, worst-first in each list."""
    base = base or PD._default_base()
    master = master or PD.load_master(base)
    cur = read_combined(base, plan["month"])
    prev = read_combined(base, plan["prior_month"])
    cs, rate, sku_stock, sku_sold, sku_rate = _shop_rollups(cur, prev)

    elapsed = sum(1 for d in range(1, plan["as_of"].day + 1)
                  if d not in plan["dry_days"])
    frac = elapsed / max(plan["selling_days"], 1)

    sitting, dry = defaultdict(list), defaultdict(list)
    for code, info in master.items():
        if info["status"].upper() == "CLOSED" or info["cat"] != "KSBC":
            continue
        bond = info["bond"]
        if bond not in plan["bonds"]:
            continue
        d = cs.get(code)
        if not d:
            continue
        stock, sold = d["stock"], d["sold"]
        mo_rate = rate.get(code, 0.0)
        expected = mo_rate * frac
        daily = mo_rate / 30.0
        cover = (stock / daily) if daily > 0.01 else (999.0 if stock > 0 else 0.0)
        shop = _short_shop(info["name"])

        # --- RUNNING DRY: shelf is empty at a shop that normally sells ----
        if mo_rate >= DRY_MIN_RATE and cover < DRY_MAX_COVER_DAYS:
            outs = sorted(
                ((r, b, pk) for (c2, b, pk), r in sku_rate.items()
                 if c2 == code and r >= 1.0
                 and sku_stock.get((code, b, pk), 0.0) <= 0.05),
                reverse=True)
            miss = ", ".join(f"{_short_brand(b)} {_short_pack(pk)}"
                             for _, b, pk in outs[:2])
            if len(outs) > 2:
                miss += f" +{len(outs) - 2}"
            dry[bond].append({
                "shop": shop, "code": code, "stock": stock, "sold": sold,
                "rate": mo_rate, "cover": cover,
                "note": miss or "shelf nearly empty",
                "lost": max(mo_rate * (1 - frac), 0.0),
            })
            continue

        # --- SITTING ON STOCK: goods are there, they are not moving -------
        slow = (sold < expected * SITTING_MAX_SELLTHRU) or (sold < 0.05)
        if stock >= SITTING_MIN_STOCK and slow:
            top = sorted(((v, b, pk) for (c2, b, pk), v in sku_stock.items()
                          if c2 == code and v > 0.05), reverse=True)
            heavy = ", ".join(f"{_short_brand(b)} {_short_pack(pk)}"
                              for _, b, pk in top[:2])
            sitting[bond].append({
                "shop": shop, "code": code, "stock": stock, "sold": sold,
                "expected": expected, "rate": mo_rate, "cover": cover,
                "note": heavy,
                "idle": stock,
            })

    out = {}
    for bond in plan["bonds"]:
        out[bond] = {
            "sitting": sorted(sitting.get(bond, []),
                              key=lambda r: -r["idle"])[:TOP_N],
            "dry": sorted(dry.get(bond, []),
                          key=lambda r: -r["rate"])[:TOP_N],
        }
    return out



def build_shop_daily(base, plan, hist_rows=None, master=None) -> dict:
    """
    bond -> [ {shop, cat, now, last, delta, per_day}, ... ] sorted by this
    month's cumulative cases, biggest first.

    CUMULATIVE cases this month against the SAME NUMBER OF SELLING DAYS last
    month — not last month's full total, which would show every shop "down"
    simply because the month isn't over. Like-for-like is the only comparison
    an executive will accept.

    KSBC outlets only. FED and BAR are order-driven invoice dispatches, so a
    per-shop sales figure there is a dispatch record, not a fact about the
    outlet.
    """
    base = base or PD._default_base()
    master = master or PD.load_master(base)
    cur = read_combined(base, plan["month"])

    elapsed = max(sum(1 for d in range(1, plan["as_of"].day + 1)
                      if d not in plan["dry_days"]), 1)

    # this month: authoritative month-to-date from the COMBINED roll-up
    now_s = defaultdict(float)
    for (code, _b, _p), v in cur.items():
        now_s[code] += v["sales"]

    # last month: same elapsed count of live days, from the daily history
    pmno = PD.MONTH_NUM[plan["prior_month"]]
    by_day = defaultdict(lambda: defaultdict(float))
    net_day = defaultdict(float)
    for r in hist_rows or []:
        d = _dt.date.fromisoformat(str(r["date"])[:10])
        if d.month != pmno:
            continue
        by_day[d.day][int(r["shop_code"])] += float(r["tertiary_cs"])
        net_day[d.day] += float(r["tertiary_cs"]) + float(r["invoice_cs"])
    live_prior = sorted(d for d, v in net_day.items() if v >= 1.0)[:elapsed]
    last_s = defaultdict(float)
    for d in live_prior:
        for code, q in by_day[d].items():
            last_s[code] += q

    out = defaultdict(list)
    for code, info in master.items():
        if info["status"].upper() == "CLOSED" or info["cat"] != "KSBC":
            continue
        bond = info["bond"]
        if bond not in plan["bonds"]:
            continue
        now = now_s.get(code, 0.0)
        last = last_s.get(code, 0.0)
        out[bond].append({
            "shop": _short_shop(info["name"]), "code": code,
            "cat": info["cat"], "now": now, "last": last,
            "delta": now - last, "per_day": now / elapsed,
        })

    for rows in out.values():
        seen = defaultdict(int)
        for r in rows:
            seen[r["shop"]] += 1
        for r in rows:
            if seen[r["shop"]] > 1:
                r["shop"] = f"{r['shop']} ({str(r['code'])[-4:]})"
    return {b: sorted(v, key=lambda r: -r["now"]) for b, v in out.items()}, \
           len(live_prior)


def channel_split(hist_rows, plan) -> dict:
    """bond -> {day: [tertiary, invoice]} — the stacking for the bond chart."""
    out = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0]))
    for r in hist_rows:
        d = _dt.date.fromisoformat(str(r["date"])[:10])
        if d.year != plan["year"] or d.month != plan["mno"]:
            continue
        v = out[r["bond"]][d.day]
        v[0] += float(r["tertiary_cs"])
        v[1] += float(r["invoice_cs"])
    return {b: dict(v) for b, v in out.items()}
