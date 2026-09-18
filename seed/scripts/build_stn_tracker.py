#!/usr/bin/env python3
"""
STN CONTROL TOWER v3 — full-lifecycle STN tracking engine.
Created 10 Jul 2026 (replaces v2 arrival-anchored tracker; research-backed redesign).

WHAT AN STN IS (public record, researched 10 Jul 2026):
  STN = Stock Transfer Note under KSBC's "STN Scheme": a SUPPLIER-initiated,
  permit-based transfer of the supplier's own IMFL brands from an FL-9 warehouse
  to specific FL-1 shops (BEVCO homepage notice, Dec 2022 — unrestricted except
  premium brands). Normal replenishment is KSBC-driven (PI/ROQ from trailing
  3-month sales, G.O. 24112/DL/85/TD of 28.10.1985); STN is the supplier's only
  lever to place stock a shop's ROQ won't pull. KSBC levies charges on the
  supplier for the transfer — KSD pays 14% of case rate per case — and KSBC
  itself tracks request -> permit -> transfer -> completion ("STN COMPLETE"
  registers). An STN that reaches the shop but does not liquidate = levy sunk,
  case value locked, and NO PI ceiling gain (MQ rises only from offtake).

DESIGN (v3, locked 10 Jul 2026):
  * LIFECYCLE per line (shop x brand x pack x cases x STN date):
      🕓 IN TRANSIT  applied, no receipt seen yet (days-waiting tracked; ⚠ flag
                     when waiting > TRANSIT_ALERT_DAYS)
      arrival        first Shop In > 0 on/after the STN date (KSBC daily raws)
      liquidation    🚫 STUCK <10% · ⚠️ SLOW 10-30% · ✅ MOVING 30-60% ·
                     🚀 FAST >=60% · 🏁 CLEARED (closing <= 0.01 cs)
  * HEADLINE: % SOLD SINCE ARRIVAL = Sold Since Arrival / Received-effective
    (cap 100%). Pre-existing stock on arrival day is its own visible column,
    never netted.
  * EXPOSURE CAP + CLEARED FREEZE (11 Jul 2026 audit, Abhay-approved):
      Received-effective = min(Received, Applied) — a routine PI indent in the
      window can no longer inflate exposure past the levy actually paid (the
      Received column still shows raw folded receipts, noted). A line FREEZES
      the first day closing reaches ≤ CLEARED_EPS after arrival: 🏁 CLEARED is
      terminal — later routine stock never reopens it.
  * SAME-SKU WINDOW PARTITION (11 Jul 2026 audit): repeat STNs on one
      shop×brand×pack partition the timeline — line i's movement window is
      [stn_i, stn_i+1) (last line open-ended), so one physical receipt/sale is
      never counted into two lines.
  * WASTAGE (cases only — Abhay 10 Jul 2026, no rupee columns):
      Unsold STN (cs) = max(Received-effective − Sold Since Arrival, 0) per
      arrived line.
      🗑 WASTE RISK    = arrived, on shelf >= WASTE_RISK_DAYS, % sold < 30%,
                        and stock still closing — the "levy paid for nothing" set.
  * Days to Arrive (receipt lag), Days on Shelf, Daily Rate, Days to Clear give
    speed context per line.
  * Movement source: KSBC monthly analysis workbooks' daily raw sheets
    (<MONTH> <N>), cases folded with loose bottles via Bottle Per Case; spans
    months. Reconciliation Pre + Received − Sold = Closing per line (tol 0.15).
  * Persistent accumulator: STN LOG sheet of the live workbook — every build
    seeds from it; new entries layer on top. Dedupe key (stn_date, shop, brand,
    pack), new wins. Malformed log rows are counted and REPORTED, never
    silently dropped. NEAR-DUP TRIPWIRE (11 Jul 2026): a new line matching an
    existing one on shop×brand×pack×cases within NEARDUP_DAYS but a different
    STN date aborts as a suspected date-shifted double-entry (the 107009
    corruption class); --allow-neardup overrides, --retire removes a wrong line.
  * ALL 15 bond sheets ALWAYS present (bonds with lines first, ranked by unsold
    STN desc; rest alphabetical with empty-state notice).
  * Sheet order: DASHBOARD · STN TRACKER · 15 bonds · STN LOG · README.
  * Static values only — no recalc step needed.

Usage:
  python3 build_stn_tracker.py [--add-xlsx REQUEST.xlsx]... [--add-json entries.json]
                               [--base DIR] [--out FILE.xlsx] [--as-of YYYY-MM-DD]
                               [--retire "SHOP|BRAND|PACK|YYYY-MM-DD"]...
                               [--allow-neardup] [--allow-gaps] [--no-total-check]
  (--year is accepted but IGNORED — workbook years are auto-assigned by walking
   back from the run clock; cross-year safe since 11 Jul 2026.)

Inputs:
  --add-xlsx  STN REQUEST workbook (matrix template, repeatable; the team's
              standard since 6 Jul 2026). R1 title; R2 WAREHOUSE/DATE/EXEC-ASM;
              R3 brand banners over R4 pack sizes; data rows from R5 to the
              TOTAL (CASES) row; last col = row TOTAL. Row + grand totals are
              reconciled against the parsed sum — mismatch aborts. Template:
              "STN tracking/STN REQUEST TEMPLATE.xlsx"; drops in
              "STN tracking/requests/".
  --add-json  [{"stn_date":"2026-06-29","shop":"5013","brand":"MWB","pack":"L",
               "cases":3,"source":"WhatsApp sojan ASM","notes":""}, ...]
              (fallback for WhatsApp screenshots / pasted lists)
Unknown shop / brand / pack in any input => hard abort (nothing written).
Field quirk: CCB = CHAIRMAN'S CHOICE, never BCB.
"""

import argparse
import json
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------- constants
BASE_DEFAULT = Path(__file__).resolve().parents[2]

MONTHS = {"JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4, "MAY": 5, "JUNE": 6,
          "JULY": 7, "AUGUST": 8, "SEPTEMBER": 9, "OCTOBER": 10, "NOVEMBER": 11, "DECEMBER": 12}
MONTH_BY_NUM = {v: k for k, v in MONTHS.items()}

# Register month tabs are hand-named and Abhay abbreviates them ("AUG", "SEPT").
# 12 Aug 2026: the August register arrived on a tab called AUG and the strict
# `in MONTHS` test rejected it as "no MONTH tab", which would have aborted the
# whole ingest. Abbreviations are matched EXACTLY from this table and are all
# >= 3 chars ON PURPOSE — a prefix match would read the register's own "NO"
# snapshot tab (a July working copy) as NOVEMBER and wholesale-replace a month
# that has nothing to do with it.
MONTH_ALIASES = {"JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "JUN": 6, "JUL": 7,
                 "AUG": 8, "SEP": 9, "SEPT": 9, "OCT": 10, "NOV": 11, "DEC": 12}

def month_of_tab(sheet_name):
    """Month number for a register MONTH tab (full name or known abbreviation),
    else None. Dated snapshot tabs ('24-7-2026', '7th AUG') return None — they
    are ignored by design, so they must not match here."""
    key = re.sub(r"\s+", " ", str(sheet_name or "").strip().upper())
    return MONTHS.get(key) or MONTH_ALIASES.get(key)

FONT = "Aptos Narrow"
NAVY_DEEP, NAVY_MID, NAVY_SOFT = "FF0D1B4A", "FF1A237E", "FF263F80"
GOLD, GOLD_DIM = "FFFFB300", "FFFFD54F"
ZEBRA, GREY_TOTAL, SHOP_BAND = "FFF2F4F8", "FF374151", "FF334466"
RISK_FILL, RISK_FONT = "FFFFEBEE", "FFB71C1C"
# dark dashboard palette (same family as the locked Secondary/Warehouse dashboards)
DK_CANVAS, DK_PANEL, DK_PANEL_ALT, DK_HDR = "FF0D1B4A", "FF14224F", "FF182A5A", "FF223466"
DK_LINE = Side(style="thin", color="FF2A3C74")
DK_BORDER = Border(left=DK_LINE, right=DK_LINE, top=DK_LINE, bottom=DK_LINE)
STEEL, STEEL_LT, RED_BRIGHT, RED_LT = "FF90A4AE", "FFCFD8DC", "FFE53935", "FFEF9A9A"

# lifecycle tiers; worst-first order for sorting
T_STUCK, T_SLOW, T_TRANSIT = "🚫 STUCK", "⚠️ SLOW", "🕓 IN TRANSIT"
T_MOVING, T_FAST, T_CLEAR = "✅ MOVING", "🚀 FAST", "🏁 CLEARED"
TIER_ORDER = [T_STUCK, T_SLOW, T_TRANSIT, T_MOVING, T_FAST, T_CLEAR]
TIER_COLOR = {
    T_STUCK:   ("FFB71C1C", "FFFFEBEE"),
    T_SLOW:    ("FFE65100", "FFFFF3E0"),
    T_TRANSIT: ("FF455A64", "FFECEFF1"),
    T_MOVING:  ("FF2E7D32", "FFDCEDC8"),
    T_FAST:    ("FF1565C0", "FFE3F2FD"),
    T_CLEAR:   ("FF1B5E20", "FFC8E6C9"),
}
FAST_TH, MOVING_TH, SLOW_TH = 0.60, 0.30, 0.10
CLEARED_EPS, DRIFT_TOL = 0.01, 0.15
TRANSIT_ALERT_DAYS = 7    # applied but not arrived after this many days -> ⚠ chase
WASTE_RISK_DAYS = 30      # arrived >= this many days AGO and still <30% sold -> 🗑
NEARDUP_DAYS = 3          # same shop×brand×pack×cases within this window = suspected double-entry
TRANSIT_WATCH_CAP = 15    # dashboard IN-TRANSIT WATCH rows shown (rest summarised) — 15 balances the bottom band (v5)

# canonical brand -> workbook Brand Name variants (2026 daily-sheet universe)
CANON = {
    "BCB CLASSIC":        ["BCB NO.1 CLASSIC BRANDY", "B.C.B NO.1 CLASSIC BRANDY", "BCB NO.1 CLASSIC BRAND"],
    "BLENDER'S CHOICE":   ["BLENDER'S CHOICE NO.1 BRANDY", "BLENDER'S CHOICE NO.1"],
    "CHAIRMAN'S CHOICE":  ["CHAIRMAN'S CHOICE XO BRANDY"],
    "KS 99":              ["K.S 99 LIFE TIME MATURED XXX RUM", "K.S.99 XXX RUM"],
    "KS OLD FORT":        ["K S OLD FORT MAT.XXX RUM"],
    "MAGIC BLEND RUM":    ["MAGIC BLEND RESERVED XXX RUM"],
    "MAGIC BLEND BRANDY": ["MAGIC BLEND NO.1 BRANDY"],
    "MORNING WALKER":     ["MORNING WALKERS XO BRANDY"],
    "OLD PEARL":          ["OLD PEARL NO.1 MATURED XXX RUM"],
    "ROYAL OLD FORT":     ["ROYAL OLD FORT NO.1 XXX RUM"],
    "GREAT WALL":         ["GREAT WALL V.S.O.P BRANDY", "GREAT WALL VSOP BDY"],
    "FRENCH DESIRE":      ["FRENCH DESIRE PREMIUM VODKA"],
    "GOLD MINE":          ["GOLD MINE SPL BRANDY"],
    "DIPLOMAT BRANDY":    ["NO.1 DIPLOMAT CHOICE BRANDY"],
    "DIPLOMAT RUM":       ["NO.1 DIPLOMAT CHOICE XXX RUM"],
    "SAVOY":              ["SAVOY DELUXE XXX RUM"],
    "SPRING TIME":        ["SPRING TIME MAT XXX RUM"],
    "WHITE & WHITE":      ["WHITE & WHITE NO.1 XXX RU", "WHITE & WHITE NO.1 XXX RUM"],
}

def _norm(s):
    return re.sub(r"[^A-Z0-9&]", "", str(s).upper())

WB2CANON = {}
for c, variants in CANON.items():
    WB2CANON[_norm(c)] = c
    for v in variants:
        WB2CANON[_norm(v)] = c

# WhatsApp/field shorthand -> canonical (field quirk: CCB = CHAIRMAN'S CHOICE, not BCB)
ALIASES = {
    "BCB": "BCB CLASSIC", "BCBC": "BCB CLASSIC", "BCBCLASSIC": "BCB CLASSIC",
    "BLEND": "BLENDER'S CHOICE", "BLENDERS": "BLENDER'S CHOICE", "BLENDERSCHOICE": "BLENDER'S CHOICE",
    "CCB": "CHAIRMAN'S CHOICE", "CHAIRMANS": "CHAIRMAN'S CHOICE", "CHAIRMANSCHOICE": "CHAIRMAN'S CHOICE",
    "KS": "KS 99", "KS99": "KS 99",
    "KSOLDFORT": "KS OLD FORT", "OLDFORT": "KS OLD FORT",
    "MB": "MAGIC BLEND RUM", "MBR": "MAGIC BLEND RUM", "MAGICBLEND": "MAGIC BLEND RUM",
    "MBB": "MAGIC BLEND BRANDY", "MAGICBLENDBRANDY": "MAGIC BLEND BRANDY",
    "MWB": "MORNING WALKER", "MORNINGWALKER": "MORNING WALKER", "MORNINGWALKERS": "MORNING WALKER",
    "OP": "OLD PEARL", "OPR": "OLD PEARL", "OLDPEARL": "OLD PEARL",
    "ROF": "ROYAL OLD FORT", "ROFR": "ROYAL OLD FORT", "ROYALOLDFORT": "ROYAL OLD FORT",
    "GW": "GREAT WALL", "GREATWALL": "GREAT WALL",
    "FD": "FRENCH DESIRE", "FRENCHDESIRE": "FRENCH DESIRE",
    "WW": "WHITE & WHITE", "WHITEWHITE": "WHITE & WHITE",
    "GOLDMINE": "GOLD MINE", "SPRINGTIME": "SPRING TIME",
    "DIPLOMATBRANDY": "DIPLOMAT BRANDY", "DIPLOMATRUM": "DIPLOMAT RUM",
}
ALIASES_N = {_norm(k): v for k, v in ALIASES.items()}

PACKS = {
    "N": "180 ML", "NIP": "180 ML", "180": "180 ML", "180ML": "180 ML",
    "P": "375 ML", "PINT": "375 ML", "375": "375 ML", "375ML": "375 ML",
    "H": "500 ML", "HL": "500 ML", "HLTR": "500 ML", "HLF": "500 ML", "500": "500 ML", "500ML": "500 ML",
    "Q": "750 ML", "QUART": "750 ML", "750": "750 ML", "750ML": "750 ML",
    "L": "1000 ML", "LTR": "1000 ML", "1L": "1000 ML", "1000": "1000 ML", "1000ML": "1000 ML",
}
BPC_DEFAULT = {"180 ML": 48, "375 ML": 24, "500 ML": 18, "750 ML": 12, "1000 ML": 9}

SHOP_CODE_ALIAS = {"111014": "11014", "11014": "111014"}  # MUKKAM legacy quirk, both directions

PACK_SORT = {"180 ML": 0, "375 ML": 1, "500 ML": 2, "750 ML": 3, "1000 ML": 4}

# ----------------------------------------------------------------------------- helpers
def canon_brand(s):
    n = _norm(s)
    if not n:
        return None
    if n in ALIASES_N:
        return ALIASES_N[n]
    if n in WB2CANON:
        return WB2CANON[n]
    return None

def canon_pack(s):
    n = re.sub(r"[^A-Z0-9]", "", str(s).upper())
    return PACKS.get(n)

def fmt_cs(v):
    return round(v + 1e-9, 2)

def parse_dt(s):
    if isinstance(s, (datetime, date)):
        return s.date() if isinstance(s, datetime) else s
    s = str(s).strip()
    for f in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d %b %Y", "%d %B %Y"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    raise ValueError(f"unparseable date: {s!r}")

# ----------------------------------------------------------------------------- master
def load_master(base):
    p = base / "MASTER DATA CONFIRMED.xlsx"
    if not p.exists():
        sys.exit(f"ABORT: master not found at {p}")
    try:
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    except Exception as e:
        sys.exit(f"ABORT: cannot open master ({e}) — is it locked by Excel (~$ file)?")
    ws = wb["16-4-25"]
    master = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        code = str(r[3]).strip() if r[3] is not None else ""
        code = re.sub(r"\D", "", code)
        if not code:
            continue  # blank-shop-code rows are never keyed (16 Jun audit rule)
        master[code] = {
            "name": re.sub(r"^\s*(FL-\d+\s+)?\d+[- ]*", "", str(r[4] or "").strip()).strip() or str(r[4] or "").strip(),
            "cat": str(r[5] or "").strip().upper(),
            "staff": str(r[6] or "").strip(),
            "bond": str(r[7] or "").strip().upper(),
            "status": str(r[8] or "").strip(),
        }
    wb.close()
    return master

def resolve_shop(raw, master):
    """Field teams write short shop numbers (e.g. 5002 for 105002 — the WhatsApp
    convention), so probe the '10'/'1' prefixes too. Returns (code, hits): code
    is the resolved shop or None. If the probes hit MORE than one master shop
    the input is AMBIGUOUS — code is None and hits lists the candidates
    (11 Jul 2026 audit: '4001' hits both 104001 ALAPPUZHA and 14001 KANNUR)."""
    d = re.sub(r"\D", "", str(raw))
    hits = []
    for cand in (d, "10" + d, "1" + d, d[1:] if d.startswith("1") else None, SHOP_CODE_ALIAS.get(d)):
        if cand and cand in master and cand not in hits:
            hits.append(cand)
    if len(hits) == 1:
        return hits[0], hits
    return None, hits

# ----------------------------------------------------------------------------- STN LOG seed + entries
LOG_HEADERS = ["Entry Date", "STN Date", "Shop Code", "Shop", "Bond", "Brand", "Pack",
               "Applied (cs)", "Source", "Notes"]

def seed_from_live(live_path):
    """Seed entries from the live workbook's STN LOG. Returns (entries, problems).
    Malformed rows are never silently dropped — they are reported (11 Jul 2026);
    a hand-edited brand/pack is re-canonicalised so it keeps matching movement."""
    if not live_path.exists():
        return [], []
    wb = openpyxl.load_workbook(live_path, read_only=True, data_only=True)
    if "STN LOG" not in wb.sheetnames:
        wb.close()
        return [], []
    ws = wb["STN LOG"]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    wb.close()
    out, problems = [], []
    for i, r in enumerate(rows, 2):
        if all(v is None or str(v).strip() == "" for v in r):
            continue  # genuinely blank spacer row
        if r[1] is None or r[2] is None:
            problems.append(f"STN LOG row {i}: blank STN date/shop code — row NOT seeded (fix the live log)")
            continue
        try:
            stn_d = parse_dt(r[1])
            entry_d = parse_dt(r[0]) if r[0] else None
        except ValueError as ex:
            sys.exit(f"ABORT — STN LOG row {i}: {ex} (fix the live log cell; nothing written)")
        brand_raw, pack_raw = str(r[5]).strip(), str(r[6]).strip()
        brand = canon_brand(brand_raw) or brand_raw
        pack = canon_pack(pack_raw) or pack_raw
        if brand != brand_raw or pack != pack_raw:
            problems.append(f"STN LOG row {i}: normalised {brand_raw!r}/{pack_raw!r} → {brand!r}/{pack!r}")
        out.append({
            "entry_date": entry_d,
            "stn_date": stn_d,
            "code": re.sub(r"\D", "", str(r[2])),
            "brand": brand,
            "pack": pack,
            "cases": float(r[7] or 0),
            "source": str(r[8] or "").strip(),
            "notes": str(r[9] or "").strip() if len(r) > 9 and r[9] else "",
        })
    return out, problems

def load_new_entries(json_path, master):
    try:
        raw = json.loads(Path(json_path).read_text(encoding="utf-8"))
    except FileNotFoundError:
        sys.exit(f"ABORT — --add-json file not found: {json_path}")
    except (json.JSONDecodeError, OSError) as ex:
        sys.exit(f"ABORT — cannot read --add-json {json_path}: {ex}")
    entries, errors = [], []
    for i, e in enumerate(raw, 1):
        code, cands = resolve_shop(e.get("shop", ""), master)
        brand = canon_brand(e.get("brand", ""))
        pack = canon_pack(e.get("pack", ""))
        try:
            stn_d = parse_dt(e["stn_date"])
            cases = float(e["cases"])
        except Exception as ex:
            errors.append(f"line {i}: bad date/cases ({ex})")
            continue
        if cases <= 0:
            errors.append(f"line {i}: cases must be > 0 (got {cases:g})")
        if code is None and len(cands) > 1:
            errors.append(f"line {i}: ambiguous shop {e.get('shop')!r} → " +
                          " / ".join(f"{c} {master[c]['name']} ({master[c]['bond']})" for c in cands))
        elif code is None:
            errors.append(f"line {i}: unknown shop {e.get('shop')!r}")
        if brand is None:
            errors.append(f"line {i}: unknown brand {e.get('brand')!r}")
        if pack is None:
            errors.append(f"line {i}: unknown pack {e.get('pack')!r}")
        if code and brand and pack and cases > 0:
            entries.append({"entry_date": date.today(), "stn_date": stn_d, "code": code,
                            "brand": brand, "pack": pack, "cases": cases,
                            "source": str(e.get("source", "")).strip(),
                            "notes": str(e.get("notes", "")).strip()})
    if errors:
        sys.exit("ABORT — input validation failed, nothing written:\n  " + "\n  ".join(errors))
    return entries

def merge_entries(seeded, new):
    store = {}
    for e in seeded:
        store[(e["stn_date"], e["code"], _norm(e["brand"]), e["pack"])] = e
    replaced, seen_new, intra = 0, set(), []
    for e in new:
        k = (e["stn_date"], e["code"], _norm(e["brand"]), e["pack"])
        if k in seen_new:
            intra.append(f"{e['code']} {e['brand']} {e['pack']} @ {e['stn_date']}")
        seen_new.add(k)
        if k in store:
            replaced += 1
        store[k] = e
    return sorted(store.values(), key=lambda x: (x["stn_date"], x["code"], x["brand"])), replaced, intra

def find_near_dups(entries, new_keys):
    """Suspected date-shifted double entries (11 Jul 2026 tripwire): same
    shop×brand×pack×cases, different STN dates within NEARDUP_DAYS, at least one
    side ingested THIS run (historical pairs never block a plain refresh)."""
    sus = []
    by_sku = defaultdict(list)
    for e in entries:
        by_sku[(e["code"], _norm(e["brand"]), e["pack"])].append(e)
    for group in by_sku.values():
        group.sort(key=lambda x: x["stn_date"])
        for i in range(len(group)):
            for j in range(i + 1, len(group)):
                a, b = group[i], group[j]
                dd = (b["stn_date"] - a["stn_date"]).days
                if dd == 0 or dd > NEARDUP_DAYS or abs(a["cases"] - b["cases"]) > 0.01:
                    continue
                ka = (a["stn_date"], a["code"], _norm(a["brand"]), a["pack"])
                kb = (b["stn_date"], b["code"], _norm(b["brand"]), b["pack"])
                if ka in new_keys or kb in new_keys:
                    sus.append(f"{a['code']} {a['brand']} {a['pack']} {a['cases']:g} cs @ "
                               f"{a['stn_date']} vs {b['stn_date']}")
    return sus

# ----------------------------------------------------------------------------- STN REQUEST workbook parser
def parse_request_xlsx(path, master, strict_totals=True):
    """Parse one filled STN REQUEST workbook (matrix template, locked 2 Jul 2026).
    Returns (entries, errors, warnings). Caller aborts on errors. Same-key rows
    within ONE file are SUMMED (split-row-loss fix, 11 Jul 2026); strict_totals
    escalates a fully-unverifiable totals state to an error."""
    path = Path(path)
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except FileNotFoundError:
        return [], [f"{path.name}: file not found"], []
    except Exception as ex:
        return [], [f"{path.name}: cannot open ({ex.__class__.__name__}: {ex})"], []
    ws = wb["STN REQUEST"] if "STN REQUEST" in wb.sheetnames else wb.worksheets[0]
    entries, errors, warns = [], [], []

    # header row 2: value = first non-empty cell right of each label
    row2 = [ws.cell(row=2, column=c).value for c in range(1, ws.max_column + 1)]
    def _after(label):
        idx = next((i for i, v in enumerate(row2) if v and label in str(v).upper()), None)
        if idx is None:
            return None
        for v in row2[idx + 1:]:
            if v is None or str(v).strip() == "":
                continue
            if any(l in str(v).upper() for l in ("WAREHOUSE", "DATE", "EXEC")):
                return None  # ran into the next label — value cell empty
            return v
        return None
    warehouse, dt_raw, execasm = _after("WAREHOUSE"), _after("DATE"), _after("EXEC")
    try:
        stn_d = parse_dt(dt_raw) if dt_raw is not None else None
    except ValueError:
        stn_d = None
    if stn_d is None:
        errors.append(f"{path.name}: DATE cell empty/unreadable (got {dt_raw!r})")
        stn_d = date.today()  # placeholder; errors abort anyway
    source = " · ".join(s for s in (str(warehouse or "").strip(), str(execasm or "").strip()) if s) or path.name

    # column map: R3 brand banners (forward-filled across merges) over R4 pack sizes
    colmap, banner, total_col = {}, None, None
    for c in range(3, ws.max_column + 1):
        b = ws.cell(row=3, column=c).value
        if b is not None and str(b).strip():
            banner = str(b).strip()
        if banner and "TOTAL" in banner.upper():
            total_col = total_col or c
            continue
        p = ws.cell(row=4, column=c).value
        if banner is None or p is None or not str(p).strip():
            continue
        br, pk = canon_brand(banner), canon_pack(p)
        if br is None:
            errors.append(f"{path.name}: unknown brand banner {banner!r} (col {get_column_letter(c)})")
            continue
        if pk is None:
            errors.append(f"{path.name}: unknown pack {p!r} under {banner!r} (col {get_column_letter(c)})")
            continue
        colmap[c] = (br, pk)
    if not colmap:
        errors.append(f"{path.name}: no brand/pack columns recognised — is this the STN REQUEST template?")

    # data rows: R5 .. row before TOTAL (CASES)
    total_row = next((rr for rr in range(5, ws.max_row + 1)
                      if ws.cell(row=rr, column=1).value and "TOTAL" in str(ws.cell(row=rr, column=1).value).upper()),
                     ws.max_row + 1)
    file_sum = 0.0
    acc, row_checked, grand_checked, rows_with_qty = {}, 0, False, 0
    for rr in range(5, total_row):
        shop_raw = ws.cell(row=rr, column=1).value
        place = str(ws.cell(row=rr, column=2).value or "").strip()
        qty = {}
        for c, bp in colmap.items():
            v = ws.cell(row=rr, column=c).value
            try:
                f = float(v)
            except (TypeError, ValueError):
                if v not in (None, ""):
                    warns.append(f"{path.name} R{rr}: non-numeric {v!r} in {get_column_letter(c)} ignored")
                continue
            if f > 0:
                qty[bp] = qty.get(bp, 0.0) + f
        if not qty:
            continue
        if shop_raw is None or not re.sub(r"\D", "", str(shop_raw)):
            errors.append(f"{path.name} R{rr}: cases entered but SHOP NO blank (place {place!r})")
            continue
        code, cands = resolve_shop(shop_raw, master)
        if code is None:
            if len(cands) > 1:
                errors.append(f"{path.name} R{rr}: ambiguous shop {shop_raw!r} → " +
                              " / ".join(f"{c} {master[c]['name']} ({master[c]['bond']})" for c in cands))
            else:
                errors.append(f"{path.name} R{rr}: unknown shop {shop_raw!r} (place {place!r})")
            continue
        rows_with_qty += 1
        row_sum = sum(qty.values())
        if total_col:
            tv = ws.cell(row=rr, column=total_col).value
            try:
                tvf = float(tv)
            except (TypeError, ValueError):
                warns.append(f"{path.name} R{rr}: row TOTAL not cached — check skipped")
            else:
                row_checked += 1
                if abs(row_sum - tvf) > 0.01:
                    errors.append(f"{path.name} R{rr}: parsed {row_sum:g} cs ≠ row TOTAL {tvf:g}")
        mname = master[code]["name"]
        if place and _norm(place)[:6] not in _norm(mname) and _norm(mname)[:6] not in _norm(place):
            warns.append(f"{path.name} R{rr}: place {place!r} vs master {mname!r} — verify shop no")
        for (br, pk), f in sorted(qty.items()):
            file_sum += f
            k = (code, br, pk)
            if k in acc:  # same shop×brand×pack on ANOTHER row of this file — SUM, never drop
                acc[k]["cases"] += f
                acc[k]["rows"] += 1
            else:
                acc[k] = {"entry_date": date.today(), "stn_date": stn_d, "code": code,
                          "brand": br, "pack": pk, "cases": f, "source": source,
                          "notes": "", "rows": 1}
    # grand-total reconciliation
    if total_col and total_row <= ws.max_row:
        gv = ws.cell(row=total_row, column=total_col).value
        try:
            gvf = float(gv)
        except (TypeError, ValueError):
            warns.append(f"{path.name}: grand TOTAL not cached — check skipped")
        else:
            grand_checked = True
            if abs(file_sum - gvf) > 0.01:
                errors.append(f"{path.name}: parsed grand total {file_sum:g} cs ≠ sheet TOTAL {gvf:g}")
    # stray numbers below the TOTAL row are ignored but flagged
    for rr in range(total_row + 1, min(ws.max_row, total_row + 300) + 1):
        if any(isinstance(ws.cell(row=rr, column=c).value, (int, float)) and ws.cell(row=rr, column=c).value
               for c in colmap):
            warns.append(f"{path.name}: stray value(s) below TOTAL row (R{rr}) — ignored")
            break
    for e in sorted(acc.values(), key=lambda x: (x["code"], x["brand"], x["pack"])):
        n_rows = e.pop("rows")
        if n_rows > 1:
            e["notes"] = "summed from split rows in request"
            warns.append(f"{path.name}: {e['code']} {e['brand']} {e['pack']} on {n_rows} rows — "
                         f"SUMMED to {e['cases']:g} cs")
        entries.append(e)
    if strict_totals and rows_with_qty and row_checked == 0 and not grand_checked:
        errors.append(f"{path.name}: NO totals could be verified (TOTAL row/col missing or formulas "
                      "uncached) — open & save the file in Excel to cache totals, or pass --no-total-check")
    wb.close()
    if not entries and not errors:
        warns.append(f"{path.name}: no request lines found (all quantities zero/blank)")
    return entries, errors, warns

# ----------------------------------------------------------------------------- STN REGISTER parser (Abhay's format, 17 Jul 2026)
CLUSTER_BONDS = {
    "1": {"ALAPPUZHA", "ATTINGAL", "NEDUMANGAD", "KOLLAM", "KOTTARAKARA", "PATHANAMTHITTA"},
    "2": {"THRISSUR", "TRIPUNITHURA", "THODUPUZHA", "KOTTAYAM", "ALUVA"},
    "3": {"KANNUR", "KOZHIKODE", "PALAKKAD", "PERINTHALMANNA"},
}

def _dm_swap_fix(sub_d, req_d):
    """Submitted-before-requested with a valid day/month swap -> the swapped date
    (the register's known '6-7 typed as Jun-7' quirk). None if no clean fix."""
    if sub_d and req_d and sub_d < req_d:
        try:
            sw = date(sub_d.year, sub_d.day, sub_d.month)
        except ValueError:
            return None
        if sw >= req_d:
            return sw
    return None

def resolve_shop_ctx(raw, master, cluster=None):
    """resolve_shop + CLUSTER NO disambiguation: '4001' hits 104001 (ALAPPUZHA,
    C1) and 14001 (KANNUR, C3) — the row's cluster picks the right one."""
    code, hits = resolve_shop(raw, master)
    if code is None and len(hits) > 1 and cluster:
        cl = re.sub(r"\D", "", str(cluster))
        cand = [h for h in hits if master[h]["bond"] in CLUSTER_BONDS.get(cl, set())]
        if len(cand) == 1:
            return cand[0], hits
    return code, hits

def parse_register_xlsx(path, master):
    """Parse Abhay's STN REGISTER workbook. Reads MONTH tabs only (JULY, ...);
    dated snapshot tabs are ignored. One row -> up to TWO tracker lines:
      line 1: SUBMITTED QTY @ SUBMITTED DATE   (the levy-bearing application)
      line 2: PENDING SUBMITTED QTY @ PENDING SUBMITTED DATE (balance submitted later)
    REQUESTED QTY is field context (never applied). PENDING QTY without a
    pending-submission date is HELD (warned, not ingested) — tracker starts at
    submission (Abhay, 17 Jul 2026). Returns (entries, errors, warns, months)."""
    path = Path(path)
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except FileNotFoundError:
        return [], [f"{path.name}: file not found"], [], set()
    except Exception as ex:
        return [], [f"{path.name}: cannot open ({ex.__class__.__name__}: {ex})"], [], set()
    entries, errors, warns, months = {}, [], [], set()

    def _emit(code, brand, pack, cases, stn_d, src, note):
        k = (stn_d, code, _norm(brand), pack)
        if k in entries:  # two field requests submitted same day for one SKU — SUM
            entries[k]["cases"] += cases
            extra = "merged 2 register rows (same SKU submitted same day)"
            if extra not in entries[k]["notes"]:
                entries[k]["notes"] = (entries[k]["notes"] + "; " if entries[k]["notes"] else "") + extra
            warns.append(f"{path.name}: {code} {brand} {pack} @ {stn_d} — {extra}, now {entries[k]['cases']:g} cs")
        else:
            entries[k] = {"entry_date": date.today(), "stn_date": stn_d, "code": code,
                          "brand": brand, "pack": pack, "cases": cases, "source": src,
                          "notes": note}
        # NOTE: months for the wholesale reconcile come from the TAB NAME
        # (see below), never from this set — a single stray-dated row must not
        # be able to delete a whole month of logged lines.

    month_tabs = [sn for sn in wb.sheetnames if month_of_tab(sn)]
    # 31 Jul 2026: the reconcile scope is the set of MONTH TABS present, resolved
    # against the same year assignment the lines use. Deriving it from emitted
    # line dates meant one row carrying a stray June date inside the JULY tab
    # would silently wipe every logged June line (15 lines / 118 cs today),
    # reported only as an informational "replaced" count.
    if not month_tabs:
        errors.append(f"{path.name}: no MONTH tab (JULY/AUGUST/…) found — dated snapshot tabs are ignored by design")
    for sn in month_tabs:
        ws = wb[sn]
        # DUPLICATE HEADERS are resolved by DATA, not by position (31 Jul 2026).
        # A register tab can carry a stray copy of a header (a leftover "PENDING
        # SUBMITTED DATE" in col O beside the real one in col M). The original
        # last-wins map silently redirected the column to the empty copy and 15
        # dated lines (115 cs) were held out of the build as "no date". Leftmost-
        # wins fixed that file but is only a positional guess — it fails the moment
        # the stray sits to the LEFT. So: keep the copy that actually carries data,
        # and ABORT if two populated copies disagree on any row.
        _seen, _dupes = {}, {}
        for idx, c in enumerate(ws[1], 0):
            key = re.sub(r"\s+", " ", str(c.value or "").strip().upper())
            if not key:
                continue
            if key in _seen:
                _dupes.setdefault(key, [_seen[key]]).append(idx)
            else:
                _seen[key] = idx
        hdr = dict(_seen)
        for key, cols in _dupes.items():
            filled = {c: sum(1 for r in ws.iter_rows(min_row=2, min_col=c + 1, max_col=c + 1,
                                                     values_only=True)
                             if r[0] is not None and str(r[0]).strip() != "")
                      for c in cols}
            populated = [c for c in cols if filled[c] > 0]
            conflict = False
            if len(populated) > 1:
                grids = [[r[0] for r in ws.iter_rows(min_row=2, min_col=c + 1, max_col=c + 1,
                                                     values_only=True)] for c in populated]
                for row_vals in zip(*grids):
                    vals = {v for v in row_vals if v is not None and str(v).strip() != ""}
                    if len(vals) > 1:
                        conflict = True
                        break
            if conflict:
                errors.append(
                    f"{path.name}/{sn}: header {key!r} appears in columns "
                    f"{', '.join(get_column_letter(c + 1) for c in populated)} and they hold "
                    f"CONFLICTING values — delete the stray column and rebuild")
                continue
            best = max(cols, key=lambda c: (filled[c], -c))
            hdr[key] = best
            warns.append(
                f"{path.name}/{sn}: duplicate header {key!r} in columns "
                f"{', '.join(get_column_letter(c + 1) for c in cols)} — using column "
                f"{get_column_letter(best + 1)} ({filled[best]} filled cell(s)); "
                f"delete the stray column to silence this")

        def col(*keys):
            for k in keys:
                for h, i in hdr.items():
                    if h.startswith(k):
                        return i
            return None
        i_req_d = col("REQUESTED DATE")
        i_clu = col("CLUSTER")
        i_wh = col("WAREHOUSE")
        i_shop = col("SHOP NO")
        i_brand = col("BRAND")
        i_pack = col("PACK")
        i_req_q = col("REQUESTED QTY", "REQUESTEDQTY")
        i_sub_q = col("SUBMITTED QTY")
        i_sub_d = col("SUBMITTED DATE")
        i_pen_q = col("PENDING QTY")
        i_psub_q = col("PENDING SUBMITTED QTY")
        i_psub_d = col("PENDING SUBMITTED DATE")
        i_status = col("STATUS")
        missing = [n for n, i in (("SHOP NO", i_shop), ("BRAND NAME", i_brand), ("PACK", i_pack),
                                  ("SUBMITTED QTY", i_sub_q), ("SUBMITTED DATE", i_sub_d)) if i is None]
        if missing:
            errors.append(f"{path.name}/{sn}: register headers missing: {', '.join(missing)}")
            continue
        for rr, r in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
            def g(i):
                return r[i] if i is not None and i < len(r) else None
            shop_raw = g(i_shop)
            first = str(g(0) or "").strip().upper()
            if first == "TOTAL":
                continue
            if shop_raw is None:
                # 31 Jul 2026: only skip a genuinely empty row. A row carrying a
                # brand/pack/quantity with the shop cell cleared used to disappear
                # silently — exactly the kind of edit that loses cases.
                if any(str(g(i) or "").strip() for i in (i_brand, i_pack) if i is not None) or \
                   any((g(i) or 0) for i in (i_sub_q, i_psub_q, i_req_q) if i is not None):
                    errors.append(f"{path.name}/{sn} R{rr}: blank SHOP NO on a row that has "
                                  f"brand/pack/quantity — fill the shop code or clear the row")
                continue
            brand_raw = str(g(i_brand) or "").strip()
            pack_raw = str(g(i_pack) or "").strip()
            brand = canon_brand(brand_raw)
            pack = canon_pack(pack_raw)
            if brand is None or (pack is None and not pack_raw):
                m = re.match(r"(.+?)\s*[-–]?\s*(\d{3,4})\s*ML\s*$", brand_raw, re.I)
                if m:  # 'CCB-750ml' style — brand cell carries the pack
                    brand = brand or canon_brand(m.group(1))
                    pack = pack or canon_pack(m.group(2))
            wh = str(g(i_wh) or "").strip()
            code, cands = resolve_shop_ctx(shop_raw, master, cluster=g(i_clu))
            if code is None:
                if len(cands) > 1:
                    errors.append(f"{path.name}/{sn} R{rr}: ambiguous shop {shop_raw!r} → " +
                                  " / ".join(f"{c} {master[c]['name']} ({master[c]['bond']})" for c in cands))
                else:
                    errors.append(f"{path.name}/{sn} R{rr}: unknown shop {shop_raw!r} (WH {wh!r})")
                continue
            if brand is None:
                errors.append(f"{path.name}/{sn} R{rr}: unknown brand {brand_raw!r}")
                continue
            if pack is None:
                errors.append(f"{path.name}/{sn} R{rr}: unknown pack {pack_raw!r} (brand {brand_raw!r})")
                continue
            try:
                req_d = parse_dt(g(i_req_d)) if g(i_req_d) else None
            except ValueError:
                req_d = None
            src = f"STN register {sn} · WH {wh}" if wh else f"STN register {sn}"

            def num(v, label=""):
                # 31 Jul 2026: a non-numeric or negative quantity used to read as
                # 0.0 and the line silently vanished from the build with no error
                # at all. The --add-json path already aborts on cases <= 0; the
                # register path had no equivalent.
                if v is None or str(v).strip() == "":
                    return 0.0
                try:
                    f = float(v)
                except (TypeError, ValueError):
                    errors.append(f"{path.name}/{sn} R{rr}: non-numeric {label} {v!r}")
                    return 0.0
                if f < 0:
                    errors.append(f"{path.name}/{sn} R{rr}: negative {label} {f:g}")
                    return 0.0
                return f
            req_q, sub_q = num(g(i_req_q), "REQUESTED QTY"), num(g(i_sub_q), "SUBMITTED QTY")
            pen_q, psub_q = num(g(i_pen_q), "PENDING QTY"), num(g(i_psub_q), "PENDING SUBMITTED QTY")
            status = str(g(i_status) or "").strip().upper()
            # identity checks (warn-only — the register is Abhay's hand-kept file)
            if req_q and abs(req_q - (sub_q + pen_q)) > 0.01 and not psub_q:
                warns.append(f"{path.name}/{sn} R{rr}: requested {req_q:g} ≠ submitted {sub_q:g} + pending {pen_q:g}")
            if psub_q > pen_q + 0.01:
                warns.append(f"{path.name}/{sn} R{rr}: pending-submitted {psub_q:g} > pending {pen_q:g}")
            # line 1 — the initial submission
            if sub_q > 0:
                sub_d = None
                try:
                    sub_d = parse_dt(g(i_sub_d)) if g(i_sub_d) else None
                except ValueError:
                    pass
                note = ""
                if sub_d is None:
                    if req_d is None:
                        errors.append(f"{path.name}/{sn} R{rr}: submitted {sub_q:g} cs but no usable date")
                        continue
                    sub_d = req_d
                    note = "submitted date blank — used requested date"
                    warns.append(f"{path.name}/{sn} R{rr}: {note}")
                else:
                    fix = _dm_swap_fix(sub_d, req_d)
                    if fix:
                        warns.append(f"{path.name}/{sn} R{rr}: submitted date {sub_d} before requested {req_d} — D/M swap fixed to {fix}")
                        sub_d = fix
                        note = "submitted date D/M-swap corrected"
                    elif req_d and sub_d < req_d:
                        # impossible date, no clean swap — floor at the requested date
                        # (portal-verified 17 Jul 2026: the '6-7 vs 7-7' batch was
                        # submitted same-day as requested)
                        warns.append(f"{path.name}/{sn} R{rr}: submitted date {sub_d} impossible (before "
                                     f"requested {req_d}), no clean swap — used requested date")
                        sub_d = req_d
                        note = "impossible submitted date — floored to requested date"
                _emit(code, brand, pack, sub_q, sub_d, src, note)
            # line 2 — the pending balance submitted later
            if psub_q > 0:
                psub_d = None
                try:
                    psub_d = parse_dt(g(i_psub_d)) if g(i_psub_d) else None
                except ValueError:
                    pass
                if psub_d is None:
                    warns.append(f"{path.name}/{sn} R{rr}: pending-submitted {psub_q:g} cs has NO date — "
                                 "HELD as pending (fill PENDING SUBMITTED DATE to ingest)")
                else:
                    fix = _dm_swap_fix(psub_d, req_d)
                    if fix:
                        warns.append(f"{path.name}/{sn} R{rr}: pending-submitted date D/M swap fixed to {fix}")
                        psub_d = fix
                    elif req_d and psub_d < req_d:
                        warns.append(f"{path.name}/{sn} R{rr}: pending-submitted date {psub_d} impossible — used requested date")
                        psub_d = req_d
                    _emit(code, brand, pack, psub_q, psub_d, src, "balance of a partially-submitted request")
            if sub_q <= 0 and psub_q <= 0 and (pen_q > 0 or status == "PENDING"):
                warns.append(f"{path.name}/{sn} R{rr}: {code} {brand} {pack} — requested, NOT YET submitted ({pen_q:g} cs pending) — not ingested")
    wb.close()
    out = sorted(entries.values(), key=lambda x: (x["stn_date"], x["code"], x["brand"]))
    # RECONCILE SCOPE (31 Jul 2026) — only months that actually have a TAB in this
    # register can be wholesale-replaced. Years still come from the data, but a
    # month with no tab of its own can never enter the set, so a stray June-dated
    # row inside the JULY tab cannot delete the logged June lines.
    tab_months = {month_of_tab(sn) for sn in month_tabs}
    for e in out:
        if e["stn_date"].month in tab_months:
            months.add((e["stn_date"].year, e["stn_date"].month))
    off_tab = sorted({(e["stn_date"].year, e["stn_date"].month) for e in out
                      if e["stn_date"].month not in tab_months})
    for y, m in off_tab:
        warns.append(f"{path.name}: {sum(1 for e in out if (e['stn_date'].year, e['stn_date'].month) == (y, m))} "
                     f"line(s) dated {MONTH_BY_NUM[m]} {y} but there is no {MONTH_BY_NUM[m]} tab — "
                     f"ingested, but that month's existing log is NOT replaced")
    if not out and not errors:
        warns.append(f"{path.name}: no submitted register lines found")
    return out, errors, warns, months

# ----------------------------------------------------------------------------- movement
def discover_workbooks(base):
    folder = base / "KSBC shop sales"
    best = {}  # month_num -> (rank, day, path)
    for p in folder.glob("*.xlsx"):
        if p.name.startswith("~$"):
            continue
        u = p.name.upper()
        m = re.fullmatch(r"([A-Z]+) SHOP SALES ANALYSIS\.XLSX", u)
        if m and m.group(1) in MONTHS:
            k = MONTHS[m.group(1)]
            if best.get(k, (0, 0, None))[0] < 2:
                best[k] = (2, 99, p)
            continue
        m = re.fullmatch(r"([A-Z]+) 1ST - (\d+)(?:ST|ND|RD|TH) ANALYSIS\.XLSX", u)
        if m and m.group(1) in MONTHS:
            k, n = MONTHS[m.group(1)], int(m.group(2))
            cur = best.get(k, (0, 0, None))
            if cur[0] < 2 and n > cur[1]:
                best[k] = (1, n, p)
    return {k: v[2] for k, v in best.items()}

def _assign_years(paths, anchor):
    """Month-named workbooks carry no year — assign each the most recent year
    such that (year, month) <= (anchor.year, anchor.month). A January run maps
    DECEMBER → anchor.year−1 (the aging-stock _window_years pattern; fixes the
    silent Dec→Jan corruption found in the 11 Jul 2026 audit)."""
    out = {}
    for mnum, p in paths.items():
        y = anchor.year if mnum <= anchor.month else anchor.year - 1
        out[(y, mnum)] = p
    return out

def _mr_names(months_read):
    yrs = {y for y, m in months_read}
    if len(yrs) > 1:
        return [f"{MONTH_BY_NUM[m]} {y}" for y, m in months_read]
    return [MONTH_BY_NUM[m] for y, m in months_read]

def assert_month_coverage(from_key, latest, months_read, allow_gaps):
    """Every month from the earliest STN to the latest sales day must have a
    discovered workbook — a silent interior gap misattributes arrivals and
    lifecycle states with NO drift flag (11 Jul 2026 audit, PI CHECK-8 pattern)."""
    if latest is None:
        return
    have = set(months_read)
    span, y, m = [], from_key[0], from_key[1]
    while (y, m) <= (latest.year, latest.month):
        span.append((y, m))
        m += 1
        if m == 13:
            y, m = y + 1, 1
    missing = [f"{MONTH_BY_NUM[mm]} {yy}" for (yy, mm) in span if (yy, mm) not in have]
    if missing:
        msg = ("KSBC workbook coverage gap — month(s) missing from the movement window: "
               + ", ".join(missing) + ". Arrivals/sales in those months would be silently "
               "misattributed. Restore the workbook(s) or pass --allow-gaps.")
        if allow_gaps:
            print(f"  ~ WARNING (--allow-gaps): {msg}")
        else:
            sys.exit(f"ABORT — {msg}")

def read_movement(base, from_key, anchor, shop_codes):
    paths = _assign_years(discover_workbooks(base), anchor)
    mov = defaultdict(dict)  # (code, canon_brand, pack) -> {date: [o,i,s,c]}
    latest, unmapped, months_read = None, set(), []
    for (yy, mnum) in sorted(paths):
        if (yy, mnum) < from_key:
            continue
        mname = MONTH_BY_NUM[mnum]
        wb = openpyxl.load_workbook(paths[(yy, mnum)], read_only=True, data_only=True)
        pat = re.compile(rf"{mname} (\d+)$")
        found = False
        for sn in wb.sheetnames:
            m = pat.fullmatch(sn)
            if not m:
                continue
            found = True
            try:
                d = date(yy, mnum, int(m.group(1)))
            except ValueError:
                continue
            latest = d if latest is None or d > latest else latest
            for r in wb[sn].iter_rows(min_row=2, values_only=True):
                if r[1] is None:
                    continue
                code = re.sub(r"\D", "", str(r[1]))
                if code not in shop_codes:
                    code = SHOP_CODE_ALIAS.get(code, code)
                    if code not in shop_codes:
                        continue
                b = str(r[4] or "").strip()
                cb = WB2CANON.get(_norm(b))
                if cb is None:
                    unmapped.add(b)
                    cb = b
                pack = re.sub(r"\s+", " ", str(r[5] or "").strip().upper())
                try:
                    bpc = float(r[6]) if r[6] else 0
                except (TypeError, ValueError):
                    bpc = 0
                bpc = bpc or BPC_DEFAULT.get(pack, 12)

                def fold(ci, bi):
                    try:
                        return float(r[ci] or 0) + float(r[bi] or 0) / bpc
                    except (TypeError, ValueError):
                        return 0.0

                vals = [fold(7, 8), fold(9, 10), fold(11, 12), fold(13, 14)]
                k = (code, cb, pack)
                cur = mov[k].get(d)
                if cur:
                    for j in range(4):
                        cur[j] += vals[j]
                else:
                    mov[k][d] = vals
        wb.close()
        if found:
            months_read.append((yy, mnum))
    return mov, latest, unmapped, months_read

# ----------------------------------------------------------------------------- compute
def compute_line(e, mov, latest, master, window_end=None):
    rec = master.get(e["code"], {})
    line = {
        "code": e["code"], "shop": rec.get("name", e["code"]), "bond": rec.get("bond", "?"),
        "staff": rec.get("staff", ""), "brand": e["brand"], "pack": e["pack"],
        "stn_date": e["stn_date"], "applied": e["cases"], "source": e["source"],
        "entry_date": e.get("entry_date"), "log_notes": e.get("notes", ""),
        "arrival": None, "lag": None, "days": None, "received": 0.0, "pre": None,
        "sold": 0.0, "pct": None, "unsold": 0.0, "closing": None, "rate": None,
        "d2c": None, "tier": T_TRANSIT, "waste_risk": False, "transit_alert": False,
        "wait": None, "notes": [], "frozen": False, "lapsed": False,
    }
    if line["log_notes"]:
        line["notes"].append(line["log_notes"])
    if rec.get("status", "").lower() == "closed":
        line["notes"].append("CLOSED SHOP")
    series = mov.get((e["code"], e["brand"], e["pack"]), {})
    if series:
        line["closing"] = fmt_cs(series[max(series)][3])
    if latest is None or e["stn_date"] > latest:
        line["notes"].append(f"awaiting data (have through {latest.strftime('%d %b') if latest else '—'})")
        if line["closing"] is not None and line["closing"] > 0.005:
            line["notes"].append(f"shelf stock {line['closing']:g} cs predates STN")
        return line
    # SAME-SKU WINDOW PARTITION (11 Jul 2026): this line only sees days in
    # [stn_date, window_end) — a later STN on the same SKU owns the days after.
    eff_latest = latest if window_end is None else min(latest, window_end - timedelta(days=1))
    wdays = sorted(d for d in series
                   if d >= e["stn_date"] and (window_end is None or d < window_end))
    arrival = next((d for d in wdays if series[d][1] > 0.005), None)
    if arrival is None:
        line["received"] = fmt_cs(sum(series[d][1] for d in wdays))
        line["wait"] = max((eff_latest - e["stn_date"]).days, 0)
        if line["wait"] >= TRANSIT_ALERT_DAYS:
            line["transit_alert"] = True
            line["notes"].append(f"⚠ {line['wait']}d since STN, no receipt — chase warehouse/permit")
        if window_end is not None and window_end <= latest:
            # 31 Jul 2026: a later STN on this SKU took over the timeline, so this
            # line can NEVER arrive — it is not "on the way". Flagged so the
            # IN TRANSIT tile can disclose how much of itself is already dead.
            line["lapsed"] = True
            line["notes"].append(f"window closed by newer STN ({window_end.strftime('%d %b')}) — never arrived")
        if not series:
            line["notes"].append("SKU not in shop's raw rows yet")
        elif line["closing"] is not None and line["closing"] > 0.005:
            line["notes"].append(f"shelf stock {line['closing']:g} cs predates STN")
        return line
    line["arrival"] = arrival
    line["lag"] = (arrival - e["stn_date"]).days
    # CLEARED FREEZE (11 Jul 2026): the line ends the first day closing reaches
    # ~zero after arrival — later routine stock never reopens it (terminal).
    # 31 Jul 2026 (Abhay-approved): the freeze must not latch onto a token
    # pre-delivery receipt. A line only clears once cumulative receipts have
    # actually reached the applied quantity. PAYYOLI 111017 BLENDER'S 180 ML was
    # the case that surfaced it — a 2 cs routine receipt on 9 Jul cleared a 15 cs
    # STN on 10 Jul, so the real 15 cs landing on 15 Jul was never seen. Network
    # cost of the old rule: 281 cs of receipts and 217 cs of sales orphaned across
    # 22 lines that were being reported as clean wins.
    _cum, _cum_at = 0.0, {}
    for _d in wdays:
        _cum += series[_d][1]
        _cum_at[_d] = _cum
    _target = e["cases"] - 0.01
    clear_day = next((d for d in wdays
                      if d >= arrival and _cum_at[d] >= _target
                      and -DRIFT_TOL <= series[d][3] <= CLEARED_EPS), None)
    if clear_day is None and (not wdays or _cum_at[wdays[-1]] < _target):
        # ONLY when the applied quantity never landed in this window at all does a
        # genuinely short delivery close on the original first-zero rule. Gating
        # this on the per-day test instead of the window TOTAL re-created the very
        # bug the fix was for: when the consignment landed but the shop never sold
        # down to zero afterwards, no day satisfied both conditions and the line
        # cleared on the token receipt again (12 lines, 149 cs, 31 Jul 2026).
        # Combined effect of both stages: 281 cs of receipts and 217 cs of sales
        # recovered across 22 lines that had been reported as clean wins.
        clear_day = next((d for d in wdays
                          if d >= arrival and -DRIFT_TOL <= series[d][3] <= CLEARED_EPS), None)
    if clear_day is not None:
        line["frozen"] = True
        wdays = [d for d in wdays if d <= clear_day]
        eff_latest = clear_day
    line["days"] = (eff_latest - arrival).days + 1
    line["pre"] = fmt_cs(series[arrival][0])
    received_raw = fmt_cs(sum(series[d][1] for d in wdays))
    line["received"] = received_raw
    line["sold"] = fmt_cs(sum(series[d][2] for d in wdays if d >= arrival))
    line["closing"] = fmt_cs(series[wdays[-1]][3])
    if window_end is not None and window_end <= latest and not line["frozen"]:
        line["notes"].append(f"window ends {(window_end - timedelta(days=1)).strftime('%d %b')} — newer STN on this SKU")
    if received_raw < e["cases"] - 0.01:
        line["notes"].append(f"partial: {received_raw:g} of {e['cases']:g} cs landed")
    elif received_raw > e["cases"] + 0.01:
        line["notes"].append(f"received {received_raw:g} cs > applied (regular indent mixed in — exposure capped at applied)")
    # EXPOSURE CAP (11 Jul 2026): the STN can only waste the levy actually paid.
    received_eff = min(received_raw, e["cases"])
    # 31 Jul 2026 (Abhay-approved): % Sold is measured on the WHOLE pool that
    # moved. The numerator counts every case that left the shelf — including the
    # shop's own routine indent — so the denominator must count them too.
    # sold/received_raw is algebraically identical to pro-rating our share
    # (sold x received_eff/received_raw) / received_eff, and is simpler to state.
    # Old basis (sold/received_eff) put 61 lines at a phantom 100% and left 49
    # lines one tier too high.
    raw_pct = line["sold"] / received_raw if received_raw > 0 else 0.0
    if raw_pct > 1.0:
        line["notes"].append("sold > received — pre-existing stock also moved")
    line["pct"] = min(raw_pct, 1.0)
    # our exposure is our cases; what is still unsold is our share of the pool
    # that did not move. Subtracting ALL sales from our cases understated it.
    line["unsold"] = fmt_cs(max(received_eff * (1.0 - line["pct"]), 0.0))
    rate = line["sold"] / line["days"] if line["days"] else 0
    line["rate"] = round(rate, 2) if rate > 0 else 0.0
    closing = line["closing"] if line["closing"] is not None else 0.0
    if closing > CLEARED_EPS:
        line["d2c"] = round(closing / rate) if rate > 0.005 else None
        if line["sold"] <= 0.005:
            line["notes"].append("NO SALES since arrival")
    # reconciliation: pre + received - sold == closing (telescoped across the window)
    if line["closing"] is not None and line["pre"] is not None:
        drift = abs(line["pre"] + received_raw - line["sold"] - line["closing"])
        if drift > DRIFT_TOL:
            line["notes"].append(f"⚠ drift {drift:.2f} cs (O+R−S≠C)")
    if line["frozen"] or (-DRIFT_TOL <= closing <= CLEARED_EPS):
        line["tier"] = T_CLEAR
        line["unsold"] = 0.0  # nothing on shelf — the STN did its job
        line["d2c"] = line["days"]  # actual days it took to clear
        if line["frozen"] and clear_day < latest:
            line["notes"].append(f"cleared {clear_day.strftime('%d %b')} — line frozen (terminal)")
    elif closing < -DRIFT_TOL:
        line["notes"].append(f"⚠ negative closing {closing:g} cs in raw (data glitch) — tier from % sold")
        line["tier"] = (T_FAST if line["pct"] >= FAST_TH else T_MOVING if line["pct"] >= MOVING_TH
                        else T_SLOW if line["pct"] >= SLOW_TH else T_STUCK)
    elif line["pct"] >= FAST_TH:
        line["tier"] = T_FAST
    elif line["pct"] >= MOVING_TH:
        line["tier"] = T_MOVING
    elif line["pct"] >= SLOW_TH:
        line["tier"] = T_SLOW
    else:
        line["tier"] = T_STUCK
    if line["tier"] == T_FAST and closing > received_raw:
        line["notes"].append("position still heavy — closing exceeds STN qty")
    if (line["tier"] in (T_STUCK, T_SLOW) and line["days"] is not None
            and (line["days"] - 1) >= WASTE_RISK_DAYS and closing > CLEARED_EPS):
        line["waste_risk"] = True
        line["notes"].append(f"🗑 WASTE RISK — {line['days']}d on shelf, {line['pct']:.0%} sold")
    return line

def sort_lines(lines):
    def key(l):
        t = TIER_ORDER.index(l["tier"])
        exposure = l["unsold"] if l["tier"] != T_TRANSIT else l["applied"]
        return (t, -exposure, l["bond"], l["shop"], l["brand"], PACK_SORT.get(l["pack"], 9))
    return sorted(lines, key=key)

# ----------------------------------------------------------------------------- styling helpers
THIN = Side(style="thin", color="FFD0D5DD")
B_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
GOLD_BOT = Border(bottom=Side(style="thin", color="FFFFD700"))

def _f(size=10, bold=False, color="FF111827", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)

def _fill(hex8):
    return PatternFill("solid", fgColor=hex8)

def _c(ws, row, col, val, font=None, fill=None, align="center", fmt=None, border=B_ALL, indent=0, wrap=False):
    cell = ws.cell(row=row, column=col, value=val)
    cell.font = font or _f()
    if fill:
        cell.fill = _fill(fill) if isinstance(fill, str) else fill
    cell.alignment = Alignment(horizontal=align, vertical="center", indent=indent, wrap_text=wrap)
    if fmt:
        cell.number_format = fmt
    if border:
        cell.border = border
    return cell

def hero(ws, ncols, title, subtitle):
    last = get_column_letter(ncols)
    ws.merge_cells(f"A1:{last}1")
    ws.merge_cells(f"A2:{last}2")
    for col in range(1, ncols + 1):
        for rr in (1, 2):
            ws.cell(row=rr, column=col).fill = _fill(NAVY_DEEP)
    _c(ws, 1, 1, title, _f(20, True, "FFFFFFFF"), NAVY_DEEP, "left", border=None, indent=1)
    _c(ws, 2, 1, subtitle, _f(11, False, GOLD_DIM), NAVY_DEEP, "left", border=None, indent=1)
    ws.row_dimensions[1].height = 34
    ws.row_dimensions[2].height = 20
    ws.sheet_view.showGridLines = False

def kpi_tiles(ws, row, tiles, start_col=1, width=2, spans=None):
    """tiles = [(label, value, sub[, accent_hex[, value_hex[, numfmt]]])].
    Numeric values are written as REAL numbers (no green 'text' triangles).
    spans = explicit [(c1, c2), ...] column spans; else width-`width` tiles with
    one gap column, starting at start_col."""
    if spans is None:
        spans, col = [], start_col
        for _ in tiles:
            spans.append((col, col + width - 1))
            col += width + 1
    edge = Side(style="thin", color="FF4A5FA8")
    for t, (c1, c2) in zip(tiles, spans):
        label, value, sub = t[0], t[1], t[2]
        accent = t[3] if len(t) > 3 and t[3] else GOLD
        valcol = t[4] if len(t) > 4 and t[4] else "FFFFFFFF"
        numfmt = t[5] if len(t) > 5 and t[5] else "#,##0.##"
        for rr in (row, row + 1, row + 2):
            if c2 > c1:
                ws.merge_cells(start_row=rr, start_column=c1, end_row=rr, end_column=c2)
        for cc in range(c1, c2 + 1):
            for rr in (row, row + 1, row + 2):
                cell = ws.cell(row=rr, column=cc)
                cell.fill = _fill(NAVY_MID)
                cell.border = Border(
                    left=Side(style="medium", color=accent) if cc == c1 else None,
                    right=edge if cc == c2 else None,
                    top=edge if rr == row else None,
                    bottom=edge if rr == row + 2 else None)
        _c(ws, row, c1, label, _f(9, True, GOLD_DIM), None, "center", border=None)
        vc = _c(ws, row + 1, c1, value, _f(18, True, valcol), None, "center", border=None)
        if isinstance(value, (int, float)):
            vc.number_format = numfmt
        _c(ws, row + 2, c1, sub, _f(9, False, "FFB9C4E8"), None, "center", border=None)
        # re-assert the card edges the label/value/sub writes just cleared
        for rr, (tp, bt) in ((row, (edge, None)), (row + 1, (None, None)), (row + 2, (None, edge))):
            ws.cell(row=rr, column=c1).border = Border(
                left=Side(style="medium", color=accent), top=tp, bottom=bt,
                right=edge if c1 == c2 else None)
            if c2 > c1:
                ws.cell(row=rr, column=c2).border = Border(right=edge, top=tp, bottom=bt)
    ws.row_dimensions[row].height = 16
    ws.row_dimensions[row + 1].height = 30
    ws.row_dimensions[row + 2].height = 15

def dark_band(ws, r1, r2, ncols):
    """Dark-canvas band behind the KPI tiles so hero + tiles read as one header block."""
    for rr in range(r1, r2 + 1):
        for cc in range(1, ncols + 1):
            ws.cell(row=rr, column=cc).fill = _fill(DK_CANVAS)

def gold_rule(ws, row, ncols, height=4):
    """Thin gold strip row — the brand accent under every hero band."""
    for cc in range(1, ncols + 1):
        ws.cell(row=row, column=cc).fill = _fill(GOLD)
    ws.row_dimensions[row].height = height

def _logo_png(base):
    """KSD shield with the black flood knocked out, sized for the hero band.
    Returns a temp PNG path or None (non-fatal — never blocks a build)."""
    try:
        from collections import deque
        from PIL import Image as PILImage
        src = base / "Internal Docs" / "KSD LOGO.png"
        if not src.exists():
            return None
        im = PILImage.open(src).convert("RGBA")
        scale = 200 / im.height
        im = im.resize((max(1, int(im.width * scale)), 200), PILImage.LANCZOS)
        px, (w, h) = im.load(), im.size
        seen, dq = set(), deque([(0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1)])
        while dq:
            x, y = dq.popleft()
            if (x, y) in seen or not (0 <= x < w and 0 <= y < h):
                continue
            seen.add((x, y))
            p = px[x, y]
            if p[3] == 0 or p[0] > 46 or p[1] > 46 or p[2] > 46:
                continue
            px[x, y] = (0, 0, 0, 0)
            dq.extend([(x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)])
        im = im.resize((max(1, int(w * 52 / h)), 52), PILImage.LANCZOS)
        import os
        import tempfile
        fd, out = tempfile.mkstemp(prefix="ksd_logo_stn_", suffix=".png")
        os.close(fd)
        im.save(out)
        return out
    except Exception:
        return None

def table_header(ws, row, headers):
    """CLASSIC KSD GOLD-NAVY header (Abhay 17 Jul 2026 — 'the classic ksd gold
    navy style'): deep navy fill, gold bold lettering, gold rule beneath."""
    for j, h in enumerate(headers, 1):
        _c(ws, row, j, h, _f(10, True, GOLD_DIM), NAVY_DEEP, "center",
           border=Border(bottom=Side(style="thin", color=GOLD)), wrap=True)
    ws.row_dimensions[row].height = 26

def status_cell(ws, row, col, tier):
    dark, light = TIER_COLOR[tier]
    _c(ws, row, col, tier, _f(10, True, dark), light, "center")

def section_title(ws, row, text):
    _c(ws, row, 1, text, _f(12, True, NAVY_DEEP), None, "left", border=None)
    ws.row_dimensions[row].height = 22

# ----------------------------------------------------------------------------- aggregation
def agg_block(lines):
    arrived = [l for l in lines if l["arrival"]]
    applied = sum(l["applied"] for l in lines)
    received = sum(l["received"] for l in lines)
    sold = sum(l["sold"] for l in arrived)
    unsold = sum(l["unsold"] for l in arrived)
    transit = [l for l in lines if l["tier"] == T_TRANSIT]
    transit_cs = sum(l["applied"] for l in transit)
    risk = [l for l in lines if l["waste_risk"]]
    risk_cs = sum(l["unsold"] for l in risk)
    return {
        "n": len(lines), "arrived": len(arrived), "applied": applied, "received": received,
        "sold": sold, "unsold": unsold, "pct": (sold / received if received > 0 else None),
        "transit_n": len(transit), "transit_cs": transit_cs,
        "risk_n": len(risk), "risk_cs": risk_cs,
        "lags": [l["lag"] for l in arrived if l["lag"] is not None],
        "waits": [l["wait"] for l in transit if l.get("wait") is not None],
        "lapsed_n": sum(1 for l in transit if l.get("lapsed")),
        "lapsed_cs": sum(l["applied"] for l in transit if l.get("lapsed")),
    }

def dk_section(ws, row, text, ncols=10):
    """Gold-caps section title on the dark canvas with a gold hairline under the full table width."""
    for cc in range(1, ncols + 1):
        ws.cell(row=row, column=cc).border = Border(bottom=Side(style="thin", color=GOLD))
    _c(ws, row, 1, text, _f(12, True, GOLD_DIM), None, "left", border=Border(bottom=Side(style="thin", color=GOLD)))
    ws.row_dimensions[row].height = 24
    return row + 1

def dk_header(ws, row, headers, start=1):
    for j, h in enumerate(headers, start):
        _c(ws, row, j, h, _f(9, True, "FFB9C4E8"), DK_HDR, "center", border=DK_BORDER, wrap=True)
    ws.row_dimensions[row].height = 22

def dk_cell(ws, row, col, val, i=0, bold=False, color="FFFFFFFF", fmt=None, align="center"):
    return _c(ws, row, col, val, _f(10, bold, color), DK_PANEL_ALT if i % 2 else DK_PANEL,
              align, fmt, border=DK_BORDER)

CS0 = '#,##0;-#,##0;"\u00b7"'          # whole cases · zero renders as a subtle dot
CS1 = '#,##0.0;-#,##0.0;"\u00b7"'      # fractional cases (Sold / Unsold)
# 31 Jul 2026: one decimal. At 0dp a 9.67% STUCK line printed "10%" against a
# tier defined as <10%, and a 29.6% SLOW line printed "30%" — the tiers were
# right, the displayed number contradicted them.
PCT = '0.0%;-0.0%;"\u00b7"'
INK, INK_SOFT, INK_DIM = "FF111827", "FF374151", "FF6B7280"
AMBER_DARK, AMBER_TINT = "FFE65100", "FFFFF3E0"
WASH = "FFF5F7FA"   # soft page background — frames the dashboard on the light canvas

def _h(ws, row, h):
    cur = ws.row_dimensions[row].height
    ws.row_dimensions[row].height = max(cur or 0, h)

def _paint_wash(ws, ncols):
    """Two-pass page wash: every still-unstyled cell inside the dashboard area
    gets the soft grey-blue WASH, so the white table cards read as a bounded,
    designed page instead of floating in the empty sheet (v5, 17 Jul 2026)."""
    for rr in range(1, ws.max_row + 3):
        for cc in range(1, ncols + 1):
            cell = ws.cell(row=rr, column=cc)
            if cell.fill is None or cell.fill.fill_type is None:
                cell.fill = _fill(WASH)

def dk_sec(ws, row, c1, c2, text):
    """LIGHT section title: navy caps over a gold hairline (bond-sheet family)."""
    for cc in range(c1, c2 + 1):
        ws.cell(row=row, column=cc).border = Border(bottom=Side(style="thin", color=GOLD))
    _c(ws, row, c1, text, _f(13, True, NAVY_DEEP), None, "left",
       border=Border(bottom=Side(style="thin", color=GOLD)))
    _h(ws, row, 28)
    return row + 1

def panel_header(ws, row, c1, labels):
    cc = c1
    for text, span in labels:
        if span > 1:
            ws.merge_cells(start_row=row, start_column=cc, end_row=row, end_column=cc + span - 1)
        _c(ws, row, cc, text, _f(9, True, GOLD_DIM), NAVY_DEEP, "center",
           border=Border(bottom=Side(style="thin", color=GOLD)), wrap=True)
        for k in range(1, span):
            cell = ws.cell(row=row, column=cc + k)
            cell.fill = _fill(NAVY_DEEP)
            cell.border = Border(bottom=Side(style="thin", color=GOLD))
        cc += span
    _h(ws, row, 24)
    return row + 1

def pcell(ws, row, col, val, i=0, span=1, bold=False, color=INK, fmt=None, align="center",
          size=10, fill=None):
    zfill = fill if fill else (ZEBRA if i % 2 else "FFFFFFFF")
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    c = _c(ws, row, col, val, _f(size, bold, color), zfill, align, fmt, border=B_ALL)
    for k in range(1, span):
        cell = ws.cell(row=row, column=col + k)
        cell.fill = _fill(zfill)
        cell.border = B_ALL
    return c

def chip(ws, row, col, tier, span=1, size=10):
    dark, light = TIER_COLOR[tier]
    if span > 1:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    _c(ws, row, col, tier, _f(size, True, dark), light, "center", border=B_ALL)
    for k in range(1, span):
        cell = ws.cell(row=row, column=col + k)
        cell.fill = _fill(light)
        cell.border = B_ALL

def group_panel(ws, row, c1, title, label, lines, keyfn):
    """8-col light panel: Label(2) | Applied | Received | Sold | % Sold | Unsold | In Transit."""
    row = dk_sec(ws, row, c1, c1 + 7, title)
    row = panel_header(ws, row, c1, [(label, 2), ("Applied", 1), ("Received", 1), ("Sold", 1),
                                     ("% Sold", 1), ("Unsold STN", 1), ("In Transit", 1)])
    groups = defaultdict(list)
    for l in lines:
        groups[keyfn(l)].append(l)
    ranked = sorted(groups.items(), key=lambda kv: (-agg_block(kv[1])["unsold"], -agg_block(kv[1])["transit_cs"]))
    for i, (g, gl) in enumerate(ranked):
        a = agg_block(gl)
        pcell(ws, row, c1, g, i, span=2, bold=True, color=NAVY_MID, align="left")
        pcell(ws, row, c1 + 2, round(a["applied"], 2), i, fmt=CS0)
        pcell(ws, row, c1 + 3, round(a["received"], 2), i, fmt=CS0)
        pcell(ws, row, c1 + 4, round(a["sold"], 2), i, color=INK_SOFT, fmt=CS1)
        pcell(ws, row, c1 + 5, (min(a["pct"], 1.0) if a["pct"] is not None else None), i,
              color=INK_SOFT, fmt=PCT)
        pcell(ws, row, c1 + 6, round(a["unsold"], 2), i, bold=a["unsold"] > 0,
              color=AMBER_DARK if a["unsold"] > 0 else INK_DIM, fmt=CS1,
              fill=AMBER_TINT if a["unsold"] > 0 else None)
        pcell(ws, row, c1 + 7, round(a["transit_cs"], 2), i, color=INK_DIM, fmt=CS0)
        _h(ws, row, 20)
        row += 1
    return row + 2

def lifecycle_panel(ws, row, c1, lines):
    row = dk_sec(ws, row, c1, c1 + 7, "BY LIFECYCLE STAGE")
    row = panel_header(ws, row, c1, [("Stage", 2), ("Lines", 1), ("Applied", 1), ("Received", 1),
                                     ("Sold", 1), ("% Sold", 1), ("Unsold STN", 1)])
    i = 0
    for t in TIER_ORDER:
        tl = [l for l in lines if l["tier"] == t]
        if not tl:
            continue
        chip(ws, row, c1, t, span=2)
        rec_t = sum(l["received"] for l in tl)
        sold_t = sum(l["sold"] for l in tl)
        pcell(ws, row, c1 + 2, len(tl), i, fmt=CS0)
        pcell(ws, row, c1 + 3, round(sum(l["applied"] for l in tl), 2), i, fmt=CS0)
        pcell(ws, row, c1 + 4, round(rec_t, 2), i, fmt=CS0)
        pcell(ws, row, c1 + 5, round(sold_t, 2), i, color=INK_SOFT, fmt=CS1)
        pcell(ws, row, c1 + 6, (min(sold_t / rec_t, 1.0) if rec_t > 0 else None), i,
              color=INK_SOFT, fmt=PCT)
        un = sum(l["unsold"] for l in tl)
        pcell(ws, row, c1 + 7, round(un, 2), i, bold=un > 0,
              color=AMBER_DARK if un > 0 else INK_DIM, fmt=CS1,
              fill=AMBER_TINT if un > 0 else None)
        _h(ws, row, 20)
        row += 1
        i += 1
    return row + 2

def watch_panel(ws, row, c1, transit):
    row = dk_sec(ws, row, c1, c1 + 7, "IN-TRANSIT WATCH — applied, not received")
    row = panel_header(ws, row, c1, [("Shop", 2), ("Brand", 2), ("Pack", 1), ("Applied", 1),
                                     ("Days Waiting", 1), ("Flag", 1)])
    for i, l in enumerate(transit[:TRANSIT_WATCH_CAP]):
        pcell(ws, row, c1, f"{l['shop']}  ({l['code']})", i, span=2, align="left", size=9, color=INK_SOFT)
        pcell(ws, row, c1 + 2, l["brand"], i, span=2, align="left", size=9, color=INK_SOFT)
        pcell(ws, row, c1 + 4, l["pack"], i, size=9, color=INK_SOFT)
        pcell(ws, row, c1 + 5, l["applied"], i, fmt=CS0)
        pcell(ws, row, c1 + 6, l["wait"], i, bold=l["transit_alert"],
              color=AMBER_DARK if l["transit_alert"] else INK, fmt="0")
        if l["transit_alert"]:
            pcell(ws, row, c1 + 7, "⚠ CHASE", i, bold=True, color=AMBER_DARK, size=9, fill=AMBER_TINT)
        else:
            pcell(ws, row, c1 + 7, "watch", i, color=INK_DIM, size=9)
        _h(ws, row, 20)
        row += 1
    if len(transit) > TRANSIT_WATCH_CAP:
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c1 + 7)
        _c(ws, row, c1, f"… and {len(transit) - TRANSIT_WATCH_CAP} more in transit — see STN TRACKER",
           _f(9, False, INK_DIM), ZEBRA, "left", border=B_ALL)
        for k in range(1, 8):
            cell = ws.cell(row=row, column=c1 + k)
            cell.fill = _fill(ZEBRA)
            cell.border = B_ALL
        _h(ws, row, 18)
        row += 1
    return row + 2

def unsold_panel(ws, row, c1, worst):
    row = dk_sec(ws, row, c1, c1 + 7, "TOP UNSOLD STN POSITIONS")
    row = panel_header(ws, row, c1, [("Shop", 2), ("Brand", 2), ("Pack", 1), ("Unsold STN", 1),
                                     ("% Sold", 1), ("Status", 1)])
    for i, l in enumerate(worst[:15]):
        pcell(ws, row, c1, f"{l['shop']}  ({l['code']})", i, span=2, align="left", size=9, color=INK_SOFT)
        pcell(ws, row, c1 + 2, l["brand"], i, span=2, align="left", size=9, color=INK_SOFT)
        pcell(ws, row, c1 + 4, l["pack"], i, size=9, color=INK_SOFT)
        pcell(ws, row, c1 + 5, l["unsold"], i, bold=True,
              color="FFB71C1C" if l["waste_risk"] else AMBER_DARK, fmt=CS1,
              fill="FFFFEBEE" if l["waste_risk"] else AMBER_TINT)
        pcell(ws, row, c1 + 6, l["pct"], i, color=INK_SOFT, fmt=PCT)
        chip(ws, row, c1 + 7, l["tier"], size=9)
        _h(ws, row, 20)
        row += 1
    # 31 Jul 2026: declare the overflow. The panel shows 15 of N positions; the
    # twin IN-TRANSIT WATCH already says "… and N more", this one said nothing
    # while hiding the large majority of the exposure.
    if len(worst) > 15:
        hidden = sum(l["unsold"] for l in worst[15:])
        _c(ws, row, c1, f"… and {len(worst) - 15} more positions holding {hidden:,.1f} cs — see STN TRACKER",
           _f(9, False, INK_DIM), None, "left")
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c1 + 7)
        _h(ws, row, 18)
        row += 1
    return row + 2

def _lum(hex8):
    r, g, b = int(hex8[2:4], 16), int(hex8[4:6], 16), int(hex8[6:8], 16)
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255

def light_tiles(ws, row, tiles, spans=None):
    """KPI CARDS with a filled colored HEADER BAND (17 Jul 2026 beautify pass):
    accent-colored label strip on top (auto-contrast text), big value below on
    white, grey sub-line, gold hairline splitting header from body, hairline box.
    Funnel trio carries the gold band; exposure tiles their semantic colors."""
    if spans is None:
        spans, col = [], 1
        for _ in tiles:
            spans.append((col, col + 1))
            col += 3
    edge = Side(style="thin", color="FFCAD2E0")
    gold_hair = Side(style="thin", color=GOLD)
    for t, (c1, c2) in zip(tiles, spans):
        label, value, sub = t[0], t[1], t[2]
        accent = t[3] if len(t) > 3 and t[3] else GOLD
        valcol = t[4] if len(t) > 4 and t[4] else NAVY_DEEP
        numfmt = t[5] if len(t) > 5 and t[5] else "#,##0.##"
        hdr_txt = NAVY_DEEP if _lum(accent) > 0.6 else "FFFFFFFF"
        for rr in (row, row + 1, row + 2):
            if c2 > c1:
                ws.merge_cells(start_row=rr, start_column=c1, end_row=rr, end_column=c2)
        # header band
        for cc in range(c1, c2 + 1):
            cell = ws.cell(row=row, column=cc)
            cell.fill = _fill(accent)
        # body white
        for rr in (row + 1, row + 2):
            for cc in range(c1, c2 + 1):
                ws.cell(row=rr, column=cc).fill = _fill("FFFFFFFF")
        _c(ws, row, c1, label, _f(9, True, hdr_txt), accent, "center", border=None)
        vc = _c(ws, row + 1, c1, value, _f(21, True, valcol), "FFFFFFFF", "center", border=None)
        if isinstance(value, (int, float)):
            vc.number_format = numfmt
        _c(ws, row + 2, c1, sub, _f(9, False, INK_DIM), "FFFFFFFF", "center", border=None)
        # box: hairline all around, gold hairline under the header band
        for cc in range(c1, c2 + 1):
            top = edge
            bottom_hdr = gold_hair
            ws.cell(row=row, column=cc).border = Border(
                left=edge if cc == c1 else None, right=edge if cc == c2 else None,
                top=top, bottom=bottom_hdr)
            ws.cell(row=row + 1, column=cc).border = Border(
                left=edge if cc == c1 else None, right=edge if cc == c2 else None)
            ws.cell(row=row + 2, column=cc).border = Border(
                left=edge if cc == c1 else None, right=edge if cc == c2 else None, bottom=edge)
    ws.row_dimensions[row].height = 17
    ws.row_dimensions[row + 1].height = 32
    ws.row_dimensions[row + 2].height = 15

# ----------------------------------------------------------------------------- sheets
TRACKER_HEADERS = ["Shop Code", "Shop", "Bond", "Brand", "Pack", "STN Date", "Arrival",
                   "Lag (d)", "Shelf (d)", "Applied (cs)", "Received (cs)", "Pre-Existing (cs)",
                   "Sold Since (cs)", "% Sold", "Unsold STN (cs)", "Closing (cs)", "Days to Clear",
                   "Status", "Notes"]

def write_tracker_sheet(wb, lines, latest):
    ws = wb.create_sheet("STN TRACKER")
    ws.sheet_view.showGridLines = False
    table_header(ws, 1, TRACKER_HEADERS)
    widths = [10, 26, 15, 19, 9, 10, 10, 7, 9, 9, 10, 11, 10, 8, 10, 10, 9, 16, 40]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    r = 2
    for i, l in enumerate(sort_lines(lines)):
        zebra = ZEBRA if i % 2 else None
        code_v = int(l["code"]) if str(l["code"]).isdigit() else l["code"]
        vals = [code_v, l["shop"], l["bond"], l["brand"], l["pack"], l["stn_date"],
                l["arrival"], l["lag"], l["days"], l["applied"], l["received"], l["pre"],
                l["sold"], l["pct"], l["unsold"] if l["arrival"] else None,
                None if l["tier"] == T_TRANSIT else l["closing"], l["d2c"]]
        fmts = ["0", None, None, None, None, "dd mmm", "dd mmm", "0", "0", "#,##0.00", "#,##0.00",
                "#,##0.00", "#,##0.00", "0.0%", "#,##0.00", "#,##0.00", "0"]
        for j, (v, fm) in enumerate(zip(vals, fmts), 1):
            align = "left" if j in (2, 4) else "center"
            _c(ws, r, j, v, _f(10), zebra, align, fm)
        if l["waste_risk"]:
            _c(ws, r, 15, l["unsold"], _f(10, True, RISK_FONT), RISK_FILL, "center", "#,##0.00")
        if l["d2c"] is None and l["tier"] not in (T_TRANSIT, T_CLEAR) and (l["sold"] or 0) <= 0.005 and l["arrival"]:
            _c(ws, r, 17, "NO SALES", _f(9, True, RISK_FONT), RISK_FILL, "center")
        status_cell(ws, r, 18, l["tier"])
        nt = "; ".join(l["notes"])
        _c(ws, r, 19, nt, _f(9, False, "FF6B7280"), zebra, "left", wrap=True)
        ws.row_dimensions[r].height = max(20, 12 * (1 + len(nt) // 58))
        r += 1
    ws.freeze_panes = "A2"
    if r > 2:
        ws.auto_filter.ref = f"A1:S{r-1}"
        ws.conditional_formatting.add(
            f"N2:N{r-1}",
            DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                        color="64B5F6", showValue=True))
        ws.conditional_formatting.add(
            f"O2:O{r-1}",
            DataBarRule(start_type="num", start_value=0, end_type="max",
                        color="FFB300", showValue=True))
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws

def _paint_canvas(ws, ncols):
    """Two-pass dark canvas (11 Jul 2026): fill every still-unstyled cell AFTER
    content is laid out, sized to real content — the old fixed range(1,150) left
    the dashboard bottom white once the in-transit book grew past ~150 rows."""
    for rr in range(1, ws.max_row + 3):
        for cc in range(1, ncols + 1):
            cell = ws.cell(row=rr, column=cc)
            if cell.fill is None or cell.fill.fill_type is None:
                cell.fill = _fill(DK_CANVAS)

def write_dashboard(wb, lines, latest, months_read, base=None):
    """LIGHT two-panel dashboard (v4, 17 Jul 2026 — Abhay: 'i dont like the dark
    background style'). White canvas; navy hero + gold rule kept for identity;
    light KPI cards; navy table headers; pastel tier chips. LEFT A:H = lifecycle
    → bonds → in-transit watch · RIGHT J:Q = brand → pack → top unsold."""
    ws = wb.create_sheet("DASHBOARD", 0)
    NC = 17
    ws.sheet_view.showGridLines = False
    a = agg_block(lines)
    avg_lag = (sum(a["lags"]) / len(a["lags"])) if a["lags"] else None
    avg_wait = (sum(a["waits"]) / len(a["waits"])) if a.get("waits") else None
    behind = (date.today() - latest).days if latest else None
    hero(ws, NC, "DASHBOARD",
         f"Applied → Received → Liquidated · every STN pays 14% of case rate — unsold STN is wasted spend"
         f" · data through {latest.strftime('%d %b %Y') if latest else '—'}"
         + (f" · sources: {', '.join(months_read)}" if months_read else "")
         + (f" · ⚠ data {behind}d behind today" if behind is not None and behind > 1 else ""))
    gold_rule(ws, 3, NC)
    # KSD logo removed from the dashboard header (Abhay, 17 Jul 2026) — _logo_png retained unused
    for j, w in enumerate([15, 11, 10.5, 10.5, 9.5, 9.5, 12.5, 12, 2,
                           15, 11, 10.5, 10.5, 9.5, 9.5, 12.5, 12], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.print_options.horizontalCentered = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    if not lines:
        ws.merge_cells("A5:N5")
        _c(ws, 5, 1, "No STN lines yet — drop STN REQUEST workbooks in  STN tracking/requests/  "
                     "(or screenshots in screenshots/) and say “update STN tracker”.",
           _f(12, True, NAVY_MID), "FFFFF8E1", "left", border=None)
        ws.row_dimensions[5].height = 30
        _paint_wash(ws, NC)
        return ws
    light_tiles(ws, 4, [
        ("① APPLIED (cs)", round(a["applied"], 2), f"{a['n']} lines · the ask", None, None, "#,##0"),
        ("② RECEIVED (cs)", round(a["received"], 2),
         f"→ {a['received']/a['applied']:.0%} of applied · {a['arrived']} arrived" if a["applied"] else "—",
         None, None, "#,##0.0"),
        ("③ SOLD SINCE ARRIVAL", round(a["sold"], 2),
         f"→ {min(a['pct'], 1.0):.0%} of received" if a["pct"] is not None else "—", None, None, "#,##0.0"),
        ("UNSOLD STN (cs)", round(a["unsold"], 2), "levy paid · not sold", "FFE65100", "FFE65100", "#,##0.0"),
        # 31 Jul 2026: this tile is about lines still WAITING, so it shows their
        # average wait. It used to print the receipt lag of the lines that had
        # already ARRIVED (4.8d vs the true 8.1d), understating its own urgency.
        ("IN TRANSIT (cs)", round(a["transit_cs"], 2),
         f"{a['transit_n']} lines" + (f" · avg wait {avg_wait:.0f}d" if avg_wait is not None else "")
         + (f" · {a['lapsed_n']} lapsed ({a['lapsed_cs']:g} cs)" if a.get("lapsed_n") else ""),
         STEEL, NAVY_SOFT, "#,##0"),
        ("WASTE RISK (cs)", round(a["risk_cs"], 2),
         f"{a['risk_n']} lines ≥{WASTE_RISK_DAYS}d & <30% sold",
         RED_BRIGHT if a["risk_cs"] else "FF43A047",
         "FFC62828" if a["risk_cs"] else "FF2E7D32", "#,##0.0"),
    ], spans=[(1, 3), (4, 6), (7, 8), (10, 12), (13, 15), (16, 17)])
    # ---- SYNCHRONISED BANDS (17 Jul 2026, "structure all the tables in a better way"):
    # each band's two tables START on the same row, so titles + headers align
    # across the page. Band 1: LIFECYCLE | PACK · Band 2: BOND | BRAND ·
    # Band 3: IN-TRANSIT WATCH | TOP UNSOLD. Legend REMOVED (Abhay).
    r0 = 8
    r_l = lifecycle_panel(ws, r0, 1, lines)
    r_r = group_panel(ws, r0, 10, "BY PACK", "Pack", lines, lambda l: l["pack"])
    r1 = max(r_l, r_r)
    r_l = group_panel(ws, r1, 1, "BY BOND", "Bond", lines, lambda l: l["bond"])
    r_r = group_panel(ws, r1, 10, "BY BRAND", "Brand", lines, lambda l: l["brand"])
    r2 = max(r_l, r_r)
    transit = sorted([l for l in lines if l["tier"] == T_TRANSIT],
                     key=lambda l: (-(l["wait"] or 0), -l["applied"]))
    worst = [l for l in lines if l["unsold"] > 0.01]
    worst.sort(key=lambda l: -l["unsold"])
    if transit:
        watch_panel(ws, r2, 1, transit)
    if worst:
        unsold_panel(ws, r2, 10, worst)
    _paint_wash(ws, NC)
    return ws

BOND_HEADERS = ["Brand", "Pack", "STN Date", "Arrival", "Lag (d)", "Shelf (d)", "Applied",
                "Received", "Pre-Exist", "Sold", "% Sold", "Unsold STN", "Closing", "Rate cs/d",
                "Status"]  # Notes + Days-to-Clear REMOVED from bond sheets (Abhay, 17 Jul 2026) — both stay on STN TRACKER
# even tile strip over the bond grid (col widths 19,9,10,10,7,8,9,10,10,9,8,10,9,9,16):
# spans chosen so the six CARDS are near-EQUAL visual widths (19/20/17/19/19/16 units) with
# even gaps (B,E,H,K,N) — the old (A:B)(D:E)... layout made APPLIED 28 units vs IN TRANSIT 9
# ("beautify this part", Abhay 17 Jul 2026). Col O = last card flush with the table edge.
BOND_TILE_SPANS = [(1, 1), (3, 4), (6, 7), (9, 10), (12, 13), (15, 15)]

def bond_tiles(a):
    """The 6 bond KPI tiles (numeric values — no Excel 'number as text' triangles)."""
    return [
        ("① APPLIED", round(a["applied"], 2), f"{a['n']} lines", None, None, "#,##0"),
        ("② RECEIVED", round(a["received"], 2),
         f"→ {a['received']/a['applied']:.0%} of applied" if a["applied"] else "—", None, None, "#,##0.0"),
        ("③ SOLD", round(a["sold"], 2),
         f"→ {min(a['pct'], 1.0):.0%} of received" if a["pct"] is not None else "—", None, None, "#,##0.0"),
        ("UNSOLD STN", round(a["unsold"], 2), "cs on shelf", GOLD, GOLD_DIM, "#,##0.0"),
        ("IN TRANSIT", round(a["transit_cs"], 2), f"{a['transit_n']} lines", STEEL, STEEL_LT, "#,##0"),
        ("WASTE RISK", round(a["risk_cs"], 2), f"{a['risk_n']} lines", RED_BRIGHT,
         RED_LT if a["risk_cs"] else "FFA5D6A7", "#,##0.0"),
    ]

def bond_staff(master, bond):
    c = Counter(rec["staff"] for rec in master.values()
                if rec.get("bond") == bond and rec.get("cat") == "KSBC"
                and rec.get("status", "").lower() == "active"
                and rec.get("staff") and rec.get("staff").upper() != "VACANT")
    return c.most_common(1)[0][0] if c else ""

def write_bond_sheets(wb, lines, latest, all_bonds, master, base=None):
    """ALL bonds get a sheet, always: bonds with STN lines first, ranked by
    unsold STN desc (wasted-levy exposure); the rest alphabetical with an
    empty-state notice — every bond is one tab away."""
    by_bond = defaultdict(list)
    for l in lines:
        by_bond[l["bond"]].append(l)
    active = sorted(by_bond, key=lambda b: (-agg_block(by_bond[b])["unsold"],
                                            -agg_block(by_bond[b])["transit_cs"]))
    rest = sorted(b for b in all_bonds if b not in by_bond)
    NC = len(BOND_HEADERS)
    for bond in active + rest:
        bl = by_bond.get(bond, [])
        ws = wb.create_sheet(str(bond)[:31])
        ws.sheet_view.showGridLines = False
        staff = bond_staff(master, bond)
        for j, w in enumerate([19, 9, 10, 10, 7, 8, 9, 10, 10, 9, 8, 10, 9, 9, 16], 1):
            ws.column_dimensions[get_column_letter(j)].width = w
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        if not bl:
            hero(ws, NC, bond, "  ·  ".join(s for s in (
                staff, "no STN lines yet",
                f"data through {latest.strftime('%d %b %Y') if latest else '—'}") if s))
            gold_rule(ws, 3, NC)
            dark_band(ws, 4, 6, NC)
            kpi_tiles(ws, 4, bond_tiles(agg_block([])), spans=BOND_TILE_SPANS)
            ws.merge_cells(start_row=8, start_column=1, end_row=8, end_column=NC)
            _c(ws, 8, 1, "No active STN lines in this bond — new applications will appear here.",
               _f(11, True, NAVY_MID), "FFFFF8E1", "left", border=None, indent=1)
            ws.row_dimensions[8].height = 26
            ws.freeze_panes = "A4"
            continue
        a = agg_block(bl)
        hero(ws, NC, bond, "  ·  ".join(s for s in (
            staff, f"{a['n']} STN lines",
            f"data through {latest.strftime('%d %b %Y') if latest else '—'}") if s))
        gold_rule(ws, 3, NC)
        dark_band(ws, 4, 6, NC)
        kpi_tiles(ws, 4, bond_tiles(a), spans=BOND_TILE_SPANS)
        r = 8
        shops = defaultdict(list)
        for l in bl:
            shops[(l["code"], l["shop"])].append(l)
        shop_rank = sorted(shops.items(),
                           key=lambda kv: (-agg_block(kv[1])["unsold"], -agg_block(kv[1])["transit_cs"]))
        for (code, shop), sl in shop_rank:
            sa = agg_block(sl)  # still feeds the per-shop TOTAL row
            head = f"{code} · {shop}"  # shop CODE · NAME, middle-dot (Abhay 17 Jul 2026); badge + unsold suffix removed
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
            for cc in range(1, NC + 1):
                ws.cell(row=r, column=cc).fill = _fill(NAVY_MID)  # royal-navy banner (Abhay 17 Jul — slate FF334466 read muddy next to the gold-navy header)
            _c(ws, r, 1, head, _f(11, True, "FFFFFFFF"), NAVY_MID, "left",
               border=Border(left=Side(style="medium", color=GOLD), bottom=Side(style="thin", color="FFFFD700")),
               indent=1)
            ws.row_dimensions[r].height = 26
            r += 1
            table_header(ws, r, BOND_HEADERS)
            r += 1
            blk_first = r
            for i, l in enumerate(sort_lines(sl)):
                zebra = ZEBRA if i % 2 else None
                vals = [l["brand"], l["pack"], l["stn_date"], l["arrival"], l["lag"], l["days"],
                        l["applied"], l["received"], l["pre"], l["sold"], l["pct"],
                        l["unsold"] if l["arrival"] else None,
                        None if l["tier"] == T_TRANSIT else l["closing"], l["rate"]]
                fmts = [None, None, "dd mmm", "dd mmm", "0", "0", "#,##0.00", "#,##0.00", "#,##0.00",
                        "#,##0.00", "0.0%", "#,##0.00", "#,##0.00", "0.00"]
                for j, (v, fm) in enumerate(zip(vals, fmts), 1):
                    _c(ws, r, j, v, _f(10), zebra, "left" if j == 1 else "center", fm)
                if l["waste_risk"]:
                    _c(ws, r, 12, l["unsold"], _f(10, True, RISK_FONT), RISK_FILL, "center", "#,##0.00")
                if l["d2c"] is None and l["tier"] not in (T_TRANSIT, T_CLEAR) and (l["sold"] or 0) <= 0.005 and l["arrival"]:
                    _c(ws, r, 14, "NO SALES", _f(9, True, RISK_FONT), RISK_FILL, "center")  # over the 0.00 rate — badge kept after d2c col drop
                status_cell(ws, r, 15, l["tier"])
                ws.row_dimensions[r].height = 20
                r += 1
            if r > blk_first:
                ws.conditional_formatting.add(
                    f"K{blk_first}:K{r-1}",
                    DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                color="64B5F6", showValue=True))
                ws.conditional_formatting.add(
                    f"L{blk_first}:L{r-1}",
                    DataBarRule(start_type="num", start_value=0, end_type="max",
                                color="FFB300", showValue=True))
            agg = [sum(l["applied"] for l in sl), sum(l["received"] for l in sl),
                   sum((l["pre"] or 0) for l in sl), sum(l["sold"] for l in sl)]
            pct = agg[3] / agg[1] if agg[1] > 0 else None
            for cc in range(1, NC + 1):
                ws.cell(row=r, column=cc).fill = _fill(GREY_TOTAL)
            _c(ws, r, 1, "TOTAL", _f(10, True, "FFFFFFFF"), GREY_TOTAL, "left", indent=1)
            for j, v in zip((7, 8, 9, 10), agg):
                _c(ws, r, j, round(v, 2), _f(10, True, "FFFFFFFF"), GREY_TOTAL, "center", "#,##0.00")
            _c(ws, r, 11, (min(pct, 1.0) if pct is not None else None), _f(10, True, GOLD_DIM), GREY_TOTAL, "center", "0.0%")
            _c(ws, r, 12, round(sa["unsold"], 2), _f(10, True, GOLD_DIM), GREY_TOTAL, "center", "#,##0.00")
            # 31 Jul 2026: exclude transit lines — their Closing cell is blanked in
            # the rows above (it is pre-STN shelf stock), so summing it made the
            # TOTAL disagree with its own column on 19 blocks (+144.58 cs).
            _c(ws, r, 13, round(sum((l["closing"] or 0) for l in sl if l["tier"] != T_TRANSIT), 2), _f(10, True, "FFFFFFFF"),
               GREY_TOTAL, "center", "#,##0.00")
            ws.row_dimensions[r].height = 22
            r += 2
        ws.freeze_panes = "A4"

def write_log(wb, entries, master):
    ws = wb.create_sheet("STN LOG")
    ws.sheet_view.showGridLines = False
    table_header(ws, 1, LOG_HEADERS)
    for j, w in enumerate([11, 11, 10, 26, 14, 19, 9, 11, 22, 34], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    r = 2
    for e in sorted(entries, key=lambda x: (x["stn_date"], x["code"], x["brand"])):
        rec = master.get(e["code"], {})
        code_v = int(e["code"]) if str(e["code"]).isdigit() else e["code"]
        vals = [e.get("entry_date"), e["stn_date"], code_v, rec.get("name", ""), rec.get("bond", ""),
                e["brand"], e["pack"], e["cases"], e["source"], e.get("notes", "")]
        fmts = ["dd mmm yyyy", "dd mmm yyyy", "0", None, None, None, None, "#,##0.00", None, None]
        for j, (v, fm) in enumerate(zip(vals, fmts), 1):
            _c(ws, r, j, v, _f(10), ZEBRA if r % 2 else None, "left" if j in (4, 9, 10) else "center", fm)
        r += 1
    if r > 2:
        ws.auto_filter.ref = f"A1:J{r-1}"
    ws.freeze_panes = "A2"

README_ROWS = [
    ("h1", "STN CONTROL TOWER — HOW TO READ THIS WORKBOOK"),
    ("p", "Built by the STN tracker pipeline (v3, 10 Jul 2026). Static values — every number was "
          "computed at build time from the KSBC daily raw sheets; nothing recalculates."),
    ("h2", "WHAT AN STN IS"),
    ("p", "STN = Stock Transfer Note under KSBC's STN Scheme: a supplier-initiated, permit-based "
          "transfer of our own IMFL brands from an FL-9 warehouse to specific FL-1 shops. KSBC "
          "grants these on request (premium brands can be restricted). Normal shop replenishment "
          "is KSBC-driven — each shop's Purchase Instruction (RL/RQ/MQ) is recomputed monthly from "
          "trailing 3-month offtake — so STN is our only lever to place stock a shop's ROQ won't pull."),
    ("p", "THE COST: KSD pays 14% of case rate per STN case as KSBC's transfer levy. The levy is "
          "sunk on transfer. If the stock reaches the shop and does not sell, we lose three ways: "
          "the levy is wasted, the case value sits locked in shop inventory (FL-1 sales settle "
          "fortnightly), and the shop's PI ceiling gets no lift (MQ rises only from offtake share). "
          "That is why UNSOLD STN cases are the workbook's headline exposure — tracked in cases, "
          "not rupees, by design."),
    ("h2", "LIFECYCLE OF EVERY LINE"),
    ("p", "Each line = one shop × brand × pack × cases × STN date. Stages:"),
    ("chip", (T_TRANSIT, "Applied, no receipt at the shop yet. Days Waiting counts from the STN "
              "date to the LATEST DATA DAY; at 7+ days the line is flagged ⚠ CHASE "
              "(permit/warehouse follow-up).")),
    ("chip", ("ARRIVAL", "First day the shop's daily raw shows Shop In > 0 on/after the STN date, "
              "within the line's window. Repeat STNs on the same shop×brand×pack split the "
              "timeline ([STN i, STN i+1)) so one receipt is never counted into two lines. "
              "Lag (d) = STN date → arrival.")),
    ("chip", (T_STUCK, "<10% of received sold since arrival — the alarm tier.")),
    ("chip", (T_SLOW, "10–30% sold since arrival.")),
    ("chip", (T_MOVING, "30–60% sold since arrival.")),
    ("chip", (T_FAST, "≥60% sold since arrival.")),
    ("chip", (T_CLEAR, "Applied quantity landed and closing ≤ 0.01 cs — the STN did its job. TERMINAL: the line freezes on "
              "its clear date; routine stock arriving later never reopens it.")),
    ("chip", ("🗑 WASTE RISK", "Arrived ≥30 days ago, still <30% sold, stock still on shelf: the "
              "levy-paid-for-nothing set. Tinted red wherever it appears.")),
    ("li", "🏁 CLEARED is terminal, and a line only clears once the quantity applied for has "
           "actually landed AND the shelf has gone to zero. A token receipt arriving ahead of the "
           "consignment can no longer close the line (31 Jul 2026 — that fault had buried 281 cs "
           "of deliveries across 22 lines). A genuinely short delivery still clears on the shelf "
           "going to zero, and is noted as a partial."),
    ("h2", "COLUMNS THAT NEED CARE"),
    ("li", "Pre-Existing (cs) — stock already on the shelf on arrival day. Never netted out of the "
           "STN metrics; old stock can absorb sales, so it stays visible."),
    ("li", "% Sold = Sold Since Arrival ÷ Received (raw receipts), capped at 100%. Both sides of "
           "the ratio count the same pool: sales come out of every case on the shelf, so the "
           "denominator counts every case received. (31 Jul 2026 — the old basis divided by the "
           "capped figure, which put 42 lines at a phantom 100%.)"),
    ("li", "Unsold STN (cs) = min(Received, Applied) × (1 − % Sold) — our own cases multiplied by "
           "the share of the pool that did not move. Zero once a line clears. Capped at Applied: a "
           "routine indent in the window can never inflate the exposure past the levy actually "
           "paid. (31 Jul 2026 — subtracting ALL sales from our cases understated exposure by "
           "159 cs across 112 lines.)"),
    ("li", "Received (raw receipts) can exceed Applied when a regular indent lands in the same "
           "window — noted on the line, and the exposure metrics cap at Applied. Partial "
           "receipts are noted too."),
    ("li", "Days to Clear = Closing ÷ daily rate since arrival; NO SALES badge when rate ≈ 0."),
    ("li", "Every line is reconciled Pre + Received − Sold = Closing (tolerance 0.15 cs for "
           "case-conversion rounding). Drift is flagged in Notes."),
    ("h2", "WORKBOOK MAP"),
    ("li", "DASHBOARD — network KPIs, lifecycle stages, BY BOND / BRAND / PACK, in-transit watch, "
           "top unsold positions."),
    ("li", "STN TRACKER — every line, worst-first (stage, then unsold exposure)."),
    ("li", "15 bond sheets — always present; bonds with lines first (ranked by unsold STN), then "
           "empty bonds. Shop banner blocks, per-shop TOTALs."),
    ("li", "STN LOG — the permanent register of every application (the accumulator every rebuild "
           "seeds from). Entry vs STN date, source (who sent it), notes."),
    ("h2", "WORKFLOW"),
    ("li", "New STNs: drop filled STN REQUEST workbooks in STN tracking/requests/ (or WhatsApp "
           "screenshots in screenshots/) and say “update STN tracker”."),
    ("li", "Refresh only (newer sales days, no new STNs): say “refresh STN tracker”."),
    ("li", "Dedupe: (STN date, shop, brand, pack) — re-submitting a line replaces it, never "
           "doubles. A near-duplicate (same shop/brand/pack/cases within 3 days, different date) "
           "aborts the build as a suspected double-entry; split rows in one request file are "
           "summed automatically."),
    ("li", "Field shorthand: CCB = CHAIRMAN'S CHOICE (never BCB) · OP/OPR = OLD PEARL · "
           "ROF = ROYAL OLD FORT · MWB = MORNING WALKER · KS/KS99 = KS 99."),
    ("p", "Kerala's 1st of month is a dry day (zero movement) — a line arriving on the 1st shows "
          "no sales that day by definition. Sources: KSBC monthly analysis workbooks (daily raw "
          "sheets), MASTER DATA CONFIRMED.xlsx. STN Scheme background: BEVCO homepage notice "
          "(Dec 2022), KSBC Rules & Regulations (Purchase), KSBC STN COMPLETE registers."),
]

def write_readme(wb):
    ws = wb.create_sheet("README")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 98
    r = 1
    for kind, text in README_ROWS:
        if kind == "h1":
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            _c(ws, r, 2, text, _f(15, True, "FFFFFFFF"), NAVY_DEEP, "left", border=None, indent=1)
            ws.row_dimensions[r].height = 30
            r += 1
            for cc in (2, 3):
                ws.cell(row=r, column=cc).fill = _fill(GOLD)
            ws.row_dimensions[r].height = 4
        elif kind == "h2":
            r += 1
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            _c(ws, r, 2, text, _f(12, True, NAVY_MID), None, "left", border=GOLD_BOT)
            ws.row_dimensions[r].height = 22
        elif kind == "chip":
            chip, desc = text
            if chip in TIER_COLOR:
                dark, light = TIER_COLOR[chip]
            elif "WASTE" in chip:
                dark, light = RISK_FONT, RISK_FILL
            else:
                dark, light = NAVY_MID, "FFE8EDF7"
            _c(ws, r, 2, chip, _f(10, True, dark), light, "center")
            _c(ws, r, 3, desc, _f(10), None, "left", border=None, wrap=True)
            ws.row_dimensions[r].height = max(18, 13 * (1 + len(desc) // 92))
        elif kind == "li":
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            _c(ws, r, 2, "•  " + text, _f(10), None, "left", border=None, wrap=True)
            ws.row_dimensions[r].height = max(15, 13 * (1 + len(text) // 108))
        else:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            _c(ws, r, 2, text, _f(10), None, "left", border=None, wrap=True)
            ws.row_dimensions[r].height = max(15, 13 * (1 + len(text) // 108))
        r += 1
    return ws

# ----------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", default=str(BASE_DEFAULT))
    ap.add_argument("--add-json", default=None)
    ap.add_argument("--add-xlsx", action="append", default=None,
                    help="STN REQUEST workbook (repeatable); the standard input since 6 Jul 2026")
    ap.add_argument("--register", action="append", default=None,
                    help="Abhay's STN REGISTER workbook (17 Jul 2026 format) — month tabs "
                         "wholesale-reconcile those months of the log")
    ap.add_argument("--out", default=None, help="output path (default: scratch next to live)")
    ap.add_argument("--year", type=int, default=None,
                    help="DEPRECATED — ignored; workbook years are auto-assigned (cross-year safe)")
    ap.add_argument("--as-of", default=None,
                    help="anchor date for year assignment (default: today)")
    ap.add_argument("--retire", action="append", default=None,
                    help='remove a logged line permanently: "SHOP|BRAND|PACK|YYYY-MM-DD" (repeatable)')
    ap.add_argument("--allow-neardup", action="store_true",
                    help="accept a suspected date-shifted duplicate batch as genuinely two batches")
    ap.add_argument("--allow-gaps", action="store_true",
                    help="build despite a missing KSBC month in the movement window (WARN instead of abort)")
    ap.add_argument("--no-total-check", action="store_true",
                    help="allow a request file whose TOTAL row/col cannot be verified")
    args = ap.parse_args()
    base = Path(args.base)
    live = base / "STN tracking" / "STN TRACKER.xlsx"
    out = Path(args.out) if args.out else live.with_name("STN TRACKER (SCRATCH).xlsx")
    anchor = parse_dt(args.as_of) if args.as_of else date.today()
    if args.year is not None:
        print(f"  ~ --year is deprecated and ignored (years auto-assigned around {anchor})")

    master = load_master(base)
    seeded, seed_problems = seed_from_live(live)
    for p in seed_problems:
        print(f"  ~ {p}")
    new = load_new_entries(args.add_json, master) if args.add_json else []
    req_errors, req_warns = [], []
    reg_months = set()
    for p in (args.register or []):
        e, errs, wrns, mset = parse_register_xlsx(p, master)
        new += e
        req_errors += errs
        req_warns += wrns
        reg_months |= mset
        if not errs:
            print(f"register {Path(p).name}: {len(e)} submitted lines, {sum(x['cases'] for x in e):g} cs "
                  f"across {', '.join(MONTH_BY_NUM[m] + ' ' + str(y) for y, m in sorted(mset)) or '—'}")
    for p in (args.add_xlsx or []):
        e, errs, wrns = parse_request_xlsx(p, master, strict_totals=not args.no_total_check)
        new += e
        req_errors += errs
        req_warns += wrns
        if not errs:
            print(f"request {Path(p).name}: {len(e)} lines, {sum(x['cases'] for x in e):g} cs"
                  + (f" (STN date {e[0]['stn_date']})" if e else ""))
    # 31 Jul 2026: print warnings BEFORE the abort. They used to be discarded on
    # a failed build, so an operator fixing one typo never saw the other problems
    # until the next run.
    for w in req_warns:
        print(f"  ~ {w}")
    if req_errors:
        sys.exit("ABORT — request file validation failed, nothing written:\n  " + "\n  ".join(req_errors))
    if reg_months:
        # REGISTER RECONCILE (17 Jul 2026): the register file is the source of
        # truth for its months — drop those months' seeded lines wholesale so
        # Abhay's in-place edits (redates, corrections, pending→submitted) flow.
        dropped = [e for e in seeded if (e["stn_date"].year, e["stn_date"].month) in reg_months]
        seeded = [e for e in seeded if (e["stn_date"].year, e["stn_date"].month) not in reg_months]
        print(f"register reconcile: {len(dropped)} logged line(s) in "
              f"{', '.join(MONTH_BY_NUM[m] + ' ' + str(y) for y, m in sorted(reg_months))} "
              f"replaced by the register file")
    entries, replaced, intra = merge_entries(seeded, new)
    if intra:
        sys.exit("ABORT — the same (STN date, shop, brand, pack) arrived from TWO inputs of this run "
                 "(split rows within one file are summed automatically; the same line in two files is "
                 "ambiguous — ingest them separately or fix the duplicates):\n  " + "\n  ".join(intra))
    print(f"entries: {len(seeded)} seeded from live log, {len(new)} new, {replaced} replaced -> {len(entries)} total")

    # --retire: permanently remove wrong log lines (11 Jul 2026 — the supported
    # correction path for date-shifted double entries)
    retired = []
    for spec in (args.retire or []):
        try:
            s_shop, s_brand, s_pack, s_date = [x.strip() for x in spec.split("|")]
        except ValueError:
            sys.exit(f'ABORT — bad --retire spec {spec!r} (want "SHOP|BRAND|PACK|YYYY-MM-DD")')
        r_code, _cands = resolve_shop(s_shop, master)
        r_brand = canon_brand(s_brand) or s_brand
        r_pack = canon_pack(s_pack) or s_pack
        r_date = parse_dt(s_date)
        keep, hit = [], False
        for e in entries:
            if (e["code"] == (r_code or s_shop) and _norm(e["brand"]) == _norm(r_brand)
                    and e["pack"] == r_pack and e["stn_date"] == r_date):
                retired.append(f"{e['code']} {e['brand']} {e['pack']} {e['cases']:g} cs @ {e['stn_date']}")
                hit = True
            else:
                keep.append(e)
        if not hit:
            sys.exit(f"ABORT — --retire matched nothing: {spec!r} (check spelling/date; nothing written)")
        entries = keep
    for x in retired:
        print(f"  RETIRED from log: {x}")

    # near-dup tripwire — only pairs involving THIS run's new entries
    new_keys = {(e["stn_date"], e["code"], _norm(e["brand"]), e["pack"]) for e in new}
    sus = find_near_dups(entries, new_keys)
    if sus and args.register:
        for s in sus:
            print(f"  ~ near-dup in register (file is authoritative — verify in the sheet): {s}")
        sus = []
    if sus and not args.allow_neardup:
        sys.exit("ABORT — suspected date-shifted duplicate batch (same shop×brand×pack×cases within "
                 f"{NEARDUP_DAYS}d, different STN dates). Confirm with the sender; re-run with "
                 "--allow-neardup if BOTH are real, or --retire the wrong one:\n  " + "\n  ".join(sus))
    for s in sus:
        print(f"  ~ near-dup allowed (--allow-neardup): {s}")

    lines, latest, unmapped, mr_names = [], None, set(), []
    if entries:
        min_stn = min(e["stn_date"] for e in entries)
        from_key = (min_stn.year, min_stn.month)
        shop_codes = {e["code"] for e in entries}
        mov, latest, unmapped, months_read = read_movement(base, from_key, anchor, shop_codes)
        assert_month_coverage(from_key, latest, months_read, args.allow_gaps)
        mr_names = _mr_names(months_read)
        # SAME-SKU WINDOW PARTITION: line i's window ends where line i+1 begins
        by_sku = defaultdict(list)
        for e in entries:
            by_sku[(e["code"], e["brand"], e["pack"])].append(e)
        wend = {}
        for group in by_sku.values():
            group.sort(key=lambda x: x["stn_date"])
            for gi, e in enumerate(group):
                wend[id(e)] = group[gi + 1]["stn_date"] if gi + 1 < len(group) else None
        lines = [compute_line(e, mov, latest, master, window_end=wend[id(e)]) for e in entries]
        bad_bond = sorted({l["code"] for l in lines if l["bond"] == "?"})
        if bad_bond:
            sys.exit("ABORT — logged shop code(s) no longer resolve in MASTER DATA (bond unknown): "
                     + ", ".join(bad_bond) + " — fix master or --retire the line(s); nothing written")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    write_tracker_sheet(wb, lines, latest)
    write_dashboard(wb, lines, latest, mr_names, base=base)
    all_bonds = sorted({rec["bond"] for rec in master.values()
                        if rec.get("cat") == "KSBC" and rec.get("status", "").lower() == "active"
                        and rec.get("bond")})
    write_bond_sheets(wb, lines, latest, all_bonds, master, base=base)
    write_log(wb, entries, master)
    write_readme(wb)
    out.parent.mkdir(parents=True, exist_ok=True)
    # 31 Jul 2026: write to a sibling temp file then rename. openpyxl truncates
    # the target the moment it opens the zip, so a crash mid-save would leave an
    # invalid file — and STN LOG is the only accumulator this stream has.
    _tmp = Path(out).with_name(Path(out).name + ".tmp")
    wb.save(_tmp)
    os.replace(_tmp, out)

    # ---- summary
    a = agg_block(lines)
    print(f"data through: {latest} · sources: {', '.join(mr_names) or '—'}")
    if unmapped:
        print(f"NOTE unmapped workbook brands (tracked verbatim): {sorted(unmapped)[:6]}")
    counts = defaultdict(int)
    for l in lines:
        counts[l["tier"]] += 1
    for t in TIER_ORDER:
        if counts[t]:
            print(f"  {t}: {counts[t]}")
    if a["applied"]:
        pct_s = f"{min(a['pct'], 1.0):.0%}" if a["pct"] is not None else "—"
        print(f"  applied {a['applied']:g} cs · received {a['received']:g} ({(a['received']/a['applied']):.0%})"
              f" · sold {a['sold']:g} ({pct_s} of received) · unsold {a['unsold']:g}"
              f" · transit {a['transit_cs']:g} · waste-risk {a['risk_cs']:g}")
    if retired:
        print(f"  retired {len(retired)} log line(s) this run — permanently out of the STN LOG")
    for l in lines:
        for n in l["notes"]:
            print(f"  ! {l['shop']} {l['brand']} {l['pack']}: {n}")
    print(f"WROTE: {out}")
    return 0

if __name__ == "__main__":
    sys.exit(main())
