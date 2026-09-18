#!/usr/bin/env python3
"""
ksbc_receipt_notes.py — bake receipt-date notes into the KSBC DETAIL sheets.

Adds an Excel note (legacy comment, author "KSD") on every Receipts cell
(column C) of the 15 `<BOND> DETAIL` sheets, listing WHICH DATE(S) the
receipts came in, exactly in the MEETING ANALYSIS-JUNE format:

    RECEIPTS — OLD PEARL NO.1 MATURED XXX RUM (all packs)
    09 Jul  ·  81 cs
    11 Jul  ·  17 cs
    Σ 98 cs · 2 days

Scope (locked 16 Jul 2026, Abhay: "DETAIL sheets only"):
  * brand summary rows  → "RECEIPTS — <BRAND> (all packs)"
  * pack rows           → "RECEIPTS — <BRAND> · <PACK>"
  * shop TOTAL rows     → "RECEIPTS — ALL BRANDS (shop)"
  * only cells whose period receipts > 0 get a note (precedent);
    the Σ footer appears only when there are 2+ receipt-day lines.

Data source: the daily raw audit sheets inside the SAME workbook
(`JULY 1` … `JULY 15`, range sheets like `JULY 3-5` label as `03–05 Jul`).
CUMULATIVE / COMBINED sheets are never scanned (period totals, no dates).
Receipts per line = Shop In Cases + Shop In Bottles / BPC (loose bottles
folded — same rule as every other KSBC consumer).

Matching mirrors ksbc_detail_nested exactly: shop codes canonicalised via
ksbc_alias_guard.canonical_code (MUKKAM 111014→11014), brand names via
_norm_brand (apostrophe/whitespace drift). A note is therefore always
consistent with how the cell value itself was aggregated.

Idempotent + self-healing: EVERY existing comment on the DETAIL sheets is
wiped first, then all notes are rebuilt from the daily sheets. Must run as
the LAST write step of the ksbc-shop-sales pipeline (after sorting, spacer
insertion and CF rebuilds — openpyxl row moves do not carry comments, so
anything earlier would misplace them; the end-of-pipeline rebuild heals any
stale anchor from the previous build).

Usage:  python3 ksbc_receipt_notes.py "<workbook.xlsx>"
Exit 0 on success; exit 1 on structural failure (no DETAIL/daily sheets).
Prints a per-class note count + any Σ-vs-cell drift > 0.15 cs (warn only —
with a CUMULATIVE override in effect the cell is period-true while the
note remains the daily audit story; informational, never a build gate).
"""
from __future__ import annotations
import os, re, sys, time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from openpyxl import load_workbook
from openpyxl.comments import Comment

from ksbc_detail_nested import _norm_brand, _is_pack_label, _find_shop_blocks, _parse_shop_code
from ksbc_alias_guard import canonical_code

AUTHOR = "KSD"
DRIFT_TOL = 0.15          # cs — same tolerance family as the other guards
EPS = 0.005

MONTHS = {m: i + 1 for i, m in enumerate(
    ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY",
     "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"])}
MON3 = {m: m[:3].title() for m in MONTHS}          # JULY -> Jul

# `JULY 7` or `JULY 3-5` — NOT `JULY 1-15 COMBINED`, NOT `JULY 1-16 CUMULATIVE`
DAILY_RE = re.compile(
    r"^(" + "|".join(MONTHS) + r")\s+(\d{1,2})(?:-(\d{1,2}))?$")


def _num(v) -> float:
    if v is None:
        return 0.0
    try:
        return float(str(v).replace(",", "").strip() or 0)
    except ValueError:
        return 0.0


def _fcs(v: float) -> str:
    """141 -> '141', 3.9097 -> '3.91' (trailing zeros trimmed)."""
    r = round(v)
    if abs(v - r) < EPS:
        return str(int(r))
    return f"{v:.2f}".rstrip("0").rstrip(".")


def scan_daily_sheets(wb):
    """Return (bp, br, labels): bp[(canon, normbrand, pack)][key]=cs,
    br[(canon, normbrand)][key]=cs; labels[key]=display label; key sorts
    chronologically (start day, end day)."""
    bp, br, labels = {}, {}, {}
    n_sheets = 0
    for name in wb.sheetnames:
        m = DAILY_RE.match(name.strip())
        if not m:
            continue
        mon, d1, d2 = m.group(1), int(m.group(2)), m.group(3)
        d2 = int(d2) if d2 else None
        key = (MONTHS[mon], d1, d2 or d1)
        labels[key] = (f"{d1:02d}–{d2:02d} {MON3[mon]}" if d2
                       else f"{d1:02d} {MON3[mon]}")
        n_sheets += 1
        ws = wb[name]
        for r in ws.iter_rows(min_row=2, max_col=11, values_only=True):
            if not r or r[1] is None or r[4] is None:
                continue
            try:
                code = canonical_code(int(str(r[1]).strip()))
            except (ValueError, TypeError):
                continue
            bpc = _num(r[6])
            cs = _num(r[9]) + (_num(r[10]) / bpc if bpc else 0.0)
            if abs(cs) < 1e-9:
                continue
            nb = _norm_brand(r[4])
            pack = str(r[5]).strip().upper() if r[5] is not None else ""
            for dct, k in ((bp, (code, nb, pack)), (br, (code, nb))):
                slot = dct.setdefault(k, {})
                slot[key] = slot.get(key, 0.0) + cs
    return bp, br, labels, n_sheets


def _note_text(title: str, daymap: dict, labels: dict) -> str:
    lines = [title]
    keys = sorted(daymap)
    for k in keys:
        lines.append(f"{labels[k]}  ·  {_fcs(daymap[k])} cs")
    if len(keys) >= 2:
        tot = sum(daymap.values())
        lines.append(f"Σ {_fcs(tot)} cs · {len(keys)} days")
    return "\n".join(lines)


def _put(cell, title, daymap, labels, drift, sheet):
    tot = sum(daymap.values())
    if tot <= EPS:
        return 0
    txt = _note_text(title, daymap, labels)
    nlines = txt.count("\n") + 1
    cell.comment = Comment(txt, AUTHOR,
                           height=min(320, 30 + 14 * nlines), width=240)
    cv = cell.value
    if isinstance(cv, (int, float)) and abs(cv - tot) > DRIFT_TOL:
        drift.append((sheet, cell.coordinate, round(tot, 2), round(cv, 2)))
    return 1


def main(path: str) -> int:
    t0 = time.time()
    wb = load_workbook(path)
    detail = [n for n in wb.sheetnames if n.endswith(" DETAIL")]
    if not detail:
        print("FAIL: no DETAIL sheets found"); return 1
    bp, br, labels, n_daily = scan_daily_sheets(wb)
    if not n_daily:
        print("FAIL: no daily raw sheets found"); return 1

    wiped = n_brand = n_pack = n_total = 0
    drift = []
    for name in detail:
        ws = wb[name]
        for row in ws.iter_rows():
            for c in row:
                if c.comment is not None:
                    c.comment = None
                    wiped += 1
        for h, t in _find_shop_blocks(ws):
            raw_code = _parse_shop_code(str(ws.cell(h, 1).value))
            if raw_code is None:
                continue
            canon = canonical_code(raw_code)
            cur_nb, cur_brand = None, None
            shop_map = {}
            for r in range(h + 2, t):
                v = ws.cell(r, 1).value
                if not isinstance(v, str) or not v.strip():
                    continue
                if _is_pack_label(v):
                    if cur_nb is None:
                        continue
                    dm = bp.get((canon, cur_nb, v.strip().upper()))
                    if dm:
                        n_pack += _put(ws.cell(r, 3),
                                       f"RECEIPTS — {cur_brand} · {v.strip()}",
                                       dm, labels, drift, name)
                else:
                    cur_brand = v.strip()
                    cur_nb = _norm_brand(v)
                    dm = br.get((canon, cur_nb))
                    if dm:
                        n_brand += _put(ws.cell(r, 3),
                                        f"RECEIPTS — {cur_brand} (all packs)",
                                        dm, labels, drift, name)
                        for k, cs in dm.items():
                            shop_map[k] = shop_map.get(k, 0.0) + cs
            if shop_map:
                n_total += _put(ws.cell(t, 3), "RECEIPTS — ALL BRANDS (shop)",
                                shop_map, labels, drift, name)

    wb.save(path)
    print(f"RECEIPT NOTES OK · {len(detail)} DETAIL sheets · {n_daily} daily "
          f"sheets scanned · notes: {n_brand} brand + {n_pack} pack + "
          f"{n_total} shop-TOTAL = {n_brand + n_pack + n_total} "
          f"(wiped {wiped} old) · {time.time() - t0:.1f}s")
    if drift:
        print(f"WARN: {len(drift)} cell(s) drift >|{DRIFT_TOL}| cs vs note Σ "
              f"(cumulative override or case-rounding — informational):")
        for s, coord, ntot, cv in drift[:10]:
            print(f"  {s}!{coord}: note Σ {ntot} vs cell {cv}")
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("usage: python3 ksbc_receipt_notes.py <workbook.xlsx>")
        sys.exit(1)
    sys.exit(main(sys.argv[1]))
