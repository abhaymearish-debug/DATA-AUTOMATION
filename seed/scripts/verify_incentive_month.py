#!/usr/bin/env python3
"""Independently verify an incentive workbook against the primary sales sources.

This deliberately ignores the liquidation file when checking values and rebuilds
every figure from the KSBC shop-sales and Secondary-sales workbooks instead. If it
only re-read the liquidation file it would confirm the copy, not the numbers --
an error in the liquidation build would pass unnoticed.

    python3 verify_incentive_month.py --workbook "July.xlsx" --month JULY \
        --base "/path/to/Claude"
"""
import argparse, os, sys, re
from collections import defaultdict, Counter
import openpyxl
from openpyxl.utils import column_index_from_string as ci

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from incentive_common import (BRAND_COLS, norm_brand, bottles_per_case,
                              match_name, pick_sheet)

TOL = 0.005          # a cent of a case; anything larger is a real difference


def load_master(base):
    ws = openpyxl.load_workbook(os.path.join(base, 'MASTER DATA CONFIRMED.xlsx'),
                                data_only=True)['16-4-25']
    hdr = next(r for r in range(1, 10)
               if str(ws.cell(r, 4).value).strip() == 'Shop Code')
    m = {}
    for r in range(hdr + 1, ws.max_row + 1):
        code = ws.cell(r, 4).value
        if code in (None, ''):
            continue
        m[str(code).strip()] = dict(name=str(ws.cell(r, 5).value).strip(),
                                    cat=str(ws.cell(r, 6).value).strip().upper(),
                                    bond=str(ws.cell(r, 8).value).strip(),
                                    status=str(ws.cell(r, 9).value).strip())
    return m


def load_ksbc(base, month):
    """Per-shop, per-brand tertiary sales from the month's COMBINED roll-up.

    The COMBINED sheet is authoritative; summing the daily sheets drifts because
    KSBC's own exports round on case conversion.
    """
    path = os.path.join(base, 'KSBC shop sales', f'{month} SHOP SALES ANALYSIS.xlsx')
    if not os.path.exists(path):
        return None, f'not found: {path}'
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheet = next((s for s in wb.sheetnames
                  if s.upper().startswith(month) and s.upper().endswith('COMBINED')), None)
    if sheet is None:
        wb.close()
        return None, f'no COMBINED sheet in {os.path.basename(path)}'
    ws = wb[sheet]
    out = defaultdict(lambda: defaultdict(float))
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1] in (None, ''):
            continue
        brand = norm_brand(row[4])
        if brand and isinstance(row[9], (int, float)):
            out[str(int(row[1]))][brand] += float(row[9])
    wb.close()
    return out, sheet


def load_secondary(base, month):
    """FED and BAR invoice dispatches, with loose bottles folded into cases."""
    path = os.path.join(base, 'Secondary sales', f'{month} SECONDARY SALES ANALYSIS.xlsx')
    if not os.path.exists(path):
        return None, f'not found: {path}'
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheet = next((s for s in wb.sheetnames
                  if s.upper().startswith(month) and 'COMBINED DISPATCH' in s.upper()), None)
    if sheet is None:
        wb.close()
        return None, f'no COMBINED DISPATCHES sheet in {os.path.basename(path)}'
    ws = wb[sheet]
    out = defaultdict(lambda: defaultdict(float))
    stray = Counter()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[6] in (None, ''):
            continue
        brand = norm_brand(row[4])
        if not brand:
            continue
        cases = float(row[12]) if isinstance(row[12], (int, float)) else 0.0
        btl = float(row[13]) if isinstance(row[13], (int, float)) else 0.0
        per = bottles_per_case(row[5])
        if btl and not per:
            stray[str(row[5])] += btl          # a pack we cannot convert = silent loss
        out[str(row[6]).strip()][brand] += cases + (btl / per if (btl and per) else 0.0)
    wb.close()
    return out, (sheet, stray)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--workbook', required=True)
    ap.add_argument('--month', required=True)
    ap.add_argument('--base', required=True)
    ap.add_argument('--sheet', default=None)
    a = ap.parse_args()
    month = a.month.strip().upper()
    base = a.base

    liq_path = os.path.join(base, 'Monthly statement-liquidation',
                            f'TOTAL LIQUIDATION -{month}.xlsx')
    lq = openpyxl.load_workbook(liq_path, data_only=True)['Brand Sales by Shop']
    n2c = {str(lq.cell(r, 2).value).strip(): str(lq.cell(r, 1).value).strip()
           for r in range(2, lq.max_row + 1) if lq.cell(r, 2).value}

    wbv = openpyxl.load_workbook(a.workbook, data_only=True)
    ws = pick_sheet(wbv, a.sheet)

    master = load_master(base)
    ksbc, ks_note = load_ksbc(base, month)
    sec, sec_note = load_secondary(base, month)
    if ksbc is None or sec is None:
        sys.exit(f'ERROR: {ks_note if ksbc is None else sec_note}')
    sec_sheet, stray = sec_note

    rows = []
    for r in range(4, ws.max_row + 1):
        name = ws.cell(r, 2).value
        if name in (None, ''):
            continue
        name = str(name).strip()
        key = match_name(name, n2c)
        if key is None:
            print(f'ERROR: row {r} {name!r} is not in {os.path.basename(liq_path)}')
            sys.exit(1)
        challan = str(ws.cell(r, 3).value or '').strip().upper()
        leg = 'KSBC' if challan.startswith('KSBC') else ('FED' if challan == 'FED' else 'BAR')
        rows.append(dict(r=r, name=name, code=n2c[key], leg=leg,
                         bond=ws.cell(r, 6).value,
                         vals={b: float(ws.cell(r, ci(c)).value or 0)
                               for c, b in BRAND_COLS.items()}))

    print(f'workbook   : {os.path.basename(a.workbook)}  sheet {ws.title!r}')
    print(f'KSBC source: {ks_note}')
    print(f'sec. source: {sec_sheet}')
    if stray:
        print(f'  WARNING loose bottles with no known pack size: {dict(stray)}')
    print()

    fails = 0
    for leg, src in (('KSBC', ksbc), ('FED', sec), ('BAR', sec)):
        sub = [x for x in rows if x['leg'] == leg]
        bad, twb, tsrc = [], 0.0, 0.0
        for x in sub:
            s = src.get(x['code'], {})
            for brand in BRAND_COLS.values():
                got, exp = x['vals'][brand], s.get(brand, 0.0)
                twb += got
                tsrc += exp
                if abs(got - round(exp, 2)) > TOL:
                    bad.append((x['r'], x['name'], brand, got, round(exp, 4)))
        label = {'KSBC': 'KSBC (tertiary)', 'FED': 'CFD / Consumer Fed', 'BAR': 'BAR'}[leg]
        flag = 'PASS' if not bad else f'FAIL ({len(bad)})'
        print(f'{label:<22} {len(sub):>4} rows  {len(sub)*8:>5} cells   '
              f'workbook {twb:>10,.2f}   source {tsrc:>10,.2f}   {flag}')
        for b in bad[:15]:
            print(f'     row {b[0]} {b[1]!r} {b[2]}: workbook {b[3]} vs source {b[4]}')
        fails += len(bad)

    # Coverage: an outlet that sold but is not on the sheet is the costly failure,
    # because nothing downstream would ever reveal it.
    print()
    codes = {x['code'] for x in rows}
    for cat in ('KSBC', 'FED', 'BAR'):
        active = {c for c, m in master.items() if m['cat'] == cat and m['status'] == 'Active'}
        absent = active - codes
        src = ksbc if cat == 'KSBC' else sec
        sold = [(c, master[c]['name'], round(sum(src.get(c, {}).values()), 2))
                for c in absent if sum(src.get(c, {}).values()) > 0]
        note = 'PASS' if not sold else f'FAIL - {len(sold)} sold but missing: {sold[:5]}'
        print(f'{cat:<6} master active {len(active):>4} | on sheet {len(active & codes):>4} '
              f'| absent {len(absent):>3}  {note}')
        fails += len(sold)

    dup = [k for k, v in Counter(x['code'] for x in rows).items() if v > 1]
    cat_bad = [(x['r'], x['name'], x['leg'], master.get(x['code'], {}).get('cat'))
               for x in rows if master.get(x['code'], {}).get('cat') != x['leg']]
    bond_bad = [(x['r'], x['name'], x['bond'], master.get(x['code'], {}).get('bond'))
                for x in rows
                if str(x['bond'] or '').strip().upper()
                != str(master.get(x['code'], {}).get('bond', '')).strip().upper()]
    print()
    print(f'duplicate shop codes      : {len(dup)} {dup[:5] if dup else "PASS"}')
    print(f'type vs master category   : {len(cat_bad)} {cat_bad[:5] if cat_bad else "PASS"}')
    print(f'bond label vs master      : {len(bond_bad)} {bond_bad[:5] if bond_bad else "PASS"}')
    fails += len(dup) + len(cat_bad) + len(bond_bad)

    tot = sum(sum(x['vals'].values()) for x in rows)
    print()
    print(f'TOTAL {month}: {tot:,.2f} cases across {len(rows)} rows')
    print('VERIFY PASS' if fails == 0 else f'VERIFY FAIL - {fails} issue(s)')
    sys.exit(0 if fails == 0 else 2)


if __name__ == '__main__':
    main()
