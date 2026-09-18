#!/usr/bin/env python3
"""Roll the marketing-head incentive workbook forward to a new month.

Writes only the 8 brand-case columns, the two header date cells and the bar-row
block. Every formula, style, table, conditional format and merge is left exactly
as the previous month's file had it.

Run it in the background and poll the log -- LibreOffice recalculation takes
roughly a minute, which is longer than a foreground shell call usually allows:

    nohup python3 build_incentive_month.py --source "June.xlsx" \
        --liquidation ".../TOTAL LIQUIDATION -JULY.xlsx" \
        --out ".../July.xlsx" --month JULY --year 2026 > build.log 2>&1 &

The log ends with either "BUILD OK" or a line starting "ERROR".
"""
import argparse, datetime, os, sys, shutil
from collections import Counter
import openpyxl
from openpyxl.utils import column_index_from_string as ci

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from incentive_common import (BRAND_COLS, read_liquidation, match_name, pick_sheet,
                              recalculate, inject_cached_values)

MONTHS = {m: i for i, m in enumerate(
    ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST',
     'SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'], 1)}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--source', required=True, help='previous month incentive workbook')
    ap.add_argument('--liquidation', required=True, help='TOTAL LIQUIDATION -<MONTH>.xlsx')
    ap.add_argument('--out', required=True)
    ap.add_argument('--month', required=True)
    ap.add_argument('--year', type=int, required=True)
    ap.add_argument('--sheet', default=None, help='defaults to the first sheet carrying shop rows')
    a = ap.parse_args()

    month = a.month.strip().upper()
    if month not in MONTHS:
        sys.exit(f'ERROR: unrecognised month {a.month!r}')

    liq = read_liquidation(a.liquidation)
    print(f'liquidation source: {len(liq)} outlets '
          f'({Counter(v["type"] for v in liq.values())})')

    wb = openpyxl.load_workbook(a.source)          # formulas preserved
    ws = pick_sheet(wb, a.sheet)
    print(f'target sheet: {ws.title!r}  (previous period {ws["B2"].value})')

    # ---- classify the template rows -------------------------------------
    # Rows are located by scanning rather than hardcoded, because the KSBC and
    # Fed blocks grow and shrink as outlets open and close.
    named, bar_slots = [], []
    for r in range(4, ws.max_row + 1):
        challan = str(ws.cell(r, 3).value or '').strip()
        name = ws.cell(r, 2).value
        if challan.lower() == 'bar':
            bar_slots.append(r)          # includes blank template rows
        if name not in (None, ''):
            named.append((r, str(name).strip(), challan))

    shop_rows = [(r, n, c) for r, n, c in named if c.lower() != 'bar']
    print(f'template: {len(shop_rows)} KSBC/Fed rows, {len(bar_slots)} bar slots')

    # ---- KSBC + Fed ------------------------------------------------------
    missing, written = [], 0
    for r, name, challan in shop_rows:
        key = match_name(name, liq)
        if key is None:
            missing.append((r, name, challan))
            continue
        vals = liq[key]['vals']
        for col, brand in BRAND_COLS.items():
            ws.cell(r, ci(col)).value = round(vals[brand], 2)
        written += 1
    if missing:
        # Never guess. A shop present in the template but absent from the month's
        # liquidation means a rename or a genuinely new outlet -- both need a human.
        print('ERROR: these template rows have no matching outlet in the liquidation file:')
        for r, n, c in missing:
            print(f'   row {r}: {n!r} ({c})')
        sys.exit(1)
    print(f'KSBC/Fed rows written: {written}')

    # ---- bars: only those with sales -------------------------------------
    # Listing every bar would add ~44 permanently blank rows, so the workbook
    # carries only the ones that actually moved stock, as it always has.
    active = sorted(((n, d) for n, d in liq.items()
                     if d['type'] == 'BAR' and sum(d['vals'].values()) > 0),
                    key=lambda x: -sum(x[1]['vals'].values()))
    if len(active) > len(bar_slots):
        print(f'ERROR: {len(active)} bars sold this month but the template only has '
              f'{len(bar_slots)} bar rows. Inserting rows would shift the table range, '
              f'so add the extra rows by hand in Excel first, then re-run.')
        sys.exit(1)
    for i, row in enumerate(bar_slots):
        if i < len(active):
            name, d = active[i]
            ws.cell(row, 2).value = name
            ws.cell(row, 3).value = 'Bar'
            ws.cell(row, 6).value = d['bond']
            for col, brand in BRAND_COLS.items():
                ws.cell(row, ci(col)).value = round(d['vals'][brand], 2)
        else:
            # Blank a slot that carried a bar last month but sold nothing this
            # month, so a stale name never sits on top of zeroes.
            if ws.cell(row, 2).value not in (None, ''):
                ws.cell(row, 2).value = None
                ws.cell(row, 4).value = None
                ws.cell(row, 6).value = None
                for col in BRAND_COLS:
                    ws.cell(row, ci(col)).value = None
    print(f'bars written: {len(active)} -> ' +
          ', '.join(f'{n} {sum(d["vals"].values()):.0f}' for n, d in active))

    # ---- header dates ----------------------------------------------------
    stamp = datetime.datetime(a.year, MONTHS[month], 1)
    for coord in ('B2', 'N1'):
        if ws[coord].value is not None:
            ws[coord].value = stamp

    tmp_fmt = a.out + '.fmt.tmp'
    wb.save(tmp_fmt)

    # ---- recalculate, then graft the values onto the formatted file ------
    print('recalculating (LibreOffice, ~1 min)...')
    valued = recalculate(tmp_fmt)
    filled, total = inject_cached_values(tmp_fmt, valued, a.out)
    os.remove(tmp_fmt)
    print(f'cached values injected: {filled}/{total} formula cells')

    v = openpyxl.load_workbook(a.out, data_only=True)[ws.title]
    grand = sum(float(v.cell(r, ci('AB')).value or 0)
                for r in range(4, v.max_row + 1) if v.cell(r, 2).value)
    pts = sum(float(v.cell(r, ci('AC')).value or 0)
              for r in range(4, v.max_row + 1) if v.cell(r, 2).value)
    print(f'{month} {a.year}: {grand:,.2f} cases | {pts:,.0f} points')
    print(f'saved -> {a.out}')
    print('BUILD OK')


if __name__ == '__main__':
    main()
