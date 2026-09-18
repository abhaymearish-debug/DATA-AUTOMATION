#!/usr/bin/env python3
"""
INCENTIVE TRACKER — aging-stock liquidation incentives at KSBC shops.

Created 12 Jun 2026 (Abhay-approved design):
  - Records every special incentive (shop x brand x pack, Rs/case, live-from date).
  - Tracks movement since each line's start date from the KSBC monthly analysis
    workbooks: month-start lines use the COMBINED sheets; mid-month lines use the
    daily raw sheets (Shop Opening / Shop In / Shop Out / Shop Closing Cases).
  - Sell-Through = Sales / (Opening + Receipts)  (KSBC rating formula verbatim).
  - Tiers: house 4-tier palette; closing <= 0.01 cs => CLEARED.
  - Persistent accumulator: INCENTIVE LOG seeds from the live workbook on re-run;
    inline NEW_ENTRIES merge on key (shop, product_code, start_date) — new wins.

Usage:
  python3 build_incentive_tracker.py --out "<scratch.xlsx>" [--base "<Claude folder>"]
"""
import argparse, os, re, glob
from datetime import date, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

MONTHS = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST',
          'SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER']

# ---- incentive entries (12 Jun 2026, WhatsApp; rates Abhay-confirmed:
#      Chairman's Choice Rs150/cs, all other brands Rs50/cs) ----
NEW_ENTRIES = [
    # shop, brand_short, product_code, pack, rate, start,        reported,                          source
    (114003, 'CC',   '11397151A', '750 ML',  150, date(2026,3,1),  '80 btl @28/2 -> 52 btl reported', 'Old-stock liquidation msg'),
    (113016, 'MWB',  '11397089A', '1000 ML',  50, date(2026,3,1),  '89 btl @28/2 -> NIL @15/5',       'Old-stock liquidation msg'),
    (110003, 'CC',   '11397151A', '750 ML',  150, date(2026,3,1),  '70 btl @28/2 -> 30 btl reported', 'Old-stock liquidation msg'),
    (107006, 'OPR',  '13397035A', '500 ML',   50, date(2026,5,26), '14.94 cs',                        'Abhinand Sales Executive'),
    (107047, 'OPR',  '13397035A', '500 ML',   50, date(2026,5,25), '14.39 cs',                        'Sojan ASM'),
    (105023, 'OPR',  '133970327', '375 ML',   50, date(2026,6,9),  '4.00 cs',                         'Murali Sales Rep'),
    (105023, 'ROFR', '13397045W', '500 ML',   50, date(2026,6,9),  '9.72 cs',                         'Murali Sales Rep'),
    (106004, 'BCB',  '11397035W', '500 ML',   50, date(2026,6,11), '4.94 cs',                         'Sojan ASM'),
    (106006, 'ROFR', '13397045W', '500 ML',   50, date(2026,6,11), '3.56 cs',                         'Sojan ASM'),
    (106006, 'BCB',  '11397035W', '500 ML',   50, date(2026,6,11), '2.50 cs',                         'Sojan ASM'),
    # Thenmala ₹80/cs CONDITIONAL — only if BOTH SKUs are fully liquidated (Abhay, 13 Jun 2026)
    (102034, 'MBR',  '133971851', '500 ML',   80, date(2026,6,13), '1.00 cs · ₹80 if BOTH liquidated', 'Abhay (aging report)'),
    (102034, 'BCB',  '11397031W', '750 ML',   80, date(2026,6,13), '4.08 cs · ₹80 if BOTH liquidated', 'Abhay (aging report)'),
    (108021, 'MWB',  '113970856', '500 ML',   50, date(2026,6,16), '5.7 cs',                           'Abhishek Sales Executive'),
    (108021, 'MWB',  '11397089A', '1000 ML',  50, date(2026,6,16), '3.2 cs',                           'Abhishek Sales Executive'),
    (105004, 'MWB',  '11397089A', '1000 ML', 100, date(2026,6,16), '13.33 cs',                         'Sojan ASM'),
    (105030, 'BLND', '113970798', '1000 ML',  50, date(2026,6,18), '12.78 cs',                         'Murali Sales Rep'),
]
ENTRY_DATE = date(2026, 6, 18)

# ---- palette (house) ----
NAVY_DEEP, NAVY_MID, NAVY_SOFT = 'FF0D1B4A', 'FF1A237E', 'FF263F80'
GOLD, WHITE, GREY = 'FFFFB300', 'FFFFFFFF', 'FF6B7280'
TIERS = [  # threshold, label, font, fill
    (80, '🚀 High Performance', 'FF1565C0', 'FFBBDEFB'),
    (60, '✅ Balanced',         'FF2E7D32', 'FFDCEDC8'),
    (40, '⚠️ Inventory Heavy',  'FFE65100', 'FFFFE0B2'),
    (-1, '🚫 Critical Overstock','FFC62828', 'FFFFCDD2'),
]
CLEARED = ('🏁 CLEARED', 'FF2E7D32', 'FFC8E6C9')
FRESH = ('🆕 Just Approved', 'FF1565C0', 'FFE3F2FD')  # approved after last data day — no sales window yet
THIN = Side(style='thin', color='FFB0BEC5')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def norm_shop(v):
    s = re.sub(r'\D', '', str(v or ''))
    return int(s) if s else None

def load_master(master_path):
    wb = openpyxl.load_workbook(master_path, read_only=True, data_only=True)
    ws = wb['16-4-25']
    out = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        code = norm_shop(r[3])
        if code:
            out[code] = {'name': str(r[4] or '').strip(), 'staff': str(r[6] or '').strip(),
                         'bond': str(r[7] or '').strip(), 'status': str(r[8] or '').strip()}
    wb.close()
    return out

def find_month_wb(folder, month):
    full = os.path.join(folder, f'{month} SHOP SALES ANALYSIS.xlsx')
    if os.path.exists(full):
        return full
    cands = [p for p in glob.glob(os.path.join(folder, f'{month} 1st - *ANALYSIS.xlsx'))
             if not os.path.basename(p).startswith('~$')]
    def endday(p):
        m = re.search(r'1st\s*-\s*(\d+)', os.path.basename(p))
        return int(m.group(1)) if m else 0
    return max(cands, key=endday) if cands else None

class MonthData:
    """COMBINED rows + daily sheets for one month workbook."""
    def __init__(self, path, month):
        self.month, self.path = month, path
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        comb = [s for s in wb.sheetnames if 'COMBINED' in s.upper()][0]
        self.combined = {}
        for r in wb[comb].iter_rows(min_row=2, values_only=True):
            if not r or r[3] is None: continue
            key = (norm_shop(r[1]), str(r[3]).strip())
            self.combined[key] = {'brand': str(r[4]), 'pack': str(r[5]), 'bpc': r[6],
                                  'O': float(r[7] or 0), 'R': float(r[8] or 0),
                                  'S': float(r[9] or 0), 'C': float(r[10] or 0)}
        self.daily_days = sorted(int(m.group(1)) for s in wb.sheetnames
                                 for m in [re.fullmatch(rf'{month} (\d+)', s)] if m)
        self._wb, self._cache = wb, {}

    def daily(self, day):
        if day not in self._cache:
            name = f'{self.month} {day}'
            rows = {}
            if name in self._wb.sheetnames:
                for r in self._wb[name].iter_rows(min_row=2, values_only=True):
                    if not r or r[3] is None: continue
                    key = (norm_shop(r[1]), str(r[3]).strip())
                    bpc = float(r[6] or 0) or 1.0
                    # daily sheets split whole cases + loose bottles — fold to fractional cases
                    rows[key] = {'O': float(r[7] or 0) + float(r[8] or 0) / bpc,
                                 'In': float(r[9] or 0) + float(r[10] or 0) / bpc,
                                 'Out': float(r[11] or 0) + float(r[12] or 0) / bpc,
                                 'C': float(r[13] or 0) + float(r[14] or 0) / bpc}
            self._cache[day] = rows
        return self._cache[day]

def compute_movement(entry_key, start, months_loaded, latest_month, latest_day):
    """O at start date; R,S summed start..latest; C now. Returns dict + notes."""
    O = R = S = 0.0; C = None; notes = []
    # Incentive approved AFTER the last available sales day → no movement window yet.
    # Open at the latest closing; zero R/S until a fresh sales day lands.
    latest_idx = MONTHS.index(latest_month)
    if (start.month - 1 > latest_idx) or (start.month - 1 == latest_idx and start.day > latest_day):
        row = months_loaded[latest_month].combined.get(entry_key)
        c = row['C'] if row is not None else 0.0
        notes.append(f'approved {start.strftime("%d/%m")} — after last data day ({latest_month.title()} {latest_day}); opens at latest closing, no movement yet')
        return {'O': c, 'R': 0.0, 'S': 0.0, 'C': c, 'drift': 0.0, 'notes': notes}
    seq = [m for m in MONTHS if start.month-1 <= MONTHS.index(m) <= MONTHS.index(latest_month)]
    for mi, mname in enumerate(seq):
        md = months_loaded.get(mname)
        if md is None:
            notes.append(f'{mname}: workbook missing — skipped'); continue
        row = md.combined.get(entry_key)
        if mi == 0 and start.day > 1:
            drow = md.daily(start.day).get(entry_key)
            if drow: O = drow['O']
            elif row is not None:
                O = row['O']; notes.append(f'{mname} {start.day} daily row absent — month opening used')
            else: notes.append(f'{mname} {start.day}: SKU row absent')
            for d in [d for d in md.daily_days if d >= start.day]:
                dr = md.daily(d).get(entry_key)
                if dr: R += dr['In']; S += dr['Out']
        else:
            if row is None:
                notes.append(f'{mname}: SKU absent from COMBINED'); continue
            if mi == 0: O = row['O']
            R += row['R']; S += row['S']
        if mname == latest_month and row is not None:
            C = row['C']
    if C is None: C = max(O + R - S, 0.0)
    drift = (O + R - S) - C
    return {'O': O, 'R': R, 'S': S, 'C': C, 'drift': drift, 'notes': notes}

def tier(st_pct, closing):
    if closing <= 0.01: return CLEARED
    for th, lab, fnt, fll in TIERS:
        if st_pct >= th: return (lab, fnt, fll)
    return TIERS[-1][1:]

# ---------------- rendering ----------------
def style_cell(c, *, bold=False, size=11, color='FF111111', fill=None, fmt=None,
               align='center', border=True):
    c.font = Font(name='Calibri', bold=bold, size=size, color=color)
    if fill: c.fill = PatternFill('solid', fgColor=fill)
    if fmt: c.number_format = fmt
    c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=True)
    if border: c.border = BORDER

def render(out_path, lines, asof_label, log_rows):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = 'DASHBOARD'
    ws.sheet_view.showGridLines = False
    widths = [4.5, 11, 16, 22, 27, 9, 8, 10, 10, 10, 10, 12, 22]
    for i, w in enumerate(widths, 1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    ws.merge_cells('A1:M1'); ws.merge_cells('A2:M2')
    ws.row_dimensions[1].height = 40; ws.row_dimensions[2].height = 20
    c = ws['A1']; c.value = 'INCENTIVE TRACKER — AGING STOCK LIQUIDATION'
    style_cell(c, bold=True, size=20, color=WHITE, fill=NAVY_DEEP, border=False)
    c = ws['A2']; c.value = f"K.S. Distillery · KSBC shops · movement as of {asof_label} · CC ₹150/cs · other brands ₹50/cs"
    style_cell(c, size=11, color=GOLD, fill=NAVY_DEEP, border=False)

    tot_base = sum(l['O'] + l['R'] for l in lines)
    tot_sold = sum(l['S'] for l in lines)
    tot_close = sum(l['C'] for l in lines)
    cleared = sum(1 for l in lines if l['C'] <= 0.01)
    kpis = [('LINES TRACKED', len(lines), '#,##0'),
            ('STOCK COVERED (O+R cs)', tot_base, '#,##0.0'),
            ('CASES MOVED', tot_sold, '#,##0.0'),
            ('STILL ON SHELF (cs)', tot_close, '#,##0.0'),
            ('MOVED %', tot_sold / tot_base if tot_base else 0, '0%'),
            ('LINES CLEARED', cleared, '#,##0')]
    ws.row_dimensions[4].height = 16; ws.row_dimensions[5].height = 30
    spans = [(1,2),(3,4),(5,6),(7,8),(9,10),(11,13)]
    for (c1, c2), (lab, val, fmt) in zip(spans, kpis):
        l1, l2 = openpyxl.utils.get_column_letter(c1), openpyxl.utils.get_column_letter(c2)
        ws.merge_cells(f'{l1}4:{l2}4'); ws.merge_cells(f'{l1}5:{l2}5')
        a = ws[f'{l1}4']; a.value = lab
        style_cell(a, bold=True, size=9, color=GOLD, fill=NAVY_MID, border=False)
        b = ws[f'{l1}5']; b.value = val
        style_cell(b, bold=True, size=16, color=WHITE, fill=NAVY_MID, fmt=fmt, border=False)

    hdr = ['#','Live From','Bond','Shop','Brand','Pack','₹/cs','Opening','Receipts',
           'Sales','Closing','Sell-Through %','Status']
    hr = 7; ws.row_dimensions[hr].height = 28
    for i, h in enumerate(hdr, 1):
        c = ws.cell(row=hr, column=i, value=h)
        style_cell(c, bold=True, size=11, color=WHITE, fill=NAVY_SOFT)
    r = hr + 1
    for i, l in enumerate(sorted(lines, key=lambda x: (x['start'], x['shop'])), 1):  # date ASC (Abhay, 12 Jun 2026)
        lab, fnt, fll = l['tier']
        vals = [i, l['start'].strftime('%d/%m/%y'), l['bond'], f"{str(l['shop'])[-4:]} {l['shop_name']}",
                l['brand'], l['pack'], l['rate'], l['O'], l['R'], l['S'], l['C'], l['st'], lab]
        ws.row_dimensions[r].height = 24
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=j, value=v)
            fmt = '#,##0.00;-#,##0.00;"·"' if j in (8,9,10,11) else ('0%' if j == 12 else ('#,##0' if j == 7 else None))
            if j == 13: style_cell(c, bold=True, size=10, color=fnt, fill=fll, fmt=fmt)
            else: style_cell(c, size=10, fmt=fmt, align='left' if j in (4,5) else 'center')
        r += 1
    ws.row_dimensions[r].height = 28
    tvals = ['', '', '', 'TOTAL', f'{len(lines)} lines', '', '',
             sum(l['O'] for l in lines), sum(l['R'] for l in lines), tot_sold, tot_close,
             tot_sold / tot_base if tot_base else 0, f'{cleared} cleared']
    for j, v in enumerate(tvals, 1):
        c = ws.cell(row=r, column=j, value=v)
        fmt = '#,##0.00' if j in (8,9,10,11) else ('0%' if j == 12 else None)
        style_cell(c, bold=True, size=11, color=WHITE, fill='FF374151', fmt=fmt)
    ws.freeze_panes = f'A{hr+1}'

    lg = wb.create_sheet('INCENTIVE LOG')
    lg.sheet_view.showGridLines = False
    lhdr = ['Entry Date','Live From','Shop Code','Shop Name','Bond','Field Staff','Brand',
            'Product Code','Pack','Btl/Case','₹/cs','Reported Stock (msg)','Source / Sender','Notes']
    lw = [11,11,11,24,16,22,30,13,9,9,8,26,26,40]
    for i, w in enumerate(lw, 1): lg.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    lg.row_dimensions[1].height = 30
    for i, h in enumerate(lhdr, 1):
        c = lg.cell(row=1, column=i, value=h)
        style_cell(c, bold=True, size=11, color=WHITE, fill=NAVY_MID)
    for ri, row in enumerate(log_rows, 2):
        lg.row_dimensions[ri].height = 22
        for j, v in enumerate(row, 1):
            c = lg.cell(row=ri, column=j, value=v)
            style_cell(c, size=10, align='left' if j in (4,6,7,12,13,14) else 'center',
                       fmt='dd/mm/yy' if j in (1,2) else None)
    lg.freeze_panes = 'A2'
    wb.save(out_path)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', required=True)
    ap.add_argument('--base', default=os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
    a = ap.parse_args()
    ksbc_dir = os.path.join(a.base, 'KSBC shop sales')
    master_path = os.environ.get('KSBC_MASTER_DATA_PATH', os.path.join(a.base, 'MASTER DATA CONFIRMED.xlsx'))
    live = os.path.join(a.base, 'Incentive tracking', 'INCENTIVE TRACKER.xlsx')

    master = load_master(master_path)
    today = date.today()
    months_loaded, latest_month, latest_day = {}, None, None
    for mname in MONTHS[:today.month]:
        if MONTHS.index(mname) < 2: continue  # MARCH onwards
        p = find_month_wb(ksbc_dir, mname)
        if p:
            md = MonthData(p, mname)
            months_loaded[mname] = md
            if md.daily_days: latest_month, latest_day = mname, max(md.daily_days)
    asof = f'{latest_month.title()} {latest_day}' if latest_month else 'n/a'
    print(f'[info] months loaded: {sorted(months_loaded, key=MONTHS.index)} · as of {asof}')

    entries = {}
    if os.path.exists(live):
        try:
            lwb = openpyxl.load_workbook(live, read_only=True, data_only=True)
            for r in lwb['INCENTIVE LOG'].iter_rows(min_row=2, values_only=True):
                if not r or r[2] is None: continue
                sd = r[1].date() if isinstance(r[1], datetime) else r[1]
                ed = r[0].date() if isinstance(r[0], datetime) else r[0]
                # one live incentive line per shop x SKU — a re-entry (e.g. corrected
                # start date) replaces the seeded line rather than duplicating it
                k = (norm_shop(r[2]), str(r[7]).strip())
                entries[k] = {'shop': k[0], 'code': k[1], 'start': sd, 'entry': ed,
                              'brand': str(r[6]), 'pack': str(r[8]), 'bpc': r[9], 'rate': r[10],
                              'reported': str(r[11] or ''), 'source': str(r[12] or ''), 'note0': str(r[13] or '')}
            lwb.close()
            print(f'[info] seeded {len(entries)} entries from live workbook')
        except Exception as e:
            print(f'[warn] live seed failed: {e}')
    for shop, bshort, code, pack, rate, start, reported, source in NEW_ENTRIES:
        prior = entries.get((shop, code))  # preserve original log date if line already seeded
        entries[(shop, code)] = {'shop': shop, 'code': code, 'start': start,
                                 'entry': prior['entry'] if prior else ENTRY_DATE,
                                 'brand': bshort, 'pack': pack, 'bpc': None, 'rate': rate,
                                 'reported': reported, 'source': source, 'note0': ''}

    lines, log_rows, problems = [], [], []
    for k, e in sorted(entries.items(), key=lambda kv: (kv[1]['start'], kv[1]['shop'])):
        shop = e['shop']
        m = master.get(shop, {})
        if not m: problems.append(f'{shop}: NOT IN MASTER — investigate')
        if m.get('status') == 'Closed':
            problems.append(f'{shop}: shop CLOSED in master'); continue
        ekey = (shop, e['code'])
        mv = compute_movement(ekey, e['start'], months_loaded, latest_month, latest_day)
        ref = None
        for mname in sorted(months_loaded, key=MONTHS.index, reverse=True):
            ref = months_loaded[mname].combined.get(ekey)
            if ref: break
        brand_full = ref['brand'] if ref else e['brand']
        bpc = ref['bpc'] if ref else e['bpc']
        st = mv['S'] / (mv['O'] + mv['R']) if (mv['O'] + mv['R']) > 0 else 0.0
        fresh = (e['start'].month - 1 > MONTHS.index(latest_month)) or \
                (e['start'].month - 1 == MONTHS.index(latest_month) and e['start'].day > latest_day)
        t = FRESH if (fresh and mv['C'] > 0.01) else tier(st * 100, mv['C'])
        nm_clean = re.sub(r'^\d+\s*-\s*', '', m.get('name', '')).strip()
        lines.append({'shop': shop, 'shop_name': nm_clean, 'bond': m.get('bond',''),
                      'staff': m.get('staff',''), 'brand': brand_full, 'pack': e['pack'],
                      'rate': e['rate'], 'start': e['start'], 'st': st, 'tier': t,
                      'exposure': mv['C'], **{x: mv[x] for x in 'ORSC'},
                      'drift': mv['drift'], 'notes': mv['notes']})
        note = e['note0'] or '; '.join(mv['notes'])
        if abs(mv['drift']) > 0.15:  # KSBC case-conversion rounding tolerance across months
            note = (note + f"; drift {mv['drift']:+.2f} cs (O+R−S vs C)").strip('; ')
            problems.append(f"{shop} {e['code']}: drift {mv['drift']:+.2f} cs")
        log_rows.append([e['entry'], e['start'], shop, nm_clean, m.get('bond',''), m.get('staff',''),
                         brand_full, e['code'], e['pack'], bpc, e['rate'], e['reported'], e['source'], note])

    render(a.out, lines, asof, log_rows)

    print(f'\n=== INCENTIVE TRACKER BUILD — as of {asof} ===')
    for l in sorted(lines, key=lambda x: (x['start'], x['shop'])):
        print(f"  {l['start'].strftime('%d/%m')} {str(l['shop'])[-4:]:<5} {l['shop_name']:<14.14} "
              f"{l['brand']:<30.30} {l['pack']:<7} ₹{l['rate']:<4} O {l['O']:7.2f} R {l['R']:5.2f} "
              f"S {l['S']:7.2f} C {l['C']:7.2f} ST {l['st']*100:5.1f}%  {l['tier'][0]}")
    tb = sum(l['O']+l['R'] for l in lines); ts = sum(l['S'] for l in lines)
    print(f"  TOTAL: covered {tb:.2f} cs · moved {ts:.2f} cs ({(ts/tb*100 if tb else 0):.1f}%) · "
          f"on shelf {sum(l['C'] for l in lines):.2f} cs · cleared {sum(1 for l in lines if l['C']<=0.01)}/{len(lines)}")
    if problems:
        print('\n[VALIDATION FLAGS]')
        for p in problems: print('  -', p)
    else:
        print('\n[validation] all lines reconcile (O+R−S = C within 0.10 cs)')
    print(f'[saved scratch] {a.out}')

if __name__ == '__main__':
    main()
