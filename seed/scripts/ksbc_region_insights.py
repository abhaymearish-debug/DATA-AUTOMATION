#!/usr/bin/env python3
"""
ksbc_region_insights.py — BOND INSIGHTS section on every region sheet, and
the pipeline's FINAL write step (supersedes the standalone
fix_region_sheet_styling run — this script calls its apply_all() in-session).

Layout approved by Abhay 16 Jul 2026 (INSIGHTS DEMO v5 screenshot):
below each region sheet's TOTAL row —
  📅 DAILY TREND — <CUR> vs <PRIOR> (same day)
     table (day · cur cs+bar · prior cs+bar · Δ) + cumulative race line chart
  🍾 SALES MIX — BY BRAND | BY PACK
     twin panels: name · sold cs (bold) · gold bar · share%
  (METHOD footnote REMOVED 16 Jul 2026 PM at Abhay's request — do not re-add.)
Explicitly REMOVED by Abhay (do not re-add): KPI tile strip, auto-insight
text lines, per-SKU cover columns, STOCKOUT WATCH table, DETAIL ⚡ pulse rows.

Column map (cols A=Shop Code and C=Field Staff are HIDDEN on region sheets
since 16 Jul 2026 — panels avoid them): day/name labels in B; values D;
bar cells E; prior values F; bar G; Δ H; chart anchored I; mix pack panel
G:J; cumulative chart helpers M:N (font size 1, white — invisible).

WHY THIS MUST BE THE FINAL WRITE STEP (both verified openpyxl 3.1.5):
  * charts: openpyxl 3.1.5 round-trips them, but every rebuild recreates
    the race charts from scratch (wipe_below clears ws._charts first —
    without it each run stacks 15 duplicates).
  * openpyxl round-trips DROP style-only merged cells (hero bands): this
    script therefore imports fix_region_sheet_styling and runs apply_all()
    in the SAME session, after wiping the old insight zone and before
    building the new one (apply_all's _rebuild_j_cf resets ALL region CF,
    so the insight data-bars are added after it — order matters).
Receipt notes (step 13.5) survive: no row above TOTAL moves, comments
round-trip fine (verified 1,332/1,332).

Idempotent: everything below each region TOTAL row is wiped and rebuilt.
Col A below TOTAL stays EMPTY — every script that anchors on A=='TOTAL'
(styling, sorter, driver, BP builder) must keep finding only the real one.

Prior-month source: '<PRIOR> SHOP SALES ANALYSIS.xlsx' (fallback: newest
'<PRIOR> 1st - *th ANALYSIS.xlsx') in the same folder; days limited to the
current max day. Missing prior workbook -> single-series build, no abort.
Trend totals are daily-sheet sums (audit basis) — under a CUMULATIVE
override they may drift from the period-true TOTAL row; WARN only.

Usage: python3 ksbc_region_insights.py "<workbook.xlsx>"
Exit 0 ok · 1 structural failure (no daily sheets / no region TOTAL).
"""
from __future__ import annotations
import os, re, sys, glob, time, datetime
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import DataBarRule
from openpyxl.chart import LineChart, Reference
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker

from ksbc_alias_guard import canonical_code
from ksbc_detail_nested import _norm_brand
import fix_region_sheet_styling as sty
import resize_comment_boxes  # FINAL box-fit pass (see wb.save below)

FONT = 'Trebuchet MS'
NAVY = 'FF1A237E'; NAVY_SOFT = 'FF263F80'
GOLD = 'FFFFB300'; GOLD_DIM = 'FFFFD54F'
RED = 'FFC62828'; GRN = 'FF2E7D32'; AMB = 'FFE65100'
GREY = 'FF6B7280'; GREY_L = 'FFF3F4F6'; DARK = 'FF374151'
GOLD_TINT = 'FFFFF8E1'; INK = 'FF111827'
HAIR = Side(style='thin', color='FFE5E7EB')

MONTHS = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY",
          "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]
MNUM = {m: i + 1 for i, m in enumerate(MONTHS)}
BRAND_SHORT = {
    'OLD PEARL NO.1 MATURED XXX RUM': 'OLD PEARL',
    'K.S 99 LIFE TIME MATURED XXX RUM': 'K.S 99',
    'BLENDERS CHOICE NO.1 BRANDY': 'BLENDERS CHOICE',
    'MORNING WALKERS XO BRANDY': 'MORNING WALKER',
    'MAGIC BLEND RESERVED XXX RUM': 'MAGIC BLEND',
    'CHAIRMANS CHOICE XO BRANDY': "CHAIRMAN'S CHOICE",
    'BCB NO.1 CLASSIC BRANDY': 'BCB',
    'ROYAL OLD FORT NO.1 XXX RUM': 'ROYAL OLD FORT',
}


def _num(v):
    try:
        return float(str(v).replace(',', '').strip() or 0)
    except (ValueError, TypeError):
        return 0.0


def detect_month(wb):
    pat = re.compile(r'^(' + '|'.join(MONTHS) + r') (\d{1,2})$')
    days = defaultdict(list)
    for n in wb.sheetnames:
        m = pat.match(n)
        if m:
            days[m.group(1)].append(int(m.group(2)))
    if not days:
        return None, 0
    mon = max(days, key=lambda k: len(days[k]))
    return mon, max(days[mon])


def bond_codes(wb):
    """code -> bond from each region sheet's (hidden) col A."""
    out = {}
    for b in sty.BONDS:
        if b not in wb.sheetnames:
            continue
        ws = wb[b]
        for r in range(5, ws.max_row + 1):
            v = ws.cell(r, 1).value
            if isinstance(v, str) and v.strip().upper() == 'TOTAL':
                break
            try:
                out[canonical_code(int(str(v).strip()))] = b
            except (ValueError, TypeError):
                continue
    return out


def scan(wb, month, maxday, code2bond):
    """bond -> {day: cs}, bond -> {brand: cs}, bond -> {pack: cs}"""
    day = defaultdict(lambda: defaultdict(float))
    brand = defaultdict(lambda: defaultdict(float))
    pack = defaultdict(lambda: defaultdict(float))
    n = 0
    for d in range(1, maxday + 1):
        sn = f'{month} {d}'
        if sn not in wb.sheetnames:
            continue
        n += 1
        for r in wb[sn].iter_rows(min_row=2, max_col=15, values_only=True):
            if not r or r[1] is None:
                continue
            try:
                code = canonical_code(int(str(r[1]).strip()))
            except (ValueError, TypeError):
                continue
            b = code2bond.get(code)
            if b is None:
                continue
            bpc = _num(r[6]) or 1
            s = _num(r[11]) + _num(r[12]) / bpc
            if not s:
                continue
            day[b][d] += s
            brand[b][_norm_brand(r[4])] += s
            pack[b][str(r[5]).strip()] += s
    return day, brand, pack, n


def find_prior(path, cur_month):
    """Prior workbook lives in the canonical 'KSBC shop sales' folder,
    resolved from THIS script's location — the build usually runs on a
    /tmp scratch copy, so the scratch's dirname is only the fallback."""
    here = os.path.dirname(os.path.abspath(__file__))
    canonical = os.path.normpath(os.path.join(here, '..', '..', 'KSBC shop sales'))
    prior = MONTHS[(MNUM[cur_month] - 2) % 12]
    for folder in (canonical, os.path.dirname(os.path.abspath(path))):
        if not os.path.isdir(folder):
            continue
        full = os.path.join(folder, f'{prior} SHOP SALES ANALYSIS.xlsx')
        if os.path.exists(full):
            return prior, full
        cands = glob.glob(os.path.join(folder, f'{prior} 1st - *ANALYSIS.xlsx'))
        if cands:
            return prior, max(cands, key=os.path.getmtime)
    return prior, None


def wipe_below(ws, total_r):
    # openpyxl 3.1.5 DOES round-trip charts (contrary to older docs) — clear
    # them here or every rebuild stacks 15 duplicates (caught on the
    # idempotency rerun, 16 Jul 2026).
    ws._charts = []
    for m in list(ws.merged_cells.ranges):
        if m.min_row > total_r:
            ws.unmerge_cells(str(m))
    for r in range(total_r + 1, ws.max_row + 1):
        for c in range(1, 15):
            cell = ws.cell(r, c)
            cell.value = None
            cell.comment = None
            cell.fill = PatternFill()
            cell.font = Font()
            cell.border = Border()
            cell.alignment = Alignment()
            cell.number_format = 'General'
    for r in [r for r in ws.row_dimensions if r > total_r]:
        del ws.row_dimensions[r]


def find_total(ws):
    for r in range(5, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.strip().upper() == 'TOTAL':
            return r
    return None


def band(ws, row, txt):
    ws.row_dimensions[row].height = 26
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
    for c in range(1, 11):
        cell = ws.cell(row, c)
        cell.fill = PatternFill('solid', fgColor=NAVY_SOFT)
        cell.border = Border(bottom=Side(style='medium', color=GOLD))
    t = ws.cell(row, 1)
    t.value = txt
    t.font = Font(name=FONT, size=12, bold=True, color=GOLD_DIM)
    t.alignment = Alignment('left', 'center', indent=1)


def hdr(ws, row, cells, merges=(), c1=2, c2=10):
    ws.row_dimensions[row].height = 19
    for a, b in merges:
        ws.merge_cells(start_row=row, start_column=a, end_row=row, end_column=b)
    for c in range(c1, c2 + 1):
        cell = ws.cell(row, c)
        cell.fill = PatternFill('solid', fgColor=NAVY)
        cell.border = Border(bottom=Side(style='thin', color=GOLD_DIM))
    for c, txt in cells:
        cell = ws.cell(row, c)
        cell.value = txt
        cell.font = Font(name=FONT, size=9, bold=True, color='FFFFFFFF')
        cell.alignment = Alignment('center', 'center')
    if any(c == 6 and t == 'Pack' for c, t in cells):
        ws.cell(row, 6).border = Border(left=Side(style='medium', color=GOLD_DIM),
                                        bottom=Side(style='thin', color=GOLD_DIM))


def brand_disp(nb):
    return BRAND_SHORT.get(nb, ' '.join(nb.split()[:2]).title().upper()[:18])


def build_section(ws, bond, cur_m, pri_m, maxday, year, cday, pday, bmix, pmix,
                  has_prior):
    total_r = find_total(ws)
    if total_r is None:
        return None
    wipe_below(ws, total_r)
    row = total_r + 2
    tag = f'{cur_m} vs {pri_m} (same day)' if has_prior else f'{cur_m} (no {pri_m} workbook found)'
    band(ws, row, f'📅 DAILY TREND — {tag}')
    row += 1
    hdr(ws, row, [(2, 'Day'), (4, f'{cur_m.title()} cs'), (5, f'{pri_m.title()} cs'), (6, 'Δ cs')],
        c1=2, c2=6)
    hr = row
    row += 1
    first = row
    cumc = cump = 0.0
    for d in range(1, maxday + 1):
        cs = round(cday.get(d, 0.0), 1)
        ps = round(pday.get(d, 0.0), 1) if has_prior else 0.0
        cumc += cs
        cump += ps
        wd = datetime.date(year, MNUM[cur_m], d).strftime('%a')
        sun = wd == 'Sun'
        lab = ws.cell(row, 2)
        lab.value = f'{wd} {d:02d} {cur_m.title()[:3]}'
        lab.font = Font(name=FONT, size=9, bold=sun, color=AMB if sun else INK)
        lab.alignment = Alignment('left', 'center', indent=1)
        for col, v in ((4, cs), (5, ps), (6, round(cs - ps, 1))):
            cell = ws.cell(row, col)
            cell.value = v
            cell.number_format = '0.0'
        ws.cell(row, 4).font = Font(name=FONT, size=9, bold=True, color=INK)
        ws.cell(row, 4).alignment = Alignment('center', 'center')
        ws.cell(row, 5).font = Font(name=FONT, size=9, color=GREY)
        ws.cell(row, 5).alignment = Alignment('center', 'center')
        dc = ws.cell(row, 6)
        dc.font = Font(name=FONT, size=9, bold=True, color=GRN if cs - ps >= 0 else RED)
        dc.alignment = Alignment('center', 'center')
        dc.number_format = '+0.0;-0.0;"·"'
        for col, cv in ((13, round(cumc, 1)), (14, round(cump, 1))):
            cell = ws.cell(row, col)
            cell.value = cv
            cell.font = Font(name=FONT, size=1, color='FFFFFFFF')
        for c in range(2, 7):
            cell = ws.cell(row, c)
            cell.border = Border(bottom=HAIR)
            if sun:
                cell.fill = PatternFill('solid', fgColor=GOLD_TINT)
        ws.row_dimensions[row].height = 16
        row += 1
    last = row - 1
    ws.conditional_formatting.add(
        f'D{first}:D{last}', DataBarRule(start_type='num', start_value=0,
                                         end_type='max', color=GOLD[2:], showValue=True))
    if has_prior:
        ws.conditional_formatting.add(
            f'E{first}:E{last}', DataBarRule(start_type='num', start_value=0,
                                             end_type='max', color='B8C2CC', showValue=True))
    for c in range(2, 7):
        cell = ws.cell(row, c)
        cell.fill = PatternFill('solid', fgColor=DARK)
        cell.border = Border(top=Side(style='medium', color=GOLD))
    tvals = {2: 'TOTAL', 4: round(cumc, 1), 5: round(cump, 1), 6: round(cumc - cump, 1)}
    for c, v in tvals.items():
        cell = ws.cell(row, c)
        cell.value = v
        cell.font = Font(name=FONT, size=10, bold=True,
                         color=GOLD_DIM if c == 6 else 'FFFFFFFF')
        cell.alignment = (Alignment('left', 'center', indent=1) if c == 2 else
                          Alignment('center', 'center'))
        if isinstance(v, float):
            cell.number_format = '+0.0;-0.0' if c == 6 else '0.0'
    ws.row_dimensions[row].height = 19
    # AVG / DAY summary row (Abhay 16 Jul 2026 PM) — per-day rate for both
    # months + the Δ per day, directly under TOTAL in a softer navy band.
    row += 1
    for c in range(2, 7):
        cell = ws.cell(row, c)
        cell.fill = PatternFill('solid', fgColor=NAVY_SOFT)
        cell.border = Border(bottom=Side(style='thin', color=GOLD_DIM))
    avals = {2: 'AVG / DAY', 4: round(cumc / maxday, 1),
             5: round(cump / maxday, 1) if has_prior else None,
             6: round((cumc - cump) / maxday, 1) if has_prior else None}
    for c, v in avals.items():
        if v is None:
            continue
        cell = ws.cell(row, c)
        cell.value = v
        cell.font = Font(name=FONT, size=9, bold=(c != 2), italic=(c == 2),
                         color=GOLD_DIM if c == 6 else 'FFFFFFFF')
        cell.alignment = (Alignment('left', 'center', indent=1) if c == 2 else
                          Alignment('center', 'center'))
        if isinstance(v, float):
            cell.number_format = '+0.0" cs/d";-0.0" cs/d"' if c == 6 else '0.0" cs/d"'
    ws.row_dimensions[row].height = 17
    trow = row
    # cumulative race chart (recreated every build — openpyxl drops charts on load)
    ws.cell(hr, 13).value = cur_m.title()
    ws.cell(hr, 14).value = pri_m.title()
    for cc in (13, 14):
        ws.cell(hr, cc).font = Font(name=FONT, size=1, color='FFFFFFFF')
    ch = LineChart()
    ch.style = 12
    ch.title = (f'CUMULATIVE RACE — {cur_m} vs {pri_m} (days 1-{maxday})'
                if has_prior else f'CUMULATIVE — {cur_m} (days 1-{maxday})')
    data = Reference(ws, min_col=13, min_row=hr, max_col=14 if has_prior else 13, max_row=last)
    cats = Reference(ws, min_col=2, min_row=first, max_row=last)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    s0 = ch.series[0]
    s0.graphicalProperties.line.solidFill = 'FFB300'
    s0.graphicalProperties.line.width = 30000
    s0.smooth = True
    if has_prior and len(ch.series) > 1:
        s1 = ch.series[1]
        s1.graphicalProperties.line.solidFill = '8FA8D8'
        s1.graphicalProperties.line.width = 16000
        s1.smooth = True
        try:
            s1.graphicalProperties.line.dashStyle = 'dash'
        except Exception:
            pass
    ch.legend.position = 'b'
    ch.y_axis.majorGridlines = None
    # GRID-SNAPPED placement (Abhay 16 Jul 2026 PM — the fixed-cm chart
    # overflowed the section on his display): TwoCellAnchor editAs='twoCell'
    # pins the frame to cells, so it renders flush on any screen/zoom —
    # top at the trend header row (below the band's gold rule), left at
    # col I, right at the end of col J, bottom at the TOTAL row's bottom.
    # Small EMU insets keep the band rule and gridline edges visible.
    anchor = TwoCellAnchor(editAs='twoCell')
    anchor._from = AnchorMarker(col=6, colOff=40000, row=hr - 1, rowOff=20000)
    anchor.to = AnchorMarker(col=10, colOff=-40000, row=trow, rowOff=-20000)
    ws.add_chart(ch, anchor)
    row += 2
    # sales mix
    band(ws, row, '🍾 SALES MIX — BY BRAND  |  BY PACK')
    row += 1
    hdr(ws, row, [(2, 'Brand'), (4, 'Sold cs'), (5, 'Share'),
                  (6, 'Pack'), (7, 'Sold cs'), (8, 'Share')],
        c1=2, c2=8)
    row += 1
    tot = sum(bmix.values()) or 1.0
    mb = sorted(((brand_disp(b), v) for b, v in bmix.items() if v > 0.05),
                key=lambda x: -x[1])
    mp = sorted(((p, v) for p, v in pmix.items() if v > 0.05), key=lambda x: -x[1])
    mstart = row
    for i in range(max(len(mb), len(mp), 1)):
        for c in range(2, 9):
            ws.cell(row, c).border = Border(bottom=HAIR)
        if i < len(mb):
            nm, v = mb[i]
            lab = ws.cell(row, 2); lab.value = nm
            lab.font = Font(name=FONT, size=9, color=INK)
            lab.alignment = Alignment('left', 'center', indent=1)
            val = ws.cell(row, 4); val.value = round(v, 1); val.number_format = '0.0'
            val.font = Font(name=FONT, size=9, bold=True, color=INK)
            val.alignment = Alignment('center', 'center')
            sh = ws.cell(row, 5); sh.value = v / tot; sh.number_format = '0%'
            sh.font = Font(name=FONT, size=9, color=GREY)
            sh.alignment = Alignment('center', 'center')
        if i < len(mp):
            nm, v = mp[i]
            lab = ws.cell(row, 6); lab.value = nm
            lab.font = Font(name=FONT, size=9, color=INK)
            lab.alignment = Alignment('left', 'center', indent=1)
            val = ws.cell(row, 7); val.value = round(v, 1); val.number_format = '0.0'
            val.font = Font(name=FONT, size=9, bold=True, color=INK)
            val.alignment = Alignment('center', 'center')
            sh = ws.cell(row, 8); sh.value = v / tot; sh.number_format = '0%'
            sh.font = Font(name=FONT, size=9, color=GREY)
            sh.alignment = Alignment('center', 'center')
        ws.cell(row, 6).border = Border(left=Side(style='medium', color=GOLD_DIM),
                                        bottom=HAIR)
        ws.row_dimensions[row].height = 16
        row += 1
    if mb:
        ws.conditional_formatting.add(
            f'D{mstart}:D{row - 1}', DataBarRule(start_type='num', start_value=0,
                                                 end_type='max', color=GOLD[2:], showValue=True))
    if mp:
        ws.conditional_formatting.add(
            f'G{mstart}:G{row - 1}', DataBarRule(start_type='num', start_value=0,
                                                 end_type='max', color=GOLD[2:], showValue=True))
    return cumc


def main(path):
    t0 = time.time()
    wb = openpyxl.load_workbook(path)
    cur_m, maxday = detect_month(wb)
    if not cur_m:
        print('FAIL: no daily raw sheets found')
        return 1
    now = datetime.date.today()
    year = now.year - 1 if MNUM[cur_m] > now.month + 1 else now.year
    code2bond = bond_codes(wb)
    if not code2bond:
        print('FAIL: no region shop codes found')
        return 1
    cday, cbrand, cpack, nsheets = scan(wb, cur_m, maxday, code2bond)
    pri_m, pri_path = find_prior(path, cur_m)
    pday = defaultdict(lambda: defaultdict(float))
    has_prior = False
    if pri_path:
        pwb = openpyxl.load_workbook(pri_path, read_only=True)
        pday, _, _, pn = scan(pwb, pri_m, maxday, code2bond)
        pwb.close()
        has_prior = pn > 0
    # order matters: wipe -> styling apply_all (resets CF, heroes, hidden
    # cols, staff header) -> build sections (adds data-bar CF + charts) -> save
    totals = {}
    for b in sty.BONDS:
        if b in wb.sheetnames:
            tr = find_total(wb[b])
            if tr:
                wipe_below(wb[b], tr)
    sty.apply_all(wb)
    warn = []
    built = 0
    for b in sty.BONDS:
        if b not in wb.sheetnames:
            continue
        ws = wb[b]
        got = build_section(ws, b, cur_m, pri_m, maxday, year,
                            cday.get(b, {}), pday.get(b, {}),
                            cbrand.get(b, {}), cpack.get(b, {}), has_prior)
        if got is None:
            print(f'  · {b:18s} SKIPPED (no TOTAL row)')
            continue
        built += 1
        tr = find_total(ws)
        cellv = ws.cell(tr, 6).value
        if isinstance(cellv, (int, float)) and abs(cellv - got) > 0.15:
            warn.append((b, round(got, 2), round(cellv, 2)))
    wb.save(path)
    # FINAL WRITE — size every comment box to fit its text. openpyxl's save
    # (line above) silently reverts all legacy-comment VML boxes to the
    # 144x79 default, clipping any receipt note with 2+ dates (a 5 cs /
    # 4-day receipt renders as '1 cs'). This pure ZIP/VML pass runs AFTER the
    # save and never reloads via openpyxl, so its sizes are what Excel shows.
    # Must stay the last file write here; non-fatal so it can never gate a build.
    try:
        _nb, _hlo, _hhi = resize_comment_boxes.resize_boxes(path)
        if _nb:
            print(f'comment boxes sized to fit: {_nb} '
                  f'(height {_hlo}..{_hhi}px, width 260px)')
    except Exception as _e:
        print(f'WARN: comment-box resize skipped '
              f'({_e.__class__.__name__}: {_e}) — notes may clip')
    pl = pri_m if has_prior else f'{pri_m} MISSING'
    print(f'REGION INSIGHTS OK · {built} bonds · {cur_m} days 1-{maxday} '
          f'({nsheets} daily sheets) vs {pl} · styling re-applied in-session · '
          f'{time.time() - t0:.1f}s')
    if warn:
        print(f'WARN: {len(warn)} bond(s) trend-total vs TOTAL-row drift >0.15 cs '
              f'(cumulative override / rounding — informational):')
        for b, a, c in warn:
            print(f'  {b}: daily-sum {a} vs TOTAL row {c}')
    return 0


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('usage: python3 ksbc_region_insights.py <workbook.xlsx>')
        sys.exit(1)
    sys.exit(main(sys.argv[1]))
