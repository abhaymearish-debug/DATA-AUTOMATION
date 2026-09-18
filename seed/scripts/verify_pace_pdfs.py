"""
Layout + figure verification for the daily-pace PDFs.

Collision test compares BASELINES, not glyph boxes. A 22pt digit's box carries
~7pt of empty descender space, so two stacked numeric rows always "overlap" by
box even when the ink is clearly separated -- that produced three false alarms
in a row. Two spans genuinely collide only when they share a baseline AND
overlap horizontally.
"""
import fitz, sys, datetime as _dt
sys.path.insert(0, '/sessions/lucid-dazzling-bohr/mnt/Claude/.claude/scripts')
import pace_data as PD, pace_engine as PE, pace_shops as PS

base = PD._default_base(); h = PE._load_history(base)
plan = PE.build_plan(base, hist_rows=h)
sd, matched = PS.build_shop_daily(base, plan, h)

def spans(pg, y0, y1):
    out = []
    for blk in pg.get_text("dict")["blocks"]:
        for ln in blk.get("lines", []):
            for sp in ln["spans"]:
                if sp["text"].strip() and y0 < sp["origin"][1] < y1:
                    out.append((sp["bbox"], sp["origin"][1], sp["size"],
                                sp["text"].strip()))
    return out

def collides(a, b):
    ba, bb = a[0], b[0]
    ox = max(0, min(ba[2], bb[2]) - max(ba[0], bb[0]))
    narrow = min(ba[2] - ba[0], bb[2] - bb[0])
    if narrow <= 0 or ox / narrow < 0.20:
        return False
    # 0.35 x font size = 'same visual line'. Anything further apart is a
    # separate row: a 30pt figure sitting 17pt above a 7.6pt sub-line is
    # 11pt of clear air, not a collision.
    return abs(a[1] - b[1]) < 0.35 * max(a[2], b[2])

elapsed = sum(1 for d in range(1, plan['as_of'].day + 1) if d not in plan['dry_days'])
pmno = PD.MONTH_NUM[plan['prior_month']]
net, byd = {}, {}
for r in h:
    d = _dt.date.fromisoformat(str(r['date'])[:10])
    if d.month != pmno: continue
    net[d.day] = net.get(d.day, 0) + float(r['tertiary_cs']) + float(r['invoice_cs'])
    byd.setdefault(d.day, {}); c = int(r['shop_code'])
    byd[d.day][c] = byd[d.day].get(c, 0) + float(r['tertiary_cs'])
live = sorted(d for d, v in net.items() if v >= 1.0)[:elapsed]
indep = {}
for d in live:
    for c, q in byd[d].items():
        indep[c] = indep.get(c, 0) + q

BOTTOM = 595.276 - 22           # nothing may print below the page margin
                                # (26pt design margin, 4pt tolerance)
bad, rows, hits, lowest = [], 0, 0, 0
for cl, bonds in PE.CLUSTERS.items():
    d = fitz.open(f'DAILY PACE - CLUSTER {cl}.pdf')
    if d.page_count != 1 + len(bonds): bad.append(f'C{cl} pages')
    for i, b in enumerate(sorted(bonds, key=lambda x: -plan['bonds'][x]['pct_done']), 1):
        pg = d[i]; t = pg.get_text(); e = plan['bonds'][b]
        for w in ('Total liquidation', 'cumulative cases', 'CASES SHORT',
                  'DAYS BEHIND', 'split by channel', 'invoice'):
            if w in t: bad.append(f'{b} stale "{w}"')
        for v in (e['per_day'], e['run_rate']):
            if f'{v:,.1f}' not in t: bad.append(f'{b} strip {v:,.1f}')
        if f"your normal day is {e['per_day']:,.1f} cs" not in t:
            bad.append(f'{b} normal-day line')
        if e.get('target_met'):
            # the card shows the surplus, not a needed-per-day of zero
            if f"+{e['surplus']:,.0f}" not in t: bad.append(f'{b} surplus')
            if 'TARGET ALREADY MET' not in t: bad.append(f'{b} met banner')
        elif f"{e['needed_per_day']:,.1f}" not in t:
            bad.append(f'{b} needed/day')
        proj = e['mtd_actual'] + e['run_rate'] * e['days_left']
        if e['run_rate'] < e['per_day'] and not e.get('target_met'):
            if f'you finish {max(e["target_month"] - proj, 0):,.0f} cs short' not in t:
                bad.append(f'{b} projection')

        pace = (e['mtd_actual'] / e['mtd_expected']) if e['mtd_expected'] else 0
        if f'{pace:.0%}' not in t: bad.append(f'{b} pace %')
        # the derived monthly shop total must NOT appear (Abhay, 27 Jul 2026)
        if f"of {e['target_month']:,.0f} cs sold" in t:
            bad.append(f'{b} monthly shop total still shown')
        if f"you should be at {e['mtd_expected']:,.0f} cs by today" not in t:
            bad.append(f'{b} should-be caption')
        for v in (e['mtd_actual'], e['mtd_expected'], abs(e['gap'])):
            if f'{v:,.0f}' not in t: bad.append(f'{b} month {v:,.0f}')
        if not e.get('target_met') and f"{e['run_rate']:,.1f} a day you have been doing" not in t:
            bad.append(f'{b} multiplier line')
        nz = [x for x in e['rows'] if not x['dry'] and x['actual'] is not None]
        pos = sorted(y['actual'] for y in nz if y['actual'] > 0.001)
        ref = max(pos[min(int(len(pos)*.85), len(pos)-1)] if pos else 0, e['per_day'])
        dpv = 0 if ref >= 10 else 1
        for x in nz:
            if f"{x['actual']:,.{dpv}f}" not in t:
                bad.append(f'{b} d{x["day"]} bar label')
        for r in sd.get(b, []):
            rows += 1
            if f"{r['now']:,.1f}" not in t: bad.append(f'{b} {r["shop"]}')
            if abs(r['last'] - indep.get(r['code'], 0)) > 0.01:
                bad.append(f'{b} {r["shop"]} last-mo')
        sp = spans(pg, 95, 205)
        for m in range(len(sp)):
            for n in range(m + 1, len(sp)):
                if collides(sp[m], sp[n]):
                    hits += 1
                    bad.append(f'{b} COLLIDE {sp[m][3][:18]!r}/{sp[n][3][:18]!r}')
        blk = [x for x in pg.get_text('blocks') if x[4].strip()]
        low = max(x[3] for x in blk); lowest = max(lowest, low)
        if low > BOTTOM: bad.append(f'{b} past bottom margin ({low:.1f})')
    d.close()
print(f'{rows} shop rows · prior window {matched}/{elapsed} days · '
      f'baseline collisions {hits} · lowest text {lowest:.1f} (margin {BOTTOM:.1f})')
print('VERIFY:', 'ALL CLEAN' if not bad else bad[:6])
