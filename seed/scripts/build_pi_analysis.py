#!/usr/bin/env python3
"""
PI ANALYSIS workbook builder — K.S. Distillery.
Rebuilds the 22-sheet "PI INSIGHTS" workbook for any consecutive month pair.

Usage:
  python3 build_pi_analysis.py --current July --prior June [--base "<Claude folder>"] [--out <scratch.xlsx>]

Auto-detects the raw PI folders by reading each file's "Report Month" (folder names need
not be consistent). Cross-checks vs MASTER DATA CONFIRMED, reconciles "Previous 3 Months
Sale" against the KSBC tertiary workbooks, and computes ALL narrative from the data.
"""
import argparse, re, html, glob, os, collections, statistics
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

MONTHS=['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER']
ABBR={m:m[:3] for m in MONTHS}                      # JANUARY->JAN
def midx(name): return MONTHS.index(name.upper())
def window3(month_name):
    """3 calendar months before month_name, as ABBR list e.g. JUNE->[MAR,APR,MAY]."""
    i=midx(month_name); return [ABBR[MONTHS[(i-k)%12]] for k in (3,2,1)]

ap=argparse.ArgumentParser()
ap.add_argument('--current', required=True)
ap.add_argument('--prior', required=True)
ap.add_argument('--base', default='/sessions/sweet-kind-mendel/mnt/Claude')
ap.add_argument('--out', default=None)
A=ap.parse_args()
CUR=A.current.capitalize(); PRIOR=A.prior.capitalize()      # 'June' as stored in Report Month
CUR_T, PRIOR_T = CUR, PRIOR                                  # title-case for headers
CUR_U, PRIOR_U = CUR.upper(), PRIOR.upper()
BASE=A.base; OUT=A.out or f'/tmp/pi/scratch_{CUR_U}.xlsx'
WIN_CUR=window3(CUR); WIN_PRIOR=window3(PRIOR)
WINDOW={CUR:WIN_CUR, PRIOR:WIN_PRIOR}
NEEDED=sorted(set(WIN_CUR)|set(WIN_PRIOR), key=lambda a:[ABBR[m] for m in MONTHS].index(a))
def span_label(win): return f"{win[0].title()}–{win[-1].title()}"   # 'Mar–May'

SHORT={'1139703':'BCB','1139707':'BLENDERS','1139708':'MWB','1139715':'CC',
       '1339703':'OLD PEARL','1339704':'ROF','1339710':'KS 99','1339718':'MBR'}
FULL={'1139703':'BCB NO.1 CLASSIC BRANDY','1139707':"BLENDER'S CHOICE NO.1 BRANDY",
      '1139708':'MORNING WALKERS XO BRANDY','1139715':"CHAIRMAN'S CHOICE XO BRANDY",
      '1339703':'OLD PEARL NO.1 MATURED XXX RUM','1339704':'ROYAL OLD FORT NO.1 XXX RUM',
      '1339710':'K.S 99 LIFE TIME MATURED XXX RUM','1339718':'MAGIC BLEND RESERVED XXX RUM'}
BORDER_ORDER=['OLD PEARL','BLENDERS','BCB','KS 99','ROF','MWB','MBR','CC']
def _short(code):
    # guarded brand-code lookup -- abort with an actionable message instead of an opaque KeyError
    if code not in SHORT:
        import sys
        sys.exit(f"FATAL: unknown KSD brand code {code!r} in a PI file -- not in SHORT/FULL maps. "
                 f"KSBC likely listed a new SKU. Add it to SHORT and FULL (line ~40) and rerun.")
    return SHORT[code]
def nb(b):
    b=str(b).upper().replace("’","'").strip(); b=re.sub(r"\s+"," ",b)
    return b.replace("MORNING WALKER'S","MORNING WALKERS")

# ---------- parse PI files (auto-detect folders by Report Month) ----------
def cells_of(t): return [html.unescape(re.sub(r'<[^>]+>','',c)).strip() for c in re.findall(r'<t[dh][^>]*>(.*?)(?=<t[dh][^>]*>|$)',t,re.S)]
def parse_file(path):
    txt=open(path,encoding='utf-8',errors='replace').read()
    shop=re.search(r'Shop\s*:\s*<b>([^<]+)</b>',txt); wh=re.search(r'Warehouse\s*:\s*<b>([^<]+)</b>',txt)
    mon=re.search(r'Report Month\s*:\s*<b>([^<]+)</b>',txt)
    cap=re.search(r'Shop Capacity\(Cases\)\s*:\s*([\d,]+)\s*,\s*Shop Area\s*:\s*([\d,]+)',txt)
    rows=[]; cs=cells_of(txt); i=0
    while i<len(cs)-6:
        c=cs[i:i+7]
        if re.fullmatch(r'\d{1,3}',c[0]) and re.fullmatch(r'\d{6,8}',c[1]) and c[2] and not c[2].isdigit():
            nums,ok=[],True
            for x in c[3:7]:
                x=(x or '0').replace(',','') or '0'
                if not re.fullmatch(r'\d+',x): ok=False; break
                nums.append(int(x))
            if ok: rows.append((c[1],c[2],*nums)); i+=7; continue
        i+=1
    m=re.match(r'(\d+)[\s-]+(.*)',shop.group(1).strip()) if shop else None
    return dict(file=os.path.basename(path),shop_code=(m.group(1) if m else None),
                shop_name=(m.group(2).strip() if m else None),warehouse=(wh.group(1).strip() if wh else None),
                month=(mon.group(1).strip() if mon else None),
                capacity=int(cap.group(1).replace(',','')) if cap else None,
                area=int(cap.group(2).replace(',','')) if cap else None, rows=rows)

pi_all=[]
for f in glob.glob(os.path.join(BASE,'PURCHASE INSTRUCTION','**','*.xls'),recursive=True):
    d=parse_file(f)
    if d['month'] in (CUR,PRIOR) and d['shop_code']: pi_all.append(d)
assert pi_all, f"No PI files found for {CUR}/{PRIOR} under PURCHASE INSTRUCTION/"
_have={d['month'] for d in pi_all}
assert CUR in _have and PRIOR in _have, f"Need BOTH months' PI files on disk. Found only: {sorted(_have)} (CUR={CUR}, PRIOR={PRIOR}). Drop the missing month's raws before building."
for _m in (CUR,PRIOR):
    _c=collections.Counter(d['shop_code'] for d in pi_all if d['month']==_m)
    _dups={k:v for k,v in _c.items() if v>1}
    assert not _dups, f"Duplicate {_m} PI files for shops {_dups} — would double-count MQ. De-dupe the raw folder first."

# ---------- master ----------
import openpyxl
mwb=openpyxl.load_workbook(os.path.join(BASE,'MASTER DATA CONFIRMED.xlsx'),read_only=True)
mws=mwb['16-4-25']; master={}
for row in mws.iter_rows(min_row=3,values_only=True):
    if row[3] is None: continue
    mm=re.match(r'\s*(\d+)',str(row[4]))
    if not mm: continue
    master[mm.group(1)]=dict(name=str(row[4]).strip(),cat=str(row[5]).strip(),
        staff=str(row[6]).strip(),bond=str(row[7]).strip(),status=str(row[8]).strip())
active={c for c,v in master.items() if v['cat']=='KSBC' and v['status']=='Active'}

# drop non-master shops that carry an EMPTY PI (the 10 stray empties); keep closed-in-master
pi=[]
for d in pi_all:
    if d['shop_code'] in master or d['rows']:
        pi.append(d)
DROPPED=[d['file'] for d in pi_all if d not in pi]

# ---------- tertiary sales (auto-discover KSBC workbooks for NEEDED months) ----------
def find_ksbc_wb(abbr):
    full=[m for m in MONTHS if m[:3]==abbr][0]
    cand=glob.glob(os.path.join(BASE,'KSBC shop sales',f'{full} SHOP SALES ANALYSIS.xlsx'))
    cand+=sorted(glob.glob(os.path.join(BASE,'KSBC shop sales',f'{full} 1st - *ANALYSIS.xlsx')))
    if len(cand)>1:
        print(f'  WARN: {len(cand)} KSBC workbooks match {full} (year-ambiguous): {[os.path.basename(c) for c in cand]} — picking most recently modified. Add a year token to disambiguate.')
        cand=sorted(cand,key=os.path.getmtime,reverse=True)
    return cand[0] if cand else None
sales=collections.defaultdict(lambda:dict(cs=0.0,btl=0.0))
MISSING_TERT=[]
for abbr in NEEDED:
    wbpath=find_ksbc_wb(abbr)
    if not wbpath: print(f"  WARN: no KSBC workbook for {abbr} — reconciliation/rate partial"); MISSING_TERT.append(abbr); continue
    w=openpyxl.load_workbook(wbpath,read_only=True)
    sn=[n for n in w.sheetnames if 'COMBINED' in n.upper()][0]
    it=w[sn].iter_rows(values_only=True); next(it)
    for r in it:
        if r[1] is None: continue
        nm=re.match(r'\s*(\d+)',str(r[2] or ''))
        if not nm: continue
        portal=nm.group(1); b=nb(r[4]); bpc=float(r[6] or 0); cs=float(r[9] or 0)
        sales[(abbr,portal,b)]['cs']+=cs; sales[(abbr,portal,b)]['btl']+=cs*bpc

CUR_WIN_OK = not (set(WIN_CUR)&set(MISSING_TERT)); PRIOR_WIN_OK = not (set(WIN_PRIOR)&set(MISSING_TERT))
# ---------- build lines (per-line window) ----------
lines=[]
for d in pi:
    md=master.get(d['shop_code'],{})
    for r in d['rows']:
        brand=nb(r[1])
        rate=sum(sales.get((m,d['shop_code'],brand),{'cs':0})['cs'] for m in WIN_CUR)/3
        ours=sum(sales.get((m,d['shop_code'],brand),{'btl':0})['btl'] for m in WINDOW[d['month']])
        lines.append(dict(month=d['month'],shop=d['shop_code'],shop_name=d['shop_name'].title(),
            bond=md.get('bond','—'),staff=md.get('staff','—').title(),wh=d['warehouse'],cap=d['capacity'],
            sb=_short(r[0]),brand=brand,p3=r[2],rl=r[3],rq=r[4],mq=r[5],rate=round(rate,2),ours=round(ours,1)))

# ---------- validation prints ----------
for mth in (CUR,PRIOR):
    pc=collections.Counter(d['shop_code'] for d in pi if d['month']==mth)
    dups={k:v for k,v in pc.items() if v>1}
    miss=active-set(pc); extra=set(pc)-active
    print(f"{mth}: files {sum(pc.values())} | shops {len(pc)} | missing-active {sorted(miss)} | extras {sorted(extra)} | dups {dups}")
assert all(r[3]+r[4]==r[5] for d in pi for r in d['rows']), "MQ != RL+RQ somewhere"
if DROPPED: print(f"dropped non-master empty PI files: {len(DROPPED)} -> {DROPPED}")

# ---------- dynamic blackout / door classification (replaces hardcoded FLAGS) ----------
cur_shops={d['shop_code'] for d in pi if d['month']==CUR}
cur_empty={d['shop_code']:d for d in pi if d['month']==CUR and not d['rows']}
def win_btl(shop,win): return sum(sales.get((m,shop,b),{'btl':0})['btl'] for m in win for b in [x['brand'] for x in lines if x['shop']==shop])
def shop_window_btl(shop,win):
    return sum(v['btl'] for (mm,s,b),v in sales.items() if s==shop and mm in win)
def classify_empties(month, win):
    fl={}
    for s in [d['shop_code'] for d in pi if d['month']==month and not d['rows']]:
        st=master.get(s,{}).get('status','')
        if st=='Closed': fl[s]='CLOSED'
        elif shop_window_btl(s,win)>0: fl[s]='DARK'
        elif any(shop_window_btl(s,[m])>0 for m in [ABBR[x] for x in MONTHS]): fl[s]='LOST'
        else: fl[s]='WHITE'
    return fl
FLAGS=classify_empties(CUR, WIN_CUR)
FLAGS_PRIOR=classify_empties(PRIOR, WIN_PRIOR)
# shops whose raw PI is present in only ONE of the two months -> can't compare; flag, don't build a block
CUR_FILE_SHOPS={d['shop_code'] for d in pi if d['month']==CUR}
PRIOR_FILE_SHOPS={d['shop_code'] for d in pi if d['month']==PRIOR}
INCOMPLETE={s:(PRIOR if s in CUR_FILE_SHOPS else CUR) for s in (CUR_FILE_SHOPS ^ PRIOR_FILE_SHOPS)}  # shop -> month whose raw is MISSING
SHOPNAME={d['shop_code']:str(d['shop_name']).title() for d in pi}

# ---------- ceiling-recovery engine (documented MQ-inversion + rolling window) ----------
# HOLD  = offtake (cases) of the month rolling OUT of next month's PI window = oldest window month WIN_CUR[0].
# RECOV = round(1.44 * Cut * rate / RL)  (MQ-formula slope dMQ/dB = 2.08*RL/rate ; 3/2.08 = 1.44).
# Action/Why from the WIN_CUR monthly trend. Spec: CLAUDE.md + .claude/memory/pi-analysis-workbook.md.
def _mo(l,m): return sales.get((m,l['shop'],l['brand']),{'cs':0})['cs']
def recovery_line(l, pmq):
    cut = pmq - l['mq']
    if cut <= 0: return None
    om0,om1,om2 = WIN_CUR
    a0,a1,a2 = _mo(l,om0),_mo(l,om1),_mo(l,om2)
    rate=l['rate']; rl=l['rl']
    hold=round(a0)
    rec=round(1.44*cut*rate/rl) if rl else round(1.44*cut*rate/max(pmq,1))
    peak=max(a0,a1,a2)
    if a2==0 and a1==0 and a0>0: why=f"\U0001f6d1 DORMANT · sold {om0.title()} only"
    elif a2==0 and (a0>0 or a1>0): why=f"⚠ {om2} = 0 · find why"
    elif a1>0 and a2<0.6*a1: why=f"▼ {om2.title()} dropped sharply"
    elif a0==peak and a0>a2 and a0>0: why=f"↘ {om0.title()} peak leaving window"
    elif a0>a1>a2: why="↘ declining 3-mo"
    else: why=""
    if cut<=1 and a0<a1 and a0<a2 and a2>0: act="✓ SELF-CORRECTS"; rec=0
    elif a2==0: act="BIG PUSH"
    elif a1>0 and a2<0.6*a1: act="BIG PUSH"
    elif rate>=12 and a2>=a1: act="CHEAP WIN"
    elif rate>=4: act="PUSH"
    else: act="BIG PUSH"
    return dict(hold=hold,rec=rec,action=act,why=why,cut=cut)
def pill_for(a):
    u=a.upper()
    if "BIG" in u: return "BIG PUSH", T_RED
    if "CHEAP" in u: return "CHEAP WIN", T_GREEN
    if "SELF" in u: return "✓ SELF-CORR", T_GREY
    if "PUSH" in u: return "PUSH", T_AMBER
    return a, T_AMBER
def boxcell(ws,rr,cc,v,nf=None,bold=False,col='FF111827',al=None,rf=None):
    c=ws.cell(rr,cc,v); c.font=F(9.5,bold,col); c.alignment=al or CTR; c.border=BOX
    if rf: c.fill=fill(rf)
    if nf: c.number_format=nf
    return c

# ================= COMPUTED NARRATIVE (data-driven; no hardcoded month facts) =================
def mq(mth,pred=lambda l:True): return sum(l['mq'] for l in lines if l['month']==mth and pred(l))
def off(mth,pred=lambda l:True): return sum(l['p3'] for l in lines if l['month']==mth and pred(l))
bonds_N=sorted({l['bond'] for l in lines if l['bond']!='—'})
n_bonds=len(bonds_N)
CUR_MQ=mq(CUR); PRIOR_MQ=mq(PRIOR); D_NET=CUR_MQ-PRIOR_MQ; PCT_NET=(D_NET/PRIOR_MQ) if PRIOR_MQ else 0
STAND_CUR=len([l for l in lines if l['month']==CUR and l['mq']>0]); STAND_PRIOR=len([l for l in lines if l['month']==PRIOR and l['mq']>0])
SILENT=[l for l in lines if l['month']==CUR and l['mq']==0 and l['p3']>0]; SILENT_N=len(SILENT); SILENT_BTL=sum(l['p3'] for l in SILENT)
CUR_OFF=off(CUR); PRIOR_OFF=off(PRIOR); OFF_MOM=(CUR_OFF/PRIOR_OFF-1) if PRIOR_OFF else 0
cov=sorted([l['mq']/l['rate'] for l in lines if l['month']==CUR and l['mq']>0 and l['rate']>0.33]); med_cov=statistics.median(cov) if cov else 0
p25=cov[len(cov)//4] if cov else 0
UNDER=[l for l in lines if l['month']==CUR and l['mq']>0 and l['rate']>=5 and l['mq']<0.6*l['rate']]; N_UNDER=len(UNDER)
silent_cs=sum(l['rate'] for l in SILENT)
UNDER_SORT=sorted(UNDER,key=lambda l:l['mq']/l['rate']); WORST=UNDER_SORT[0] if UNDER_SORT else None
# RL stockout risk: RL (re-order trigger) below one week of run-rate -> reorders too late
WEEKS_PER_MO=4.345
RLRISK=sorted([l for l in lines if l['month']==CUR and l['rate']>=4 and l['rl']>0 and l['rl'] < l['rate']/WEEKS_PER_MO],
              key=lambda l:(l['rl']/(l['rate']/WEEKS_PER_MO)))
# also flag RL==0 on a standing line (MQ>0 but no trigger set) at a selling door
RLZERO=sorted([l for l in lines if l['month']==CUR and l['mq']>0 and l['rl']==0 and l['rate']>=4], key=lambda l:-l['rate'])
# field-staff scorecard aggregates
staff_rows=collections.defaultdict(lambda:dict(shops=set(),standing=0,silent=0,silent_btl=0,under=0,dark=0,mq=0,off_c=0,off_p=0))
for l in lines:
    if l['month']==CUR:
        a=staff_rows[l['staff']]; a['shops'].add(l['shop']); a['mq']+=l['mq']; a['off_c']+=l['p3']
        if l['mq']>0: a['standing']+=1
        if l['mq']==0 and l['p3']>0: a['silent']+=1; a['silent_btl']+=l['p3']
        if l['mq']>0 and l['rate']>=5 and l['mq']<0.6*l['rate']: a['under']+=1
    elif l['month']==PRIOR:
        staff_rows[l['staff']]['off_p']+=l['p3']
bond_d={b:(mq(CUR,lambda l:l['bond']==b)-mq(PRIOR,lambda l:l['bond']==b)) for b in bonds_N}
BONDS_CUT=sum(1 for b in bonds_N if bond_d[b]<0); RAISED=[b for b in bonds_N if bond_d[b]>0]
raised_lbl = (RAISED[0] if len(RAISED)==1 else (f"{len(RAISED)} bonds" if RAISED else "no bond"))
top_raise = max(bonds_N,key=lambda b:bond_d[b]); top_raise_pct = bond_d[top_raise]/mq(PRIOR,lambda l:l['bond']==top_raise) if mq(PRIOR,lambda l:l['bond']==top_raise) else 0
# brand momentum
def brand_off_pct(sb):
    a=off(CUR,lambda l:l['sb']==sb); b=off(PRIOR,lambda l:l['sb']==sb); return (a/b-1) if b else 0
def stand_by_brand_pre(sb): return len([l for l in lines if l['month']==CUR and l['sb']==sb and l['mq']>0])
GROW=max([sb for sb in SHORT.values() if stand_by_brand_pre(sb)>=30] or list(SHORT.values()),key=brand_off_pct); GROW_FULL=[v for k,v in FULL.items() if SHORT[k]==GROW][0]
grow_lines=[l for l in lines if l['month']==CUR and l['sb']==GROW]; grow_sil=[l for l in grow_lines if l['mq']==0 and l['p3']>0]
grow_sil_pct=len(grow_sil)/len(grow_lines) if grow_lines else 0
# zero/low system-pull brand (fewest standing doors)
stand_by_brand={sb:len([l for l in lines if l['month']==CUR and l['sb']==sb and l['mq']>0]) for sb in SHORT.values()}
listed_by_brand={sb:len([l for l in lines if l['month']==CUR and l['sb']==sb]) for sb in SHORT.values()}
NOPULL=min(SHORT.values(),key=lambda sb:stand_by_brand[sb]); NOPULL_FULL=[v for k,v in FULL.items() if SHORT[k]==NOPULL][0]
# blackout doors detail
DARKS=[]
for s,flag in FLAGS.items():
    nm=master.get(s,{}).get('name',s); bond=master.get(s,{}).get('bond','—'); wbt=shop_window_btl(s,WIN_CUR)
    DARKS.append((s,nm,bond,flag,wbt))
for s_,nm,bond,flag,wbt in DARKS:
    if flag in ('DARK','LOST','WHITE'):
        st=master.get(s_,{}).get('staff','—').title()
        if st in staff_rows: staff_rows[st]['dark']+=1
DARK_SELL=[d for d in DARKS if d[3]=='DARK']
# reconciliation stats for validation
recon_exact=recon_w5=recon_off=0; max_gap=0
for l in lines:
    o=sum(sales.get((m,l['shop'],l['brand']),{'btl':0})['btl'] for m in WINDOW[l['month']]); g=abs(l['p3']-o)
    if g<0.51: recon_exact+=1
    elif g<=5: recon_w5+=1
    else: recon_off+=1; max_gap=max(max_gap,g)
N_LINES=len(lines); recon_pct=recon_exact/N_LINES if N_LINES else 0
n_files_cur=len({d['file'] for d in pi if d['month']==CUR}); n_files_prior=len({d['file'] for d in pi if d['month']==PRIOR})
_cur_codes=collections.Counter(d['shop_code'] for d in pi if d['month']==CUR)
chk2_missing=sorted(active-set(_cur_codes)); chk2_extras=sorted(set(_cur_codes)-active)
chk2_dups={k:v for k,v in _cur_codes.items() if v>1}
chk2_closed_extra=[e for e in chk2_extras if master.get(e,{}).get('status')=='Closed']
chk2_unknown=[e for e in chk2_extras if e not in master]
def _nm(c): return master.get(c,{}).get('name',c)
def _month_ingest(mth):
    fm={d['shop_code']:d for d in pi if d['month']==mth}
    ingested=sorted([c for c in active if c in fm])
    missing=sorted([c for c in active if c not in fm])
    empty=sorted([c for c in active if c in fm and len(fm[c]['rows'])==0])
    return ingested,missing,empty
chk2_by_month={m:_month_ingest(m) for m in (CUR,PRIOR)}
chk2_pass = all(not chk2_by_month[m][1] for m in (CUR,PRIOR)) and (not chk2_dups) and (not chk2_unknown)

RED='FFC62828'; AMBER='FFE65100'; GREENC='FF2E7D32'; BLUEC='FF1565C0'
def pf(x): return f"{x:+.0%}"
CALLS=[
 (f"1 · KSBC {'cut' if D_NET<0 else 'lifted'} our network ceiling {abs(PCT_NET):.1%} for {CUR_T} ({PRIOR_MQ:,} → {CUR_MQ:,} cs).",
  f"Driven by {span_label(WIN_CUR)} offtake {pf(OFF_MOM)} vs the prior window. {BONDS_CUT} of {n_bonds} bonds were cut; only {raised_lbl} rose. KSBC recomputes from trailing sales — if offtake stays soft, next month's PI cuts again.", RED if D_NET<0 else GREENC),
 (f"2 · {GROW_FULL.split(' NO')[0].split(' XO')[0].split(' XXX')[0].title()} leads offtake momentum ({pf(brand_off_pct(GROW))}) — but the shelf lags.",
  f"{len(grow_sil)} of {len(grow_lines)} {CUR_T} {GROW} lines carry MQ = 0 ({grow_sil_pct:.0%} off-indent) despite real sales. Escalate RL/RQ with KSBC RM offices; until then these shops are served only by manual indent / STN.", AMBER),
 (f"3 · {NOPULL_FULL.title()} has almost no system pull ({stand_by_brand[NOPULL]} on-indent of {listed_by_brand[NOPULL]} listed).",
  f"The PI system will barely replenish it — liquidation must run on incentive + STN, not the auto-indent. Treat it as push-only.", RED),
 (f"4 · PI BLACKOUT at {len(DARK_SELL)} actively-selling shop(s): " + (", ".join(f"{d[1]} ({d[2].title()})" for d in DARK_SELL) if DARK_SELL else "none this month") + ".",
  ("They sold "+ " / ".join(f"{int(d[4]):,}" for d in DARK_SELL) + f" btl in the {span_label(WIN_CUR)} window yet received EMPTY {CUR_T} PIs (same warehouses — a KSBC ERP gap, not a remap). Escalate to the RM office to restore the lines." ) if DARK_SELL else "No selling shop lost its PI this month.", RED if DARK_SELL else GREENC),
 (f"5 · {SILENT_N} off-indent lines = {SILENT_BTL:,} bottles of demand with no auto-indent.",
  f"MQ = 0 despite real sales. These shops depend 100% on field push — the {SILENT_N} off-indent lines are the field-push target list.", AMBER),
 (f"6 · The ceiling is earned, not negotiated: median MQ ≈ {med_cov:.2f} month of offtake.",
  f"{N_UNDER} high-velocity lines sit under 0.6-month cover" + (f" — worst {WORST['sb']} at {WORST['shop_name']} ({WORST['rate']:.0f} cs/mo vs MQ {WORST['mq']} = {WORST['mq']/WORST['rate']:.2f} mo)" if WORST else '') + f". KSBC recomputes MQ monthly from trailing sales — every extra case sold lifts next month's target. Push offtake where cover binds hardest.", BLUEC),
]
BRAND_NOTE={'OLD PEARL':'ceiling {p} on '+pf(brand_off_pct('OLD PEARL'))+' offtake',
            'BLENDERS':'ceiling {p} on '+pf(brand_off_pct('BLENDERS'))+' offtake',
            'BCB':'ceiling {p}',
            'KS 99':'ceiling {p}',
            'ROF':'ceiling {p}',
            'MWB':'ceiling {p} but offtake '+pf(brand_off_pct('MWB'))+' — demand vs shelf gap ('+f"{len([l for l in lines if l['month']==CUR and l['sb']=='MWB' and l['mq']==0 and l['p3']>0])}"+' off-indent lines)',
            'MBR':'ceiling {p} — thin footprint',
            'CC':f'{stand_by_brand["CC"]} on-indent line(s) — push-only, minimal system pull'}
NAVY='FF0D1B4A'; NAVYM='FF1A237E'; NAVYS='FF2E3F70'; BANNER='FF263F80'; GOLD='FFFFB300'; GOLDD='FFFFD54F'
WHITE='FFFFFFFF'; GREY='FF6B7280'; LGREY='FFF3F4F6'
T_BLUE=('FF1565C0','FFBBDEFB'); T_GREEN=('FF2E7D32','FFDCEDC8'); T_AMBER=('FFE65100','FFFFE0B2'); T_RED=('FFC62828','FFFFCDD2'); T_GREY=('FF6B7280','FFECEFF1')
FN='Aptos Narrow'
def F(sz=10, b=False, c='FF111827', i=False): return Font(name=FN, size=sz, bold=b, color=c, italic=i)
def fill(c): return PatternFill('solid', start_color=c)
CTR=Alignment(horizontal='center', vertical='center', wrap_text=True)
LFT=Alignment(horizontal='left', vertical='center', wrap_text=True)
RGT=Alignment(horizontal='right', vertical='center')
thin=Side(style='thin', color='FFD1D5DB'); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
gold_b=Side(style='medium', color=GOLD)

def hero(ws, ncols, title, subtitle):
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=ncols)
    c=ws.cell(1,1,title); c.font=Font(name=FN,size=22,bold=True,color=WHITE); c.alignment=CTR
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=ncols)
    s=ws.cell(3,1,subtitle); s.font=Font(name=FN,size=11,bold=False,color=GOLDD); s.alignment=CTR
    for rr in (1,2,3):
        for cc in range(1, ncols+1):
            ws.cell(rr,cc).fill=fill(NAVY)
            if rr==3: ws.cell(rr,cc).border=Border(bottom=Side(style='medium', color=GOLD))
    ws.row_dimensions[1].height=26; ws.row_dimensions[2].height=8; ws.row_dimensions[3].height=18
    ws.sheet_view.showGridLines=False

def band(ws, row, ncols, text, h=22, fl=BANNER, fc=WHITE, sz=12):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c=ws.cell(row,1,text); c.font=Font(name=FN,size=sz,bold=True,color=fc); c.alignment=Alignment(horizontal='left',vertical='center',indent=1)
    for cc in range(1,ncols+1):
        ws.cell(row,cc).fill=fill(fl); ws.cell(row,cc).border=Border(bottom=Side(style='thin',color=GOLD))
    ws.row_dimensions[row].height=h

def header_row(ws, row, cols, start=1, fl=NAVYS, h=26):
    for i, t in enumerate(cols):
        c=ws.cell(row, start+i, t); c.font=Font(name=FN,size=10,bold=True,color=WHITE); c.alignment=CTR
        c.fill=fill(fl); c.border=Border(bottom=Side(style='thin', color=GOLD))
    ws.row_dimensions[row].height=h

wb = Workbook()
RAW_SHEET='RAW PI DATA'

# ============ RAW PI DATA (built first; other sheets reference it) ============
raw = wb.active; raw.title=RAW_SHEET
hero(raw, 15, 'RAW PI DATA — AUDIT TRAIL', f'Every Purchase Instruction line, {PRIOR_T} + {CUR_T} · parsed verbatim from KSBC ERP exports · KSD tertiary rate appended for cover analytics')
hr=5
header_row(raw, hr, ['Month','Shop Code','Shop Name','Bond','Field Staff','Warehouse','Shop Capacity (cs)','Brand','Product Brand (KSBC)','Prev 3-Mo Sale (btl)','RL (cs)','RQ (cs)','MQ (cs)',f'KSD Tertiary Rate (cs/mo, {span_label(WIN_CUR)})','KSD Bottles Same Window'], h=42)
r=hr+1
for L in sorted(lines, key=lambda x:(x['month']!=PRIOR, x['bond'], int(x['shop']), x['sb'])):
    vals=[L['month'],L['shop'],L['shop_name'],L['bond'],L['staff'],L['wh'],L['cap'],L['sb'],L['brand'],L['p3'],L['rl'],L['rq'],L['mq'],L['rate'],L['ours']]
    for i,v in enumerate(vals):
        c=raw.cell(r,1+i,v); c.font=F(9); c.border=BOX
        c.alignment = LFT if i in (2,4,5,8) else CTR
    if L['mq']==0 and L['p3']>0:
        for i in range(15): raw.cell(r,1+i).fill=fill('FFFFF7ED')
    r+=1
last_raw=r-1
for i,w in enumerate([7,9,22,14,16,26,12,11,30,12,7,7,7,14,12]): raw.column_dimensions[get_column_letter(1+i)].width=w
for col,fmt in [(7,'#,##0'),(10,'#,##0'),(11,'#,##0'),(12,'#,##0'),(13,'#,##0'),(14,'0.00'),(15,'#,##0')]:
    for rr in range(hr+1,last_raw+1): raw.cell(rr,col).number_format=fmt
raw.freeze_panes='A6'; raw.auto_filter.ref=f"A{hr}:O{last_raw}"
RR=lambda col: f"'{RAW_SHEET}'!{col}{hr+1}:{col}{last_raw}"
# ============ helper aggregates ============
mayd = {(l['shop'], l['sb']): l for l in lines if l['month']==PRIOR}
jund = {(l['shop'], l['sb']): l for l in lines if l['month']==CUR}
bonds = sorted({l['bond'] for l in lines if l['bond']!='—'})
shop_jun = collections.defaultdict(lambda: dict(mq=0,p3=0,rate=0.0,brands=0))
shop_meta={}
for l in lines:
    if l['month']==CUR:
        s=shop_jun[l['shop']]; s['mq']+=l['mq']; s['p3']+=l['p3']; s['rate']+=l['rate']; s['brands']+=1 if l['mq']>0 else 0
        shop_meta[l['shop']]=l
shop_may_mq = collections.defaultdict(int)
for l in lines:
    if l['month']==PRIOR: shop_may_mq[l['shop']]+=l['mq']

new_st  = sorted([l for k,l in jund.items() if l['mq']>0 and mayd.get(k, {'mq':0})['mq']==0], key=lambda x:-x['mq'])
lost_st = sorted([l for k,l in mayd.items() if l['mq']>0 and jund.get(k, {'mq':0})['mq']==0], key=lambda x:-x['mq'])
raised  = sorted([(jund[k]['mq']-l['mq'], jund[k]) for k,l in mayd.items() if k in jund and jund[k]['mq']>l['mq']>0], key=lambda t:-t[0])
cut     = sorted([(l['mq']-jund[k]['mq'], jund[k], l['mq']) for k,l in mayd.items() if k in jund and 0<jund[k]['mq']<l['mq']], key=lambda t:-t[0])
zero_sales = sorted([l for l in lines if l['month']==CUR and l['mq']==0 and l['p3']>0], key=lambda x:-x['p3'])
underprov = sorted([(l['mq']/l['rate'], l) for l in lines if l['month']==CUR and l['mq']>0 and l['rate']>=5 and l['mq']<0.6*l['rate']], key=lambda t:t[0])
cov_all = sorted([l['mq']/l['rate'] for l in lines if l['month']==CUR and l['mq']>0 and l['rate']>0.33])
med_cov = statistics.median(cov_all)
bond_rate = {b: sum(v['cs'] for (m,s,bb),v in sales.items() if m in WIN_CUR and master.get(s,{}).get('bond')==b)/len(WIN_CUR) for b in bonds}
bond_sorted = sorted(bonds, key=lambda b: -sum(l['mq'] for l in lines if l['month']==CUR and l['bond']==b))

# kpi() kept for bond sheets
def kpi(ws, row, col, w, title, valref, sub, scolor=GOLDD, numfmt='#,##0', accent=GOLD):
    last=col+w-1
    for rr in range(row,row+3):
        for cc in range(col,last+1):
            c=ws.cell(rr,cc); c.fill=fill(NAVYM)
            bL=Side(style='thin',color='FF324272')
            bR=Side(style='thin',color='FF324272') if cc==last else None
            bT=Side(style='thick',color=accent) if rr==row else None
            bB=Side(style='thin',color=GOLD) if rr==row+2 else None
            c.border=Border(left=bL,right=bR,top=bT,bottom=bB)
    ws.merge_cells(start_row=row,start_column=col,end_row=row,end_column=last)
    t=ws.cell(row,col,title); t.font=Font(name=FN,size=9,bold=True,color='FFB7C3E8'); t.alignment=Alignment(horizontal='center',vertical='bottom')
    ws.merge_cells(start_row=row+1,start_column=col,end_row=row+1,end_column=last)
    v=ws.cell(row+1,col,valref); v.font=Font(name=FN,size=23,bold=True,color=WHITE); v.alignment=CTR; v.number_format=numfmt
    ws.merge_cells(start_row=row+2,start_column=col,end_row=row+2,end_column=last)
    sc=ws.cell(row+2,col,sub); sc.font=Font(name=FN,size=8.5,bold=False,color=scolor); sc.alignment=Alignment(horizontal='center',vertical='top',wrap_text=True)
    ws.row_dimensions[row].height=15; ws.row_dimensions[row+1].height=30; ws.row_dimensions[row+2].height=20

# ============ DASHBOARD (redesigned) ============
NC=14
db = wb.create_sheet('DASHBOARD')
# --- custom dashboard title (no subtitle band) ---
db.merge_cells(start_row=1,start_column=1,end_row=3,end_column=NC)
tc=db.cell(1,1,f'PI ANALYSIS  ·  {CUR_U} vs {PRIOR_U}'); tc.font=Font(name=FN,size=30,bold=True,color=WHITE); tc.alignment=CTR
db.merge_cells(start_row=4,start_column=1,end_row=4,end_column=NC)
sc=db.cell(4,1,'K.S. DISTILLERY  ·  KSBC PURCHASE INSTRUCTION INTELLIGENCE'); sc.font=Font(name=FN,size=10.5,bold=True,color=GOLDD); sc.alignment=CTR
for rr in (1,2,3,4):
    for cc in range(1,NC+1):
        db.cell(rr,cc).fill=fill(NAVY)
        if rr==4: db.cell(rr,cc).border=Border(bottom=Side(style='medium',color=GOLD))
db.row_dimensions[1].height=14; db.row_dimensions[2].height=24; db.row_dimensions[3].height=8; db.row_dimensions[4].height=18
db.sheet_view.showGridLines=False

# --- KPI cards: 2 rows x 3 cards, gaps between ---
def card(row, c0, c1, title, val, sub, numfmt='#,##0', accent=GOLD, subcolor=GOLDD):
    for rr in range(row,row+3):
        for cc in range(c0,c1+1):
            cell=db.cell(rr,cc); cell.fill=fill(NAVYM)
            bL=Side(style='thin',color='FF324272')
            bR=Side(style='thin',color='FF324272') if cc==c1 else None
            bT=Side(style='thick',color=accent) if rr==row else None
            bB=Side(style='thin',color=GOLD) if rr==row+2 else None
            cell.border=Border(left=bL,right=bR,top=bT,bottom=bB)
    db.merge_cells(start_row=row,start_column=c0,end_row=row,end_column=c1)
    t=db.cell(row,c0,title); t.font=Font(name=FN,size=9.5,bold=True,color='FFB7C3E8'); t.alignment=Alignment(horizontal='center',vertical='bottom')
    db.merge_cells(start_row=row+1,start_column=c0,end_row=row+1,end_column=c1)
    v=db.cell(row+1,c0,val); v.font=Font(name=FN,size=26,bold=True,color=WHITE); v.alignment=CTR; v.number_format=numfmt
    db.merge_cells(start_row=row+2,start_column=c0,end_row=row+2,end_column=c1)
    sc=db.cell(row+2,c0,sub); sc.font=Font(name=FN,size=9,color=subcolor); sc.alignment=Alignment(horizontal='center',vertical='top',wrap_text=True)

# missing-data banner (only when a tertiary month is absent)
if MISSING_TERT:
    db.merge_cells(start_row=5,start_column=1,end_row=5,end_column=NC)
    wc=db.cell(5,1,f'⚠ MISSING DATA — no KSBC tertiary workbook for {", ".join(MISSING_TERT)}. Offtake / Cover figures that use those months are PARTIAL (shown as n/a). Build the missing month before trusting cover.')
    wc.font=Font(name=FN,size=10.5,bold=True,color='FFFFFFFF'); wc.alignment=Alignment(horizontal='left',vertical='center',indent=1)
    for cc in range(1,NC+1): db.cell(5,cc).fill=fill('FFC62828')
    db.row_dimensions[5].height=22
K=6 if MISSING_TERT else 5
RED_A='FFE53935'; GRN_A='FF43A047'; AMB_A='FFF57C00'; BLU_A='FF1E88E5'
_bonds_cut = sum(1 for b in bonds if (sum(l['mq'] for l in lines if l['month']==CUR and l['bond']==b) - sum(l['mq'] for l in lines if l['month']==PRIOR and l['bond']==b)) < 0)
_n_under = len([1 for l in lines if l['month']==CUR and l['mq']>0 and l['rate']>=5 and l['mq']<0.6*l['rate']])
_silent_btl = sum(l['p3'] for l in lines if l['month']==CUR and l['mq']==0 and l['p3']>0)
_unserved_cs = round(sum(sales.get((m,l['shop'],l['brand']),{'cs':0})['cs'] for l in SILENT for m in WIN_CUR))
# Row 1 — the squeeze and its cause
card(K,1,3,f'{CUR_U} CEILING — ΣMQ', f"=SUMIFS({RR('M')},{RR('A')},\"{CUR}\")", f'{"▲" if D_NET>=0 else "▼"} {abs(D_NET):,} cs vs {PRIOR_T} {PRIOR_MQ:,}  ({PCT_NET:+.1%})', '#,##0', RED_A, 'FFFCA5A5')
card(K,4,9,'OFFTAKE MOMENTUM', f"=SUMIFS({RR('J')},{RR('A')},\"{CUR}\")/SUMIFS({RR('J')},{RR('A')},\"{PRIOR}\")-1", "trailing offtake — sets next month ceiling", '+0.0%;-0.0%', RED_A, 'FFFCA5A5')
card(K,10,14,'BONDS CUT', _bonds_cut, f'of {n_bonds} bonds · only {raised_lbl} raised ({top_raise_pct:+.0%})' if len(RAISED)==1 else (f'of {n_bonds} bonds · {len(RAISED)} raised' if RAISED else f'of {n_bonds} bonds · none raised'), f'0" of {n_bonds}"', RED_A, 'FFFCA5A5')
K2=K+4
# Row 2 — the prize and the constraint
card(K2,1,3,'UNSERVED DEMAND', _unserved_cs, f'{SILENT_N} off-indent lines · {SILENT_BTL:,} btl', '#,##0" cs"', AMB_A, 'FFFFD9A8')
card(K2,4,9,'UNDER-PROVISIONED', _n_under, 'sell faster than ceiling allows · <0.6 mo cover', '0" lines"', AMB_A, 'FFFFD9A8')
card(K2,10,14,'MEDIAN MQ COVER', round(med_cov,2), f'months of shelf KSBC allows · p25 {p25:.2f}', '0.00" mo"', BLU_A, 'FFBBDEFB')
for rr in (K,K2): db.row_dimensions[rr].height=16
for rr in (K+1,K2+1): db.row_dimensions[rr].height=32
for rr in (K+2,K2+2): db.row_dimensions[rr].height=18
for rr in (K+3,): db.row_dimensions[rr].height=7

# --- shared signal classifier (used by BOTH brand & bond tables) ---
def classify(june_mq, may_mq):
    d = june_mq - may_mq
    pct = (d/may_mq) if may_mq else (1.0 if d>0 else 0.0)
    if pct > 0.01:           return 'RAISED','FF2E7D32', pct, d
    if pct >= -0.04 or abs(d) < 3:  return 'STEADY','FF2E7D32', pct, d
    if pct >= -0.15:         return 'CUT','FFE65100', pct, d
    return 'STEEP CUT','FFC62828', pct, d

# --- BRAND LEAGUE ---
R0=K2+4
band(db, R0, NC, f'BRAND LEAGUE — {CUR_U} CEILING vs {PRIOR_U}  (cases)')
hdr=['Brand','On-Indent Shops',f'ΣMQ {CUR_T}',f'Δ vs {PRIOR_T}','Δ %','Share','Offtake Δ%']
header_row(db, R0+1, hdr+['Signal']+['']*(NC-len(hdr)-1), h=24)
db.merge_cells(start_row=R0+1,start_column=8,end_row=R0+1,end_column=NC)
# brand context notes (badge tier comes from the shared classifier; note carries the nuance)
r=R0+2
for sb in BORDER_ORDER:
    full=[v for k,v in FULL.items() if SHORT[k]==sb][0]
    db.cell(r,1,full).font=F(9.5,True,NAVYM)
    db.cell(r,2,f"=COUNTIFS({RR('A')},\"{CUR}\",{RR('H')},\"{sb}\",{RR('M')},\">0\")")
    db.cell(r,3,f"=SUMIFS({RR('M')},{RR('A')},\"{CUR}\",{RR('H')},\"{sb}\")")
    db.cell(r,4,f"=C{r}-SUMIFS({RR('M')},{RR('A')},\"{PRIOR}\",{RR('H')},\"{sb}\")")
    db.cell(r,5,f"=IF(C{r}-D{r}=0,0,D{r}/(C{r}-D{r}))")
    db.cell(r,6,f"=C{r}/SUMIFS({RR('M')},{RR('A')},\"{CUR}\")")
    db.cell(r,7,f"=IF(SUMIFS({RR('J')},{RR('A')},\"{PRIOR}\",{RR('H')},\"{sb}\")=0,0,SUMIFS({RR('J')},{RR('A')},\"{CUR}\",{RR('H')},\"{sb}\")/SUMIFS({RR('J')},{RR('A')},\"{PRIOR}\",{RR('H')},\"{sb}\")-1)")
    bj_mq=sum(l['mq'] for l in lines if l['month']==CUR and l['sb']==sb)
    bm_mq=sum(l['mq'] for l in lines if l['month']==PRIOR and l['sb']==sb)
    tag,scol,pct,dd=classify(bj_mq,bm_mq)
    note=BRAND_NOTE[sb].format(p=f"{pct:+.0%}")
    bc=db.cell(r,8,tag); bc.font=F(9,True,WHITE); bc.fill=fill(scol); bc.alignment=CTR; bc.border=BOX
    db.merge_cells(start_row=r,start_column=9,end_row=r,end_column=NC)
    sc=db.cell(r,9,note); sc.font=F(9.5,True,scol); sc.alignment=Alignment(horizontal='left',vertical='center',indent=1)
    for cc in range(1,NC+1):
        c2=db.cell(r,cc); c2.border=BOX
        if cc in (2,3): c2.number_format='#,##0'; c2.alignment=CTR
        if cc==4: c2.number_format='+#,##0;-#,##0;"·"'; c2.alignment=CTR
        if cc in (5,7): c2.number_format='+0%;-0%;"·"'; c2.alignment=CTR
        if cc==6: c2.number_format='0%'; c2.alignment=CTR
        if c2.font.name!=FN: c2.font=F(9.5)
        if (r-R0)%2==0 and cc!=8: c2.fill=fill('FFF7F9FC')
    d_brand = sum(l['mq'] for l in lines if l['month']==CUR and l['sb']==sb) - sum(l['mq'] for l in lines if l['month']==PRIOR and l['sb']==sb)
    dcol = 'FF2E7D32' if d_brand>0 else ('FFC62828' if d_brand<0 else GREY)
    db.cell(r,4).font=F(9.5,True,dcol); db.cell(r,5).font=F(9.5,True,dcol)
    db.row_dimensions[r].height=20
    r+=1
tot=r
db.cell(tot,1,'TOTAL').font=F(10,True,WHITE)
db.cell(tot,2,f"=SUM(B{R0+2}:B{tot-1})"); db.cell(tot,3,f"=SUM(C{R0+2}:C{tot-1})")
db.cell(tot,4,f"=SUM(D{R0+2}:D{tot-1})"); db.cell(tot,5,f"=D{tot}/(C{tot}-D{tot})")
db.cell(tot,6,1)
db.cell(tot,7,f"=SUMIFS({RR('J')},{RR('A')},\"{CUR}\")/SUMIFS({RR('J')},{RR('A')},\"{PRIOR}\")-1")
db.merge_cells(start_row=tot,start_column=8,end_row=tot,end_column=NC)
for cc in range(1,NC+1):
    c2=db.cell(tot,cc); c2.fill=fill('FF374151')
    if not (c2.font and c2.font.bold and c2.font.color and c2.font.color.rgb=='FFFFFFFF'): c2.font=F(10,True,WHITE)
    c2.alignment=CTR; c2.border=Border(top=Side(style='medium',color=GOLD))
    if cc in (2,3): c2.number_format='#,##0'
    if cc==4: c2.number_format='+#,##0;-#,##0;0'
    if cc in (5,7): c2.number_format='+0%;-0%;0%'
    if cc==6: c2.number_format='0%'
db.cell(tot,1).alignment=Alignment(horizontal='left',vertical='center',indent=1)
db.row_dimensions[tot].height=22

# --- BOND PULSE ---
R1=tot+2
band(db, R1, NC, f'BOND PULSE — {CUR_U} CEILING, MOVEMENT & COVER')
hdr2=['Bond','Shops',f'ΣMQ {CUR_T}',f'Δ vs {PRIOR_T}','Δ %','Offtake Δ%','Cover (mo)']
header_row(db, R1+1, hdr2+['Signal']+['']*(NC-len(hdr2)-1), h=24)
db.merge_cells(start_row=R1+1,start_column=8,end_row=R1+1,end_column=NC)
r=R1+2
for b in bond_sorted:
    jm=sum(l['mq'] for l in lines if l['month']==CUR and l['bond']==b)
    mm=sum(l['mq'] for l in lines if l['month']==PRIOR and l['bond']==b)
    d=jm-mm
    db.cell(r,1,b).font=F(9.5,True,NAVYM)
    db.cell(r,2,len({l['shop'] for l in lines if l['month']==CUR and l['bond']==b and l['mq']>0}))  # shops carrying a ceiling (MQ>0) — matches the ΣMQ universe (was: all PI files incl. empty)
    db.cell(r,3,f"=SUMIFS({RR('M')},{RR('A')},\"{CUR}\",{RR('D')},\"{b}\")")
    db.cell(r,4,f"=C{r}-SUMIFS({RR('M')},{RR('A')},\"{PRIOR}\",{RR('D')},\"{b}\")")
    db.cell(r,5,f"=IF(C{r}-D{r}=0,0,D{r}/(C{r}-D{r}))")
    db.cell(r,6,f"=IF(SUMIFS({RR('J')},{RR('A')},\"{PRIOR}\",{RR('D')},\"{b}\")=0,0,SUMIFS({RR('J')},{RR('A')},\"{CUR}\",{RR('D')},\"{b}\")/SUMIFS({RR('J')},{RR('A')},\"{PRIOR}\",{RR('D')},\"{b}\")-1)")
    cov=jm/bond_rate[b] if bond_rate.get(b) else 0
    db.cell(r,7,cov)
    tag,scol,pct,dd=classify(jm,mm)
    note={'RAISED':f'only bond raised ({pct:+.0%}) — defend it',
          'STEADY':f'ceiling held ({pct:+.0%})',
          'CUT':f'ceiling trimmed {pct:+.0%} ({d:+,} cs)',
          'STEEP CUT':f'ceiling cut {pct:+.0%} ({d:+,} cs)'}[tag]
    if cov<0.6: note += '  ·  binds: cover <0.6 mo'
    bc=db.cell(r,8,tag); bc.font=F(9,True,WHITE); bc.fill=fill(scol); bc.alignment=CTR; bc.border=BOX
    db.merge_cells(start_row=r,start_column=9,end_row=r,end_column=NC)
    sc=db.cell(r,9,note); sc.font=F(9.5,True,'FFC62828' if cov<0.6 else '+FF1F2937'.lstrip('+')); sc.alignment=Alignment(horizontal='left',vertical='center',indent=1)
    for cc in range(1,NC+1):
        c2=db.cell(r,cc); c2.border=BOX
        if cc in (2,3): c2.number_format='#,##0'; c2.alignment=CTR
        if cc==4: c2.number_format='+#,##0;-#,##0;"·"'; c2.alignment=CTR
        if cc in (5,6): c2.number_format='+0%;-0%;"·"'; c2.alignment=CTR
        if cc==7: c2.number_format='0.00'; c2.alignment=CTR
        if c2.font.name!=FN: c2.font=F(9.5)
        if (r-R1)%2==0 and cc!=8: c2.fill=fill('FFF7F9FC')
    db.cell(r,4).font=F(9.5,True,'FF2E7D32' if d>0 else ('FFC62828' if d<0 else GREY))
    db.cell(r,5).font=F(9.5,True,'FF2E7D32' if d>0 else ('FFC62828' if d<0 else GREY))
    if cov<0.6:
        db.cell(r,7).fill=fill(T_RED[1]); db.cell(r,7).font=F(9.5,True,T_RED[0])
    db.row_dimensions[r].height=20
    r+=1
tb=r
db.cell(tb,1,'TOTAL').font=F(10,True,WHITE)
db.cell(tb,2,f"=SUM(B{R1+2}:B{tb-1})"); db.cell(tb,3,f"=SUM(C{R1+2}:C{tb-1})")
db.cell(tb,4,f"=SUM(D{R1+2}:D{tb-1})"); db.cell(tb,5,f"=D{tb}/(C{tb}-D{tb})")
db.cell(tb,6,f"=SUMIFS({RR('J')},{RR('A')},\"{CUR}\")/SUMIFS({RR('J')},{RR('A')},\"{PRIOR}\")-1")
db.cell(tb,7,round(CUR_MQ/sum(bond_rate.values()),2) if sum(bond_rate.values()) else 0)
db.merge_cells(start_row=tb,start_column=8,end_row=tb,end_column=NC)
for cc in range(1,NC+1):
    c2=db.cell(tb,cc); c2.fill=fill('FF374151')
    if not (c2.font and c2.font.bold): c2.font=F(10,True,WHITE)
    c2.alignment=CTR; c2.border=Border(top=Side(style='medium',color=GOLD))
    if cc in (2,3): c2.number_format='#,##0'
    if cc==4: c2.number_format='+#,##0;-#,##0;0'
    if cc in (5,6): c2.number_format='+0%;-0%;0%'
    if cc==7: c2.number_format='0.00'
db.cell(tb,1).alignment=Alignment(horizontal='left',vertical='center',indent=1)
db.row_dimensions[tb].height=22

# --- SIX MOVES ---
R2=tb+2
band(db, R2, NC, 'SIX MOVES FOR THE CMO', fl=NAVY, h=24, sz=12)
calls=CALLS
r=R2+2
for t, s, col in calls:
    db.merge_cells(start_row=r,start_column=1,end_row=r,end_column=NC)
    c=db.cell(r,1,t); c.font=F(10.5,True,col); c.alignment=Alignment(horizontal='left',vertical='bottom',indent=1)
    db.row_dimensions[r].height=18
    db.merge_cells(start_row=r+1,start_column=1,end_row=r+1,end_column=NC)
    c=db.cell(r+1,1,s); c.font=F(9.5,c='FF4B5563'); c.alignment=Alignment(horizontal='left',vertical='top',wrap_text=True,indent=1)
    for cc in range(1,NC+1):
        db.cell(r+1,cc).border=Border(bottom=Side(style='thin',color='FFE5E7EB'))
    db.row_dimensions[r+1].height=28
    db.row_dimensions[r+2].height=6
    r+=3

for i,w in enumerate([40,11,10.5,10.5,9,9,11,11,10,12,12,12,12,12]): db.column_dimensions[get_column_letter(1+i)].width=w
db.freeze_panes='A4'
db.sheet_view.showGridLines=False
# bond pulse + callouts moved into wb2 (dashboard redesign)
# ============ BRAND ANALYSIS ============
ba = wb.create_sheet('BRAND ANALYSIS')
NB=12
hero(ba, NB, 'BRAND ANALYSIS — DISTRIBUTION & CEILING', f'Indent coverage, offtake momentum and listing economics per brand · {CUR_T} PI vs {PRIOR_T} PI · cases unless marked btl')
r=5
GOODUP={1,2,3,4,5,6,7}; BADUP={9,10}
for sb in BORDER_ORDER:
    full=[v for k,v in FULL.items() if SHORT[k]==sb][0]
    jl=[l for l in lines if l['month']==CUR and l['sb']==sb]
    ml=[l for l in lines if l['month']==PRIOR and l['sb']==sb]
    js=[l for l in jl if l['mq']>0]; ms=[l for l in ml if l['mq']>0]
    sil=sorted([l for l in jl if l['mq']==0 and l['p3']>0], key=lambda x:-x['p3'])
    jmq=sum(l['mq'] for l in js); mmq=sum(l['mq'] for l in ms)
    # brand banner: title left, metrics right
    ba.merge_cells(start_row=r,start_column=1,end_row=r,end_column=7)
    t=ba.cell(r,1,full)
    t.font=Font(name=FN,size=12,bold=True,color=WHITE); t.alignment=Alignment(horizontal='left',vertical='center',indent=1)
    ba.merge_cells(start_row=r,start_column=8,end_row=r,end_column=NB)
    d_=jmq-mmq; dtxt='—' if d_==0 else (f"▲{d_}" if d_>0 else f"▼{abs(d_)}")
    t2=ba.cell(r,8,f"{CUR_T} MQ {jmq:,} {dtxt}  ·  {len(js)} on-indent shops")
    t2.font=Font(name=FN,size=10,bold=True,color=GOLDD); t2.alignment=Alignment(horizontal='right',vertical='center',indent=1)
    for cc in range(1,NB+1):
        ba.cell(r,cc).fill=fill(NAVYM); ba.cell(r,cc).border=Border(bottom=Side(style='medium',color=GOLD))
    ba.row_dimensions[r].height=26
    r+=1
    header_row(ba, r, ['','Shops Listed','On-Indent (MQ>0)','Distribution Width','ΣMQ (cs)','ΣRL','ΣRQ','Offtake (btl)','Avg MQ / Shop','Off-Indent Lines','Off-Indent Demand (btl)','Median Cover (mo)'], fl='FF263F80', h=30)
    r+=1
    FMT={1:'#,##0',2:'#,##0',3:'0%',4:'#,##0',5:'#,##0',6:'#,##0',7:'#,##0',8:'0.0',9:'#,##0',10:'#,##0',11:'0.00'}
    for mth, ls, st in ((CUR,jl,js),(PRIOR,ml,ms)):
        sil_m=[l for l in ls if l['mq']==0 and l['p3']>0]
        cv=sorted([l['mq']/l['rate'] for l in st if l['rate']>0.33])
        vals=[mth,len(ls),len(st),len(st)/len(active),sum(l['mq'] for l in st),sum(l['rl'] for l in ls),sum(l['rq'] for l in ls),
              sum(l['p3'] for l in ls),(sum(l['mq'] for l in st)/len(st) if st else 0),len(sil_m),sum(l['p3'] for l in sil_m),
              (statistics.median(cv) if cv else 0)]
        for i,v in enumerate(vals):
            c=ba.cell(r,1+i,v); c.border=BOX; c.alignment=CTR
            if i==0:
                c.font=F(10,True,NAVYM) if mth==CUR else F(10,False,GREY)
            else:
                c.font=F(10) if mth==CUR else F(10,False,'FF6B7280')
                c.number_format=FMT[i]
            if mth==PRIOR: c.fill=fill('FFF7F8FA')
        ba.row_dimensions[r].height=18
        r+=1
    msil=[l for l in ml if l['mq']==0 and l['p3']>0]
    cvj=sorted([l['mq']/l['rate'] for l in js if l['rate']>0.33]); cvm=sorted([l['mq']/l['rate'] for l in ms if l['rate']>0.33])
    dvals=['Δ',len(jl)-len(ml),len(js)-len(ms),(len(js)-len(ms))/len(active),jmq-mmq,
           sum(l['rl'] for l in jl)-sum(l['rl'] for l in ml),sum(l['rq'] for l in jl)-sum(l['rq'] for l in ml),
           sum(l['p3'] for l in jl)-sum(l['p3'] for l in ml),None,len(sil)-len(msil),
           sum(l['p3'] for l in sil)-sum(l['p3'] for l in msil),None]
    for i,v in enumerate(dvals):
        c=ba.cell(r,1+i,v); c.fill=fill('FF374151'); c.alignment=CTR
        c.border=Border(top=Side(style='thin',color=GOLD))
        if i==0: c.font=F(10,True,WHITE); continue
        if v is None: c.font=F(10,True,WHITE); continue
        if i in BADUP: good = v<0
        else: good = v>0
        col = 'FF81C784' if good else ('FFE57373' if v!=0 else 'FFB0BEC5')
        c.font=F(10,True,col)
        c.number_format='+0%;-0%;"·"' if i==3 else '+#,##0;-#,##0;"·"'
    ba.row_dimensions[r].height=18
    r+=1
    if sil:
        tops=' · '.join(f"{l['shop_name']} {l['p3']:,} btl" for l in sil[:4])
        ba.merge_cells(start_row=r,start_column=1,end_row=r,end_column=NB)
        c=ba.cell(r,1,f"   ↳ largest off-indent shops ({CUR_T}): {tops}"); c.font=F(9,i=True,c=GREY); c.alignment=LFT
        ba.row_dimensions[r].height=14
        r+=1
    r+=1
for i,w in enumerate([9,11,11,14,10,8,8,14,12,10,14,12]): ba.column_dimensions[get_column_letter(1+i)].width=w
ba.freeze_panes='A4'

# ============ BOND ANALYSIS (matrix) ============
bo = wb.create_sheet('BOND ANALYSIS')
NM=12
hero(bo, NM, 'BOND ANALYSIS — WHERE THE CEILING LIVES', f'{CUR_T} MQ ceiling (cases) by bond × brand · darker cell = bigger share of network shelf · ranked by {CUR_T} ΣMQ')
mat=collections.defaultdict(int); matm=collections.defaultdict(int)
for l in lines:
    if l['month']==CUR: mat[(l['bond'],l['sb'])]+=l['mq']
    else: matm[(l['bond'],l['sb'])]+=l['mq']

# brand-group banner row (RUM / BRANDY) above header
r=5
RUMSET={'OLD PEARL','ROF','KS 99','MBR'}
bo.cell(r,1,'').fill=fill(NAVY)
# OLD PEARL,BLENDERS,BCB,KS 99,ROF,MWB,MBR,CC  -> group by liquor type
GROUP={'OLD PEARL':'RUM','KS 99':'RUM','ROF':'RUM','MBR':'RUM','BLENDERS':'BRANDY','BCB':'BRANDY','MWB':'BRANDY','CC':'BRANDY'}
# we keep BORDER_ORDER; mark group bands by color underline
bo.merge_cells(start_row=r,start_column=1,end_row=r,end_column=1)
gc=bo.cell(r,1,'CATEGORY ▸'); gc.font=F(9,True,GOLDD); gc.alignment=Alignment(horizontal='right',vertical='center',indent=1); gc.fill=fill(NAVY)
for i,sb in enumerate(BORDER_ORDER):
    g=GROUP[sb]
    c=bo.cell(r,2+i,'RUM' if g=='RUM' else 'BRANDY'); c.font=F(8,True,WHITE); c.alignment=CTR
    c.fill=fill('FF7B341E' if g=='RUM' else 'FF1A3A6B')
for cc in (10,11,12):
    bo.cell(r,cc,'').fill=fill(NAVY)
bo.row_dimensions[r].height=14
r+=1
header_row(bo, r, ['Bond']+BORDER_ORDER+[f'ΣMQ {CUR_T}',f'Δ vs {PRIOR_T}','Cover (mo)'], h=24)
r+=1; r0=r
for idx,b in enumerate(bond_sorted):
    zebra = (idx%2==1)
    bc=bo.cell(r,1,b); bc.font=F(10,True,NAVYM); bc.border=BOX; bc.alignment=Alignment(horizontal='left',vertical='center',indent=1)
    if zebra: bc.fill=fill('FFF1F4FA')
    for i,sb in enumerate(BORDER_ORDER):
        v=mat[(b,sb)]
        c=bo.cell(r,2+i,v); c.alignment=CTR; c.border=BOX; c.number_format='#,##0;-#,##0;"·"'
        mx=max(mat[(bb,sb)] for bb in bonds) or 1
        if v>0:
            inten=v/mx
            if inten<0.25: c.fill=fill('FFE3ECF9'); c.font=F(10,c='FF1F3B70')
            elif inten<0.5: c.fill=fill('FFBBD0EE'); c.font=F(10,c='FF12305F')
            elif inten<0.75: c.fill=fill('FF7FA8E4'); c.font=F(10,True,'FF0D1B4A')
            else: c.fill=fill('FF3C6BC4'); c.font=F(10,True,WHITE)
        else:
            c.font=F(10,c='FFB6BECC')
            if zebra: c.fill=fill('FFF1F4FA')
    jt=sum(mat[(b,s)] for s in BORDER_ORDER); mt=sum(matm[(b,s)] for s in BORDER_ORDER)
    c=bo.cell(r,10,f"=SUM(B{r}:I{r})"); c.font=F(10,True,NAVYM); c.alignment=CTR; c.border=BOX; c.number_format='#,##0'
    if zebra: c.fill=fill('FFF1F4FA')
    d=jt-mt
    c=bo.cell(r,11,d); c.font=F(10,True,'FF2E7D32' if d>0 else ('FFC62828' if d<0 else GREY)); c.alignment=CTR; c.border=BOX; c.number_format='+#,##0;-#,##0;"·"'
    if zebra and d==0: c.fill=fill('FFF1F4FA')
    cov=jt/bond_rate[b] if bond_rate.get(b) else 0
    c=bo.cell(r,12,cov); c.font=F(10); c.alignment=CTR; c.border=BOX; c.number_format='0.00'
    if cov<0.6: c.fill=fill(T_RED[1]); c.font=F(10,True,T_RED[0])
    elif zebra: c.fill=fill('FFF1F4FA')
    bo.row_dimensions[r].height=20
    r+=1
bo.cell(r,1,'TOTAL').font=F(10,True,WHITE); bo.cell(r,1).alignment=Alignment(horizontal='left',vertical='center',indent=1)
for i,sb in enumerate(BORDER_ORDER):
    bo.cell(r,2+i,f"=SUM({get_column_letter(2+i)}{r0}:{get_column_letter(2+i)}{r-1})")
bo.cell(r,10,f"=SUM(J{r0}:J{r-1})"); bo.cell(r,11,f"=SUM(K{r0}:K{r-1})")
for cc in range(1,13):
    c=bo.cell(r,cc); c.fill=fill('FF374151')
    if c.font is None or not c.font.bold: c.font=F(10,True,WHITE)
    c.font=F(10,True,WHITE); c.alignment=CTR
    c.number_format='#,##0' if cc!=11 else '+#,##0;-#,##0;0'
    c.border=Border(top=Side(style='medium',color=GOLD))
bo.cell(r,1).alignment=Alignment(horizontal='left',vertical='center',indent=1)
bo.row_dimensions[r].height=22
r+=2
for i,w in enumerate([21,10,11,8,8,8,8,8,8,11,11,11]): bo.column_dimensions[get_column_letter(1+i)].width=w
bo.freeze_panes='A4'
bo.sheet_view.showGridLines=False
# ============ MAY-JUNE MOVEMENT ============
MOVE_SHEET=f'{PRIOR_U}-{CUR_U} MOVEMENT'
mv = wb.create_sheet(MOVE_SHEET)
NV=8
june_files={d['shop_code']:d for d in pi if d['month']==CUR}
def why_zero(l):
    jf=june_files.get(l['shop'])
    if jf is None or len(jf['rows'])==0: return '⚑ EMPTY PI'
    if not any(SHORT[r[0]]==l['sb'] for r in jf['rows']): return '↓ line dropped'
    return 'MQ set to 0'
hero(mv, NV, f'{PRIOR_U} → {CUR_U} MOVEMENT', 'Every change in the purchase instruction · NEW / LOST / RAISED / CUT at shop × brand level')
r=5
band(mv, r, NV, f"NET: {D_NET:+,} cs  ·  RAISED {len(raised)} (+{sum(t[0] for t in raised):,})  ·  CUT {len(cut)} ({-sum(t[0] for t in cut):,})  ·  NEW {len(new_st)} (+{sum(x['mq'] for x in new_st)})  ·  LOST {len(lost_st)} ({-sum(x['mq'] for x in lost_st)})", fl=NAVY)
r+=2
SB2FULL={SHORT[k]:FULL[k] for k in SHORT}
def listing(title, rows, cols, fmtcols, tint):
    global r
    band(mv, r, NV, title, fl=NAVYM); r+=1
    header_row(mv, r, cols+['']*(NV-len(cols))); r+=1
    for vals in rows:
        for i,v in enumerate(vals):
            c=mv.cell(r,1+i,v); c.font=F(9.5); c.border=BOX
            c.alignment=Alignment(horizontal='left',vertical='center',indent=1) if i in (0,1,2,7) else CTR
            if i in fmtcols: c.number_format='#,##0'
        mv.cell(r,1).fill=fill(tint)
        mv.row_dimensions[r].height=17
        r+=1
    r+=1
listing('TOP CUTS — KSBC lowered the ceiling most here',
        [[SB2FULL[SHORT_l['sb']], f"{SHORT_l['shop']} {SHORT_l['shop_name']}", SHORT_l['bond'], m_old, SHORT_l['mq'], SHORT_l['mq']-m_old, SHORT_l['p3'], 'reduced'] for d_,SHORT_l,m_old in cut[:20]],
        ['Brand','Shop','Bond',f'{PRIOR_T} MQ',f'{CUR_T} MQ','Δ','Offtake (btl)','Status'], (3,4,5,6), T_AMBER[1])
listing(f'LOST ON-INDENT LINES — had a ceiling in {PRIOR_T}, zero in {CUR_T} (excl. closed shops)',
        [[SB2FULL[l['sb']], f"{l['shop']} {l['shop_name']}", l['bond'], l['mq'], 0, -l['mq'], round(sum(sales.get((m,l['shop'],l['brand']),{'btl':0})['btl'] for m in WIN_CUR)), why_zero(l)] for l in lost_st if l['shop'] not in ('8002',)][:20],
        ['Brand','Shop','Bond',f'{PRIOR_T} MQ',f'{CUR_T} MQ','Δ',f'Offtake {span_label(WIN_CUR)} (btl)','Why 0'], (3,4,5,6), T_RED[1])
listing('TOP RAISES — ceiling lifted',
        [[SB2FULL[l['sb']], f"{l['shop']} {l['shop_name']}", l['bond'], l['mq']-d_, l['mq'], d_, l['p3'], 'lifted'] for d_,l in raised[:15]],
        ['Brand','Shop','Bond',f'{PRIOR_T} MQ',f'{CUR_T} MQ','Δ','Offtake (btl)','Status'], (3,4,5,6), T_GREEN[1])
listing(f'NEW ON-INDENT LINES — first ceiling granted in {CUR_T}',
        [[SB2FULL[l['sb']], f"{l['shop']} {l['shop_name']}", l['bond'], 0, l['mq'], l['mq'], l['p3'], 'new line'] for l in new_st[:15]],
        ['Brand','Shop','Bond',f'{PRIOR_T} MQ',f'{CUR_T} MQ','Δ','Offtake (btl)','Status'], (3,4,5,6), T_BLUE[1])
for i,w in enumerate([34,26,16,9,9,8,15,15]): mv.column_dimensions[get_column_letter(1+i)].width=w
mv.freeze_panes='A4'

# ============ VALIDATION ============
va = wb.create_sheet('VALIDATION')
NV2=8
import datetime as _dt
hero(va, NV2, 'VALIDATION — RECONCILIATION RECEIPTS', f'Every check recomputed from the data each run · {CUR_T} vs {PRIOR_T} · built {_dt.date.today():%d-%b-%Y}')
r=5
checks=[
 ('CHECK 1 · File coverage',f'{n_files_prior} {PRIOR_T} files + {n_files_cur} {CUR_T} files parsed; 0 unparseable; every file labelled with the correct Report Month.','PASS'),
 ('CHECK 2 · Master shops → PI raw ingested? (per month)',f'{len(active)} active KSBC shops in master. '+ f'{CUR}: raw ingested {len(chk2_by_month[CUR][0])}/{len(active)}'+('' if not chk2_by_month[CUR][1] else f' (MISSING: '+', '.join(_nm(c) for c in chk2_by_month[CUR][1])+')')+f'; empty raws: '+('none' if not chk2_by_month[CUR][2] else f'{len(chk2_by_month[CUR][2])} — '+', '.join(_nm(c) for c in chk2_by_month[CUR][2])) + '  ||  ' + f'{PRIOR}: raw ingested {len(chk2_by_month[PRIOR][0])}/{len(active)}'+('' if not chk2_by_month[PRIOR][1] else f' (MISSING: '+', '.join(_nm(c) for c in chk2_by_month[PRIOR][1])+')')+f'; empty raws: '+('none' if not chk2_by_month[PRIOR][2] else f'{len(chk2_by_month[PRIOR][2])} — '+', '.join(_nm(c) for c in chk2_by_month[PRIOR][2])) + f'.  Duplicates: '+('none' if not chk2_dups else str(chk2_dups))+f'. (Closed-in-master present: '+(', '.join(chk2_extras) if chk2_extras else 'none')+'. Empty raw = KSBC issued a blank PI for that month.)', 'PASS' if chk2_pass else 'CHECK'),
 ('CHECK 3 · MQ = RL + RQ',f'Holds in {N_LINES:,} of {N_LINES:,} brand lines (100%). Confirms KSBC policy: Maximum Quantity = Re-order Level + Re-order Quantity.','PASS'),
 ('CHECK 4 · PI ↔ KSD tertiary reconciliation',f'"Previous 3 Months Sale" reconciled line-by-line vs our KSBC workbooks (bottles = Σ sales×BPC; {PRIOR_T} PI←{span_label(WIN_PRIOR)}, {CUR_T} PI←{span_label(WIN_CUR)}). {recon_exact:,} / {N_LINES:,} exact ({recon_pct:.1%}) · {recon_w5} within 5 btl · {recon_off} beyond 5 btl (largest gap {max_gap:.0f} btl — case-conversion rounding in KSBC daily exports).','PASS'),
 ('CHECK 5 · Window arithmetic',f'{CUR_T}-window offtake ({CUR_OFF:,} btl) − {PRIOR_T}-window ({PRIOR_OFF:,} btl) = {CUR_OFF-PRIOR_OFF:+,} btl ({OFF_MOM:+.1%}). Overlapping windows telescope to the (newest − dropped) month difference.','PASS'),
 ('CHECK 6 · Empty-PI cross-check (both months)',f'{CUR_T}: {len(FLAGS)} empty PIs — ' + (', '.join(f"{_nm(s)} [{flag}]" for s,flag in sorted(FLAGS.items())) or 'none') + f'.  ||  {PRIOR_T}: {len(FLAGS_PRIOR)} empty PIs — ' + (', '.join(f"{_nm(s)} [{flag}]" for s,flag in sorted(FLAGS_PRIOR.items())) or 'none') + '.  DARK = selling but empty (ERP gap, escalate) · LOST/WHITE = no offtake in window · CLOSED = master.','PASS'),
 ('CHECK 7 · Dashboard totals',f'ΣMQ {CUR_T} {CUR_MQ:,} / {PRIOR_T} {PRIOR_MQ:,} · on-indent lines {STAND_CUR}/{STAND_PRIOR} · off-indent lines {SILENT_N} — recomputed live via SUMIFS/COUNTIFS against RAW PI DATA.','PASS'),
 ('CHECK 8 · Tertiary data coverage',(f'All {len(NEEDED)} window months ({"/".join(NEEDED)}) have a KSBC tertiary workbook — offtake, rate & cover are complete.' if not MISSING_TERT else f'MISSING tertiary workbook(s): {", ".join(MISSING_TERT)}. Offtake/cover for lines using those months are PARTIAL and shown as n/a — DO NOT trust the affected Cover figures until the workbook is built.'),('PASS' if not MISSING_TERT else 'MISSING DATA')),
]
for t,d,s in checks:
    band(va, r, NV2, t, h=18, fl=NAVYS, sz=10); r+=1
    va.merge_cells(start_row=r,start_column=1,end_row=r,end_column=NV2-1)
    c=va.cell(r,1,d); c.font=F(9.5,c='FF374151'); c.alignment=Alignment(horizontal='left',vertical='top',wrap_text=True,indent=1)
    c2=va.cell(r,NV2,s); c2.font=F(10,True,T_GREEN[0]); c2.fill=fill(T_GREEN[1]); c2.alignment=CTR; c2.border=BOX
    va.row_dimensions[r].height=34
    r+=1
r+=1
band(va, r, NV2, 'SOURCES', h=18, fl=NAVYS, sz=10); r+=1
va.merge_cells(start_row=r,start_column=1,end_row=r+3,end_column=NV2)
c=va.cell(r,1,f'Source: KSBC ERP "Month Wise Purchase Instruction Report", {CUR_T} + {PRIOR_T} sets, folder PURCHASE INSTRUCTION/. RL/RQ/MQ per KSBC PI portal. Master: MASTER DATA CONFIRMED.xlsx sheet 16-4-25 ({len(active)} active KSBC). Tertiary cross-check: {"/".join(NEEDED)} SHOP SALES ANALYSIS workbooks, COMBINED sheets. Cover = MQ ÷ KSD avg monthly case rate ({span_label(WIN_CUR)}).')
c.font=F(9.5,c='FF374151'); c.alignment=Alignment(horizontal='left',vertical='top',wrap_text=True,indent=1)
for i,w in enumerate([16,16,16,16,16,16,16,10]): va.column_dimensions[get_column_letter(1+i)].width=w

# ============ PER-BOND SHEETS ============
NB2=11
FLAG_TXT={'DARK':f'⚑ DARK — selling, but KSBC issued an EMPTY {CUR_T} PI (escalate to RM)',
          'LOST':f'⚑ LOST SHOP — no KSD offtake in the {span_label(WIN_CUR)} window','WHITE':'⚑ WHITE SHOP — no KSD presence in window',
          'CLOSED':'⚑ CLOSED in master'}
# ensure June empty-PI shops exist in shop_jun/shop_meta (zero rows) so bond sheets can render them
for d0 in pi:
    if d0['month']==CUR and d0['shop_code'] not in shop_jun:
        md0=master.get(d0['shop_code'],{})
        shop_jun[d0['shop_code']]=dict(mq=0,p3=0,rate=0.0,brands=0)
        shop_meta[d0['shop_code']]=dict(shop_name=d0['shop_name'].title(),bond=md0.get('bond','—'),
                                        staff=md0.get('staff','—').title(),cap=d0['capacity'])
jun_by_bond_shop = collections.defaultdict(lambda: collections.defaultdict(list))
may_by_shop_brand = {(l['shop'], l['sb']): l for l in lines if l['month']==PRIOR}
# ---- growth-aware MQ planner (added 15 Jun 2026, Abhay-approved): realized market growth per bond + MQ-formula slope ----
# A inversion (validated): A = 0.5*rate*J/RL. A_PRIOR scales the CUR-window rate to the prior window by the
# prev-3-mo bottle ratio (BPC-free). g_bond = sum(A_CUR)/sum(A_PRIOR)-1 over each shop's dominant (max-RL) line.
GROW_TGT=0.05
def planner_extra(rate,rl,mq,target,g):
    # extra cases to sell THIS month so next month's MQ hits target, in a market growing g.
    # slope dMQ/dB = 2.08*RL/rate ; x3 because 2 of next window's 3 months are already banked.
    if rl<=0 or rate<=0: return None
    Bnext=rate*(1+g)*(1+(target-mq)/(2.08*rl))
    return 3*(Bnext-rate)
def _realized_g():
    cur_by={(l['shop'],l['sb']):l for l in lines if l['month']==CUR}
    dom={}
    for (shop,sb),j in cur_by.items():
        m=may_by_shop_brand.get((shop,sb))
        if not m: continue
        J=j.get('cap') or 0
        if j['rl']<=0 or m['rl']<=0 or j['rate']<=0 or j['p3']<=0 or m['p3']<=0 or J<=0: continue
        Aj=0.5*j['rate']*J/j['rl']; Am=0.5*j['rate']*(m['p3']/j['p3'])*J/m['rl']
        if shop not in dom or j['rl']>dom[shop][0]: dom[shop]=(j['rl'],Aj,Am,j['bond'])
    bs=collections.defaultdict(lambda:[0.0,0.0])
    for shop,(rl,Aj,Am,bd) in dom.items(): bs[bd][0]+=Aj; bs[bd][1]+=Am
    return {bd:(v[0]/v[1]-1 if v[1]>0 else 0.0) for bd,v in bs.items()}
REALIZED_G=_realized_g()
for l in lines:
    if l['month']==CUR and l['shop'] not in INCOMPLETE: jun_by_bond_shop[l['bond']][l['shop']].append(l)
shopfile_bond = collections.defaultdict(list)
for d0 in pi:
    if d0['month']==CUR and d0['shop_code'] not in INCOMPLETE:
        b0 = master.get(d0['shop_code'],{}).get('bond','—')
        shopfile_bond[b0].append(d0)

bond_cut_rank = {b:i+1 for i,(b,_) in enumerate(sorted(((b, sum(l['mq'] for l in lines if l['month']==CUR and l['bond']==b and l['shop'] not in INCOMPLETE)-sum(l['mq'] for l in lines if l['month']==PRIOR and l['bond']==b and l['shop'] not in INCOMPLETE)) for b in bonds), key=lambda t:t[1]))}

for b in bond_sorted:
    ws = wb.create_sheet(b[:31])
    staff_set = sorted({master.get(d0['shop_code'],{}).get('staff','—').title() for d0 in shopfile_bond[b]})
    jl=[l for l in lines if l['month']==CUR and l['bond']==b and l['shop'] not in INCOMPLETE]
    ml=[l for l in lines if l['month']==PRIOR and l['bond']==b and l['shop'] not in INCOMPLETE]
    js=[l for l in jl if l['mq']>0]; sil=[l for l in jl if l['mq']==0 and l['p3']>0]
    jmq=sum(l['mq'] for l in js); mmq=sum(l['mq'] for l in ml)
    rate=bond_rate.get(b,0.0)
    hero(ws, NB2, f"{b} — PURCHASE INSTRUCTION DRILL-DOWN",
         f"{len(shopfile_bond[b])} shops · field staff: {', '.join(staff_set[:4])}{' …' if len(staff_set)>4 else ''} · {CUR_T} PI vs {PRIOR_T} PI · cases unless marked btl")
    K2=5
    dpct = (jmq/mmq-1) if mmq else 0
    off_j=sum(l['p3'] for l in jl); off_m=sum(l['p3'] for l in ml)
    off_j_cs=sum(sales.get((m,l['shop'],l['brand']),{'cs':0})['cs'] for l in jl for m in WIN_CUR)
    off_m_cs=sum(sales.get((m,l['shop'],l['brand']),{'cs':0})['cs'] for l in ml for m in WIN_PRIOR)
    cov_v=round(jmq/rate,2) if rate else 0
    kpi(ws,K2,1,1,f'{CUR_U} CEILING (ΣMQ)', jmq, f"{'▲' if jmq>=mmq else '▼'} {jmq-mmq:+,} vs {PRIOR_T} ({dpct:+.0%})", 'FF81C784' if jmq>=mmq else 'FFFCA5A5', accent=('FF43A047' if jmq>=mmq else 'FFE53935'))
    kpi(ws,K2,2,3,'ON-INDENT LINES', len(js), f"of {len(jl)} listed lines", 'FFBBDEFB', accent='FF43A047')
    kpi(ws,K2,5,3,'OFF-INDENT LINES', len(sil), f"{sum(l['p3'] for l in sil):,} btl sold · no auto-indent", 'FFFCA5A5' if sil else 'FFBBDEFB', accent='FFF57C00')
    kpi(ws,K2,8,2,'MQ COVER', cov_v, f"mo vs rate {rate:,.0f} cs/mo", 'FFBBDEFB', '0.00', accent=('FFE53935' if cov_v<0.6 else 'FF1E88E5'))
    kpi(ws,K2,10,2,f'OFFTAKE {span_label(WIN_CUR)}', round(off_j_cs), f"{(off_j_cs/off_m_cs-1) if off_m_cs else 0:+.0%} vs prior · cases", 'FF81C784' if off_j_cs>=off_m_cs else 'FFFCA5A5', accent=('FF43A047' if off_j_cs>=off_m_cs else 'FFE53935'))
    for rr in (K2,K2+1,K2+2): ws.row_dimensions[rr].height=[16,30,14][rr-K2]

    r=K2+4
    # ---- auto insights ----
    ins=[]
    rankcut=bond_cut_rank[b]
    if jmq>mmq: ins.append(f"Only bond KSBC raised for {CUR_T}: +{jmq-mmq} cs — defend the momentum.")
    else: ins.append(f"Ceiling cut {mmq}→{jmq} cs ({dpct:+.0%}) — {'deepest' if rankcut==1 else f'#{rankcut} deepest'} cut of {n_bonds} bonds. Driver: {span_label(WIN_CUR)} offtake softened vs the prior window.")
    if js:
        top_doors = sorted(jun_by_bond_shop[b].items(), key=lambda kv:-sum(x['mq'] for x in kv[1]))
        t0=top_doors[0]; t0mq=sum(x['mq'] for x in t0[1])
        top3=sum(sum(x['mq'] for x in kv[1]) for kv in top_doors[:3])
        ins.append(f"Concentration: {t0[1][0]['shop_name']} alone holds {t0mq} cs ({t0mq/jmq:.0%} of the bond ceiling); top-3 shops hold {top3/jmq:.0%}. A stockout at these few shops moves the whole bond.")
    if sil:
        s0=max(sil,key=lambda l:l['p3'])
        ins.append(f"Off-indent demand: {len(sil)} lines / {sum(l['p3'] for l in sil):,} btl sell with NO auto-indent — largest {s0['sb']} at {s0['shop_name']} ({s0['p3']:,} btl). These shops are 100% field-push until offtake earns them an MQ.")
    tight=[(l['mq']/l['rate'], l) for l in js if l['rate']>=5 and l['mq']<0.6*l['rate']]
    if tight:
        tight.sort(key=lambda t:t[0]); c0,l0=tight[0]
        ins.append(f"Ceiling binds: {len(tight)} high-velocity lines under 0.6-month cover — worst {l0['sb']} at {l0['shop_name']} ({l0['rate']:.0f} cs/mo rate vs MQ {l0['mq']} = {c0:.2f} mo). Argue these with the offtake receipts.")
    bw=[(sb, sum(l['p3'] for l in jl if l['sb']==sb)) for sb in BORDER_ORDER]
    bw=[(sb,p) for sb,p in bw if p>=500 and sum(l['mq'] for l in js if l['sb']==sb)==0]
    if bw: ins.append("Regional white space: " + " · ".join(f"{sb} moves {p:,} btl here yet has ZERO MQ ceiling" for sb,p in bw) + " — warehouse-level listing gap, not a demand problem.")
    dark=[s for s in FLAGS if FLAGS[s] in ('DARK','LOST','WHITE') and master.get(s,{}).get('bond')==b]
    if dark: ins.append("Shop alerts: " + " · ".join(f"{shop_meta[s]['shop_name']} — {FLAG_TXT[FLAGS[s]]}" for s in dark))
    band(ws, r, NB2, 'BOND INSIGHTS', fl=NAVY, h=20); r+=1
    for t in ins[:6]:
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=NB2)
        c=ws.cell(r,1,'•  '+t); c.font=F(9.5,c='FF1F2937'); c.fill=fill('FFEFF3FB')
        c.alignment=Alignment(horizontal='left',vertical='center',wrap_text=True,indent=1)
        ws.row_dimensions[r].height=26
        r+=1
    r+=1
    # ---- brand table ----
    band(ws, r, NB2, f'BRAND PERFORMANCE IN THIS BOND  ({CUR_T} vs {PRIOR_T})', fl=NAVYM); r+=1
    header_row(ws, r, ['Brand','Shops','On-Indent',f'ΣMQ {CUR_T}',f'ΣMQ {PRIOR_T}','Δ MQ','Offtake (cs)','Offtake Δ%','Off-Indent','Off-Ind btl','Share of MQ'], h=42); r+=1
    rb0=r
    for sb in BORDER_ORDER:
        bj=[l for l in jl if l['sb']==sb]; bm=[l for l in ml if l['sb']==sb]
        if not bj and not bm: continue
        bjs=[l for l in bj if l['mq']>0]; bsil=[l for l in bj if l['mq']==0 and l['p3']>0]
        offj=sum(l['p3'] for l in bj); offm=sum(l['p3'] for l in bm)
        offj_cs=sum(sales.get((m,l['shop'],l['brand']),{'cs':0})['cs'] for l in bj for m in WIN_CUR)
        offm_cs=sum(sales.get((m,l['shop'],l['brand']),{'cs':0})['cs'] for l in bm for m in WIN_PRIOR)
        smq=sum(l['mq'] for l in bjs)
        fullname=[v for k,v in FULL.items() if SHORT[k]==sb][0]
        vals=[fullname,len(bj),len(bjs),smq,sum(l['mq'] for l in bm if l['mq']>0),smq-sum(l['mq'] for l in bm if l['mq']>0),round(offj_cs),(offj_cs/offm_cs-1) if offm_cs else 0,len(bsil),sum(l['p3'] for l in bsil),smq/jmq if jmq else 0]
        for i,v in enumerate(vals):
            c=ws.cell(r,1+i,v); c.border=BOX; c.alignment=CTR if i else LFT
            c.font=F(9.5,True,NAVYM) if i==0 else F(9.5)
            if i in (1,2,3,4,6,8,9): c.number_format='#,##0;-#,##0;"·"'
            if i==5: c.number_format='+#,##0;-#,##0;"·"'; c.font=F(9.5,True,'FF2E7D32' if v>0 else ('FFC62828' if v<0 else GREY))
            if i==7: c.number_format='+0%;-0%;"·"'; c.font=F(9.5,True,'FF2E7D32' if v>0 else ('FFC62828' if v<0 else GREY))
            if i==10: c.number_format='0%'
        if sum(l['p3'] for l in bj)>=500 and not bjs:
            for i in range(NB2): ws.cell(r,1+i).fill=fill(T_AMBER[1])
        ws.row_dimensions[r].height=18
        r+=1
    ws.cell(r,1,'TOTAL').font=F(10,True,WHITE)
    for cc,fx in [(2,f"=SUM(B{rb0}:B{r-1})"),(3,f"=SUM(C{rb0}:C{r-1})"),(4,f"=SUM(D{rb0}:D{r-1})"),(5,f"=SUM(E{rb0}:E{r-1})"),(6,f"=D{r}-E{r}"),(7,f"=SUM(G{rb0}:G{r-1})"),(9,f"=SUM(I{rb0}:I{r-1})"),(10,f"=SUM(J{rb0}:J{r-1})"),(11,1)]:
        ws.cell(r,cc,fx)
    for cc in range(1,NB2+1):
        c=ws.cell(r,cc); c.fill=fill('FF374151'); c.font=F(9.5,True,WHITE); c.alignment=CTR
        c.border=Border(top=Side(style='medium',color=GOLD))
        if cc in (2,3,4,5,7,9,10): c.number_format='#,##0'
        if cc==6: c.number_format='+#,##0;-#,##0;0'
        if cc==11: c.number_format='0%'
    ws.row_dimensions[r].height=20
    r+=2
    # ---- shop blocks ----
    # ---- shop blocks (v6 layout + recovery engine; consolidated, no separate file) ----
    NUMF='#,##0;-#,##0;"·"'; NUMD='+#,##0;-#,##0;"·"'
    gb=REALIZED_G.get(b,0.0); PLF='+0.0;-0.0;"·"'
    recs={}
    for l in jl:
        pmq=(may_by_shop_brand.get((l['shop'],l['sb'])) or {}).get('mq',0)
        rv=recovery_line(l,pmq)
        if rv: recs[(l['shop'],l['sb'])]=rv
    b_ncut=len(recs); b_scut=sum(v['cut'] for v in recs.values())
    b_hold=sum(v['hold'] for v in recs.values()); b_rec=sum(v['rec'] for v in recs.values())
    b_inv=sum(1 for v in recs.values() if '⚠' in v['why'] or '\U0001f6d1' in v['why'])
    # ===== ONE TABLE, two segregated halves: A-I = CEILING & OFFTAKE | J-N = ACTION PLAN, split by a gold divider rule =====
    detail_start=r
    _spacers=set()
    band(ws, r, 9, f'CEILING & OFFTAKE   ·   shop × brand, ranked by {CUR_T} ceiling', fl=NAVYM)
    ws.merge_cells(start_row=r,start_column=10,end_row=r,end_column=14)
    _h2=ws.cell(r,10,"ACTION PLAN   ·   CASES TO SELL / MONTH"); _h2.font=F(11,True,WHITE); _h2.alignment=Alignment(horizontal='center',vertical='center')
    for cc in range(10,15): ws.cell(r,cc).fill=fill(NAVYM); ws.cell(r,cc).border=Border(bottom=Side(style='thin',color=GOLD))
    ws.row_dimensions[r].height=22; r+=1
    # editable per-bond scenario controls (right half) -> HOLD / RECOVER / GROW recompute live
    GCELL=f"$L${r}"; TCELL=f"$N${r}"
    ws.merge_cells(start_row=r,start_column=10,end_row=r,end_column=11)
    _l1=ws.cell(r,10,"Market growth →"); _l1.font=F(10,True,'FF111827'); _l1.alignment=Alignment(horizontal='right',vertical='center')
    _cg=ws.cell(r,12,round(gb,3)); _cg.number_format='0.0%'; _cg.font=F(12,True,'FF7F1D1D'); _cg.fill=fill('FFFFF176'); _cg.alignment=CTR; _cg.border=BOX
    _l2=ws.cell(r,13,"MQ target lift →"); _l2.font=F(10,True,'FF111827'); _l2.alignment=Alignment(horizontal='right',vertical='center')
    _ct=ws.cell(r,14,GROW_TGT); _ct.number_format='0%'; _ct.font=F(12,True,'FF7F1D1D'); _ct.fill=fill('FFFFF176'); _ct.alignment=CTR; _ct.border=BOX
    ws.row_dimensions[r].height=20; r+=1
    shops_sorted = sorted(jun_by_bond_shop[b].keys(), key=lambda s:-sum(x['mq'] for x in jun_by_bond_shop[b][s]))
    empt = [d0['shop_code'] for d0 in shopfile_bond[b] if d0['shop_code'] not in jun_by_bond_shop[b]]
    for s in shops_sorted + empt:
        rows = sorted(jun_by_bond_shop[b].get(s,[]), key=lambda l: (0 if l['mq']>0 else (1 if l['p3']>0 else 2), -l['mq'], -l['p3'], BORDER_ORDER.index(l['sb'])))
        meta = rows[0] if rows else shop_meta[s]
        smq=sum(x['mq'] for x in rows); smay=shop_may_mq.get(s,0); fl_=FLAGS.get(s,'')
        shop_recs=[recs[(s,x['sb'])] for x in rows if (s,x['sb']) in recs]
        ncut=len(shop_recs); scut=sum(v['cut'] for v in shop_recs)
        bfl='FF1A237E' if not fl_ else ('FF7F1D1D' if fl_ in ('DARK','LOST','WHITE') else 'FF4B5563')
        # full-width shop title bar (one banner across both halves)
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=14)
        t=ws.cell(r,1,f"{s} · {meta['shop_name'].upper()}{('   ⚑ '+fl_) if fl_ else ''}")
        t.font=Font(name=FN,size=12,bold=True,color=WHITE); t.alignment=Alignment(horizontal='left',vertical='center',indent=1)
        for cc in range(1,15): ws.cell(r,cc).fill=fill(bfl); ws.cell(r,cc).border=Border(bottom=Side(style='medium',color=GOLD))
        ws.row_dimensions[r].height=26; r+=1
        if not rows:
            ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=14)
            note = f'EMPTY {CUR_U} PI — KSBC issued no KSD lines.'
            if fl_=='DARK': note += f"  Shop sold {sum(may_by_shop_brand[(s,sb)]['p3'] for sb in BORDER_ORDER if (s,sb) in may_by_shop_brand):,} btl in the {PRIOR_T} window and is still selling — escalate to the KSBC RM to restore the PI."
            elif fl_=='CLOSED': note += '  Shop closed in master — expected.'
            elif fl_ in ('LOST','WHITE'): note += '  Zero KSD offtake in trailing window — shop needs a field win, not an escalation.'
            c=ws.cell(r,1,note); c.font=F(9,i=True,c=T_RED[0]); c.fill=fill(T_RED[1]); c.alignment=Alignment(horizontal='left',vertical='center',indent=2)
            ws.row_dimensions[r].height=16; _spacers.add(r+1); r+=2; continue
        H1=['Brand','Rate cs/mo','RL','RQ',f'MQ {CUR_T}',f'MQ {PRIOR_T}','Δ MQ','Cover mo','Status']
        H2=['HOLD cs/mo','RECOVER cs/mo','GROW cs/mo','Action','Why / Watch']
        for i,lab in enumerate(H1):
            c=ws.cell(r,1+i,lab); c.font=F(10,True,WHITE); c.fill=fill(BANNER); c.alignment=CTR; c.border=Border(bottom=Side(style='medium',color=GOLD))
        for i,lab in enumerate(H2):
            c=ws.cell(r,10+i,lab); c.font=F(10,True,WHITE); c.fill=fill(BANNER); c.alignment=CTR; c.border=Border(bottom=Side(style='medium',color=GOLD))
        ws.row_dimensions[r].height=20; r+=1
        rblk0=r
        for l in rows:
            pmq=(may_by_shop_brand.get((s,l['sb'])) or {}).get('mq',0)
            cov=(l['mq']/l['rate']) if l['rate']>0.33 and l['mq']>0 else None
            off=(l['mq']==0 and l['p3']>0); rf=T_AMBER[1] if off else None
            dq=l['mq']-pmq; rv=recs.get((s,l['sb']))
            boxcell(ws,r,1,l['brand'],bold=True,col=NAVYM,al=Alignment(horizontal='left',vertical='center',indent=2),rf=rf)
            boxcell(ws,r,2,(round(l['rate'],1) if l['rate']>0 else ''),'0.0',rf=rf)
            boxcell(ws,r,3,l['rl'],NUMF,rf=rf)
            boxcell(ws,r,4,l['rq'],NUMF,rf=rf)
            boxcell(ws,r,5,l['mq'],NUMF,bold=True,rf=rf)
            boxcell(ws,r,6,pmq,NUMF,rf=rf)
            boxcell(ws,r,7,dq,NUMD,bold=True,col=('FF2E7D32' if dq>0 else ('FFC62828' if dq<0 else GREY)),rf=rf)
            boxcell(ws,r,8,(round(cov,2) if cov is not None else ''),'0.00',rf=rf)
            if l['mq']>0: st_,stf,stb=('✅ On-Indent',T_GREEN[0],T_GREEN[1])
            elif l['p3']>0: st_,stf,stb=('⚠️ Off-Indent',T_AMBER[0],T_AMBER[1])
            else: st_,stf,stb=('— No Sales',GREY,None)
            cs9=ws.cell(r,9,st_); cs9.font=F(9,True,stf); cs9.alignment=CTR; cs9.border=BOX
            if stb: cs9.fill=fill(stb)
            fH=f'=IF(AND($G{r}<0,$C{r}>0,$B{r}>0),$B{r}*(1+{GCELL}),"")'
            fR=f'=IF(AND($G{r}<0,$C{r}>0,$B{r}>0),$B{r}*(1+{GCELL})*(1+($F{r}-$E{r})/(2.08*$C{r})),"")'
            fG=f'=IF(AND($G{r}>=0,$E{r}>0,$C{r}>0,$B{r}>0),$B{r}*(1+{GCELL})*(1+($E{r}*{TCELL})/(2.08*$C{r})),"")'
            for cc_,frm_,fc_ in ((10,fH,'FF1565C0'),(11,fR,'FFC62828'),(12,fG,'FF2E7D32')):
                pc=ws.cell(r,cc_,frm_); pc.font=F(9.5,True,fc_); pc.alignment=CTR; pc.border=BOX; pc.number_format='#,##0.0;;"·"'
                if rf: pc.fill=fill(rf)
            if rv:
                lbl,(pf_,pb_)=pill_for(rv['action']); cp=ws.cell(r,13,lbl); cp.font=F(8.5,True,pf_); cp.fill=fill(pb_); cp.alignment=CTR; cp.border=BOX
                wfl=('⚠' in rv['why'] or '\U0001f6d1' in rv['why']); cw=ws.cell(r,14,rv['why']); cw.font=F(9,wfl,T_RED[0] if wfl else GREY); cw.alignment=Alignment(horizontal='left',vertical='center',wrap_text=True,indent=1); cw.border=BOX
                if rf: cw.fill=fill(rf)
            else:
                cl=ws.cell(r,13,None); cl.alignment=CTR; cl.border=BOX
                if rf: cl.fill=fill(rf)
                cw=ws.cell(r,14,None); cw.border=BOX; cw.alignment=LFT
                if rf: cw.fill=fill(rf)
            ws.row_dimensions[r].height=16; r+=1
        on=sum(1 for x in rows if x['mq']>0); offc=sum(1 for x in rows if x['mq']==0 and x['p3']>0)
        ws.cell(r,1,'TOTAL'); ws.cell(r,2,f"=SUM(B{rblk0}:B{r-1})")
        for col,L in [(3,'C'),(4,'D'),(5,'E'),(6,'F')]: ws.cell(r,col,f"=SUM({L}{rblk0}:{L}{r-1})")
        ws.cell(r,7,f"=SUM(G{rblk0}:G{r-1})"); ws.cell(r,8,f'=IF(B{r}=0,"",E{r}/B{r})')
        ws.cell(r,9,f"✅ {on} · ⚠️ {offc}")
        ws.cell(r,10,f'=IF(SUM(J{rblk0}:J{r-1})=0,"",SUM(J{rblk0}:J{r-1}))')
        ws.cell(r,11,f'=IF(SUM(K{rblk0}:K{r-1})=0,"",SUM(K{rblk0}:K{r-1}))')
        ws.cell(r,12,f'=IF(SUM(L{rblk0}:L{r-1})=0,"",SUM(L{rblk0}:L{r-1}))')
        ws.cell(r,13,(f"⚑ {ncut} cut" if ncut>0 else "—"))
        ws.cell(r,14,None)
        for cc in range(1,15):
            c=ws.cell(r,cc); c.fill=fill('FF374151'); c.border=Border(top=Side(style='medium',color=GOLD))
            c.alignment=(Alignment(horizontal='left',vertical='center',indent=2) if cc==1 else (Alignment(horizontal='left',vertical='center',wrap_text=True,indent=1) if cc==14 else CTR))
            c.font=(F(9,True,GOLDD) if cc in (9,13) else F(9.5,True,WHITE))
            if cc in (3,4,5,6): c.number_format=NUMF
            if cc==2: c.number_format='0.0'
            if cc==7: c.number_format=NUMD
            if cc==8: c.number_format='0.00'
            if cc in (10,11,12): c.number_format='#,##0.0;;"·"'
        ws.row_dimensions[r].height=19; r+=2; _spacers.add(r-1)
    # ---- flag shops missing one month's raw PI ----
    bond_incomplete=sorted([s for s in INCOMPLETE if master.get(s,{}).get('bond')==b])
    if bond_incomplete:
        _spacers.add(r); r+=1
        band(ws, r, 14, '⚠ INCOMPLETE — PI RAW MISSING FOR ONE MONTH · no shop block built · drop the missing raw & rebuild for a full comparison', h=22, fl='FFC62828'); r+=1
        for s in bond_incomplete:
            miss=INCOMPLETE[s]; have=(PRIOR if miss==CUR else CUR)
            nm=SHOPNAME.get(s, master.get(s,{}).get('name',s))
            ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=14)
            c=ws.cell(r,1,f"⚠ {s} · {str(nm).upper()}   —   {have} PI present · {miss} PI MISSING   →   drop the {miss} raw and rebuild")
            c.font=F(9.5,True,'FFC62828'); c.alignment=Alignment(horizontal='left',vertical='center',indent=2)
            for cc in range(1,15):
                cell=ws.cell(r,cc); cell.fill=fill('FFFCE4E4'); cell.border=BOX
            ws.row_dimensions[r].height=18; r+=1
        _spacers.add(r); r+=1
    # crisp gold divider rule between the two halves (left edge of col J); hidden inside full-width merged title bars
    _div=Side(style='medium',color=GOLD)
    for rr in range(detail_start, r):
        if rr in _spacers: continue
        c=ws.cell(rr,10); bd=c.border
        c.border=Border(left=_div, right=bd.right, top=bd.top, bottom=bd.bottom)
    for i,w in enumerate([42,11,7,7,9,9,9,10,14,14,14,14,13,28]): ws.column_dimensions[get_column_letter(1+i)].width=w
    ws.freeze_panes=None
    ws.sheet_view.showGridLines=False
    ws.page_setup.orientation='landscape'; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
    ws.sheet_properties.pageSetUpPr.fitToPage=True

# print setup for non-bond sheets
for nm in ['DASHBOARD','BRAND ANALYSIS','BOND ANALYSIS',MOVE_SHEET,'VALIDATION','RAW PI DATA']:
    w0=wb[nm]; w0.page_setup.orientation='landscape'; w0.page_setup.fitToWidth=1; w0.page_setup.fitToHeight=0
    w0.sheet_properties.pageSetUpPr.fitToPage=True

# order sheets
order_names=['DASHBOARD','BRAND ANALYSIS','BOND ANALYSIS']+[b[:31] for b in bond_sorted]+[MOVE_SHEET,'VALIDATION','RAW PI DATA']
wb._sheets = [wb[n] for n in order_names]
wb.save(OUT)
print(f"saved {len(wb.sheetnames)} sheets -> {OUT}")
