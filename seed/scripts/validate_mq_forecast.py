#!/usr/bin/env python3
"""Validate the MQ planner forecast against KSBC's actual MQ once a NEW month's PI lands.
Run after building the new PI INSIGHTS workbook (it carries both months in RAW PI DATA):
    python3 .claude/scripts/validate_mq_forecast.py "PURCHASE INSTRUCTION/PI INSIGHTS - JUNE & JULY 2026.xlsx"
Checks (all on KSBC's own published numbers — no external data):
  1. FORWARD-MODEL TEST: predict NEW-month MQ from OLD-month MQ/RL + the realized offtake change
     (prev-3-mo bottle ratio) + realized per-bond market growth g; compare to KSBC's actual NEW MQ.
     If the planner formula is KSBC's real method this reproduces published MQ within ~1 case.
  2. GROWTH-SEED TEST: realized g per bond (A_OLD->A_NEW) — the number the planner seeds its slider with.
A = 0.5*rate*J/RL (validated inversion); A_OLD scales NEW-window rate by the bottle ratio (BPC-free).
"""
import sys, openpyxl, collections, statistics as st
MORD=['January','February','March','April','May','June','July','August','September','October','November','December']
def mi(m): return MORD.index(m) if m in MORD else -1

wbpath=sys.argv[1] if len(sys.argv)>1 else None
if not wbpath:
    import glob,os
    c=sorted(glob.glob("PURCHASE INSTRUCTION/PI INSIGHTS - * 2026.xlsx"),key=os.path.getmtime)
    wbpath=c[-1] if c else sys.exit("No PI INSIGHTS workbook found; pass the path.")
wb=openpyxl.load_workbook(wbpath, read_only=True, data_only=True)
ws=wb["RAW PI DATA"]; rows=list(ws.iter_rows(values_only=True))
# header at first row whose col0=='Month'
h=[i for i,r in enumerate(rows) if r and r[0]=='Month'][0]
recs=[]
for r in rows[h+1:]:
    if not r or not r[0] or mi(r[0])<0: continue
    recs.append(dict(month=r[0],shop=str(r[1]),bond=r[3],J=r[6] or 0,brand=r[7],
                     p3=r[9] or 0,rl=r[10] or 0,rq=r[11] or 0,mq=r[12] or 0,rate=r[13] or 0))
months=sorted({x['month'] for x in recs}, key=mi)
if len(months)<2: sys.exit(f"Need two months in RAW PI DATA; found {months}")
OLD,NEW=months[-2],months[-1]
print(f"Validating {NEW} forecast vs actual, baseline {OLD}  ({wbpath})\n")
by={}
for x in recs: by.setdefault((x['shop'],x['brand']),{})[x['month']]=x

# realized g per bond — MIRRORS build_pi_analysis._realized_g exactly (dominant max-RL_NEW line per shop)
dom={}
for (shop,brand),mm in by.items():
    if OLD not in mm or NEW not in mm: continue
    m,j=mm[OLD],mm[NEW]
    Jc=j['J'] or 0
    if j['rl']<=0 or m['rl']<=0 or j['rate']<=0 or j['p3']<=0 or m['p3']<=0 or Jc<=0: continue
    Aj=0.5*j['rate']*Jc/j['rl']                       # A_NEW
    Am=0.5*j['rate']*(m['p3']/j['p3'])*Jc/m['rl']     # A_OLD (scale NEW-window rate to OLD by bottle ratio)
    if shop not in dom or j['rl']>dom[shop][0]: dom[shop]=(j['rl'],Am,Aj,j['bond'])
bs=collections.defaultdict(lambda:[0.0,0.0])
for shop,(rl,Aold,Anew,bd) in dom.items(): bs[bd][0]+=Aold; bs[bd][1]+=Anew
G={bd:(v[1]/v[0]-1 if v[0]>0 else 0.0) for bd,v in bs.items()}

# forward-model MQ prediction
errs=[]; within1=0; tot=0
for (shop,brand),mm in by.items():
    if OLD not in mm or NEW not in mm: continue
    o,n=mm[OLD],mm[NEW]
    if o['rl']<=0 or o['p3']<=0 or n['p3']<=0: continue
    g=G.get(o['bond'],0.0)
    W=o['mq']-2.08*o['rl']
    pred=2.08*o['rl']*(n['p3']/o['p3'])/(1+g)+W
    e=pred-n['mq']; errs.append(abs(e)); tot+=1; within1+= (abs(e)<=1)
print("CHECK 1 — FORWARD-MODEL MQ PREDICTION vs KSBC actual:")
print(f"  lines tested: {tot}")
print(f"  within +-1 case: {within1/tot*100:.1f}%   median abs err: {st.median(errs):.2f} cs   mean: {sum(errs)/tot:.2f} cs")
print("\nCHECK 2 — REALIZED MARKET GROWTH per bond (A_{}->A_{}):".format(OLD,NEW))
for bd in sorted(G): print(f"  {bd:16s} {G[bd]*100:+6.1f}%")
print("\nInterpretation: high within-+-1% in CHECK 1 = the planner's MQ math holds on the new month too;")
print("CHECK 2 g vs the value the planner seeded last month = whether the growth assumption materialized.")
