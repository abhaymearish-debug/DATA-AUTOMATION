#!/usr/bin/env python3
"""Audit the 15 bond PDFs against the PI INSIGHTS workbook (cached Excel values)."""
import openpyxl, glob, os, re, subprocess
from collections import defaultdict
BASE="/sessions/great-serene-rubin/mnt/Claude"
OUTDIR=os.path.join(BASE,"PURCHASE INSTRUCTION","BOND PI DRILL-DOWNS")
WB=max([c for c in glob.glob(BASE+"/PURCHASE INSTRUCTION/PI INSIGHTS - * & * *.xlsx") if not os.path.basename(c).startswith('~$')], key=os.path.getmtime)
MONTHS=['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER']
m=re.search(r'PI INSIGHTS\s*-\s*([A-Za-z]+)\s*&\s*([A-Za-z]+)\s*(\d{4})', os.path.basename(WB))
PRIOR,CUR,YEAR=m.group(1),m.group(2),m.group(3)
i=MONTHS.index(CUR.upper()); AB=[MONTHS[(i-k)%12].title()[:3] for k in (3,2,1)]
wb=openpyxl.load_workbook(WB, data_only=True)
def txt(ws,r,c): v=ws.cell(r,c).value; return "" if v is None else str(v)
def pdftext(p): return subprocess.run(['pdftotext','-layout',p,'-'],capture_output=True,text=True).stdout
def toks(s): return re.findall(r'[+\-]?\d[\d,]*(?:\.\d+)?%?|·', s.replace('−','-'))
def fi(v):
    if v in (None,""): return '·'
    iv=int(round(float(v))); return '·' if iv==0 else f'{iv:,}'
def fd(v):
    if v in (None,""): return '·'
    iv=int(round(float(v))); return '·' if iv==0 else (f'+{iv:,}' if iv>0 else f'-{abs(iv):,}')
def fp(v):
    if v in (None,""): return '·'
    p=int(round(float(v)*100)); return '·' if p==0 else (f'+{p}%' if p>0 else f'-{abs(p)}%')
def markers(ws):
    BP=bt=co=None
    for r in range(1,ws.max_row+1):
        a=txt(ws,r,1)
        if BP is None and a.startswith("BRAND PERFORMANCE IN THIS BOND"): BP=r
        if BP and bt is None and r>BP and a.strip()=="TOTAL": bt=r
        if co is None and a.startswith("CEILING & OFFTAKE"): co=r
    last=0
    for r in range(1,ws.max_row+1):
        if any(txt(ws,r,c) for c in range(1,15)): last=r
    return BP,bt,co,last

bonds=[n for n in wb.sheetnames if txt(wb[n],1,1).endswith("PURCHASE INSTRUCTION DRILL-DOWN")]
G={'bc':0,'bb':0,'sc':0,'sb':0,'struct':0,'jul':0}
for name in bonds:
    ws=wb[name]; BP,bt,co,last=markers(ws)
    pdf=os.path.join(OUTDIR, f"{name} — PI DRILL-DOWN ({CUR.title()} vs {PRIOR.title()} {YEAR}).pdf")
    T=pdftext(pdf); lines=T.splitlines()
    # (1) BRAND TABLE positional cross-check PDF<->workbook
    bbad=[]; bc=0
    for r in list(range(BP+2,bt))+[bt]:
        nm=txt(ws,r,1)
        if not nm: continue
        exp=[fi(ws.cell(r,2).value),fi(ws.cell(r,3).value),fi(ws.cell(r,4).value),fi(ws.cell(r,5).value),
             fd(ws.cell(r,6).value),fi(ws.cell(r,7).value),fp(ws.cell(r,8).value)]
        if nm=='TOTAL': exp=exp[:6]
        ln=[l for l in lines if nm in l]
        if not ln: bbad.append(f"'{nm}' row missing"); continue
        got=[g for g in toks(ln[0].replace(nm,'')) if g!='·']; exp=[e for e in exp if e!='·']; bc+=len(exp)
        for e in exp:
            if e in got: continue
            mm=re.match(r'([+\-])(\d+)%',e); ok=mm and any(f'{mm.group(1)}{int(mm.group(2))+d}%' in got for d in(-1,0,1))
            if not ok: bbad.append(f"{nm}:{e}∉{got}")
    # (2) SHOP-BLOCK TOTAL rows PDF<->workbook (Rate,RL,RQ,MQjul,MQjun)
    wtot=[]  # (title,[rate,rl,rq,mqjul,mqjun])
    r=co
    while r<=last:
        if re.match(r'^\s*\d{3,5}\s*·', txt(ws,r,1)):
            title=txt(ws,r,1); tr=None
            for rr in range(r+1,last+1):
                if txt(ws,rr,1).strip()=='TOTAL': tr=rr; break
                if re.match(r'^\s*\d{3,5}\s*·', txt(ws,rr,1)): break
            if tr:
                wtot.append((title,[f"{float(ws.cell(tr,2).value or 0):.1f}",fi(ws.cell(tr,3).value),
                             fi(ws.cell(tr,4).value),fi(ws.cell(tr,5).value),fi(ws.cell(tr,6).value)]))
        r+=1
    # segment PDF by shop titles, find each block's TOTAL line
    idx=[(k,l) for k,l in enumerate(lines) if re.match(r'^\s*\d{3,5}\s*·',l)]
    sbad=[]; sc=0
    for bi,(title,exp) in enumerate(wtot):
        start=None
        for k,l in idx:
            if title.split('·')[0].strip() in l and title.split('·')[1].strip()[:6] in l: start=k; break
        if start is None: sbad.append(f"{title} block missing"); continue
        tl=None
        for k in range(start+1,min(start+40,len(lines))):
            if 'TOTAL' in lines[k]: tl=lines[k]; break
        if tl is None: sbad.append(f"{title} TOTAL missing"); continue
        got=toks(tl); expnz=[e for e in exp if e!='·']; sc+=len(expnz)
        for e in expnz:
            if e not in got: sbad.append(f"{title} {e}∉TOTAL")
    # (3) reconciliation (informational) + structural
    strows=[rr for rr in range(co+1,last+1) if txt(ws,rr,1).strip()=='TOTAL']
    smj=sum(float(ws.cell(rr,5).value or 0) for rr in strows); bmj=float(ws.cell(bt,4).value or 0)
    julok=abs(smj-bmj)<0.5
    st=[]
    if txt(ws,1,1) not in T: st.append("heading")
    if "3-Mo" not in T: st.append("3-Mo header")
    if f"{AB[0]}, {AB[1]} & {AB[2]} {YEAR}" not in T: st.append("footnote window")
    if "ACTION PLAN" in T: st.append("ACTION PLAN leaked")
    if "Share of MQ" in T: st.append("Share of MQ leaked")
    if "Why / Watch" in T: st.append("action cols leaked")
    G['bc']+=bc;G['bb']+=len(bbad);G['sc']+=sc;G['sb']+=len(sbad);G['struct']+=len(st);G['jul']+=0 if julok else 1
    ok=not bbad and not sbad and not st and julok
    print(f"[{'OK ' if ok else '!!!'}] {name:15} brand {bc-len(bbad)}/{bc} | shopTOTAL {sc-len(sbad)}/{sc} ({len(wtot)} shops) | MQ{CUR[:3]} recon {'✓' if julok else '✗'} | struct {'clean' if not st else st}")
    for x in (bbad+sbad)[:4]: print("      -",x)

print("\n==== GRAND ====")
print(f"brand cells {G['bc']-G['bb']}/{G['bc']} | shop-TOTAL cells {G['sc']-G['sb']}/{G['sc']} | MQ{CUR[:3]} recon {15-G['jul']}/15 | structural issues {G['struct']}")
print("RESULT:", "ALL PASS ✅" if G['bb']==0 and G['sb']==0 and G['struct']==0 and G['jul']==0 else "REVIEW")
