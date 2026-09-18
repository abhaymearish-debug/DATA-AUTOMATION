#!/usr/bin/env python3
"""MEETING ANALYSIS pipeline - stage 1: discover sources, prep base (visit cols + staff), compute bond_data.json.
Usage: python3 meeting_build_data.py [--month JUNE] [--root <Claude folder>]
Outputs (to outputs dir given by $MEETING_OUT or cwd): meeting_base.xlsx, bond_data.json
"""
import openpyxl, json, re, csv, sys, os, shutil, glob
from copy import copy as _copy
from collections import defaultdict, Counter
from openpyxl.styles import Font
MONTHS=['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER']
SHORT={m:m.capitalize()[:3] for m in MONTHS}
BONDS=['KOLLAM','KOZHIKODE','ATTINGAL','PALAKKAD','KOTTARAKARA','KANNUR','ALAPPUZHA','NEDUMANGAD','ALUVA','PERINTHALMANNA','THODUPUZHA','PATHANAMTHITTA','KOTTAYAM','THRISSUR','TRIPUNITHURA']
VISIT_ALIAS={111014:11014}  # MUKKAM legacy 6-digit region code <-> 5-digit visit-file code
# ASMs (Area Sales Managers) -- their field hours/visits are NOT counted; only Sales Executives (Abhay 17 Jun 2026).
def _nm(x): return re.sub(r'[^A-Z0-9]','',str(x).upper())
ASM_ROSTER={'DINESHKUMARP','HARIDASANM','SOJANTT'}
def is_asm(name): return bool(name) and _nm(name) in ASM_ROSTER
def num(x):
    try: return float(x)
    except: return 0.0
def arg(name,default=None):
    if name in sys.argv: return sys.argv[sys.argv.index(name)+1]
    return default
def _detect_root():
    import glob as _g
    cands=sorted(_g.glob('/sessions/*/mnt/Claude'))
    withm=[p for p in cands if os.path.exists(os.path.join(p,'MASTER DATA CONFIRMED.xlsx'))]
    for p in withm:  # prefer a session that also has a sibling outputs dir (keeps ROOT/OUT in one session)
        if os.path.isdir(os.path.join(os.path.dirname(p),'outputs')): return p
    return withm[0] if withm else os.getcwd()
ROOT=arg('--root', os.environ.get('MEETING_ROOT') or _detect_root())
def _detect_out():
    sib=os.path.join(os.path.dirname(ROOT),'outputs')
    if os.path.isdir(sib): return sib
    import glob as _g
    g=sorted(_g.glob('/sessions/*/mnt/outputs'))
    return g[0] if g else os.getcwd()
OUT=os.environ.get('MEETING_OUT') or _detect_out()
YEAR=arg('--year', os.environ.get('MEETING_YEAR') or '2026')   # overridable; fixes cross-year (P2)

def find_ksbc(month):
    d=os.path.join(ROOT,'KSBC shop sales')
    full=os.path.join(d,f'{month} SHOP SALES ANALYSIS.xlsx')
    if os.path.exists(full): return full,'full'
    mids=sorted(glob.glob(os.path.join(d,f'{month} 1st - *th ANALYSIS.xlsx'))+glob.glob(os.path.join(d,f'{month} 1st - *st ANALYSIS.xlsx'))+glob.glob(os.path.join(d,f'{month} 1st - *nd ANALYSIS.xlsx'))+glob.glob(os.path.join(d,f'{month} 1st - *rd ANALYSIS.xlsx')), key=os.path.getmtime)
    if mids: return mids[-1],'mid'
    return None,None
def find_sec(month):
    d=os.path.join(ROOT,'Secondary sales')
    full=os.path.join(d,f'{month} SECONDARY SALES ANALYSIS.xlsx')
    if os.path.exists(full): return full
    mids=sorted(glob.glob(os.path.join(d,f'{month} 1st - * SECONDARY SALES ANALYSIS.xlsx')), key=os.path.getmtime)
    return mids[-1] if mids else None

# determine target month
tm=arg('--month')
if tm: target=tm.upper()
else:
    avail=[m for m in MONTHS if find_ksbc(m)[0] and find_sec(m)]   # P1: need BOTH KSBC and secondary
    target=avail[-1] if avail else 'MAY'
    ksbc_only=[m for m in MONTHS if find_ksbc(m)[0] and not find_sec(m)]
    if ksbc_only and (not avail or MONTHS.index(ksbc_only[-1])>MONTHS.index(target)):
        print(f'WARN: {ksbc_only[-1]} has KSBC but no secondary analysis yet -- building {target}; pass --month {ksbc_only[-1]} to force')
ti=MONTHS.index(target); prev=MONTHS[ti-1]; prev_year=str(int(YEAR)-1) if ti==0 else YEAR   # P2 cross-year prev
ksbc_cur,kind=find_ksbc(target); ksbc_prev,_=find_ksbc(prev)
sec_cur=find_sec(target); sec_prev=find_sec(prev)
master=os.path.join(ROOT,'MASTER DATA CONFIRMED.xlsx')
# Visit data source = "Field visit analysis" folder: FIELD VISIT ANALYSIS -<MONTH>.xlsx (renamed 2 Jul 2026; legacy pattern fallback below).
FVDIR=os.path.join(ROOT,'Field visit analysis')
visit_sum=None
for base_dir in (FVDIR, ROOT):
    # naming locked 2 Jul 2026: FIELD VISIT ANALYSIS -<MONTH>.xlsx (legacy EMPLOYEE-WISE pattern kept as fallback)
    for pat in (f'FIELD VISIT ANALYSIS -{target}.xlsx', f'*CLIENT VISIT ANALYSIS - {target} {YEAR}*FULL MONTH*.xlsx'):
        for cand in glob.glob(os.path.join(base_dir,pat)):
            if not os.path.basename(cand).startswith('~$'): visit_sum=cand
        if visit_sum: break
    if visit_sum: break
# raw log (for Active Field Days) - only if present in the folder; else degrade gracefully
visit_raw=None
for base_dir in (FVDIR, ROOT):
    for cand in glob.glob(os.path.join(base_dir,'*.xlsx')):
        b=os.path.basename(cand)
        if b.startswith('~$') or 'ANALYSIS' in b.upper(): continue
        if target.lower() not in b.lower(): continue   # require target month in filename (don't fall back to another month's log)
        try:
            w=openpyxl.load_workbook(cand,read_only=True); first=w.sheetnames[0] if w.sheetnames else ''; w.close()
        except: first=''
        if 'Client Visit' in first: visit_raw=cand
    if visit_raw: break
hist=os.path.join(ROOT,'Warehouse stock','_history','stock_history.csv')
print(f'TARGET={target} ({kind})  ksbc_cur={os.path.basename(ksbc_cur or "?")}')
print(f'  sec_cur={os.path.basename(sec_cur or "MISSING")}  visit_raw={"yes" if (visit_raw and os.path.exists(visit_raw)) else "NONE"}  visit_sum={"yes" if visit_sum else "NONE"}')
if not ksbc_cur or not sec_cur:
    print('FATAL: missing KSBC or secondary source'); sys.exit(1)

# ---------- master maps ----------
wb=openpyxl.load_workbook(master, read_only=True, data_only=True); ws=wb['16-4-25']
code_staff={}; code_bond={}; code_status={}; code_cat={}; code_name={}; closed_ksbc=set(); bond_staff_cnt=defaultdict(Counter)
for r in ws.iter_rows(min_row=3,values_only=True):
    if r[3] is None: continue
    try: c=int(r[3])
    except: continue
    code_staff[c]=r[6]; code_bond[c]=r[7]; code_status[c]=r[8]; code_cat[c]=(str(r[5]).strip().upper() if r[5] else ''); code_name[c]=(str(r[4]).strip() if r[4] else '')
    if r[5]=='KSBC' and r[8]=='Closed': closed_ksbc.add(c)   # D4: closed shops pruned from counts
    # bond Sales Executive = most-common KSBC active staff, EXCLUDING ASMs (D1)
    if r[7] in BONDS and r[5]=='KSBC' and r[6] and r[8]!='Closed' and not is_asm(r[6]): bond_staff_cnt[r[7]][r[6]]+=1
wb.close()
bond_staff={b:(bond_staff_cnt[b].most_common(1)[0][0] if bond_staff_cnt[b] else None) for b in BONDS}

# ---------- prep base: copy ksbc, add visit cols, fix staff ----------
base=os.path.join(OUT,'meeting_base.xlsx')
shutil.copy(ksbc_cur, base)
wb=openpyxl.load_workbook(base)
# visit per-shop from raw log
def dsec(s):
    try: h,m,sec=str(s).split(':'); return int(h)*3600+int(m)*60+int(sec)
    except: return None
# Per-shop visits from the FULL MONTH per-executive sheets (keyed by 6-digit shop code; matches locked workbook).
# Per-shop visits = the ASSIGNED executive's count only (a shop may appear in other execs' sheets too).
# Per-shop visits: each per-exec sheet lists shops with that executive's visit counts. Sum non-ASM sheets by
# shop code (reproducible; no current-master re-derivation, D2). Also track which non-ASM exec actually covered
# each bond (max visits) so FIELD COVERAGE shows the real coverer + matching hours, not a stale master assignment.
vmap=defaultdict(lambda:[0,0]); bond_exec_v=defaultdict(lambda:defaultdict(int))
if visit_sum:
    vw=openpyxl.load_workbook(visit_sum, read_only=True, data_only=True)
    for sh in vw.sheetnames:
        if sh=='SUMMARY' or is_asm(sh): continue   # skip SUMMARY + ASM executives (D1)
        vs=vw[sh]
        for r in vs.iter_rows(min_row=6,values_only=True):
            if r[0] is None: continue
            try: code=int(r[0])
            except: continue
            bd=r[3] if len(r)>3 else None
            if bd not in BONDS: continue   # drop CFD/warehouse rows
            try: vct=int(r[5])
            except: vct=0
            vmap[code][0]+=vct; vmap[code][1]+=(dsec(r[7]) or 0)
            if vct>0: bond_exec_v[bd][sh]+=vct
    vw.close()
vmap={k:(v[0],v[1]) for k,v in vmap.items()}
bond_cover={b:(max(ex,key=ex.get) if ex else None) for b,ex in bond_exec_v.items()}   # actual coverer per bond (D2)
def fmt(s): s=int(round(s)); return f'{s//3600}:{(s%3600)//60:02d}:{s%60:02d}'
present=[b for b in BONDS if b in wb.sheetnames]
def _clone(dst,src,numfmt=None):
    dst.font=_copy(src.font); dst.fill=_copy(src.fill); dst.border=_copy(src.border); dst.alignment=_copy(src.alignment)
    if numfmt: dst.number_format=numfmt
for b in present:
    ws=wb[b]
    tot=ws.max_row
    for rr in range(5,ws.max_row+1):
        if str(ws.cell(row=rr,column=1).value).strip().upper()=='TOTAL': tot=rr; break
    # extend title-row merges (rows 1-3) to col 12 so the band covers K/L
    # 16 Jul 2026: KSBC region row 2 is a SPLIT hero (A2:E2 'All figures in cases.' +
    # F2:J2 'Field Staff: ...').  Widening only the col-1 merge left F2:J2 in place and
    # produced an OVERLAPPING merge (invalid OOXML; crashed slim_region's unmerge).
    # Drop every top merge first, then re-merge the col-1-anchored bands out to col 12.
    _bands=set()
    for m in list(ws.merged_cells.ranges):
        if m.min_row<=3:
            if m.min_col==1: _bands.add(m.min_row)
            ws.unmerge_cells(str(m))
    for _rr in sorted(_bands):
        ws.merge_cells(start_row=_rr,start_column=1,end_row=_rr,end_column=12)
    jh=ws.cell(row=4,column=10)  # Rating header style
    for col,label in ((11,'No. of Visits'),(12,'Avg Time / Visit')):
        c=ws.cell(row=4,column=col,value=label); _clone(c,jh)
    alld=[]
    for rr in range(5,tot):
        v=ws.cell(row=rr,column=1).value
        try: code=int(v)
        except: continue
        srcI=ws.cell(row=rr,column=9)
        vct,tt=(vmap.get(code) or vmap.get(VISIT_ALIAS.get(code)) or (0,0))
        kc=ws.cell(row=rr,column=11,value=vct); _clone(kc,srcI,'0')
        lc=ws.cell(row=rr,column=12,value=(fmt(tt/vct) if vct>0 else '—')); _clone(lc,srcI)
        alld.append((vct,tt))
        mas=code_staff.get(code)
        if mas and str(ws.cell(row=rr,column=3).value).strip()!=str(mas).strip(): ws.cell(row=rr,column=3,value=mas)
    jt=ws.cell(row=tot,column=10)
    kt=ws.cell(row=tot,column=11,value=f'=SUM(K5:K{tot-1})'); _clone(kt,jt,'0')
    sv=sum(a for a,_ in alld); st=sum(bb for _,bb in alld)
    lt=ws.cell(row=tot,column=12,value=(fmt(st/sv) if sv>0 else '—')); _clone(lt,jt)
    for _c in (kt,lt):  # F2: visit totals aren't a rating -- white text, not col J tier colour
        _ff=_c.font; _c.font=Font(name=_ff.name,size=_ff.size,bold=_ff.bold,color='FFFFFFFF')
    ws.column_dimensions['K'].width=14; ws.column_dimensions['L'].width=16
    # DETAIL staff
    dn=b+' DETAIL'
    if dn in wb.sheetnames:
        ds=wb[dn]
        for rr in range(1,ds.max_row+1):
            a=ds.cell(row=rr,column=1).value
            if a and isinstance(a,str) and 'Staff:' in a:
                m=re.match(r'\s*(\d+)',a)
                if m:
                    mas=code_staff.get(int(m.group(1)))
                    if mas: ds.cell(row=rr,column=1,value=re.sub(r'Staff:\s*.*$',f'Staff: {mas}',a))
wb.save(base)
print('base prepped:', base)

# ---------- compute bond_data ----------
def find_combined(path):
    w=openpyxl.load_workbook(path, read_only=True)
    nm=[s for s in w.sheetnames if 'COMBINED' in s.upper() and 'DISPATCH' not in s.upper()]
    w.close()
    if not nm: return None
    def endday(s):
        m=re.findall(r'(\d+)', s); return int(m[-1]) if m else 0
    return max(nm, key=endday)

wb=openpyxl.load_workbook(base, read_only=True, data_only=True)
tert_cur={}; bond_shops={}; shop_rows={}
for b in present:
    ws=wb[b]; shops=[]; codes=set(); O=R=Sl=Cl=V=0.0; nv=0
    for r in ws.iter_rows(min_row=5,values_only=True):
        if r[0] is None or str(r[0]).strip().upper()=='TOTAL': continue
        try: code=int(r[0])
        except: continue
        if code in closed_ksbc: continue   # D4: drop master-Closed shops (e.g. 101042 KALLAMBALAM) -> active 283
        o,rc,s,c=num(r[3]),num(r[4]),num(r[5]),num(r[6]); vis=num(r[10]) if len(r)>10 else 0
        shops.append(dict(code=code,name=r[1],staff=r[2],O=o,R=rc,S=s,C=c,vis=vis))
        codes.add(code); O+=o;R+=rc;Sl+=s;Cl+=c;V+=vis
        if vis>0: nv+=1
    tert_cur[b]=dict(O=O,R=R,S=Sl,C=Cl,st=(Sl/(O+R) if (O+R)>0 else 0),nshops=len(shops),visits=V,visited=nv)
    bond_shops[b]=codes; shop_rows[b]=shops
comb_name=find_combined(base)
wb.close()
if comb_name is None:
    print('FATAL: no COMBINED roll-up sheet in KSBC base'); sys.exit(2)

def bondperf_regions(path):
    """Prev-month per-bond figures REBUCKETED BY CURRENT MASTER bond (Abhay, 2 Jul 2026).
    The prev workbook's own region placement can lag master (e.g. 101011 POWERHOUSE ROAD sat in
    ATTINGAL's MAY sheet but is a NEDUMANGAD shop in master), making the prev column non-comparable
    with the current month's bond composition and inconsistent with the Active-shops count.
    Read every region sheet's shop rows from the prev workbook, assign each shop to its CURRENT
    master bond (fallback: the sheet it appeared on; MUKKAM alias handled), sum O/R/S/C + nshops.
    Network totals are invariant under the re-bucket (every shop still lands in exactly one bond)."""
    w=openpyxl.load_workbook(path, read_only=True, data_only=True)
    agg={b:dict(O=0.0,R=0.0,S=0.0,C=0.0,nshops=0) for b in BONDS}
    for b in BONDS:
        if b not in w.sheetnames: continue
        for r in w[b].iter_rows(min_row=5,values_only=True):
            if r[0] is None or str(r[0]).strip().upper()=='TOTAL': continue
            try: code=int(r[0])
            except: continue
            mb=code_bond.get(code) or code_bond.get(VISIT_ALIAS.get(code))
            a=agg[mb if mb in BONDS else b]
            a['O']+=num(r[3]); a['R']+=num(r[4]); a['S']+=num(r[5]); a['C']+=num(r[6]); a['nshops']+=1
    w.close()
    for a in agg.values():
        a['st']=a['S']/(a['O']+a['R']) if (a['O']+a['R'])>0 else 0
    return {b:a for b,a in agg.items() if a['nshops']>0}
tert_prev=bondperf_regions(ksbc_prev) if ksbc_prev else {}
def secperf(path):
    if not path: return {}
    w=openpyxl.load_workbook(path, read_only=True, data_only=True); ws=w['BOND PERFORMANCE']; d={}
    for r in ws.iter_rows(min_row=3,values_only=True):
        if r[0] in BONDS: d[r[0]]=dict(ksbc=num(r[1]),fed=num(r[2]),bar=num(r[3]),total=num(r[4]))
    w.close(); return d
sec_curd=secperf(sec_cur); sec_prevd=secperf(sec_prev)

code2bond={}
for b in present:
    for c in bond_shops[b]: code2bond[c]=b
for _big,_small in VISIT_ALIAS.items():
    if _big in code2bond: code2bond[_small]=code2bond[_big]  # MUKKAM 111014 -> 11014 (COMBINED uses 5-digit)
def norm(s): return str(s).strip().upper().replace("B.C.B","BCB").replace("MORNING WALKER'S","MORNING WALKERS").replace("WALKER'S","WALKERS")
# tertiary brand/pack from base COMBINED
tbs=defaultdict(lambda:defaultdict(float)); tbc=defaultdict(lambda:defaultdict(float)); tbp=defaultdict(lambda:defaultdict(lambda:defaultdict(float))); tps=defaultdict(lambda:defaultdict(float))
wb=openpyxl.load_workbook(base, read_only=True, data_only=True); ws=wb[comb_name]
hdr=[c.value for c in next(ws.iter_rows(min_row=1,max_row=1))]
ci={str(h).strip():i for i,h in enumerate(hdr)}
def col(*names):
    for n in names:
        if n in ci: return ci[n]
    return None
c_code=col('Shop Code'); c_brand=col('Brand Name'); c_pack=col('Packing','Pack'); c_sales=col('Sales (Cases)','Sales'); c_close=col('Closing (Cases)','Closing')
if c_code is None or c_brand is None or c_sales is None:
    print('FATAL: COMBINED header columns missing (Shop Code/Brand Name/Sales) -- would silently zero brand mix'); sys.exit(2)
for r in ws.iter_rows(min_row=2,values_only=True):
    try: code=int(r[c_code])
    except: continue
    b=code2bond.get(code)
    if not b: continue
    br=norm(r[c_brand]); pk=r[c_pack]; sa=num(r[c_sales]); cl=num(r[c_close])
    tbs[b][br]+=sa; tbc[b][br]+=cl; tbp[b][br][pk]+=sa; tps[b][pk]+=sa
wb.close()
# invoice brand/pack from secondary COMBINED DISPATCHES
ibs=defaultdict(lambda:defaultdict(float)); ibp=defaultdict(lambda:defaultdict(lambda:defaultdict(float))); ips=defaultdict(lambda:defaultdict(float))
fed_cur=defaultdict(float); fed_prev=defaultdict(float)   # per-outlet Consumer-fed dispatches, cur/prev month (Abhay 2 Jul 2026)
bond_wh=defaultdict(set)
def canon_wh(nm):
    m=re.search(r'WH-([A-Z .]+?)\s+[A-Z]?FL',str(nm)); return m.group(1).strip() if m else None
wb=openpyxl.load_workbook(sec_cur, read_only=True, data_only=True)
disp=[s for s in wb.sheetnames if 'COMBINED DISPATCH' in s.upper()]
sec_outlets=defaultdict(list)
def _cs(pack,cases,bottles):   # D5: fold loose Issue Bottles into cases by pack BPC
    cs=num(cases); bo=num(bottles)
    if bo:
        m=re.search(r'(\d+)',str(pack)); bpc={'180':48,'375':24,'500':18,'750':12,'1000':9}.get(m.group(1)) if m else None
        if bpc: cs+=bo/bpc
    return cs
for b in present:
    if b in wb.sheetnames:
        for r in wb[b].iter_rows(min_row=5,values_only=True):
            if r[3] in ('FED','BAR'): sec_outlets[b].append((r[1],r[3],num(r[4]),r[2]))
if not disp:
    print('WARN: no COMBINED DISPATCHES sheet -- invoice brand/pack will be tertiary-only')
else:
    ws=wb[disp[-1]]
    for r in ws.iter_rows(min_row=2,values_only=True):
        lic=r[6]
        try: lici=int(lic)
        except: continue
        cat={'1':'KSBC','2':'FED','4':'BAR'}.get(str(lic)[0],'?'); b=code2bond.get(lici); bm=code_bond.get(lici)
        w=canon_wh(r[1])
        if (b or bm) and w: bond_wh[bm if bm in BONDS else b].add(w)
        if cat in ('FED','BAR') and bm in BONDS:
            br=norm(r[4]); pk=r[5]; cs=_cs(r[5], r[12], r[13] if len(r)>13 else 0); ibs[bm][br]+=cs; ibp[bm][br][pk]+=cs; ips[bm][pk]+=cs
            if cat=='FED': fed_cur[lici]+=cs
        elif cat=='FED' and bm not in BONDS:
            print(f'WARN: FED licensee {lici} not bond-mapped in master -- excluded from Consumer-fed roster')
wb.close()
if sec_prev:
    _w=openpyxl.load_workbook(sec_prev, read_only=True, data_only=True)
    _dsp=[x for x in _w.sheetnames if 'COMBINED DISPATCH' in x.upper()]
    if _dsp:
        for r in _w[_dsp[-1]].iter_rows(min_row=2,values_only=True):
            try: lici=int(r[6])
            except: continue
            if str(r[6])[0]=='2' and code_bond.get(lici) in BONDS:
                fed_prev[lici]+=_cs(r[5], r[12], r[13] if len(r)>13 else 0)
    _w.close()

# warehouse stock
wh_latest={}; wh_min={}
if os.path.exists(hist):
    H=list(csv.DictReader(open(hist)))
    monthnum=f'{YEAR}-{ti+1:02d}'
    mrows=[h for h in H if h['date'][:7]<=monthnum]
    if mrows:
        mend=max(h['date'] for h in mrows if h['date'][:7]==monthnum) if any(h['date'][:7]==monthnum for h in mrows) else max(h['date'] for h in mrows)
        for h in H:
            if h['date']==mend: wh_latest[h['warehouse']]=dict(phys=num(h['physical']),allot=num(h['allotable']),pend=num(h['pending']))
        for h in [x for x in H if x['date'][:7]==monthnum]:
            w=h['warehouse']; p=num(h['physical'])
            if w not in wh_min or p<wh_min[w]: wh_min[w]=p
hist_wh=set(wh_latest)
def match_wh(w):
    if w in hist_wh: return w
    for hw in hist_wh:
        if hw in w or w in hw: return hw
    return None

# visit insights
summ={}; vdays=defaultdict(set)
if visit_sum:
    w=openpyxl.load_workbook(visit_sum, read_only=True, data_only=True)
    if 'SUMMARY' in w.sheetnames:   # legacy workbooks (pre-2 Jul 2026)
        ws=w['SUMMARY']
        _hdr=[str(c.value).strip() if c.value else '' for c in next(ws.iter_rows(min_row=4,max_row=4))]
        _ix={h:i for i,h in enumerate(_hdr)}
        def _gi(*names,default=None):
            for n in names:
                if n in _ix: return _ix[n]
            return default
        i_tv=_gi('Total Visits'); i_at=_gi('Avg Time / Visit'); i_fh=_gi('Total Field Time (h)'); i_wh=_gi('Warehouse Stops'); i_ad=_gi('Active Field Days')
        for r in ws.iter_rows(min_row=5,values_only=True):
            if r[0] and r[0]!='TOTAL / NETWORK' and not is_asm(r[0]):   # D1: exclude ASMs -- SE hours only
                ad=r[i_ad] if (i_ad is not None and len(r)>i_ad) else None
                summ[r[0]]=dict(total_visits=(r[i_tv] if i_tv is not None else 0), avg_time=(str(r[i_at]) if i_at is not None else ''),
                    field_h=(r[i_fh] if i_fh is not None else 0), wh=(r[i_wh] if i_wh is not None else 0),
                    active_days=(ad if isinstance(ad,(int,float)) else None))
    else:   # v2 workbooks (2 Jul 2026+, SUMMARY removed): read each exec sheet's KPI tiles (row 4 labels / row 5 values)
        _SKIP={'OVERVIEW','ASM','DATA QUALITY','RAW VISIT LOG'}
        for sh in w.sheetnames:
            if sh in _SKIP or is_asm(sh): continue
            ws2=w[sh]
            top=list(ws2.iter_rows(min_row=1,max_row=5,values_only=True))
            if len(top)<5: continue
            labels=[str(x).strip() if x else '' for x in top[3]]; vals=top[4]
            _ix={h:i for i,h in enumerate(labels)}
            if 'TOTAL VISITS' not in _ix: continue
            def _g(name):
                i=_ix.get(name); return vals[i] if (i is not None and i<len(vals)) else None
            whn=0; inwh=False   # warehouse stops = visit count under the WAREHOUSE / NON-SHOP STOPS banner
            for r2 in ws2.iter_rows(min_row=6,values_only=True):
                a=str(r2[0]).strip() if r2[0] is not None else ''
                if a.startswith('WAREHOUSE / NON-SHOP'): inwh=True; continue
                if inwh:
                    if a=='Shop Code': continue
                    if r2[0] is None and (len(r2)<2 or r2[1] is None): break
                    v6=r2[5] if len(r2)>5 else None
                    if isinstance(v6,(int,float)): whn+=int(v6)
            emp=str(top[0][0]).strip() if top[0][0] else sh   # hero row carries the full name (sheet tab may be truncated)
            if is_asm(emp): continue
            ad=_g('ACTIVE DAYS')
            summ[emp]=dict(total_visits=_g('TOTAL VISITS') or 0, avg_time=str(_g('AVG TIME / VISIT') or ''),
                field_h=_g('FIELD HOURS') or 0, wh=whn,
                active_days=(ad if isinstance(ad,(int,float)) else None))
    w.close()
if visit_raw and os.path.exists(visit_raw):
    w=openpyxl.load_workbook(visit_raw, read_only=True, data_only=True); ws=w[w.sheetnames[0]]
    for r in ws.iter_rows(min_row=5,values_only=True):
        if r[0] and len(r)>2 and r[2] is not None: vdays[r[0]].add(str(r[2]))
    w.close()
def vmatch(staff):
    if not staff or staff=='VACANT': return None
    ns=str(staff).replace(' ','').replace('.','').upper()
    for k in summ:
        nk=k.replace(' ','').replace('.','').upper()
        if ns==nk or nk.startswith(ns) or ns.startswith(nk) or ns in nk or nk in ns: return k
    return None

out={'__meta__':dict(cur=f'{target} {YEAR}', prev=f'{prev} {prev_year}', curm=SHORT[target], prevm=SHORT[prev], target=target, kind=kind)}
for b in present:
    tmv=tert_cur[b]; whs=sorted(bond_wh.get(b,[])); whstock=[]; seen=set()
    for w in whs:
        mw=match_wh(w)
        if mw and mw not in seen and mw in wh_latest:
            seen.add(mw); l=wh_latest[mw]; whstock.append(dict(name=mw,phys=l['phys'],allot=l['allot'],pend=l['pend'],minmay=wh_min.get(mw,0)))
    # brand liquidation
    brs=set(tbs[b])|set(ibs[b]); liq_tot=0; rows=[]
    for br in brs:
        lq=tbs[b].get(br,0)+ibs[b].get(br,0)
        if lq<=0 and tbc[b].get(br,0)<=0: continue
        liq_tot+=lq; rows.append((br,lq,tbc[b].get(br,0)))
    blist=[]
    for br,lq,cl in sorted(rows,key=lambda x:-x[1]):
        pk=defaultdict(float)
        for p,v in tbp[b][br].items(): pk[p]+=v
        for p,v in ibp[b][br].items(): pk[p]+=v
        tp=max(pk.items(),key=lambda x:x[1])[0] if pk else ''
        blist.append(dict(brand=br,sales=lq,pct=(lq/liq_tot if liq_tot else 0),close=cl,toppack=tp))
    pset=set(tps[b])|set(ips[b]); plist=sorted([[p,tps[b].get(p,0)+ips[b].get(p,0)] for p in pset if (tps[b].get(p,0)+ips[b].get(p,0))>0], key=lambda x:-x[1])
    # visit insight
    staff=bond_cover.get(b) or bond_staff.get(b); key=vmatch(staff); vis=None   # prefer the real coverer (D2)
    if key and summ.get(key):
        sdat=summ[key]; ad=(int(sdat['active_days']) if sdat.get('active_days') else len(vdays.get(key,[]))); tv=sdat['total_visits'] or 0
        vis=dict(staff=staff,active_days=(ad if ad>0 else None),staff_total_visits=tv,avg_visits_day=((tv/ad) if ad>0 else None),avg_time=sdat['avg_time'],field_h=sdat['field_h'] or 0,wh_stops=sdat['wh'] or 0)
    fedo=[]
    for cc,cat in code_cat.items():
        if cat=='FED' and str(code_bond.get(cc,'')).strip()==b:
            mj=float(fed_cur.get(cc,0)); mm=float(fed_prev.get(cc,0))
            if str(code_status.get(cc,'')).strip().lower()=='active' or mj>0 or mm>0:
                fedo.append(dict(code=cc,name=code_name.get(cc,''),status=str(code_status.get(cc,'')).strip(),may=mm,jun=mj))
    fedo.sort(key=lambda x:(-x['jun'],-x['may'],x['name']))
    out[b]=dict(tert_may=tmv,tert_apr=tert_prev.get(b,{}),sec_may=sec_curd.get(b,{}),sec_apr=sec_prevd.get(b,{}),
        sec_outlets=sec_outlets.get(b,[]),brands=blist,packs=plist,whstock=whstock,vis=vis,fed_outlets=fedo)
# network Sales-Executive field totals (ASM-excluded, de-duplicated) -- D1/D2
out['__field__']=dict(field_h=round(sum(num(v['field_h']) for v in summ.values()),1),
    wh=int(sum(num(v['wh']) for v in summ.values())),
    total_visits=int(sum(num(v['total_visits']) for v in summ.values())),
    active=len(summ))
# P3: hard guards so a silently-wrong build aborts (set -e) before anything copies to live
if len(present)!=15:
    print(f'FATAL: expected 15 bond sheets, found {len(present)}: {present}'); sys.exit(2)
_tert=sum(out[b]['tert_may']['S'] for b in present); _sec=sum(out[b]['sec_may'].get('total',0) for b in present)
if _tert<=0 or _sec<=0:
    print(f'FATAL: zero tertiary ({_tert}) or secondary ({_sec}) total -- source/header drift'); sys.exit(2)
json.dump(out, open(os.path.join(OUT,'bond_data.json'),'w'), default=str, indent=1)
print('bond_data.json written. bonds:', len(present), '| tert total:', round(_tert,1), '| sec total:', round(_sec,1),
      '| SE field_h:', out['__field__']['field_h'], '| SE active:', out['__field__']['active'])
