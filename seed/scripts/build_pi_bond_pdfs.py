#!/usr/bin/env python3
"""
Build per-bond PI DRILL-DOWN PDFs from the PI INSIGHTS workbook.  [LOCKED FORMAT — 7 Jul 2026, Abhay-approved]

Per-bond PDF layout (page 1 then shop pages):
  * Page-1 heading bar   : the sheet's A1 title ("<BOND> — PURCHASE INSTRUCTION DRILL-DOWN"),
                           navy fill / gold bold text / gold underline.
  * BRAND PERFORMANCE tbl : cut at "Offtake Δ%" (cols A:H). The Off-Indent / Off-Ind btl /
                           Share of MQ columns (I:K) are DROPPED. The "Offtake (cs)" header is
                           relabelled "Offtake 3-Mo (cs)" and an italic footnote states the
                           exact trailing window (M-3..M-1) that KSBC uses to set the CUR PI.
  * Shop pages           : the CEILING & OFFTAKE shop x brand tables at A:I only — the ACTION
                           PLAN block (HOLD/RECOVER/GROW/Action/Why-Watch, cols J:N) and the
                           market-growth control row are removed. Every navy shop-title banner
                           is shrunk from A:N to A:I so it ends flush at the Status column.

Everything is auto-detected per sheet (row positions differ by shop count), so this works for
every bond and every future month pair. Month labels/footnote are derived from the workbook name.

Usage:
  python3 build_pi_bond_pdfs.py [--workbook PATH] [--outdir DIR] [--bonds "KOLLAM,ATTINGAL"] [--base DIR]
Requires: LibreOffice (soffice) + python3-uno + openpyxl. Self-launches a headless soffice.
"""
import os, sys, re, time, glob, argparse, subprocess, shutil
import openpyxl
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.table import CellRangeAddress, BorderLine2
from com.sun.star.awt.FontWeight import BOLD
from com.sun.star.awt.FontSlant import ITALIC
from com.sun.star.table.CellHoriJustify import CENTER as HCENTER, LEFT as HLEFT
from com.sun.star.table.CellVertJustify import CENTER as VCENTER

MONTHS=['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER']
NAVY=0x0D1B4A; GOLD=0xFFB300; GREY=0x6B7280; PORT=2002
FONT="Aptos Narrow"

def win3(cur):
    i=MONTHS.index(cur.upper()); return [MONTHS[(i-k)%12] for k in (3,2,1)]

def parse_wb_name(path):
    m=re.search(r'PI INSIGHTS\s*-\s*([A-Za-z]+)\s*&\s*([A-Za-z]+)\s*(\d{4})', os.path.basename(path))
    if not m: raise SystemExit("cannot parse month pair from workbook name: "+os.path.basename(path))
    return m.group(1).upper(), m.group(2).upper(), m.group(3)          # PRIOR, CUR, YEAR

def newest_workbook(base):
    cand=glob.glob(os.path.join(base,'PURCHASE INSTRUCTION','PI INSIGHTS - * & * *.xlsx'))
    cand=[c for c in cand if not os.path.basename(c).startswith('~$')]
    if not cand: raise SystemExit("no PI INSIGHTS workbook found under "+base)
    return max(cand, key=os.path.getmtime)

def precompute(xlsx):
    wb=openpyxl.load_workbook(xlsx, data_only=True)
    out=[]
    for name in wb.sheetnames:
        ws=wb[name]
        a1=ws.cell(1,1).value or ""
        if not str(a1).endswith("PURCHASE INSTRUCTION DRILL-DOWN"): continue
        BP=bt=co=None
        for r in range(1, ws.max_row+1):
            a=ws.cell(r,1).value; a=str(a) if a is not None else ""
            if BP is None and a.startswith("BRAND PERFORMANCE IN THIS BOND"): BP=r
            if BP and bt is None and r>BP and a.strip()=="TOTAL": bt=r
            if co is None and a.startswith("CEILING & OFFTAKE"): co=r
        if not (BP and bt and co):
            print("  ! skipping %s (markers BP=%s bt=%s co=%s)"%(name,BP,bt,co)); continue
        titlerows=sorted(m.min_row for m in ws.merged_cells.ranges
                         if m.min_col==1 and m.max_col>9 and m.min_row==m.max_row and m.min_row>co)
        last=0
        for r in range(1, ws.max_row+1):
            if any(ws.cell(r,c).value not in (None,"") for c in range(1,15)): last=r
        # empty-PI (blackout-door) band note rows — get the 'no brands on indent' mention + wrap
        noterows=[r for r in range(co+1,ws.max_row+1) if "issued no KSD lines" in str(ws.cell(r,1).value or "")]
        # INCOMPLETE (one-month-only) band rows — reword + wrap (dev-instruction wording, overflows A:I)
        incrows=[r for r in range(co+1,ws.max_row+1)
                 if ("INCOMPLETE — PI RAW MISSING" in str(ws.cell(r,1).value or ""))
                 or ("PI MISSING" in str(ws.cell(r,1).value or "") and "→" in str(ws.cell(r,1).value or ""))]
        out.append(dict(name=name,a1=str(a1),BP=BP,bt=bt,co=co,last=last,titlerows=titlerows,noterows=noterows,incrows=incrows))
    return out

def mkp(n,v): p=PropertyValue(); p.Name=n; p.Value=v; return p

def apply_transform(doc, sheets, idx, b, foot):
    ks=sheets.getByIndex(idx)
    BP,bt,co,last=b['BP'],b['bt'],b['co'],b['last']; hr=BP-1; fr=bt+1
    # 1. BRAND PERFORMANCE banner  A:* -> A:H
    ks.getCellRangeByPosition(0,BP-1,13,BP-1).merge(False)
    ks.getCellRangeByPosition(0,BP-1,7,BP-1).merge(True)
    # 2. relabel Offtake header (col G = idx6, header row = BP+1 -> 0-based BP)
    ks.getCellByPosition(6,BP).setString("Offtake 3-Mo (cs)")
    # 3. page-1 heading (row BP-1)
    ks.getCellByPosition(0,hr-1).setString(b['a1'])
    ks.getCellRangeByPosition(0,hr-1,7,hr-1).merge(False)
    ks.getCellRangeByPosition(0,hr-1,7,hr-1).merge(True)
    h=ks.getCellRangeByPosition(0,hr-1,7,hr-1)
    h.CellBackColor=NAVY; h.CharColor=GOLD; h.CharWeight=BOLD; h.CharFontName=FONT
    h.HoriJustify=HCENTER; h.VertJustify=VCENTER
    # long bond names overflow the A:H width at 19pt -> auto-shrink so they never clip
    L=len(b['a1']); h.CharHeight=19.0 if L<=40 else max(13.0, 19.0*40.0/L)
    try: h.ShrinkToFit=True
    except Exception: pass
    gl=BorderLine2(); gl.Color=GOLD; gl.LineWidth=50; gl.LineStyle=0; h.BottomBorder=gl
    ks.Rows.getByIndex(hr-1).Height=1450
    # 4. footnote (row bt+1)
    ks.getCellByPosition(0,fr-1).setString(foot)
    ks.getCellRangeByPosition(0,fr-1,7,fr-1).merge(False)
    ks.getCellRangeByPosition(0,fr-1,7,fr-1).merge(True)
    f=ks.getCellRangeByPosition(0,fr-1,7,fr-1)
    f.CharColor=GREY; f.CharHeight=9.0; f.CharFontName=FONT; f.CharPosture=ITALIC
    f.HoriJustify=HLEFT; f.VertJustify=VCENTER; f.IsTextWrapped=True
    ks.Rows.getByIndex(fr-1).Height=650
    # 5. shrink shop-title banners A:N -> A:I
    for r in b['titlerows']:
        ks.getCellRangeByPosition(0,r-1,13,r-1).merge(False)
        ks.getCellRangeByPosition(0,r-1,8,r-1).merge(True)
    # 5b. empty-PI (blackout) band notes: simplify to just 'NO <MONTH> PI — no brands on indent'
    #     (drop the 'sold X btl / still selling / escalate to RM' tail — Abhay 7 Jul 2026)
    for r in b.get('noterows',[]):
        cell=ks.getCellByPosition(0,r-1)
        cell.setString(re.sub(r'^EMPTY (\w+) PI — .*', r'NO \1 PI — no brands on indent.',
                              cell.getString(), flags=re.DOTALL))
        rr=ks.getCellRangeByPosition(0,r-1,8,r-1)
        rr.IsTextWrapped=False; rr.HoriJustify=HLEFT; rr.VertJustify=VCENTER
        ks.Rows.getByIndex(r-1).Height=520
    # 5c. INCOMPLETE (one-month-only) band: drop dev-instruction wording + wrap so it fits A:I (else overflows)
    for r in b.get('incrows',[]):
        cell=ks.getCellByPosition(0,r-1); t=cell.getString()
        if "INCOMPLETE — PI RAW MISSING" in t:
            t="⚠ INCOMPLETE — shop has a PI for only one of the two months (no comparison shown below)."
        elif "→" in t:
            t=t.split("→")[0].strip()
        cell.setString(t)
        rr=ks.getCellRangeByPosition(0,r-1,8,r-1)
        rr.IsTextWrapped=True; rr.HoriJustify=HLEFT; rr.VertJustify=VCENTER
        ks.Rows.getByIndex(r-1).Height=620
    # 5d. ✅ (U+2705) has no glyph in the headless renderer -> swap to ✓ (DejaVu-safe) so the status never boxes
    repl=ks.createReplaceDescriptor(); repl.SearchString="✅"; repl.ReplaceString="✓"
    ks.replaceAll(repl)
    # 6. remember print ranges — APPLIED PER-EXPORT (only the target sheet may carry a print
    #    area at export time, else LibreOffice's PDF export bundles every print-area sheet).
    b['pr']=((0,hr-1,7,fr-1),(0,co-1,8,last-1))
    # 7. landscape fit-to-width
    try:
        ps=doc.StyleFamilies.getByName("PageStyles").getByName(ks.PageStyle)
        ps.ScaleToPagesX=1; ps.ScaleToPagesY=0; ps.IsLandscape=True
    except Exception as e:
        print("   page-style warn %s: %s"%(b['name'],e))

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('--workbook'); ap.add_argument('--outdir'); ap.add_argument('--bonds', default='')
    ap.add_argument('--base', default='/sessions/great-serene-rubin/mnt/Claude')
    A=ap.parse_args()
    wb_path=A.workbook or newest_workbook(A.base)
    PRIOR,CUR,YEAR=parse_wb_name(wb_path)
    ab=[m.title()[:3] for m in win3(CUR)]
    foot=("Offtake (cs) = past 3 months' tertiary sales — %s, %s & %s %s — "
          "the trailing window KSBC uses to set the %s PI."%(ab[0],ab[1],ab[2],YEAR,CUR.title()))
    outdir=A.outdir or os.path.join(A.base,'PURCHASE INSTRUCTION','BOND PI DRILL-DOWNS')
    os.makedirs(outdir, exist_ok=True)
    print("workbook:",os.path.basename(wb_path),"| window:",",".join(ab),YEAR,"| CUR PI:",CUR.title())

    bonds=precompute(wb_path)
    want=[x.strip().upper() for x in A.bonds.split(',') if x.strip()]
    if want: bonds=[b for b in bonds if b['name'].upper() in want]
    print("bonds to build:",", ".join(b['name'] for b in bonds))

    # work on a copy so the live (Excel-open) workbook is never touched
    tmp="/tmp/pi_bond_src.xlsx"; shutil.copy(wb_path, tmp)

    subprocess.call(['pkill','-9','-x','soffice.bin']); time.sleep(1)
    subprocess.call(['rm','-rf','/tmp/lo_bondpdf'])
    subprocess.Popen(['soffice','--headless','--invisible','--nologo','--nofirststartwizard','--norestore',
        '-env:UserInstallation=file:///tmp/lo_bondpdf',
        '--accept=socket,host=localhost,port=%d;urp;StarOffice.ComponentContext'%PORT],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(6)
    lc=uno.getComponentContext()
    res=lc.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", lc)
    ctx=None
    for _ in range(40):
        try: ctx=res.resolve("uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext"%PORT); break
        except Exception: time.sleep(1)
    if ctx is None: raise SystemExit("no soffice connect")
    smgr=ctx.ServiceManager
    desktop=smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc=None
    for _ in range(8):
        doc=desktop.loadComponentFromURL("file://"+tmp,"_blank",0,(mkp("Hidden",True),))
        if doc: break
        time.sleep(2)
    if not doc: raise SystemExit("workbook load fail")
    sheets=doc.Sheets; names=list(sheets.ElementNames)

    for b in bonds:
        apply_transform(doc, sheets, names.index(b['name']), b, foot)

    def setpa(idx, prs):
        ks=sheets.getByIndex(idx)
        if not prs: ks.setPrintAreas(()); return
        addrs=[]
        for (sc,sr,ec,er) in prs:
            a=CellRangeAddress(); a.Sheet=idx; a.StartColumn=sc; a.StartRow=sr; a.EndColumn=ec; a.EndRow=er; addrs.append(a)
        ks.setPrintAreas(tuple(addrs))

    tmpout="/tmp/bondpdfs"; os.makedirs(tmpout, exist_ok=True)  # write here first (keeps LO temp/lock churn out of the deliverable folder)
    written=[]
    for b in bonds:
        # isolate: only the target bond carries a print area, and only it is visible
        for b2 in bonds: setpa(names.index(b2['name']), b['pr'] if b2['name']==b['name'] else None)
        for nm in names: sheets.getByName(nm).IsVisible=(nm==b['name'])
        fn="%s — PI DRILL-DOWN (%s vs %s %s).pdf"%(b['name'],CUR.title(),PRIOR.title(),YEAR)
        tp=os.path.join(tmpout,fn)
        doc.storeToURL("file://"+tp,(mkp("FilterName","calc_pdf_Export"),))
        shutil.copy(tp, os.path.join(outdir,fn))
        written.append(fn); print("WROTE", fn)
    doc.close(False)
    subprocess.call(['pkill','-9','-x','soffice.bin'])
    print("DONE %d PDF(s) -> %s"%(len(written), outdir))

if __name__=="__main__": main()
