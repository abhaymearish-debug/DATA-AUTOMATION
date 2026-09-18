"""Shared helpers for the marketing-head incentive workbook."""
import re, zipfile, subprocess, shutil, os, sys
from collections import defaultdict
from lxml import etree

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

# Incentive sheet column -> brand key. These 8 columns are the ONLY data cells
# in the workbook; everything else (H,I,L,M,N,Q,R,S,V,W,Z,AA,AB,AC) is a formula.
BRAND_COLS = {'G':'KS99', 'J':'OPR', 'K':'BLENDERS', 'O':'ROR',
              'P':'BCB',  'T':'MBR', 'U':'MWB',      'Y':'CCB'}

# Brand-name normaliser. Sources spell the same brand several ways across months
# (e.g. "MORNING WALKERS" vs "MORNING WALKER'S"), so match on a stripped prefix
# rather than the literal string, or real volume silently vanishes.
_BRAND_PREFIXES = [('BCBNO','BCB'), ('BLENDERSCHOICE','BLENDERS'), ('CHAIRMANSCHOICE','CCB'),
                   ('KS99','KS99'), ('MAGICBLENDRESERVED','MBR'), ('MORNINGWALKER','MWB'),
                   ('OLDPEARL','OPR'), ('ROYALOLDFORT','ROR')]

def norm_brand(raw):
    """Return the canonical brand key, or None for a non-KSD / discontinued SKU."""
    s = re.sub(r'[^A-Z0-9]', '', str(raw).upper())
    for prefix, key in _BRAND_PREFIXES:
        if s.startswith(prefix):
            return key
    return None

BOTTLES_PER_CASE = {'180':48, '375':24, '500':18, '750':12, '1000':9}

def bottles_per_case(pack):
    m = re.search(r'(\d+)', str(pack))
    return BOTTLES_PER_CASE.get(m.group(1)) if m else None

def pick_sheet(wb, explicit=None):
    """Choose the sheet to roll forward.

    The workbook keeps one sheet per month, so the newest is the one to overwrite.
    The header date in B2 says which month a sheet holds -- far more reliable than
    row counts, since an older sheet can easily carry more rows than a newer one.
    """
    if explicit:
        return wb[explicit]
    import datetime as _dt
    best, best_key = None, None
    for s in wb.worksheets:
        n = sum(1 for r in range(4, 500) if s.cell(r, 2).value)
        if n == 0:
            continue
        d = s['B2'].value
        key = (d if isinstance(d, _dt.datetime) else _dt.datetime.min, n)
        if best_key is None or key > best_key:
            best, best_key = s, key
    if best is None:
        raise SystemExit('ERROR: no sheet in this workbook carries shop rows')
    return best


def match_name(name, lookup):
    """Workbook labels drift in case and spacing between months, so fall back
    progressively rather than dropping a shop that is really there."""
    n = str(name).strip()
    if n in lookup:
        return n
    fold = {k.strip().casefold(): k for k in lookup}
    if n.casefold() in fold:
        return fold[n.casefold()]
    squash = {re.sub(r'\s+', ' ', k).strip().casefold(): k for k in lookup}
    key = re.sub(r'\s+', ' ', n).casefold()
    return squash.get(key)


def read_liquidation(path):
    """TOTAL LIQUIDATION -<MONTH>.xlsx -> {shop name: {...}}.

    Column order in that workbook is alphabetical by brand, which is NOT the
    incentive sheet's segment order, so map by header text rather than position.
    """
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True)['Brand Sales by Shop']
    hdr = {}
    for c in range(6, 14):
        key = norm_brand(ws.cell(1, c).value)
        if key:
            hdr[c] = key
    if len(hdr) != 8:
        raise SystemExit(f'ERROR: expected 8 brand columns in {path}, resolved {len(hdr)}')
    out = {}
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 2).value
        if name in (None, ''):
            continue
        out[str(name).strip()] = dict(
            code=str(ws.cell(r, 1).value).strip(),
            bond=ws.cell(r, 4).value,
            type=str(ws.cell(r, 5).value).strip().upper(),
            vals={k: float(ws.cell(r, c).value or 0) for c, k in hdr.items()})
    return out


def sheet_paths(z):
    """Map sheet display name -> xml path inside the xlsx zip."""
    rels = {r.get('Id'): r.get('Target')
            for r in etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    out = {}
    for sh in etree.fromstring(z.read('xl/workbook.xml')).iter(NS + 'sheet'):
        rid = sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        t = rels[rid].lstrip('/')
        # openpyxl writes relative targets ("worksheets/sheet1.xml") while
        # LibreOffice writes absolute ones ("/xl/worksheets/sheet1.xml"), and we
        # read both, so normalise instead of assuming either shape.
        out[sh.get('name')] = t if t.startswith('xl/') else 'xl/' + t
    return out


def recalculate(path, timeout=300):
    """Recalculate via LibreOffice into a temp copy and return its path.

    Never hand the LibreOffice output to the user directly: it rewrites accounting
    number formats and rounds column widths, which shows up as thousands of
    cosmetic diffs against the original workbook. Use it purely as a value oracle.
    """
    work = '/tmp/_inc_recalc'
    shutil.rmtree(work, ignore_errors=True)
    outdir = os.path.join(work, 'out')
    os.makedirs(outdir)
    src = os.path.join(work, 'src.xlsx')
    shutil.copy(path, src)
    # Convert into a SEPARATE directory. Pointing --outdir at the source folder
    # makes LibreOffice write over its own input, and since the file then still
    # exists an existence check passes while the "recalculated" copy is really
    # just the original, with no cached values at all.
    profile = '/tmp/_inc_lo_profile'
    shutil.rmtree(profile, ignore_errors=True)
    r = subprocess.run(['soffice', f'-env:UserInstallation=file://{profile}',
                        '--headless', '--norestore', '--invisible', '--nolockcheck',
                        '--convert-to', 'xlsx:Calc MS Excel 2007 XML',
                        '--outdir', outdir, src],
                       capture_output=True, timeout=timeout, text=True)
    out = os.path.join(outdir, 'src.xlsx')
    if not os.path.exists(out):
        raise SystemExit('ERROR: LibreOffice recalculation produced no output\n'
                         f'  exit {r.returncode}\n  stdout: {r.stdout.strip()}\n'
                         f'  stderr: {r.stderr.strip()}')
    return out


def inject_cached_values(formatted_path, valued_path, out_path):
    """Copy computed results out of the recalculated file into the openpyxl file.

    openpyxl keeps the original formatting byte-for-byte but drops every cached
    formula result, so the workbook would look blank in any viewer that does not
    calculate (Quick Look, phone, Google preview). This grafts the values back on
    without touching a single style, which is what lets us keep both.
    """
    zf, zv = zipfile.ZipFile(formatted_path), zipfile.ZipFile(valued_path)
    ps, pv = sheet_paths(zf), sheet_paths(zv)
    sst = []
    if 'xl/sharedStrings.xml' in zv.namelist():
        for si in etree.fromstring(zv.read('xl/sharedStrings.xml')).iter(NS + 'si'):
            sst.append(''.join(t.text or '' for t in si.iter(NS + 't')))
    patched, filled, total = {}, 0, 0
    for name, path in ps.items():
        if name not in pv:
            continue
        vals = {}
        for c in etree.fromstring(zv.read(pv[name])).iter(NS + 'c'):
            v = c.find(NS + 'v')
            if v is None:
                continue
            t = c.get('t')
            if t == 's' and (v.text or '').isdigit():
                vals[c.get('r')] = ('str', sst[int(v.text)])
            elif t in ('e', 'str'):
                vals[c.get('r')] = (t, v.text)
            else:
                vals[c.get('r')] = ('n', v.text)
        tree = etree.fromstring(zf.read(path))
        for c in tree.iter(NS + 'c'):
            if c.find(NS + 'f') is None:
                continue
            total += 1
            hit = vals.get(c.get('r'))
            if not hit or hit[1] is None:
                continue
            kind, val = hit
            old = c.find(NS + 'v')
            if old is not None:
                c.remove(old)
            etree.SubElement(c, NS + 'v').text = val
            if kind == 'str':
                c.set('t', 'str')
            elif kind == 'e':
                c.set('t', 'e')
            elif c.get('t') in ('s', 'str', 'e'):
                del c.attrib['t']
            filled += 1
        patched[path] = etree.tostring(tree, xml_declaration=True,
                                       encoding='UTF-8', standalone=True)
    if total and filled == 0:
        # A zero fill means the "recalculated" workbook carried no results, so the
        # file would open showing blanks everywhere. Fail loudly -- shipping it
        # quietly is far worse than stopping here.
        raise SystemExit(f'ERROR: no cached values found in {valued_path} '
                         f'({total} formula cells needed values). The recalculation '
                         f'step did not actually produce a calculated workbook.')
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item in zf.infolist():
            zo.writestr(item, patched.get(item.filename, zf.read(item.filename)))
    return filled, total
