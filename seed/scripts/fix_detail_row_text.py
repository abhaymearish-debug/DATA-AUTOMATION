"""
Re-paint cols A-F of every DETAIL-sheet data row so the row's font/fill match
its data tier — not its history.

Background (regression Abhay flagged 10 May 2026):

  Pack rows that started life as all-zero ("— No activity") had italic + grey
  font (FF6B7280) and grey fill (FFF3F4F6) baked in at sheet build time. When
  later daily updates poured real numbers into B-E, the value cells refreshed
  but the row text kept its no-activity styling — so PALAKKAD DETAIL row 109
  showed `5.00 / 0.00 / 3.40 / 1.60 / Balanced` but the row body was still
  italic-grey, looking misleadingly like a dead row.

  No script in the pipeline was repainting row text after a tier change.
  ksbc_detail_nested.py only updates B-E values; ksbc_sort_by_rating.py copies
  font+fill on sort and so faithfully preserves the staleness; fix_nested_cf.py
  and fix_rating_cell_format.py only touch col G (Rating).

This script closes that gap. For every DETAIL sheet in the workbook it:

  1. Walks every shop block.
  2. Skips the navy header (`<code> — <name>  |  Staff: <staff>`) and the
     column-header row that follows it, and skips the per-block TOTAL row
     (those are styled by other steps).
  3. For every remaining data row, classifies it as either a BRAND summary
     (col A is a brand name) or a PACK row (col A matches r"^\s*\d+\s*ML\s*$").
  4. Reads B/C/D/E. If ALL FOUR are zero → "no activity" styling
     (italic + FF6B7280 grey font + FFF3F4F6 grey fill). Otherwise →
     canonical clean styling for that row class.

Canonical styling spec (locked 10 May 2026, matches the rows that have always
rendered correctly):

  PACK with data       : 10pt, NOT bold, NOT italic, font FF000000, fill FFFFFFFF
  PACK no activity     : 10pt, NOT bold, italic,    font FF6B7280, fill FFF3F4F6
  BRAND with data      : 10pt, bold,    NOT italic, font FF1A237E, fill FFE5EAF0
  BRAND no activity    : 10pt, bold,    italic,     font FF6B7280, fill FFF3F4F6

Alignment:
  col A → horizontal=left, indent=3 (pack) or 1 (brand), vertical=center
  cols B-F → horizontal=right, indent=1, vertical=center

Col G (Rating) — color and fill are owned by `fix_rating_cell_format.py`
(tier-driven palette). This script forces col G to UNIFORM BOLD +
NOT-italic on every data row regardless of row class or tier (Abhay's
rule, 10 May 2026: "all ratings identical, all bold"). The tier color
is the only differentiator. TOTAL rows are skipped at the block level
and keep their own bold + bright-tier styling.

The script is idempotent: re-running on a clean workbook is a no-op.

Usage:
    python3 .claude/scripts/fix_detail_row_text.py <workbook.xlsx>
"""
from __future__ import annotations
import re
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
from openpyxl.formatting.formatting import ConditionalFormattingList


BONDS = ['KOLLAM', 'KOZHIKODE', 'ATTINGAL', 'PALAKKAD', 'KOTTARAKARA', 'KANNUR',
         'ALAPPUZHA', 'NEDUMANGAD', 'ALUVA', 'PERINTHALMANNA', 'THODUPUZHA',
         'PATHANAMTHITTA', 'KOTTAYAM', 'THRISSUR', 'TRIPUNITHURA']

PACK_RE = re.compile(r'^\s*\d+\s*ML\s*$', re.IGNORECASE)
HEADER_CODE_RE = re.compile(r'^(\d+)\s*—')

# MUKKAM alias — KSBC's portal exports the shop as 111014; master uses 11014.
ALIAS_CODE = {111014: 11014}

def _canon_code(c):
    return ALIAS_CODE.get(c, c)

FONT_NAME = 'Trebuchet MS'
FONT_SIZE = 10

# Uniform row heights across all DETAIL sheets (locked 10 May 2026 per Abhay):
# brand summary rows are slightly taller than pack rows so the brand band
# reads as the heavier visual anchor at a glance.
ROW_HEIGHT_PACK   = 18
ROW_HEIGHT_BRAND  = 22
ROW_HEIGHT_TOTAL  = 26  # slightly taller than brand rows so TOTAL anchors visually
ROW_HEIGHT_SPACER = 12  # blank breathing-room row between consecutive blocks
ROW_HEIGHT_HEADER = 32  # navy block-header banner — tall + airy
ROW_HEIGHT_COLHDR = 26  # column-header strip ("Brand / Pack | Opening | ...")

# Block-header / col-header palette + typography (locked 10 May 2026 for the
# "more elegant" pass).
HEADER_FILL_HEX     = 'FF1A237E'   # deep navy banner
HEADER_FONT_HEX     = 'FFFFFFFF'   # white text
HEADER_FONT_SIZE    = 13
HEADER_ACCENT_HEX   = 'FFFFD700'   # gold thin bottom-border accent

COLHDR_FILL_HEX     = 'FF263F80'   # softer navy (one step lighter than banner)
COLHDR_FONT_HEX     = 'FFFFFFFF'
COLHDR_FONT_SIZE    = 11
COLHDR_ACCENT_HEX   = 'FFFFD700'   # same gold thin bottom border

# (bold, italic, font_hex, fill_hex)
STYLE_PACK_DATA   = (False, False, 'FF000000', 'FFFFFFFF')
STYLE_PACK_NONE   = (False, True,  'FF6B7280', 'FFF3F4F6')
STYLE_BRAND_DATA  = (True,  False, 'FF1A237E', 'FFE5EAF0')
STYLE_BRAND_NONE  = (True,  True,  'FF6B7280', 'FFF3F4F6')

# TOTAL row styling — uniform dark grey across cols A-G. Cols A-F use white
# bold text. Col G uses a BRIGHT tier-color font on the same dark-grey fill
# (matches the canonical "TOTAL on dark grey + tier color font" look of
# every rating-bearing sheet — see fix_rating_cell_format.py palette).
TOTAL_FILL_HEX   = 'FF374151'   # dark grey
TOTAL_FONT_HEX   = 'FFFFFFFF'   # white (cols A-F)
TOTAL_BORDER_HEX = 'FFFFD700'   # gold accent on the top edge
TOTAL_FONT_SIZE  = 12           # bumped from 10 so TOTAL pops vs data rows

# Bright tier font colors used on the dark-grey TOTAL fill in col G.
# Mirrors fix_rating_cell_format.py's BRIGHT_FONT mapping.
TOTAL_TIER_FONT = {
    'high':     'FF4FC3F7',   # cyan
    'balanced': 'FF81C784',   # light green
    'inv':      'FFFFB74D',   # amber
    'crit':     'FFE57373',   # coral
    'none':     'FFB0BEC5',   # light grey
}


def _compute_tier(opening, receipts, sales, closing):
    """Mirrors fix_rating_cell_format._tier — five-tier classifier."""
    o = opening or 0
    r = receipts or 0
    s = sales or 0
    c = closing or 0
    if o == 0 and r == 0 and s == 0 and c == 0:
        return 'none'
    denom = o + r
    st = (s / denom) if denom else 0.0
    if st >= 0.8:  return 'high'
    if st >= 0.6:  return 'balanced'
    if st >= 0.4:  return 'inv'
    return 'crit'


def _is_pack_label(v) -> bool:
    return isinstance(v, str) and bool(PACK_RE.match(v))


def _is_header_label(v) -> bool:
    return isinstance(v, str) and ' — ' in v and 'Staff:' in v


def _num(v) -> float:
    if isinstance(v, (int, float)):
        return float(v)
    return 0.0


def _has_data(ws, r) -> bool:
    return any(_num(ws.cell(r, c).value) != 0 for c in (2, 3, 4, 5))


def _find_blocks(ws):
    """Return list of (header_row, total_row) for each shop block."""
    headers = [r for r in range(1, ws.max_row + 1)
               if _is_header_label(ws.cell(r, 1).value)]
    blocks = []
    for i, h in enumerate(headers):
        end_limit = headers[i + 1] - 1 if i + 1 < len(headers) else ws.max_row
        for r in range(h + 2, end_limit + 1):
            if ws.cell(r, 1).value == 'TOTAL':
                blocks.append((h, r))
                break
    return blocks


# ---- Block spacing ----------------------------------------------------------
#
# Insert one blank row before every shop-block header (after the first) so
# consecutive blocks have visible breathing room. Idempotent: skip if a
# blank row already sits above the header.
#
# Inserting a row in openpyxl shifts cells down but does NOT update formula
# row references. Every TOTAL/cluster row in the sheet uses formulas like
# =SUM(B{start}:B{end}) — those would silently point to the wrong rows after
# an insert. So after each insert we walk the entire sheet and increment
# every row reference >= the insert position by 1. Merged cell ranges are
# handled by openpyxl's insert_rows in 3.0+.

CELL_REF_RE = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def _shift_formula_refs(ws, from_row: int):
    """Increment every row reference >= from_row in every formula in the sheet."""
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not (isinstance(v, str) and v.startswith('=')):
                continue
            def _bump(m):
                col_dollar, col, row_dollar, n_str = m.group(1), m.group(2), m.group(3), m.group(4)
                n = int(n_str)
                if n >= from_row:
                    n += 1
                return f"{col_dollar}{col}{row_dollar}{n}"
            new_v = CELL_REF_RE.sub(_bump, v)
            if new_v != v:
                cell.value = new_v


def _detail_f_formula(r):
    """Canonical self-referencing Sell-Through formula for DETAIL col F at row r."""
    return (f'=IF(AND(B{r}=0,C{r}=0,D{r}=0,E{r}=0),"—",'
            f'IFERROR(D{r}/(B{r}+C{r}),0))')


def _detail_g_formula(r):
    """Canonical self-referencing Rating formula for DETAIL col G at row r."""
    return (f'=IF(AND(B{r}=0,C{r}=0,D{r}=0,E{r}=0),"— No activity",'
            f'IF(F{r}>=0.8,"🚀 High Performance",'
            f'IF(F{r}>=0.6,"✅ Balanced",'
            f'IF(F{r}>=0.4,"⚠️ Inventory Heavy","🚫 Critical Overstock"))))')


def _normalize_rating_formulas(ws):
    """Rewrite every DETAIL Sell-Through (col F) and Rating (col G) formula so
    it self-references its OWN current row.

    Every F/G formula in a DETAIL sheet is a pure self-reference, so the
    correct ref is ALWAYS the cell's own row — no matter how rows have been
    moved by spacer inserts, pack inserts, or block sorting. Running this as
    the final formula pass makes the sheet immune to the row-shift class of
    bug that left 148 cells (ATTINGAL DETAIL worst: +39 rows) pointing at
    another shop's row (16 Jun 2026 fix). A row is treated as a rating row iff
    its col-G cell currently holds an `=IF(AND(B...` formula, so headers /
    spacers / title rows are never touched. Returns the count of cells fixed.
    """
    fixed = 0
    for r in range(1, ws.max_row + 1):
        g = ws.cell(r, 7).value
        if not (isinstance(g, str) and g.startswith('=IF(AND(B')):
            continue
        want_f = _detail_f_formula(r)
        want_g = _detail_g_formula(r)
        if ws.cell(r, 6).value != want_f:
            ws.cell(r, 6).value = want_f
            fixed += 1
        if g != want_g:
            ws.cell(r, 7).value = want_g
            fixed += 1
    return fixed


def _ensure_block_spacing(ws):
    """Insert a blank row above every shop-block header (except the first).

    The structural-insert approach used on the first attempt corrupted 254
    TOTAL rows because openpyxl's insert_rows() leaves merged ranges
    behind, and a stale A:G merge sitting on a TOTAL row swallows its B-G
    values. The fix that works:

      1. UNMERGE every A:G block-header merge across the whole sheet first
         (preserve title r1 and subtitle r2). With no merges in play, no
         insert can trap anyone.
      2. Snapshot the set of header rows (by content match) so we know
         where to re-merge at the end.
      3. Process headers top-down. For each header (after the first),
         insert a blank row above it at its CURRENT (post-shift) position.
         After the insert, shift every formula's row references >= the
         insert position by +1.
      4. After all inserts, walk the sheet, find every header row, and
         re-merge A:G on it.

    Idempotent: if a blank row already sits above a header, skip the
    insert (still re-merge in case prior runs left a header unmerged).
    """
    # 1. Unmerge all single-row A:G block-header merges (skip title/subtitle).
    preserve = {1, 2}
    for mr in list(ws.merged_cells.ranges):
        s = str(mr)
        m = re.match(r"^A(\d+):G(\d+)$", s)
        if not m:
            continue
        r1, r2 = int(m.group(1)), int(m.group(2))
        if r1 != r2 or r1 in preserve:
            continue
        ws.unmerge_cells(s)

    # 2. Snapshot header positions (by content) — these are the rows that
    #    need re-merging at the end.
    header_rows = [r for r in range(1, ws.max_row + 1)
                   if _is_header_label(ws.cell(r, 1).value)]
    if len(header_rows) < 2:
        # Re-merge whatever headers exist and bail.
        for hr in header_rows:
            ws.merge_cells(f"A{hr}:G{hr}")
        return 0

    # 3. Insert spacers top-down, tracking cumulative offset.
    inserts = 0
    for original_h in header_rows[1:]:
        actual_h = original_h + inserts
        prev_row = actual_h - 1
        is_blank = all(
            ws.cell(prev_row, c).value in (None, '') for c in range(1, 8)
        )
        if is_blank:
            ws.row_dimensions[prev_row].height = ROW_HEIGHT_SPACER
            continue
        ws.insert_rows(actual_h)
        _shift_formula_refs(ws, actual_h)
        ws.row_dimensions[actual_h].height = ROW_HEIGHT_SPACER
        inserts += 1

    # 4. Re-merge every header row at its post-insert position. Heights +
    #    full styling are applied by _paint_block_headers in fix_detail_sheet.
    for r in range(1, ws.max_row + 1):
        if _is_header_label(ws.cell(r, 1).value):
            ws.merge_cells(f"A{r}:G{r}")

    return inserts


def _paint_row(ws, r, style, is_pack):
    bold, italic, font_color, fill_color = style
    fill = PatternFill('solid', fgColor=fill_color)
    indent_a = 3 if is_pack else 1
    # Uniform thin border on every cell (matches the canonical look of pack /
    # brand rows on existing DETAIL blocks). Newly-inserted rows from
    # `_insert_missing_pack_rows` have no borders by default, so we always
    # set them here.
    thin_grey = Side(style='thin', color='FFE5E7EB')
    border_all = Border(top=thin_grey, bottom=thin_grey, left=thin_grey, right=thin_grey)
    # Uniform row height — pack rows shorter, brand rows taller. Header /
    # col-header / TOTAL rows are skipped at the block level so their heights
    # (set by the build script) are preserved.
    ws.row_dimensions[r].height = ROW_HEIGHT_PACK if is_pack else ROW_HEIGHT_BRAND
    for c in range(1, 7):  # cols A..F  (G/Rating color+fill handled below)
        cell = ws.cell(r, c)
        cell.font = Font(
            name=FONT_NAME, size=FONT_SIZE,
            bold=bold, italic=italic, color=font_color,
        )
        cell.fill = fill
        cell.border = border_all
        if c == 1:
            cell.alignment = Alignment(horizontal='left', vertical='center', indent=indent_a)
        else:
            # Cols B-F centred (Abhay's preference 10 May 2026, "centre align B-G").
            cell.alignment = Alignment(horizontal='center', vertical='center')
        # Number formats: B-E whole numbers (with 2dp), F whole-number %.
        if c in (2, 3, 4, 5):
            cell.number_format = '0.00'
        elif c == 6:
            cell.number_format = '0%'   # whole-number % per Abhay's rule
    # Col G (Rating): force uniform BOLD + NOT italic + FONT_SIZE for every
    # data row (locked 10 May 2026 per Abhay — "all ratings identical, all
    # bold"). For inserted rows that have no prior cell formatting, we ALSO
    # apply a sensible direct fill + font color based on the row's data tier
    # so the rating cell looks identical to existing rows even before
    # fix_rating_cell_format.py runs.
    g = ws.cell(r, 7)
    # Compute tier from B-E to pick the right direct fill + font color.
    op = _num(ws.cell(r, 2).value)
    rc = _num(ws.cell(r, 3).value)
    sa = _num(ws.cell(r, 4).value)
    cl = _num(ws.cell(r, 5).value)
    tier_idx = 4  # default to 'none' (— No activity)
    if not (op == 0 and rc == 0 and sa == 0 and cl == 0):
        denom = op + rc
        st = (sa / denom) if denom else 0.0
        if   st >= 0.8: tier_idx = 0   # high
        elif st >= 0.6: tier_idx = 1   # balanced
        elif st >= 0.4: tier_idx = 2   # inv
        else:           tier_idx = 3   # crit
    g_fill_hex = DATA_RATING_STYLES[tier_idx][1]
    g_font_hex = DATA_RATING_STYLES[tier_idx][2]
    g.fill = PatternFill('solid', fgColor=g_fill_hex)
    g.font = Font(
        name=FONT_NAME, size=FONT_SIZE,
        bold=True, italic=False, color=g_font_hex,
    )
    g.alignment = Alignment(horizontal='center', vertical='center')
    g.border = border_all


def _norm_brand(v):
    """Normalise a brand string for matching: upper-case, drop apostrophes
    (curly/straight/back-tick) and collapse internal whitespace. Mirrors
    ksbc_detail_nested._norm_brand. (H4 fix 1 Jun 2026.)

    The raw KSBC export and the DETAIL row templates disagree on punctuation
    for some brands (raw 'MORNING WALKERS XO BRANDY' vs template
    "MORNING WALKER'S XO BRANDY"). Exact-string matching silently dropped
    those rows to 0 here too — the self-heal wrote {O:0,R:0,S:0,C:0} into a
    blanked row, re-opening the MAY 2026 Morning Walker's regression. Used as
    the brand component of every COMBINED-aggregate key AND every lookup, so
    both sides are normalised consistently. NEVER written to a cell — only a
    dict key — so display spelling is untouched.
    """
    if v is None:
        return ''
    return ' '.join(str(v).strip().upper()
                    .replace("’", "").replace("'", "").replace("`", "").split())


def _norm_pack(v):
    """Normalise a pack label for matching: strip + upper-case. Kills the
    `500 ml` vs `500 ML` mismatch that broke self-heal lookups and missing-pack
    idempotency (audit Suspicion C). (H4 fix 1 Jun 2026.)"""
    return str(v).strip().upper() if v is not None else ''


def _build_combined_aggregates(wb):
    """Walk the MAY-N COMBINED sheet (or any *COMBINED* sheet) and build:
       shop_brand_pack[(code, norm_brand, norm_pack)] = {O,R,S,C}
       shop_brand[(code, norm_brand)] = {O,R,S,C}
    Brand and pack components are NORMALISED (see _norm_brand / _norm_pack) so
    apostrophe/case drift between the raw export and the DETAIL templates can't
    cause a lookup miss. Returns ({}, {}) if no COMBINED sheet exists."""
    from collections import defaultdict
    combined_name = next((n for n in wb.sheetnames if 'COMBINED' in n.upper()), None)
    if combined_name is None:
        return {}, {}
    cs = wb[combined_name]
    sbp = defaultdict(lambda: {"O": 0.0, "R": 0.0, "S": 0.0, "C": 0.0})
    sb  = defaultdict(lambda: {"O": 0.0, "R": 0.0, "S": 0.0, "C": 0.0})
    for r in range(2, cs.max_row + 1):
        code = cs.cell(r, 2).value
        brand = cs.cell(r, 5).value
        pack = cs.cell(r, 6).value
        if code is None or brand is None or pack is None:
            continue
        try:
            code_i = _canon_code(int(code))
        except (TypeError, ValueError):
            continue
        o = cs.cell(r, 8).value or 0
        rc = cs.cell(r, 9).value or 0
        s = cs.cell(r, 10).value or 0
        cl = cs.cell(r, 11).value or 0
        nb = _norm_brand(brand)
        np = _norm_pack(pack)
        sbp[(code_i, nb, np)]["O"] += o
        sbp[(code_i, nb, np)]["R"] += rc
        sbp[(code_i, nb, np)]["S"] += s
        sbp[(code_i, nb, np)]["C"] += cl
        sb[(code_i, nb)]["O"] += o
        sb[(code_i, nb)]["R"] += rc
        sb[(code_i, nb)]["S"] += s
        sb[(code_i, nb)]["C"] += cl
    return dict(sbp), dict(sb)


def _insert_missing_pack_rows(ws, shop_brand_pack):
    """Auto-insert any (brand, pack) combo present in shop_brand_pack but
    missing from the DETAIL sheet block for that shop.

    Background — regression Abhay flagged 11 May 2026: when KSBC starts
    dispatching a NEW pack size to a shop mid-month (e.g. shop 105015 began
    receiving K.S 99 180 ML in May for the first time), there's no row in
    the DETAIL sheet to receive the data. The brand total picks it up but
    the pack-level breakdown shows 0, masking the inconsistency.

    For each shop block:
      1. Identify all existing (brand → [packs]) pairs in the block.
      2. Compare against shop_brand_pack entries for this shop's code.
      3. For each (brand, pack) in data but not in the block, insert a new
         pack row directly after the LAST existing pack of that brand
         (or after the brand row itself if no packs exist).

    Each insert uses the same safe-insert pattern as _ensure_block_spacing:
      - All block-header A:G merges are unmerged FIRST (sheet-wide) so
        trapped values can't get blanked.
      - Insert at the target row.
      - Shift every formula's row refs >= insert_row by +1.
      - At the end, re-merge all header rows at their post-insert positions.

    Processes inserts BOTTOM-UP within and across blocks so earlier inserts
    don't shift positions we haven't processed yet.

    Returns the count of pack rows inserted.
    """
    if not shop_brand_pack:
        return 0

    # Snapshot block boundaries before any inserts.
    blocks = []
    headers = [r for r in range(1, ws.max_row + 1)
               if _is_header_label(ws.cell(r, 1).value)]
    for i, h in enumerate(headers):
        end = headers[i + 1] - 1 if i + 1 < len(headers) else ws.max_row
        for r in range(h + 2, end + 1):
            if ws.cell(r, 1).value == 'TOTAL':
                blocks.append((h, r))
                break

    # For each block, plan the inserts: (block_index, brand_row_pos,
    # last_pack_or_brand_row_pos, brand_name, missing_packs_to_insert).
    # Then we'll process them in row-descending order.
    planned_inserts = []  # list of (insert_at_row, brand, pack_label, values_dict)

    for block_idx, (header_r, total_r) in enumerate(blocks):
        code = None
        a = ws.cell(header_r, 1).value
        m = HEADER_CODE_RE.match(a) if isinstance(a, str) else None
        if not m:
            continue
        code = _canon_code(int(m.group(1)))

        # Walk the block, building brand → {existing packs, last position}.
        current_brand = None
        existing_packs_by_brand = {}     # brand → set of pack labels
        last_row_for_brand = {}          # brand → last row of that brand's section
        brand_row_for_brand = {}         # brand → row of the brand summary row
        for r in range(header_r + 2, total_r):
            label = ws.cell(r, 1).value
            if label is None or (isinstance(label, str) and label.strip() == ''):
                continue
            # Key brand/pack by NORMALISED form so the block side matches the
            # normalised keys in shop_brand_pack (apostrophe/case drift). The
            # raw COMBINED pack `p` is still what gets WRITTEN for display.
            if _is_pack_label(label):
                if current_brand is not None:
                    existing_packs_by_brand.setdefault(current_brand, set()).add(_norm_pack(label))
                    last_row_for_brand[current_brand] = r
            else:
                current_brand = _norm_brand(label)
                existing_packs_by_brand.setdefault(current_brand, set())
                brand_row_for_brand[current_brand] = r
                last_row_for_brand[current_brand] = r  # initially the brand row itself

        # Build expected packs for this shop from shop_brand_pack (keys already
        # normalised). Compare on normalised pack; write the normalised pack
        # label (uppercase ML form — display-identical to the export).
        expected_packs_by_brand = {}
        for (c, b, p), vals in shop_brand_pack.items():
            if c != code:
                continue
            expected_packs_by_brand.setdefault(b, {})[_norm_pack(p)] = vals

        # For each brand, find missing packs and plan inserts.
        for brand, expected in expected_packs_by_brand.items():
            if brand not in brand_row_for_brand:
                # The brand row itself is missing from the block — different
                # problem; brand-row insertion isn't supported here.
                continue
            existing = existing_packs_by_brand.get(brand, set())
            for pack, vals in expected.items():
                if pack in existing:
                    continue
                insert_row = last_row_for_brand[brand] + 1
                planned_inserts.append((insert_row, brand, pack, vals))

    if not planned_inserts:
        return 0

    # Sort inserts by row DESCENDING so each ws.insert_rows() call doesn't
    # shift positions for inserts we haven't done yet. Stable within same row.
    planned_inserts.sort(key=lambda x: -x[0])
    original_insert_rows = sorted([p[0] for p in planned_inserts])  # ascending

    # Unmerge all block-header A:G merges sheet-wide first (preserve r1/r2).
    preserve = {1, 2}
    for mr in list(ws.merged_cells.ranges):
        s = str(mr)
        m = re.match(r"^A(\d+):G(\d+)$", s)
        if m and int(m.group(1)) == int(m.group(2)):
            r1 = int(m.group(1))
            if r1 not in preserve:
                ws.unmerge_cells(s)

    # Step A: do all the inserts bottom-up. Each insert_rows(P) places a blank
    # at the ORIGINAL P (no prior insert affected it because we're DESC).
    #
    # 16 Jun 2026 fix: every DETAIL F/G formula is a PURE self-reference to its
    # own row, so they never need a cross-row "shift". The previous Step-B
    # batched bump double-shifted formulas written by earlier passes and left
    # 148 cells (ATTINGAL DETAIL worst: +39) pointing at another shop's row.
    # We now write each inserted row's formula as a plain self-ref; the buggy
    # Step-B pass is removed, and _normalize_rating_formulas() (called at the
    # end of fix_detail_sheet) rewrites EVERY F/G cell to reference its own
    # current row, making the sheet immune to any insert/spacer/sort shift.
    inserted = 0
    for (insert_row, brand, pack, vals) in planned_inserts:
        ws.insert_rows(insert_row)
        ws.cell(insert_row, 1).value = pack
        ws.cell(insert_row, 2).value = vals.get('O', vals.get('Opening', 0))
        ws.cell(insert_row, 3).value = vals.get('R', vals.get('Receipts', 0))
        ws.cell(insert_row, 4).value = vals.get('S', vals.get('Sales', 0))
        ws.cell(insert_row, 5).value = vals.get('C', vals.get('Closing', 0))
        for c in (2, 3, 4, 5):
            ws.cell(insert_row, c).number_format = '0.00'
        ws.cell(insert_row, 6).number_format = '0%'
        ws.cell(insert_row, 6).value = _detail_f_formula(insert_row)
        ws.cell(insert_row, 7).value = _detail_g_formula(insert_row)
        inserted += 1

    # Re-merge all header rows at their post-insert positions.
    for r in range(1, ws.max_row + 1):
        if _is_header_label(ws.cell(r, 1).value):
            ws.merge_cells(f"A{r}:G{r}")

    return inserted


def _recover_broken_rows(ws, shop_brand_pack, shop_brand):
    """Self-heal: any data row whose B-E are None (a casualty of the spacer
    insert) gets its values re-derived from the COMBINED aggregates and its
    F + G formulas restored. Idempotent (skips rows that already have data).
    Without this step, the very first run on a clean workbook would leave
    27 of ~7700 data rows blank — the corner-case where openpyxl's
    insert_rows leaks B-G of the row immediately preceding a TOTAL.
    """
    if not shop_brand_pack and not shop_brand:
        return 0  # nothing to recover from
    recovered = 0
    current_code = None
    current_brand = None
    in_block = False
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is None or (isinstance(a, str) and a.strip() == ''):
            continue
        if _is_header_label(a):
            in_block = True
            m = HEADER_CODE_RE.match(a)
            current_code = _canon_code(int(m.group(1))) if m else None
            current_brand = None
            continue
        if not in_block or a in ('Brand / Pack', 'TOTAL'):
            if a == 'TOTAL':
                # End of a shop block: drop brand context so a stray blank row
                # before the next header can't be healed with the prior shop's
                # brand (16 Jun 2026 fix).
                current_brand = None
                in_block = False
            continue
        if _is_pack_label(a):
            pack_label = _norm_pack(a)
        else:
            current_brand = a
            pack_label = None
        # Only recover if B is None (the canary)
        if ws.cell(r, 2).value is not None:
            continue
        if current_code is None or current_brand is None:
            continue
        # Look up by NORMALISED brand/pack — keys in shop_brand_pack/shop_brand
        # are normalised, so a raw display label (apostrophes/case) still hits.
        if pack_label is not None:
            t = shop_brand_pack.get((current_code, _norm_brand(current_brand), pack_label),
                                    {"O": 0, "R": 0, "S": 0, "C": 0})
        else:
            t = shop_brand.get((current_code, _norm_brand(current_brand)),
                               {"O": 0, "R": 0, "S": 0, "C": 0})
        ws.cell(r, 2).value = t["O"]
        ws.cell(r, 3).value = t["R"]
        ws.cell(r, 4).value = t["S"]
        ws.cell(r, 5).value = t["C"]
        for c in (2, 3, 4, 5):
            ws.cell(r, c).number_format = '0.00'
        ws.cell(r, 6).number_format = '0%'   # whole-number % per Abhay's rule (10 May 2026)
        ws.cell(r, 6).value = (
            f'=IF(AND(B{r}=0,C{r}=0,D{r}=0,E{r}=0),"—",'
            f'IFERROR(D{r}/(B{r}+C{r}),0))'
        )
        ws.cell(r, 7).value = (
            f'=IF(AND(B{r}=0,C{r}=0,D{r}=0,E{r}=0),"— No activity",'
            f'IF(F{r}>=0.8,"🚀 High Performance",'
            f'IF(F{r}>=0.6,"✅ Balanced",'
            f'IF(F{r}>=0.4,"⚠️ Inventory Heavy","🚫 Critical Overstock"))))'
        )
        recovered += 1
    return recovered


def _paint_block_headers(ws, header_r):
    """Style the navy banner row + the column-header row right beneath it.

    Locked 10 May 2026 (Abhay's "more elegant" pass):
      - Banner: 32pt, deep navy fill, 13pt bold white, vertically centred,
        thin GOLD bottom border to separate it from the col-header row.
      - Col-header: 26pt, softer navy fill, 11pt bold white, centred, thin
        gold bottom border to separate from the first data row.

    The two-tone navy + twin gold accents read as a single titled section
    that clearly anchors each shop block. Idempotent.
    """
    fill_h = PatternFill('solid', fgColor=HEADER_FILL_HEX)
    fill_c = PatternFill('solid', fgColor=COLHDR_FILL_HEX)
    gold_thin = Side(style='thin', color=HEADER_ACCENT_HEX)

    # Banner row (header_r) — A:G is merged, so col 1 is the master cell.
    ws.row_dimensions[header_r].height = ROW_HEIGHT_HEADER
    for c in range(1, 8):
        cell = ws.cell(header_r, c)
        cell.fill = fill_h
        cell.font = Font(name=FONT_NAME, size=HEADER_FONT_SIZE,
                         bold=True, italic=False, color=HEADER_FONT_HEX)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        existing = cell.border
        cell.border = Border(
            top=existing.top if existing and existing.top else None,
            bottom=gold_thin,
            left=existing.left if existing and existing.left else None,
            right=existing.right if existing and existing.right else None,
        )

    # Column-header row (header_r + 1)
    colhdr_r = header_r + 1
    if colhdr_r > ws.max_row:
        return
    ws.row_dimensions[colhdr_r].height = ROW_HEIGHT_COLHDR
    for c in range(1, 8):
        cell = ws.cell(colhdr_r, c)
        cell.fill = fill_c
        cell.font = Font(name=FONT_NAME, size=COLHDR_FONT_SIZE,
                         bold=True, italic=False, color=COLHDR_FONT_HEX)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        existing = cell.border
        cell.border = Border(
            top=existing.top if existing and existing.top else None,
            bottom=gold_thin,
            left=existing.left if existing and existing.left else None,
            right=existing.right if existing and existing.right else None,
        )


def _paint_total_row(ws, total_r):
    """Apply canonical TOTAL row styling — uniform dark-grey fill across
    cols A-G with gold medium top border. Cols A-F use white bold text;
    col G uses a BRIGHT tier-color font (cyan/green/amber/coral/grey)
    on the same dark-grey fill so it pops as the rating callout while
    the row itself reads as one solid TOTAL band.

    Tier on col G is computed from this row's own B-E, mirroring the
    rating formula. Re-running on a clean row is a no-op.
    """
    fill = PatternFill('solid', fgColor=TOTAL_FILL_HEX)
    top_border = Side(style='medium', color=TOTAL_BORDER_HEX)

    # Compute tier from this TOTAL row's B-E for col G's bright font color.
    op = ws.cell(total_r, 2).value
    rc = ws.cell(total_r, 3).value
    sa = ws.cell(total_r, 4).value
    cl = ws.cell(total_r, 5).value
    op = op if isinstance(op, (int, float)) else 0
    rc = rc if isinstance(rc, (int, float)) else 0
    sa = sa if isinstance(sa, (int, float)) else 0
    cl = cl if isinstance(cl, (int, float)) else 0
    tier = _compute_tier(op, rc, sa, cl)
    g_font_color = TOTAL_TIER_FONT[tier]

    for c in range(1, 8):  # cols A..G — uniform dark-grey fill across the band
        cell = ws.cell(total_r, c)
        cell.fill = fill
        existing = cell.font
        font_color = g_font_color if c == 7 else TOTAL_FONT_HEX
        cell.font = Font(
            name=existing.name or FONT_NAME,
            size=TOTAL_FONT_SIZE,
            bold=True, italic=False,
            color=font_color,
        )
        # All cols A-G centre-aligned on the TOTAL row (locked 10 May 2026,
        # Abhay's "centre align B-G" + symmetry with col A's centre).
        cell.alignment = Alignment(horizontal='center', vertical='center')
        # Number formats: B-E whole numbers (with 2dp), F whole-number %.
        if c in (2, 3, 4, 5):
            cell.number_format = '0.00'
        elif c == 6:
            cell.number_format = '0%'   # whole-number %
        existing_border = cell.border
        cell.border = Border(
            top=top_border,
            bottom=existing_border.bottom if existing_border and existing_border.bottom else None,
            left=existing_border.left if existing_border and existing_border.left else None,
            right=existing_border.right if existing_border and existing_border.right else None,
        )


# Per-tier rating styling on col G of data rows. Locked 10 May 2026 per
# Abhay's "uniform across all DETAIL sheets" rule. CF rules set BOTH fill
# AND font (bold + color) so the styling can't drift via direct cell edits
# and looks identical on every DETAIL sheet.
#
#   High Performance   →  blue text   on light-blue   fill
#   Balanced           →  green text  on light-green  fill
#   Inventory Heavy    →  amber text  on light-amber  fill
#   Critical Overstock →  red text    on light-red    fill
#   No activity        →  grey text   on light-grey   fill
#
# (formula_text, fill_hex, font_hex)
DATA_RATING_STYLES = [
    ('"🚀 High Performance"',  'FFBBDEFB', 'FF1565C0'),  # blue
    ('"✅ Balanced"',           'FFDCEDC8', 'FF2E7D32'),  # green
    ('"⚠️ Inventory Heavy"',   'FFFFE0B2', 'FFE65100'),  # amber-orange
    ('"🚫 Critical Overstock"', 'FFFFCDD2', 'FFC62828'),  # red
    ('"— No activity"',         'FFECEFF1', 'FF6B7280'),  # grey
]


def _rebuild_g_cf(ws):
    """Rebuild col G conditional formatting using current block positions.

    The previous run's CF ranges are stale (don't follow row inserts), which
    leaks the data-row tier-fill onto TOTAL rows whose positions shifted past
    the original ranges. We wipe ALL CF on this sheet and re-apply only the
    data-row fill rules per current block. TOTAL row col G is styled directly
    by _paint_total_row (uniform dark grey + bright tier font), so it gets
    NO CF — guaranteeing nothing can override.

    Excel quirk: CF solid fills must be set via bgColor inside the dxf, not
    fgColor — see .claude/memory/ksbc-detail-cf.md.
    """
    # Find every shop block on this sheet (post-insert positions).
    blocks = []
    headers = [r for r in range(1, ws.max_row + 1)
               if _is_header_label(ws.cell(r, 1).value)]
    for i, h in enumerate(headers):
        end_limit = headers[i + 1] - 1 if i + 1 < len(headers) else ws.max_row
        for r in range(h + 2, end_limit + 1):
            if ws.cell(r, 1).value == 'TOTAL':
                blocks.append((h, r))
                break
    if not blocks:
        return 0
    data_ranges = []
    for header_r, total_r in blocks:
        data_start = header_r + 2  # row after col header
        data_end = total_r - 1     # row before TOTAL
        if data_end >= data_start:
            data_ranges.append(f'G{data_start}:G{data_end}')
    # Wipe ALL existing CF on the sheet — ensures no stale rules linger.
    ws.conditional_formatting = ConditionalFormattingList()
    if not data_ranges:
        return 0
    range_str = ' '.join(data_ranges)
    for priority, (txt, fill_color, font_color) in enumerate(DATA_RATING_STYLES, start=1):
        # Set BOTH fill (via bgColor — Excel CF quirk) AND font (bold + tier
        # color) in one dxf, so each tier renders identically on every sheet.
        dxf = DifferentialStyle(
            fill=PatternFill(bgColor=fill_color),
            font=Font(name=FONT_NAME, size=FONT_SIZE, bold=True, color=font_color),
        )
        rule = Rule(type='cellIs', operator='equal', formula=[txt],
                    stopIfTrue=False, dxf=dxf)
        rule.priority = priority
        ws.conditional_formatting.add(range_str, rule)
    return len(blocks)


def fix_detail_sheet(ws, shop_brand_pack=None, shop_brand=None):
    # Step 1 — ensure blank spacer rows between consecutive shop blocks. Done
    # FIRST so the row indices we walk in step 2 are the post-spacing ones.
    _ensure_block_spacing(ws)
    # Step 1b — self-heal any data rows the insert may have left blank.
    if shop_brand_pack is not None and shop_brand is not None:
        _recover_broken_rows(ws, shop_brand_pack, shop_brand)
    # Step 1c — auto-insert any (brand, pack) combos present in COMBINED data
    # but missing from the DETAIL sheet block. Without this, new pack sizes
    # dispatched mid-month get summed into the brand total but have nowhere
    # to land at the pack level (regression on shop 105015 K.S 99 180 ML).
    if shop_brand_pack is not None:
        _insert_missing_pack_rows(ws, shop_brand_pack)
    # Step 1d — normalize every Sell-Through / Rating formula to self-reference
    # its own current row. Immune to insert/spacer/sort shifts; repairs and
    # prevents the 148-cell row-shift bug (16 Jun 2026).
    _normalize_rating_formulas(ws)
    # Step 2 — repaint and resize.
    blocks = _find_blocks(ws)
    if not blocks:
        return 0
    repainted = 0
    for header_r, total_r in blocks:
        # Navy banner + col-header strip — styled tall + elegant.
        _paint_block_headers(ws, header_r)
        # per-block TOTAL row — set canonical dark-grey fill + white bold
        # text on cols A-G, plus uniform 26pt height.
        ws.row_dimensions[total_r].height = ROW_HEIGHT_TOTAL
        _paint_total_row(ws, total_r)
        for r in range(header_r + 2, total_r):
            label = ws.cell(r, 1).value
            if label is None or (isinstance(label, str) and label.strip() == ''):
                continue
            is_pack = _is_pack_label(label)
            has_data = _has_data(ws, r)
            if is_pack:
                style = STYLE_PACK_DATA if has_data else STYLE_PACK_NONE
            else:
                style = STYLE_BRAND_DATA if has_data else STYLE_BRAND_NONE
            _paint_row(ws, r, style, is_pack)
            repainted += 1
    # Step 3 — rebuild col G CF using the post-insert block positions so
    # stale data-fill rules can't leak onto the TOTAL row.
    _rebuild_g_cf(ws)
    return repainted


def main():
    if len(sys.argv) < 2:
        print('Usage: python3 fix_detail_row_text.py <workbook.xlsx>')
        sys.exit(1)
    target = sys.argv[1]
    wb = load_workbook(target, data_only=False)

    # Build COMBINED-aggregates once so every DETAIL sheet can self-heal any
    # rows that the spacer-insert leaves blank.
    sbp, sb = _build_combined_aggregates(wb)

    total = 0
    for bond in BONDS:
        sn = f'{bond} DETAIL'
        if sn not in wb.sheetnames:
            print(f'  · {sn:22s}  (missing)')
            continue
        n = fix_detail_sheet(wb[sn], sbp, sb)
        total += n
        print(f'  ✓ {sn:22s}  {n} rows repainted')

    wb.save(target)
    print(f'\nRepainted {total} DETAIL rows. Saved {target}.')


if __name__ == '__main__':
    main()
