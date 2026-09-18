"""K.S. Distillery -- Warehouse Requirement Breakup builder (LOCKED format).

Produces a two-sheet workbook from a parsed-screenshot DATA dict:

  Sheet 1: WAREHOUSE REQUIREMENT
    - Title + subtitle bands
    - Brand-grouped matrix (warehouses × brand-pack cases)
    - TOTAL row
    - BLEND & ENA REQUIREMENT section
        BLEND (L)         = sum(cases × pack_factor)
        BLEND ROUNDED (L) = CEILING(blend, 1000)  [overridable per brand]
        ENA  (L)          = Blend Rounded × (0.75 / 1.68)
                            = strength balance: 75°LP / 168°LP = 44.64%

  Sheet 2: BLEND BREAKUP
    - K.S. DISTILLERY wordmark + gold rule
    - Simple table: BRAND | CASES | SHARE  (sorted by cases desc)
    - Navy TOTAL row

Usage:
    from build_warehouse_requirement import build_workbook

    data = {
        "Kollam": {("BCB", "L"): 5, ("OPR", "HL"): 250, ...},
        "Alappuzha": {...},
        ...
    }
    build_workbook(
        data,
        output_path="/path/to/output.xlsx",
        compile_date="11 May 2026",
        # round_overrides normally OMITTED (locked policy: CEILING to 1000).
    )
"""
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Canonical brand & pack metadata (LOCKED -- do not change order)
# ---------------------------------------------------------------------------
BRAND_ORDER = ["BCB", "Blenders", "MWB", "KS", "OPR", "ROFR", "MBR"]

# Uniform 4-pack layout per brand (locked 13 May 2026 -- every brand has
# all four pack sizes so factory stock can be entered for every cell).
_ALL_PACKS = ["NIP", "PINT", "HL", "L"]
PACKS_BY_BRAND = {brand: list(_ALL_PACKS) for brand in
                  ["BCB", "Blenders", "MWB", "KS", "OPR", "ROFR", "MBR"]}
PACK_LABEL = {"NIP": "180ml", "PINT": "375ml", "HL": "500ml", "L": "1000ml"}

# Brand & pack aliases -- CANONICAL REFERENCE for the WhatsApp-screenshot
# parsing step (warehouse-requirement SKILL.md steps 2-3). build_workbook()
# itself expects keys that are ALREADY canonical and does NOT consult these
# tables, but keep them in sync with warehouse-requirement-spec.md so the
# parsing step has one correct source of truth (and so a future code-path
# parser can import them).
BRAND_ALIASES = {
    "bcb": "BCB", "bcbc": "BCB",
    "blenders": "Blenders", "blend": "Blenders",
    "mwb": "MWB",
    "ks": "KS", "ks99": "KS", "k.s.": "KS", "k.s": "KS",
    "opr": "OPR", "op": "OPR",
    "rofr": "ROFR",
    "mbr": "MBR",
}
PACK_ALIASES = {
    "n": "NIP", "nip": "NIP", "180": "NIP", "180ml": "NIP",
    "p": "PINT", "pint": "PINT", "375": "PINT", "375ml": "PINT",
    "360": "PINT", "360ml": "PINT",          # KS99 pint bottle is 360ml -> PINT slot
    "h": "HL", "hl": "HL", "hltr": "HL", "hlf": "HL", "500": "HL", "500ml": "HL",
    "l": "L", "ltr": "L", "1000": "L", "1000ml": "L",
}

# Conversion factors (LOCKED)
BLEND_FACTOR = {"NIP": 8.64, "PINT": 9.0, "HL": 9.0, "L": 9.0}
ROUND_TO = 1000
ENA_NUM = 0.75   # finished spirit  75 LP = 42.8 % v/v
ENA_DEN = 1.68   # ENA              168 LP = 96.0 % v/v
# -> Blend Rounded x (0.75 / 1.68) = ENA litres per brand  (~44.64%)

# Palette (LOCKED)
NAVY        = "FF1A237E"
NAVY_SOFT   = "FF263F80"
GOLD        = "FFFFD700"
GOLD_DIV    = "FFD4A106"
GOLD_RULE   = "FFB8860B"
WHITE       = "FFFFFFFF"
DARK_GREY   = "FF374151"
INK         = "FF1A237E"
SOFT_INK    = "FF6B7280"
WH_BAND     = "FFF8F4E8"
WH_BAND_ALT = "FFFAF1DC"
ZEBRA       = "FFF6F8FB"
HAIRLINE    = "FFE5E7EB"
ROW_TINT    = "FFF6F8FB"
CLUSTER_FILL = "FF263F80"   # navy banner fill (matches BOND PERFORMANCE)


# ---------------------------------------------------------------------------
# Cluster mapping (depot -> cluster) -- canonical, extensible.
# Cluster identity follows Abhay's locked KSBC cluster list (by bond);
# depot/town names roll up to their bond's cluster. Banners are labelled
# simply "CLUSTER 1 / 2 / 3" -- NO South/Central/North suffix, no ASM names.
# clusters_from_depots() raises on any unmapped depot so a new town surfaces
# instead of being silently misplaced.
# ---------------------------------------------------------------------------
CLUSTER_ORDER = ["CLUSTER 1", "CLUSTER 2", "CLUSTER 3"]
DEPOT_CLUSTER = {
    # Cluster 1 -- South (ALAPPUZHA, ATTINGAL, NEDUMANGAD, KOLLAM,
    #                     KOTTARAKARA, PATHANAMTHITTA + their depots)
    "BALARAMAPURAM": "CLUSTER 1", "NEDUMANGAD": "CLUSTER 1",
    "ATTINGAL": "CLUSTER 1", "MENAMKULAM": "CLUSTER 1",
    "KOLLAM": "CLUSTER 1", "KARUNAGAPPALLY": "CLUSTER 1",
    "THIRUVALLA": "CLUSTER 1", "KOTTARAKKARA": "CLUSTER 1",
    "KOTTARAKARA": "CLUSTER 1", "PATHANAMTHITTA": "CLUSTER 1",
    "ALAPPUZHA": "CLUSTER 1", "ADOOR": "CLUSTER 1",
    "KAYAMKULAM": "CLUSTER 1", "PARASSALA": "CLUSTER 1",
    # Cluster 2 -- Central (THRISSUR, TRIPUNITHURA, THODUPUZHA,
    #                       KOTTAYAM, ALUVA + their depots)
    "THRISSUR": "CLUSTER 2", "TRIPUNITHURA": "CLUSTER 2",
    "THODUPUZHA": "CLUSTER 2", "KOTTAYAM": "CLUSTER 2",
    "ALUVA": "CLUSTER 2", "ERNAKULAM": "CLUSTER 2",
    "MUVATTUPUZHA": "CLUSTER 2", "PERUMBAVOOR": "CLUSTER 2",
    "CHALAKUDY": "CLUSTER 2", "KOTHAMANGALAM": "CLUSTER 2",
    "ANGAMALY": "CLUSTER 2", "PALA": "CLUSTER 2",
    "AYARKUNNAM": "CLUSTER 2", "THRIPPUNNITHURA": "CLUSTER 2",
    "KADAVANTHRA": "CLUSTER 2",
    # Cluster 3 -- North (KANNUR, KOZHIKODE, PALAKKAD, PERINTHALMANNA
    #                     + their depots)
    "MENONPARA": "CLUSTER 3", "PALAKKAD": "CLUSTER 3",
    "PERINTHALMANNA": "CLUSTER 3", "NADUVANNUR": "CLUSTER 3",
    "CALICUT": "CLUSTER 3", "KOZHIKODE": "CLUSTER 3",
    "KALPATTA": "CLUSTER 3", "KALPETTA": "CLUSTER 3",
    "BETTATHUR": "CLUSTER 3", "KANNUR": "CLUSTER 3",
    "MANANTHAVADY": "CLUSTER 3", "VADAKARA": "CLUSTER 3",
    "KOYILANDY": "CLUSTER 3", "MUKKAM": "CLUSTER 3",
    "KASARAGOD": "CLUSTER 3", "OTTAPALAM": "CLUSTER 3",
    "SHORANUR": "CLUSTER 3", "MANNARKKAD": "CLUSTER 3",
    "TIRUR": "CLUSTER 3", "PAYYANNUR": "CLUSTER 3",
    "THALASSERY": "CLUSTER 3", "BATHERY": "CLUSTER 3",
    "SULTHAN BATHERY": "CLUSTER 3",
}


def clusters_from_depots(data):
    """Group data's warehouses into ordered clusters via DEPOT_CLUSTER.

    Returns [(cluster_label, [warehouse, ...]), ...] in CLUSTER_ORDER,
    skipping empty clusters and preserving each warehouse's insertion order
    within its cluster. Raises ValueError listing any unmapped depot.
    """
    unmapped = sorted({wh for wh in data
                       if str(wh).strip().upper() not in DEPOT_CLUSTER})
    if unmapped:
        raise ValueError(
            "Unmapped depot(s) -- add to DEPOT_CLUSTER before clustering: "
            + ", ".join(unmapped))
    groups = {c: [] for c in CLUSTER_ORDER}
    for wh in data:
        groups[DEPOT_CLUSTER[str(wh).strip().upper()]].append(wh)
    return [(c, groups[c]) for c in CLUSTER_ORDER if groups[c]]


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def build_workbook(data, output_path, compile_date=None, round_overrides=None,
                   physical_stock=None, clusters=None):
    """Build the locked K.S. Distillery warehouse requirement workbook.

    Args:
        data: {warehouse_name: {(brand, pack): cases}}
              Brand must be canonical (BRAND_ORDER); pack must be one of
              NIP / PINT / HL / L. Empty pack quantities can be omitted.
        output_path: absolute path to the .xlsx to write.
        compile_date: "DD Month YYYY" string. Defaults to today.
        clusters: optional [(cluster_label, [warehouse_name, ...]), ...].
                  When supplied, the matrix groups depot rows under a navy
                  cluster banner that carries live SUM subtotals across every
                  pack column + a cluster total; TOTAL REQUIREMENT then sums
                  the banners so nothing double-counts. Must be an exact
                  partition of data's warehouses. Use clusters_from_depots(data)
                  to derive it from DEPOT_CLUSTER. Omit for the flat layout.
        round_overrides: optional {brand: rounded_blend_litres} dict.
                         NORMALLY OMITTED -- locked 13 May 2026 policy is
                         BLEND ROUNDED = CEILING(blend, 1000) for every brand.
                         If a brand has an override, BLEND ROUNDED becomes a
                         hard value instead of CEILING(blend, 1000), which also
                         FREEZES that brand's ENA (ENA is computed off the
                         rounded value) so it will NOT follow a LESS: PHYSICAL
                         STOCK deduction. Only pass an override when you
                         explicitly intend that.
        physical_stock: optional {(brand, pack): cases} dict of cases held
                        at the factory as finished goods. The Production
                        row (= Requirement - Stock, floor 0) drives the
                        BLEND / ENA section. Missing entries default to 0.
    """
    if compile_date is None:
        compile_date = datetime.now().strftime("%d %B %Y")
    round_overrides = round_overrides or {}
    physical_stock = physical_stock or {}

    # Fail loud on non-canonical (brand, pack) keys -- never silently drop
    # them from the matrix / totals.
    _valid_cols = {(b, p) for b in BRAND_ORDER for p in PACKS_BY_BRAND[b]}
    _bad = [f"{wh}:{k}" for wh, breakup in data.items()
            for k in breakup if tuple(k) not in _valid_cols]
    _bad += [f"physical_stock:{k}" for k in physical_stock
             if tuple(k) not in _valid_cols]
    if _bad:
        raise ValueError(
            "Non-canonical (brand, pack) key(s) would be silently dropped -- "
            "canonicalise first (brands: " + "/".join(BRAND_ORDER) +
            "; packs: NIP/PINT/HL/L): " + ", ".join(map(str, _bad)))

    # Clusters (if given) must be an EXACT partition of data's warehouses --
    # never silently drop a depot or place one in two clusters.
    if clusters:
        flat = [wh for _, whs in clusters for wh in whs]
        missing = [wh for wh in data if wh not in flat]
        extra = [wh for wh in flat if wh not in data]
        dup = sorted({wh for wh in flat if flat.count(wh) > 1})
        if missing or extra or dup:
            raise ValueError(
                "clusters must partition data's warehouses exactly: "
                f"missing={missing} extra={extra} dup={dup}")

    wb = Workbook()
    layout = _build_matrix_sheet(wb.active, data, compile_date, round_overrides,
                                 physical_stock, clusters)
    _build_brand_pack_matrix_sheet(
        wb.create_sheet("BRAND × PACK MATRIX", 1),
        data, compile_date, layout)
    _build_blend_breakup_sheet(wb.create_sheet("BLEND BREAKUP"),
                                data, compile_date, layout)
    wb.active = 0
    wb.save(output_path)


# ---------------------------------------------------------------------------
# Sheet 1 -- WAREHOUSE REQUIREMENT
# ---------------------------------------------------------------------------
def _build_matrix_sheet(ws, data, compile_date, round_overrides, physical_stock,
                        clusters=None):
    ws.title = "WAREHOUSE REQUIREMENT"

    warehouses = list(data.keys())
    columns = [(b, p) for b in BRAND_ORDER for p in PACKS_BY_BRAND[b]]

    n_data_cols = len(columns)
    last_data_col = 1 + n_data_cols
    total_col = last_data_col + 1
    last_letter = get_column_letter(total_col)
    data_first_letter = get_column_letter(2)
    data_last_letter = get_column_letter(last_data_col)

    # Brand-group end columns (right edge of each brand's pack span)
    brand_end_cols = set()
    col = 2
    for brand in BRAND_ORDER:
        brand_end_cols.add(col + len(PACKS_BY_BRAND[brand]) - 1)
        col += len(PACKS_BY_BRAND[brand])

    hair_side    = Side(style="thin", color=HAIRLINE)
    gold_div     = Side(style="thin", color=GOLD_DIV)
    gold_medium  = Side(style="medium", color=GOLD_DIV)

    def data_border(col_idx):
        right = gold_div if col_idx in brand_end_cols else hair_side
        return Border(left=hair_side, right=right,
                      top=hair_side, bottom=hair_side)

    # --- Row 1: Title bar ---
    ws.merge_cells(f"A1:{last_letter}1")
    c = ws["A1"]
    c.value = "K.S. DISTILLERY    •    WAREHOUSE REQUIREMENT BREAKUP"
    c.font = Font(name="Arial", bold=True, size=18, color=WHITE)
    c.fill = PatternFill("solid", start_color=NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 44

    # --- Row 2: Subtitle ---
    ws.merge_cells(f"A2:{last_letter}2")
    c = ws["A2"]
    c.value = (f"Warehouse-wise dispatch requirement   |   "
               f"Compiled {compile_date}   |   All quantities in CASES")
    c.font = Font(name="Arial", italic=True, size=10, color=WHITE)
    c.fill = PatternFill("solid", start_color=NAVY_SOFT)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    # Rows 3-4: blank spacer
    for r in (3, 4):
        ws.row_dimensions[r].height = 8
        for cc in range(1, total_col + 1):
            ws.cell(row=r, column=cc).fill = PatternFill("solid", start_color=WHITE)

    # --- Row 5: Brand banner ---
    ws.merge_cells(start_row=5, start_column=1, end_row=6, end_column=1)
    wh_head = ws.cell(row=5, column=1, value="WAREHOUSE")
    wh_head.font = Font(name="Arial", bold=True, size=11, color=WHITE)
    wh_head.fill = PatternFill("solid", start_color=NAVY)
    wh_head.alignment = Alignment(horizontal="center", vertical="center")
    wh_head.border = Border(right=gold_div, top=gold_medium, bottom=gold_medium)

    col = 2
    for brand in BRAND_ORDER:
        n = len(PACKS_BY_BRAND[brand])
        if n > 1:
            ws.merge_cells(start_row=5, start_column=col,
                           end_row=5, end_column=col + n - 1)
        cell = ws.cell(row=5, column=col, value=brand.upper())
        cell.font = Font(name="Arial", bold=True, size=13, color=WHITE)
        cell.fill = PatternFill("solid", start_color=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(right=gold_div, top=gold_medium)
        col += n

    ws.merge_cells(start_row=5, start_column=total_col,
                   end_row=6, end_column=total_col)
    cell = ws.cell(row=5, column=total_col, value="TOTAL\n(cases)")
    cell.font = Font(name="Arial", bold=True, size=11, color=WHITE)
    cell.fill = PatternFill("solid", start_color=DARK_GREY)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(top=gold_medium, bottom=gold_medium, left=gold_div)
    ws.row_dimensions[5].height = 34

    # --- Row 6: Pack header ---
    col = 2
    for brand in BRAND_ORDER:
        for pack in PACKS_BY_BRAND[brand]:
            cell = ws.cell(row=6, column=col, value=PACK_LABEL[pack])
            cell.font = Font(name="Arial", bold=True, size=9, color=WHITE)
            cell.fill = PatternFill("solid", start_color=NAVY_SOFT)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            right = gold_div if col in brand_end_cols else hair_side
            cell.border = Border(top=hair_side, bottom=gold_medium, right=right)
            col += 1
    ws.row_dimensions[6].height = 28

    # --- Data rows (optionally grouped into cluster banners) ---
    data_start_row = 7

    if clusters:
        groups = [(lbl, [w for w in whs]) for lbl, whs in clusters]
    else:
        groups = [(None, warehouses)]

    # Assign row positions: each non-empty labelled group gets a banner row,
    # then its warehouse rows beneath it. Banners are rendered AFTER the
    # warehouse rows so their SUM ranges are known.
    row = data_start_row
    banner_spans = []        # (banner_row, first_wh_row, last_wh_row, label)
    wh_positions = []        # (warehouse_name, row, global_idx) in render order
    g_idx = 0
    for label, whs in groups:
        if not whs:
            continue
        first_wh = row
        for wh in whs:
            wh_positions.append((wh, row, g_idx))
            g_idx += 1
            row += 1
        last_wh = row - 1
        if label is not None:
            banner_r = row          # subtotal row sits AFTER its depots
            row += 1
            banner_spans.append((banner_r, first_wh, last_wh, label))
    total_row = row

    # Warehouse rows
    for wh, r, idx in wh_positions:
        stripe_color = ZEBRA if idx % 2 == 1 else WHITE
        stripe = PatternFill("solid", start_color=stripe_color)

        name_cell = ws.cell(row=r, column=1, value=wh.upper())
        name_cell.font = Font(name="Arial", bold=True, size=11, color=INK)
        name_cell.fill = PatternFill("solid",
                                     start_color=WH_BAND if idx % 2 == 0 else WH_BAND_ALT)
        name_cell.alignment = Alignment(horizontal="left", vertical="center", indent=2)
        name_cell.border = Border(right=gold_div, top=hair_side, bottom=hair_side)

        for j, (brand, pack) in enumerate(columns):
            col_idx = 2 + j
            qty = data[wh].get((brand, pack), 0)
            cell = ws.cell(row=r, column=col_idx, value=qty if qty else None)
            if qty:
                cell.font = Font(name="Arial", bold=True, size=11, color=INK)
            else:
                cell.font = Font(name="Arial", size=10, color=SOFT_INK)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.number_format = "#,##0;-#,##0;"
            cell.fill = stripe
            cell.border = data_border(col_idx)

        total_cell = ws.cell(row=r, column=total_col,
                             value=f"=SUM({data_first_letter}{r}:{data_last_letter}{r})")
        total_cell.font = Font(name="Arial", bold=True, size=12, color=WHITE)
        total_cell.fill = PatternFill("solid", start_color=DARK_GREY)
        total_cell.alignment = Alignment(horizontal="center", vertical="center")
        total_cell.number_format = "#,##0"
        total_cell.border = Border(left=gold_div, right=hair_side,
                                   top=hair_side, bottom=hair_side)

        ws.row_dimensions[r].height = 28

    # Cluster subtotal rows -- navy, placed AFTER each cluster's depots,
    # carrying live SUM subtotals of the depots above.
    for banner_r, first_wh, last_wh, label in banner_spans:
        lab = ws.cell(row=banner_r, column=1, value=f"{str(label).upper()}  TOTAL")
        lab.font = Font(name="Arial", bold=True, size=11, color=GOLD)
        lab.fill = PatternFill("solid", start_color=CLUSTER_FILL)
        lab.alignment = Alignment(horizontal="left", vertical="center", indent=2)
        lab.border = Border(top=gold_medium, bottom=gold_medium, right=gold_div)
        for j in range(n_data_cols):
            col_idx = 2 + j
            cl = get_column_letter(col_idx)
            cell = ws.cell(row=banner_r, column=col_idx,
                           value=f"=SUM({cl}{first_wh}:{cl}{last_wh})")
            cell.font = Font(name="Arial", bold=True, size=11, color=WHITE)
            cell.fill = PatternFill("solid", start_color=CLUSTER_FILL)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.number_format = "#,##0;-#,##0;"
            right = gold_div if col_idx in brand_end_cols else None
            cell.border = Border(top=gold_medium, bottom=gold_medium, right=right)
        bt = ws.cell(row=banner_r, column=total_col,
                     value=f"=SUM({last_letter}{first_wh}:{last_letter}{last_wh})")
        bt.font = Font(name="Arial", bold=True, size=12, color=GOLD)
        bt.fill = PatternFill("solid", start_color=CLUSTER_FILL)
        bt.alignment = Alignment(horizontal="center", vertical="center")
        bt.number_format = "#,##0"
        bt.border = Border(top=gold_medium, bottom=gold_medium, left=gold_div)
        ws.row_dimensions[banner_r].height = 30

    # --- TOTAL REQUIREMENT row ---
    def _grand(col_letter):
        # Sum the cluster banners (no double-count) when clustered, else the
        # contiguous warehouse range.
        if banner_spans:
            return "=" + "+".join(f"{col_letter}{b[0]}" for b in banner_spans)
        return f"=SUM({col_letter}{data_start_row}:{col_letter}{total_row - 1})"

    cell = ws.cell(row=total_row, column=1, value="TOTAL  REQUIREMENT")
    cell.font = Font(name="Arial", bold=True, size=12, color=GOLD)
    cell.fill = PatternFill("solid", start_color=NAVY)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=2)
    cell.border = Border(top=gold_medium, bottom=gold_medium, right=gold_div)

    for j in range(n_data_cols):
        col_idx = 2 + j
        col_letter = get_column_letter(col_idx)
        cell = ws.cell(row=total_row, column=col_idx, value=_grand(col_letter))
        cell.font = Font(name="Arial", bold=True, size=12, color=WHITE)
        cell.fill = PatternFill("solid", start_color=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.number_format = "#,##0;-#,##0;"
        right = gold_div if col_idx in brand_end_cols else None
        cell.border = Border(top=gold_medium, bottom=gold_medium, right=right)

    gt = ws.cell(row=total_row, column=total_col, value=_grand(last_letter))
    gt.font = Font(name="Arial", bold=True, size=15, color=GOLD)
    gt.fill = PatternFill("solid", start_color=NAVY)
    gt.alignment = Alignment(horizontal="center", vertical="center")
    gt.number_format = "#,##0"
    gt.border = Border(top=gold_medium, bottom=gold_medium, left=gold_div)
    ws.row_dimensions[total_row].height = 38

    # --- LESS: PHYSICAL STOCK AT FACTORY row (manual / pre-filled values) ---
    stock_row = total_row + 1
    STOCK_FILL  = "FFFFF6E0"   # very pale cream (input cell)
    STOCK_LABEL = "FFF7E7B8"   # warmer cream for the label
    STOCK_INK   = "FFB45309"   # amber for stock numbers
    label = ws.cell(row=stock_row, column=1, value="LESS:  PHYSICAL  STOCK")
    label.font = Font(name="Arial", bold=True, size=11, color=STOCK_INK)
    label.fill = PatternFill("solid", start_color=STOCK_LABEL)
    label.alignment = Alignment(horizontal="left", vertical="center", indent=2)
    label.border = Border(top=hair_side, bottom=hair_side, right=gold_div)

    for j, (brand, pack) in enumerate(columns):
        col_idx = 2 + j
        qty = physical_stock.get((brand, pack), 0)
        cell = ws.cell(row=stock_row, column=col_idx, value=qty if qty else None)
        if qty:
            cell.font = Font(name="Arial", bold=True, size=11, color=STOCK_INK)
        else:
            cell.font = Font(name="Arial", size=10, color=SOFT_INK)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.number_format = "#,##0;-#,##0;"
        cell.fill = PatternFill("solid", start_color=STOCK_FILL)
        cell.border = data_border(col_idx)

    stock_total = ws.cell(row=stock_row, column=total_col,
                          value=f"=SUM({data_first_letter}{stock_row}:{data_last_letter}{stock_row})")
    stock_total.font = Font(name="Arial", bold=True, size=12, color=STOCK_INK)
    stock_total.fill = PatternFill("solid", start_color=STOCK_FILL)
    stock_total.alignment = Alignment(horizontal="center", vertical="center")
    stock_total.number_format = "#,##0"
    stock_total.border = Border(left=gold_div, right=hair_side,
                                top=hair_side, bottom=hair_side)
    ws.row_dimensions[stock_row].height = 28

    # --- PRODUCTION TO BE TAKEN row -- Requirement - Stock, floor 0 ---
    production_row = total_row + 2
    cell = ws.cell(row=production_row, column=1, value="PRODUCTION  TO  BE  TAKEN")
    cell.font = Font(name="Arial", bold=True, size=12, color=GOLD)
    cell.fill = PatternFill("solid", start_color=NAVY)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=2)
    cell.border = Border(top=gold_medium, bottom=gold_medium, right=gold_div)

    for j in range(n_data_cols):
        col_idx = 2 + j
        col_letter = get_column_letter(col_idx)
        cell = ws.cell(row=production_row, column=col_idx,
                       value=f"=MAX(0,{col_letter}{total_row}-{col_letter}{stock_row})")
        cell.font = Font(name="Arial", bold=True, size=12, color=WHITE)
        cell.fill = PatternFill("solid", start_color=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.number_format = "#,##0;-#,##0;"
        right = gold_div if col_idx in brand_end_cols else None
        cell.border = Border(top=gold_medium, bottom=gold_medium, right=right)

    pgt = ws.cell(row=production_row, column=total_col,
                  value=f"=SUM({data_first_letter}{production_row}:{data_last_letter}{production_row})")
    pgt.font = Font(name="Arial", bold=True, size=15, color=GOLD)
    pgt.fill = PatternFill("solid", start_color=NAVY)
    pgt.alignment = Alignment(horizontal="center", vertical="center")
    pgt.number_format = "#,##0"
    pgt.border = Border(top=gold_medium, bottom=gold_medium, left=gold_div)
    ws.row_dimensions[production_row].height = 38

    foot_row = production_row + 2
    ws.row_dimensions[foot_row].height = 8

    # ----------- BLEND & ENA REQUIREMENT section -----------
    brand_col_ranges = {}
    col = 2
    for brand in BRAND_ORDER:
        cols = []
        for pack in PACKS_BY_BRAND[brand]:
            cols.append((col, pack))
            col += 1
        brand_col_ranges[brand] = cols

    calc_top    = foot_row + 2
    banner_row  = calc_top
    header_row2 = calc_top + 1
    blend_row   = calc_top + 2
    round_row   = calc_top + 3
    ena_row     = calc_top + 4

    ws.row_dimensions[foot_row + 1].height = 12

    ws.merge_cells(start_row=banner_row, start_column=1,
                   end_row=banner_row, end_column=total_col)
    cell = ws.cell(row=banner_row, column=1)
    cell.value = "BLEND  &  ENA  REQUIREMENT"
    cell.font = Font(name="Arial", bold=True, size=13, color=WHITE)
    cell.fill = PatternFill("solid", start_color=NAVY)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = Border(top=gold_medium, bottom=gold_medium)
    ws.row_dimensions[banner_row].height = 32

    hcell = ws.cell(row=header_row2, column=1, value="")
    hcell.fill = PatternFill("solid", start_color=NAVY)
    hcell.border = Border(right=gold_div, bottom=gold_medium)
    for brand, cols in brand_col_ranges.items():
        first_col = cols[0][0]
        last_col_b = cols[-1][0]
        if first_col != last_col_b:
            ws.merge_cells(start_row=header_row2, start_column=first_col,
                           end_row=header_row2, end_column=last_col_b)
        cell = ws.cell(row=header_row2, column=first_col, value=brand.upper())
        cell.font = Font(name="Arial", bold=True, size=11, color=WHITE)
        cell.fill = PatternFill("solid", start_color=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(right=gold_div, bottom=gold_medium)

    cell = ws.cell(row=header_row2, column=total_col, value="TOTAL")
    cell.font = Font(name="Arial", bold=True, size=11, color=WHITE)
    cell.fill = PatternFill("solid", start_color=DARK_GREY)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = Border(left=gold_div, bottom=gold_medium)
    ws.row_dimensions[header_row2].height = 26

    def _label_cell(row, text, big=False):
        cell = ws.cell(row=row, column=1, value=text)
        cell.font = Font(name="Arial", bold=True,
                         size=11 if not big else 12, color=INK)
        cell.fill = PatternFill("solid", start_color=WH_BAND)
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=2)
        cell.border = Border(right=gold_div, top=hair_side, bottom=hair_side)
        return cell

    # BLEND (L) row -- driven by PRODUCTION TO BE TAKEN row (post-stock)
    _label_cell(blend_row, "BLEND  (L)")
    for brand, cols in brand_col_ranges.items():
        parts = []
        for col_idx, pack in cols:
            col_letter = get_column_letter(col_idx)
            parts.append(f"{col_letter}{production_row}*{BLEND_FACTOR[pack]}")
        formula = "=" + "+".join(parts)
        first_col = cols[0][0]
        last_col_b = cols[-1][0]
        if first_col != last_col_b:
            ws.merge_cells(start_row=blend_row, start_column=first_col,
                           end_row=blend_row, end_column=last_col_b)
        cell = ws.cell(row=blend_row, column=first_col, value=formula)
        cell.font = Font(name="Arial", bold=True, size=11, color=INK)
        cell.fill = PatternFill("solid", start_color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.number_format = "#,##0.00"
        cell.border = Border(top=hair_side, bottom=hair_side, right=gold_div)

    blend_sum_parts = [get_column_letter(cols[0][0]) + str(blend_row)
                        for cols in brand_col_ranges.values()]
    tot_blend_cell = ws.cell(row=blend_row, column=total_col,
                             value="=" + "+".join(blend_sum_parts))
    tot_blend_cell.font = Font(name="Arial", bold=True, size=12, color=WHITE)
    tot_blend_cell.fill = PatternFill("solid", start_color=DARK_GREY)
    tot_blend_cell.alignment = Alignment(horizontal="center", vertical="center")
    tot_blend_cell.number_format = "#,##0.00"
    tot_blend_cell.border = Border(left=gold_div, top=hair_side, bottom=hair_side)
    ws.row_dimensions[blend_row].height = 28

    # BLEND ROUNDED (L) row
    _label_cell(round_row, "BLEND  ROUNDED  (L)")
    for brand, cols in brand_col_ranges.items():
        first_col = cols[0][0]
        last_col_b = cols[-1][0]
        first_letter = get_column_letter(first_col)
        if first_col != last_col_b:
            ws.merge_cells(start_row=round_row, start_column=first_col,
                           end_row=round_row, end_column=last_col_b)
        if brand in round_overrides:
            cell = ws.cell(row=round_row, column=first_col,
                           value=round_overrides[brand])
        else:
            cell = ws.cell(row=round_row, column=first_col,
                           value=f"=CEILING({first_letter}{blend_row},{ROUND_TO})")
        cell.font = Font(name="Arial", bold=True, size=11, color=INK)
        cell.fill = PatternFill("solid", start_color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.number_format = "#,##0"
        cell.border = Border(top=hair_side, bottom=hair_side, right=gold_div)

    round_sum_parts = [get_column_letter(cols[0][0]) + str(round_row)
                        for cols in brand_col_ranges.values()]
    tot_round_cell = ws.cell(row=round_row, column=total_col,
                             value="=" + "+".join(round_sum_parts))
    tot_round_cell.font = Font(name="Arial", bold=True, size=12, color=WHITE)
    tot_round_cell.fill = PatternFill("solid", start_color=DARK_GREY)
    tot_round_cell.alignment = Alignment(horizontal="center", vertical="center")
    tot_round_cell.number_format = "#,##0"
    tot_round_cell.border = Border(left=gold_div, top=hair_side, bottom=hair_side)
    ws.row_dimensions[round_row].height = 28

    # ENA REQUIREMENT (L) row -- per brand = Blend Rounded x (0.75 / 1.68)
    _label_cell(ena_row, "ENA  REQUIREMENT  (L)", big=True)
    ena_first_letters = []
    for brand, cols in brand_col_ranges.items():
        first_col = cols[0][0]
        last_col_b = cols[-1][0]
        first_letter = get_column_letter(first_col)
        if first_col != last_col_b:
            ws.merge_cells(start_row=ena_row, start_column=first_col,
                           end_row=ena_row, end_column=last_col_b)
        cell = ws.cell(row=ena_row, column=first_col,
                       value=f"={first_letter}{round_row}*({ENA_NUM}/{ENA_DEN})")
        cell.font = Font(name="Arial", bold=True, size=11, color=INK)
        cell.fill = PatternFill("solid", start_color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.number_format = "#,##0.00"
        cell.border = Border(top=hair_side, bottom=gold_medium, right=gold_div)
        ena_first_letters.append(first_letter)

    tot_ena_cell = ws.cell(row=ena_row, column=total_col,
                           value="=" + "+".join(f"{c}{ena_row}" for c in ena_first_letters))
    tot_ena_cell.font = Font(name="Arial", bold=True, size=14, color=GOLD)
    tot_ena_cell.fill = PatternFill("solid", start_color=NAVY)
    tot_ena_cell.alignment = Alignment(horizontal="center", vertical="center")
    tot_ena_cell.number_format = "#,##0.00"
    tot_ena_cell.border = Border(left=gold_div, top=hair_side, bottom=gold_medium)
    ws.row_dimensions[ena_row].height = 34

    # ----------- column widths + sheet props -----------
    ws.column_dimensions["A"].width = 36
    for j in range(n_data_cols):
        ws.column_dimensions[get_column_letter(2 + j)].width = 9.5
    ws.column_dimensions[get_column_letter(total_col)].width = 13

    ws.freeze_panes = "B7"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5

    # Layout positions the BLEND BREAKUP sheet needs to cross-reference.
    return {
        "total_row": total_row,
        "total_col_letter": last_letter,
        "data_first_letter": data_first_letter,
        "data_last_letter": data_last_letter,
    }


# ---------------------------------------------------------------------------
# Sheet 2 -- BLEND BREAKUP
# ---------------------------------------------------------------------------
def _build_blend_breakup_sheet(ws, data, compile_date, layout):
    warehouses = list(data.keys())

    # Compute brand totals (drives sort order; cells stay live formulas)
    brand_totals = {}
    for brand in BRAND_ORDER:
        t = 0
        for wh in warehouses:
            for pack in PACKS_BY_BRAND[brand]:
                t += data[wh].get((brand, pack), 0)
        brand_totals[brand] = t
    brands_ranked = sorted(BRAND_ORDER, key=lambda b: -brand_totals[b])

    # Brand-col map for cross-sheet references
    brand_col_map = {}
    col = 2
    for brand in BRAND_ORDER:
        n = len(PACKS_BY_BRAND[brand])
        first = get_column_letter(col)
        last = get_column_letter(col + n - 1)
        brand_col_map[brand] = (first, last)
        col += n
    # Cross-reference the TOTAL REQUIREMENT row (not the raw warehouse rows) so
    # brand cases are correct whether or not cluster banner rows are present --
    # the TOTAL row already nets the banners, so this can never double-count.
    total_col_letter = layout["total_col_letter"]
    total_row_main = layout["total_row"]

    main = "'WAREHOUSE REQUIREMENT'"
    grand_ref = f"{main}!{total_col_letter}{total_row_main}"

    hair = Side(style="thin", color=HAIRLINE)
    gold_med = Side(style="medium", color=GOLD_RULE)

    ws.sheet_view.showGridLines = False

    # Row 1: wordmark
    ws.merge_cells("B1:D1")
    c = ws["B1"]
    c.value = "K.S. DISTILLERY"
    c.font = Font(name="Arial", bold=True, size=22, color=NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 38

    # Row 2: subtitle + date
    ws["B2"] = "Brand-wise Blend Breakup"
    ws["B2"].font = Font(name="Arial", italic=True, size=12, color=SOFT_INK)
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("C2:D2")
    ws["C2"] = f"Compiled  {compile_date}"
    ws["C2"].font = Font(name="Arial", italic=True, size=10, color=SOFT_INK)
    ws["C2"].alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[2].height = 22

    # Row 3: gold rule
    for col in range(2, 5):
        ws.cell(row=3, column=col).border = Border(bottom=gold_med)
    ws.row_dimensions[3].height = 8
    ws.row_dimensions[4].height = 10

    # Header row
    HEADER_ROW = 5
    ws.cell(row=HEADER_ROW, column=2, value="BRAND")
    ws.cell(row=HEADER_ROW, column=3, value="CASES")
    ws.cell(row=HEADER_ROW, column=4, value="SHARE")
    for col, align in [(2, "left"), (3, "right"), (4, "right")]:
        cell = ws.cell(row=HEADER_ROW, column=col)
        cell.font = Font(name="Arial", bold=True, size=11, color=WHITE)
        cell.fill = PatternFill("solid", start_color=NAVY)
        cell.alignment = Alignment(horizontal=align, vertical="center", indent=1)
        cell.border = Border(top=hair, bottom=hair, left=hair, right=hair)
    ws.row_dimensions[HEADER_ROW].height = 28

    # Brand rows
    DATA_START = HEADER_ROW + 1
    for i, brand in enumerate(brands_ranked):
        r = DATA_START + i
        stripe = PatternFill("solid", start_color=ROW_TINT) if i % 2 == 1 else None

        bcell = ws.cell(row=r, column=2, value=brand.upper())
        bcell.font = Font(name="Arial", bold=True, size=12, color=INK)
        bcell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        bcell.border = Border(top=hair, bottom=hair, left=hair, right=hair)
        if stripe:
            bcell.fill = stripe

        first, last = brand_col_map[brand]
        case_formula = (f"=SUM({main}!{first}{total_row_main}:"
                        f"{last}{total_row_main})")
        ccell = ws.cell(row=r, column=3, value=case_formula)
        ccell.font = Font(name="Arial", bold=True, size=12, color=INK)
        ccell.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        ccell.number_format = "#,##0"
        ccell.border = Border(top=hair, bottom=hair, left=hair, right=hair)
        if stripe:
            ccell.fill = stripe

        pcell = ws.cell(row=r, column=4, value=f"=C{r}/{grand_ref}")
        pcell.font = Font(name="Arial", size=12, color=SOFT_INK)
        pcell.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        pcell.number_format = "0.0%"
        pcell.border = Border(top=hair, bottom=hair, left=hair, right=hair)
        if stripe:
            pcell.fill = stripe

        ws.row_dimensions[r].height = 26

    # TOTAL row
    total_row_2 = DATA_START + len(BRAND_ORDER)
    ws.cell(row=total_row_2, column=2, value="TOTAL")
    ws.cell(row=total_row_2, column=3, value=f"={grand_ref}")
    ws.cell(row=total_row_2, column=4,
             value=f"=SUM(D{DATA_START}:D{total_row_2 - 1})")
    for col, align in [(2, "left"), (3, "right"), (4, "right")]:
        cell = ws.cell(row=total_row_2, column=col)
        cell.font = Font(name="Arial", bold=True, size=12, color=WHITE)
        cell.fill = PatternFill("solid", start_color=NAVY)
        cell.alignment = Alignment(horizontal=align, vertical="center", indent=1)
        cell.border = Border(top=hair, bottom=hair, left=hair, right=hair)
    ws.cell(row=total_row_2, column=3).number_format = "#,##0"
    ws.cell(row=total_row_2, column=4).number_format = "0.0%"
    ws.row_dimensions[total_row_2].height = 30

    # ---- PACK-WISE BREAKUP (second table, stacked below the brand table) ----
    # Each pack's cases = sum of that pack's column across all 7 brands on the
    # matrix TOTAL REQUIREMENT row -> live + always reconciles to the grand total.
    PACK_KEYS = list(_ALL_PACKS)                       # NIP, PINT, HL, L
    brand_starts, _c = [], 2
    for brand in BRAND_ORDER:
        brand_starts.append(_c)
        _c += len(PACKS_BY_BRAND[brand])
    pack_cols = {pk: [get_column_letter(bs + i) for bs in brand_starts]
                 for i, pk in enumerate(PACK_KEYS)}
    pack_totals = {pk: sum(data[wh].get((b, pk), 0)
                           for wh in warehouses for b in BRAND_ORDER)
                   for pk in PACK_KEYS}
    packs_ranked = sorted(PACK_KEYS, key=lambda p: -pack_totals[p])

    pk_head_row = total_row_2 + 2
    hcell = ws.cell(row=pk_head_row, column=2, value="Pack-wise Blend Breakup")
    hcell.font = Font(name="Arial", italic=True, size=12, color=SOFT_INK)
    hcell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[pk_head_row].height = 22
    for col in range(2, 5):
        ws.cell(row=pk_head_row, column=col).border = Border(bottom=gold_med)

    pk_header_row = pk_head_row + 1
    for col, txt, align in [(2, "PACK", "left"), (3, "CASES", "right"),
                            (4, "SHARE", "right")]:
        cell = ws.cell(row=pk_header_row, column=col, value=txt)
        cell.font = Font(name="Arial", bold=True, size=11, color=WHITE)
        cell.fill = PatternFill("solid", start_color=NAVY)
        cell.alignment = Alignment(horizontal=align, vertical="center", indent=1)
        cell.border = Border(top=hair, bottom=hair, left=hair, right=hair)
    ws.row_dimensions[pk_header_row].height = 28

    PK_DATA = pk_header_row + 1
    for i, pk in enumerate(packs_ranked):
        r = PK_DATA + i
        stripe = PatternFill("solid", start_color=ROW_TINT) if i % 2 == 1 else None
        pcell = ws.cell(row=r, column=2, value=PACK_LABEL[pk])
        pcell.font = Font(name="Arial", bold=True, size=12, color=INK)
        pcell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        pcell.border = Border(top=hair, bottom=hair, left=hair, right=hair)
        if stripe:
            pcell.fill = stripe
        case_f = "=" + "+".join(f"{main}!{c}{total_row_main}" for c in pack_cols[pk])
        ccell = ws.cell(row=r, column=3, value=case_f)
        ccell.font = Font(name="Arial", bold=True, size=12, color=INK)
        ccell.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        ccell.number_format = "#,##0"
        ccell.border = Border(top=hair, bottom=hair, left=hair, right=hair)
        if stripe:
            ccell.fill = stripe
        scell = ws.cell(row=r, column=4, value=f"=C{r}/{grand_ref}")
        scell.font = Font(name="Arial", size=12, color=SOFT_INK)
        scell.alignment = Alignment(horizontal="right", vertical="center", indent=1)
        scell.number_format = "0.0%"
        scell.border = Border(top=hair, bottom=hair, left=hair, right=hair)
        if stripe:
            scell.fill = stripe
        ws.row_dimensions[r].height = 26

    pk_total_row = PK_DATA + len(PACK_KEYS)
    ws.cell(row=pk_total_row, column=2, value="TOTAL")
    ws.cell(row=pk_total_row, column=3, value=f"={grand_ref}")
    ws.cell(row=pk_total_row, column=4,
            value=f"=SUM(D{PK_DATA}:D{pk_total_row - 1})")
    for col, align in [(2, "left"), (3, "right"), (4, "right")]:
        cell = ws.cell(row=pk_total_row, column=col)
        cell.font = Font(name="Arial", bold=True, size=12, color=WHITE)
        cell.fill = PatternFill("solid", start_color=NAVY)
        cell.alignment = Alignment(horizontal=align, vertical="center", indent=1)
        cell.border = Border(top=hair, bottom=hair, left=hair, right=hair)
    ws.cell(row=pk_total_row, column=3).number_format = "#,##0"
    ws.cell(row=pk_total_row, column=4).number_format = "0.0%"
    ws.row_dimensions[pk_total_row].height = 30

    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 27
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 4

    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True


# ---------------------------------------------------------------------------
# Sheet -- BRAND × PACK MATRIX  (production-manager aggregate view)
# ---------------------------------------------------------------------------
def _build_brand_pack_matrix_sheet(ws, data, compile_date, layout):
    """Brands down the rows, pack sizes across the columns; each cell is the
    network-wide requirement (cases) for that brand × pack, pulled LIVE from
    the WAREHOUSE REQUIREMENT sheet's TOTAL REQUIREMENT row -- so it always
    reconciles to the warehouse matrix and re-totals on any edit."""
    main = "'WAREHOUSE REQUIREMENT'"
    tr = layout["total_row"]
    PACK_KEYS = list(_ALL_PACKS)                       # NIP PINT HL L

    def col1(bi, pj):                                  # sheet-1 col for brand bi, pack pj
        return get_column_letter(2 + 4 * bi + pj)

    hair = Side(style="thin", color=HAIRLINE)
    gdiv = Side(style="thin", color=GOLD_DIV)
    gmed = Side(style="medium", color=GOLD_DIV)

    n_cols = 1 + len(PACK_KEYS) + 1                     # brand + 4 packs + total
    last = get_column_letter(n_cols)
    tot_col = n_cols
    first_pack_L = get_column_letter(2)
    last_pack_L = get_column_letter(1 + len(PACK_KEYS))

    # Row 1 -- title
    ws.merge_cells(f"A1:{last}1")
    c = ws["A1"]; c.value = "K.S. DISTILLERY    •    BRAND × PACK REQUIREMENT MATRIX"
    c.font = Font(name="Arial", bold=True, size=16, color=WHITE)
    c.fill = PatternFill("solid", start_color=NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40
    # Row 2 -- subtitle
    ws.merge_cells(f"A2:{last}2")
    c = ws["A2"]; c.value = (f"Aggregate requirement by brand & pack size   |   "
                             f"Compiled {compile_date}   |   All quantities in CASES")
    c.font = Font(name="Arial", italic=True, size=10, color=WHITE)
    c.fill = PatternFill("solid", start_color=NAVY_SOFT)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 8

    # Row 4 -- header
    hdr = 4
    head = [("BRAND", "left")] + [(PACK_LABEL[p], "center") for p in PACK_KEYS] \
        + [("TOTAL (cases)", "center")]
    for i, (txt, al) in enumerate(head):
        cell = ws.cell(row=hdr, column=1 + i, value=txt)
        cell.font = Font(name="Arial", bold=True, size=11, color=WHITE)
        cell.fill = PatternFill("solid",
                                start_color=NAVY if i in (0, len(head) - 1) else NAVY_SOFT)
        cell.alignment = Alignment(horizontal=al, vertical="center",
                                   indent=1 if al == "left" else 0)
        cell.border = Border(top=gmed, bottom=gmed,
                             right=gdiv if i in (0, len(head) - 2) else hair)
    ws.row_dimensions[hdr].height = 30

    # Brand rows
    start = hdr + 1
    for bi, brand in enumerate(BRAND_ORDER):
        r = start + bi
        stripe = ZEBRA if bi % 2 == 1 else WHITE
        bcell = ws.cell(row=r, column=1, value=brand.upper())
        bcell.font = Font(name="Arial", bold=True, size=11, color=INK)
        bcell.fill = PatternFill("solid",
                                 start_color=WH_BAND if bi % 2 == 0 else WH_BAND_ALT)
        bcell.alignment = Alignment(horizontal="left", vertical="center", indent=2)
        bcell.border = Border(left=hair, right=gdiv, top=hair, bottom=hair)
        for pj in range(len(PACK_KEYS)):
            cc = ws.cell(row=r, column=2 + pj,
                         value=f"={main}!{col1(bi, pj)}{tr}")
            cc.font = Font(name="Arial", bold=True, size=11, color=INK)
            cc.fill = PatternFill("solid", start_color=stripe)
            cc.alignment = Alignment(horizontal="center", vertical="center")
            cc.number_format = "#,##0;-#,##0;"
            cc.border = Border(left=hair, right=hair, top=hair, bottom=hair)
        tc = ws.cell(row=r, column=tot_col,
                     value=f"=SUM({first_pack_L}{r}:{last_pack_L}{r})")
        tc.font = Font(name="Arial", bold=True, size=12, color=WHITE)
        tc.fill = PatternFill("solid", start_color=DARK_GREY)
        tc.alignment = Alignment(horizontal="center", vertical="center")
        tc.number_format = "#,##0"
        tc.border = Border(left=gdiv, right=hair, top=hair, bottom=hair)
        ws.row_dimensions[r].height = 26

    # TOTAL row
    trow = start + len(BRAND_ORDER)
    lab = ws.cell(row=trow, column=1, value="TOTAL")
    lab.font = Font(name="Arial", bold=True, size=12, color=GOLD)
    lab.fill = PatternFill("solid", start_color=NAVY)
    lab.alignment = Alignment(horizontal="left", vertical="center", indent=2)
    lab.border = Border(top=gmed, bottom=gmed, right=gdiv)
    for pj in range(len(PACK_KEYS)):
        L = get_column_letter(2 + pj)
        cc = ws.cell(row=trow, column=2 + pj, value=f"=SUM({L}{start}:{L}{trow - 1})")
        cc.font = Font(name="Arial", bold=True, size=12, color=WHITE)
        cc.fill = PatternFill("solid", start_color=NAVY)
        cc.alignment = Alignment(horizontal="center", vertical="center")
        cc.number_format = "#,##0"
        cc.border = Border(top=gmed, bottom=gmed, right=hair)
    gc = ws.cell(row=trow, column=tot_col,
                 value=f"=SUM({first_pack_L}{trow}:{last_pack_L}{trow})")
    gc.font = Font(name="Arial", bold=True, size=14, color=GOLD)
    gc.fill = PatternFill("solid", start_color=NAVY)
    gc.alignment = Alignment(horizontal="center", vertical="center")
    gc.number_format = "#,##0"
    gc.border = Border(top=gmed, bottom=gmed, left=gdiv)
    ws.row_dimensions[trow].height = 32

    ws.column_dimensions["A"].width = 24
    for pj in range(len(PACK_KEYS)):
        ws.column_dimensions[get_column_letter(2 + pj)].width = 13
    ws.column_dimensions[get_column_letter(tot_col)].width = 15
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B5"
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True


# ---------------------------------------------------------------------------
# CLI entry -- smoke test when running the script standalone.
# Callers should normally import build_workbook directly.
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import sys
    import json
    if len(sys.argv) != 3:
        print("Usage: build_warehouse_requirement.py DATA.json OUTPUT.xlsx")
        print('DATA.json schema:')
        print('  { "compile_date": "DD Month YYYY",')
        print('    "round_overrides": {},   # normally omitted -> CEILING to 1000')
        print('    "physical_stock": {"BCB|L": 10, ...},  # factory finished-goods')
        print('    "data": { "Kollam": {"BCB|L": 5, ...}, ... } }')
        sys.exit(1)
    with open(sys.argv[1]) as f:
        cfg = json.load(f)
    data = {wh: {tuple(k.split("|")): v for k, v in breakup.items()}
            for wh, breakup in cfg["data"].items()}
    physical_stock = {tuple(k.split("|")): v
                      for k, v in cfg.get("physical_stock", {}).items()}
    # clusters: pass explicit [[label,[wh,...]],...], OR set "cluster": true to
    # auto-group via DEPOT_CLUSTER. Omit / false -> flat layout.
    clusters = None
    if cfg.get("clusters"):
        clusters = [(lbl, list(whs)) for lbl, whs in cfg["clusters"]]
    elif cfg.get("cluster"):
        clusters = clusters_from_depots(data)
    build_workbook(
        data, sys.argv[2],
        compile_date=cfg.get("compile_date"),
        round_overrides=cfg.get("round_overrides"),
        physical_stock=physical_stock,
        clusters=clusters,
    )
    print(f"Saved: {sys.argv[2]}")
