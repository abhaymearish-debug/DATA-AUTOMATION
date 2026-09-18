"""
KSBC DETAIL-sheet value updater for the NESTED Brand → Pack layout.

Each DETAIL sheet contains N shop blocks. Block shape:

    [start]     "<code> — <shop name>  |  Staff: <staff>"   (merged A:G header)
    [start+1]   "Brand / Pack" | Opening | Receipts | Sales | Closing | ST% | Rating
    [start+2]   brand summary row (e.g. "OLD PEARL NO.1 MATURED XXX RUM")
    [start+3..] indented pack rows for that brand (e.g. "500 ML", "1000 ML")
                ... next brand summary row ... its packs ...
    [last]      TOTAL

Pack rows are identified by label matching r"^\s*\d+\s*ML\s*$".
Brand rows are the remaining string labels in col A (within the block).

update_detail_sheet(ws, shop_brand, shop_brand_pack, canonical_code_fn)
overwrites B..E for every brand/pack row + TOTAL row (TOTAL is sum of brand
rows — NOT brand+pack, which would double-count).
"""
from __future__ import annotations
import re

PACK_RE = re.compile(r"^\s*\d+\s*ML\s*$", re.IGNORECASE)


def _norm_brand(v):
    """Normalise a brand string for matching: upper-case, drop apostrophes
    (curly/straight/back-tick) and collapse internal whitespace.

    The raw KSBC export and the DETAIL row templates disagree on punctuation
    for some brands (e.g. raw 'MORNING WALKERS XO BRANDY' vs template
    "MORNING WALKER'S XO BRANDY"). Exact-string matching silently dropped
    those rows to 0 (MAY 2026: 318.58 cs of Morning Walker's vanished from
    every DETAIL sheet). Normalising both sides of the lookup fixes it.
    """
    if v is None:
        return ''
    return ' '.join(str(v).strip().upper()
                    .replace("’", "").replace("'", "").replace("`", "").split())


def _is_pack_label(v):
    return isinstance(v, str) and bool(PACK_RE.match(v))


def _find_shop_blocks(ws):
    """Return list of (header_row, total_row) for each shop block."""
    headers = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and " — " in v and "Staff:" in v:
            headers.append(r)
    blocks = []
    for i, h in enumerate(headers):
        end_limit = headers[i + 1] - 1 if i + 1 < len(headers) else ws.max_row
        total_row = None
        for r in range(h + 2, end_limit + 1):
            if ws.cell(r, 1).value == "TOTAL":
                total_row = r
                break
        if total_row is not None:
            blocks.append((h, total_row))
    return blocks


def _parse_shop_code(header_val: str):
    m = re.match(r"\s*(\d+)\s*—", header_val)
    return int(m.group(1)) if m else None


def update_detail_sheet(ws, shop_brand, shop_brand_pack, canonical_code_fn):
    # Index the data by NORMALISED brand so apostrophe/punctuation drift between
    # the raw export spelling and the DETAIL row label can't drop a brand to 0.
    nb = {}
    for (c, b), v in shop_brand.items():
        nb[(c, _norm_brand(b))] = v
    nbp = {}
    for (c, b, pk), v in shop_brand_pack.items():
        nbp[(c, _norm_brand(b), pk)] = v

    blocks = _find_shop_blocks(ws)
    n = 0
    for start, total_row in blocks:
        code = _parse_shop_code(ws.cell(start, 1).value)
        if code is None:
            continue
        canon = canonical_code_fn(code)

        current_brand = None
        for r in range(start + 2, total_row):
            label = ws.cell(r, 1).value
            if label is None or (isinstance(label, str) and label.strip() == ""):
                continue
            if _is_pack_label(label):
                if current_brand is None:
                    continue
                pack = label.strip()
                t = nbp.get((canon, _norm_brand(current_brand), pack),
                            {"Opening": 0, "Receipts": 0, "Sales": 0, "Closing": 0})
            else:
                current_brand = label
                t = nb.get((canon, _norm_brand(current_brand)),
                           {"Opening": 0, "Receipts": 0, "Sales": 0, "Closing": 0})
            ws.cell(r, 2).value = t["Opening"]
            ws.cell(r, 3).value = t["Receipts"]
            ws.cell(r, 4).value = t["Sales"]
            ws.cell(r, 5).value = t["Closing"]

        # TOTAL row: sum of BRAND rows only
        op = rc = sl = cl = 0.0
        for r in range(start + 2, total_row):
            label = ws.cell(r, 1).value
            if label is None:
                continue
            if not _is_pack_label(label):
                op += ws.cell(r, 2).value or 0
                rc += ws.cell(r, 3).value or 0
                sl += ws.cell(r, 4).value or 0
                cl += ws.cell(r, 5).value or 0
        ws.cell(total_row, 2).value = op
        ws.cell(total_row, 3).value = rc
        ws.cell(total_row, 4).value = sl
        ws.cell(total_row, 5).value = cl
        n += 1
    return n
