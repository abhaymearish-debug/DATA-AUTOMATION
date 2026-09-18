#!/usr/bin/env python3
"""
Insert MISSING BRAND rows (plus their pack rows) into DETAIL shop blocks.

WHY THIS EXISTS
---------------
`ksbc_detail_nested.py` is a pure VALUE-OVERWRITE pass: it walks the rows that
already exist in a DETAIL block and fills them from COMBINED. It can never add
a row. `fix_detail_row_text.py::_insert_missing_pack_rows` auto-inserts a
missing (brand, pack) combo — but ONLY when the brand row already exists; its
own source says:

    if brand not in brand_row_for_brand:
        # The brand row itself is missing from the block — different
        # problem; brand-row insertion isn't supported here.
        continue

So when a shop starts stocking a BRAND it has never carried before, the whole
brand (and all its packs) is silently absent from the DETAIL block. The region
sheet, BOND PERFORMANCE and DASHBOARD are all correct because they are built
from COMBINED — only the DETAIL block under-reports, which makes the block
TOTAL disagree with the region row for that shop.

Found 17 Aug 2026 on shop 107013 GANDHINAGAR (TRIPUNITHURA): COMBINED carried
K.S 99 LIFE TIME MATURED XXX RUM · 375 ML (O=8.79, S=0.29, C=8.50) but the
DETAIL block had no K.S 99 brand at all, so the block TOTAL read
O=15.69 / S=7.11 / C=18.58 against the region row's 24.48 / 7.40 / 27.08.

WHERE IT RUNS
-------------
Immediately BEFORE `ksbc_detail_nested.py` in the KSBC pipeline (it only
creates correctly-labelled rows; the nested pass fills the values, the sort
pass orders them, `fix_detail_row_text.py` styles them, and
`ksbc_region_insights.py` remains the final writer).

SAFETY
------
Mirrors the proven merged-cell sequence used by `fix_detail_row_text.py`:
unmerge every block-header A:G range sheet-wide FIRST, do all inserts
bottom-up (so pending insert positions never shift), then re-merge from the
freshly-recomputed block list. Formula row-references are NOT hand-shifted
here — the inserted rows carry self-referencing F/G formulas and
`fix_detail_row_text.py::_normalize_rating_formulas` re-points every F/G
formula to its own row on the very next pass.

Idempotent: a brand already present in a block is never re-inserted.

Usage:
    python3 ksbc_insert_missing_brand_rows.py <workbook.xlsx> [--dry-run]
"""

from __future__ import annotations
import sys
from pathlib import Path
from openpyxl import load_workbook

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

import fix_detail_row_text as fdrt


def _block_brands(ws, header_row: int, total_row: int) -> dict:
    """Return {normalised_brand: last_row_of_that_brand} for a shop block."""
    out = {}
    current = None
    for r in range(header_row + 2, total_row):
        label = ws.cell(r, 1).value
        if label is None or (isinstance(label, str) and not label.strip()):
            continue
        if fdrt._is_pack_label(label):
            if current is not None:
                out[current] = r
        else:
            current = fdrt._norm_brand(label)
            out[current] = r
    return out


def insert_missing_brand_rows(wb, dry_run: bool = False) -> list:
    """Insert brand (+pack) rows present in COMBINED but absent from a DETAIL
    block. Returns a list of (sheet, shop_code, brand, [packs])."""
    sbp, sb = fdrt._build_combined_aggregates(wb)
    if not sbp:
        raise SystemExit('ERROR: no COMBINED sheet found — refusing to guess.')

    # Display spelling for each normalised brand/pack, taken from COMBINED.
    combined_name = next(n for n in wb.sheetnames if 'COMBINED' in n.upper())
    cs = wb[combined_name]
    brand_disp, pack_disp = {}, {}
    for r in range(2, cs.max_row + 1):
        b, p = cs.cell(r, 5).value, cs.cell(r, 6).value
        if b is not None:
            brand_disp.setdefault(fdrt._norm_brand(b), str(b).strip())
        if p is not None:
            pack_disp.setdefault(fdrt._norm_pack(p), str(p).strip())

    # shop -> {brand: [packs]}
    by_shop = {}
    for (code, nb, np_) in sbp:
        by_shop.setdefault(code, {}).setdefault(nb, []).append(np_)

    applied = []
    for sn in [n for n in wb.sheetnames if n.endswith('DETAIL')]:
        ws = wb[sn]
        plans = []  # (insert_at, code, brand, [packs])
        for header_row, total_row in fdrt._find_blocks(ws):
            code = fdrt._parse_shop_code(ws.cell(header_row, 1).value) \
                if hasattr(fdrt, '_parse_shop_code') else None
            if code is None:
                import re as _re
                m = _re.match(r'^\s*(\d{5,6})\b', str(ws.cell(header_row, 1).value or ''))
                if not m:
                    continue
                code = int(m.group(1))
            code = fdrt._canon_code(code)
            present = _block_brands(ws, header_row, total_row)
            for nb, packs in by_shop.get(code, {}).items():
                if nb in present:
                    continue
                # Insert the brand and every pack COMBINED knows about,
                # directly ABOVE the block's TOTAL row.
                plans.append((total_row, code, nb, sorted(set(packs))))

        if not plans:
            continue

        applied.extend((sn, c, brand_disp.get(b, b),
                        [pack_disp.get(p, p) for p in pk]) for (_, c, b, pk) in plans)
        if dry_run:
            continue

        # --- unmerge every block-header A:G range sheet-wide, then insert
        #     bottom-up so pending positions never shift. ---
        for rng in [str(m) for m in ws.merged_cells.ranges]:
            ws.unmerge_cells(rng)

        for insert_at, code, nb, packs in sorted(plans, key=lambda x: -x[0]):
            n_new = 1 + len(packs)
            ws.insert_rows(insert_at, n_new)
            ws.cell(insert_at, 1).value = brand_disp.get(nb, nb)
            for k, p in enumerate(packs, start=1):
                ws.cell(insert_at + k, 1).value = pack_disp.get(p, p)
            for r in range(insert_at, insert_at + n_new):
                for c in range(2, 6):
                    ws.cell(r, c).value = 0
                ws.cell(r, 6).value = fdrt._detail_f_formula(r)
                ws.cell(r, 7).value = fdrt._detail_g_formula(r)

        for header_row, _total_row in fdrt._find_blocks(ws):
            ws.merge_cells(f'A{header_row}:G{header_row}')

    return applied


def main():
    if len(sys.argv) < 2:
        print('usage: ksbc_insert_missing_brand_rows.py <workbook.xlsx> [--dry-run]')
        sys.exit(2)
    path = sys.argv[1]
    dry = '--dry-run' in sys.argv
    wb = load_workbook(path)
    applied = insert_missing_brand_rows(wb, dry_run=dry)
    if not applied:
        print('MISSING-BRAND CHECK OK · no DETAIL block is missing a brand.')
        return
    for sn, code, brand, packs in applied:
        print(f'  {"would insert" if dry else "inserted"}: {sn} · shop {code} · '
              f'{brand} · packs {", ".join(packs)}')
    if not dry:
        wb.save(path)
        print(f'MISSING-BRAND REPAIR · {len(applied)} brand row(s) inserted · saved {path}')


if __name__ == '__main__':
    main()
