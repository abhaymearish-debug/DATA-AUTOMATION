"""
KSBC cumulative-period ingest module.

When KSBC's portal exports a 1-16 or 17-end cumulative file, this module:
  1. Detects *CUMULATIVE.xlsx files in the KSBC shop sales folder
  2. Copies their SupplierWiseShopSaleReport sheet into the workbook as
     `<MONTH> 1-16 CUMULATIVE` or `<MONTH> 17-<lastday> CUMULATIVE`
  3. Returns the canonical day-range tuples for downstream aggregation

The existing aggregate_month() in ksbc_daily_update.py already supports the
cumulative-aware aggregation pattern (Opening from first sheet, Receipts/Sales
SUMMED across sheets, Closing OVERWRITTEN from last sheet). Feed it
[(1, 16, '<MONTH> 1-16 CUMULATIVE'), (17, N, '<MONTH> 17-<N> CUMULATIVE')]
and it will produce the correct period-totals-based COMBINED.

Spec: see CLAUDE.md "Cumulative-period rule" + .claude/memory/ksbc-cumulative-period.md

Usage as a module:
    from ksbc_ingest_cumulative import (
        find_cumulative_files,
        ingest_cumulative_file,
        canonical_cumulative_sheets,
    )

Usage as a CLI (diagnostic):
    python3 ksbc_ingest_cumulative.py <folder>
        - lists detected *CUMULATIVE.xlsx files with their parsed periods
"""

from __future__ import annotations
import os, re, sys
from pathlib import Path
from openpyxl import load_workbook

MONTHS = {"JANUARY":1,"FEBRUARY":2,"MARCH":3,"APRIL":4,"MAY":5,"JUNE":6,
          "JULY":7,"AUGUST":8,"SEPTEMBER":9,"OCTOBER":10,"NOVEMBER":11,"DECEMBER":12}

# Max valid day per month. February is lenient (29) to accept leap years; the
# workbook/sheet names don't carry the year so we can't pin 28 vs 29. Used to
# reject impossible cumulative windows (e.g. a 17-31 in 30-day April).
_MONTH_MAX = {1:31, 2:29, 3:31, 4:30, 5:31, 6:30,
              7:31, 8:31, 9:30, 10:31, 11:30, 12:31}


def _validate_window(s: int, e: int, month_u: str, where: str) -> None:
    """Raise ValueError if a cumulative window (s..e) is impossible for the
    month: s must be >=1, e>=s, and e<=the month's last day. (Month-length
    check added 1 Jun 2026 — a short/over-long 17-N window would otherwise
    under/over-report the full month silently.)"""
    last = _MONTH_MAX.get(MONTHS.get(month_u, 0), 31)
    if not (1 <= s <= e <= last):
        raise ValueError(
            f"Invalid cumulative window {s}-{e} for {month_u} ({where}): "
            f"days must satisfy 1 <= start <= end <= {last} (the last day of "
            f"{month_u}). Fix the cumulative file's day range and rerun."
        )


def parse_cumulative_filename(filename: str) -> tuple[int, int, str] | None:
    """Parse `<month> <M>-<N> CUMULATIVE.xlsx` (case-insensitive, lowercase
    month convention). Returns (start_day, end_day, MONTH_UPPER) or None if
    the filename doesn't match the cumulative convention.

    Per spec: a file is treated as cumulative ONLY if its basename matches
    `<month> <M>-<N>(th|st|nd|rd)? CUMULATIVE.xlsx`. The CUMULATIVE suffix is
    the only signal — daily/range drops without it are NOT cumulatives.
    """
    base = os.path.basename(filename)
    if 'CUMULATIVE' not in base.upper():
        return None
    # Strip extension and uppercase for matching
    stem = base.rsplit('.', 1)[0]
    # Pattern: <month> <M>-<N>(th|st|nd|rd)? CUMULATIVE
    # Allow an optional ordinal suffix on BOTH days (fix 1 Jun 2026): the old
    # pattern only tolerated it on the END day, so a plausibly-named file like
    # `may 1st-16th CUMULATIVE.xlsx` silently failed to be recognised as a
    # cumulative and got ignored entirely.
    m = re.match(
        r'^([A-Za-z]+)\s+(\d+)(?:st|nd|rd|th)?\s*-\s*(\d+)(?:st|nd|rd|th)?\s+CUMULATIVE$',
        stem,
        re.IGNORECASE,
    )
    if not m:
        return None
    month_part = m.group(1).upper()
    if month_part not in MONTHS:
        return None
    return int(m.group(2)), int(m.group(3)), month_part


def find_cumulative_files(folder: str) -> list[tuple[int, int, str, str]]:
    """Scan folder for cumulative-period raw files. Returns sorted list of
    (start_day, end_day, MONTH_UPPER, full_path)."""
    out = []
    for fn in os.listdir(folder):
        if not fn.lower().endswith('.xlsx'):
            continue
        if fn.startswith('~'):  # Excel lock files
            continue
        parsed = parse_cumulative_filename(fn)
        if parsed:
            s, e, m = parsed
            out.append((s, e, m, os.path.join(folder, fn)))
    out.sort(key=lambda x: (x[2], x[0]))
    return out


def cumulative_sheet_name(start_day: int, end_day: int, month_u: str) -> str:
    """Canonical CUMULATIVE sheet name inside the workbook."""
    return f'{month_u} {start_day}-{end_day} CUMULATIVE'


def ingest_cumulative_file(wb, raw_path: str, start_day: int, end_day: int,
                           month_u: str) -> tuple[str, int]:
    """Copy SupplierWiseShopSaleReport from raw_path into wb as a new
    `<MONTH> <M>-<N> CUMULATIVE` sheet. Replaces any existing sheet of
    that name. Returns (sheet_name, n_rows_copied)."""
    sn = cumulative_sheet_name(start_day, end_day, month_u)
    src = load_workbook(raw_path, read_only=True, data_only=False)
    src_ws = src['SupplierWiseShopSaleReport']
    if sn in wb.sheetnames:
        del wb[sn]
    new_ws = wb.create_sheet(sn)
    n = 0
    for row in src_ws.iter_rows(values_only=True):
        new_ws.append(list(row))
        n += 1
    src.close()
    return sn, n


def canonical_cumulative_sheets(wb, month_u: str) -> list[tuple[int, int, str]]:
    """Return [(start, end, sheet_name)] for all CUMULATIVE sheets present
    inside the workbook for the given month, sorted by start_day."""
    out = []
    pat = re.compile(rf'^{month_u}\s+(\d+)\s*-\s*(\d+)\s+CUMULATIVE$', re.IGNORECASE)
    for sn in wb.sheetnames:
        m = pat.match(sn)
        if m:
            out.append((int(m.group(1)), int(m.group(2)), sn))
    out.sort(key=lambda x: x[0])
    return out


def daily_sheets_outside_cumulatives(wb, month_u: str,
                                      cumulatives: list[tuple[int,int,str]]
                                      ) -> list[tuple[int, int, str]]:
    """Return [(start, end, sheet_name)] for daily/range sheets in the
    workbook that fall OUTSIDE every cumulative period. These are the
    fallback sources for any date range without a cumulative."""
    covered_days = set()
    for s, e, _ in cumulatives:
        covered_days.update(range(s, e + 1))
    out = []
    pat = re.compile(rf'^{month_u}\s+(\d+)(?:\s*-\s*(\d+))?$', re.IGNORECASE)
    for sn in wb.sheetnames:
        m = pat.match(sn)
        if m and 'CUMULATIVE' not in sn.upper():
            s = int(m.group(1))
            e = int(m.group(2)) if m.group(2) else s
            days = set(range(s, e + 1))
            covered = days & covered_days
            if not covered:
                # Fully OUTSIDE every cumulative window — use as a fallback.
                out.append((s, e, sn))
            elif covered == days:
                # Fully SHADOWED by a cumulative — the cumulative is source of
                # truth for these days, so skip the daily (audit trail only).
                continue
            else:
                # PARTIAL straddle (H1 fix 1 Jun 2026): some days are covered by
                # a cumulative and some aren't. A pre-aggregated range sheet
                # can't be sliced, so silently dropping it (the old behaviour)
                # lost the uncovered tail, and keeping it double-counts the
                # covered days. Surface it — never guess.
                uncovered = sorted(days - covered)
                raise ValueError(
                    f"Daily/range sheet '{sn}' (days {s}-{e}) partially overlaps "
                    f"a cumulative window: days {sorted(covered)} are covered by a "
                    f"cumulative but {uncovered} are not. A pre-aggregated range "
                    f"sheet cannot be split. Re-drop the uncovered tail "
                    f"({uncovered}) as its own daily/range file, or supply the "
                    f"matching cumulative, then rerun."
                )
    out.sort(key=lambda x: x[0])
    return out


def aggregation_sources(wb, month_u: str) -> tuple[list[tuple[int,int,str]], dict]:
    """Return the canonical list of (start, end, sheet_name) tuples to feed
    aggregate_month(), plus a status dict describing each period's source.

    Logic:
      Period 1 (1-16):   <MONTH> 1-16 CUMULATIVE  if present, else dailies 1..16 (provisional)
      Period 2 (17-end): <MONTH> 17-<N> CUMULATIVE if present, else dailies 17..N (provisional)
    """
    cumulatives = canonical_cumulative_sheets(wb, month_u)
    # Month-length / sanity validation of every cumulative window (Fix 3).
    for s, e, sn in cumulatives:
        _validate_window(s, e, month_u, sn)
    fallbacks = daily_sheets_outside_cumulatives(wb, month_u, cumulatives)

    # Merge — cumulatives take precedence by construction
    sources = sorted(cumulatives + fallbacks, key=lambda x: x[0])

    # Overlap guard (H2 fix 1 Jun 2026): no two aggregation sources may share a
    # day. Catches a range drop + single-day drop of the same day (e.g. a
    # `MAY 17-23` range AND a `MAY 17` single both present), which would
    # double-count Receipts/Sales for the shared days. Cumulative windows are
    # disjoint by construction, so this only fires on genuine duplicate dailies.
    for i in range(1, len(sources)):
        ps, pe, pn = sources[i - 1]
        cs_, ce_, cn = sources[i]
        if cs_ <= pe:
            raise ValueError(
                f"Overlapping aggregation sources: '{pn}' (days {ps}-{pe}) and "
                f"'{cn}' (days {cs_}-{ce_}) share days {cs_}-{min(pe, ce_)}. "
                f"Remove the duplicate/overlapping sheet before aggregating to "
                f"avoid double-counting."
            )

    # Classify by OVERLAP with the period windows, not exact start day (Fix 2).
    # The old `s == 1` / `s == 17` test mislabelled a corrected window like
    # `2-16` (day 1 is a Kerala dry day) as "provisional" even though the
    # cumulative WAS being used as the source. A cumulative covers a period if
    # its range intersects that period's day window.
    def _overlaps(s, e, a, b):
        return s <= b and e >= a
    period1_covered = any(_overlaps(s, e, 1, 16) for s, e, _ in cumulatives)
    period2_covered = any(_overlaps(s, e, 17, 31) for s, e, _ in cumulatives)

    status = {
        'period1_source': 'cumulative' if period1_covered else 'dailies (provisional)',
        'period2_source': 'cumulative' if period2_covered else 'dailies (provisional)',
        'cumulatives': cumulatives,
        'fallbacks': fallbacks,
        'sources': sources,
    }
    return sources, status


def main():
    if len(sys.argv) < 2:
        print('usage: ksbc_ingest_cumulative.py <folder>')
        sys.exit(2)
    folder = sys.argv[1]
    files = find_cumulative_files(folder)
    if not files:
        print(f'No *CUMULATIVE.xlsx files found in {folder}')
        return
    print(f'Detected {len(files)} cumulative file(s):')
    for s, e, m, p in files:
        print(f'  {m} days {s}-{e}: {p}')


if __name__ == '__main__':
    main()
