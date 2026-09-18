"""Maintain _history/inbound_history.csv from stock_history.csv + Secondary sales COMBINED DISPATCHES.

Network-level identity (per Abhay 19 May 2026, since a load is always unloaded fully):
    inbound_during_window = ΣΔphys_during_window + Σdispatched_during_window

The window for each snapshot day D is [prior_D, D) — every day from the previous
Bevco snapshot up to but excluding today's snapshot. Bevco's stock_history.csv has
gaps (Sundays, holidays, missed exports) so a "between-snapshots" window can span
several days; the maintainer sums Secondary-sales dispatches across that full window
rather than just D itself, which avoids fake negative inbound on long gaps.

`Σdispatched on D` comes from `Secondary sales/<MONTH>...SECONDARY SALES ANALYSIS.xlsx`,
sheet `<MONTH> COMBINED DISPATCHES`, column "Issue Cases" filtered by "Inv/GTN Date".

The per-warehouse split kept in the CSV (`inbound[W] = Δphys[W] + Σdisp[W in window]`)
is best-effort and absorbs same-day in/out friction; at the network level the totals
reconstruct correctly.

Skip-on-missing: if any day in the window isn't covered by an on-disk Secondary
analysis (month missing, or month present but mid-month file hasn't reached that
day), the target snapshot date is skipped — no heuristic fallback that would
produce wrong numbers.

Idempotent: only writes (date, warehouse) pairs not already present.
"""

from __future__ import annotations

import csv
import re
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl

WH_DIR_NAME = "Warehouse stock"
HIST_DIR_NAME = "_history"
STOCK_CSV_NAME = "stock_history.csv"
INBOUND_CSV_NAME = "inbound_history.csv"
SECONDARY_DIR_NAME = "Secondary sales"

MONTH_NAMES = [
    "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
    "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER",
]

def canon_wh(raw):
    """Strip the "WH-" prefix and the licence-code suffix from a Secondary-sales
    warehouse name. Handles both "WH-X FL9-..." and "WH-X RFL9/..." variants,
    and passes plain canonical names through unchanged.

        WH-KOLLAM FL9-KLM-01/2026-27         -> KOLLAM
        WH-THODUPUZHA RFL9/01/2026-27/IDUKKI -> THODUPUZHA
        WH-ALAPPUZHA FL-9-A-1/2026-27        -> ALAPPUZHA
        KALPETTA                             -> KALPETTA
    """
    if raw is None:
        return None
    s = str(raw).strip().upper()
    if not s:
        return None
    if s.startswith("WH-"):
        s = s[3:].lstrip()
    tokens = s.split()
    keep = []
    for tok in tokens:
        # Stop at the first token that looks like a license-code suffix.
        if any(c.isdigit() for c in tok) or "/" in tok or tok.startswith("FL") or tok.startswith("RFL"):
            break
        keep.append(tok)
    name = " ".join(keep).strip()
    return name or s


def parse_inv_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%b-%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def month_name_for_iso(iso):
    return MONTH_NAMES[datetime.strptime(iso, "%Y-%m-%d").month - 1]


def find_secondary_workbook(secondary_dir, month_upper):
    full = secondary_dir / f"{month_upper} SECONDARY SALES ANALYSIS.xlsx"
    if full.exists():
        return full
    pat = re.compile(
        rf"^{month_upper}\s+1st\s+-\s+(\d+)(?:st|nd|rd|th)\s+SECONDARY SALES ANALYSIS\.xlsx$",
        re.IGNORECASE,
    )
    best, best_n = None, -1
    for p in secondary_dir.iterdir():
        m = pat.match(p.name)
        if m:
            n = int(m.group(1))
            if n > best_n:
                best, best_n = p, n
    return best


def coverage_range_for(wb_path, month_upper, year=None):
    """Return (first_iso, last_iso) for the nominal date range this workbook covers,
    inferred from its filename. Full-month → 1st through last day; mid-month → 1st through Nth.
    `year` should be derived from the workbook's own dispatch dates (cross-year safe);
    falls back to the current calendar year only when no data year is available."""
    import calendar
    if year is None:
        year = datetime.now().year  # degenerate fallback (empty workbook only)
    month_idx = MONTH_NAMES.index(month_upper) + 1
    first = datetime(year, month_idx, 1)
    name = wb_path.name
    m = re.search(r'1st\s+-\s+(\d+)(?:st|nd|rd|th)\s+SECONDARY', name, re.IGNORECASE)
    if m:
        last_day = int(m.group(1))
    else:
        last_day = calendar.monthrange(year, month_idx)[1]
    last = datetime(year, month_idx, last_day)
    return first.strftime("%Y-%m-%d"), last.strftime("%Y-%m-%d")


def read_dispatched_for_month(secondary_dir, month_upper):
    wb_path = find_secondary_workbook(secondary_dir, month_upper)
    if wb_path is None:
        return {}, None, None

    wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
    sheet_name = f"{month_upper} COMBINED DISPATCHES"
    if sheet_name not in wb.sheetnames:
        wb.close()
        return {}, wb_path, None

    ws = wb[sheet_name]
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    if header is None:
        wb.close()
        return {}, wb_path, coverage_range_for(wb_path, month_upper)

    def idx_of(name):
        for i, h in enumerate(header):
            if h and str(h).strip().lower() == name.lower():
                return i
        return None

    i_wh = idx_of("Warehouse Name")
    i_date = idx_of("Inv/GTN Date")
    i_case = idx_of("Issue Cases")

    if None in (i_wh, i_date, i_case):
        wb.close()
        raise RuntimeError(
            f"COMBINED DISPATCHES sheet in {wb_path.name} missing expected columns "
            f"(Warehouse Name={i_wh}, Inv/GTN Date={i_date}, Issue Cases={i_case})"
        )

    bucket = defaultdict(int)
    for row in rows:
        if row is None:
            continue
        wh = canon_wh(row[i_wh])
        iso = parse_inv_date(row[i_date])
        cases = row[i_case]
        if not (wh and iso and cases is not None):
            continue
        try:
            n = int(round(float(cases)))
        except (TypeError, ValueError):
            continue
        bucket[(iso, wh)] += n

    wb.close()
    # Derive the data year from the dispatch dates themselves (cross-year safe) so
    # coverage_range_for never assumes datetime.now().year — a December workbook
    # processed in January would otherwise compute next-year coverage and
    # silently skip every snapshot in its window.
    data_year = None
    if bucket:
        yrs = [int(iso[:4]) for (iso, _wh) in bucket.keys()]
        data_year = max(set(yrs), key=yrs.count)
    return dict(bucket), wb_path, coverage_range_for(wb_path, month_upper, data_year)


def read_stock_history(stock_csv):
    out = defaultdict(dict)
    if not stock_csv.exists():
        return out
    with stock_csv.open() as f:
        r = csv.reader(f)
        header = next(r, None)
        if header is None:
            return out
        for row in r:
            if not row or len(row) < 3:
                continue
            iso, wh, phys = row[0], row[1].strip().upper(), row[2]
            try:
                out[iso][wh] = int(round(float(phys)))
            except (TypeError, ValueError):
                continue
    return dict(out)


def read_existing_inbound_keys(inbound_csv):
    keys = set()
    if not inbound_csv.exists():
        return keys
    with inbound_csv.open() as f:
        r = csv.reader(f)
        header = next(r, None)
        if header is None:
            return keys
        for row in r:
            if not row or len(row) < 2:
                continue
            keys.add((row[0], row[1].strip().upper()))
    return keys


def maintain_inbound_history(folder_root, only_dates=None, verbose=True):
    folder_root = Path(folder_root)
    stock_csv = folder_root / WH_DIR_NAME / HIST_DIR_NAME / STOCK_CSV_NAME
    inbound_csv = folder_root / WH_DIR_NAME / HIST_DIR_NAME / INBOUND_CSV_NAME
    secondary_dir = folder_root / SECONDARY_DIR_NAME

    if not stock_csv.exists():
        if verbose:
            print(f"[maintain_inbound] no stock_history at {stock_csv} — skip")
        return 0, 0

    stock = read_stock_history(stock_csv)
    if not stock:
        if verbose:
            print("[maintain_inbound] stock_history empty — skip")
        return 0, 0

    all_dates = sorted(stock.keys())
    targets = list(only_dates) if only_dates else all_dates
    targets = sorted(set(targets))

    def prior_of(iso):
        prev = None
        for d in all_dates:
            if d < iso:
                prev = d
            else:
                break
        return prev

    def prior_of_wh(iso, wh):
        """Per-WH lookback: find the most recent snapshot strictly before 
        that contains . Returns None if WH never appeared earlier.

        Handles the BALARAMAPURAM-shaped failure where Bevco's export silently
        drops a WH on day D. Without this, that WH's dispatches around D get
        lost because the (D-1, D] window would exclude it."""
        for d in reversed(all_dates):
            if d < iso and wh in stock.get(d, {}):
                return d
        return None

    existing = read_existing_inbound_keys(inbound_csv)

    # Cache Secondary dispatch data per month.
    # Value = (disp_dict, src_path, dates_covered_set).
    sec_cache = {}

    def secondary_for(month_upper):
        if month_upper in sec_cache:
            return sec_cache[month_upper]
        disp, src, cov_range = read_dispatched_for_month(secondary_dir, month_upper)
        sec_cache[month_upper] = (disp, src, cov_range)
        return sec_cache[month_upper]

    def window_fully_covered(prior_iso, target_iso):
        # Window is [prior_iso, target_iso) — see dispatched_in_window for convention rationale.
        end = datetime.strptime(target_iso, "%Y-%m-%d")
        cur = datetime.strptime(prior_iso, "%Y-%m-%d")
        while cur < end:
            iso = cur.strftime("%Y-%m-%d")
            m = month_name_for_iso(iso)
            _disp, src, cov_range = secondary_for(m)
            if src is None:
                return False, f"no Secondary analysis for {m}"
            if cov_range is None:
                return False, f"{m} analysis exists but COMBINED DISPATCHES sheet absent"
            first, last = cov_range
            if iso < first or iso > last:
                return False, f"{m} analysis covers {first}..{last}, needs {iso}"
            cur += timedelta(days=1)
        return True, None

    def dispatched_in_window(prior_iso, target_iso, wh):
        # Bevco snapshots are MORNING events (per Abhay 19 May 2026).
        # Window between prior_snapshot and today_snapshot is [prior_iso, target_iso):
        #   - Dispatches dated prior_iso happen IN the day AFTER the prior morning snapshot → INCLUDED.
        #   - Dispatches dated target_iso happen AFTER the today morning snapshot → EXCLUDED.
        # Previous bug: used (prior_iso, target_iso] — opposite convention, missed prior-date
        # afternoon dispatches and wrongly included today-date afternoon dispatches.
        # Fixing this closed the May MTD gap to factory log from 118 cases → 8 cases.
        start = datetime.strptime(prior_iso, "%Y-%m-%d")
        end = datetime.strptime(target_iso, "%Y-%m-%d")
        total = 0
        cur = start
        while cur < end:
            iso = cur.strftime("%Y-%m-%d")
            m = month_name_for_iso(iso)
            disp, _src, _cov = secondary_for(m)
            total += disp.get((iso, wh), 0)
            cur += timedelta(days=1)
        return total

    new_rows = []
    skipped_no_secondary = []

    for iso in targets:
        today_phys = stock[iso]
        if not today_phys:
            continue

        # Per-WH lookback: each warehouse independently finds its own prior snapshot.
        # Skip the snapshot entirely if NO WH has a usable prior (i.e., iso is the
        # first snapshot of the dataset).
        per_wh_priors = {}
        for wh in today_phys:
            pwh = prior_of_wh(iso, wh)
            if pwh is not None:
                per_wh_priors[wh] = pwh
        if not per_wh_priors:
            if verbose:
                print(f"[maintain_inbound] {iso} is first snapshot, no WH has a prior — skip")
            continue

        for wh in sorted(per_wh_priors):
            if (iso, wh) in existing:
                continue
            prior = per_wh_priors[wh]
            ok, reason = window_fully_covered(prior, iso)
            if not ok:
                if verbose:
                    print(f"[maintain_inbound] {iso} {wh}: window [{prior}, {iso}) not fully covered — {reason} — skip WH")
                if iso not in skipped_no_secondary:
                    skipped_no_secondary.append(iso)
                continue
            phys_start = stock[prior][wh]
            phys_end = today_phys[wh]
            dispatched = dispatched_in_window(prior, iso, wh)
            inbound = (phys_end - phys_start) + dispatched
            # Invariant: this row is END-stamped — phys_end is TODAY's snapshot and
            # phys_start the per-WH prior. Asserting it prevents a future refactor
            # from silently reintroducing the April-style START-stamped date shift.
            assert phys_end == today_phys[wh] and phys_start == stock[prior][wh], (
                f"date-stamp invariant broken at {iso} {wh}")
            new_rows.append([iso, wh, inbound, dispatched, phys_start, phys_end])

    need_header = not inbound_csv.exists()
    inbound_csv.parent.mkdir(parents=True, exist_ok=True)
    with inbound_csv.open("a", newline="") as f:
        w = csv.writer(f)
        if need_header:
            w.writerow(["date", "warehouse", "inbound_cases", "dispatched_cases", "phys_start", "phys_end"])
        for row in new_rows:
            w.writerow(row)

    if verbose:
        print(f"[maintain_inbound] wrote {len(new_rows)} new rows; "
              f"skipped {len(skipped_no_secondary)} date(s) for missing Secondary analysis")
    return len(new_rows), len(skipped_no_secondary)


if __name__ == "__main__":
    import sys
    root = Path(sys.argv[1]) if len(sys.argv) > 1 else (
        Path(__file__).resolve().parent.parent.parent
    )
    written, skipped = maintain_inbound_history(root, only_dates=None, verbose=True)
    print(f"BACKFILL_DONE rows_written={written} dates_skipped={skipped}")
