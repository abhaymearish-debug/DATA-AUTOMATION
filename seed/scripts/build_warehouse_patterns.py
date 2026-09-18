#!/usr/bin/env python3
"""
build_warehouse_patterns.py — weekly digest of emergent stock-history patterns.

Reads:
    Warehouse stock/_history/stock_history.csv   (date, warehouse, physical, allotable, pending)
    Warehouse stock/_history/inbound_history.csv (date, warehouse, inbound_cases, dispatched_cases, phys_start, phys_end)

Writes:
    Warehouse stock/_history/PATTERN DIGEST.md

Sections (locked to match the 13-May-2026 digest format):
    1. Network position (daily totals + window deltas)
    2. Replenishment events (Δphys > +250 between consecutive snapshots per warehouse)
    3. Sustained drainers (no replenishment event in window, sorted by Δ% desc)
    4. Pending pressure — latest snapshot (pending% desc, min absolute pending=30)
    5. Builds vs drains across the window (first → latest, both directions)
    6. Open questions for production (data-driven, narrative)

Invocation:
    python3 .claude/scripts/build_warehouse_patterns.py
    python3 .claude/scripts/build_warehouse_patterns.py --root /Users/abhaymearish/Downloads/Claude
    python3 .claude/scripts/build_warehouse_patterns.py --out-extra /path/to/outputs/PATTERN_DIGEST.md

Idempotent. Overwrites the digest file in place.
"""

from __future__ import annotations

import argparse
import csv
import sys
from collections import defaultdict
from datetime import date as _date, datetime
from pathlib import Path
from typing import Dict, List, Tuple

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

DEFAULT_ROOT = Path("/Users/abhaymearish/Downloads/Claude")
HISTORY_SUBDIR = Path("Warehouse stock/_history")
STOCK_CSV = "stock_history.csv"
INBOUND_CSV = "inbound_history.csv"
DIGEST_FILE = "PATTERN DIGEST.md"

REFILL_THRESHOLD = 250          # Δphys (cases) qualifying as a "replenishment event"
PENDING_MIN_ABS = 30            # ignore tiny pending under this absolute case count
PENDING_TABLE_LIMIT = 12        # rows in §4
BUILDER_DRAINER_LIMIT = 8       # rows per side in §5
DRAINER_PCT_THRESHOLD = -20.0   # at least this drained to land in §3

# ---------------------------------------------------------------------------
# I/O
# ---------------------------------------------------------------------------


def parse_date(s: str) -> _date:
    return datetime.strptime(s.strip(), "%Y-%m-%d").date()


def load_stock(path: Path) -> List[dict]:
    rows: List[dict] = []
    with path.open(newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            rows.append(
                {
                    "date": parse_date(r["date"]),
                    "warehouse": r["warehouse"].strip(),
                    "physical": int(r["physical"] or 0),
                    "allotable": int(r["allotable"] or 0),
                    "pending": int(r["pending"] or 0),
                }
            )
    rows.sort(key=lambda x: (x["date"], x["warehouse"]))
    return rows


def load_inbound(path: Path) -> List[dict]:
    if not path.exists():
        return []
    rows: List[dict] = []
    with path.open(newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            rows.append(
                {
                    "date": parse_date(r["date"]),
                    "warehouse": r["warehouse"].strip(),
                    "inbound_cases": int(float(r.get("inbound_cases") or 0)),
                    "dispatched_cases": int(float(r.get("dispatched_cases") or 0)),
                    "phys_start": int(float(r.get("phys_start") or 0)),
                    "phys_end": int(float(r.get("phys_end") or 0)),
                }
            )
    rows.sort(key=lambda x: (x["date"], x["warehouse"]))
    return rows


# ---------------------------------------------------------------------------
# Analytics
# ---------------------------------------------------------------------------


def daily_network_totals(stock: List[dict]) -> List[Tuple[_date, int, int, int]]:
    by_date: Dict[_date, List[int]] = defaultdict(lambda: [0, 0, 0])
    for r in stock:
        agg = by_date[r["date"]]
        agg[0] += r["physical"]
        agg[1] += r["allotable"]
        agg[2] += r["pending"]
    return [(d, v[0], v[1], v[2]) for d, v in sorted(by_date.items())]


def per_warehouse_series(stock: List[dict]) -> Dict[str, List[Tuple[_date, int, int, int]]]:
    series: Dict[str, List[Tuple[_date, int, int, int]]] = defaultdict(list)
    for r in stock:
        series[r["warehouse"]].append(
            (r["date"], r["physical"], r["allotable"], r["pending"])
        )
    for wh in series:
        series[wh].sort(key=lambda x: x[0])
    return series


def find_refills(
    series: Dict[str, List[Tuple[_date, int, int, int]]],
    threshold: int = REFILL_THRESHOLD,
) -> List[Tuple[_date, str, int, int, int]]:
    """Return (date, warehouse, before_phys, after_phys, delta) for Δ > threshold."""
    events: List[Tuple[_date, str, int, int, int]] = []
    for wh, points in series.items():
        for prev, curr in zip(points, points[1:]):
            delta = curr[1] - prev[1]
            if delta > threshold:
                events.append((curr[0], wh, prev[1], curr[1], delta))
    events.sort(key=lambda x: (x[0], -x[4]))
    return events


def first_latest_per_warehouse(
    series: Dict[str, List[Tuple[_date, int, int, int]]]
) -> Dict[str, Tuple[int, int, float]]:
    """warehouse -> (first_phys, latest_phys, pct_change)."""
    out: Dict[str, Tuple[int, int, float]] = {}
    for wh, points in series.items():
        if not points:
            continue
        first = points[0][1]
        latest = points[-1][1]
        pct = (latest - first) / first * 100.0 if first else 0.0
        out[wh] = (first, latest, pct)
    return out


def latest_snapshot(stock: List[dict]) -> Tuple[_date, List[dict]]:
    latest = max(r["date"] for r in stock)
    return latest, [r for r in stock if r["date"] == latest]


# ---------------------------------------------------------------------------
# Markdown rendering
# ---------------------------------------------------------------------------


def fmt(n):
    if isinstance(n, float):
        return f"{n:,.1f}"
    return f"{n:,}"


def fmt_pct(p: float) -> str:
    sign = "+" if p >= 0 else ""
    return f"{sign}{p:.1f}%"


def render_digest(
    stock: List[dict],
    inbound: List[dict],
    refill_threshold: int = REFILL_THRESHOLD,
) -> str:
    today = max(r["date"] for r in stock)
    first_day = min(r["date"] for r in stock)
    distinct_dates = sorted({r["date"] for r in stock})
    warehouses = sorted({r["warehouse"] for r in stock})

    out: List[str] = []
    out.append(f"# Warehouse Pattern Digest — {today.isoformat()}")
    out.append("")
    out.append(
        f"_History window: **{first_day.isoformat()} → {today.isoformat()}** "
        f"({len(distinct_dates)} snapshot days, {len(warehouses)} Bevco warehouses). "
        f"Auto-generated by the `warehouse-patterns-weekly` task via "
        f"`.claude/scripts/build_warehouse_patterns.py`._"
    )
    out.append("")

    # --- s1 Network position --------------------------------------------
    out.append("## 1. Network position")
    out.append("")
    out.append("| Date | Physical | Allotable | Pending |")
    out.append("|---|---:|---:|---:|")
    totals = daily_network_totals(stock)
    for d, ph, al, pe in totals:
        out.append(f"| {d.isoformat()} | {fmt(ph)} | {fmt(al)} | {fmt(pe)} |")
    out.append("")
    first_ph = totals[0][1]
    last_ph = totals[-1][1]
    net_delta = last_ph - first_ph
    net_pct = net_delta / first_ph * 100.0 if first_ph else 0.0
    low_date, low_ph = min(((d, ph) for d, ph, _, _ in totals), key=lambda x: x[1])
    first_pe = totals[0][3]
    last_pe = totals[-1][3]
    out.append(
        f"Net change physical: **{fmt(first_ph)} → {fmt(last_ph)} "
        f"({'+' if net_delta >= 0 else ''}{fmt(net_delta)}, {fmt_pct(net_pct)})**. "
        f"Network bottomed at **{fmt(low_ph)} cases on {low_date.strftime('%d %b')}**. "
        f"Pending: **{fmt(first_pe)} → {fmt(last_pe)}** across the window."
    )
    out.append("")

    # --- s2 Replenishment events ----------------------------------------
    out.append(f"## 2. Replenishment events (Δphys > +{refill_threshold} in one snapshot)")
    out.append("")
    series = per_warehouse_series(stock)
    refills = find_refills(series, refill_threshold)
    if refills:
        out.append("| Day | Warehouse | Before → After | Δ |")
        out.append("|---|---|---:|---:|")
        for d, wh, before, after, delta in refills:
            out.append(
                f"| {d.isoformat()} | {wh} | {fmt(before)} → {fmt(after)} | **+{fmt(delta)}** |"
            )
    else:
        out.append("_No replenishment events crossed the threshold this window._")
    out.append("")

    # --- s3 Sustained drainers ------------------------------------------
    out.append("## 3. Sustained drainers — no replenishment in window")
    out.append("")
    refilled_warehouses = {wh for _, wh, _, _, _ in refills}
    fl = first_latest_per_warehouse(series)
    drainers = sorted(
        (
            (wh, first, latest, pct)
            for wh, (first, latest, pct) in fl.items()
            if wh not in refilled_warehouses and pct <= DRAINER_PCT_THRESHOLD
        ),
        key=lambda x: x[3],  # most-drained first (most-negative pct)
    )
    if drainers:
        out.append(
            f"| Warehouse | First ({first_day.strftime('%b %d')}) "
            f"| Latest ({today.strftime('%b %d')}) | Δ % | Latest physical |"
        )
        out.append("|---|---:|---:|---:|---:|")
        for wh, first, latest, pct in drainers:
            out.append(
                f"| {wh} | {fmt(first)} | {fmt(latest)} | **{fmt_pct(pct)}** | {fmt(latest)} |"
            )
    else:
        out.append("_No warehouse drained at least 20% without a recorded refill this window._")
    out.append("")
    if drainers:
        thin = [d for d in drainers if d[2] <= 200]
        if thin:
            names = ", ".join(f"**{wh}** ({fmt(latest)})" for wh, _, latest, _ in thin[:3])
            out.append(f"Critically low in absolute terms: {names}.")
            out.append("")

    # --- s4 Pending pressure --------------------------------------------
    out.append("## 4. Pending pressure — latest snapshot")
    out.append("")
    latest_date, latest_rows = latest_snapshot(stock)
    ranked = sorted(
        (
            (
                r["warehouse"],
                r["physical"],
                r["pending"],
                (r["pending"] / r["physical"] * 100.0) if r["physical"] else 0.0,
            )
            for r in latest_rows
            if r["pending"] >= PENDING_MIN_ABS
        ),
        key=lambda x: -x[3],
    )[:PENDING_TABLE_LIMIT]
    if ranked:
        out.append("| Warehouse | Physical | Pending | Pending % |")
        out.append("|---|---:|---:|---:|")
        for wh, ph, pe, pct in ranked:
            out.append(f"| {wh} | {fmt(ph)} | {fmt(pe)} | **{pct:.1f}%** |")
    else:
        out.append("_No pending pressure above threshold in latest snapshot._")
    out.append("")

    # --- s5 Builds vs drains across the window --------------------------
    out.append("## 5. Builds vs drains across the window")
    out.append("")
    builders = sorted(
        ((wh, first, latest, pct) for wh, (first, latest, pct) in fl.items() if pct > 0),
        key=lambda x: -x[3],
    )[:BUILDER_DRAINER_LIMIT]
    losers = sorted(
        ((wh, first, latest, pct) for wh, (first, latest, pct) in fl.items() if pct < 0),
        key=lambda x: x[3],
    )[:BUILDER_DRAINER_LIMIT]

    out.append("**Biggest builders (first → latest):**")
    out.append("")
    if builders:
        for wh, first, latest, pct in builders:
            out.append(f"- **{wh}** — {fmt(first)} → {fmt(latest)} ({fmt_pct(pct)})")
    else:
        out.append("_No warehouse net-built this window._")
    out.append("")
    out.append("**Biggest drainers (first → latest):**")
    out.append("")
    if losers:
        for wh, first, latest, pct in losers:
            out.append(f"- **{wh}** — {fmt(first)} → {fmt(latest)} ({fmt_pct(pct)})")
    else:
        out.append("_No warehouse net-drained this window._")
    out.append("")

    # --- s6 Open questions ----------------------------------------------
    out.append("## 6. Open questions for production")
    out.append("")
    bullets: List[str] = []
    if drainers and refills:
        skipped = ", ".join(wh for wh, _, _, _ in drainers[:7])
        recent_refills = ", ".join(sorted({wh for _, wh, _, _, _ in refills[-5:]}))
        bullets.append(
            f"The most recent refill wave hit **{recent_refills}** but bypassed "
            f"**{skipped}** — deliberate routing or queue?"
        )
    if ranked:
        top_wh, top_ph, top_pe, top_pct = ranked[0]
        bullets.append(
            f"**{top_wh}** sits at {top_pct:.1f}% pending ({fmt(top_pe)}/{fmt(top_ph)}) — "
            f"check whether indents are running ahead of stock or backlog isn't clearing."
        )
    if drainers:
        thin_critical = [d for d in drainers if d[2] <= 200]
        if thin_critical:
            names = ", ".join(wh for wh, _, _, _ in thin_critical[:5])
            bullets.append(
                f"Critically thin absolute stock at **{names}** — risk of stockout "
                f"if next dispatch slot is missed."
            )
    if not bullets:
        bullets.append("_Nothing flagged this week — network in steady state._")
    for b in bullets:
        out.append(f"- {b}")
    out.append("")

    # --- footer ---------------------------------------------------------
    out.append("---")
    out.append("")
    out.append(
        f"_Sources: `Warehouse stock/_history/stock_history.csv` "
        f"({len(distinct_dates)} days × {len(warehouses)} warehouses), "
        f"`Warehouse stock/_history/inbound_history.csv` ({len(inbound)} daily flow rows)._"
    )
    out.append("")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        type=Path,
        default=DEFAULT_ROOT,
        help="Project root containing Warehouse stock/ (default: %(default)s).",
    )
    parser.add_argument(
        "--out-extra",
        type=Path,
        default=None,
        help="Optional second output path (e.g. an outputs folder copy).",
    )
    parser.add_argument(
        "--refill-threshold",
        type=int,
        default=REFILL_THRESHOLD,
        help="Δphys cases that constitute a replenishment event (default: %(default)s).",
    )
    parser.add_argument(
        "--print",
        action="store_true",
        help="Print the digest to stdout in addition to writing the file.",
    )
    args = parser.parse_args(argv)

    history_dir = args.root / HISTORY_SUBDIR
    stock_path = history_dir / STOCK_CSV
    inbound_path = history_dir / INBOUND_CSV
    digest_path = history_dir / DIGEST_FILE

    if not stock_path.exists():
        print(f"ERROR: missing {stock_path}", file=sys.stderr)
        return 2
    stock = load_stock(stock_path)
    if not stock:
        print(f"ERROR: {stock_path} has no rows", file=sys.stderr)
        return 2
    inbound = load_inbound(inbound_path)

    md = render_digest(stock, inbound, refill_threshold=args.refill_threshold)
    digest_path.write_text(md, encoding="utf-8")
    print(f"Wrote {digest_path} ({len(md):,} bytes)")
    if args.out_extra:
        args.out_extra.parent.mkdir(parents=True, exist_ok=True)
        args.out_extra.write_text(md, encoding="utf-8")
        print(f"Wrote {args.out_extra}")
    if args.print:
        print()
        print(md)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
