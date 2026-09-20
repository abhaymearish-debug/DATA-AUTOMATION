"""Codified build pipelines.

WHY THIS FILE EXISTS
--------------------
Today the step order for each stream is not in code anywhere. It lives in the
prose of scheduled-task prompts and in CLAUDE.md, and it is executed by a human
or an assistant reading those instructions. That is the most fragile part of the
whole setup: the ordering rules are load-bearing and subtle, and nothing
enforces them.

Examples of ordering that is genuinely load-bearing:

  * ksbc_region_insights.py MUST be the final write of a KSBC build. It calls
    fix_region_sheet_styling.apply_all() in-session, because hero merged-cell
    styles only survive from the last save (openpyxl 3.1.5 drops style-only
    merged-range interiors on load+save).
  * ksbc_receipt_notes.py MUST be penultimate. openpyxl row moves do not carry
    comments, so any sort pass after it wipes the notes.
  * resize_comment_boxes.py MUST run after the last openpyxl save — any reload
    reverts every comment box to openpyxl's hardcoded 144x79 default.
  * ksbc_sort_by_rating.py runs BEFORE fix_rating_cell_format.py, or sorted
    rows keep the previous tier's font colour.
  * fix_nested_cf.py is DEPRECATED (10 May 2026) and must NOT run.

Encoding that here turns a prose convention into something deterministic,
testable, and reviewable in a diff.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Callable, Sequence

from . import config

MONTH_NAMES = [
    "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
    "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER",
]


def ordinal(n: int) -> str:
    """1 -> '1st', 2 -> '2nd', 11 -> '11th', 22 -> '22nd'."""
    if 10 <= n % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


class UploadRejected(Exception):
    """Raised when an uploaded file cannot be placed safely.

    Rejecting loudly is deliberate. The failure mode these pipelines punish is
    the silent one: a file that parses into the wrong period, or does not parse
    at all and leaves the workbook quietly short a day.
    """


# ---------------------------------------------------------------------------
# Steps
# ---------------------------------------------------------------------------


# Reports this app owns, as opposed to the pre-existing build scripts.
APP_REPORTS = Path(__file__).resolve().parent.parent / "reports"


@dataclass(frozen=True)
class Step:
    """One script invocation in a pipeline."""

    script: str
    # Given the job context, produce the argv tail after the script path.
    args: Callable[["JobContext"], Sequence[str]] = lambda ctx: ()
    # Extra environment for this step.
    env: Callable[["JobContext"], dict[str, str]] = lambda ctx: {}
    label: str = ""
    # A non-fatal step logs its failure and the pipeline continues. Used only
    # where the upstream script itself treats the call as non-fatal.
    fatal: bool = True
    # True when the script lives in this app's reports/ folder rather than in
    # the existing .claude/scripts pipeline.
    app_script: bool = False

    def display(self) -> str:
        return self.label or self.script


@dataclass
class JobContext:
    """Everything a step needs to know about the run in progress."""

    job_id: str
    stream_key: str
    uploaded: list[Path] = field(default_factory=list)
    # The workbook being built/mutated, once known.
    workbook: Path | None = None
    scratch_dir: Path | None = None
    meta: dict = field(default_factory=dict)
    # name -> mtime_ns of the stream folder's workbooks before the build, for
    # the streams that write in place and so cannot be asked afterwards which
    # file this run produced.
    before: dict = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Filename canonicalisation
# ---------------------------------------------------------------------------
#
# This is one of the places the web layer earns its keep. The scripts require
# exact filenames; a real person uploading from the KSBC portal will hand you
# "SupplierWiseShopSaleReport (3).xlsx" or "KSBC 17.09.2026.xlsx". Rather than
# letting parse_day_range raise deep inside a build, the upload is renamed to
# the canonical form up front, or rejected with a message that says what is
# needed.

_CUMULATIVE_RE = re.compile(
    r"(?P<month>[a-z]+)\s+(?P<start>\d{1,2})\s*-\s*(?P<end>\d{1,2})\s+CUMULATIVE",
    re.IGNORECASE,
)


def canonical_shop_sales_name(
    original: str, *, day: int | None, month: int, year: int,
    cumulative: tuple[int, int] | None = None,
) -> str:
    """Return the filename ksbc_daily_update.py's parser will accept.

    Daily:      'september 17th.xlsx'
    Cumulative: 'september 1-16 CUMULATIVE.xlsx'

    The cumulative suffix is what tells the pipeline the file is a period
    override rather than a daily drop, so it is never inferred — the caller
    states it explicitly.
    """
    month_lower = MONTH_NAMES[month - 1].lower()
    if cumulative:
        start, end = cumulative
        return f"{month_lower} {start}-{end} CUMULATIVE.xlsx"
    if day is None:
        raise UploadRejected("A daily shop-sales upload needs the day it covers.")
    return f"{month_lower} {ordinal(day)}.xlsx"


def canonical_secondary_name(*, day: int, month: int, year: int) -> str:
    """'RAW DATA -SEPTEMBER 17TH SECONDARY SALES.xlsx'

    build_secondary.py parses the export window END out of this filename, so a
    trailing dry day still extends the labelled period instead of silently
    shortening it.
    """
    return f"RAW DATA -{MONTH_NAMES[month - 1]} {day}TH SECONDARY SALES.xlsx"


# Chrome/Safari rename duplicate downloads, so the same portal export arrives
# as 'Report_17-Sep-2026.xls', 'Report 17-Sep-2026 (1).xls' and
# 'Report_17-Sep-2026__1_.xls'. build_warehouse_stock.py globs
# r'^Report[ _]' precisely because of this, so the gate here must not be
# stricter than the script it wraps.
_WAREHOUSE_RE = re.compile(r"^Report[ _]\d{1,2}-[A-Za-z]{3}-\d{4}", re.IGNORECASE)


def validate_warehouse_name(original: str) -> str:
    """Warehouse stock raws keep their Bevco export names.

    build_warehouse_stock.py globs 'Report_*.xls' and reads the report date out
    of the file itself, aborting on mixed report dates. Renaming them would
    destroy information, so these are validated rather than canonicalised.
    """
    stem = Path(original).name
    if not _WAREHOUSE_RE.match(stem):
        raise UploadRejected(
            f"'{stem}' does not look like a Bevco stock export. Expected files "
            "named like 'Report 17-Sep-2026.xls' or 'Report_17-Sep-2026 (1).xls' "
            "straight from the portal — the browser's '(1)', '(2)' copies are fine."
        )
    return stem


# ---------------------------------------------------------------------------
# Stream definitions
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Stream:
    key: str
    label: str
    blurb: str
    # Folder under CLAUDE_ROOT that the raws are dropped into.
    input_dir: str
    # Accepted upload extensions.
    extensions: tuple[str, ...]
    multi_file: bool
    steps: tuple[Step, ...]

    def input_path(self) -> Path:
        return config.CLAUDE_ROOT / self.input_dir


def _ksbc_driver_env(ctx: JobContext) -> dict[str, str]:
    """Point the KSBC driver's scratch output at this job's directory.

    ksbc_daily_update.py honours KSBC_SCRATCH_XLSX and KSBC_SUMMARY_JSON, which
    is what makes it safe to run under a web service: the build never writes to
    the live workbook, and the summary comes back machine-readable for the UI.
    """
    assert ctx.scratch_dir is not None
    return {
        "KSBC_SCRATCH_XLSX": str(ctx.scratch_dir / "out.xlsx"),
        "KSBC_SUMMARY_JSON": str(ctx.scratch_dir / "summary.json"),
    }


def _workbook_arg(ctx: JobContext) -> Sequence[str]:
    assert ctx.workbook is not None, "workbook not resolved yet"
    return (str(ctx.workbook),)


SHOP_SALES = Stream(
    key="shop_sales",
    label="Shop Sales",
    blurb="KSBC tertiary sales — daily, cumulative and analysis, across 15 bonds.",
    input_dir="KSBC shop sales",
    extensions=(".xlsx",),
    multi_file=True,
    steps=(
        # The driver does aggregation, the alias guard (apply/zero-drop/
        # double-count) and the nested Brand->Pack DETAIL rebuild internally.
        Step(
            "ksbc_daily_update.py",
            args=lambda ctx: (str(ctx.meta["analysis_workbook"]),) + tuple(str(p) for p in ctx.uploaded),
            env=_ksbc_driver_env,
            label="Ingest raw + aggregate (alias guard, nested DETAIL)",
        ),
        Step("ksbc_bond_performance_clusterwise.py", args=_workbook_arg,
             label="Rebuild BOND PERFORMANCE clusterwise"),
        # Sort BEFORE the rating-format pass, never after.
        Step("ksbc_sort_by_rating.py", args=_workbook_arg,
             label="Sort region / DETAIL / BOND PERFORMANCE by rating"),
        Step("fix_rating_cell_format.py", args=_workbook_arg,
             label="Re-apply rating tier fills and fonts"),
        Step("fix_detail_row_text.py", args=_workbook_arg,
             label="DETAIL sheet styling sweep"),
        # Penultimate: comments do not survive later row moves.
        Step("ksbc_receipt_notes.py", args=_workbook_arg,
             label="Bake receipt-date notes on DETAIL sheets"),
        # FINAL write. Calls fix_region_sheet_styling.apply_all() and
        # resize_comment_boxes() in-session; nothing may write after it.
        Step("ksbc_region_insights.py", args=_workbook_arg,
             label="BOND INSIGHTS + final styling (last write)"),
    ),
)


SECONDARY_SALES = Stream(
    key="secondary_sales",
    label="Secondary Sales",
    blurb="Warehouse to outlet — KSBC dispatch plus Consumer Fed and BAR invoice sales.",
    input_dir="Secondary sales",
    extensions=(".xlsx",),
    multi_file=False,
    steps=(
        # Self-contained: seeds from the live month workbook's COMBINED
        # DISPATCHES accumulator, layers the new day on top, runs the
        # raw-vs-retained reconciliation guard, then invokes
        # restyle_dashboard.py itself and verifies the locked dark dashboard
        # rendered before declaring success.
        Step("build_secondary.py",
             label="Build secondary analysis (seed + day merge + dashboard)"),
        # Brandwise cluster PDFs, built from the SCRATCH workbook this run just
        # produced — never from the live file — so the PDFs and the workbook
        # awaiting approval always describe the same data.
        Step("build_secondary_brandwise_pdfs.py",
             app_script=True,
             args=lambda ctx: (
                 "--base", str(config.CLAUDE_ROOT),
                 "--workbook", str(config.WORKSPACE_ROOT / "sec_scratch.xlsx"),
                 "--outdir", str(ctx.scratch_dir),
                 "--verify",
             ),
             label="Brandwise cluster PDFs (3 clusters)"),
    ),
)


WAREHOUSE_STOCK = Stream(
    key="warehouse_stock",
    label="Warehouse Stock",
    blurb="Bevco upstream stock position across ~28 warehouses. Snapshot, not sales.",
    input_dir="Warehouse stock",
    extensions=(".xls", ".xlsx"),
    multi_file=True,
    steps=(
        # Builds the dated dashboard workbook, appends stock_history.csv and
        # brand_pack_history.csv, and calls maintain_inbound_history.py as a
        # non-fatal hook.
        Step("build_warehouse_stock.py",
             label="Build stock dashboard + append history"),
    ),
)


# Shop sales arrives two ways and they are not the same job.
#
#   * A DAY's export is reported on its own. The daily report reads the day's
#     raw directly, so the 17th needs nothing but the 17th — no month, no
#     contiguity, no three-minute build. Hence: no steps. Storing the file IS
#     the work.
#   * A PERIOD export (1-16, or 17-month-end) is what builds the month's
#     analysis workbook, and that runs the full locked pipeline below.
SHOP_SALES_DAILY = Stream(
    key="shop_sales_daily",
    label="Shop Sales - Daily",
    blurb="One day's shop sales export, reported on its own.",
    input_dir="KSBC shop sales",
    extensions=(".xlsx",),
    multi_file=False,
    steps=(),
)

# The cumulative period export. Like the daily file, storing it IS the work:
# both PDFs it feeds - the per-bond cumulative books and the one-page
# comparative - are drawn on demand from the raw, so a period can be reopened,
# re-read and re-downloaded months later without rebuilding anything.
SHOP_SALES_CUMULATIVE = Stream(
    key="shop_sales_cumulative",
    label="Shop Sales - Cumulative",
    blurb="A period's cumulative shop sales export.",
    input_dir="KSBC shop sales/_cumulative",
    extensions=(".xlsx",),
    multi_file=False,
    steps=(),
)

# One month's purchase instruction is ~295 files, one per shop, each an HTML
# table KSBC saves with an .xls extension. Storing them IS the work: the report
# reads the raws, so a month answers the moment it is uploaded.
PURCHASE_INSTRUCTION = Stream(
    key="purchase_instruction",
    label="Purchase Instruction",
    blurb="KSBC's monthly buying instruction, one file per shop.",
    input_dir="PURCHASE INSTRUCTION/_months",
    extensions=(".xls", ".xlsx"),
    multi_file=True,
    steps=(),
)

ITEM_ISSUE = Stream(
    key="item_issue",
    label="Item Issue Consolidation",
    blurb="KSBC's issue consolidation, one file per warehouse, for one date range.",
    input_dir="ITEM ISSUE/_periods",
    extensions=(".xls", ".xlsx"),
    multi_file=True,
    steps=(),
)

STREAMS: dict[str, Stream] = {
    s.key: s for s in (SHOP_SALES, SHOP_SALES_DAILY, SHOP_SALES_CUMULATIVE,
                       SECONDARY_SALES, WAREHOUSE_STOCK, PURCHASE_INSTRUCTION,
                       ITEM_ISSUE)
}


# ---------------------------------------------------------------------------
# Preflight
# ---------------------------------------------------------------------------


def preflight() -> list[str]:
    """Verify every script a pipeline names actually exists.

    Cheap, and it catches the failure that would otherwise appear halfway
    through a build with a half-written workbook: a renamed or archived script.
    """
    problems: list[str] = []
    for stream in STREAMS.values():
        # An input folder is this app's to own, so create it rather than refuse
        # to start. A missing one used to hold every upload at 503 until someone
        # made the directory by hand — including on a fresh deploy, where none
        # of them exist yet. Only a folder we cannot create is a real problem.
        folder = stream.input_path()
        if not folder.exists():
            try:
                folder.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                problems.append(f"{stream.label}: cannot create input folder {folder} — {exc}")
        for step in stream.steps:
            base = APP_REPORTS if step.app_script else config.SCRIPTS_DIR
            script = base / step.script
            if not script.exists():
                problems.append(f"{stream.label}: script not found — {step.script} (looked in {base})")
    return problems


# Scripts that must never be run, with the reason, so a future edit that adds
# one back fails a test instead of silently corrupting formatting.
FORBIDDEN_SCRIPTS = {
    "fix_nested_cf.py": "Deprecated 10 May 2026 — superseded by fix_detail_row_text.py "
                        "and fix_region_sheet_styling.py.",
    "fix_detail_cf.py": "Legacy 4-tier CF; must not run on the nested workbook — "
                        "leaves no-activity rows blank.",
    "build_mobile_summary.py": "Mobile summary retired 25 Aug 2026; not part of any pipeline.",
}


def forbidden_in_pipelines() -> list[str]:
    """Assert no pipeline references a forbidden script."""
    hits = []
    for stream in STREAMS.values():
        for step in stream.steps:
            if step.script in FORBIDDEN_SCRIPTS:
                hits.append(f"{stream.label} references {step.script}: {FORBIDDEN_SCRIPTS[step.script]}")
    return hits
