"""Promotion: the only code in this service that writes to a live workbook.

Kept in one small module on purpose. If you are auditing what can overwrite
months of accumulated data, this file is the whole surface area.

Rules enforced here:
  * A backup of the existing live workbook is taken first, every time.
  * Promotion only ever happens from a job in AWAITING_APPROVAL.
  * Raw files are moved into the job directory, not deleted. The cleanup policy
    in CLAUDE.md hard-deletes them; a web service with no one watching should
    keep them until someone says otherwise.
"""

from __future__ import annotations

import re
import os
import shutil
from datetime import datetime
from pathlib import Path

from . import config
from .jobs import Job, JobStatus, STORE
from .pipelines import STREAMS, MONTH_NAMES


class PromotionError(Exception):
    pass


_MIDMONTH_RE = re.compile(r"^(?P<month>[A-Z]+) 1st - (?P<day>\d+)(?:st|nd|rd|th) ", re.IGNORECASE)


def live_target(job: Job) -> Path:
    """Where this job's output belongs once approved."""
    stream = STREAMS[job.stream_key]

    if job.stream_key == "shop_sales":
        # The driver was handed the analysis workbook; it goes back to the same
        # path. Renaming a mid-month file to the clean full-month name is a
        # separate, deliberate act — not a side effect of a daily ingest.
        target = job.summary.get("analysis_workbook") or job.summary.get("workbook")
        if target:
            return Path(target)
        raise PromotionError(
            "Could not determine the target workbook for this build. "
            "Promote manually and report this — it should never happen."
        )

    if job.stream_key == "secondary_sales":
        window = _secondary_period(job, stream.input_path())
        if window:
            month_u, end_day = window
            return stream.input_path() / (
                f"{month_u} 1st - {_ordinal(end_day)} SECONDARY SALES ANALYSIS.xlsx")

        # Deliberately NOT "the newest workbook in the folder". That guess once
        # wrote a September build over AUGUST SECONDARY SALES ANALYSIS.xlsx,
        # because the raw had already been consumed and there was nothing left
        # in the folder to name the period from. A month's live workbook is
        # months of accumulated data; refusing costs one re-upload, guessing
        # wrong costs the month.
        raise PromotionError(
            "Could not tell which period this build covers, so nothing was written - "
            "overwriting another month's workbook is not a safe guess. Upload the raw "
            "again with the date it covers, and it will save itself.")

    if job.stream_key == "warehouse_stock":
        # Already written in place by the build script.
        return Path(job.output_path)

    raise PromotionError(f"Unknown stream {job.stream_key}.")


_RAW_SECONDARY_RE = re.compile(
    r"^RAW DATA -([A-Z]+) (\d{1,2})(?:ST|ND|RD|TH) SECONDARY SALES\.xlsx$", re.IGNORECASE)


def _secondary_period(job: Job, folder: Path) -> tuple[str, int] | None:
    """(MONTH, last day) for a secondary build, from the most reliable source up.

    Four answers, in order of how much they can be trusted: what the build
    itself reported, the dated raws still in the folder, the name of the file
    that was uploaded, and the date the upload was recorded against. The last
    two survive a folder the build has already emptied, which is exactly when
    this used to fall through to a guess.
    """
    month = (job.summary.get("month") or "").upper()
    day = job.summary.get("latest_day")
    if month in MONTH_NAMES and day:
        try:
            return month, int(day)
        except (TypeError, ValueError):
            pass

    window = _secondary_window(folder, job)
    if window:
        return window

    for name in job.uploaded_names or []:
        m = _RAW_SECONDARY_RE.match(name)
        if m and m.group(1).upper() in MONTH_NAMES:
            return m.group(1).upper(), int(m.group(2))

    if job.covers:
        try:
            d = datetime.strptime(job.covers[:10], "%Y-%m-%d").date()
            return MONTH_NAMES[d.month - 1], d.day
        except (ValueError, IndexError):
            pass
    return None


def _secondary_window(folder: Path, job: Job) -> tuple[str, int] | None:
    """(MONTH, last day) across the dated raws sitting in the secondary folder."""
    best: dict[str, int] = {}
    for p in folder.glob("RAW DATA -*SECONDARY SALES.xlsx"):
        m = _RAW_SECONDARY_RE.match(p.name)
        if not m:
            continue
        month_u = m.group(1).upper()
        if month_u in MONTH_NAMES:
            best[month_u] = max(best.get(month_u, 0), int(m.group(2)))
    if not best:
        return None

    # One upload, one month. If several months are somehow present, the month
    # this job covers decides, so a back-dated raw cannot rename the period.
    if job.covers and len(best) > 1:
        try:
            wanted = MONTH_NAMES[int(job.covers[5:7]) - 1]
        except (ValueError, IndexError):
            wanted = ""
        if wanted in best:
            return wanted, best[wanted]
    month_u = max(best, key=lambda k: best[k])
    return month_u, best[month_u]


def _ordinal(n: int) -> str:
    if 10 <= n % 100 <= 20:
        return f"{n}th"
    return f"{n}{ {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th') }"


def promote(job: Job) -> Path:
    if job.status is not JobStatus.AWAITING_APPROVAL:
        raise PromotionError(
            f"This build is '{job.status.value}', not awaiting approval. Nothing was written."
        )

    source = Path(job.output_path)
    if not source.is_file():
        raise PromotionError("The build output is missing. Re-run the build.")

    target = live_target(job)

    if target.resolve() == source.resolve():
        # warehouse_stock: the script already wrote in place.
        job.status = JobStatus.PROMOTED
        job.promoted_to = str(target)
        STORE.persist(job)
        return target

    target.parent.mkdir(parents=True, exist_ok=True)

    # An open Excel file leaves a lock alongside it and the overwrite will
    # either fail or be silently reverted when the user saves. Catch it here
    # with a message a person can act on.
    lock = target.parent / f"~${target.name}"
    if lock.exists():
        raise PromotionError(
            f"'{target.name}' is open in Excel on someone's machine. Close it and approve again."
        )

    if target.is_file():
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_dir = STORE.dir_for(job) / "replaced"
        backup_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(target, backup_dir / f"{target.stem}.bak_{stamp}{target.suffix}")

    # Copied to a sibling and moved into place, never written over the live
    # path directly. shutil.copyfile truncates the target and then streams into
    # it, so an interruption - a full disk, a killed worker - left a live
    # workbook that is half a file. build_secondary.py treats a workbook it
    # cannot open as "no prior month" and starts the month from zero, so a
    # half-written file does not fail loudly; it quietly costs a month of
    # accumulated days. os.replace is atomic within a filesystem, so the live
    # path is either the old workbook or the new one and never something in
    # between.
    staging = target.with_name(target.name + ".part")
    try:
        shutil.copyfile(source, staging)
        os.replace(staging, target)
    finally:
        staging.unlink(missing_ok=True)
    _copy_artifacts(job, target.parent)

    job.status = JobStatus.PROMOTED
    job.promoted_to = str(target)
    STORE.persist(job)
    return target


def _copy_artifacts(job: Job, destination: Path) -> None:
    """Place PDFs and other extras beside the workbook they describe."""
    for name in job.artifacts:
        src = STORE.dir_for(job) / name
        if src.is_file():
            try:
                shutil.copyfile(src, destination / name)
            except OSError:
                # An extra failing to copy must not undo a good promotion.
                pass


def discard(job: Job) -> None:
    if job.status is not JobStatus.AWAITING_APPROVAL:
        raise PromotionError(f"This build is '{job.status.value}' and cannot be discarded.")
    job.status = JobStatus.DISCARDED
    STORE.persist(job)
