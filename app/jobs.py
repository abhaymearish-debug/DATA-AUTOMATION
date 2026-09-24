"""Job runner: executes a pipeline as a sequence of subprocesses.

BUILD TO SCRATCH, THEN SAVE
---------------------------
Every data-updating run builds into a scratch directory first and only touches a
live workbook through promote(). The service used to stop between those two
steps and wait for a human 'approve'; it no longer does — a clean build saves
itself:

    queued -> running -> promoted
                      \\-> failed

What the manual gate was protecting is still enforced inside promote(): a
timestamped backup of the workbook being replaced, and a hard stop if that
workbook is open in Excel. A failed build never reaches promote() at all.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import threading
import time
import uuid
from dataclasses import dataclass, field, asdict
from datetime import datetime, timezone
from enum import Enum
from pathlib import Path

from . import config
from .pipelines import STREAMS, JobContext, Stream


class JobStatus(str, Enum):
    QUEUED = "queued"
    RUNNING = "running"
    AWAITING_APPROVAL = "awaiting_approval"
    PROMOTED = "promoted"
    DISCARDED = "discarded"
    FAILED = "failed"


@dataclass
class StepResult:
    label: str
    script: str
    status: str = "pending"      # pending | running | ok | failed | skipped
    returncode: int | None = None
    seconds: float | None = None
    stdout_tail: str = ""
    stderr_tail: str = ""


@dataclass
class Job:
    id: str
    stream_key: str
    user_email: str
    created_at: str
    status: JobStatus = JobStatus.QUEUED
    steps: list[StepResult] = field(default_factory=list)
    uploaded_names: list[str] = field(default_factory=list)
    summary: dict = field(default_factory=dict)
    error: str = ""
    output_path: str = ""        # scratch output, before promotion
    promoted_to: str = ""        # live path, after promotion
    artifacts: list[str] = field(default_factory=list)   # extra files, e.g. PDFs
    finished_at: str = ""
    # The day the DATA covers, which is not the day it was uploaded. History is
    # read by business date — "did the 17th ever go in?" — and a day is often
    # uploaded late, or re-run three times before it sticks.
    raw_kept: bool = False      # this job archived its own copy of the raws
    covers: str = ""            # ISO date, or the first day of a period
    covers_to: str = ""         # ISO date, set only for a cumulative period
    # KSBC's secondary export is cumulative: 'SEPTEMBER 16TH' is every dispatch
    # up to the 16th, not the 16th alone. Read off the file itself at upload,
    # so history can say what a raw actually holds instead of only when it
    # arrived - which is what makes one upload look like fifteen days of data.
    spans_from: str = ""        # ISO date: the earliest day inside the raw
    # A line the upload itself worked out and history should keep - what a
    # batch turned out to contain, when the files decide that rather than the
    # dialog. Written at upload, shown in history, never inferred later.
    note: str = ""

    def to_dict(self) -> dict:
        d = asdict(self)
        d["status"] = self.status.value
        return d


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _tail(text: str, limit: int = 4000) -> str:
    text = text or ""
    return text if len(text) <= limit else "…\n" + text[-limit:]


class JobStore:
    """In-memory index over on-disk job directories.

    The authoritative record is the job directory on the persistent volume;
    this is just the fast lookup. A restart mid-build marks that job failed
    rather than leaving it eternally "running" — an honest state beats a
    hopeful one.
    """

    def __init__(self) -> None:
        self._jobs: dict[str, Job] = {}
        self._lock = threading.Lock()

    def create(self, stream_key: str, user_email: str) -> Job:
        job = Job(
            id=uuid.uuid4().hex[:12],
            stream_key=stream_key,
            user_email=user_email,
            created_at=_now(),
        )
        with self._lock:
            self._jobs[job.id] = job
        self.dir_for(job).mkdir(parents=True, exist_ok=True)
        self.persist(job)
        return job

    def dir_for(self, job: Job) -> Path:
        return config.JOBS_ROOT / job.id

    def get(self, job_id: str) -> Job | None:
        with self._lock:
            return self._jobs.get(job_id)

    def recent(self, limit: int = 40) -> list[Job]:
        with self._lock:
            jobs = sorted(self._jobs.values(), key=lambda j: j.created_at, reverse=True)
        return jobs[:limit]

    def forget(self, job_id: str) -> None:
        """Drop a job from the index. The caller removes its directory."""
        with self._lock:
            self._jobs.pop(job_id, None)

    def persist(self, job: Job) -> None:
        try:
            (self.dir_for(job) / "job.json").write_text(json.dumps(job.to_dict(), indent=2))
        except OSError:
            # Persistence of the index is best-effort; the build itself is what
            # matters and its outputs are already on disk.
            pass

    def load_existing(self) -> None:
        """Rehydrate the index from disk on startup."""
        if not config.JOBS_ROOT.exists():
            return
        for d in sorted(config.JOBS_ROOT.iterdir()):
            f = d / "job.json"
            if not f.is_file():
                continue
            try:
                raw = json.loads(f.read_text())
                raw["status"] = JobStatus(raw.get("status", "failed"))
                raw["steps"] = [StepResult(**s) for s in raw.get("steps", [])]
                job = Job(**raw)
            except Exception:
                continue
            if job.status in (JobStatus.RUNNING, JobStatus.QUEUED):
                job.status = JobStatus.FAILED
                job.error = "Service restarted while this build was in progress."
            with self._lock:
                self._jobs[job.id] = job


STORE = JobStore()

# One build at a time — the pipelines mutate shared workbook state in fixed
# folders, so concurrent runs on the same stream would interleave writes.
_BUILD_LOCK = threading.Semaphore(config.MAX_CONCURRENT_BUILDS)


def run_pipeline(job: Job, ctx: JobContext) -> None:
    """Execute the stream's steps in order. Blocking; call on a worker thread."""
    stream: Stream = STREAMS[job.stream_key]
    job.steps = [StepResult(label=s.display(), script=s.script) for s in stream.steps]
    # A stream with no steps has nothing to build: the upload itself is the
    # deliverable, and the report reads it where it lies.
    if not stream.steps:
        job.status = JobStatus.PROMOTED
        job.promoted_to = str(stream.input_path())
        job.finished_at = _now()
        STORE.persist(job)
        return

    job.status = JobStatus.RUNNING
    # warehouse_stock writes its dated workbook straight into the stream
    # folder, so "what did this build produce" cannot be answered by looking
    # at the folder afterwards - the newest file there may be last week's.
    # Remembered here so _resolve_output can tell a new workbook from an old
    # one, rather than reporting the latest it finds as this build's.
    ctx.before = {}
    if job.stream_key == "warehouse_stock":
        folder = STREAMS["warehouse_stock"].input_path()
        if folder.is_dir():
            ctx.before = {p.name: p.stat().st_mtime_ns
                          for p in folder.glob("* WAREHOUSE STOCK.xlsx")}
    STORE.persist(job)

    log_path = STORE.dir_for(job) / "build.log"
    log_lines: list[str] = []

    with _BUILD_LOCK:
        for i, step in enumerate(stream.steps):
            result = job.steps[i]
            result.status = "running"
            STORE.persist(job)

            env = os.environ.copy()
            env["PYTHONUNBUFFERED"] = "1"
            # The scripts import each other as siblings.
            env["PYTHONPATH"] = str(config.SCRIPTS_DIR) + os.pathsep + env.get("PYTHONPATH", "")
            env.update(step.env(ctx))

            from .pipelines import APP_REPORTS
            script_dir = APP_REPORTS if step.app_script else config.SCRIPTS_DIR
            argv = ["python3", str(script_dir / step.script), *step.args(ctx)]
            started = time.time()
            log_lines.append(f"\n=== [{i+1}/{len(stream.steps)}] {step.display()} ===\n$ {' '.join(argv)}\n")

            try:
                proc = subprocess.run(
                    argv,
                    cwd=str(config.CLAUDE_ROOT),   # scripts resolve paths from here
                    env=env,
                    capture_output=True,
                    text=True,
                    timeout=config.STEP_TIMEOUT_SECONDS,
                )
                result.returncode = proc.returncode
                result.stdout_tail = _tail(proc.stdout)
                result.stderr_tail = _tail(proc.stderr)
                log_lines.append(proc.stdout or "")
                if proc.stderr:
                    log_lines.append("\n--- stderr ---\n" + proc.stderr)
            except subprocess.TimeoutExpired:
                result.status = "failed"
                result.seconds = time.time() - started
                job.status = JobStatus.FAILED
                job.error = (
                    f"Step '{step.display()}' exceeded {config.STEP_TIMEOUT_SECONDS}s and was "
                    "stopped. Nothing was written to the live workbook."
                )
                log_lines.append("\n*** TIMEOUT ***\n")
                _finish(job, log_path, log_lines)
                return

            result.seconds = round(time.time() - started, 1)

            if result.returncode != 0:
                result.status = "failed"
                if step.fatal:
                    job.status = JobStatus.FAILED
                    job.error = _explain(step, result, proc.stdout,
                                        proc.stderr, job.stream_key)
                    for later in job.steps[i + 1:]:
                        later.status = "skipped"
                    _finish(job, log_path, log_lines)
                    return
            else:
                result.status = "ok"

            STORE.persist(job)

    # Resolve what the build produced and read any machine-readable summary.
    out = _resolve_output(job, ctx)
    if out is None:
        job.status = JobStatus.FAILED
        job.error = (
            "The pipeline finished but no output workbook was found. "
            "See the build log — this usually means a guard aborted quietly."
        )
        _finish(job, log_path, log_lines)
        return

    job.output_path = str(out)
    # Anything else the pipeline produced into the job directory is a
    # deliverable in its own right.
    job.artifacts = sorted(p.name for p in STORE.dir_for(job).glob("*.pdf"))
    summary_file = STORE.dir_for(job) / "summary.json"
    if summary_file.is_file():
        try:
            job.summary = json.loads(summary_file.read_text())
        except json.JSONDecodeError:
            pass

    # Auto-save. promote() still refuses to overwrite a workbook that is open
    # in Excel and still backs up whatever it replaces, so nothing that the
    # manual gate caught is lost — only the click.
    job.status = JobStatus.AWAITING_APPROVAL
    from . import promote as _promote  # local import: promote imports this module
    try:
        target = _promote.promote(job)
        log_lines.append(f"\n[auto-save] live workbook updated: {target}\n")
    except Exception as exc:  # PromotionError, or anything the filesystem threw
        job.status = JobStatus.FAILED
        job.error = f"The build finished but saving it failed: {exc}"
        log_lines.append(f"\n[auto-save] FAILED: {exc}\n")
    _finish(job, log_path, log_lines)


def _explain(step: Step, result: StepResult, stdout: str, stderr: str,
             stream_key: str = "") -> str:
    """Say what the script said, not just that it exited non-zero.

    These scripts fail loudly and usefully — 'ERROR: gap between existing
    coverage (up to day 0) and new raws (start at day 17)' tells the operator
    exactly what to do next. Reporting 'exit 4' throws that away and sends them
    to the log to find it. So the script's own last word is the error, with the
    exit code kept as a suffix for anyone debugging.
    """
    lines = ((stdout or "") + "\n" + (stderr or "")).splitlines()
    said = ""
    action = ""
    for i in range(len(lines) - 1, -1, -1):
        text = lines[i].strip()
        # The scripts name their failures - ERROR_MIXED_DATES,
        # ERROR_NO_ROWS_PARSED, ERROR_PARTIAL_DAY - not just "ERROR:". The
        # narrower test matched none of those, so the most useful sentence the
        # script printed was thrown away and the operator got the last line of
        # a traceback instead.
        if re.match(r"^(ERROR|ABORT|FATAL)[A-Z_]*\s*:", text, re.IGNORECASE):
            said = text.split(":", 1)[1].strip()
            # The ACTION line under it says what to do about it, which is the
            # half the operator actually needs.
            for follow in lines[i + 1:]:
                nxt = follow.strip()
                if nxt.upper().startswith("ACTION:"):
                    action = nxt.split(":", 1)[1].strip()
                    break
                if re.match(r"^[A-Z_]+\s*:", nxt):
                    break
            break
    if not said:
        for line in reversed((stderr or "").splitlines()):
            if line.strip():
                said = line.strip()
                break

    head = f"{step.display()} failed"
    # Only true for the streams that build in a scratch folder and promote
    # afterwards. build_warehouse_stock.py writes its workbook into the live
    # folder and appends the history CSVs as it goes, so by the time a step
    # fails the live data may already have changed - and telling the operator
    # otherwise sends them away without checking.
    if stream_key == "warehouse_stock" and result.returncode in (2, 3, 4, 5):
        # Every one of those guards stops before the workbook is saved and
        # before the history is touched, and the day's files are left where
        # they are. Saying "check the folder" after a clean refusal sends the
        # operator looking for damage that cannot be there.
        tail = ("Nothing was written and the day's files are still uploaded, so "
                "fixing this and uploading again is all it takes.")
    elif stream_key == "warehouse_stock":
        tail = ("The warehouse stock workbook and history are written in place, "
                "so check the stream folder before retrying — this run may have "
                "changed them.")
    else:
        tail = "The live workbook was NOT modified."
    if said:
        body = f"{head}: {said}"
        if action:
            body += f" What to do: {action}"
        return f"{body} (exit {result.returncode}) — {tail}"
    return (f"{head} (exit {result.returncode}). {tail} "
            "See the build log for what the script printed.")


def _finish(job: Job, log_path: Path, log_lines: list[str]) -> None:
    job.finished_at = _now()
    try:
        log_path.write_text("".join(log_lines))
    except OSError:
        pass
    STORE.persist(job)


def _resolve_output(job: Job, ctx: JobContext) -> Path | None:
    """Find the workbook a finished pipeline produced.

    Each stream writes somewhere different, and two of them were never designed
    around an explicit output argument — so the location is knowledge encoded
    here rather than guessed.
    """
    if job.stream_key == "shop_sales":
        candidate = STORE.dir_for(job) / "out.xlsx"     # KSBC_SCRATCH_XLSX
        return candidate if candidate.is_file() else None

    if job.stream_key == "secondary_sales":
        # build_secondary.py writes SCRATCH = f"{SESSION_ROOT}/sec_scratch.xlsx"
        scratch = config.WORKSPACE_ROOT / "sec_scratch.xlsx"
        if not scratch.is_file():
            return None
        landed = STORE.dir_for(job) / "out.xlsx"
        shutil.copyfile(scratch, landed)
        return landed

    if job.stream_key == "warehouse_stock":
        # Writes its dated workbook straight into the stream folder and appends
        # the history CSVs. Snapshot semantics, so there is no scratch stage —
        # see docs/KNOWN_DIFFERENCES.md.
        folder = STREAMS["warehouse_stock"].input_path()
        before = getattr(ctx, "before", None) or {}
        fresh = [p for p in folder.glob("* WAREHOUSE STOCK.xlsx")
                 if before.get(p.name) != p.stat().st_mtime_ns]
        if not fresh:
            # A build that wrote nothing used to hand back the newest workbook
            # in the folder - very often another date's - and the job reported
            # PROMOTED, pointing the operator at a file this run never touched.
            return None
        return max(fresh, key=lambda p: p.stat().st_mtime_ns)

    return None
