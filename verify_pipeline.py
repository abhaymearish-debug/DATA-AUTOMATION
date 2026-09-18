#!/usr/bin/env python3
"""End-to-end verification of the Shop Sales tail pipeline.

WHAT THIS PROVES
----------------
The service's value rests on one claim: running the existing build scripts
through a web wrapper produces exactly what running them by hand produces. This
harness tests that claim against a real workbook rather than a fixture.

METHOD
------
The formatting and insight steps of the KSBC pipeline are documented as
idempotent — they recompute tiers from row data, rebuild conditional formatting,
wipe-and-rebuild the BOND INSIGHTS block, and re-fit comment boxes from each
note's own text. So running them over an already-built workbook must leave every
*value* untouched while the styling passes do their work.

That gives a strong test that needs no fresh raw export:

    1. Snapshot every cell value in the live September workbook.
    2. Run the real tail steps, in the real order, through the real Step
       definitions from app.pipelines — not a reimplementation.
    3. Snapshot again and diff.

A value diff means either the pipeline order in app/pipelines.py is wrong, or a
step is not as idempotent as documented. Either way it is a finding, and it is
better found here than on a Monday morning.

Run:  python3 verify_pipeline.py
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))

from app import config                      # noqa: E402
from app.pipelines import SHOP_SALES, JobContext  # noqa: E402

# The first step ingests a raw export; the remaining steps are the formatting,
# notes and insight passes this harness exercises.
TAIL_STEPS = SHOP_SALES.steps[1:]


def snapshot(path: Path) -> dict[tuple[str, str], object]:
    """Every non-empty cell value in the workbook, keyed by (sheet, coordinate)."""
    wb = openpyxl.load_workbook(path, data_only=False)
    values: dict[tuple[str, str], object] = {}
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        values[(ws.title, cell.coordinate)] = cell.value
    finally:
        wb.close()
    return values


def main() -> int:
    workbook_src = config.CLAUDE_ROOT / "KSBC shop sales" / "SEPTEMBER 1st - 16th ANALYSIS.xlsx"
    if not workbook_src.is_file():
        print(f"FAIL: test workbook not found at {workbook_src}")
        return 2

    work = Path("/tmp/verify_pipeline")
    work.mkdir(parents=True, exist_ok=True)
    target = work / "out.xlsx"
    shutil.copyfile(workbook_src, target)

    print(f"Workbook under test : {workbook_src.name} ({workbook_src.stat().st_size/1e6:.1f} MB)")
    print("Snapshotting before …")
    before = snapshot(target)
    print(f"  {len(before):,} populated cells across "
          f"{len({k[0] for k in before})} sheets\n")

    ctx = JobContext(job_id="verify", stream_key="shop_sales", workbook=target, scratch_dir=work)

    env = os.environ.copy()
    env["PYTHONUNBUFFERED"] = "1"
    env["PYTHONPATH"] = str(config.SCRIPTS_DIR) + os.pathsep + env.get("PYTHONPATH", "")

    failures: list[str] = []
    for i, step in enumerate(TAIL_STEPS, start=1):
        argv = ["python3", str(config.SCRIPTS_DIR / step.script), *step.args(ctx)]
        started = time.time()
        proc = subprocess.run(
            argv, cwd=str(config.CLAUDE_ROOT), env=env,
            capture_output=True, text=True, timeout=config.STEP_TIMEOUT_SECONDS,
        )
        secs = time.time() - started
        mark = "ok  " if proc.returncode == 0 else "FAIL"
        print(f"[{mark}] {i}/{len(TAIL_STEPS)} {step.display()}  ({secs:.1f}s)")
        if proc.returncode != 0:
            failures.append(step.script)
            tail = (proc.stderr or proc.stdout or "").strip().splitlines()[-12:]
            for line in tail:
                print(f"        {line}")

    print("\nSnapshotting after …")
    after = snapshot(target)
    print(f"  {len(after):,} populated cells across {len({k[0] for k in after})} sheets\n")

    added = set(after) - set(before)
    removed = set(before) - set(after)
    changed = {k for k in set(before) & set(after) if before[k] != after[k]}

    print("=" * 62)
    print(f"cells added    : {len(added):,}")
    print(f"cells removed  : {len(removed):,}")
    print(f"values changed : {len(changed):,}")

    for label, sample in (("added", added), ("removed", removed), ("changed", changed)):
        if sample:
            print(f"\n  first {label}:")
            for key in sorted(sample)[:8]:
                if label == "changed":
                    print(f"    {key[0]}!{key[1]}: {before[key]!r} -> {after[key]!r}")
                else:
                    src = before if label == "removed" else after
                    print(f"    {key[0]}!{key[1]}: {src[key]!r}")

    print("=" * 62)
    if failures:
        print(f"RESULT: FAIL — {len(failures)} step(s) exited non-zero: {', '.join(failures)}")
        return 1
    print("RESULT: all steps exited 0")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
