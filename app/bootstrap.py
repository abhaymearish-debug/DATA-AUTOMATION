"""First-boot seeding for a workspace that starts empty.

On this Mac the workspace is built by run_local.sh; on a server the disk comes
up bare, so the same job has to be done by the app itself. What gets seeded is
only ever REFERENCE material - the build scripts and the master data - never
sales figures. Anything already on the disk is left exactly as it is, so this
runs on every boot and does nothing after the first.
"""

from __future__ import annotations

import shutil
from pathlib import Path

from . import config

SEED = Path(__file__).resolve().parent.parent / "seed"


def _copy_if_absent(src: Path, dst: Path) -> bool:
    if dst.exists() or not src.exists():
        return False
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    return True


def _refresh(src: Path, dst: Path) -> bool:
    """Copy when the copy on the disk is not the one in this image.

    The build scripts are CODE, and the image is where they are versioned.
    Seeding them only when absent meant a fix shipped in a deploy never
    reached the disk they actually run from: the first boot's copy ran for
    ever, and a corrected script sat in the image doing nothing. Data is a
    different matter and stays on _copy_if_absent.
    """
    if not src.exists():
        return False
    if dst.exists() and dst.read_bytes() == src.read_bytes():
        return False
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    return True


def ensure_workspace() -> list[str]:
    """Create the layout the build scripts expect. Returns what was seeded."""
    done: list[str] = []

    for folder in (
        config.CLAUDE_ROOT / "KSBC shop sales" / "_cumulative",
        config.CLAUDE_ROOT / "Secondary sales",
        config.CLAUDE_ROOT / "Warehouse stock" / "_history",
        config.SCRIPTS_DIR,
        config.CLAUDE_ROOT / ".claude" / "memory",
        config.JOBS_ROOT,
    ):
        folder.mkdir(parents=True, exist_ok=True)

    # The 81 build scripts, verbatim. They are the whole point of this service:
    # the reports are their output, not a reimplementation of them.
    copied = 0
    for src in sorted((SEED / "scripts").glob("*.py")):
        if _refresh(src, config.SCRIPTS_DIR / src.name):
            copied += 1
    if copied:
        done.append(f"{copied} build script(s)")

    # Master data is reference, not sales: shop codes, names, bonds, clusters.
    # Without it every outlet reports as unmapped.
    if _copy_if_absent(SEED / config.MASTER_DATA.name, config.MASTER_DATA):
        done.append("master data")

    return done
