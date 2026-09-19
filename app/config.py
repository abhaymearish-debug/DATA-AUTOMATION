"""Configuration for the KSD Report Transformer.

THE CENTRAL DESIGN FACT
-----------------------
The existing build scripts are *directory-coupled*. They do not take a
"--base" argument and then behave; they resolve their own paths from
__file__ and then glob fixed folder names. For example:

    build_secondary.py
        CLAUDE       = <two levels up from scripts/>
        SESSION_ROOT = <two levels above that>
        SEC_DIR      = f"{CLAUDE}/Secondary sales"
        SCRATCH      = f"{SESSION_ROOT}/sec_scratch.xlsx"

    build_warehouse_stock.py
        FOLDER = Path(__file__).resolve().parent.parent.parent / "Warehouse stock"

So this service does NOT refactor the scripts. It reproduces the exact
directory layout they expect, drops the uploaded raw file into the right
stream folder under its canonical filename, and runs the script as a
subprocess. The scripts stay byte-identical to the ones that have been
producing these workbooks for months, which means every validated business
rule — the alias guard, the 1-16/17-end cumulative rule, loose-bottle
folding, raw-vs-retained reconciliation — carries over with zero
re-implementation risk.

Required layout on the server:

    WORKSPACE_ROOT/                 <- "SESSION_ROOT" to the scripts
      mnt/
        Claude/                     <- "CLAUDE" to the scripts
          MASTER DATA CONFIRMED.xlsx
          KSBC shop sales/
          Secondary sales/
          Warehouse stock/
            _history/
          .claude/
            scripts/                <- the build scripts, verbatim
            memory/
"""

from __future__ import annotations

import os
from pathlib import Path

# --------------------------------------------------------------------------
# Paths
# --------------------------------------------------------------------------

# Everything hangs off this. In development it points at a local checkout; in
# production it is a mounted persistent volume (the workbooks are STATE, not
# build artefacts — see the note on statefulness below).
WORKSPACE_ROOT = Path(
    os.environ.get("KSD_WORKSPACE_ROOT", Path(__file__).resolve().parent.parent / "workspace")
).resolve()

CLAUDE_ROOT = WORKSPACE_ROOT / "mnt" / "Claude"
SCRIPTS_DIR = CLAUDE_ROOT / ".claude" / "scripts"
MASTER_DATA = CLAUDE_ROOT / "MASTER DATA CONFIRMED.xlsx"

# Per-job scratch and the durable record of what was built.
JOBS_ROOT = Path(os.environ.get("KSD_JOBS_ROOT", WORKSPACE_ROOT / "_jobs")).resolve()

# --------------------------------------------------------------------------
# WHY THIS SERVICE IS STATEFUL
# --------------------------------------------------------------------------
# These pipelines are accumulators, not pure functions of their input:
#
#   * build_secondary.py SEEDS from the existing live month workbook's
#     "<MONTH> COMBINED DISPATCHES" sheet and layers the new single-day raw on
#     top. Raws are uploaded one day at a time and deleted after ingest, so the
#     month's prior days exist ONLY inside that workbook.
#   * ksbc_daily_update.py appends a day sheet to the month workbook and
#     refuses to ingest a day that would leave a gap.
#   * build_warehouse_stock.py appends to _history/stock_history.csv, which is
#     the canonical trend data every downstream artefact reads.
#
# Consequence: the workbook folders must live on a PERSISTENT VOLUME. A
# stateless container filesystem would silently reset the month on every
# redeploy. This is the single most important operational property of the
# deployment — see docs/DEPLOYMENT.md.
# --------------------------------------------------------------------------

REQUIRE_PERSISTENT_VOLUME = os.environ.get("KSD_ALLOW_EPHEMERAL", "") != "1"

# --------------------------------------------------------------------------
# Auth
# --------------------------------------------------------------------------

SESSION_SECRET = os.environ.get("KSD_SESSION_SECRET", "")

# Login is restricted to company addresses. K.S. Distillery uses Rediffmail for
# company email, so the allowed domains are configured rather than hardcoded to
# a single provider.
ALLOWED_EMAIL_DOMAINS = [
    d.strip().lower()
    for d in os.environ.get("KSD_ALLOWED_EMAIL_DOMAINS", "").split(",")
    if d.strip()
]

# Explicit allowlist wins over the domain rule when set. v1 is Abhay + Neelima.
ALLOWED_EMAILS = [
    e.strip().lower()
    for e in os.environ.get("KSD_ALLOWED_EMAILS", "").split(",")
    if e.strip()
]

# Who may add and remove accounts. Normally the first account created owns the
# install and this is left unset; setting it here overrides that, which is the
# way back in if the owner's own account is ever lost - no data is touched.
OWNER_EMAIL = os.environ.get("KSD_OWNER_EMAIL", "").strip().lower()

SESSION_MAX_AGE_SECONDS = int(os.environ.get("KSD_SESSION_MAX_AGE", 60 * 60 * 12))

# Session cookies are Secure by default. Only ever set KSD_COOKIE_SECURE=0 for
# local testing over plain HTTP — never in production, where the whole session
# would be readable on the wire.
COOKIE_SECURE = os.environ.get("KSD_COOKIE_SECURE", "1") != "0"

# --------------------------------------------------------------------------
# Build execution
# --------------------------------------------------------------------------

# Longest a single script step may run before it is killed. The heavy builds
# (KSBC region insights, PI analysis) genuinely take tens of seconds; the
# liquidation month extraction is ~30s per month. This is why the service needs
# a real backend host and not serverless functions.
STEP_TIMEOUT_SECONDS = int(os.environ.get("KSD_STEP_TIMEOUT", 900))

# One build at a time. The pipelines mutate shared workbook state in a fixed
# folder layout, so two concurrent runs on the same stream would interleave
# writes to the same files. Serialising is correct, not lazy.
MAX_CONCURRENT_BUILDS = 1


def config_problems() -> list[str]:
    """Startup checks. Returns human-readable problems; empty list means OK."""
    problems: list[str] = []

    if not SESSION_SECRET:
        problems.append(
            "KSD_SESSION_SECRET is not set. Generate one with "
            "`python3 -c \"import secrets; print(secrets.token_urlsafe(48))\"`."
        )
    if not ALLOWED_EMAILS and not ALLOWED_EMAIL_DOMAINS:
        problems.append(
            "Neither KSD_ALLOWED_EMAILS nor KSD_ALLOWED_EMAIL_DOMAINS is set — "
            "nobody would be able to sign in."
        )
    if not CLAUDE_ROOT.exists():
        problems.append(f"Workspace not found at {CLAUDE_ROOT} (set KSD_WORKSPACE_ROOT).")
    if not SCRIPTS_DIR.exists():
        problems.append(f"Build scripts not found at {SCRIPTS_DIR}.")
    if not MASTER_DATA.exists():
        problems.append(f"MASTER DATA CONFIRMED.xlsx not found at {MASTER_DATA}.")

    return problems
