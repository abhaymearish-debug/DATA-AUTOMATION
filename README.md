# K.S. Distillery — Report Transformer

A web front end over the KSBC/Bevco report pipelines, so the office can upload a
raw export and get the finished workbook back without anyone running scripts by
hand.

---

## The design decision everything rests on

**The build scripts are not rewritten. They are wrapped.**

The `.claude/scripts/` pipelines have been producing these workbooks for months
and carry a great deal of hard-won business logic: the MUKKAM alias guard, the
KSBC 1–16 / 17–end cumulative rule, loose-bottle-to-case folding, raw-vs-retained
reconciliation, closed-shop pruning, the ordering constraints around comments and
merged-cell styling. Re-implementing that is where the previous attempt spent
seven months.

So this service reproduces the exact directory layout the scripts resolve from
`__file__`, drops the uploaded file into the right folder under the filename the
parser expects, and runs the script as a subprocess. The scripts stay
byte-identical.

**Verified:** the full Shop Sales formatting and insight tail run through this
wrapper over the live September workbook — 1,483,329 cells, 51 sheets — produced
**0 cells added, 0 removed, 5 values changed**, all five being floating-point
noise at the 13th decimal place. See `docs/VERIFICATION.md`.

---

## What it adds over running the scripts by hand

**The approval gate is a state machine, not a convention.** Every build goes
scratch → summary → explicit approval → live. Nothing touches a live workbook
until someone clicks Approve, and the file being replaced is backed up first.

**The pipeline order is in code.** It previously lived in the prose of scheduled
task prompts. That ordering is load-bearing and subtle — `ksbc_region_insights.py`
must be the last write because it applies hero styling in-session;
`ksbc_receipt_notes.py` must be penultimate because row moves drop comments;
`resize_comment_boxes.py` must run after the final openpyxl save or every comment
box reverts to a 144×79 default. `app/pipelines.py` encodes all of it, and a
`FORBIDDEN_SCRIPTS` check fails the build if a deprecated script ever creeps back.

**Filenames are canonicalised on upload.** The parsers need
`september 17th.xlsx`; the portal gives you `SupplierWiseShopSaleReport (3).xlsx`.
The app renames on the way in, or rejects with a message that says what is needed
— rather than failing deep inside a build.

**There is an audit trail.** Who uploaded what, when, which steps ran, how long
each took, and the full build log.

---

## Layout

```
app/
  config.py      Paths, auth config, startup checks
  pipelines.py   The codified step sequences  <- the heart of it
  jobs.py        Job runner and state machine
  promote.py     The ONLY code that writes to a live workbook
  auth.py        Company-email login, PBKDF2, signed sessions
  main.py        Routes
  templates/     Login and console
scripts/
  bootstrap_user.py
docs/
  DEPLOYMENT.md        How to deploy it, and who must own the accounts
  VERIFICATION.md      What was tested, and what has not been yet
  KNOWN_DIFFERENCES.md Where this behaves differently to the manual process
verify_pipeline.py     End-to-end harness against a real workbook
```

---

## Running locally

```bash
pip install -r requirements.txt

export KSD_WORKSPACE_ROOT=$PWD/workspace
export KSD_SESSION_SECRET=$(python3 -c "import secrets; print(secrets.token_urlsafe(48))")
export KSD_ALLOWED_EMAILS=you@example.com
export KSD_COOKIE_SECURE=0          # local HTTP only — never in production

python3 scripts/bootstrap_user.py you@example.com
uvicorn app.main:app --reload --port 8000
```

`GET /healthz` returns 503 and lists what is wrong until the workspace, secret
and allowlist are all in place.

The workspace layout is described in `docs/DEPLOYMENT.md` §5.

---

## v1 scope

Three streams: **Shop Sales**, **Secondary Sales**, **Warehouse Stock**. Two
users. Purchase Instruction, Target vs Achievement, Liquidation Summary and
Permit Status are deliberately out — shipping three reports that work beats
eleven that half do.

Adding a stream is a new entry in `STREAMS` with its ordered steps; no other
file needs to change.

## Not yet verified

The three **ingest** steps need raw exports that did not exist when this was
built. See `docs/VERIFICATION.md` — run each once, watched, against a copy of
the workspace when the next real raws land.
