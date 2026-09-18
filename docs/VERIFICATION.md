# Verification record

## Run 1 — Shop Sales formatting & insight tail

**Date:** 17 September 2026
**Workbook:** `SEPTEMBER 1st - 16th ANALYSIS.xlsx` (5.6 MB, the live file)
**Harness:** `verify_pipeline.py`, driving the real `Step` definitions from
`app/pipelines.py` — not a reimplementation of the order.

### What was tested

The six formatting, notes and insight steps of the Shop Sales pipeline, run in
the order this service defines, over a real workbook. These steps are documented
as idempotent, so a correct run must leave every cell **value** untouched while
the styling passes rebuild fills, conditional formatting, notes and the BOND
INSIGHTS block.

### Result

| Step | | Time |
|---|---|---|
| 1 | Rebuild BOND PERFORMANCE clusterwise | 26.4s |
| 2 | Sort region / DETAIL / BOND PERFORMANCE by rating | 68.2s |
| 3 | Re-apply rating tier fills and fonts | 27.3s |
| 4 | DETAIL sheet styling sweep | 33.8s |
| 5 | Bake receipt-date notes on DETAIL sheets | 28.5s |
| 6 | BOND INSIGHTS + final styling (last write) | 39.3s |

All six exited 0. Total **223.5s**.

```
1,483,329 populated cells across 51 sheets
cells added    : 0
cells removed  : 0
values changed : 5
```

The five differences are floating-point representation only, at the **13th
decimal place**:

```
BOND PERFORMANCE!B5:  386.1319444444446 -> 386.1319444444445
BOND PERFORMANCE!B8:  484.2916666666666 -> 484.2916666666667
BOND PERFORMANCE!D13: 292.9444444444443 -> 292.9444444444444
BOND PERFORMANCE!D20: 245.6458333333333 -> 245.6458333333334
BOND PERFORMANCE!E7:  544.1041666666666 -> 544.1041666666665
```

Cases are twelfths and forty-eighths, so summation order shifts the last bit.
The magnitude is ~1e-13 cases against a reconciliation tolerance of 0.15 cases —
roughly a trillion times smaller than anything the pipeline treats as a
difference. This is not a data change.

**Conclusion:** the wrapper reproduces the hand-run pipeline faithfully, and the
step order encoded in `app/pipelines.py` is correct.

### The operational number this produced

A single step takes **26–68 seconds**, and the tail alone runs **3.7 minutes**.
With the ingest step a full Shop Sales build is roughly 4–5 minutes.

This settles the hosting question definitively. Vercel's serverless functions cap
out well below this. The service needs a long-running backend process — see
`docs/DEPLOYMENT.md`.

---

## Still to verify

The three **ingest** steps could not be exercised because they need raw exports
that do not exist yet:

- `ksbc_daily_update.py` — needs a Shop Sales raw for 17 September or later
  (the workbook currently covers 1–16).
- `build_secondary.py` — needs a single-day Secondary raw, to confirm the seed
  from `COMBINED DISPATCHES` plus day-merge and the raw-vs-retained guard.
- `build_warehouse_stock.py` — needs a day's set of ~28 `Report_*.xls` exports.

Each should be run through the app once, against a **copy** of the workspace, on
the next day real raws land. The approval gate means a bad build cannot reach a
live workbook, but the first run of each stream should still be watched.
