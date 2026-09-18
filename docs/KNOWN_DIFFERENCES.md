# Where this differs from the manual process

Deliberate divergences, and why. Anything here is a decision, not an oversight —
if one turns out to be wrong, change it knowingly.

---

## 1. Raw files are kept, not deleted

**Manual:** the cleanup policy hard-`rm`s the raw export and any superseded
workbook in the same turn the new one is approved, on the grounds that the raw is
preserved inside the workbook as an audit sheet.

**Here:** raws stay where they were placed, and the replaced workbook is copied
into `_jobs/<id>/replaced/` before being overwritten.

**Why:** that policy was written for sessions with a person watching, who could
answer for an unexpected result immediately. A web service takes uploads from
someone who may not be looking at the output, so the cheapest possible undo is
worth more than a tidy folder. Disk is not the constraint — the whole workspace
is well under 100 MB.

If the folders need trimming later, add a scheduled job that prunes raws older
than N days. Do not put deletion in the build path.

---

## 2. Warehouse Stock has no scratch stage

**Manual and here:** `build_warehouse_stock.py` writes its dated workbook
straight into `Warehouse stock/` and appends the history CSVs.

**Consequence:** for this one stream the approval gate is informational. By the
time you see the summary, the workbook and the history append have already
happened.

**Why it was left alone:** the daily workbook is a *snapshot*, not a cumulative
document — a bad day can be rebuilt by re-uploading that day's exports. Adding a
scratch stage would mean intercepting the script's own output path and the
history append, which means modifying the script, which is exactly the thing this
design avoids.

**What to watch:** the history CSVs are appended before you approve. A duplicate
run for the same date updates that date's row rather than duplicating it (fixed
17 Jun 2026), so a re-run is safe. A run with the *wrong* date's files is not, and
would need the CSV rows removed by hand.

---

## 3. The build script's own raw deletion is inherited

`build_warehouse_stock.py` deletes the ingested `Report_*.xls` files at the end
of a successful build. That is inside the script, so it still happens.

It is safe — the history CSV is written before the delete — but it does mean the
warehouse raws behave differently to the other two streams, despite point 1.

---

## 4. Uploads are renamed; warehouse exports are not

Shop Sales and Secondary raws are renamed to the canonical form the parsers
require. Bevco stock exports keep their original names, because
`build_warehouse_stock.py` reads the report date out of the file itself and
aborts on mixed report dates — a rename would destroy information the build
relies on. Those are validated against the expected shape instead.

---

## 5. One build at a time

`MAX_CONCURRENT_BUILDS = 1`, and the container runs a single worker.

The pipelines mutate shared workbook state in fixed folder locations. Two
concurrent runs would interleave writes to the same files. If two people press
Build at once the second waits, which is correct.

Do not "fix" this by raising the worker count.

---

## 6. New months still need bootstrapping by hand

`ksbc_daily_update.py` only extends an existing month workbook; it has no
"start a new month" path. On the 1st of a month someone must still run:

```bash
python3 .claude/scripts/ksbc_bootstrap_month.py --month <MONTH> --folder "KSBC shop sales"
```

The app detects the missing workbook and says so plainly rather than failing
inside the driver, but it does not bootstrap automatically — creating a month is
a deliberate act, and doing it by accident from a mistyped date would be worse
than an error message.

Worth adding to the UI as an explicit, confirmed button once the three streams
have a few weeks of real use.

---

## 7. The Kerala dry day is not special-cased

The 1st of each month has no liquor activity, so a day-1 build legitimately shows
every shop at 0% sell-through and 🚫 Critical Overstock. That is accurate, not a
bug, and it normalises as real sales days land.

The app does not annotate it. If it confuses people in practice, a note on the
job summary for day-1 builds would be a small addition.
