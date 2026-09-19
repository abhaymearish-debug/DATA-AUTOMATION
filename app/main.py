"""FastAPI application."""

from __future__ import annotations

import re
import shutil
import threading
from datetime import date, datetime
from urllib.parse import urlencode
import subprocess
import sys
from pathlib import Path

from fastapi import FastAPI, Form, HTTPException, Request, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from . import (bootstrap, auth, bondmap, config, pi as pi_mod, promote as promote_mod,
               reports_api, reports_pdf, targets as targets_mod)
from .jobs import STORE, Job, JobContext, JobStatus, run_pipeline
from .pipelines import (
    MONTH_NAMES,
    STREAMS,
    UploadRejected,
    canonical_secondary_name,
    canonical_shop_sales_name,
    forbidden_in_pipelines,
    preflight,
    validate_warehouse_name,
)

app = FastAPI(title="KSD Report Transformer")
templates = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))

# The brand mark, and anywhere else the UI needs a file rather than markup.
# Mounted rather than inlined so the browser caches it once instead of
# re-downloading it inside every page.
_STATIC = Path(__file__).parent / "static"
_STATIC.mkdir(exist_ok=True)
app.mount("/static", StaticFiles(directory=str(_STATIC)), name="static")

MAX_UPLOAD_BYTES = 80 * 1024 * 1024

# The upload page is organised by DATA TYPE, which is how the office thinks
# about it, rather than by pipeline. Several cards can feed the same pipeline:
# a daily shop-sales file and a cumulative-period file are the same stream with
# a different date mode, and making that a card each is clearer than a free-text
# "cumulative period" box people have to know the convention for.
#
# mode: "day"   -> one date
#       "range" -> from/to, sent to the pipeline as a cumulative period
#       "files" -> no date, the export carries its own
UPLOAD_CARDS = [
    {"id": "wh_stock", "title": "Warehouse Physical Stock", "stream": "warehouse_stock",
     "mode": "files", "icon": "warehouse", "ready": True},
    {"id": "shop_daily", "title": "Shop Sales - Daily", "stream": "shop_sales_daily",
     "mode": "day", "icon": "shop", "ready": True},
    {"id": "shop_cum", "title": "Shop Sales - Cumulative", "stream": "shop_sales_cumulative",
     "mode": "range", "icon": "calendar", "ready": True},
    {"id": "secondary", "title": "Secondary Sales - Daily", "stream": "secondary_sales",
     "mode": "day", "icon": "truck", "ready": True},
    {"id": "pi_variance", "title": "Purchase Instruction", "stream": "purchase_instruction",
     "mode": "batch", "icon": "clipboard", "ready": True},
]

LEAVE_CARDS = [
    {"title": "Warehouse Leaves", "icon": "calendar-off"},
    {"title": "Shop Leaves", "icon": "calendar-off"},
]

# When the process started. Shown in the header so it is obvious at a glance
# whether the server you are looking at is the one you just restarted.
STARTED_AT = datetime.now().strftime("%d %b %H:%M")


@app.exception_handler(404)
async def page_not_found(request: Request, exc):
    """A missing PAGE gets an explanation; a missing API call stays JSON.

    Templates reload on refresh but routes do not, so a sidebar link can point
    at a route the running process has never heard of. Raw {"detail":"Not Found"}
    gives the operator nothing to act on; this says which half is stale.
    """
    wants_json = (request.url.path.startswith(("/api/", "/jobs/"))
                  or "application/json" in request.headers.get("accept", ""))
    if wants_json or not current_user(request):
        return JSONResponse({"detail": "Not Found"}, status_code=404)
    return templates.TemplateResponse(
        request, "notfound.html",
        {"user": current_user(request), "page": "", "started_at": STARTED_AT,
         "problems": getattr(app.state, "problems", [])},
        status_code=404,
    )


@app.middleware("http")
async def no_cache_html(request: Request, call_next):
    """Never let a browser cache a page of this app.

    All the CSS is inline, so a cached HTML page shows an old interface with no
    hint that anything is stale — which is indistinguishable from the code not
    having been deployed.
    """
    response = await call_next(request)
    if "text/html" in response.headers.get("content-type", ""):
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
    return response


def _drain_pending() -> None:
    """Clear builds still sitting at the approval step that no longer exists.

    Deliberately NOT promoted. A build that has been waiting since before this
    restart was made from an older snapshot of the data, and writing it over
    whatever is live now could quietly undo newer work. Discarding costs one
    re-upload; promoting a stale build costs a day of numbers. The scratch
    output stays on disk either way, so nothing is actually destroyed.
    """
    stuck = [j for j in STORE.recent(200) if j.status is JobStatus.AWAITING_APPROVAL]
    for job in stuck:
        job.status = JobStatus.DISCARDED
        job.error = "Left over from the old approval step — upload it again and it will save itself."
        STORE.persist(job)
    if stuck:
        print(f"[startup] discarded {len(stuck)} build(s) left at the old approval step")


@app.on_event("startup")
def _startup() -> None:
    # A server's disk comes up bare. Lay out the folders the build scripts
    # expect and seed the reference material before anything reads it.
    try:
        seeded = bootstrap.ensure_workspace()
        if seeded:
            print("[startup] seeded " + ", ".join(seeded))
    except OSError as exc:
        print(f"[startup] PROBLEM: could not prepare the workspace — {exc}")

    config.JOBS_ROOT.mkdir(parents=True, exist_ok=True)
    STORE.load_existing()
    _drain_pending()

    problems = config.config_problems() + forbidden_in_pipelines()
    app.state.problems = problems
    # Pipeline preflight only makes sense once the workspace exists.
    if not problems:
        app.state.problems = preflight()

    for p in app.state.problems:
        print(f"[startup] PROBLEM: {p}")


# ---------------------------------------------------------------------------
# Auth plumbing
# ---------------------------------------------------------------------------


def current_user(request: Request) -> str | None:
    return auth.read_session(request.cookies.get(auth.COOKIE_NAME))


def require_user(request: Request) -> str:
    user = current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Sign in first.")
    return user


@app.get("/login", response_class=HTMLResponse)
def login_form(request: Request, error: str = ""):
    return templates.TemplateResponse(
        request, "login.html", {"error": error, "first_run": not auth.load_users()})


# A server has no terminal to run bootstrap_user.py in, so the first password
# is set through the app itself. It is open only while there are NO accounts at
# all, and only to an address already on the allowlist - so it is not a sign-up
# page, it is the one-time handover of an install to its owner.
@app.get("/first-run", response_class=HTMLResponse)
def first_run_form(request: Request, error: str = ""):
    if auth.load_users():
        return RedirectResponse("/login", status_code=303)
    return templates.TemplateResponse(
        request, "first_run.html", {"error": error, "allowed": sorted(config.ALLOWED_EMAILS)})


@app.post("/first-run")
def first_run(request: Request, email: str = Form(...), password: str = Form(...),
              confirm: str = Form("")):
    if auth.load_users():
        return RedirectResponse("/login", status_code=303)

    def again(msg: str):
        return templates.TemplateResponse(
            request, "first_run.html",
            {"error": msg, "allowed": sorted(config.ALLOWED_EMAILS)}, status_code=400)

    if password != confirm:
        return again("Those two passwords are not the same.")
    try:
        auth.set_password(email, password)
    except ValueError as exc:
        return again(str(exc))
    return RedirectResponse("/login?error=Password+set.+Sign+in+with+it.", status_code=303)


@app.post("/login")
def login(request: Request, email: str = Form(...), password: str = Form(...)):
    who = auth.authenticate(email, password)
    if not who:
        # first_run belongs on THIS render too. Without it, one failed attempt
        # hid the only way to create the first account - which is exactly when
        # somebody signing in to a brand new install needs to see it.
        return templates.TemplateResponse(
            request,
            "login.html",
            {"error": ("No account has been set up on this install yet — "
                       "set your password first."
                       if not auth.load_users()
                       else "That email and password combination was not accepted."),
             "first_run": not auth.load_users()},
            status_code=401,
        )
    response = RedirectResponse("/", status_code=303)
    response.set_cookie(
        auth.COOKIE_NAME,
        auth.issue_session(who),
        httponly=True,
        samesite="lax",
        secure=config.COOKIE_SECURE,
        max_age=config.SESSION_MAX_AGE_SECONDS,
    )
    return response


# ---------------------------------------------------------------------------
# People
# ---------------------------------------------------------------------------
#
# Adding a colleague used to mean editing render.yaml and redeploying. It is a
# normal Tuesday thing to do, so it belongs in the app: anybody already signed
# in can add an account, and the address is allowed from that moment.


@app.get("/settings/people", response_class=HTMLResponse)
def people_page(request: Request, note: str = "", error: str = ""):
    user = require_user(request)
    return templates.TemplateResponse(
        request, "settings_people.html",
        {"user": user, "page": "people", "started_at": STARTED_AT,
         "problems": getattr(app.state, "problems", []),
         "people": sorted(auth.load_users()), "note": note, "error": error},
    )


@app.post("/settings/people/add")
def people_add(request: Request, email: str = Form(...), password: str = Form(...),
               confirm: str = Form("")):
    require_user(request)
    email = (email or "").strip().lower()

    def back(**kw):
        return RedirectResponse("/settings/people?" + urlencode(kw), status_code=303)

    if password != confirm:
        return back(error="Those two passwords are not the same.")
    if email in auth.load_users():
        return back(error=f"{email} already has an account.")
    try:
        auth.invite(email)          # allowed from now on
        auth.set_password(email, password)
    except ValueError as exc:
        return back(error=str(exc))
    return back(note=f"{email} can sign in now. Send them the password yourself — "
                     "it is not stored anywhere readable.")


@app.post("/settings/people/remove")
def people_remove(request: Request, email: str = Form(...)):
    me = require_user(request)
    email = (email or "").strip().lower()
    if email == me:
        return RedirectResponse(
            "/settings/people?" + urlencode({"error": "You cannot remove your own account."}),
            status_code=303)
    users = auth.load_users()
    users.pop(email, None)
    auth.save_users(users)
    auth.save_invited([e for e in auth.load_invited() if e != email])
    return RedirectResponse(
        "/settings/people?" + urlencode({"note": f"{email} can no longer sign in."}),
        status_code=303)


@app.post("/logout")
def logout():
    response = RedirectResponse("/login", status_code=303)
    response.delete_cookie(auth.COOKIE_NAME)
    return response


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    user = current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "user": user,
            "page": "upload",
            "started_at": STARTED_AT,
            "streams": list(STREAMS.values()),
            "streams_by_key": STREAMS,
            "cards": UPLOAD_CARDS,
            "leave_cards": LEAVE_CARDS,
            "jobs": STORE.recent(),
            "problems": getattr(app.state, "problems", []),
            "today": date.today().isoformat(),
        },
    )


@app.get("/healthz")
def healthz():
    problems = getattr(app.state, "problems", [])
    return JSONResponse(
        {"ok": not problems, "problems": problems},
        status_code=200 if not problems else 503,
    )


# ---------------------------------------------------------------------------
# Uploads and builds
# ---------------------------------------------------------------------------


def _resolve_shop_sales_workbook(month: int, year: int) -> Path:
    """Find the month's analysis workbook the driver should extend.

    A full-month file wins over a mid-month one. If neither exists the month
    has not been bootstrapped — say so plainly rather than letting the driver
    fail deep inside with a confusing message.
    """
    folder = STREAMS["shop_sales"].input_path()
    month_u = MONTH_NAMES[month - 1]

    full = folder / f"{month_u} SHOP SALES ANALYSIS.xlsx"
    if full.is_file():
        return full

    mid = sorted(
        folder.glob(f"{month_u} 1st - *ANALYSIS.xlsx"),
        key=lambda p: p.stat().st_mtime,
    )
    if mid:
        return mid[-1]

    # No workbook for this month yet. The driver cannot create one, so clone the
    # structural template — a workbook with every data sheet stripped out. The
    # month therefore starts empty and fills up only from what is uploaded.
    return _bootstrap_month(month_u, folder)


def _bootstrap_month(month_u: str, folder: Path) -> Path:
    """Start an empty month from the stripped template, via the locked script."""
    template = config.WORKSPACE_ROOT / "_seed" / "SHOP SALES TEMPLATE.xlsx"
    if not template.is_file():
        raise UploadRejected(
            f"No {month_u} workbook exists yet, and there is no shop-sales structure "
            "template to start one from. Shop Sales is the one pipeline that extends a "
            "month rather than building it, so it needs that template once. Re-run "
            "run_local.sh to build it from your own workbook."
        )

    out = folder / f"{month_u} 1st - 1st ANALYSIS.xlsx"
    script = config.SCRIPTS_DIR / "ksbc_bootstrap_month.py"
    proc = subprocess.run(
        [sys.executable, str(script), "--month", month_u,
         "--from", str(template), "--out", str(out), "--folder", str(folder)],
        capture_output=True, text=True, timeout=300,
    )
    if not out.is_file():
        raise UploadRejected(
            f"Could not start {month_u}: {(proc.stderr or proc.stdout or '').strip()[:300]}"
        )
    return out


def _save_upload(upload: UploadFile, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    size = 0
    with destination.open("wb") as fh:
        while chunk := upload.file.read(1024 * 1024):
            size += len(chunk)
            if size > MAX_UPLOAD_BYTES:
                fh.close()
                destination.unlink(missing_ok=True)
                raise UploadRejected(
                    f"'{upload.filename}' is larger than {MAX_UPLOAD_BYTES // (1024*1024)}MB."
                )
            fh.write(chunk)


@app.post("/build/{stream_key}")
async def build(
    request: Request,
    stream_key: str,
    files: list[UploadFile] = File(...),
    covers_date: str = Form(""),
    cumulative: str = Form(""),
):
    user = require_user(request)
    stream = STREAMS.get(stream_key)
    if stream is None:
        raise HTTPException(404, "Unknown report stream.")
    if getattr(app.state, "problems", []):
        raise HTTPException(503, "Service is not correctly configured; see /healthz.")

    job = STORE.create(stream_key, user)
    scratch_dir = STORE.dir_for(job)
    ctx = JobContext(job_id=job.id, stream_key=stream_key, scratch_dir=scratch_dir)

    try:
        placed: list[Path] = []

        for upload in files:
            suffix = Path(upload.filename or "").suffix.lower()
            if suffix not in stream.extensions:
                raise UploadRejected(
                    f"'{upload.filename}' is a {suffix or 'unknown'} file; "
                    f"{stream.label} expects {' or '.join(stream.extensions)}."
                )

        if stream_key == "warehouse_stock":
            for upload in files:
                name = validate_warehouse_name(upload.filename or "")
                target = stream.input_path() / name
                _save_upload(upload, target)
                placed.append(target)
            # Bevco writes the report date into the filename and the build
            # aborts on a mixed batch, so the first name speaks for all of them.
            # A date typed in the dialog wins: the operator is looking at the
            # file, and a renamed export should not silently file itself wrong.
            typed = _parse_date(covers_date) if covers_date else None
            job.covers = (typed.isoformat() if typed else
                          _warehouse_covers(job.uploaded_names or
                                            [p.name for p in placed]) or "")

        elif stream_key == "shop_sales_daily":
            # Standalone day: name it canonically and store it. No month
            # workbook to extend, so no contiguity question to answer.
            covers = _parse_date(covers_date)
            if len(files) != 1:
                raise UploadRejected("Upload one day at a time.")
            name = canonical_shop_sales_name(
                files[0].filename or "",
                day=covers.day, month=covers.month, year=covers.year,
            )
            target = stream.input_path() / name
            _save_upload(files[0], target)
            placed.append(target)
            job.covers = covers.isoformat()

        elif stream_key == "shop_sales":
            covers = _parse_date(covers_date)
            workbook = _resolve_shop_sales_workbook(covers.month, covers.year)
            ctx.meta["analysis_workbook"] = workbook
            job.summary["analysis_workbook"] = str(workbook)

            cum_range = _parse_cumulative(cumulative)
            if cum_range and len(files) != 1:
                raise UploadRejected("Upload one file at a time for a cumulative period.")
            # Answer the contiguity question now, in a second, instead of after
            # a three-minute build that the driver would abort anyway.
            _check_shop_day_gap(workbook, MONTH_NAMES[covers.month - 1],
                                covers.day, cum_range)

            for upload in files:
                name = canonical_shop_sales_name(
                    upload.filename or "",
                    day=covers.day,
                    month=covers.month,
                    year=covers.year,
                    cumulative=cum_range,
                )
                target = stream.input_path() / name
                _save_upload(upload, target)
                placed.append(target)

            if cum_range:
                job.covers = covers.replace(day=cum_range[0]).isoformat()
                job.covers_to = covers.replace(day=cum_range[1]).isoformat()
            else:
                job.covers = covers.isoformat()

        elif stream_key == "shop_sales_cumulative":
            # A period, stored whole. The two cumulative PDFs are drawn from
            # this file on demand, so nothing is built here and nothing can go
            # stale between the raw and the report.
            covers = _parse_date(covers_date)
            cum_range = _parse_cumulative(cumulative)
            if len(files) != 1:
                raise UploadRejected("Upload one cumulative period at a time.")
            if not cum_range:
                raise UploadRejected(
                    "A cumulative upload needs the period it covers - pick a From and a To date.")
            # Year-stamped, unlike the script-facing daily names: these files
            # are kept forever in one folder, and 'september 1-16' alone would
            # have next September overwrite this one.
            name = (f"{covers.year}-{covers.month:02d} {MONTH_NAMES[covers.month - 1]} "
                    f"{cum_range[0]}-{cum_range[1]} CUMULATIVE.xlsx")
            target = stream.input_path() / name
            target.parent.mkdir(parents=True, exist_ok=True)
            _save_upload(files[0], target)
            placed.append(target)
            job.covers = covers.replace(day=cum_range[0]).isoformat()
            job.covers_to = covers.replace(day=cum_range[1]).isoformat()

        elif stream_key == "secondary_sales":
            covers = _parse_date(covers_date)
            if len(files) != 1:
                raise UploadRejected("Secondary sales raws are uploaded one day at a time.")
            name = canonical_secondary_name(day=covers.day, month=covers.month, year=covers.year)
            target = stream.input_path() / name
            _save_upload(files[0], target)
            placed.append(target)
            job.covers = covers.isoformat()
            # What the file actually holds. A KSBC secondary pull is cumulative,
            # so one upload can carry a fortnight of dispatch; saying so here is
            # the difference between a report that looks invented and one that
            # can be traced back to the raw it came from.
            job.spans_from = _first_dispatch_day(target)

        elif stream_key == "purchase_instruction":
            # The month is in the files, not in the dialog: every instruction
            # names its own Report Month, and ~295 of them agreeing is a better
            # answer than one typed date. pi.store() files them under it, and
            # refuses a batch that turns out to span two months.
            staged = []
            for upload in files:
                target = scratch_dir / (upload.filename or "pi.xls")
                _save_upload(upload, target)
                staged.append(target)
            got = pi_mod.store(staged)
            if "error" in got:
                raise UploadRejected(got["error"])
            placed = sorted(pi_mod.month_dir(got["month"]).glob("*.xls*"))
            job.covers = f"{got['month']}-01"
            job.note = (f"{got['label']}: {got['shops']} shops"
                        + (f", {got['blank']} with no instruction" if got["blank"] else ""))

        ctx.uploaded = placed
        job.uploaded_names = [p.name for p in placed]
        # Keep the raws with the job. The live input folder is shared and the
        # warehouse build deletes what it consumes, so this is the only copy
        # that survives — and history is worth little if you cannot get back
        # the file a day was built from.
        _archive_raws(job, placed)
        STORE.persist(job)

    except UploadRejected as exc:
        job.status = JobStatus.FAILED
        job.error = str(exc)
        STORE.persist(job)
        return JSONResponse({"job_id": job.id, "error": str(exc)}, status_code=400)

    threading.Thread(target=run_pipeline, args=(job, ctx), daemon=True).start()
    return JSONResponse({"job_id": job.id}, status_code=202)


_WH_DATE_RE = re.compile(r"^Report[ _](\d{1,2})-([A-Za-z]{3})-(\d{4})", re.IGNORECASE)
_MON3 = {m[:3].upper(): i for i, m in enumerate(MONTH_NAMES, 1)}


def _warehouse_covers(names: list[str]) -> str:
    """The report date written into a Bevco export's filename, as ISO."""
    for name in names:
        m = _WH_DATE_RE.match(name)
        if not m:
            continue
        month = _MON3.get(m.group(2).upper())
        if not month:
            continue
        try:
            return date(int(m.group(3)), month, int(m.group(1))).isoformat()
        except ValueError:
            continue
    return ""


_DAY_SHEET_RE = r"^{month}\s+(\d+)(?:-(\d+))?$"
_CUM_SHEET_RE = r"^{month}\s+(\d+)-(\d+)\s+CUMULATIVE$"


def _shop_days_covered(workbook: Path, month_u: str) -> int:
    """Last day of the month this workbook already holds, 0 if none.

    Mirrors what ksbc_daily_update.py works out for itself, so the app can give
    the same answer before spending the time.
    """
    import openpyxl

    try:
        wb = openpyxl.load_workbook(workbook, read_only=True)
    except OSError:
        return 0
    last = 0
    for name in wb.sheetnames:
        for pattern in (_CUM_SHEET_RE, _DAY_SHEET_RE):
            m = re.match(pattern.format(month=month_u), name.strip(), re.IGNORECASE)
            if m:
                end = int(m.group(2) or m.group(1))
                last = max(last, end)
                break
    wb.close()
    return last


def _check_shop_day_gap(workbook: Path, month_u: str, day: int,
                        cum_range: tuple[int, int] | None) -> None:
    covered = _shop_days_covered(workbook, month_u)
    start = cum_range[0] if cum_range else day
    if start <= covered + 1:
        return

    missing = f"{covered + 1}" if covered + 1 == start - 1 else f"{covered + 1}–{start - 1}"
    month_t = month_u.title()
    if covered == 0:
        raise UploadRejected(
            f"{month_t} has nothing in it yet, and this file covers day {start}. "
            f"Shop sales is built day by day, so it cannot start at {start} — the "
            f"month's totals would be missing {month_t} {missing}. Either upload "
            f"those days first, or upload the {month_t} 1–{start - 1} period export "
            f"on the Shop Sales - Cumulative card and then this one."
        )
    raise UploadRejected(
        f"{month_t} currently covers up to day {covered}, and this file covers day "
        f"{start}. {month_t} {missing} would be missing, which would understate the "
        f"month. Upload the missing day(s) first."
    )


def _first_dispatch_day(raw: Path) -> str:
    """The earliest dispatch date inside a secondary raw, as ISO. '' if unreadable."""
    try:
        dates = [l["date"] for l in reports_api.load_dispatch_lines(raw) if l.get("date")]
    except Exception:
        return ""
    return min(dates).isoformat() if dates else ""


def _archive_raws(job: Job, placed: list[Path]) -> None:
    keep = STORE.dir_for(job) / "raw"
    keep.mkdir(parents=True, exist_ok=True)
    saved = 0
    for src in placed:
        try:
            shutil.copy2(src, keep / src.name)
            saved += 1
        except OSError:
            # A missing archive copy must never fail the upload itself — but
            # the flag has to stay honest, or the screen offers a download
            # that 404s.
            pass
    job.raw_kept = saved == len(placed) and saved > 0


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError:
        raise UploadRejected("Tell the app which date this export covers.")


def _parse_cumulative(value: str) -> tuple[int, int] | None:
    if not value.strip():
        return None
    m = re.match(r"^\s*(\d{1,2})\s*-\s*(\d{1,2})\s*$", value)
    if not m:
        raise UploadRejected("A cumulative period looks like '1-16' or '17-30'.")
    return int(m.group(1)), int(m.group(2))


# ---------------------------------------------------------------------------
# Job status, approval, download
# ---------------------------------------------------------------------------


@app.get("/jobs/{job_id}")
def job_status(request: Request, job_id: str):
    require_user(request)
    job = STORE.get(job_id)
    if job is None:
        raise HTTPException(404, "No such build.")
    return JSONResponse(job.to_dict())


@app.get("/jobs/{job_id}/log", response_class=HTMLResponse)
def job_log(request: Request, job_id: str):
    require_user(request)
    job = STORE.get(job_id)
    if job is None:
        raise HTTPException(404, "No such build.")
    log = STORE.dir_for(job) / "build.log"
    text = log.read_text() if log.is_file() else "(no log)"
    return HTMLResponse(f"<pre style='white-space:pre-wrap;font:12px/1.5 ui-monospace,monospace;padding:16px'>{_escape(text)}</pre>")


def _escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


@app.post("/jobs/{job_id}/approve")
def approve(request: Request, job_id: str):
    require_user(request)
    job = STORE.get(job_id)
    if job is None:
        raise HTTPException(404, "No such build.")
    try:
        target = promote_mod.promote(job)
    except promote_mod.PromotionError as exc:
        return JSONResponse({"error": str(exc)}, status_code=409)
    return JSONResponse({"ok": True, "promoted_to": str(target)})


@app.post("/jobs/{job_id}/discard")
def discard(request: Request, job_id: str):
    require_user(request)
    job = STORE.get(job_id)
    if job is None:
        raise HTTPException(404, "No such build.")
    try:
        promote_mod.discard(job)
    except promote_mod.PromotionError as exc:
        return JSONResponse({"error": str(exc)}, status_code=409)
    return JSONResponse({"ok": True})


@app.post("/api/uploads/{stream_key}/delete")
def delete_upload_date(request: Request, stream_key: str, day: str = Form(...)):
    """Remove an uploaded date: its runs, its raws, and the data it produced.

    Deleting an upload has to mean the reports stop showing it. The raws are
    the truth here; a month's analysis workbook is something the build made out
    of them, so once one of its days is deleted the workbook no longer
    describes anything that was uploaded and it goes too. Whatever raws remain
    answer on their own, which is how a month rebuilds itself down to nothing.

    Nothing is erased: the workbook is moved to _deleted/ in the workspace,
    where it can be fetched back by hand if this was a mistake.
    """
    require_user(request)
    if stream_key not in STREAMS:
        raise HTTPException(404, "No such data type.")
    stream = STREAMS[stream_key]

    jobs = [j for j in STORE.recent(500)
            if j.stream_key == stream_key and (j.covers or j.created_at[:10]) == day]
    if not jobs:
        raise HTTPException(404, "Nothing uploaded for that date.")

    removed = 0
    for job in jobs:
        for name in job.uploaded_names:
            live = stream.input_path() / name
            try:
                if live.is_file():
                    live.unlink()
                    removed += 1
            except OSError:
                pass
        shutil.rmtree(STORE.dir_for(job), ignore_errors=True)
        STORE.forget(job.id)

    note = ""
    if stream_key == "warehouse_stock":
        rows, books = _purge_stock_date(day)
        note = (f"Removed {rows} history row(s) and {books} day workbook(s) — "
                "the stock report will stop showing that date.")
    else:
        moved = _retire_derived(stream_key, day)
        left = _uploads_left(stream_key)
        if moved:
            note = (f"Removed {removed} raw file(s) and retired {len(moved)} built "
                    f"workbook(s) ({', '.join(moved)}) to _deleted/. ")
            note += ("The reports now show only what is still uploaded."
                     if left else "Nothing is uploaded for this report any more, "
                                  "so it will read as empty.")
        else:
            note = ("Removed the raw file(s). The reports now show only what is "
                    "still uploaded.")

    return JSONResponse({"ok": True, "runs": len(jobs), "raws": removed, "note": note})


# Which built workbooks belong to which stream. They are outputs, not sources:
# every one of them was made from raws that were uploaded, so when those raws
# go, so does the workbook - otherwise a report keeps quoting a file nobody can
# trace back to an upload.
_DERIVED = {
    "secondary_sales": ("Secondary sales", "*SECONDARY SALES ANALYSIS.xlsx"),
    "shop_sales": ("KSBC shop sales", "*SHOP SALES ANALYSIS.xlsx"),
}


def _retire_derived(stream_key: str, day: str) -> list[str]:
    """Move that month's built workbooks out of the way. Returns their names."""
    where = _DERIVED.get(stream_key)
    if not where:
        return []
    folder = config.CLAUDE_ROOT / where[0]
    if not folder.is_dir():
        return []

    try:
        month_u = MONTH_NAMES[date.fromisoformat(day).month - 1]
    except (ValueError, IndexError):
        return []

    bin_dir = config.WORKSPACE_ROOT / "_deleted" / datetime.now().strftime("%Y%m%d_%H%M%S")
    moved: list[str] = []
    for path in sorted(folder.glob(where[1])):
        if month_u not in path.name.upper():
            continue
        try:
            bin_dir.mkdir(parents=True, exist_ok=True)
            shutil.move(str(path), str(bin_dir / path.name))
            moved.append(path.name)
        except OSError:
            pass
    return moved


def _uploads_left(stream_key: str) -> bool:
    """Is anything still uploaded for this stream?"""
    return any(j.stream_key == stream_key for j in STORE.recent(500))


def _purge_stock_date(day: str) -> tuple[int, int]:
    """Drop one report date from the warehouse stock history and day workbooks."""
    import csv

    rows_removed = 0
    hist = config.CLAUDE_ROOT / "Warehouse stock" / "_history"
    if hist.is_dir():
        for path in sorted(hist.glob("*.csv")):
            try:
                with path.open(newline="") as fh:
                    rows = list(csv.reader(fh))
            except OSError:
                continue
            if not rows:
                continue
            header = rows[0]
            if "date" not in header:
                continue
            ix = header.index("date")
            kept = [r for r in rows[1:] if not (len(r) > ix and r[ix] == day)]
            dropped = len(rows) - 1 - len(kept)
            if not dropped:
                continue
            try:
                with path.open("w", newline="") as fh:
                    writer = csv.writer(fh)
                    writer.writerow(header)
                    writer.writerows(kept)
                rows_removed += dropped
            except OSError:
                continue

    books = 0
    try:
        d = date.fromisoformat(day)
    except ValueError:
        return rows_removed, books
    target = (config.CLAUDE_ROOT / "Warehouse stock" /
              f"{MONTH_NAMES[d.month - 1]} {_day_ordinal(d.day)} WAREHOUSE STOCK.xlsx")
    try:
        if target.is_file():
            target.unlink()
            books += 1
    except OSError:
        pass
    return rows_removed, books


def _day_ordinal(n: int) -> str:
    if 10 <= n % 100 <= 20:
        return f"{n}th"
    return f"{n}{ {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th') }"


@app.get("/jobs/{job_id}/raw")
def download_raw(request: Request, job_id: str, name: str = ""):
    """The raw file(s) this upload was built from — one file, or all as a zip."""
    require_user(request)
    job = STORE.get(job_id)
    if job is None:
        raise HTTPException(404, "No such build.")

    folder = STORE.dir_for(job) / "raw"
    have = [n for n in job.uploaded_names if (folder / n).is_file()]
    if not have:
        raise HTTPException(404, "The raw files for this upload are no longer on disk.")

    if name:
        if name not in job.uploaded_names:
            raise HTTPException(404, "No such file in this upload.")
        one = folder / name
        if not one.is_file():
            raise HTTPException(404, "That file is no longer on disk.")
        return FileResponse(one, filename=name)

    if len(have) == 1:
        return FileResponse(folder / have[0], filename=have[0])

    import tempfile
    import zipfile

    label = STREAMS[job.stream_key].label
    day = job.covers or job.created_at[:10]
    out = Path(tempfile.mkdtemp()) / f"{label} — {day} (raw files).zip"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in have:
            z.write(folder / n, n)
    return FileResponse(out, filename=out.name)


@app.get("/jobs/{job_id}/download")
def download(request: Request, job_id: str, name: str = ""):
    require_user(request)
    job = STORE.get(job_id)
    if job is None:
        raise HTTPException(404, "No such build.")

    if name:
        # Only ever serve a file this job actually produced, by its own listed
        # name — never a path assembled from user input.
        if name not in job.artifacts:
            raise HTTPException(404, "No such file in this build.")
        artifact = STORE.dir_for(job) / name
        if not artifact.is_file():
            raise HTTPException(404, "That file is no longer on disk.")
        return FileResponse(artifact, filename=name)

    if not job.output_path:
        raise HTTPException(404, "Nothing to download.")
    path = Path(job.output_path)
    if not path.is_file():
        raise HTTPException(404, "The output file is no longer on disk.")

    stream = STREAMS[job.stream_key]
    nice = f"{stream.label} — {job.created_at[:10]}.xlsx"
    return FileResponse(path, filename=nice)


# ---------------------------------------------------------------------------
# Reports (read-only)
# ---------------------------------------------------------------------------


@app.get("/reports/brandwise", response_class=HTMLResponse)
def brandwise_page(request: Request):
    user = current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    return templates.TemplateResponse(
        request, "report_brandwise.html",
        {"user": user, "page": "brandwise", "started_at": STARTED_AT,
         "problems": getattr(app.state, "problems", [])},
    )


def _report_args(date_from: str, date_to: str):
    def d(v):
        try:
            return date.fromisoformat(v) if v else None
        except ValueError:
            return None
    return d(date_from), d(date_to)


@app.get("/api/reports/brandwise")
def brandwise_data(
    request: Request,
    view: str = "bond",
    date_from: str = "",
    date_to: str = "",
    bond: str = "",
    warehouse: str = "",
    round_off: str = "",
):
    require_user(request)
    if view not in ("bond", "warehouse", "shop"):
        raise HTTPException(400, "Unknown view.")
    f, t = _report_args(date_from, date_to)
    return JSONResponse(reports_api.brandwise(
        view=view, date_from=f, date_to=t, bond=bond,
        warehouse=warehouse, round_off=bool(round_off),
    ))


@app.get("/api/reports/brandwise/shops")
def brandwise_shops_data(
    request: Request,
    view: str = "bond",
    key: str = "",
    date_from: str = "",
    date_to: str = "",
    brands: str = "",
    round_off: str = "",
):
    """One group's shops, for the drill-down under a bond or warehouse row."""
    require_user(request)
    if view not in ("bond", "warehouse"):
        raise HTTPException(400, "Unknown view.")
    f, t = _report_args(date_from, date_to)
    cols = [b for b in brands.split("\u001f") if b] if brands else None
    return JSONResponse(reports_api.brandwise_shops(
        view=view, key=key, date_from=f, date_to=t,
        round_off=bool(round_off), brands=cols,
    ))


@app.get("/reports/brandwise/export.xlsx")
def brandwise_xlsx(
    request: Request,
    view: str = "bond",
    date_from: str = "",
    date_to: str = "",
    bond: str = "",
    warehouse: str = "",
    round_off: str = "",
):
    require_user(request)
    f, t = _report_args(date_from, date_to)
    data = reports_api.brandwise(view=view, date_from=f, date_to=t, bond=bond,
                                 warehouse=warehouse, round_off=bool(round_off))
    if "error" in data:
        raise HTTPException(404, data["error"])

    import tempfile
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = f"BRANDWISE {view.upper()}"[:31]

    navy, gold, white = "FF0D1B4A", "FFFFB300", "FFFFFFFF"
    ws.append([data["label"]] + data["brands"] + ["Total"])
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(bold=True, color=white, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 42

    for r in data["rows"]:
        ws.append([r["name"]] + [r["cells"].get(b, 0) for b in data["brands"]] + [r["total"]])
        if r["kind"] == "cluster":
            for cell in ws[ws.max_row]:
                cell.fill = PatternFill("solid", fgColor=navy)
                cell.font = Font(bold=True, color=gold, size=10)

    ws.append(["GRAND TOTAL"] + [data["grand"]["cells"].get(b, 0) for b in data["brands"]]
              + [data["grand"]["total"]])
    for cell in ws[ws.max_row]:
        cell.fill = PatternFill("solid", fgColor="FF374151")
        cell.font = Font(bold=True, color=white, size=10)

    ws.column_dimensions["A"].width = 34
    for i in range(2, len(data["brands"]) + 3):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = 15
    ws.freeze_panes = "B2"

    tmp = Path(tempfile.mkdtemp()) / "Secondary Sales - Cumulative.xlsx"
    wb.save(tmp)
    return FileResponse(tmp, filename=tmp.name)


@app.get("/reports/brandwise/export.pdf")
def brandwise_pdf(
    request: Request,
    scope: str = "cluster",
    cluster: int = 1,
    view: str = "bond",
    date_from: str = "",
    date_to: str = "",
    bond: str = "",
    warehouse: str = "",
    round_off: str = "",
):
    """Two shapes of PDF from one endpoint.

    scope="cluster"  one ASM cluster, in the fixed warehouse-per-page layout the
                     field team already receives. Built by the same script the
                     Secondary Sales pipeline runs, so a PDF downloaded here is
                     byte-for-byte the kind of artefact a build produces.
    scope="current"  exactly what the page is showing — the chosen view, the
                     live filters and date range, cluster subtotals and grand
                     total. Rendered from the same brandwise() result that fed
                     the table, so the two cannot drift apart.
    """
    require_user(request)

    # Only the official cluster books need the month's workbook - they are built
    # by the same script the pipeline runs. The current view renders from the
    # figures already on screen, which the raws alone can answer for.
    workbook = reports_api.find_secondary_workbook()
    if workbook is None and scope != "current":
        raise HTTPException(404, "No secondary-sales workbook has been built yet. "
                                 "The current view still exports.")

    import subprocess, tempfile
    from .pipelines import APP_REPORTS

    outdir = Path(tempfile.mkdtemp())

    if scope == "current":
        if view not in ("bond", "warehouse", "shop"):
            raise HTTPException(400, "Unknown view.")
        f, t = _report_args(date_from, date_to)
        data = reports_api.brandwise(view=view, date_from=f, date_to=t, bond=bond,
                                     warehouse=warehouse, round_off=bool(round_off))
        if "error" in data:
            raise HTTPException(404, data["error"])
        if not data["rows"]:
            raise HTTPException(404, "Nothing matches those filters.")

        name = f"Secondary Sales - Cumulative ({data['label']} View).pdf"
        out = reports_pdf.build_view_pdf(
            data, outdir / name, round_off=bool(round_off),
            filters={"bond": bond, "warehouse": warehouse,
                     "date_from": date_from, "date_to": date_to},
        )
        return FileResponse(out, filename=out.name)

    if cluster not in (1, 2, 3):
        raise HTTPException(400, "Cluster must be 1, 2 or 3.")

    argv = ["python3", str(APP_REPORTS / "build_secondary_brandwise_pdfs.py"),
            "--base", str(config.CLAUDE_ROOT), "--workbook", str(workbook),
            "--outdir", str(outdir), "--cluster", str(cluster)]
    if date_from:
        argv += ["--from", date_from]
    if date_to:
        argv += ["--to", date_to]

    proc = subprocess.run(argv, capture_output=True, text=True,
                          timeout=config.STEP_TIMEOUT_SECONDS)
    if proc.returncode != 0:
        raise HTTPException(500, f"PDF build failed: {(proc.stderr or proc.stdout)[-400:]}")

    pdfs = list(outdir.glob("*.pdf"))
    if not pdfs:
        raise HTTPException(404, f"No dispatches for cluster {cluster} in the selected range.")
    return FileResponse(pdfs[0], filename=pdfs[0].name)


@app.get("/api/jobs")
def jobs_list(request: Request, stream: str = ""):
    """Recent builds, optionally for one stream.

    The upload page reads this instead of scraping a table out of the DOM, so
    the History dialog and the in-progress strip stay correct even though the
    page no longer renders a build list of its own.
    """
    require_user(request)
    # History is a permanent record: when a stream is named, every run it has
    # ever had comes back, not a recent window. The upload page asks without a
    # stream and only needs enough to colour the cards.
    limit = 5000 if stream else 60
    jobs = [j for j in STORE.recent(limit) if not stream or j.stream_key == stream]
    return JSONResponse({
        "jobs": [j.to_dict() for j in jobs],
        "labels": {k: v.label for k, v in STREAMS.items()},
    })


# ---------------------------------------------------------------------------
# Warehouse Stock Report
# ---------------------------------------------------------------------------


@app.get("/status-calendar", response_class=HTMLResponse)
def status_calendar(request: Request):
    """Which days the reports can answer for, and which they cannot."""
    user = require_user(request)
    return templates.TemplateResponse(
        request, "status_calendar.html",
        {"user": user, "page": "calendar", "started_at": STARTED_AT,
         "problems": getattr(app.state, "problems", []),
         "calendar": reports_api.upload_calendar()},
    )


@app.get("/reports/warehouse-stock", response_class=HTMLResponse)
def stock_page(request: Request):
    user = current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    return templates.TemplateResponse(
        request, "report_stock.html",
        {"user": user, "page": "stock", "started_at": STARTED_AT,
         "problems": getattr(app.state, "problems", [])},
    )


@app.get("/api/reports/warehouse-stock")
def stock_data(request: Request, as_of: str = "", cluster: str = "", warehouse: str = ""):
    require_user(request)
    cl = int(cluster) if cluster in ("1", "2", "3") else None
    return JSONResponse(reports_api.warehouse_stock(as_of=as_of, cluster=cl,
                                                    warehouse=warehouse))


# ---------------------------------------------------------------------------
# Shop sales - cumulative and analysis
# ---------------------------------------------------------------------------
#
# Both are drawn from the stored cumulative raw on demand rather than baked at
# upload time. That is what makes a period from months ago still openable, and
# it means a change to Bond Mapping shows up in every past period at once
# instead of only in whatever is rebuilt next.


def _analysis_module():
    """The PDF builder, imported as a module so the page and the PDF agree."""
    import importlib.util
    from .pipelines import APP_REPORTS
    spec = importlib.util.spec_from_file_location(
        "build_shop_analysis_pdf", APP_REPORTS / "build_shop_analysis_pdf.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _period_or_404(period: str) -> dict:
    chosen = reports_api.cumulative_period(period)
    if chosen is None:
        raise HTTPException(404, "No cumulative period has been uploaded yet.")
    return chosen


def _asked_dates(date_from: str, date_to: str, period: str = ""):
    """The window the caller actually asked for, so a refusal can name it."""
    if not (date_from and date_to) and period and ".." in period:
        date_from, date_to = period.split("..", 1)
    if not (date_from and date_to):
        return None
    try:
        start = datetime.strptime(date_from, "%Y-%m-%d").date()
        end = datetime.strptime(date_to, "%Y-%m-%d").date()
    except ValueError:
        return None
    return (start, end) if start <= end else (end, start)


def _asked_period(date_from: str, date_to: str, period: str = ""):
    got = _asked_dates(date_from, date_to, period)
    return reports_api.period_for(*got) if got else None


def _suggested_window():
    """The widest answerable window, for a screen that has to refuse one."""
    widest = reports_api.widest_window()
    if not widest:
        return None
    return {"from": widest[0].isoformat(), "to": widest[1].isoformat(),
            "short": reports_api.period_for(*widest)["short"]}


def _window(date_from: str, date_to: str, period: str = ""):
    """Whatever the caller asked for, as a window and a source.

    A calendar range is the primary input now; `period` is still honoured so an
    older bookmark or a link from another screen keeps working.
    """
    if not (date_from and date_to) and period and ".." in period:
        date_from, date_to = period.split("..", 1)
    if date_from and date_to:
        try:
            start = datetime.strptime(date_from, "%Y-%m-%d").date()
            end = datetime.strptime(date_to, "%Y-%m-%d").date()
        except ValueError:
            raise HTTPException(400, "Those dates are not readable.")
        if end < start:
            start, end = end, start
        return reports_api.resolve_window(start, end)

    # Default to the widest window the uploads can answer for, rather than the
    # newest file: with 1-16 and 17-23 on disk, 1-23 is the report you want.
    widest = reports_api.widest_window()
    if widest is None:
        return {"error": "Nothing has been uploaded yet."}
    return reports_api.resolve_window(*widest)


def _window_data(win: dict, cluster: int, bond: str, warehouse: str = "",
                 group_by: str = "bond") -> dict:
    """The grouping for a resolved window, whatever its source."""
    if win["source"] == "cumulative":
        return reports_api.shop_cumulative(win["period"], cluster or None, bond,
                                           warehouse, group_by)
    slim = {k: win["period"][k] for k in ("key", "short", "long", "days")}
    return reports_api._group_shops(win["shops"], slim, cluster or None, bond,
                                    warehouse, group_by)


@app.get("/reports/shop-cumulative", response_class=HTMLResponse)
def shop_cumulative_page(request: Request, date_from: str = "", date_to: str = "",
                         period: str = ""):
    user = require_user(request)
    periods = reports_api.cumulative_periods()
    # Seed the page with the same window the API would default to - the widest
    # the uploads can answer for. Taking the newest file instead would open on
    # 17-23 while a 1-16 sits next to it, which is not the report anyone wants.
    # An explicit window in the link wins, so a shared URL opens on its period.
    win = _window(date_from, date_to, period)
    chosen = None if "error" in win else win["period"]
    if chosen is None:
        asked = _asked_period(date_from, date_to, period)
        if asked:
            chosen = asked
    return templates.TemplateResponse(
        request, "report_shop_cumulative.html",
        {"user": user, "page": "shop_cumulative", "started_at": STARTED_AT,
         "problems": getattr(app.state, "problems", []),
         "periods": periods, "chosen": chosen},
    )


@app.get("/api/shop-cumulative")
def shop_cumulative_api(request: Request, date_from: str = "", date_to: str = "",
                        period: str = "", cluster: int = 0, bond: str = "",
                        warehouse: str = "", group_by: str = "bond"):
    """Bond (or warehouse) rows with their shops underneath, for any window."""
    require_user(request)
    win = _window(date_from, date_to, period)
    if "error" in win:
        # A window we cannot tile is a normal thing to ask for, not a failure:
        # say what is missing, and offer the widest window that IS answerable
        # so the screen has something to act on rather than a dead end.
        asked = _asked_period(date_from, date_to, period)
        return JSONResponse({"error": win["error"], "calendar": _calendar_days(),
                             "asked": {"short": asked["short"]} if asked else None,
                             "suggest": _suggested_window()})

    data = _window_data(win, cluster, bond, warehouse, group_by)
    if "error" not in data:
        data["source"] = win["source"]
        data["chain"] = win.get("chain", [])
        data["all_warehouses"] = data.pop("warehouses", [])
        data["cluster"] = cluster
        data["bond"] = (bond or "").upper()
        data["warehouse"] = (warehouse or "").upper()
        data["calendar"] = _calendar_days()
    return JSONResponse(data)


def _calendar_days() -> dict:
    """What the calendar can offer: whole uploaded periods, and single days."""
    return {
        "periods": [{"from": p["start"].isoformat(), "to": p["end"].isoformat(),
                     "short": p["short"]} for p in reports_api.cumulative_periods()],
        "days": sorted(d.isoformat() for d in reports_api.shop_day_files()),
    }


@app.get("/api/shop-cumulative/shop")
def shop_cumulative_detail(request: Request, code: str, date_from: str = "",
                           date_to: str = "", period: str = ""):
    """One shop's brand and pack lines, fetched when its row is opened."""
    require_user(request)
    win = _window(date_from, date_to, period)
    if "error" in win:
        return JSONResponse({"error": win["error"]})
    cur = win["period"]
    return JSONResponse(reports_api.shop_detail(cur["start"], cur["end"], code))


@app.get("/reports/shop-cumulative/export.xlsx")
def shop_cumulative_xlsx(request: Request, date_from: str = "", date_to: str = "",
                         period: str = "", cluster: int = 0, bond: str = "",
                         warehouse: str = "", group_by: str = "bond"):
    """The screen, as a workbook: bonds grouped, shops collapsible under them.

    The earlier flat dump pivoted well and read badly - four hundred rows with
    the bond repeated on every one, and nothing to tell a shop line from a
    total. This keeps the grouping people actually work with and lets Excel
    fold each bond away, while the autofilter still gives the flat view to
    anyone who wants to pivot.
    """
    require_user(request)
    win = _window(date_from, date_to, period)
    if "error" in win:
        raise HTTPException(404, win["error"])
    chosen = win["period"]
    data = _window_data(win, cluster, bond, warehouse, group_by)
    if "error" in data:
        raise HTTPException(404, data["error"])
    if not data["bonds"]:
        raise HTTPException(404, "Nothing matches that filter.")

    import tempfile
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    NAVY, GOLD, PAPER = "FF0A294F", "FFFFBD30", "FFF5F7FC"
    INK, HAIR = "FF28324A", "FFD9DEE9"
    thin = Side(style="thin", color=HAIR)
    box = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = Workbook()
    ws = wb.active
    ws.title = "SHOP SALES CUMULATIVE"
    grouping = "WAREHOUSE" if group_by == "warehouse" else "BOND"
    headings = (["CLUSTER"] if grouping == "BOND" else []) + [
        grouping, "SHOP CODE", "SHOP", "OPENING", "RECEIPT", "SALES", "CLOSING"]
    last_col = get_column_letter(len(headings))

    # A title anyone can read six months later, without opening the filename.
    ws.merge_cells(f"A1:{last_col}1")
    t = ws["A1"]
    t.value = f"SHOP SALES CUMULATIVE   ·   {chosen['long']}"
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.font = Font(bold=True, color=GOLD, size=13)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    scope = (bond.title() if bond else warehouse.title() if warehouse
             else f"Cluster {cluster}" if cluster else "All bonds")
    ws.merge_cells(f"A2:{last_col}2")
    sub = ws["A2"]
    sub.value = (f"{scope}   ·   {data['shop_count']} shops in {len(data['bonds'])} bonds"
                 f"   ·   cases")
    sub.fill = PatternFill("solid", fgColor=GOLD)
    sub.font = Font(bold=True, color=NAVY, size=10)
    sub.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    ws.append(headings)
    for cell in ws[3]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(bold=True, color=GOLD, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = box
    ws.row_dimensions[3].height = 22

    detail = reports_api.window_lines(chosen["start"], chosen["end"])
    lines = {} if "error" in detail else detail["lines"]
    by_shop: dict = {}
    for key, values in lines.items():
        shop_code, brand, pack = key.split("|", 2)
        by_shop.setdefault(shop_code, {}).setdefault(brand, {})[pack] = values

    for g in data["bonds"]:
        for i, shop in enumerate(g["shops"]):
            code = shop["code"]
            ws.append([g["cluster"] or "", g["bond"],
                       int(code) if code.isdigit() and not code.startswith("0") else code,
                       shop["name"], *shop["values"]])
            r = ws.max_row
            # Three collapsible levels, the way the screen folds: bond, then
            # shop, then the brand and pack lines under it.
            ws.row_dimensions[r].outlineLevel = 1
            band = PatternFill("solid", fgColor=PAPER) if i % 2 else None
            for cell in ws[r]:
                cell.border = box
                cell.font = Font(bold=True, size=10, color=INK)
                if band:
                    cell.fill = band

            for brand in sorted(by_shop.get(code, {})):
                packs = by_shop[code][brand]
                sub = [sum(packs[pk][k] for pk in packs) for k in range(4)]
                ws.append(["", "", "", f"    {brand}", *[round(v, 2) for v in sub]])
                ws.row_dimensions[ws.max_row].outlineLevel = 2
                for cell in ws[ws.max_row]:
                    cell.border = box
                    cell.font = Font(bold=True, size=9, color=INK)
                for pack in sorted(packs):
                    ws.append(["", "", "", f"        {pack}",
                               *[round(v, 2) for v in packs[pack]]])
                    ws.row_dimensions[ws.max_row].outlineLevel = 3
                    for cell in ws[ws.max_row]:
                        cell.border = box
                        cell.font = Font(size=9, color="FF6B7280")

        ws.append([g["cluster"] or "", g["bond"], "", f"{g['bond'].title()} total",
                   *g["totals"]])
        for cell in ws[ws.max_row]:
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.font = Font(bold=True, color=GOLD, size=10)
            cell.border = box

    # The grand total rounds once from the unrounded figures, never from the
    # bond totals - adding fifteen rounded numbers drifts.
    ws.append(["", "", "", "GRAND TOTAL", *data["total"]])
    for cell in ws[ws.max_row]:
        cell.fill = PatternFill("solid", fgColor=GOLD)
        cell.font = Font(bold=True, color=NAVY, size=11)
        cell.border = box

    for row in ws.iter_rows(min_row=4, min_col=5):
        for cell in row:
            cell.alignment = Alignment(horizontal="right")
            cell.number_format = "#,##0.00"
    for row in ws.iter_rows(min_row=4, min_col=1, max_col=3):
        for cell in row:
            cell.alignment = Alignment(horizontal="center")

    widths = {"A": 9, "B": 18, "C": 12, "D": 36, "E": 13, "F": 13, "G": 13, "H": 13}
    for col, wide in widths.items():
        ws.column_dimensions[col].width = wide
    ws.freeze_panes = "E4"
    ws.auto_filter.ref = f"A3:{last_col}{ws.max_row - 1}"
    ws.sheet_properties.outlinePr.summaryBelow = True
    ws.sheet_view.showGridLines = False

    tmp = Path(tempfile.mkdtemp()) / f"Shop Sales Cumulative - {scope} ({chosen['short']}).xlsx"
    wb.save(tmp)
    return FileResponse(tmp, filename=tmp.name)


@app.get("/reports/shop-cumulative/export.pdf")
def shop_cumulative_pdf(request: Request, date_from: str = "", date_to: str = "",
                        period: str = "", bond: str = "", cluster: int = 0,
                        warehouse: str = "", group_by: str = "bond",
                        scope: str = ""):
    """Two shapes from one endpoint.

    scope="current"  the screen, printed - the bonds you filtered to with their
                     shops under them, as one document.
    anything else    the office's own books: one PDF per bond, one page per
                     shop, zipped when more than one bond is asked for.
    """
    require_user(request)
    win = _window(date_from, date_to, period)
    if "error" in win:
        raise HTTPException(404, win["error"])
    chosen = win["period"]

    import json, subprocess, tempfile, zipfile
    from .pipelines import APP_REPORTS

    outdir = Path(tempfile.mkdtemp())

    if scope == "current":
        label = (bond.title() if bond else warehouse.title() if warehouse
                 else f"Cluster {cluster}" if cluster else "All bonds")
        data = _window_data(win, cluster, bond, warehouse, group_by)
        if "error" in data or not data.get("bonds"):
            raise HTTPException(404, data.get("error", "Nothing matches that filter."))
        out = outdir / f"Shop Sales Cumulative - {label} ({chosen['short']}).pdf"
        reports_pdf.build_cumulative_view_pdf(data, out, scope=label)
        return FileResponse(out, filename=out.name)

    argv = ["python3", str(APP_REPORTS / "build_shop_cumulative_pdfs.py"),
            "--outdir", str(outdir), "--period", chosen["long"]]
    if win["source"] == "cumulative":
        argv += ["--raw", str(chosen["path"])]
    else:
        # A window stitched from several files has no file of its own, so the
        # shop lines go to the builder as data rather than as a path.
        feed = outdir / "_window.json"
        feed.write_text(json.dumps(reports_api.shop_range_lines(
            chosen["start"], chosen["end"])))
        argv += ["--lines", str(feed)]

    proc = subprocess.run(argv, capture_output=True, text=True,
                          timeout=config.STEP_TIMEOUT_SECONDS)
    if proc.returncode != 0:
        raise HTTPException(500, f"PDF build failed: {(proc.stderr or proc.stdout)[-400:]}")

    pdfs = sorted(outdir.glob("*.pdf"))
    if not pdfs:
        raise HTTPException(404, "Nothing in that period mapped to a bond.")

    if bond and bond.lower() != "all":
        want = f"Shop Sales Cumulative - {bond.title()}.pdf"
        for f in pdfs:
            if f.name.lower() == want.lower():
                return FileResponse(f, filename=f.name)
        raise HTTPException(404, f"No pages for {bond} in that period.")

    if cluster:
        keep = {b.title().lower() for b, c in reports_api.cluster_of_bond().items()
                if c == cluster}
        pdfs = [f for f in pdfs if f.stem.split(" - ", 1)[-1].lower() in keep]
        if not pdfs:
            raise HTTPException(404, f"No bonds in cluster {cluster} for that period.")

    stamp = f"{chosen['start'].isoformat()} to {chosen['end'].isoformat()}"
    label = f"Cluster {cluster} " if cluster else ""
    bundle = outdir / f"Shop Sales Cumulative {label}({stamp}).zip"
    with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as z:
        for f in pdfs:
            z.write(f, f.name)
    return FileResponse(bundle, filename=bundle.name)


def _bond_totals(win: dict) -> dict:
    """bond -> [opening, receipt, sales, closing], whichever way it was sourced."""
    data = _window_data(win, 0, "")
    return {g["bond"]: g["totals"] for g in data.get("bonds", [])}


def _previous_window(period: dict):
    """The same day-window one month back, from a cumulative or the dailies.

    Matched on the day numbers rather than on elapsed days: 'the first half of
    last month' is the comparison the trade actually makes.
    """
    start, end = period["start"], period["end"]
    month = start.month - 1 or 12
    year = start.year - (1 if start.month == 1 else 0)
    try:
        back_from = date(year, month, start.day)
        back_to = date(year, month, end.day)
    except ValueError:
        return None
    prev = reports_api.resolve_window(back_from, back_to)
    return None if "error" in prev else prev


def _analysis_rows(win: dict, cluster: int, bond: str):
    """The rows the screen shows, the PDF prints and the workbook exports."""
    mod = _analysis_module()
    chosen = win["period"]
    prev = _previous_window(chosen)
    now = _bond_totals(win)
    before = _bond_totals(prev) if prev else {}
    raw = mod.build_rows(now, before, chosen["days"],
                         prev["period"]["days"] if prev else chosen["days"], cluster, bond)
    rows = [{
        "kind": r["kind"], "label": r["label"], "cells": r["cells"],
        "net_pct": mod.pct(r["net_pct"]), "sell": mod.pct(r["sell"]),
        "amber": r["sell"] is not None and r["sell"] >= mod.SELL_AMBER_AT,
        "cm": r["cm"], "lm": r["lm"],
        "up": r["trend"] >= 0, "trend": mod.whole(abs(r["trend"])),
    } for r in raw]
    return rows, (prev["period"] if prev else None), sorted(now)


@app.get("/reports/shop-analysis", response_class=HTMLResponse)
def shop_analysis_page(request: Request, date_from: str = "", date_to: str = "",
                       period: str = "", cluster: int = 0, bond: str = ""):
    user = require_user(request)
    periods = reports_api.cumulative_periods()
    win = _window(date_from, date_to, period)
    chosen = None if "error" in win else win["period"]

    rows, prev, bonds = [], None, []
    if chosen:
        rows, prev, bonds = _analysis_rows(win, cluster, bond)
    return templates.TemplateResponse(
        request, "report_shop_analysis.html",
        {"user": user, "page": "shop_analysis", "started_at": STARTED_AT,
         "problems": getattr(app.state, "problems", []),
         "periods": periods, "chosen": chosen, "prev": prev, "rows": rows,
         "bonds": bonds, "cluster": cluster, "bond": bond.upper(),
         "calendar": _calendar_days(),
         "source": win.get("source", ""), "chain": win.get("chain", []),
         "no_data": win.get("error", ""),
         "asked": _asked_period(date_from, date_to, period),
         "suggest": _suggested_window()},
    )


@app.get("/reports/shop-analysis/export.xlsx")
def shop_analysis_xlsx(request: Request, date_from: str = "", date_to: str = "",
                       period: str = "", cluster: int = 0, bond: str = ""):
    require_user(request)
    win = _window(date_from, date_to, period)
    if "error" in win:
        raise HTTPException(404, win["error"])
    chosen = win["period"]
    rows, _prev, _bonds = _analysis_rows(win, cluster, bond)
    if not rows:
        raise HTTPException(404, "Nothing matches that filter.")

    import tempfile
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    navy, gold = "FF0A294F", "FFFFBD30"
    wb = Workbook()
    ws = wb.active
    ws.title = "SHOPSALES COMPARATIVE"
    ws.append(["BOND", "OPENING", "RECEIPT", "SALES", "CLOSING", "STOCK NET",
               "STOCK NET %", "SELL-THROUGH %", "AVG/DAY CM", "AVG/DAY LM", "TREND"])
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(bold=True, color=gold, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    for r in rows:
        ws.append([r["label"], *r["cells"], r["net_pct"], r["sell"],
                   r["cm"], r["lm"], ("+" if r["up"] else "-") + str(r["trend"])])
        if r["kind"] != "bond":
            for cell in ws[ws.max_row]:
                cell.fill = PatternFill("solid", fgColor=navy)
                cell.font = Font(bold=True, color=gold, size=10)
    for row in ws.iter_rows(min_row=2, min_col=2):
        for cell in row:
            cell.alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 22
    for col in "BCDEFGHIJK":
        ws.column_dimensions[col].width = 13
    ws.freeze_panes = "B2"

    tmp = Path(tempfile.mkdtemp()) / f"Shop Sales Analysis ({chosen['short']}).xlsx"
    wb.save(tmp)
    return FileResponse(tmp, filename=tmp.name)


@app.get("/reports/shop-analysis/export.pdf")
def shop_analysis_pdf(request: Request, date_from: str = "", date_to: str = "",
                      period: str = "", cluster: int = 0, bond: str = ""):
    require_user(request)
    win = _window(date_from, date_to, period)
    if "error" in win:
        raise HTTPException(404, win["error"])
    chosen = win["period"]
    prev = _previous_window(chosen)

    import json, subprocess, tempfile
    from .pipelines import APP_REPORTS

    outdir = Path(tempfile.mkdtemp())
    # Totals rather than a path: the window may have been rebuilt from a dozen
    # daily files, and either way the builder should draw the same figures the
    # screen just showed instead of re-deriving them from a different read.
    totals = outdir / "_totals.json"
    totals.write_text(json.dumps(_bond_totals(win)))
    argv = ["python3", str(APP_REPORTS / "build_shop_analysis_pdf.py"),
            "--totals", str(totals), "--outdir", str(outdir),
            "--period", chosen["short"], "--days", str(chosen["days"])]
    if prev:
        before = outdir / "_prev.json"
        before.write_text(json.dumps(_bond_totals(prev)))
        argv += ["--prev-totals", str(before), "--prev-days", str(prev["period"]["days"])]
    if cluster:
        argv += ["--cluster", str(cluster)]
    if bond:
        argv += ["--bond", bond]

    proc = subprocess.run(argv, capture_output=True, text=True,
                          timeout=config.STEP_TIMEOUT_SECONDS)
    if proc.returncode != 0:
        raise HTTPException(500, f"PDF build failed: {(proc.stderr or proc.stdout)[-400:]}")

    out = outdir / "Shop Sales Analysis.pdf"
    if not out.is_file():
        raise HTTPException(404, "Nothing in that period mapped to a bond.")
    name = f"Shop Sales Analysis ({chosen['short']}).pdf"
    return FileResponse(out, filename=name)


# ---------------------------------------------------------------------------
# Liquidation Summary
# ---------------------------------------------------------------------------


def _liquidation(date_from: str, date_to: str, period: str = ""):
    """The scorecard for a window, plus the labels every export needs."""
    win = _window(date_from, date_to, period)
    if "error" in win:
        return {"error": win["error"]}
    cur = win["period"]
    data = reports_api.liquidation(cur["start"], cur["end"])
    if "error" in data:
        return data
    prev = data["previous"]
    data["headings"] = [cur["start"].strftime("%b").upper(),
                        prev["start"].strftime("%b").upper() if prev else "PREV", "CS", "%"]
    data["subtitle"] = (
        f'{cur["start"].strftime("%B").upper()} vs '
        f'{prev["start"].strftime("%B").upper() if prev else "LAST MONTH"} {cur["start"].year}'
        f'  \u00b7  DAYS {cur["start"].day}-{cur["end"].day}  \u00b7  cases')
    return data


@app.get("/reports/liquidation", response_class=HTMLResponse)
def liquidation_page(request: Request, date_from: str = "", date_to: str = "", period: str = ""):
    user = require_user(request)
    data = _liquidation(date_from, date_to, period)
    return templates.TemplateResponse(
        request, "report_liquidation.html",
        {"user": user, "page": "liquidation", "started_at": STARTED_AT,
         "problems": getattr(app.state, "problems", []),
         "calendar": _calendar_days(),
         "groups": ["Shop liquidation (KSBC)", "Secondary sales",
                    "Fed / Bar invoice", "Total liquidation"],
         "data": data, "error": data.get("error", "")},
    )


@app.get("/reports/liquidation/export.pdf")
def liquidation_pdf(request: Request, date_from: str = "", date_to: str = "", period: str = ""):
    require_user(request)
    data = _liquidation(date_from, date_to, period)
    if "error" in data:
        raise HTTPException(404, data["error"])

    import json, subprocess, tempfile
    from .pipelines import APP_REPORTS

    outdir = Path(tempfile.mkdtemp())
    feed = outdir / "scorecard.json"
    feed.write_text(json.dumps({"rows": data["rows"], "headings": data["headings"]}, default=str))
    out = outdir / f"Liquidation Summary ({data['period']['short']}).pdf"
    proc = subprocess.run(
        ["python3", str(APP_REPORTS / "build_liquidation_pdf.py"),
         "--data", str(feed), "--out", str(out), "--title", data["subtitle"]],
        capture_output=True, text=True, timeout=config.STEP_TIMEOUT_SECONDS)
    if proc.returncode != 0:
        raise HTTPException(500, f"PDF build failed: {(proc.stderr or proc.stdout)[-400:]}")
    return FileResponse(out, filename=out.name)


@app.get("/reports/liquidation/export.xlsx")
def liquidation_xlsx(request: Request, date_from: str = "", date_to: str = "", period: str = ""):
    require_user(request)
    data = _liquidation(date_from, date_to, period)
    if "error" in data:
        raise HTTPException(404, data["error"])

    import tempfile
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    navy, gold = "FF0A294F", "FFFFBD30"
    now_label, was_label = data["headings"][0], data["headings"][1]
    wb = Workbook()
    ws = wb.active
    ws.title = "LIQUIDATION"
    head = ["BOND"]
    for g in ("SHOP LIQUIDATION", "SECONDARY SALES", "FED / BAR INVOICE", "TOTAL LIQUIDATION"):
        head += [f"{g} {now_label}", f"{g} {was_label}", f"{g} +/- CS", f"{g} +/- %"]
    ws.append(head)
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(bold=True, color=gold, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    for row in data["rows"]:
        line = [row["label"]]
        for now, was in row["blocks"]:
            delta = now - was
            pct = ((now - was) / was * 100) if was else None
            line += [round(now), round(was), round(delta),
                     (round(pct) / 100 if pct is not None else None)]
        ws.append(line)
        if row["kind"] != "bond":
            for cell in ws[ws.max_row]:
                cell.fill = PatternFill("solid", fgColor=navy)
                cell.font = Font(bold=True, color=gold, size=10)

    for r in ws.iter_rows(min_row=2, min_col=2):
        for cell in r:
            cell.alignment = Alignment(horizontal="center")
    for col in range(5, ws.max_column + 1, 4):
        for r in ws.iter_rows(min_row=2, min_col=col, max_col=col):
            r[0].number_format = "0%"

    ws.column_dimensions["A"].width = 22
    for i in range(2, ws.max_column + 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = 13
    ws.freeze_panes = "B2"

    tmp = Path(tempfile.mkdtemp()) / f"Liquidation Summary ({data['period']['short']}).xlsx"
    wb.save(tmp)
    return FileResponse(tmp, filename=tmp.name)


@app.get("/reports/warehouse-stock/export.pdf")
def stock_pdf(request: Request, scope: str = "cluster", cluster: str = "1",
              as_of: str = "", warehouse: str = ""):
    """scope='cluster' -> one cluster; scope='current' -> whatever is filtered."""
    require_user(request)
    import tempfile

    if scope == "current":
        cl = int(cluster) if cluster in ("1", "2", "3") else None
        data = reports_api.warehouse_stock(as_of=as_of, cluster=cl, warehouse=warehouse)
        label = warehouse or (f"Cluster {cl}" if cl else "All Warehouses")
    else:
        if cluster not in ("1", "2", "3"):
            raise HTTPException(400, "Cluster must be 1, 2 or 3.")
        data = reports_api.warehouse_stock(as_of=as_of, cluster=int(cluster))
        label = f"Cluster {cluster}"

    if "error" in data:
        raise HTTPException(404, data["error"])
    if not data["warehouses"]:
        raise HTTPException(404, f"No stock rows for {label} on that date.")

    out = Path(tempfile.mkdtemp()) / f"Warehouse Stock Report - {label}.pdf"
    reports_pdf.build_stock_pdf(data, out)
    return FileResponse(out, filename=out.name)


@app.get("/reports/warehouse-stock/export.xlsx")
def stock_xlsx(request: Request, as_of: str = "", cluster: str = "", warehouse: str = ""):
    require_user(request)
    import tempfile
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    cl = int(cluster) if cluster in ("1", "2", "3") else None
    data = reports_api.warehouse_stock(as_of=as_of, cluster=cl, warehouse=warehouse)
    if "error" in data:
        raise HTTPException(404, data["error"])

    navy, gold, white = "FF0A294F", "FFFFBD30", "FFFFFFFF"
    wb = Workbook()
    ws = wb.active
    ws.title = "WAREHOUSE STOCK"
    ws.append(["WAREHOUSE", "ITEM NAME", "PACK", "PHYSICAL", "ALLOTABLE", "PENDING"])
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(bold=True, color=gold, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    for wh in data["warehouses"]:
        for r in wh["rows"]:
            ws.append([wh["name"], r["brand"], r["pack"],
                       r["physical"], r["allotable"], r["pending"]])
        ws.append([wh["name"], "TOTAL", "", wh["totals"]["physical"],
                   wh["totals"]["allotable"], wh["totals"]["pending"]])
        for cell in ws[ws.max_row]:
            cell.fill = PatternFill("solid", fgColor=navy)
            cell.font = Font(bold=True, color=gold, size=10)

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 38
    for col in ("C", "D", "E", "F"):
        ws.column_dimensions[col].width = 13
    ws.freeze_panes = "A2"

    label = warehouse or (f"Cluster {cl}" if cl else "All Warehouses")
    tmp = Path(tempfile.mkdtemp()) / f"Warehouse Stock Report - {label}.xlsx"
    wb.save(tmp)
    return FileResponse(tmp, filename=tmp.name)


# ---------------------------------------------------------------------------
# Daily grids: Shop Sales - Daily and Secondary Sales - Daily
#
# One screen and one PDF shape serve both — the reports differ only in where
# the cases come from, so the route takes the kind and everything else is
# shared. Keeping them apart would mean maintaining the same grid twice.
# ---------------------------------------------------------------------------

_DAILY_PAGES = {"shop": "shop_daily", "secondary": "secondary_daily"}


def _daily_kind(path_kind: str) -> str:
    if path_kind not in reports_api.DAILY_KINDS:
        raise HTTPException(404, "No such report.")
    return path_kind


@app.get("/reports/daily/{kind}", response_class=HTMLResponse)
def daily_page(request: Request, kind: str):
    user = current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    _daily_kind(kind)
    return templates.TemplateResponse(
        request, "report_daily.html",
        {"user": user, "page": _DAILY_PAGES[kind], "kind": kind,
         "title": reports_api.DAILY_KINDS[kind]["title"],
         "started_at": STARTED_AT,
         "problems": getattr(app.state, "problems", [])},
    )


@app.get("/api/reports/daily/{kind}")
def daily_data(request: Request, kind: str, date_from: str = "", date_to: str = "",
               cluster: str = "", view: str = "bond"):
    require_user(request)
    _daily_kind(kind)
    if view not in ("bond", "warehouse"):
        raise HTTPException(400, "Unknown view.")
    cl = int(cluster) if cluster in ("1", "2", "3") else None
    return JSONResponse(reports_api.daily_grid(kind, date_from=date_from,
                                               date_to=date_to, cluster=cl, view=view))


@app.get("/reports/daily/{kind}/export.pdf")
def daily_pdf(request: Request, kind: str, date_from: str = "", date_to: str = "",
              cluster: str = "", view: str = "bond"):
    require_user(request)
    _daily_kind(kind)
    import tempfile

    cl = int(cluster) if cluster in ("1", "2", "3") else None
    data = reports_api.daily_grid(kind, date_from=date_from, date_to=date_to,
                                  cluster=cl, view=view)
    if "error" in data:
        raise HTTPException(404, data["error"])

    suffix = f" (Cluster {cl})" if cl else ""
    out = Path(tempfile.mkdtemp()) / f"{data['title']}{suffix}.pdf"
    reports_pdf.build_daily_pdf(data, out)
    return FileResponse(out, filename=out.name)


@app.get("/reports/daily/{kind}/export.xlsx")
def daily_xlsx(request: Request, kind: str, date_from: str = "", date_to: str = "",
               cluster: str = "", view: str = "bond"):
    require_user(request)
    _daily_kind(kind)
    import tempfile
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    cl = int(cluster) if cluster in ("1", "2", "3") else None
    data = reports_api.daily_grid(kind, date_from=date_from, date_to=date_to,
                                  cluster=cl, view=view)
    if "error" in data:
        raise HTTPException(404, data["error"])

    navy, gold, white = "FF0A294F", "FFFFBD30", "FFFFFFFF"
    wb = Workbook()
    ws = wb.active
    ws.title = data["title"][:31]

    ws.append([data.get("group_label", "Bond").upper()]
              + [f"{d['dow']} {d['dom']}" for d in data["days"]] + ["TOTAL"])
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.font = Font(bold=True, color=gold, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    for r in data["rows"]:
        ws.append([r["label"]] + r["cells"] + [r["total"]])
        row = ws[ws.max_row]
        for cell in row:
            cell.alignment = Alignment(horizontal="center")
            if r["kind"] == "cluster":
                cell.fill = PatternFill("solid", fgColor=navy)
                cell.font = Font(bold=True, color=gold)
            elif r["kind"] == "grand":
                cell.fill = PatternFill("solid", fgColor=gold)
                cell.font = Font(bold=True, color=navy)
        row[0].alignment = Alignment(horizontal="left")
        row[-1].font = Font(bold=True, color=(gold if r["kind"] == "cluster"
                                              else navy if r["kind"] == "grand" else "FF0F192D"))

    ws.freeze_panes = "B2"
    ws.column_dimensions["A"].width = 22
    for i in range(2, len(data["days"]) + 3):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = 7

    out = Path(tempfile.mkdtemp()) / f"{data['title']}.xlsx"
    wb.save(out)
    return FileResponse(out, filename=out.name)


# ---------------------------------------------------------------------------
# Settings: bond mapping
# ---------------------------------------------------------------------------


@app.get("/settings/bond-mapping", response_class=HTMLResponse)
def bond_mapping_page(request: Request):
    user = current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    return templates.TemplateResponse(
        request, "settings_bonds.html",
        {"user": user, "page": "bonds", "started_at": STARTED_AT,
         "problems": getattr(app.state, "problems", [])},
    )


@app.get("/api/bond-mapping")
def bond_mapping_data(request: Request):
    require_user(request)
    return JSONResponse(bondmap.load_shops())


@app.post("/api/bond-mapping")
async def bond_mapping_save(request: Request):
    """Save shop -> bond edits, the cluster split, or both."""
    require_user(request)
    body = await request.json()

    result: dict = {}
    clusters = body.get("clusters")
    if isinstance(clusters, dict):
        bondmap.save_clusters(clusters)
        result["clusters_saved"] = True

    changes = body.get("bonds") or {}
    if changes:
        written = bondmap.save_bonds(changes)
        if "error" in written:
            raise HTTPException(400, written["error"])
        result.update(written)
        # The reports cache master data on its mtime, so the next read picks
        # the new mapping up on its own.

    return JSONResponse({"ok": True, **result})


# ---------------------------------------------------------------------------
# Target vs achievement
# ---------------------------------------------------------------------------
#
# The targets are typed here rather than uploaded, because they are not KSBC
# data: nobody exports them, they are a decision taken in a meeting and revised
# mid-month. The achievement beside them is total liquidation - shop sales plus
# the FED and BAR invoices - which reports_api assembles from the raws already
# in the workspace.


def _tva_window(date_from: str, date_to: str):
    """The window to report on: what was asked, or the current month so far."""
    asked = _asked_dates(date_from, date_to)
    if asked:
        return asked
    widest = reports_api.widest_window()
    if not widest:
        return None
    start, end = widest
    return (max(start, end.replace(day=1)), end)


@app.get("/reports/target-achievement", response_class=HTMLResponse)
def target_achievement_page(request: Request):
    user = current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    return templates.TemplateResponse(
        request, "report_target.html",
        {"user": user, "page": "target", "started_at": STARTED_AT,
         "problems": getattr(app.state, "problems", [])},
    )


@app.get("/api/target-achievement")
def target_achievement_data(request: Request, date_from: str = "", date_to: str = "",
                            cluster: int = 0, month: str = ""):
    require_user(request)
    window = _tva_window(date_from, date_to)
    if window is None:
        return JSONResponse({"error": "No shop sales have been uploaded yet, so there is "
                                      "nothing to measure a target against."})
    data = reports_api.target_vs_achievement(
        window[0], window[1], cluster=cluster or None, month=month)
    data.setdefault("bonds", sorted({b for members in reports_api.bond_clusters().values()
                                     for b in members}))
    data["calendar"] = _calendar_days()
    data["suggest"] = _suggested_window()
    return JSONResponse(data)


@app.get("/api/targets")
def targets_read(request: Request, month: str = ""):
    require_user(request)
    month = month or date.today().strftime("%Y-%m")
    try:
        cells = targets_mod.load(month)
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    return JSONResponse({
        "month": month,
        "bonds": cells,
        "families": [{"key": k, "label": l} for k, l in targets_mod.FAMILIES],
        "all_bonds": sorted({b for members in reports_api.bond_clusters().values()
                             for b in members}),
        "clusters": {str(c): sorted(m) for c, m in reports_api.bond_clusters().items()},
        "months": targets_mod.months(),
        "meta": targets_mod.meta(month),
    })


@app.post("/api/targets")
async def targets_write(request: Request):
    user = require_user(request)
    body = await request.json()
    month = str(body.get("month") or "")
    try:
        saved = targets_mod.save(month, body.get("bonds") or {}, by=user)
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    return JSONResponse({"ok": True, **saved})


@app.get("/reports/target-achievement/export.xlsx")
def target_achievement_xlsx(request: Request, date_from: str = "", date_to: str = "",
                            cluster: int = 0, month: str = ""):
    """The sheet as it prints, in the shape the office already circulates."""
    require_user(request)
    window = _tva_window(date_from, date_to)
    if window is None:
        raise HTTPException(404, "No shop sales have been uploaded yet.")
    data = reports_api.target_vs_achievement(
        window[0], window[1], cluster=cluster or None, month=month)
    if "error" in data:
        raise HTTPException(404, data["error"])

    import tempfile
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    # The same sheet as the PDF, cell for cell: the office should not be able
    # to tell which of the two it is looking at.
    NAVY, GOLD, GOLD_D = "FF0A294F", "FFFFBD30", "FFF2B22E"
    SHADE, INK, RED = "FFF2F2F2", "FF1B2A4A", "FFCF1322"
    LINE = Side(style="thin", color="FFC7C7C7")
    BOX = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)

    cols = data["columns"]
    labels = ["BOND", "CAT"] + [reports_pdf.TA_LABELS.get(c["key"], str(c["label"]).upper())
                                for c in cols] + ["GRAND TOTAL", "ACH %"]
    span = len(labels)
    last_col = get_column_letter(span)

    wb = Workbook()
    ws = wb.active
    ws.title = f"TGT vs ACH {data['month'][:4]}"[:31]
    ws.sheet_view.showGridLines = False

    def band(row, text, fill, colour, size, align="center", height=None):
        ws.cell(row=row, column=1, value=text)
        ws.merge_cells(f"A{row}:{last_col}{row}")
        for col in range(1, span + 1):
            c = ws.cell(row=row, column=col)
            c.fill = PatternFill("solid", fgColor=fill)
            c.font = Font(bold=True, size=size, color=colour)
            c.alignment = Alignment(horizontal=align, vertical="center")
        if height:
            ws.row_dimensions[row].height = height

    end_day = data["period"]["to"]
    as_on = (f"AS ON {int(end_day[8:10])} "
             f"{reports_pdf._MONTH_ABBR[int(end_day[5:7]) - 1]} {end_day[:4]}")
    title = "TARGET VS ACHIEVEMENT"
    if cluster in (1, 2, 3):
        title += f" - CLUSTER {cluster}"

    band(1, "K.S DISTILLERY", NAVY, GOLD, 18, height=32)
    band(2, title, GOLD, NAVY, 12, align="left", height=22)
    # the as-on date sits at the far end of the same gold band
    ws.unmerge_cells(f"A2:{last_col}2")
    ws.merge_cells(f"A2:{get_column_letter(span - 4)}2")
    ws.merge_cells(f"{get_column_letter(span - 3)}2:{last_col}2")
    right = ws.cell(row=2, column=span - 3, value=as_on)
    right.font = Font(bold=True, size=12, color=NAVY)
    right.alignment = Alignment(horizontal="right", vertical="center")

    ws.append([])
    head_row = 4
    for i, label in enumerate(labels, start=1):
        c = ws.cell(row=head_row, column=i, value=label)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.font = Font(bold=True, size=9.5, color="FFFFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[head_row].height = 30

    rows = data["rows"]
    has_bonds = any(r.get("kind") == "bond" for r in rows)
    r = head_row

    for idx, row in enumerate(rows):
        total = idx == len(rows) - 1
        strong = total or (has_bonds and row.get("kind") == "cluster")
        top, bottom = r + 1, r + 2
        for which, tag, rr, fill in (("tgt", "TGT", top, GOLD_D if total else SHADE),
                                     ("ach", "ACH", bottom, GOLD if total else "FFFFFFFF")):
            line = [row["label"] if rr == top else None, tag]
            line += [round(float(row[which].get(c["key"], 0) or 0)) for c in cols]
            line.append(round(float(row.get(which + "_total", 0) or 0)))
            line.append((row["pct"] / 100) if (rr == top and row["pct"] is not None) else None)
            for i, value in enumerate(line, start=1):
                c = ws.cell(row=rr, column=i, value=value)
                c.border = BOX
                c.alignment = Alignment(horizontal="center", vertical="center")
                if i == 1:
                    c.fill = PatternFill("solid", fgColor=NAVY)
                    c.font = Font(bold=True, size=10, color=GOLD if total else "FFFFFFFF")
                elif i == span:
                    c.fill = PatternFill("solid", fgColor=GOLD if total else "FFFFFFFF")
                    c.font = Font(bold=True, size=9.5, color=INK if total else RED)
                    c.number_format = "0.0%"
                else:
                    c.fill = PatternFill("solid", fgColor=fill)
                    c.font = Font(bold=strong or i in (2, span - 1), size=9.5, color=INK)

        ws.merge_cells(start_row=top, start_column=1, end_row=bottom, end_column=1)
        ws.merge_cells(start_row=top, start_column=span, end_row=bottom, end_column=span)
        ws.row_dimensions[top].height = 18
        ws.row_dimensions[bottom].height = 18
        r = bottom

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 7
    for i in range(3, span + 1):
        ws.column_dimensions[get_column_letter(i)].width = 13
    ws.freeze_panes = ws.cell(row=head_row + 1, column=3)

    label = data["period"]["short"].replace(" to ", " - ")
    tmp = Path(tempfile.mkdtemp()) / f"TARGET vs ACHIEVEMENT ({label}).xlsx"
    wb.save(tmp)
    return FileResponse(tmp, filename=tmp.name)


@app.get("/reports/target-achievement/export.pdf")
def target_achievement_pdf(request: Request, date_from: str = "", date_to: str = "",
                           cluster: int = 0, month: str = "", scope: str = "current"):
    """Two ways out: what is on screen, or the set the office circulates.

    'all' is four sheets - one per cluster and the summary that sits on top of
    them - zipped, because that is how they go out together. 'current' is the
    page as it stands, cluster filter and all.
    """
    require_user(request)
    if scope not in ("current", "all"):
        raise HTTPException(400, "Unknown scope.")
    window = _tva_window(date_from, date_to)
    if window is None:
        raise HTTPException(404, "No shop sales have been uploaded yet.")

    import tempfile, zipfile

    def grid(which: int | None):
        got = reports_api.target_vs_achievement(
            window[0], window[1], cluster=which, month=month)
        if "error" in got:
            raise HTTPException(404, got["error"])
        return got

    out = Path(tempfile.mkdtemp())
    span = None

    if scope == "current":
        data = grid(cluster or None)
        span = data["period"]["short"].replace(" to ", " - ")
        name = f"CLUSTER {cluster}" if cluster in (1, 2, 3) else ""
        stem = f"TARGET vs ACHIEVEMENT{' - ' + name if name else ''} ({span})"
        pdf = reports_pdf.build_target_pdf(data, out / f"{stem}.pdf", scope=name)
        return FileResponse(pdf, filename=pdf.name, media_type="application/pdf")

    whole = grid(None)
    span = whole["period"]["short"].replace(" to ", " - ")
    made: list[Path] = []
    for cid in (1, 2, 3):
        one = grid(cid)
        if not one["rows"]:
            continue
        made.append(reports_pdf.build_target_pdf(
            one, out / f"TARGET vs ACHIEVEMENT - CLUSTER {cid} ({span}).pdf",
            scope=f"CLUSTER {cid}"))

    summary = [r for r in whole["rows"] if r.get("kind") in ("cluster", "grand")]
    made.append(reports_pdf.build_target_pdf(
        whole, out / f"TARGET vs ACHIEVEMENT - CLUSTER SUMMARY ({span}).pdf",
        rows=summary, scope="CLUSTER SUMMARY"))

    bundle = out / f"TARGET vs ACHIEVEMENT ({span}).zip"
    with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as z:
        for pdf in made:
            z.write(pdf, pdf.name)
    return FileResponse(bundle, filename=bundle.name, media_type="application/zip")


# ---------------------------------------------------------------------------
# Purchase instruction
# ---------------------------------------------------------------------------


@app.get("/reports/pi-variance", response_class=HTMLResponse)
def pi_page(request: Request):
    user = current_user(request)
    if not user:
        return RedirectResponse("/login", status_code=303)
    return templates.TemplateResponse(
        request, "report_pi.html",
        {"user": user, "page": "pi", "started_at": STARTED_AT,
         "problems": getattr(app.state, "problems", [])},
    )


def _pi_months(month: str = "", prior: str = "") -> tuple[str, str]:
    """The month to show and the one to compare it with.

    Default to the newest stored month against the one before it, because the
    only question anyone opens this report with is 'what changed'.
    """
    have = [m["key"] for m in pi_mod.months()]
    if not have:
        return "", ""
    month = month if month in have else have[0]
    if prior == "-":
        return month, ""
    if prior in have and prior != month:
        return month, prior
    later = [m for m in have if m < month]
    return month, (later[0] if later else "")


@app.get("/api/pi-variance")
def pi_data(request: Request, month: str = "", prior: str = "",
            view: str = "bond", cluster: int = 0):
    require_user(request)
    if view not in ("bond", "warehouse"):
        raise HTTPException(400, "Unknown view.")
    cur, prev = _pi_months(month, prior)
    if not cur:
        return JSONResponse({"error": "No purchase instruction has been uploaded yet. "
                                      "Upload a month's files on the Raw Data Upload page.",
                             "months": []})
    return JSONResponse(pi_mod.variance(cur, view=view,
                                        cluster=cluster or None, prior=prev))


@app.get("/reports/pi-variance/export.xlsx")
def pi_xlsx(request: Request, month: str = "", prior: str = "",
            view: str = "bond", cluster: int = 0):
    """The workbook the office already circulates, cell for cell.

    Bond, then its shops, then the next bond; eight brands four columns wide;
    a zero prints as a dash. No cluster bands and no totals column - this is
    the shape people already read, and a better one they have to re-learn is
    not better.
    """
    require_user(request)
    cur, prev = _pi_months(month, prior)
    if not cur:
        raise HTTPException(404, "No purchase instruction has been uploaded yet.")
    data = pi_mod.variance(cur, view=view, cluster=cluster or None, prior=prev)
    if "error" in data:
        raise HTTPException(404, data["error"])

    import tempfile
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    NAVY, GOLD, INK, ROW, TOTAL, DASH = ("FF0B2C52", "FFFAAF19", "FF0F192D",
                                         "FFF5F7FC", "FF2C3540", "FF8C8C8C")
    thin, medium = Side(style="thin", color="FFE3E8F0"), Side(style="medium", color="FFC9D2E2")
    # Brands in the workbook's own order: alphabetical by the full KSBC name.
    brands = sorted(data["brands"], key=lambda b: b["full"])
    span = 1 + len(brands) * 4
    last = get_column_letter(span)

    wb = Workbook()
    ws = wb.active
    ws.title = "PI Variance"
    ws.sheet_view.showGridLines = False

    def band(row, text, fill, colour, size, align="center"):
        ws.cell(row=row, column=1, value=text)
        ws.merge_cells(f"A{row}:{last}{row}")
        for col in range(1, span + 1):
            c = ws.cell(row=row, column=col)
            c.fill = PatternFill("solid", fgColor=fill)
            c.font = Font(bold=True, size=size, color=colour)
            c.alignment = Alignment(horizontal=align, vertical="center")

    band(1, "K.S DISTILLERY", NAVY, GOLD, 18)
    band(2, f"PURCHASE INSTRUCTION \u00b7 1 {data['month_label'].upper()}", GOLD, NAVY, 12)
    ws.row_dimensions[1].height = 36
    ws.row_dimensions[2].height = 24

    ws.cell(row=3, column=1, value="Row Labels")
    ws.merge_cells("A3:A4")
    for i, b in enumerate(brands):
        col = 2 + i * 4
        ws.cell(row=3, column=col, value=b["full"])
        ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + 3)
        for j, m in enumerate(("L3MS", "RL", "RQ", "MQ")):
            ws.cell(row=4, column=col + j, value=m)
    for r in (3, 4):
        ws.row_dimensions[r].height = 20
        for col in range(1, span + 1):
            c = ws.cell(row=r, column=col)
            c.fill = PatternFill("solid", fgColor=NAVY)
            c.font = Font(bold=True, size=9.5, color=GOLD)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def line(label, cellset, fill, colour, bold, zero_colour=None):
        r = ws.max_row + 1
        ws.row_dimensions[r].height = 20
        ws.cell(row=r, column=1, value=label)
        for i, b in enumerate(brands):
            cell = cellset.get(b["code"], {})
            for j, m in enumerate(("l3ms", "rl", "rq", "mq")):
                v = cell.get(m, 0)
                ws.cell(row=r, column=2 + i * 4 + j, value=v if v else "-")
        for col in range(1, span + 1):
            c = ws.cell(row=r, column=col)
            c.fill = PatternFill("solid", fgColor=fill)
            blank = col > 1 and c.value == "-"
            c.font = Font(bold=bold, size=9.5,
                          color=(zero_colour if (blank and zero_colour) else colour))
            c.alignment = Alignment(horizontal="left" if col == 1 else "center",
                                    vertical="center")
            c.number_format = "#,##0"
            c.border = Border(bottom=thin, left=medium if (col - 2) % 4 == 0 or col == 1 else thin)
        return r

    for row in data["rows"]:
        if row["kind"] != "group":
            continue
        line(row["label"], row["cells"], NAVY, GOLD, True)
        for shop in (data["shops"].get(row["key"]) or []):
            line(shop["name"], shop["cells"], ROW, INK, False, zero_colour=DASH)

    grand = next((r for r in data["rows"] if r["kind"] == "grand"), None)
    if grand:
        r = line("Grand Total", grand["cells"], TOTAL, "FFFFFFFF", True)
        ws.cell(row=r, column=1).font = Font(bold=True, size=9.5, color=GOLD)

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 12
    for i in range(3, span + 1):
        ws.column_dimensions[get_column_letter(i)].width = 13
    ws.freeze_panes = "B5"

    # The blank-PI list rides along on its own sheet: it changes nothing about
    # the sheet above, and it is the one thing the figures cannot tell you.
    if data["blanks"]:
        b = wb.create_sheet("No instruction")
        b.append(["Bond", "Shop", f"{data['prior_label'] or 'Last month'} MQ"])
        for cell in b[1]:
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.font = Font(bold=True, color=GOLD, size=10)
        for x in data["blanks"]:
            b.append([x["bond"], x["name"], x["prior_mq"] or ""])
        b.column_dimensions["A"].width = 20
        b.column_dimensions["B"].width = 38
        b.column_dimensions["C"].width = 16

    tmp = Path(tempfile.mkdtemp()) / f"pi_variance_report_{data['month']}.xlsx"
    wb.save(tmp)
    return FileResponse(tmp, filename=tmp.name)
