# Deployment

## The rule that matters most

**Every account is created in K.S. Distillery's name, with credentials the
company holds.** Not a developer's personal account, not a vendor's. That
applies to the Git host, the app host, and anything added later.

The point of this rebuild is that nobody outside KSD can switch it off.

---

## 1. Accounts to create

| What | Purpose | Who owns it |
|---|---|---|
| GitHub (or GitLab) organisation | The source code | KSD, on a company email |
| Railway **or** Render | Runs the service | KSD, on a company email |

Use a company address that more than one person can reach — a shared
`reports@` style mailbox is better than an individual's, so access does not walk
out of the door with one person.

Turn on two-factor authentication and store the recovery codes wherever the
company keeps its other credentials.

---

## 2. Why it cannot be serverless

A single formatting step in the Shop Sales pipeline takes **26–68 seconds**, and
the full tail runs **3.7 minutes** (measured — see `VERIFICATION.md`). Serverless
function platforms cut execution off far below that.

So: a long-running container on Railway or Render. Not Vercel functions, not
Lambda.

## 3. Why it needs a persistent disk

These pipelines are **accumulators, not pure functions**:

- `build_secondary.py` seeds from the live workbook's `COMBINED DISPATCHES`
  sheet, then layers the new single-day raw on top. Raws are uploaded one day at
  a time and removed after ingest, so **the month's earlier days exist only
  inside that workbook**.
- `ksbc_daily_update.py` appends a day sheet and refuses a gap.
- `build_warehouse_stock.py` appends to `_history/stock_history.csv`, which is
  the canonical trend data.

On an ephemeral container filesystem, every redeploy would silently reset the
month. **Mount a persistent volume at `/data`.** This is not optional and it is
the single easiest thing to get wrong.

Budget ~2 GB to start. Current usage is around 60 MB of workbooks plus history,
and it grows by roughly 10–15 MB a month.

---

## 4. Environment variables

| Variable | Value | Notes |
|---|---|---|
| `KSD_WORKSPACE_ROOT` | `/data/workspace` | Must be on the mounted volume |
| `KSD_SESSION_SECRET` | 48 random bytes | `python3 -c "import secrets; print(secrets.token_urlsafe(48))"` |
| `KSD_ALLOWED_EMAILS` | `abhay@…,neelima@…` | Comma separated. v1 is these two |
| `KSD_ALLOWED_EMAIL_DOMAINS` | your company domain | Optional; widens access to a whole domain |
| `KSD_STEP_TIMEOUT` | `900` | Seconds per step. Leave generous |

Never set `KSD_COOKIE_SECURE=0` outside local testing — it would send the
session cookie in the clear.

---

## 5. Seeding the workspace

The service expects the exact folder layout the build scripts resolve by
`__file__`:

```
/data/workspace/
  mnt/Claude/
    MASTER DATA CONFIRMED.xlsx
    KSBC shop sales/          <- current + prior month workbooks
    Secondary sales/          <- current + prior month workbooks
    Warehouse stock/
      _history/               <- stock_history.csv, inbound_history.csv,
                                 brand_pack_history.csv
    .claude/
      scripts/                <- the build scripts, copied verbatim
      memory/
```

Copy these from the existing Claude folder. Two things people get wrong:

1. **The prior month's workbooks are required, not optional.** The BOND INSIGHTS
   pass builds a current-vs-prior daily trend, and the Secondary build reads the
   prior month for comparisons. Copy at least the current and previous month for
   each stream.
2. **`_history/` must come across intact.** Those CSVs are years of trend data
   and nothing regenerates them.

Then create the logins:

```bash
python3 scripts/bootstrap_user.py abhay@<company-domain>
python3 scripts/bootstrap_user.py neelima@<company-domain>
```

---

## 6. Deploy

Both hosts read the `Dockerfile`.

**Railway:** New Project → Deploy from GitHub repo → add a Volume mounted at
`/data` → set the variables above → deploy.

**Render:** New → Web Service → connect the repo → Docker runtime → add a Disk
mounted at `/data` → set the variables → deploy.

Set the health check path to `/healthz`. It returns 503 with a list of specific
problems when the service is misconfigured, so a bad deploy announces what is
wrong instead of just failing.

Running cost is roughly $7–15/month plus the disk, depending on host and plan.
Check current pricing — it moves.

---

## 7. Backups

The volume holds months of accumulated work. The host's own snapshots are the
first line, but take your own copy too:

- The workbooks already live in the Claude folder on the Mac, which syncs to
  Google Drive. Keep that habit — download the promoted workbook periodically,
  or keep running the local pipeline in parallel for the first few weeks.
- Every promotion in this service backs up the file it replaced into the job
  directory under `/data/workspace/_jobs/<job id>/replaced/`, so an overwrite is
  recoverable without going to a host snapshot.

---

## 8. Handover checklist

Before considering this done, confirm someone other than the person who set it
up can do all of these:

- [ ] Sign in to the GitHub account and see the repository
- [ ] Sign in to the Railway/Render account and see the service
- [ ] Find the environment variables and the volume
- [ ] Trigger a redeploy
- [ ] Download a backup of `/data`
- [ ] Add a new user with `bootstrap_user.py`

If any of those depends on one person, the key-person risk has moved rather
than gone away.
