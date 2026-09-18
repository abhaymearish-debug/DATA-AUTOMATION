#!/bin/bash
# ---------------------------------------------------------------------------
# Run the Report Transformer on this Mac, for testing.
#
#   bash "run_local.sh"
#
# It sets everything up in ~/ksd-report-app and works on a COPY of the Claude
# folder, so nothing in your live workbooks can be touched while you try it.
# Stop it any time with Ctrl-C. Delete ~/ksd-report-app to remove it entirely.
# ---------------------------------------------------------------------------

set -e

APP_HOME="${KSD_APP_HOME:-$HOME/ksd-report-app}"
# Override with KSD_CLAUDE_FOLDER if your Claude folder lives somewhere else.
CLAUDE_SRC="${KSD_CLAUDE_FOLDER:-$HOME/Downloads/Claude}"
SRC_CODE="$CLAUDE_SRC/Report Transformer App"
WS="$APP_HOME/workspace/mnt/Claude"
# Reference material the app keeps but never reports on. Must match
# WORKSPACE_ROOT/_seed in app/main.py.
SEED="$APP_HOME/workspace/_seed"
PORT=8000

bold() { printf "\033[1m%s\033[0m\n" "$1"; }
ok()   { printf "  \033[32m✓\033[0m %s\n" "$1"; }
warn() { printf "  \033[33m!\033[0m %s\n" "$1"; }

echo
bold "K.S. Distillery — Report Transformer (local test)"
echo

# --- 1. Python -------------------------------------------------------------
# 3.9 is enough. Every module in this app and all 81 build scripts were checked
# against 3.9 syntax and none of them use anything newer. A more recent
# interpreter is preferred when one is installed, because the server runs 3.11.
PY=""
for candidate in python3.13 python3.12 python3.11 python3.10 python3; do
  if command -v "$candidate" >/dev/null 2>&1 \
     && "$candidate" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3,9) else 1)' 2>/dev/null; then
    PY="$candidate"; break
  fi
done
if [ -z "$PY" ]; then
  echo "No usable Python found (3.9 or newer is needed)."
  echo "Run this, let it finish, then run this script again:"
  echo "    xcode-select --install"
  exit 1
fi
ok "$PY $("$PY" -c 'import sys; print("%d.%d.%d" % sys.version_info[:3])')"

# --- 2. Source code --------------------------------------------------------
if [ ! -d "$SRC_CODE" ]; then
  echo "Could not find the app at:"
  echo "    $SRC_CODE"
  exit 1
fi
ok "found the app source"

# --- 3. Install ------------------------------------------------------------
# The app is RUN FROM THE SOURCE FOLDER, not from a copy of it. Only the
# virtualenv, the config and the data workspace live in $APP_HOME. That means an
# edit to a template or its CSS/JS is live on the next browser refresh, because
# Jinja re-reads templates from disk on every render. Python changes still need
# a restart — the modules are already in memory. (KSD_RELOAD=1 restarts the
# server automatically on a .py change; it will kill a build that is running.)
bold "Installing (first run takes a minute or two)"
mkdir -p "$APP_HOME"

if [ ! -d "$APP_HOME/.venv" ]; then
  "$PY" -m venv "$APP_HOME/.venv"
fi
# shellcheck disable=SC1091
source "$APP_HOME/.venv/bin/activate"

pip install --quiet --upgrade pip
# Plain uvicorn rather than uvicorn[standard]: the extras compile C code and
# that is a needless way for a first run to fail.
pip install --quiet "openpyxl==3.1.5" fastapi uvicorn python-multipart itsdangerous jinja2 "reportlab>=4.0"
[ "${KSD_RELOAD:-0}" = "1" ] && pip install --quiet watchfiles
ok "dependencies installed"

# --- 4. Workspace ----------------------------------------------------------
# NOTHING but reference data is seeded. The reports read the workspace, so
# anything seeded here would appear as a figure the operator never uploaded —
# which is exactly the confusion this avoids. Sales workbooks and warehouse
# history are built by the app, from raws uploaded through it.
#
# Shop sales is the one exception that needs explaining. Its locked driver
# (ksbc_daily_update.py) only ever EXTENDS a month's workbook; it cannot create
# one, and ksbc_bootstrap_month.py can only template a new month from an
# existing workbook's structure. So one structural TEMPLATE is built here —
# every data sheet stripped out, no sales figures in it — and parked outside the
# report folders. The app clones it whenever a new month is first uploaded.
# The workspace used to be seeded with copies of the live workbooks, so reports
# showed months nobody had uploaded. Clear that once, on the first run under the
# new policy, or those old copies would linger and keep doing it.
POLICY="$SEED/.policy"
if [ -d "$APP_HOME/workspace" ] && [ "$(cat "$POLICY" 2>/dev/null)" != "uploads-only" ]; then
  warn "clearing seeded workbooks — from now on the reports show only what you upload"
  rm -rf "$WS/KSBC shop sales" "$WS/Secondary sales" "$WS/Warehouse stock" "$SEED"
fi

bold "Preparing the workspace"
mkdir -p "$WS/.claude" "$WS/KSBC shop sales" "$WS/KSBC shop sales/_cumulative" \
         "$WS/Secondary sales" "$WS/Warehouse stock/_history" "$SEED"

rm -rf "$WS/.claude/scripts"
cp -Rp "$CLAUDE_SRC/.claude/scripts" "$WS/.claude/scripts"
[ -d "$CLAUDE_SRC/.claude/memory" ] && cp -Rp "$CLAUDE_SRC/.claude/memory" "$WS/.claude/memory"
cp -p "$CLAUDE_SRC/MASTER DATA CONFIRMED.xlsx" "$WS/"
ok "scripts and master data (reference only — no sales figures)"

TEMPLATE="$SEED/SHOP SALES TEMPLATE.xlsx"
if [ ! -f "$TEMPLATE" ]; then
  SRC_WB=""
  while IFS= read -r nm; do
    SRC_WB="$CLAUDE_SRC/KSBC shop sales/$nm"; break
  done < <(ls -t "$CLAUDE_SRC/KSBC shop sales" 2>/dev/null | grep -iE "SHOP SALES ANALYSIS\.xlsx$")

  if [ -n "$SRC_WB" ] && [ -f "$SRC_WB" ]; then
    BOOT_LOG="$APP_HOME/bootstrap_template.log"
    python3 "$WS/.claude/scripts/ksbc_bootstrap_month.py" \
      --month JANUARY --from "$SRC_WB" --out "$TEMPLATE" \
      --folder "$SEED" > "$BOOT_LOG" 2>&1
    # Judge it on the file, not the exit code: the script exits 0 on a SKIP.
    if [ -f "$TEMPLATE" ]; then
      ok "shop-sales structure template built (data sheets stripped)"
    else
      warn "could not build the shop-sales template — see $BOOT_LOG"
      tail -3 "$BOOT_LOG" | sed 's/^/      /'
    fi
  else
    warn "no shop-sales workbook found to take structure from"
  fi
else
  ok "shop-sales structure template already in place"
fi
printf 'uploads-only' > "$POLICY"

# --- 5. Config -------------------------------------------------------------
ENV_FILE="$APP_HOME/.env"
if [ ! -f "$ENV_FILE" ]; then
  SECRET=$(python3 -c 'import secrets; print(secrets.token_urlsafe(48))')
  cat > "$ENV_FILE" <<EOF
KSD_WORKSPACE_ROOT=$APP_HOME/workspace
KSD_SESSION_SECRET=$SECRET
KSD_ALLOWED_EMAILS=abhaymearish@gmail.com
KSD_COOKIE_SECURE=0
EOF
  chmod 600 "$ENV_FILE"
fi
set -a; # shellcheck disable=SC1090
source "$ENV_FILE"; set +a
ok "configuration ready"

# --- 6. Login --------------------------------------------------------------
cd "$SRC_CODE"
if [ ! -f "$APP_HOME/workspace/_auth/users.json" ]; then
  echo
  bold "Set a password for $KSD_ALLOWED_EMAILS"
  echo "(at least 12 characters — you will type it each time you sign in)"
  python3 scripts/bootstrap_user.py "$KSD_ALLOWED_EMAILS"
fi

# --- 7. Stop any previous instance ------------------------------------------
# Without this, a second run fails to bind the port, exits, and leaves the OLD
# server answering — so the browser shows stale code and nothing looks wrong.
OLD=$(lsof -ti tcp:$PORT 2>/dev/null || true)
if [ -n "$OLD" ]; then
  warn "an instance is already running on port $PORT — stopping it"
  kill $OLD 2>/dev/null || true
  sleep 2
  STILL=$(lsof -ti tcp:$PORT 2>/dev/null || true)
  [ -n "$STILL" ] && kill -9 $STILL 2>/dev/null || true
  sleep 1
fi

# --- 8. Go -----------------------------------------------------------------
echo
bold "Starting — open this in your browser:"
printf "\n      \033[1;36mhttp://localhost:%s\033[0m\n\n" "$PORT"
echo "  Sign in with $KSD_ALLOWED_EMAILS and the password you just set."
echo "  Screen changes (layout, colours, buttons): just refresh the browser."
echo "  Changes to the Python behind them need a Ctrl-C and a re-run."
echo "  Press Ctrl-C here to stop it."
echo

( sleep 3; open "http://localhost:$PORT" >/dev/null 2>&1 || true ) &
RELOAD=""
if [ "${KSD_RELOAD:-0}" = "1" ]; then
  RELOAD="--reload --reload-dir app --reload-dir reports"
  warn "auto-reload on — a code change mid-build will abandon that build"
fi
# shellcheck disable=SC2086
exec python3 -m uvicorn app.main:app --host 127.0.0.1 --port "$PORT" $RELOAD
