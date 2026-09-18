#!/bin/bash
# Double-click this to send whatever Claude changed up to GitHub.
# Render sees the push and redeploys on its own, about five minutes.
cd "$(dirname "$0")" || exit 1

printf '\n  K.S. Distillery — Report Transformer\n'
printf '  Sending changes to Render...\n\n'

if [ -n "$(git status --porcelain)" ]; then
  printf '  Some files were changed but not committed:\n'
  git status --short | sed 's/^/    /'
  printf '\n  Committing them as "Manual changes"...\n\n'
  git add -A
  git commit -q -m "Manual changes"
fi

if git push; then
  printf '\n  Done. Render is rebuilding — give it about five minutes,\n'
  printf '  then open ksd-report-transformer.onrender.com\n\n'
else
  printf '\n  That did not work. Screenshot this window and send it to Claude.\n\n'
fi

printf '  You can close this window.\n\n'
