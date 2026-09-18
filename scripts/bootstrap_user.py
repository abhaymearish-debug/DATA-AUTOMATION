#!/usr/bin/env python3
"""Create or reset a login.

    python3 scripts/bootstrap_user.py neelima@<company-domain>

Prompts for the password rather than taking it as an argument, so it does not
land in shell history or in the host's process list.

The address must already be permitted by KSD_ALLOWED_EMAILS or
KSD_ALLOWED_EMAIL_DOMAINS — this script grants a password, never access.
"""

from __future__ import annotations

import getpass
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import auth, config  # noqa: E402


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2

    email = sys.argv[1].strip().lower()

    if not config.ALLOWED_EMAILS and not config.ALLOWED_EMAIL_DOMAINS:
        print("Set KSD_ALLOWED_EMAILS or KSD_ALLOWED_EMAIL_DOMAINS first.")
        return 1

    if not auth.email_is_allowed(email):
        print(f"{email} is not permitted by the current allowlist.")
        print(f"  KSD_ALLOWED_EMAILS        = {config.ALLOWED_EMAILS or '(unset)'}")
        print(f"  KSD_ALLOWED_EMAIL_DOMAINS = {config.ALLOWED_EMAIL_DOMAINS or '(unset)'}")
        return 1

    password = getpass.getpass("New password (min 12 characters): ")
    if password != getpass.getpass("Repeat: "):
        print("Those did not match.")
        return 1

    try:
        auth.set_password(email, password)
    except ValueError as exc:
        print(str(exc))
        return 1

    existing = sorted(auth.load_users())
    print(f"\nPassword set for {email}.")
    print(f"Accounts that can now sign in: {', '.join(existing)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
