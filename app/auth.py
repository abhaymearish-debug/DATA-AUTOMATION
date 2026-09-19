"""Company-email login.

v1 is two people, so this deliberately avoids an email round-trip: no SMTP
credentials to hold, no magic-link delivery to fail at 8am when a report is
due. Passwords are PBKDF2-hashed in a file on the persistent volume, and access
is gated on an allowlist of company addresses.

The audit trail matters as much as the gate. Every build records who ran it,
which is something the previous setup could not tell you at all.
"""

from __future__ import annotations

import hashlib
import hmac
import json
import os
import secrets
from pathlib import Path

from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer

from . import config

_ITERATIONS = 240_000
_COOKIE = "ksd_session"


def _users_file() -> Path:
    return config.WORKSPACE_ROOT / "_auth" / "users.json"


def hash_password(password: str, salt: bytes | None = None) -> str:
    salt = salt or secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, _ITERATIONS)
    return f"pbkdf2_sha256${_ITERATIONS}${salt.hex()}${dk.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        algo, iters, salt_hex, hash_hex = stored.split("$")
        if algo != "pbkdf2_sha256":
            return False
        dk = hashlib.pbkdf2_hmac(
            "sha256", password.encode(), bytes.fromhex(salt_hex), int(iters)
        )
        return hmac.compare_digest(dk.hex(), hash_hex)
    except (ValueError, AttributeError):
        return False


def _invited_file() -> Path:
    return config.WORKSPACE_ROOT / "_auth" / "invited.json"


def load_invited() -> list[str]:
    """Addresses a signed-in user has invited from inside the app.

    The env allowlist is set at deploy time and needs a redeploy to change,
    which is no way to add a colleague on a Tuesday. This file is the same
    allowlist, editable by somebody who is already signed in - a deliberate act
    by a person with an account, not a way for a stranger to get one.
    """
    f = _invited_file()
    if not f.is_file():
        return []
    try:
        got = json.loads(f.read_text())
    except (json.JSONDecodeError, OSError):
        return []
    return [str(e).strip().lower() for e in got if str(e).strip()]


def save_invited(emails: list[str]) -> None:
    f = _invited_file()
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(sorted(set(emails)), indent=2))
    try:
        os.chmod(f, 0o600)
    except OSError:
        pass


def invite(email: str) -> None:
    """Let this address hold an account. Does not create one."""
    email = (email or "").strip().lower()
    if "@" not in email:
        raise ValueError("That does not look like an email address.")
    save_invited(load_invited() + [email])


def email_is_allowed(email: str) -> bool:
    email = email.strip().lower()
    if email in load_invited():
        return True
    if config.ALLOWED_EMAILS:
        if email in config.ALLOWED_EMAILS:
            return True
        # An explicit allowlist is a closed set; the domain rule does not widen it.
        if not config.ALLOWED_EMAIL_DOMAINS:
            return False
    domain = email.rpartition("@")[2]
    return bool(domain) and domain in config.ALLOWED_EMAIL_DOMAINS


def load_users() -> dict[str, str]:
    f = _users_file()
    if not f.is_file():
        return {}
    try:
        return json.loads(f.read_text())
    except json.JSONDecodeError:
        return {}


def save_users(users: dict[str, str]) -> None:
    f = _users_file()
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(users, indent=2))
    try:
        os.chmod(f, 0o600)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# The owner
#
# Everyone who can sign in reads every report; only one account manages who
# those people are. Without this, any colleague could remove any other - the
# install included its own owner - and the only way back was wiping the auth
# file on the server.
# ---------------------------------------------------------------------------


def _owner_file() -> Path:
    return config.WORKSPACE_ROOT / "_auth" / "owner.json"


def set_owner(email: str) -> None:
    email = (email or "").strip().lower()
    if not email:
        return
    f = _owner_file()
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps({"email": email}, indent=2))
    try:
        os.chmod(f, 0o600)
    except OSError:
        pass


def owner() -> str:
    """The one account that can add or remove people.

    Claimed by the first account created. An install that was already running
    before this existed has no file, so the first account in users.json - the
    one the handover page made - is taken as the owner and written down: the
    same answer, without anybody having to do anything.

    KSD_OWNER_EMAIL overrides both, and is the way back in. The owner cannot be
    removed from inside the app, so the only remaining lock-out is losing that
    account, and this settles it in Render without touching any data.
    """
    forced = (getattr(config, "OWNER_EMAIL", "") or "").strip().lower()
    if forced:
        # It outranks the file, and corrects it: an install that had guessed
        # its owner before this was set should not keep the guess on disk.
        if _stored_owner() != forced:
            set_owner(forced)
        return forced

    stored = _stored_owner()
    if stored:
        return stored

    users = load_users()
    if not users:
        return ""
    # Nothing on record says who set this install up, so this is a guess: the
    # first account written, preferring one named in the deploy-time allowlist
    # over a colleague invited from inside the app. It is NOT written down -
    # a guess that files itself is a guess nobody can correct without going
    # into the server. KSD_OWNER_EMAIL settles it for good.
    return next((e for e in users if e in config.ALLOWED_EMAILS), next(iter(users)))


def _stored_owner() -> str:
    f = _owner_file()
    if not f.is_file():
        return ""
    try:
        return str(json.loads(f.read_text()).get("email") or "").strip().lower()
    except (json.JSONDecodeError, OSError):
        return ""


def is_owner(email: str) -> bool:
    who = owner()
    return bool(who) and (email or "").strip().lower() == who


def set_password(email: str, password: str) -> None:
    """Used by scripts/bootstrap_user.py. Refuses addresses outside the allowlist."""
    email = email.strip().lower()
    if not email_is_allowed(email):
        raise ValueError(
            f"{email} is not in KSD_ALLOWED_EMAILS / KSD_ALLOWED_EMAIL_DOMAINS."
        )
    if len(password) < 12:
        raise ValueError("Use at least 12 characters.")
    users = load_users()
    users[email] = hash_password(password)
    save_users(users)


def authenticate(email: str, password: str) -> str | None:
    """Return the canonical email on success, None otherwise."""
    email = (email or "").strip().lower()
    users = load_users()
    stored = users.get(email)

    # Always run a verification so a wrong address and a wrong password take
    # the same time — otherwise the response time tells an attacker which
    # addresses exist.
    reference = stored or hash_password(secrets.token_urlsafe(16))
    ok = verify_password(password or "", reference)

    if stored and ok and email_is_allowed(email):
        return email
    return None


def _serializer() -> URLSafeTimedSerializer:
    return URLSafeTimedSerializer(config.SESSION_SECRET, salt="ksd-session")


def issue_session(email: str) -> str:
    return _serializer().dumps({"email": email})


def read_session(token: str | None) -> str | None:
    if not token:
        return None
    try:
        data = _serializer().loads(token, max_age=config.SESSION_MAX_AGE_SECONDS)
    except (BadSignature, SignatureExpired):
        return None
    email = data.get("email")
    # Re-check the allowlist on every request: revoking access should take
    # effect on the next page load, not whenever the cookie happens to expire.
    return email if email and email_is_allowed(email) else None


COOKIE_NAME = _COOKIE
