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


def email_is_allowed(email: str) -> bool:
    email = email.strip().lower()
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
