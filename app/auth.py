"""Small, local-only password/session layer for the household app."""
from __future__ import annotations

import hashlib
import hmac
import secrets
import time

from . import config

SESSION_COOKIE = "sheila_session"
CSRF_COOKIE = "sheila_csrf"
_sessions: dict[str, tuple[float, str]] = {}
_failures: dict[str, list[float]] = {}


def _hash_password(password: str, salt: bytes | None = None) -> str:
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, 310_000)
    return f"pbkdf2_sha256$310000${salt.hex()}${digest.hex()}"


def _check_password(password: str, encoded: str) -> bool:
    try:
        scheme, rounds, salt_hex, digest_hex = encoded.split("$", 3)
        if scheme != "pbkdf2_sha256":
            return False
        actual = hashlib.pbkdf2_hmac(
            "sha256", password.encode(), bytes.fromhex(salt_hex), int(rounds)
        ).hex()
        return hmac.compare_digest(actual, digest_hex)
    except (TypeError, ValueError):
        return False


def ensure_password(conn) -> str:
    """Return the configured password, creating a one-time local secret if needed."""
    row = conn.execute("SELECT value FROM meta WHERE key='auth.password_hash'").fetchone()
    if row and row["value"]:
        return ""
    password = config.AUTH_PASSWORD
    if not password:
        secret_path = config.DATA_DIR / "auth.secret"
        if secret_path.is_file():
            password = secret_path.read_text(encoding="utf-8").strip()
        else:
            password = secrets.token_urlsafe(24)
            secret_path.write_text(password + "\n", encoding="utf-8")
            try:
                secret_path.chmod(0o600)
            except OSError:
                pass
    if len(password) < 10:
        raise RuntimeError("AUTH_PASSWORD must be at least 10 characters")
    conn.execute(
        "INSERT OR REPLACE INTO meta(key, value) VALUES ('auth.password_hash', ?)",
        (_hash_password(password),),
    )
    return password


def login(conn, password: str, identity: str = "local") -> tuple[str, str] | None:
    now = time.time()
    recent = [t for t in _failures.get(identity, []) if now - t < 300]
    _failures[identity] = recent
    if len(recent) >= 5:
        return None
    row = conn.execute("SELECT value FROM meta WHERE key='auth.password_hash'").fetchone()
    if not row or not _check_password(password, row["value"]):
        recent.append(now)
        return None
    _failures.pop(identity, None)
    token, csrf = secrets.token_urlsafe(32), secrets.token_urlsafe(32)
    _sessions[token] = (time.time() + 8 * 3600, csrf)
    return token, csrf


def rate_limited(identity: str) -> bool:
    now = time.time()
    recent = [t for t in _failures.get(identity, []) if now - t < 300]
    _failures[identity] = recent
    return len(recent) >= 5


def session(request):
    token = request.cookies.get(SESSION_COOKIE, "")
    item = _sessions.get(token)
    if not item or item[0] <= time.time():
        _sessions.pop(token, None)
        return None
    return token, item[1]


def csrf_ok(request, csrf: str) -> bool:
    item = session(request)
    supplied = request.headers.get("X-CSRF-Token") or csrf
    return bool(item and hmac.compare_digest(item[1], supplied))
