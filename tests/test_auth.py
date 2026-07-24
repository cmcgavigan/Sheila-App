"""Authentication primitives and no-default-credential guards."""
from __future__ import annotations

import sqlite3
from pathlib import Path


def test_password_is_hashed_and_login_is_rate_limited(monkeypatch):
    from app import auth, config

    monkeypatch.setattr(config, "AUTH_PASSWORD", "a-test-password-longer-than-16")
    conn = sqlite3.connect(":memory:")
    conn.row_factory = sqlite3.Row
    conn.execute("CREATE TABLE meta(key TEXT PRIMARY KEY, value TEXT)")
    auth.ensure_password(conn)
    encoded = conn.execute(
        "SELECT value FROM meta WHERE key='auth.password_hash'"
    ).fetchone()[0]
    assert encoded.startswith("pbkdf2_sha256$")
    assert "a-test-password-longer-than-16" not in encoded
    assert auth.login(conn, "a-test-password-longer-than-16", "auth-test")
    for _ in range(5):
        assert auth.login(conn, "wrong", "rate-test") is None
    assert auth.rate_limited("rate-test")
    conn.close()


def test_treatment_page_has_no_prefilled_pin():
    page = Path(__file__).resolve().parents[1] / "public" / "treatments.html"
    text = page.read_text(encoding="utf-8")
    assert 'value="9245"' not in text
    assert "Auto-unlock" not in text


def test_certificate_reuse_validates_sans_and_key(monkeypatch, tmp_path):
    from app import netinfo

    monkeypatch.setattr(netinfo, "CERT_KEY", tmp_path / "key.pem")
    monkeypatch.setattr(netinfo, "CERT_PEM", tmp_path / "cert.pem")
    assert netinfo.load_or_generate_cert(["192.0.2.10"], ["sheila.test"])
    assert not netinfo.load_or_generate_cert(["192.0.2.10"], ["sheila.test"])
    assert netinfo.load_or_generate_cert(["192.0.2.11"], ["sheila.test"])
