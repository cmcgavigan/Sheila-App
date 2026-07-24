"""In-process background jobs — replaces v1's separate backup.py, watchdog.py
and PowerShell scheduled tasks. One asyncio task, checked every 10 minutes:

- nightly backup (02:30): zip DB + photos into data/backups, apply retention,
  optional second copy to GDRIVE_BACKUP_DIR (a synced folder).
- heartbeat/inactivity alert via ntfy (optional, WATCHDOG_ALERT_HOURS).
- update check is on-demand via /api/update-check (PWA banner), not polled.
"""
from __future__ import annotations

import asyncio
import shutil
import sqlite3
import subprocess
import sys
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

from . import config, db, services

CHECK_EVERY = 600  # seconds
BACKUP_HOUR, BACKUP_MINUTE = 2, 30


def make_backup() -> Path:
    stamp = datetime.now().strftime("%y%m%d-%H%M")
    out = config.BACKUPS_DIR / f"sheila-backup-{stamp}.zip"
    tmp_db = config.BACKUPS_DIR / f"_snapshot-{stamp}.db"
    # Consistent snapshot even mid-write (SQLite online backup API).
    src = sqlite3.connect(config.DB_PATH)
    dst = sqlite3.connect(tmp_db)
    with dst:
        src.backup(dst)
    src.close()
    dst.close()
    try:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(tmp_db, "sheila.db")
            for p in sorted(config.PHOTOS_DIR.glob("*.jpg")):
                z.write(p, f"photos/{p.name}")
            env = config.ROOT / ".env"
            if env.exists():
                z.write(env, ".env")
    finally:
        tmp_db.unlink(missing_ok=True)

    # Retention.
    cutoff = datetime.now() - timedelta(days=config.BACKUP_RETENTION_DAYS)
    for p in config.BACKUPS_DIR.glob("sheila-backup-*.zip"):
        if datetime.fromtimestamp(p.stat().st_mtime) < cutoff:
            p.unlink(missing_ok=True)

    # Optional mirrored copy into a cloud-synced folder.
    if config.GDRIVE_BACKUP_DIR:
        try:
            mirror = Path(config.GDRIVE_BACKUP_DIR)
            mirror.mkdir(parents=True, exist_ok=True)
            shutil.copy2(out, mirror / out.name)
        except OSError as e:
            print(f"[backup] mirror copy failed: {e}")
    return out


async def _maybe_backup() -> None:
    now = datetime.now()
    due = now.replace(hour=BACKUP_HOUR, minute=BACKUP_MINUTE, second=0, microsecond=0)
    if now < due:
        return
    with db.get_db() as conn:
        last = db.get_meta(conn, "last_backup_date")
    if last == now.strftime("%Y-%m-%d"):
        return
    try:
        out = await asyncio.to_thread(make_backup)
        with db.get_db() as conn:
            db.set_meta(conn, "last_backup_date", now.strftime("%Y-%m-%d"))
        print(f"[backup] wrote {out.name}")
    except Exception as e:  # noqa: BLE001
        print(f"[backup] FAILED: {e}")
        await services.ntfy("Sheila app - backup failed", str(e), priority="high")


async def _maybe_inactivity_alert() -> None:
    if config.WATCHDOG_ALERT_HOURS <= 0:
        return
    with db.get_db() as conn:
        last = db.get_meta(conn, "last_activity")
        alerted = db.get_meta(conn, "inactivity_alerted")
    if not last:
        return
    try:
        last_dt = datetime.fromisoformat(last)
    except ValueError:
        return
    idle = datetime.now() - last_dt
    if idle > timedelta(hours=config.WATCHDOG_ALERT_HOURS):
        if alerted != last:  # alert once per idle stretch
            await services.ntfy(
                "Sheila app - quiet for a while",
                f"No receipts or trips logged in {idle.days * 24 + idle.seconds // 3600}h. "
                "Server is up - just a nudge.")
            with db.get_db() as conn:
                db.set_meta(conn, "inactivity_alerted", last)


async def background_loop() -> None:
    await services.ntfy("Sheila app", "Server started.", priority="low")
    while True:
        try:
            await _maybe_backup()
            await _maybe_inactivity_alert()
        except Exception as e:  # noqa: BLE001 — the loop must never die
            print(f"[jobs] error: {e}")
        await asyncio.sleep(CHECK_EVERY)


# --------------------------------------------------------------------------
# Auto-update (git installs only). NSSM restarts us after exit, so "update"
# is: git pull --ff-only, reinstall requirements if they changed, exit(0).
# --------------------------------------------------------------------------

def _git(args: list, timeout: int = 20):
    try:
        out = subprocess.run(
            ["git", "-C", str(config.ROOT)] + args,
            capture_output=True, text=True, timeout=timeout,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return out.stdout.strip() if out.returncode == 0 else None
    except (OSError, subprocess.TimeoutExpired):
        return None


def is_git_install() -> bool:
    return (config.ROOT / ".git").exists()


_last_fetch = 0.0


def update_status() -> dict:
    global _last_fetch
    if not is_git_install():
        return {"mode": "non-git", "updateAvailable": False}
    import time
    if time.time() - _last_fetch > 60:
        _last_fetch = time.time()
        _git(["fetch", "--quiet", "--tags", "origin"])
    branch = _git(["rev-parse", "--abbrev-ref", "HEAD"]) or "master"
    ref = f"origin/{branch}"
    behind = int(_git(["rev-list", "--count", f"HEAD..{ref}"]) or 0)
    current = _git(["describe", "--tags", "--always"]) or "unknown"
    latest = _git(["describe", "--tags", "--always", ref]) or current
    return {"mode": "git", "current": current, "latest": latest,
            "behind": behind, "updateAvailable": behind > 0}


def perform_update() -> dict:
    if not is_git_install():
        return {"error": "Not a git install - update manually."}
    st = update_status()
    if not st.get("updateAvailable"):
        return {"ok": True, "updated": False, "message": "Already up to date."}
    req = config.ROOT / "requirements.txt"
    before = req.read_text() if req.exists() else ""
    branch = _git(["rev-parse", "--abbrev-ref", "HEAD"]) or "master"
    pulled = _git(["pull", "--ff-only", "origin", branch], timeout=60)
    if pulled is None:
        return {"error": "git pull failed (local edits or conflict). Update manually."}
    after = req.read_text() if req.exists() else ""
    if before != after:
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "-r", str(req)],
                           timeout=300, check=False)
        except (OSError, subprocess.TimeoutExpired) as e:
            print(f"[update] pip install failed: {e}")
    return {"ok": True, "updated": True, "restarting": True}


def schedule_restart() -> None:
    """Exit shortly after the HTTP response flushes; NSSM restarts the service."""
    import threading

    def _exit():
        import time
        time.sleep(0.7)
        import os
        os._exit(0)  # hard exit — uvicorn's graceful shutdown can hang on open SSE

    threading.Thread(target=_exit, daemon=True).start()
