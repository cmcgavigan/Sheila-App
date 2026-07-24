#!/usr/bin/env python3
"""
backup.py - daily backup of Sheila's data to Google Drive, 30-day retention.

Simple + no credentials: it zips the data and copies the zip INTO your Google
Drive desktop folder. Google Drive for Desktop uploads it to the cloud for you.
No API, no OAuth, no service account.

Each run:
  1. Zips her-expenses.xlsx + her-photos/ + treatments.json into
     backups/sheila-backup-YYMMDD-HHMM.zip  (a local copy is always kept).
  2. Copies that zip into your Drive folder (GDRIVE_BACKUP_DIR).
  3. Deletes backups older than BACKUP_RETENTION_DAYS in BOTH places.

.env keys:
  GDRIVE_BACKUP_DIR      full path to a folder inside your Google Drive, e.g.
                         G:\My Drive\Sheila-Backups
  BACKUP_RETENTION_DAYS  optional, default 30

Run manually:  python backup.py
Schedule:      see enable-backup.ps1
"""
import os, sys, glob, shutil, zipfile, datetime

HERE = os.path.dirname(os.path.abspath(__file__))


def env(key, default=None):
    v = os.environ.get(key)
    if v is not None:
        return v
    p = os.path.join(HERE, ".env")
    if os.path.exists(p):
        for line in open(p, encoding="utf-8"):
            line = line.strip()
            if line.startswith(key + "="):
                return line.split("=", 1)[1].strip()
    return default


EXCEL_PATH = os.path.join(HERE, env("EXCEL_PATH", "her-expenses.xlsx"))
PHOTOS_DIR = os.path.join(HERE, env("PHOTOS_DIR", "her-photos"))
TREATMENTS = os.path.join(HERE, env("TREATMENTS_PATH", "treatments.json"))
LOCAL_BACKUP_DIR = os.path.join(HERE, "backups")
DRIVE_DIR = env("GDRIVE_BACKUP_DIR", "")
RETENTION_DAYS = int(env("BACKUP_RETENTION_DAYS", "30"))
PREFIX = "sheila-backup-"


def log(msg):
    print(f"[backup] {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}  {msg}")


def make_zip():
    os.makedirs(LOCAL_BACKUP_DIR, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%y%m%d-%H%M")
    zip_path = os.path.join(LOCAL_BACKUP_DIR, f"{PREFIX}{stamp}.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        if os.path.exists(EXCEL_PATH):
            z.write(EXCEL_PATH, os.path.basename(EXCEL_PATH))
        if os.path.exists(TREATMENTS):
            z.write(TREATMENTS, os.path.basename(TREATMENTS))
        if os.path.isdir(PHOTOS_DIR):
            base = os.path.basename(PHOTOS_DIR.rstrip("/\\"))
            for root, _d, files in os.walk(PHOTOS_DIR):
                for f in files:
                    full = os.path.join(root, f)
                    z.write(full, os.path.join(base, os.path.relpath(full, PHOTOS_DIR)))
    log(f"zipped -> {os.path.basename(zip_path)} ({os.path.getsize(zip_path)//1024} KB)")
    return zip_path


def copy_to_drive(zip_path):
    if not DRIVE_DIR:
        log("GDRIVE_BACKUP_DIR not set in .env - keeping local copy only.")
        return
    if not os.path.isdir(DRIVE_DIR):
        log(f"Drive folder not found: {DRIVE_DIR}")
        log("  Is Google Drive for Desktop running, and is the path right? Local copy is safe.")
        return
    try:
        dest = os.path.join(DRIVE_DIR, os.path.basename(zip_path))
        shutil.copy2(zip_path, dest)
        log(f"copied to Google Drive: {dest}")
    except Exception as e:
        log(f"could not copy to Drive ({e}); local copy is safe.")


def prune(folder):
    if not folder or not os.path.isdir(folder):
        return
    cutoff = datetime.datetime.now() - datetime.timedelta(days=RETENTION_DAYS)
    removed = 0
    for p in glob.glob(os.path.join(folder, f"{PREFIX}*.zip")):
        try:
            if datetime.datetime.fromtimestamp(os.path.getmtime(p)) < cutoff:
                os.remove(p)
                removed += 1
        except OSError as e:
            log(f"could not delete {os.path.basename(p)}: {e}")
    if removed:
        log(f"pruned {removed} backup(s) older than {RETENTION_DAYS} days in {folder}")


def main():
    zip_path = make_zip()
    copy_to_drive(zip_path)
    prune(LOCAL_BACKUP_DIR)
    prune(DRIVE_DIR)
    log("done.")


if __name__ == "__main__":
    main()
