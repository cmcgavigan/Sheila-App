# Sheila App v2

Receipt / income / business-trip PWA for Sheila McGavigan PMU. Full rebuild of
v1 as **one Python service** with **SQLite as the source of truth** — the Excel
workbook is now a generated export, not a live file.

## What changed vs v1

| | v1 | v2 |
|---|---|---|
| Server | Node/Express + Python helpers spawned per save | One FastAPI process |
| Data | her-expenses.xlsx written in place (lock/chart fragility) | SQLite (`data/sheila.db`), never locked |
| Excel | live workbook, SheetJS forbidden, append-only | generated on demand at `/export` — Steuerberater layout (§19, bilingual, auto-deductibility) + dashboard charts |
| Keep-alive | tray.js + watchdog.py + 3× .ps1 + .vbs + .cmd | one Windows service (NSSM), auto-restart, in-process backups/alerts |
| Install | manual | double-click `INSTALL.cmd` |

The PWA frontend (Out/In/Trip capture, offline queue, treatments editor) is the
proven v1 code, unchanged except a cache-version bump — the API contract is
identical.

## Install (laptop)

1. Double-click `INSTALL.cmd` (it elevates itself). It creates a venv, installs
   dependencies, copies your v1 `.env` (ports adjusted to 3002/8444), installs
   the `SheilaApp` Windows service via NSSM, opens the firewall, and starts it.
2. Open `https://localhost:3002/setup` and scan the QR wall with the phone
   (same one-minute flow as v1: Tailscale → app → Add to Home Screen).

Nothing else to babysit: starts at boot, restarts on crash, logs to
`data/logs/service.log`.

## Pages

- `/` — the PWA (Out / In / Trip)
- `/setup` — phone setup QR wall
- `/treatments` — PIN-gated treatments editor
- `/export` — generate + download the Steuerberater workbook

## Excel export

`app/export/steuer.py` is the ONLY file that writes Excel. Tabs: Info (rates +
deductibility table) · Reisen · Fahrten · Ausgaben · Einnahmen · Dashboard
(native charts). Formulas reference the Info rates, so Herr Schwarz can see the
derivation. When his format wishes arrive, change that one file and re-export —
the data never moves.

Tax parameters (0,30 €/km, €14/€28 per-diem, €5,60 breakfast reduction) are
seeded into the DB `meta` table and surfaced on the Info tab. **Pending
confirmation by Herr Schwarz.**

## Update flow

The PWA shows an "Update now" banner when GitHub is ahead. Tapping it runs
`git pull --ff-only`, reinstalls requirements if they changed, and exits —
NSSM restarts the service with the new code. Data (`data/`, `.env`) is
gitignored and never touched.

## Dev

```
python -m venv .venv && .venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python run.py
```
