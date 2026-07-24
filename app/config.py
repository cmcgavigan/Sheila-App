"""Central configuration — everything comes from .env next to the project root.

Variable names match v1 where the meaning is identical, so the old .env can be
copied over and only the port entries need attention (v2 defaults to 3002/8444
so it can run alongside v1 during the switchover).
"""
from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
load_dotenv(ROOT / ".env")


def _int(name: str, default: int) -> int:
    try:
        return int(os.environ.get(name, "") or default)
    except ValueError:
        return default


# --- server ---------------------------------------------------------------
PORT = _int("PORT", 3002)
HOSTNAME_LOCAL = os.environ.get("HOSTNAME_LOCAL", "sheila.local")
TS_HTTPS_PORT = _int("TS_HTTPS_PORT", 8444)  # Tailscale-served HTTPS port
PUBLIC_URL = os.environ.get("PUBLIC_URL", "")  # override; else Tailscale auto

# --- integrations ----------------------------------------------------------
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")
GROQ_MODEL = os.environ.get("GROQ_MODEL", "qwen/qwen3.6-27b")
ORS_API_KEY = os.environ.get("ORS_API_KEY", "")
TREATMENTS_PIN = str(os.environ.get("TREATMENTS_PIN", ""))
AUTH_PASSWORD = os.environ.get("AUTH_PASSWORD", "")
AUTH_DISABLED = os.environ.get("AUTH_DISABLED", "") == "1"
AUTH_SECURE_COOKIE = os.environ.get("AUTH_SECURE_COOKIE", "1") != "0"
TRIP_AUTOMATION_TOKEN = os.environ.get("TRIP_AUTOMATION_TOKEN", "")

# --- alerts / watchdog (in-process) ----------------------------------------
NTFY_TOPIC = os.environ.get("NTFY_TOPIC", "")
NTFY_SERVER = os.environ.get("NTFY_SERVER", "https://ntfy.sh").rstrip("/")
WATCHDOG_ALERT_HOURS = _int("WATCHDOG_ALERT_HOURS", 0)  # 0 = no inactivity alert

# --- backups ----------------------------------------------------------------
BACKUP_RETENTION_DAYS = _int("BACKUP_RETENTION_DAYS", 30)
GDRIVE_BACKUP_DIR = os.environ.get("GDRIVE_BACKUP_DIR", "")  # optional 2nd copy

# --- data layout (all gitignored) -------------------------------------------
DATA_DIR = Path(os.environ.get("DATA_DIR", "") or (ROOT / "data"))
DB_PATH = DATA_DIR / "sheila.db"
PHOTOS_DIR = DATA_DIR / "photos"
BACKUPS_DIR = DATA_DIR / "backups"
EXPORTS_DIR = DATA_DIR / "exports"
CERT_DIR = DATA_DIR / "cert"
LOG_DIR = DATA_DIR / "logs"

for _d in (DATA_DIR, PHOTOS_DIR, BACKUPS_DIR, EXPORTS_DIR, CERT_DIR, LOG_DIR):
    _d.mkdir(parents=True, exist_ok=True)

APP_VERSION = "2.0.0"

# Out-sheet expense categories. Single source of truth — the PWA fetches these
# from /api/articles and the export's deductibility table is keyed on them.
ARTICLES = [
    "Fuel", "Meals (Travel)", "Client Entertainment", "Accommodation", "Parking",
    "Vehicle expenses", "Phone & Internet", "Insurance", "Training & Education",
    "Marketing & Advertising", "Equipment & Tools", "General Business Expenses",
    "PMU Pigments & Inks", "PMU Needles & Cartridges", "PMU Equipment & Machines",
    "Skincare & Treatment Products", "Disposables & PPE", "Salon Supplies",
]

TRIP_PURPOSES = [
    "Client appointment", "Supply / stock run",
    "Training / course", "Marketing / networking",
]

# --- agreed tax parameters (pending Herr Schwarz's confirmation) -------------
# These are DEFAULTS; live values are stored in the DB (meta table) so they can
# be changed without touching code, and they surface on the export's Info tab.
TAX_DEFAULTS = {
    "rate_km": 0.30,          # € per km, Kilometerpauschale
    "rate_meal_full": 28.0,   # € per full day / per overnight on multi-day trips
    "rate_meal_partial": 14.0,  # € single-day trip > 8h
    "breakfast_reduction": 5.60,  # € per provided breakfast
}

# Auto-deductibility classification: category -> (percent, treatment note DE/EN)
# fuel + own-trip travel meals are €0 because the Pauschalen already cover them;
# Bewirtung 70%; everything else 100%.
DEDUCTIBILITY = {
    "Fuel": (0, "Durch Kilometerpauschale abgegolten (covered by mileage allowance)"),
    "Meals (Travel)": (0, "Durch Verpflegungspauschale abgegolten (covered by per-diem)"),
    "Client Entertainment": (70, "Bewirtung §4 Abs. 5 Nr. 2 EStG (business entertainment)"),
}
DEDUCTIBILITY_DEFAULT = (100, "Voll abziehbar (fully deductible)")
