"""Test isolation (audit A-02 / repair plan P0-2).

Every test runs against a throwaway data directory under pytest's tmp root:
- DATA_DIR is forced to that tmp dir BEFORE any app module is imported
  (app.config resolves paths at import time), via the autouse session fixture.
- HARD REFUSAL: if DATA_DIR was pre-set to anything inside the project trees
  (the v1 original or v2), or app.config was somehow imported first, or the
  resolved path is not our tmp dir, the whole run aborts.
- The API is exercised in-process through Starlette's TestClient (httpx ASGI
  transport underneath) — no live server, no port, no TLS, no production URL.

Import app modules inside fixtures/tests only, never at test-module top level.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

V2_ROOT = Path(__file__).resolve().parents[1]
V1_ROOT = V2_ROOT.parent  # the original app — must never be read from or written to
TEST_PIN = "4711"

sys.path.insert(0, str(V2_ROOT))


def _inside(p: Path, root: Path) -> bool:
    try:
        p = p.resolve()
        root = root.resolve()
    except OSError:
        return False
    return p == root or root in p.parents


# ---- hard refusals, checked before anything else --------------------------
_preset = os.environ.get("DATA_DIR", "")
if _preset and _inside(Path(_preset), V1_ROOT):
    pytest.exit(
        f"REFUSING to run tests: DATA_DIR={_preset!r} points inside the production "
        f"tree {V1_ROOT}. Unset it — tests always use a temporary directory.",
        returncode=3)
if "app.config" in sys.modules:
    pytest.exit(
        "REFUSING to run tests: app.config was imported before test isolation "
        "was established, so its data paths may point at production.",
        returncode=3)


@pytest.fixture(scope="session", autouse=True)
def data_dir(tmp_path_factory) -> Path:
    """Force all app data into a session tmp dir; verify the redirect took."""
    d = tmp_path_factory.mktemp("sheila-data")
    os.environ["DATA_DIR"] = str(d)
    os.environ["TREATMENTS_PIN"] = TEST_PIN
    os.environ["AUTH_DISABLED"] = "1"
    # Make sure no external integration can fire, whatever a stray .env says.
    for k in ("GROQ_API_KEY", "ORS_API_KEY", "NTFY_TOPIC", "GDRIVE_BACKUP_DIR",
              "PUBLIC_URL"):
        os.environ[k] = ""

    from app import config

    resolved = Path(config.DB_PATH).resolve()
    if _inside(resolved, V1_ROOT) or not _inside(resolved, d):
        pytest.exit(
            f"REFUSING to run tests: resolved DB path {resolved} is not inside "
            f"the isolated tmp dir {d}.", returncode=3)
    return d


@pytest.fixture(scope="session")
def client(data_dir):
    """In-process ASGI client against a freshly initialised isolated DB."""
    from starlette.testclient import TestClient

    from app import db
    from app.main import app as fastapi_app

    db.init_db()
    # No context manager: startup events (cert generation, background jobs,
    # current-url.txt) must NOT run in tests.
    return TestClient(fastapi_app)
