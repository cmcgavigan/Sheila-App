"""Failure, retry and concurrency checks for receipt persistence."""
from __future__ import annotations

import base64
import io
import sqlite3
from concurrent.futures import ThreadPoolExecutor

import pytest


@pytest.fixture(autouse=True)
def _remove_atomicity_test_receipts(data_dir):
    """Keep this module order-independent despite the legacy session DB fixture."""
    yield
    from app import db

    conn = db.connect()
    try:
        rows = conn.execute(
            "SELECT photo FROM receipts WHERE "
            "operation_id IN ("
            "'corrupt-image', 'filesystem-failure', 'processing-failure', "
            "'serial-retry', 'simultaneous-retry', 'crash-after-commit')"
        ).fetchall()
        conn.execute(
            "DELETE FROM receipts WHERE "
            "operation_id IN ("
            "'corrupt-image', 'filesystem-failure', 'processing-failure', "
            "'serial-retry', 'simultaneous-retry', 'crash-after-commit')"
        )
        conn.commit()
    finally:
        conn.close()
    for row in rows:
        (data_dir / "photos" / row["photo"]).unlink(missing_ok=True)
        (data_dir / "photos" / f"{row['photo']}.pending").unlink(missing_ok=True)

def _receipt_count():
    from app import db

    conn = db.connect()
    try:
        return conn.execute("SELECT COUNT(*) FROM receipts").fetchone()[0]
    finally:
        conn.close()


def _files(data_dir):
    return sorted(
        p.name for p in (data_dir / "photos").iterdir() if p.is_file()
    )


def _payload(operation_id: str, **changes):
    body = {
        "operationId": operation_id,
        "merchant": "Atomic Test",
        "article": "Salon Supplies",
        "date": "2026-07-24",
        "totalCost": 19.5,
        "currency": "EUR",
        "businessPersonal": "Business",
        "note": "same semantic request",
    }
    body.update(changes)
    return body


def test_corrupt_image_leaves_no_row_or_file(client, data_dir):
    before_n, before_files = _receipt_count(), _files(data_dir)
    r = client.post("/api/save-out", json=_payload(
        "corrupt-image", image="data:image/jpeg;base64,not-valid-%%%"
    ))
    assert r.status_code == 400
    assert _receipt_count() == before_n
    assert _files(data_dir) == before_files


def test_finalization_failure_is_compensated(client, data_dir, monkeypatch):
    from app import images

    before_n, before_files = _receipt_count(), _files(data_dir)

    def fail_finalization(*_args, **_kwargs):
        raise PermissionError("simulated read-only destination")

    monkeypatch.setattr(images, "finalize_pending_photo", fail_finalization)
    r = client.post("/api/save-out", json=_payload("filesystem-failure"))
    assert r.status_code == 500
    assert _receipt_count() == before_n
    assert _files(data_dir) == before_files
    assert not list((data_dir / ".receipt-staging").glob("*"))


def test_unexpected_image_processing_failure_cleans_staging(
        client, data_dir, monkeypatch):
    from PIL import Image

    from app import images

    buf = io.BytesIO()
    Image.new("RGB", (20, 20), "white").save(buf, "JPEG")
    encoded = base64.b64encode(buf.getvalue()).decode()
    before_n, before_files = _receipt_count(), _files(data_dir)

    def fail_processing(*_args, **_kwargs):
        raise RuntimeError("simulated decoder crash")

    monkeypatch.setattr(images, "watermark_and_save", fail_processing)
    r = client.post("/api/save-out", json=_payload(
        "processing-failure", image="data:image/jpeg;base64," + encoded
    ))
    assert r.status_code == 500
    assert _receipt_count() == before_n
    assert _files(data_dir) == before_files
    assert not list((data_dir / ".receipt-staging").glob("*"))


def test_serial_retry_is_idempotent_and_conflict_is_rejected(client, data_dir):
    before_n, before_files = _receipt_count(), _files(data_dir)
    body = _payload("serial-retry")
    first = client.post("/api/save-out", json=body)
    second = client.post("/api/save-out", json=body)

    assert first.status_code == second.status_code == 200
    assert first.json()["serial"] == second.json()["serial"]
    assert second.json()["idempotent"] is True
    assert _receipt_count() == before_n + 1
    assert len(_files(data_dir)) == len(before_files) + 1

    conflict = client.post(
        "/api/save-out", json=_payload("serial-retry", note="different")
    )
    assert conflict.status_code == 409
    assert _receipt_count() == before_n + 1


def test_simultaneous_same_operation_creates_one_receipt(client, data_dir):
    before_n, before_files = _receipt_count(), _files(data_dir)
    body = _payload("simultaneous-retry")

    with ThreadPoolExecutor(max_workers=2) as pool:
        responses = list(pool.map(
            lambda _n: client.post("/api/save-out", json=body), range(2)
        ))

    assert [r.status_code for r in responses] == [200, 200]
    assert len({r.json()["serial"] for r in responses}) == 1
    assert sorted(r.json()["idempotent"] for r in responses) == [False, True]
    assert _receipt_count() == before_n + 1
    assert len(_files(data_dir)) == len(before_files) + 1


def test_startup_recovery_promotes_committed_pending_photo(data_dir):
    from app import db, images

    conn = db.connect()
    try:
        serial = db.next_serial(conn, "out")
        photo = f"O{serial}.jpg"
        db.insert_receipt(conn, {
            "kind": "out", "serial": serial, "photo": photo,
            "operation_id": "crash-after-commit", "operation_hash": "test",
        })
        conn.commit()

        pending = data_dir / "photos" / f"{photo}.pending"
        pending.write_bytes(b"pending-photo")
        orphan = data_dir / "photos" / "O999999.jpg.pending"
        orphan.write_bytes(b"orphan")
        staging = data_dir / ".receipt-staging"
        staging.mkdir(exist_ok=True)
        (staging / "abandoned.tmp").write_bytes(b"temp")

        images.recover_pending_photos(conn, data_dir / "photos", staging)
        assert (data_dir / "photos" / photo).read_bytes() == b"pending-photo"
        assert not pending.exists()
        assert not orphan.exists()
        assert not list(staging.iterdir())
    finally:
        conn.close()


def test_existing_database_gets_operation_columns():
    from app import db

    conn = sqlite3.connect(":memory:")
    conn.row_factory = sqlite3.Row
    old_schema = db.SCHEMA.replace(
        "  operation_id TEXT,\n  operation_hash TEXT,\n", ""
    )
    conn.executescript(old_schema)
    db.init_schema(conn)
    cols = {r["name"] for r in conn.execute("PRAGMA table_info(receipts)")}
    assert {"operation_id", "operation_hash"} <= cols
    indexes = {r["name"] for r in conn.execute("PRAGMA index_list(receipts)")}
    assert "idx_receipts_operation_id" in indexes
    conn.close()
