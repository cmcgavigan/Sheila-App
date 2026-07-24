"""All /api routes. The contract matches v1 exactly, so the proven PWA frontend
carries over with nothing but a cache-version bump."""
from __future__ import annotations

import asyncio
import base64
import binascii
import hashlib
import hmac
import json
import math
import os
import tempfile
import uuid
from datetime import datetime
from pathlib import Path

from fastapi import APIRouter, Request
from fastapi.responses import FileResponse, JSONResponse

from . import auth, config, db, images, services

router = APIRouter(prefix="/api")


def _automation_authorized(request: Request) -> bool:
    configured = config.TRIP_AUTOMATION_TOKEN
    supplied = request.headers.get("x-trip-automation-token", "")
    return bool(configured) and hmac.compare_digest(supplied, configured)


def _now_local_stamp() -> str:
    return datetime.now().strftime("%d.%m.%Y %H:%M")


def _fmt_coord(v) -> str:
    return f"{float(v):.6f}"


def _bounded_text(value, limit: int, default: str = "") -> str:
    text = str(value or "").strip()
    return text[:limit] if len(text) <= limit else text[:limit]


def _money(value, field: str) -> float:
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        raise ValueError(f"{field} must be a finite non-negative number")
    if not math.isfinite(number) or number < 0:
        raise ValueError(f"{field} must be a finite non-negative number")
    return number


# --------------------------------------------------------------------------
# OCR
# --------------------------------------------------------------------------
@router.post("/process-receipt")
async def process_receipt(request: Request):
    body = await request.json()
    image = body.get("image")
    mime = body.get("mimeType", "image/jpeg")
    mode = body.get("mode", "out")
    if not image or not isinstance(image, str):
        return JSONResponse({"error": 'Missing "image" (base64 string)'}, status_code=400)
    if not config.GROQ_API_KEY:
        return JSONResponse({"error": "GROQ_API_KEY not configured on server"}, status_code=500)
    b64 = image.split(",", 1)[1] if "," in image else image
    try:
        parsed = await services.groq_ocr(b64, mime, mode)
    except services.QuotaError as e:
        return JSONResponse({"error": "quota", "quota": True, "retryAfter": e.retry_after},
                            status_code=429)
    except Exception as e:  # noqa: BLE001
        print("process-receipt (groq) error:", e)
        return JSONResponse({"error": str(e) or "OCR failed"}, status_code=502)

    def num(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0

    conf = parsed.get("confidence")
    conf = float(conf) if isinstance(conf, (int, float, str)) and str(conf) != "" else None
    if mode == "in":
        return {
            "clientName": parsed.get("clientName") or "",
            "date": parsed.get("date") or "",
            "totalCost": num(parsed.get("totalCost")),
            "currency": parsed.get("currency") or "",
            "receiptCode": parsed.get("receiptCode") or "",
            "treatment": parsed.get("treatment") or "",
            "confidence": conf,
        }
    article = parsed.get("article")
    if article not in config.ARTICLES:
        article = "General Business Expenses"
    return {
        "merchant": parsed.get("merchant") or "",
        "date": parsed.get("date") or "",
        "totalCost": num(parsed.get("totalCost")),
        "currency": parsed.get("currency") or "",
        "receiptCode": parsed.get("receiptCode") or "",
        "article": article,
        "confidence": conf,
    }


# --------------------------------------------------------------------------
# Save (out / in)
# --------------------------------------------------------------------------
async def _handle_save(kind: str, body: dict):
    image = body.get("image")
    has_image = isinstance(image, str) and len(image) > 0
    operation_id = str(body.get("operationId") or uuid.uuid4()).strip()
    if not operation_id or len(operation_id) > 128:
        return JSONResponse({"error": "Invalid operationId"}, status_code=400)

    image_bytes = None
    if has_image:
        b64 = image.split(",", 1)[1] if "," in image else image
        try:
            image_bytes = base64.b64decode(b64, validate=True)
        except (binascii.Error, ValueError):
            return JSONResponse({"error": "Invalid receipt image"}, status_code=400)

    date_ = _bounded_text(body.get("date"), 32)
    if date_:
        try:
            datetime.fromisoformat(date_)
        except ValueError:
            return JSONResponse({"error": "date must be ISO formatted"}, status_code=422)
    try:
        total = _money(body.get("totalCost"), "totalCost")
    except ValueError as e:
        return JSONResponse({"error": str(e)}, status_code=422)
    currency = _bounded_text(body.get("currency"), 8).upper()
    receipt_code = _bounded_text(body.get("receiptCode"), 128)
    note = _bounded_text(body.get("note"), 2000)
    lat, lng = body.get("lat"), body.get("lng")
    try:
        location = f"{_fmt_coord(lat)}, {_fmt_coord(lng)}" if lat not in (None, "") and lng not in (None, "") else ""
    except (TypeError, ValueError):
        location = ""
    map_url = f"https://www.google.com/maps?q={location.replace(' ', '')}" if location else ""

    if kind == "out":
        if not body.get("merchant") and not date_ and not total and not receipt_code:
            return JSONResponse({"error": "Nothing to save"}, status_code=400)
    else:
        if not body.get("clientName") and not date_ and not total and not receipt_code:
            return JSONResponse({"error": "Nothing to save"}, status_code=400)

    fx = await services.to_eur(total, currency, date_)
    prefix = "O" if kind == "out" else "I"

    merchant = _bounded_text(body.get("merchant"), 256)
    article = _bounded_text(body.get("article"), 128)
    if article not in config.ARTICLES:
        article = "General Business Expenses"
    business_personal = (
        "Personal" if body.get("businessPersonal") == "Personal" else "Business"
    )
    client_name = _bounded_text(body.get("clientName"), 256)
    treatment = _bounded_text(body.get("treatment"), 128)
    captured_at = _bounded_text(body.get("capturedAt"), 64)
    semantic = {
        "kind": kind, "date": date_, "total": total, "currency": currency,
        "receipt_code": receipt_code, "note": note, "location": location,
        "merchant": merchant, "article": article,
        "business_personal": business_personal, "client_name": client_name,
        "treatment": treatment, "captured_at": captured_at,
        "image_sha256": hashlib.sha256(image_bytes).hexdigest() if image_bytes else "",
    }
    operation_hash = hashlib.sha256(
        json.dumps(semantic, sort_keys=True, separators=(",", ":")).encode()
    ).hexdigest()

    def response_for(row, *, idempotent=False):
        return {
            "ok": True, "serial": row["serial"], "photo": row["photo"],
            "generated": bool(row["generated_card"]), "eur": row["eur"],
            "fxFlagged": row["status"] == "CHECK FX",
            "operationId": operation_id, "idempotent": idempotent,
        }

    staging_dir = config.DATA_DIR / ".receipt-staging"
    staging_dir.mkdir(parents=True, exist_ok=True)
    config.PHOTOS_DIR.mkdir(parents=True, exist_ok=True)
    conn = db.connect()
    staged_path = None
    pending_path = None
    inserted = False
    try:
        # Serialize receipt-number allocation and operation-ID checks. A second
        # copy of the same operation waits here, then returns the committed row.
        conn.execute("BEGIN IMMEDIATE")
        existing = db.receipt_by_operation(conn, operation_id)
        if existing:
            conn.rollback()
            if existing["operation_hash"] != operation_hash:
                return JSONResponse(
                    {"error": "operationId was already used for different receipt data"},
                    status_code=409,
                )
            images.finalize_pending_photo(config.PHOTOS_DIR, existing["photo"])
            return response_for(existing, idempotent=True)

        serial = db.next_serial(conn, kind)
        photo_name = f"{prefix}{serial}.jpg"
        label = f"{prefix}{serial}"
        fd, temp_name = tempfile.mkstemp(
            prefix="receipt-", suffix=".jpg", dir=staging_dir
        )
        os.close(fd)
        staged_path = Path(temp_name)

        if has_image:
            images.watermark_and_save(image_bytes, label, staged_path)
        else:
            amount = f"{total:.2f} {currency}".strip() if total else ""
            fields = (
                [("Merchant", merchant), ("Category", article),
                 ("Amount", amount), ("Date", date_), ("Type", business_personal),
                 ("Receipt", receipt_code), ("Note", note)]
                if kind == "out" else
                [("Client", client_name), ("Treatment", treatment),
                 ("Amount", amount), ("Date", date_), ("Receipt", receipt_code),
                 ("Note", note)]
            )
            images.generate_receipt_card(label, fields, staged_path)

        pending_path = config.PHOTOS_DIR / f"{photo_name}.pending"
        os.replace(staged_path, pending_path)
        staged_path = None
        row = {
            "kind": kind, "serial": serial, "date": date_,
            "total": total, "currency": currency,
            "eur": fx["eur"], "fx_rate": fx["rate"], "fx_date": fx["rate_date"],
            "receipt_code": receipt_code, "note": note,
            "location": location, "map_url": map_url,
            "status": "CHECK FX" if fx["flagged"] else "OK",
            "photo": photo_name, "generated_card": 0 if has_image else 1,
            "captured_at": captured_at, "saved_at": _now_local_stamp(),
            "operation_id": operation_id, "operation_hash": operation_hash,
        }
        if kind == "out":
            row.update(merchant=merchant, article=article,
                       business_personal=business_personal)
        else:
            row.update(client_name=client_name, treatment=treatment)
        db.insert_receipt(conn, row)
        db.touch_last_activity(conn)
        conn.commit()
        inserted = True

        try:
            images.finalize_pending_photo(config.PHOTOS_DIR, photo_name)
        except OSError:
            # An ordinary finalization failure is compensated immediately. A
            # hard process crash is recovered from the .pending file at startup.
            cleanup = db.connect()
            try:
                cleanup.execute(
                    "DELETE FROM receipts WHERE operation_id = ?", (operation_id,)
                )
                cleanup.commit()
            finally:
                cleanup.close()
            inserted = False
            pending_path.unlink(missing_ok=True)
            raise

        saved = db.receipt_by_operation(conn, operation_id)
        return response_for(saved)
    except Exception:  # noqa: BLE001 - cleanup must cover all processing failures
        if conn.in_transaction:
            conn.rollback()
        # If compensation itself failed after the row committed, preserve the
        # pending photo so startup recovery can restore consistency.
        if pending_path and not inserted:
            pending_path.unlink(missing_ok=True)
        if staged_path:
            staged_path.unlink(missing_ok=True)
        return JSONResponse({"error": "Receipt could not be saved"}, status_code=500)
    finally:
        conn.close()


@router.post("/save-out")
async def save_out(request: Request):
    return await _handle_save("out", await request.json())


@router.post("/save-in")
async def save_in(request: Request):
    return await _handle_save("in", await request.json())


# --------------------------------------------------------------------------
# Treatments (PIN-gated editor)
# --------------------------------------------------------------------------
@router.get("/treatments")
async def get_treatments():
    with db.get_db() as conn:
        rows = conn.execute("SELECT name, price FROM treatments ORDER BY pos, id").fetchall()
    return {"treatments": [{"name": r["name"], "price": r["price"]} for r in rows]}


@router.post("/treatments/verify")
async def verify_pin(request: Request):
    body = await request.json()
    return {"ok": bool(config.TREATMENTS_PIN) and
            str(body.get("pin") or "") == config.TREATMENTS_PIN}


@router.post("/treatments")
async def set_treatments(request: Request):
    body = await request.json()
    if not config.TREATMENTS_PIN or str(body.get("pin") or "") != config.TREATMENTS_PIN:
        return JSONResponse({"error": "Wrong PIN"}, status_code=403)
    items = body.get("treatments")
    if not isinstance(items, list):
        return JSONResponse({"error": "treatments must be an array"}, status_code=400)
    clean = []
    for t in items:
        name = str((t or {}).get("name") or "").strip()
        if not name:
            continue
        try:
            price = float((t or {}).get("price") or 0)
        except (TypeError, ValueError):
            price = 0
        clean.append({"name": name, "price": price})
    with db.get_db() as conn:
        conn.execute("DELETE FROM treatments")
        for i, t in enumerate(clean):
            conn.execute(
                "INSERT OR REPLACE INTO treatments(name, price, pos) VALUES (?,?,?)",
                (t["name"], t["price"], i),
            )
    return {"ok": True, "treatments": clean}


# --------------------------------------------------------------------------
# Trips (Reisekosten) — same lifecycle as v1
# --------------------------------------------------------------------------
@router.get("/trip/purposes")
async def trip_purposes():
    return {"purposes": config.TRIP_PURPOSES}


@router.get("/trip/active")
async def trip_active():
    with db.get_db() as conn:
        return {"trip": db.read_open_trip(conn)}


@router.post("/trip/start")
async def trip_start(request: Request):
    body = await request.json()
    purpose = _bounded_text(body.get("purpose"), 128)
    if not purpose:
        return JSONResponse({"error": "purpose is required"}, status_code=400)
    if purpose not in config.TRIP_PURPOSES:
        return JSONResponse({"error": "unsupported trip purpose"}, status_code=422)
    now = datetime.now().strftime("%Y-%m-%dT%H:%M")
    trip = {"purpose": purpose, "start": now, "legs": [], "createdAt": now,
            "tripToken": str(uuid.uuid4())}
    with db.get_db() as conn:
        conn.execute("BEGIN IMMEDIATE")
        if db.read_open_trip(conn):
            return JSONResponse({"error": "A trip is already open. End or cancel it first."},
                                status_code=409)
        db.write_open_trip(conn, trip)
    return {"ok": True, "trip": trip}


@router.post("/trip/automation/start")
async def trip_automation_start(request: Request):
    """Bluetooth/Shortcuts start: no browser session or CSRF required."""
    if not _automation_authorized(request):
        return JSONResponse({"error": "Invalid automation token"}, status_code=401)
    purpose = config.TRIP_PURPOSES[0]
    now = datetime.now().strftime("%Y-%m-%dT%H:%M")
    trip = {"purpose": purpose, "start": now, "legs": [], "createdAt": now,
            "tripToken": str(uuid.uuid4())}
    with db.get_db() as conn:
        conn.execute("BEGIN IMMEDIATE")
        if db.read_open_trip(conn):
            return JSONResponse({"error": "A trip is already open."}, status_code=409)
        db.write_open_trip(conn, trip)
    return {"ok": True, "trip": trip}


@router.post("/trip/automation/end")
async def trip_automation_end(request: Request):
    """Bluetooth/Shortcuts finish: records the trip, even with zero legs."""
    if not _automation_authorized(request):
        return JSONResponse({"error": "Invalid automation token"}, status_code=401)
    with db.get_db() as conn:
        conn.execute("BEGIN IMMEDIATE")
        trip = db.read_open_trip(conn)
        if not trip:
            return JSONResponse({"error": "No open trip to end."}, status_code=409)
        legs = trip.get("legs") or []
        finished = {"purpose": trip.get("purpose", ""), "start": trip.get("start", ""),
                    "end": datetime.now().strftime("%Y-%m-%dT%H:%M"), "nights": 0,
                    "meal_reduction": 0, "overnight_cost": 0, "incidentals": 0,
                    "note": "Bluetooth: MB Bluetooth 13593"}
        trip_id = db.insert_trip(conn, finished, legs)
        db.clear_open_trip(conn)
        db.touch_last_activity(conn)
        km = round(sum(float(l.get("km") or 0) for l in legs) * 10) / 10
    return {"ok": True, "trip": trip_id, "legs": len(legs), "km": km}


@router.post("/trip/leg")
async def trip_leg(request: Request):
    body = await request.json()
    with db.get_db() as conn:
        active_trip = db.read_open_trip(conn)
    if not active_trip:
        return JSONResponse({"error": "No open trip. Start one first."}, status_code=409)
    requested_token = str(body.get("tripToken") or "").strip()
    if requested_token and requested_token != str(active_trip.get("tripToken") or ""):
        return JSONResponse({"error": "This trip is stale. Refresh and try again."},
                            status_code=409)
    from_ = _bounded_text(body.get("from"), 256)
    to = _bounded_text(body.get("to"), 256)
    overnight = bool(body.get("overnight"))
    if not from_ or not to:
        return JSONResponse({"error": "from and to are required"}, status_code=400)

    raw_km = body.get("km")
    if raw_km not in (None, ""):
        try:
            km = _money(raw_km, "km")
        except ValueError as e:
            return JSONResponse({"error": str(e)}, status_code=422)
    else:
        km = 0
    auto = False
    if km <= 0:
        try:
            km = await services.ors_distance_km(from_, to)
            auto = True
        except services.NoOrsKey:
            return JSONResponse(
                {"error": "no-distance",
                 "message": "Enter the km (auto-distance needs an OpenRouteService key)."},
                status_code=422)
        except Exception as e:  # noqa: BLE001
            return JSONResponse({"error": "distance-failed", "message": str(e)}, status_code=502)

    leg = {"date": datetime.now().strftime("%Y-%m-%d"), "from": from_, "to": to,
           "km": round(km * 10) / 10, "overnight": overnight}
    operation_id = _bounded_text(body.get("operationId") or uuid.uuid4(), 128)
    # Legacy clients did not send a token; bind those requests to the trip that
    # was just observed while new clients receive stale-trip protection.
    trip_token = requested_token or str(active_trip.get("tripToken") or "")
    operation_hash = hashlib.sha256(json.dumps(
        {"tripToken": trip_token, **leg}, sort_keys=True,
        separators=(",", ":")
    ).encode()).hexdigest()
    with db.get_db() as conn:
        try:
            saved_leg, idempotent = db.append_open_trip_leg(
                conn, trip_token, operation_id, operation_hash, leg
            )
            trip = db.read_open_trip(conn)
        except db.TripLegConflict as e:
            return JSONResponse({"error": str(e)}, status_code=409)
    return {"ok": True, "leg": saved_leg, "idempotent": idempotent,
            "autoDistance": auto, "trip": trip}


@router.post("/trip/cancel")
async def trip_cancel():
    with db.get_db() as conn:
        conn.execute("BEGIN IMMEDIATE")
        db.clear_open_trip(conn)
    return {"ok": True}


@router.post("/trip/end")
async def trip_end(request: Request):
    body = await request.json()
    with db.get_db() as conn:
        conn.execute("BEGIN IMMEDIATE")
        trip = db.read_open_trip(conn)
        if not trip:
            return JSONResponse({"error": "No open trip to end."}, status_code=409)
        legs = trip.get("legs") or []
        # A trip can be finished with no manually entered legs.  The simple
        # wife-facing flow only records start/end; detailed mileage can be
        # added later from the accounting view.
        nights = sum(1 for l in legs if l.get("overnight"))
        try:
            meal_reduction = _money(body.get("mealReduction"), "mealReduction")
            overnight_cost = _money(body.get("overnightCost"), "overnightCost")
            incidentals = _money(body.get("incidentals"), "incidentals")
        except ValueError as e:
            return JSONResponse({"error": str(e)}, status_code=422)
        try:
            hours = max(0.0, (datetime.now() -
                              datetime.fromisoformat(str(trip.get("start")))).total_seconds() / 3600)
        except (TypeError, ValueError):
            hours = 0
        allowance = (config.TAX_DEFAULTS["rate_meal_full"] * nights if nights > 0
                     else config.TAX_DEFAULTS["rate_meal_partial"] if hours > 8 else 0)
        if meal_reduction > allowance:
            return JSONResponse({"error": "mealReduction exceeds the calculated allowance"},
                                status_code=422)
        finished = {
            "purpose": trip.get("purpose", ""),
            "start": trip.get("start", ""),
            "end": datetime.now().strftime("%Y-%m-%dT%H:%M"),
            "nights": nights,
            "meal_reduction": meal_reduction,
            "overnight_cost": overnight_cost,
            "incidentals": incidentals,
            "note": _bounded_text(body.get("note"), 2000),
        }
        trip_id = db.insert_trip(conn, finished, legs)
        db.clear_open_trip(conn)
        db.touch_last_activity(conn)
        km = round(sum(float(l.get("km") or 0) for l in legs) * 10) / 10
    return {"ok": True, "trip": trip_id, "legs": len(legs), "km": km, "nights": nights}


# --------------------------------------------------------------------------
# Places + distance
# --------------------------------------------------------------------------
@router.get("/places")
async def get_places():
    with db.get_db() as conn:
        rows = conn.execute("SELECT name, address FROM places ORDER BY pos, id").fetchall()
    return {"places": [{"name": r["name"], "address": r["address"]} for r in rows]}


@router.post("/places")
async def set_places(request: Request):
    body = await request.json()
    items = body.get("places")
    if not isinstance(items, list):
        return JSONResponse({"error": "places must be an array"}, status_code=400)
    clean = []
    for p in items:
        name = str((p or {}).get("name") or "").strip()
        if name:
            clean.append({"name": name, "address": str((p or {}).get("address") or "").strip()})
    with db.get_db() as conn:
        conn.execute("DELETE FROM places")
        for i, p in enumerate(clean):
            conn.execute("INSERT OR REPLACE INTO places(name, address, pos) VALUES (?,?,?)",
                         (p["name"], p["address"], i))
    return {"ok": True, "places": clean}


@router.get("/distance")
async def distance(request: Request):
    # "from" is a Python reserved word, so read the query params directly.
    from_ = str(request.query_params.get("from") or "").strip()
    to = str(request.query_params.get("to") or "").strip()
    if not from_ or not to:
        return JSONResponse({"error": "from and to required"}, status_code=400)
    try:
        return {"km": await services.ors_distance_km(from_, to)}
    except services.NoOrsKey:
        return JSONResponse({"error": "no-ors-key",
                             "message": "No OpenRouteService key configured."}, status_code=422)
    except Exception as e:  # noqa: BLE001
        return JSONResponse({"error": str(e)}, status_code=502)


# --------------------------------------------------------------------------
# Health / misc
# --------------------------------------------------------------------------
@router.get("/health")
async def health(request: Request):
    with db.get_db() as conn:
        counts = {
            k: conn.execute(f"SELECT COUNT(*) c FROM receipts WHERE kind='{k}'").fetchone()["c"]
            for k in ("out", "in")
        }
        trips_n = conn.execute("SELECT COUNT(*) c FROM trips").fetchone()["c"]
    result = {
        "ok": True, "version": config.APP_VERSION, "db": str(config.DB_PATH),
        "photosDir": str(config.PHOTOS_DIR), "receipts": counts, "trips": trips_n,
        "excelReady": True, "pythonOk": True,  # legacy keys the PWA checks
    }
    if not auth.session(request):
        result.pop("db", None)
        result.pop("photosDir", None)
    return result


@router.get("/articles")
async def articles():
    return {"articles": config.ARTICLES}


# --------------------------------------------------------------------------
# Excel export — generate the Steuerberater workbook from the DB.
# --------------------------------------------------------------------------
@router.post("/export")
async def make_export():
    from .export.steuer import generate_workbook  # lazy import, openpyxl is heavy
    out = await asyncio.to_thread(generate_workbook)
    return {"ok": True, "file": out.name, "path": str(out),
            "download": f"/api/export/download/{out.name}"}


@router.get("/export/download/{name}")
async def download_export(name: str):
    safe = Path(name).name  # no traversal
    p = config.EXPORTS_DIR / safe
    if not p.exists():
        return JSONResponse({"error": "not found"}, status_code=404)
    return FileResponse(
        p, filename=safe,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
