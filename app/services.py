"""External integrations: Groq vision OCR, Frankfurter EUR, OpenRouteService."""
from __future__ import annotations

import json
import re
from datetime import date
from typing import Optional

import httpx

from . import config

OCR_PROMPT_OUT = "\n".join([
    "You are a receipt data extractor. Read this EXPENSE receipt photo and reply with ONLY a JSON object",
    "(no markdown, no commentary) using exactly these keys:",
    "- merchant: store/company name (string)",
    "- date: transaction date as YYYY-MM-DD (string)",
    "- totalCost: final total paid, a number with no currency symbol",
    '- currency: ISO code (EUR, CHF, USD, GBP, ...) or "" if unclear',
    '- receiptCode: receipt/transaction/invoice reference, or "" if none',
    "- article: best-fit category, EXACTLY one of: " + ", ".join(config.ARTICLES) + ".",
    "    Choose the closest fit for a permanent-makeup / beauty business. PMU = permanent makeup.",
    '    Use "General Business Expenses" only if nothing else fits. Always pick one; never empty.',
    "- confidence: number 0.0-1.0, your honest confidence that merchant, date and total are correct.",
    "    Use below 0.5 if the photo is blurry, cropped, dark, glare-washed or hard to read.",
    'If a field cannot be read, use "" for strings or 0 for numbers (but article must always be chosen).',
])

OCR_PROMPT_IN = "\n".join([
    "You are an invoice/receipt data extractor for a permanent-makeup (PMU) artist who is being PAID by a client.",
    "Read this photo and reply with ONLY a JSON object (no markdown, no commentary) using exactly these keys:",
    '- clientName: the client/customer name if shown, else "" (string)',
    "- date: transaction date as YYYY-MM-DD (string)",
    "- totalCost: final total received, a number with no currency symbol",
    '- currency: ISO code (EUR, CHF, USD, GBP, ...) or "" if unclear',
    '- receiptCode: invoice/receipt reference, or "" if none',
    '- treatment: the service/treatment described, or "" if unclear (string)',
    "- confidence: number 0.0-1.0, your honest confidence that the total and date are correct.",
    'If a field cannot be read, use "" for strings or 0 for numbers.',
])


class QuotaError(Exception):
    def __init__(self, retry_after: int = 0):
        super().__init__("Groq quota exceeded")
        self.retry_after = retry_after


async def groq_ocr(base64_image: str, mime_type: str, mode: str) -> dict:
    if not config.GROQ_API_KEY:
        raise RuntimeError("GROQ_API_KEY not configured on server")
    prompt = OCR_PROMPT_IN if mode == "in" else OCR_PROMPT_OUT
    async with httpx.AsyncClient(timeout=60) as client:
        resp = await client.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={"Authorization": f"Bearer {config.GROQ_API_KEY}"},
            json={
                "model": config.GROQ_MODEL,
                "temperature": 0,
                "response_format": {"type": "json_object"},
                "messages": [{
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url",
                         "image_url": {"url": f"data:{mime_type};base64,{base64_image}"}},
                    ],
                }],
            },
        )
    if resp.status_code == 429:
        retry = 0
        m = re.search(r"try again in ([\d.]+)s", resp.text, re.I)
        if m:
            retry = int(float(m.group(1))) + 1
        raise QuotaError(retry)
    if resp.status_code != 200:
        raise RuntimeError(f"Groq HTTP {resp.status_code} {resp.text[:200]}")
    data = resp.json()
    text = data["choices"][0]["message"]["content"]
    return json.loads(text)


async def to_eur(total, currency: str, receipt_date: str) -> dict:
    """Frankfurter.app, rate for the receipt's exact date.

    Returns {eur, rate, rate_date, flagged}. flagged=True -> row status CHECK FX.
    """
    amount = float(total or 0)
    cur = str(currency or "").upper().strip()
    if not amount:
        return {"eur": 0.0, "rate": 1.0, "rate_date": None, "flagged": False}
    if not cur or cur == "EUR":
        return {"eur": amount, "rate": 1.0, "rate_date": receipt_date or None, "flagged": False}

    iso = receipt_date if re.match(r"^\d{4}-\d{2}-\d{2}$", str(receipt_date or "")) else None
    use_date = iso or date.today().isoformat()
    try:
        async with httpx.AsyncClient(timeout=8) as client:
            r = await client.get(
                f"https://api.frankfurter.app/{use_date}",
                params={"from": cur, "to": "EUR"},
            )
        r.raise_for_status()
        j = r.json()
        rate = j.get("rates", {}).get("EUR")
        if not isinstance(rate, (int, float)):
            raise RuntimeError("No EUR rate in response")
        return {
            "eur": round(amount * rate * 100) / 100,
            "rate": rate,
            "rate_date": j.get("date", use_date),
            "flagged": iso is None,  # had to substitute the date
        }
    except Exception as e:  # noqa: BLE001 — any failure just flags the row
        print(f"[eur] conversion failed for {cur} {use_date} - {e}")
        return {"eur": None, "rate": None, "rate_date": None, "flagged": True}


class NoOrsKey(Exception):
    pass


async def ors_distance_km(from_text: str, to_text: str) -> float:
    if not config.ORS_API_KEY:
        raise NoOrsKey()
    async with httpx.AsyncClient(timeout=15) as client:
        async def geocode(text: str):
            r = await client.get(
                "https://api.openrouteservice.org/geocode/search",
                params={
                    "api_key": config.ORS_API_KEY, "text": text,
                    "boundary.country": "DE", "size": 1,
                },
            )
            r.raise_for_status()
            feats = r.json().get("features") or []
            if not feats:
                raise RuntimeError(f'could not locate "{text}"')
            return feats[0]["geometry"]["coordinates"]  # [lng, lat]

        a = await geocode(from_text)
        b = await geocode(to_text)
        r = await client.get(
            "https://api.openrouteservice.org/v2/directions/driving-car",
            params={
                "api_key": config.ORS_API_KEY,
                "start": f"{a[0]},{a[1]}", "end": f"{b[0]},{b[1]}",
            },
        )
        r.raise_for_status()
        feats = r.json().get("features") or []
        try:
            metres = feats[0]["properties"]["summary"]["distance"]
        except (IndexError, KeyError):
            raise RuntimeError("no distance in ORS response")
    return round(metres / 100) / 10  # km, 1 dp


async def ntfy(title: str, message: str, priority: str = "default") -> None:
    """Fire-and-forget phone notification. Silently no-ops without a topic."""
    if not config.NTFY_TOPIC:
        return
    try:
        async with httpx.AsyncClient(timeout=10) as client:
            await client.post(
                f"{config.NTFY_SERVER}/{config.NTFY_TOPIC}",
                content=message.encode("utf-8"),
                headers={"Title": title, "Priority": priority},
            )
    except Exception as e:  # noqa: BLE001
        print(f"[ntfy] failed: {e}")
