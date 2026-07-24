"""Boundary validation for money, dates, distances and request size."""
from __future__ import annotations


def test_invalid_receipt_values_are_rejected(client):
    assert client.post("/api/save-out", json={
        "operationId": "negative-money", "merchant": "x", "totalCost": -1,
    }).status_code == 422
    assert client.post("/api/save-out", json={
        "operationId": "nan-money", "merchant": "x", "totalCost": "NaN",
    }).status_code == 422
    assert client.post("/api/save-out", json={
        "operationId": "bad-date", "merchant": "x", "date": "24/07/2026",
    }).status_code == 422


def test_negative_trip_distance_and_unknown_purpose_are_rejected(client):
    assert client.post("/api/trip/start", json={"purpose": "not-a-purpose"}).status_code == 422
    assert client.post("/api/trip/start", json={"purpose": "Client appointment"}).status_code == 200
    active = client.get("/api/trip/active").json()["trip"]
    response = client.post("/api/trip/leg", json={
        "tripToken": active["tripToken"], "from": "A", "to": "B", "km": -1,
    })
    assert response.status_code == 422
    client.post("/api/trip/cancel")


def test_oversized_api_body_is_rejected(client):
    response = client.post("/api/save-out", content=b"x" * (8 * 1024 * 1024 + 1),
                           headers={"content-type": "application/json"})
    assert response.status_code == 413
