"""Concurrent open-trip leg append and stale-client protection."""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor

import pytest


@pytest.fixture(autouse=True)
def _clear_trip(client):
    client.post("/api/trip/cancel")
    yield
    client.post("/api/trip/cancel")


def _start(client):
    response = client.post(
        "/api/trip/start", json={"purpose": "Client appointment"}
    )
    assert response.status_code == 200
    return response.json()["trip"]["tripToken"]


def test_two_concurrent_legs_both_survive_and_retry_is_idempotent(client):
    token = _start(client)

    def add(op):
        return client.post("/api/trip/leg", json={
            "tripToken": token, "operationId": op,
            "from": "Zuhause", "to": op, "km": 10, "overnight": False,
        })

    with ThreadPoolExecutor(max_workers=2) as pool:
        responses = list(pool.map(add, ["leg-op-a", "leg-op-b"]))

    assert [r.status_code for r in responses] == [200, 200]
    active = client.get("/api/trip/active").json()["trip"]
    assert {leg["operationId"] for leg in active["legs"]} == {
        "leg-op-a", "leg-op-b"
    }
    retry = add("leg-op-a")
    assert retry.status_code == 200
    assert retry.json()["idempotent"] is True
    assert len(client.get("/api/trip/active").json()["trip"]["legs"]) == 2


def test_stale_trip_token_is_rejected(client):
    old_token = _start(client)
    assert client.post("/api/trip/cancel").status_code == 200
    new_token = _start(client)
    assert old_token != new_token
    stale = client.post("/api/trip/leg", json={
        "tripToken": old_token, "operationId": "stale-leg",
        "from": "A", "to": "B", "km": 1,
    })
    assert stale.status_code == 409
    assert client.get("/api/trip/active").json()["trip"]["tripToken"] == new_token
