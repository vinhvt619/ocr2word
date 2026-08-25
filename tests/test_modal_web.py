from __future__ import annotations

import base64

import pytest
from fastapi.testclient import TestClient


pytest.importorskip("modal")
from modal_app import create_web_app  # noqa: E402


def _auth(username: str, password: str) -> dict[str, str]:
    token = base64.b64encode(f"{username}:{password}".encode("utf-8")).decode("ascii")
    return {"Authorization": f"Basic {token}"}


def test_modal_site_requires_family_password(monkeypatch):
    monkeypatch.setenv("OCR2WORD_BASIC_USER", "family")
    monkeypatch.setenv("OCR2WORD_BASIC_PASSWORD", "safe-password")
    client = TestClient(create_web_app())

    health = client.get("/healthz")
    assert health.status_code == 200
    assert health.json() == {"ok": True}

    unauthorized = client.get("/")
    assert unauthorized.status_code == 401
    assert unauthorized.headers["www-authenticate"].startswith("Basic")
    assert client.get("/api/status", headers=_auth("family", "wrong")).status_code == 401


def test_modal_site_can_be_public_without_family_password(monkeypatch):
    monkeypatch.delenv("OCR2WORD_BASIC_USER", raising=False)
    monkeypatch.delenv("OCR2WORD_BASIC_PASSWORD", raising=False)
    client = TestClient(create_web_app())

    assert client.get("/").status_code == 200
    assert client.get("/api/status").status_code == 200


def test_modal_status_and_static_interface_are_available(monkeypatch):
    monkeypatch.setenv("OCR2WORD_BASIC_USER", "family")
    monkeypatch.setenv("OCR2WORD_BASIC_PASSWORD", "safe-password")
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    headers = _auth("family", "safe-password")
    client = TestClient(create_web_app())

    status = client.get("/api/status", headers=headers)
    assert status.status_code == 200
    payload = status.json()
    assert payload["deployment"]["mode"] == "modal"
    assert payload["ocr"]["available"] is True
    assert payload["limits"]["max_file_mb"] == 80
    assert payload["limits"]["max_total_mb"] == 90
    assert status.headers["cache-control"] == "no-store"

    index = client.get("/", headers=headers)
    assert index.status_code == 200
    assert "OCR2Word" in index.text
    script = client.get("/static/app.js", headers=headers)
    assert script.status_code == 200
    assert "/api/modal/jobs" in script.text
