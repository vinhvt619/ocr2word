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

def test_modal_create_job_xlsx_accepted(monkeypatch):
    monkeypatch.delenv("OCR2WORD_BASIC_USER", raising=False)
    monkeypatch.delenv("OCR2WORD_BASIC_PASSWORD", raising=False)
    stored_jobs = {}

    async def mock_put_aio(jid, data):
        stored_jobs[jid] = data

    import modal_app
    from unittest.mock import AsyncMock, MagicMock

    spawned = []
    mock_call = MagicMock()
    mock_call.object_id = "mock-call-id"

    async def mock_spawn_aio(jid, up, opt):
        spawned.append(jid)
        return mock_call

    mock_spawn = MagicMock()
    mock_spawn.aio = mock_spawn_aio

    mock_dict = MagicMock()
    mock_dict.put.aio = mock_put_aio

    monkeypatch.setattr(modal_app, "jobs", mock_dict)
    monkeypatch.setattr(modal_app.ocr_job, "spawn", mock_spawn)

    client = TestClient(create_web_app(), raise_server_exceptions=False)
    tiny_png = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15c4\x00\x00\x00\nIDATx\x9cc\x00\x01\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82"
    response = client.post(
        "/api/modal/jobs",
        data={"options": '{"output_format": "xlsx"}'},
        files={"files": ("test.png", tiny_png, "image/png")},
    )
    assert response.status_code == 202
    res_data = response.json()
    assert "job_id" in res_data
    assert res_data["status"] == "queued"
    assert len(spawned) == 1
    assert spawned[0] == res_data["job_id"]
    assert res_data["job_id"] in stored_jobs


def test_modal_create_job_xlsx_rejects_excel_input(monkeypatch):
    monkeypatch.delenv("OCR2WORD_BASIC_USER", raising=False)
    monkeypatch.delenv("OCR2WORD_BASIC_PASSWORD", raising=False)
    client = TestClient(create_web_app(), raise_server_exceptions=False)
    response = client.post(
        "/api/modal/jobs",
        data={"options": '{"output_format": "xlsx"}'},
        files={"files": ("sheet.xlsx", b"dummycontent", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response.status_code == 400
    assert "ch\u1ec9 nh\u1eadn file PDF ho\u1eb7c h\u00ecnh \u1ea3nh" in response.json().get("detail", "")


def test_modal_create_job_xlsx_rejects_incompatible_options(monkeypatch):
    monkeypatch.delenv("OCR2WORD_BASIC_USER", raising=False)
    monkeypatch.delenv("OCR2WORD_BASIC_PASSWORD", raising=False)
    client = TestClient(create_web_app(), raise_server_exceptions=False)
    tiny_png = b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15c4\x00\x00\x00\nIDATx\x9cc\x00\x01\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82"

    # Math OCR with xlsx
    resp_math = client.post(
        "/api/modal/jobs",
        data={"options": '{"output_format": "xlsx", "math_ocr": true}'},
        files={"files": ("test.png", tiny_png, "image/png")},
    )
    assert resp_math.status_code == 400
    assert "c\u00f4ng th\u1ee9c to\u00e1n" in resp_math.json().get("detail", "")

    # Invalid output format
    resp_inv = client.post(
        "/api/modal/jobs",
        data={"options": '{"output_format": "ppt"}'},
        files={"files": ("test.png", tiny_png, "image/png")},
    )
    assert resp_inv.status_code == 400
    assert "ch\u1ec9 h\u1ed7 tr\u1ee3 docx ho\u1eb7c xlsx" in resp_inv.json().get("detail", "")
