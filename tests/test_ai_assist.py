from __future__ import annotations

import json
import urllib.error
import urllib.request
from pathlib import Path

from docx import Document

from app.services import ai_assist
from app.services.ai_assist import AiEngineInfo, apply_ai_assistance_to_docx


def _ready_engine(model: str = "gemini-3.7-flash") -> AiEngineInfo:
    return AiEngineInfo(
        available=True,
        installed=True,
        server_running=True,
        model_available=True,
        model=model,
        provider="Google Gemini API",
        endpoint="https://generativelanguage.googleapis.com/v1beta/openai",
        requires_api_key=True,
        message="AI sẵn sàng cho kiểm thử.",
    )


def test_ai_corrects_docx_text_and_preserves_runs_and_tables(
    tmp_path: Path,
    monkeypatch,
) -> None:
    output = tmp_path / "ocr.docx"
    document = Document()
    paragraph = document.add_paragraph()
    title_run = paragraph.add_run("THƯ CAM ")
    title_run.bold = True
    accent_run = paragraph.add_run("KET")
    accent_run.italic = True
    table = document.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "Sốhóa đơn"
    document.save(output)

    calls: list[dict[str, object]] = []

    def fake_gemini_request(payload, **_kwargs):
        calls.append(payload)
        user_prompt = payload["messages"][1]["content"]
        source = json.loads(user_prompt.split("\n", 1)[1])
        replacements = {
            "THƯ CAM KET": "THƯ CAM KẾT",
            "Sốhóa đơn": "Số hóa đơn",
        }
        parsed = ai_assist._CorrectionBatch(
            corrections=[
                ai_assist._Correction(
                    id=segment["id"],
                    text=replacements.get(segment["text"], segment["text"]),
                )
                for segment in source["segments"]
            ]
        )
        return {"choices": [{"message": {"role": "assistant", "content": parsed.model_dump_json()}}]}

    monkeypatch.setattr(ai_assist, "get_ai_engine_info", _ready_engine)
    monkeypatch.setattr(ai_assist, "_request_gemini_json", fake_gemini_request)

    progress: list[int] = []
    result = apply_ai_assistance_to_docx(
        output,
        progress_callback=lambda percent, _message: progress.append(percent),
    )

    assert result.processed_paragraphs == 2
    assert result.changed_paragraphs == 2
    assert result.rejected_changes == 0
    assert progress[0] == 0
    assert progress[-1] == 100
    assert calls
    request_payload = calls[0]
    assert request_payload["model"] == "gemini-3.7-flash"
    assert request_payload["reasoning_effort"] == "low"
    assert request_payload["response_format"]["type"] == "json_schema"

    corrected = Document(output)
    assert corrected.paragraphs[0].text == "THƯ CAM KẾT"
    assert corrected.paragraphs[0].runs[0].bold is True
    assert corrected.paragraphs[0].runs[1].italic is True
    assert corrected.tables[0].cell(0, 0).text == "Số hóa đơn"


def test_ai_rejects_changes_that_alter_numbers(tmp_path: Path, monkeypatch) -> None:
    output = tmp_path / "numbers.docx"
    document = Document()
    document.add_paragraph("Tổng tiền 100.000 đồng")
    document.save(output)

    monkeypatch.setattr(ai_assist, "get_ai_engine_info", _ready_engine)
    monkeypatch.setattr(
        ai_assist,
        "_correct_batch",
        lambda _endpoint, _model, _units: ["Tổng tiền 900.000 đồng"],
    )

    result = apply_ai_assistance_to_docx(output)

    assert result.changed_paragraphs == 0
    assert result.rejected_changes == 1
    assert Document(output).paragraphs[0].text == "Tổng tiền 100.000 đồng"


def test_gemini_status_detects_api_key(monkeypatch) -> None:
    monkeypatch.setenv("GEMINI_API_KEY", "test-gemini-key")
    monkeypatch.delenv("OCR2WORD_AI_MODEL", raising=False)
    info = ai_assist.get_ai_engine_info()

    assert info.available is True
    assert info.server_running is True
    assert info.model_available is True
    assert info.requires_api_key is True
    assert info.provider == "Google Gemini API"
    assert info.model == "gemini-3.7-flash"
    assert info.endpoint == "https://generativelanguage.googleapis.com/v1beta/openai"


def test_gemini_request_uses_official_endpoint_and_bearer_key(monkeypatch) -> None:
    monkeypatch.setenv("GEMINI_API_KEY", "test-gemini-key-for-request")

    class FakeResponse:
        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

        def read(self) -> bytes:
            return b'{"choices":[]}'

    def fake_urlopen(request: urllib.request.Request, *, timeout: float):
        assert request.full_url == (
            "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions"
        )
        assert request.get_header("Authorization") == "Bearer test-gemini-key-for-request"
        assert timeout == ai_assist.AI_REQUEST_TIMEOUT_SECONDS
        return FakeResponse()

    monkeypatch.setattr(urllib.request, "urlopen", fake_urlopen)

    assert ai_assist._request_gemini_json({"model": "gemini-3.7-flash"}) == {
        "choices": []
    }


def test_gemini_request_retries_transient_failures(monkeypatch) -> None:
    monkeypatch.setenv("GEMINI_API_KEY", "test-gemini-key-for-retry")
    attempts = 0
    delays: list[float] = []

    class FakeResponse:
        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

        def read(self) -> bytes:
            return b'{"choices":[{"message":{"content":"ok"}}]}'

    def fake_urlopen(_request: urllib.request.Request, *, timeout: float):
        nonlocal attempts
        attempts += 1
        assert timeout == ai_assist.AI_REQUEST_TIMEOUT_SECONDS
        if attempts < 3:
            raise urllib.error.HTTPError(
                ai_assist.DEFAULT_AI_ENDPOINT,
                503,
                "Service Unavailable",
                {"Retry-After": "0"},
                None,
            )
        return FakeResponse()

    monkeypatch.setattr(urllib.request, "urlopen", fake_urlopen)
    monkeypatch.setattr(ai_assist.random, "uniform", lambda _start, _end: 0.0)
    monkeypatch.setattr(ai_assist.time, "sleep", delays.append)

    response = ai_assist._request_gemini_json({"model": "gemini-3.7-flash"})

    assert response["choices"]
    assert attempts == 3
    assert delays == [2.0, 4.0]


def test_gemini_corrects_docx(tmp_path: Path, monkeypatch) -> None:
    output = tmp_path / "gemini.docx"
    document = Document()
    document.add_paragraph("Sốhóa đơn")
    document.save(output)
    monkeypatch.setenv("GEMINI_API_KEY", "test-gemini-key")
    calls: list[dict[str, object]] = []

    def fake_gemini_request(payload, **_kwargs):
        calls.append(payload)
        user_prompt = payload["messages"][1]["content"]
        source = json.loads(user_prompt.split("\n", 1)[1])
        content = ai_assist._CorrectionBatch(
            corrections=[
                ai_assist._Correction(id=segment["id"], text="Số hóa đơn")
                for segment in source["segments"]
            ]
        ).model_dump_json()
        return {"choices": [{"message": {"role": "assistant", "content": content}}]}

    monkeypatch.setattr(ai_assist, "_request_gemini_json", fake_gemini_request)

    result = apply_ai_assistance_to_docx(output)

    assert result.changed_paragraphs == 1
    assert Document(output).paragraphs[0].text == "Số hóa đơn"
    assert calls[0]["model"] == "gemini-3.7-flash"
    assert calls[0]["response_format"]["type"] == "json_schema"
