from __future__ import annotations

import codecs
import json
from pathlib import Path


def test_windows_powershell_scripts_use_utf8_bom() -> None:
    project_root = Path(__file__).resolve().parents[1]
    scripts = [
        project_root / "setup.ps1",
        project_root / "run.ps1",
        project_root / "scripts" / "download_languages.ps1",
    ]

    for script in scripts:
        assert script.read_bytes().startswith(codecs.BOM_UTF8), (
            f"{script.relative_to(project_root)} must use UTF-8 BOM for Windows PowerShell 5.1"
        )


def test_vercel_local_scripts_are_ascii_for_windows_powershell_51() -> None:
    project_root = Path(__file__).resolve().parents[1]
    scripts = [
        project_root / "setup-vercel-local.ps1",
        project_root / "run-vercel-local.ps1",
        project_root / "verify-vercel-local.ps1",
    ]

    for script in scripts:
        script.read_bytes().decode("ascii")


def test_vercel_local_parity_configuration_stays_in_sync() -> None:
    project_root = Path(__file__).resolve().parents[1]
    config = json.loads((project_root / "vercel.json").read_text(encoding="utf-8"))
    package = json.loads((project_root / "frontend" / "package.json").read_text(encoding="utf-8"))
    trigger = config["services"]["web"]["functions"]["pages/api/queue/convert.ts"][
        "experimentalTriggers"
    ][0]

    assert package["engines"]["node"] == "24.x"
    assert trigger["type"] == "queue/v2beta"
    assert trigger["topic"] == "ocr2word-convert"
    assert config["services"]["worker"]["entrypoint"] == "Dockerfile.vercel"

    dockerfile = (project_root / "Dockerfile.vercel").read_text(encoding="utf-8")
    assert dockerfile.startswith("FROM python:3.12-slim")
    assert "tesseract-ocr-vie" in dockerfile
    assert "tesseract-ocr-eng" in dockerfile

    env_example = (project_root / "frontend" / ".env.example").read_text(encoding="utf-8")
    assert "OCR2WORD_QUEUE_REGION=sin1" in env_example
