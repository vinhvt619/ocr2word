# OCR2Word

OCR2Word converts PDFs, scanned documents, images, and spreadsheets into editable Word (`.docx`) files or extracts structured tables and text into Excel (`.xlsx`) workbooks, with Vietnamese (`vie`) and English (`eng`) OCR support.

## Supported Formats

| Input Format | Word Export (`.docx`) | Excel Export (`.xlsx`) |
| --- | --- | --- |
| **Digital PDF** | Editable text flow, tables, and optional layout reconstruction | Extracted tables into sheets; non-table text in `Văn bản` |
| **Scanned PDF & images** (`.png`, `.jpg`, `.jpeg`, `.tif`, `.tiff`, `.bmp`, `.webp`) | OCR text and OpenCV-cleaned raster tables | Detected grid tables into sheets; remaining OCR text in `Văn bản` |
| **Spreadsheets** (`.xlsx`, `.xls`, `.xlsm`) | Converts worksheets into Word tables | *Not supported* (Word export only) |

## Quickstart

### Windows (PowerShell)

Requires Python 3.11+ and Windows PowerShell 5.1 or PowerShell 7.

1. Run [`setup.ps1`](setup.ps1) once to create `.venv`, install Python dependencies ([`requirements.txt`](requirements.txt)), install Tesseract via WinGet (if missing), and download the default `best` OCR language data (`vie`, `eng`, `osd`):
   ```powershell
   .\setup.ps1 -InstallTesseract
   ```
2. Start the local FastAPI server with [`run.ps1`](run.ps1):
   ```powershell
   .\run.ps1
   ```
3. Open [http://localhost:8000](http://localhost:8000), upload your files, open Conversion Settings to choose **Word (`.docx`)** or **Excel (`.xlsx`)**, and click the convert button to start conversion.
4. Useful local endpoints:
   - OCR readiness and installed languages (`vie`, `eng`): [http://localhost:8000/api/status](http://localhost:8000/api/status)
   - Interactive OpenAPI docs: [http://localhost:8000/api/docs](http://localhost:8000/api/docs)
5. **Port conflict (`WinError 10048`)**: Do not terminate running processes. Check [http://localhost:8000/api/status](http://localhost:8000/api/status) first—if OCR2Word is already running there, reuse that server; otherwise start on another port:
   ```powershell
   .\run.ps1 -Port 8001
   ```
   Then open [http://localhost:8001](http://localhost:8001).

### Docker Compose

Using [`compose.yaml`](compose.yaml) and [`Dockerfile`](Dockerfile) (bundled with Tesseract `vie` and `eng`):

```sh
docker compose up --build
```

Open [http://localhost:8000](http://localhost:8000), select your files, choose **Word (`.docx`)** or **Excel (`.xlsx`)**, and start conversion.

## Conversion Workflow & OCR Modes

- **Single vs. Batch Output**: Single-file jobs download `.docx` or `.xlsx` directly. Batch jobs (up to 20 files) bundle outputs into `ocr2word-ket-qua.zip` and include `bao-cao-chuyen-doi.txt` whenever any file fails or produces warnings.
- **OCR Modes** (evaluated per page for PDFs):
  - `auto` *(Default)*:
    - **Word export**: Weighs selectable digital text against largest-image coverage, using the optional digital Word layout path when visible digital text sits over a photo or graphic background, and routing scans or sparse/hidden-text pages to OCR.
    - **Excel export**: Applies separate per-page heuristics based on digital text and detected digital tables before falling back to OCR.
  - `force`: Runs OCR on every page, helpful when scans contain corrupted or misaligned hidden text layers.
  - `text`: Extracts digital text only and skips OCR entirely (no OCR engine needed).

## Equations, Excel Output & Limits

- **Editable Word Equations (`.docx` only)**: Reconstructs supported expressions into Office Math (`m:oMath` / `m:oMathPara`) via built-in PDF rules or an optional local CPU neural model (`pix2tex` + CPU PyTorch). Equation recognition is disabled for `.xlsx`, and running `pip install -r` on [`requirements-math-local.txt`](requirements-math-local.txt) alone does not install the model—follow the prerequisite commands in its header first.
- **Excel Workbooks (`.xlsx`)**: Writes a single detected table to `Bảng 1` or multiple tables to `T{page}_B{table}` (e.g. `T1_B1`), and places remaining text lines in `Văn bản` (`Trang`, `Dòng`, `Nội dung`). Strings with leading zeros and formula-like prefixes (`=`, `+`, `-`, `@`) stay text, while plain signed integers (without leading zeros) may remain numeric.
- **Practical Limits**:
  - **Excel bounds**: Maximum 200 PDF pages, 50,000 rows per sheet, 256 columns per sheet, 500,000 cells per workbook, 32,767 characters per cell, and 25,000,000 pixels per rasterized page.
  - **Tables, layouts & scans**: Borderless or low-contrast tables may merge or fall back to text, complex multi-column layouts may need manual edits in Word, and OCR accuracy depends on upright, sharp scans (300 DPI recommended).

## Troubleshooting

- **Refused `.pdf` upload**: Normal PDFs starting with `%PDF-` pass immediately. As a fallback only, if `%PDF-` follows prefix bytes within the first 1,024 bytes, OCR2Word accepts the file when PyMuPDF validates its structure. Files merely renamed to `.pdf` must be exported as actual PDFs from their source app; unreadable or damaged files should be re-exported.
- **Missing languages or low OCR quality**: Confirm `vie` and `eng` at `/api/status` (or re-run [`scripts/download_languages.ps1`](scripts/download_languages.ps1)), scan upright at 300–400 DPI, and switch OCR mode to `force` when a scan embeds faulty hidden text.

## Architecture & Cloud Deployment

- **Local FastAPI App**: Python conversion engines ([`app/converters`](app/converters)) and web API ([`app/main.py`](app/main.py)) serving the browser UI ([`app/static`](app/static)).
- **Modal Serverless**: CPU container configuration defined in [`modal_app.py`](modal_app.py), managed via [`setup-modal.ps1`](setup-modal.ps1), [`deploy-modal.ps1`](deploy-modal.ps1), and [`verify-modal.ps1`](verify-modal.ps1).
- **Vercel Full-Stack**: Next.js web app ([`frontend`](frontend)) in ESM JavaScript with JSDoc type checking ([`frontend/jsconfig.json`](frontend/jsconfig.json)), Vercel Blob, Vercel Queue, Upstash Redis, and the Python worker ([`Dockerfile.vercel`](Dockerfile.vercel), [`vercel.json`](vercel.json), [`setup-vercel-local.ps1`](setup-vercel-local.ps1), [`run-vercel-local.ps1`](run-vercel-local.ps1), [`verify-vercel-local.ps1`](verify-vercel-local.ps1)).
- **Vercel Job Routing**: Eligible single-file Word PDFs can be split into page ranges and reassembled into one `.docx`, whereas Excel jobs run as a direct single conversion.

## Development & Tests

Install Python test dependencies from [`requirements-dev.txt`](requirements-dev.txt) and run the suite, or run the Node.js 24 checks in [`frontend/`](frontend):

```sh
# Python backend & script tests
python -m pip install -r requirements-dev.txt
python -m pytest

# Vercel frontend checks (Node.js 24)
npm --prefix frontend ci
npm --prefix frontend test
npm --prefix frontend run typecheck
npm --prefix frontend run build
```
