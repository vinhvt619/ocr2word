# OCR2Word

OCR2Word is a web app for turning PDFs and scanned documents into editable Word files and extracting tables into Excel. It reads Vietnamese and English text and processes up to 20 files per batch.

| Input | Word (`.docx`) | Excel (`.xlsx`) |
| --- | --- | --- |
| PDFs with a text layer | Text, tables, and page layout | Tables and page text |
| Scanned PDFs and images | OCR text and reconstructed tables | Detected tables and OCR text |
| Excel spreadsheets (`.xls`, `.xlsx`, `.xlsm`) | Worksheets as Word tables | |

The browser shows upload and conversion progress. A single-file job returns a Word or Excel document. A batch returns a ZIP with the converted files and includes a report when a file fails or produces warnings.

## Run locally

**Docker.** The image includes Tesseract with Vietnamese and English language data.

```sh
docker compose up --build
```

**Windows.** Install Python 3.11 or newer, then run these commands from the repository root. The setup script installs Tesseract through WinGet when needed and downloads the OCR language data.

```powershell
.\setup.ps1 -InstallTesseract
.\run.ps1
```

Open [localhost:8000](http://localhost:8000). Choose one or more files, select Word or Excel, and start the conversion. The local FastAPI service exposes interactive API docs at [localhost:8000/api/docs](http://localhost:8000/api/docs) and OCR status at [`/api/status`](http://localhost:8000/api/status).

## How it works

- **PDFs with selectable text.** PyMuPDF extracts text, geometry, and tables. Word export uses those structures to reconstruct editable content.
- **Scans and images.** Tesseract reads text and its position. OpenCV locates table gridlines, estimates text styling for Word, and cleans table images before a second OCR pass. The converter assigns recognized text to table cells using its position on the page.
- **Office export.** Word output uses `python-docx`. Excel output uses `openpyxl` and places detected tables in separate worksheets. Values with leading zeros and source text that resembles a formula remain text.
- **Cloud jobs.** Modal starts conversion jobs on demand and stores progress in Modal Dict. The Vercel setup uses Next.js, private Blob uploads, a Queue, Redis for job status, and a FastAPI worker. Large scanned PDFs can be split into page ranges and merged into a Word file. The browser polls by job ID in both deployments.

The core conversion code is in [`app/converters`](app/converters). The local API is in [`app/main.py`](app/main.py), with cloud entry points in [`modal_app.py`](modal_app.py) and [`frontend/`](frontend).

## Tests

From an activated Python virtual environment:

```sh
python -m pip install -r requirements-dev.txt
python -m pytest
```

For the Vercel frontend, use Node.js 24:

```sh
npm --prefix frontend ci
npm --prefix frontend test
npm --prefix frontend run build
```

## Limitations

OCR accuracy depends on scan quality, font size, and page orientation. Faint or missing gridlines can reduce table detection accuracy. Complex layouts may need manual adjustment in Word, and Excel export is limited to 200 pages per document. Editable Word equations are optional. The default local image includes standard OCR; recognizing equations from scans requires the optional local model.
