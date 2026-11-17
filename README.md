# OCR2Word

[Tiếng Việt](#tiếng-việt) | [English](#english)

---

## Tiếng Việt

**OCR2Word** là ứng dụng web chuyển đổi PDF, tài liệu scan, hình ảnh và bảng tính sang file Word (`.docx`) có thể chỉnh sửa, hoặc trích xuất bảng và văn bản sang Excel (`.xlsx`), hỗ trợ OCR tiếng Việt (`vie`) và tiếng Anh (`eng`).

### Định dạng hỗ trợ

| Định dạng đầu vào | Xuất Word (`.docx`) | Xuất Excel (`.xlsx`) |
| --- | --- | --- |
| **PDF có lớp chữ** | Văn bản, bảng và tùy chọn giữ bố cục | Bảng tách thành các sheet; chữ ngoài bảng vào sheet `Văn bản` |
| **PDF scan & ảnh** (`.png`, `.jpg`, `.jpeg`, `.tif`, `.tiff`, `.bmp`, `.webp`) | Chữ OCR và bảng dựng lại theo lưới | Bảng nhận diện tách thành sheet; chữ OCR còn lại vào `Văn bản` |
| **Bảng tính** (`.xlsx`, `.xls`, `.xlsm`) | Chuyển các sheet thành bảng Word | *Không hỗ trợ* (chỉ xuất Word) |

### Chạy nhanh (Local)

#### Windows (PowerShell)
Yêu cầu Python 3.11+ và Windows PowerShell 5.1 hoặc PowerShell 7.

1. Tạo `.venv`, cài thư viện Python ([`requirements.txt`](requirements.txt)), cài Tesseract qua WinGet (nếu chưa có) và tải dữ liệu OCR `best` (`vie`, `eng`, `osd`):
   ```powershell
   .\setup.ps1 -InstallTesseract
   ```
2. Khởi động máy chủ FastAPI cục bộ:
   ```powershell
   .\run.ps1
   ```
3. Mở [http://127.0.0.1:8000](http://127.0.0.1:8000), tải file lên, chọn **Word (`.docx`)** hoặc **Excel (`.xlsx`)** rồi bắt đầu chuyển đổi.
   - Kiểm tra trạng thái OCR và ngôn ngữ: [http://127.0.0.1:8000/api/status](http://127.0.0.1:8000/api/status)
   - Tài liệu OpenAPI: [http://127.0.0.1:8000/api/docs](http://127.0.0.1:8000/api/docs)
   - Nếu trùng cổng (`WinError 10048`), kiểm tra `/api/status` xem máy chủ đã chạy chưa, hoặc đổi cổng bằng `.\run.ps1 -Port 8001`.

#### Docker Compose
Dùng [`compose.yaml`](compose.yaml) và [`Dockerfile`](Dockerfile) (đã tích hợp sẵn Tesseract `vie` và `eng`):

```sh
docker compose up --build
```

Sau đó truy cập [http://localhost:8000](http://localhost:8000).

### Chế độ OCR & Lưu ý giới hạn

- **Đầu ra đơn lẻ & hàng loạt**: Tải 1 file sẽ nhận trực tiếp `.docx` hoặc `.xlsx`. Tải nhiều file (tối đa **20 file**) sẽ nén thành `ocr2word-ket-qua.zip`, kèm `bao-cao-chuyen-doi.txt` nếu có cảnh báo hoặc lỗi.
- **Chế độ OCR (đánh giá theo từng trang PDF)**:
  - `auto` *(Mặc định)*: Tự phân biệt trang có lớp chữ số và trang scan để chọn luồng trích xuất hoặc OCR phù hợp.
  - `force`: Ép chạy OCR trên mọi trang (hữu ích khi trang scan có lớp chữ ẩn bị lỗi hoặc lệch).
  - `text`: Chỉ trích xuất lớp chữ có sẵn trong PDF và bỏ qua hoàn toàn bước OCR.
- **Giới hạn chính**:
  - **Upload cục bộ**: Tối đa **20 file/lượt**, **150 MB/file**, tổng **300 MB/lượt**.
  - **Xuất Excel (`.xlsx`)**: Tối đa **200 trang PDF**, **50.000 dòng/sheet**, **256 cột/sheet**, **500.000 ô/workbook**, **32.767 ký tự/ô** và **25.000.000 pixel/trang**. Các chuỗi có số `0` ở đầu hoặc bắt đầu bằng ký tự công thức (`=`, `+`, `-`, `@`) được giữ ở dạng văn bản.
  - **Chất lượng OCR & bảng**: Độ chính xác phụ thuộc vào ảnh rõ nét, ngay ngắn (khuyến nghị 300 DPI); bảng không viền hoặc bố cục nhiều cột phức tạp có thể cần chỉnh tay lại trong Word.
  - **Công thức toán học (chỉ `.docx`, tùy chọn)**: Hỗ trợ dựng Office Math qua quy tắc PDF tích hợp hoặc mô hình CPU tùy chọn (xem hướng dẫn cài trong [`requirements-math-local.txt`](requirements-math-local.txt)).

> Ngoài bản chạy cục bộ ([`app/main.py`](app/main.py)), repo còn có cấu hình triển khai cloud cho Modal ([`modal_app.py`](modal_app.py)) và Vercel ([`frontend/`](frontend), [`vercel.json`](vercel.json)).

---

## English

**OCR2Word** is a web application that converts PDFs, scanned documents, images, and spreadsheets into editable Word (`.docx`) files or extracts structured tables and text into Excel (`.xlsx`) workbooks, with Vietnamese (`vie`) and English (`eng`) OCR support.

### Supported Formats

| Input Format | Word Export (`.docx`) | Excel Export (`.xlsx`) |
| --- | --- | --- |
| **Digital PDF** | Editable text flow, tables, and optional layout reconstruction | Extracted tables into sheets; non-table text in `Văn bản` |
| **Scanned PDF & images** (`.png`, `.jpg`, `.jpeg`, `.tif`, `.tiff`, `.bmp`, `.webp`) | OCR text and reconstructed grid tables | Detected grid tables into sheets; remaining OCR text in `Văn bản` |
| **Spreadsheets** (`.xlsx`, `.xls`, `.xlsm`) | Converts worksheets into Word tables | *Not supported* (Word export only) |

### Quickstart (Local)

#### Windows (PowerShell)
Requires Python 3.11+ and Windows PowerShell 5.1 or PowerShell 7.

1. Create `.venv`, install Python dependencies ([`requirements.txt`](requirements.txt)), install Tesseract via WinGet (if missing), and download `best` OCR language data (`vie`, `eng`, `osd`):
   ```powershell
   .\setup.ps1 -InstallTesseract
   ```
2. Start the local FastAPI server:
   ```powershell
   .\run.ps1
   ```
3. Open [http://127.0.0.1:8000](http://127.0.0.1:8000), upload your files, select **Word (`.docx`)** or **Excel (`.xlsx`)**, and start conversion.
   - OCR readiness & languages: [http://127.0.0.1:8000/api/status](http://127.0.0.1:8000/api/status)
   - Interactive OpenAPI docs: [http://127.0.0.1:8000/api/docs](http://127.0.0.1:8000/api/docs)
   - On port conflict (`WinError 10048`), check `/api/status` first to see if OCR2Word is already running, or start on another port with `.\run.ps1 -Port 8001`.

#### Docker Compose
Using [`compose.yaml`](compose.yaml) and [`Dockerfile`](Dockerfile) (pre-bundled with Tesseract `vie` and `eng`):

```sh
docker compose up --build
```

Then open [http://localhost:8000](http://localhost:8000).

### OCR Modes & Key Limits

- **Single vs. Batch Output**: Single-file jobs download `.docx` or `.xlsx` directly. Batch jobs (up to **20 files**) bundle outputs into `ocr2word-ket-qua.zip`, including `bao-cao-chuyen-doi.txt` when any file produces warnings or fails.
- **OCR Modes (evaluated per PDF page)**:
  - `auto` *(Default)*: Automatically chooses digital extraction or OCR per page based on selectable text and page imagery.
  - `force`: Runs OCR on every page (useful when scans contain corrupted or misaligned hidden text layers).
  - `text`: Extracts digital text only and skips OCR entirely.
- **Key Limits**:
  - **Local uploads**: Up to **20 files per batch**, **150 MB per file**, and **300 MB total per request**.
  - **Excel export (`.xlsx`)**: Maximum **200 PDF pages**, **50,000 rows/sheet**, **256 columns/sheet**, **500,000 cells/workbook**, **32,767 characters/cell**, and **25,000,000 pixels/page**. Strings with leading zeros or formula-like prefixes (`=`, `+`, `-`, `@`) remain text.
  - **OCR & table quality**: Accuracy depends on upright, sharp scans (300 DPI recommended); borderless tables or complex multi-column layouts may need minor manual edits in Word.
  - **Equations (`.docx` only, optional)**: Reconstructs supported expressions via built-in PDF rules or an optional local CPU model (see setup comments in [`requirements-math-local.txt`](requirements-math-local.txt)).

> Beyond local FastAPI ([`app/main.py`](app/main.py)), this repository also includes cloud deployment configs for Modal ([`modal_app.py`](modal_app.py)) and Vercel ([`frontend/`](frontend), [`vercel.json`](vercel.json)).
