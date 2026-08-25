# OCR2Word

Ứng dụng web chạy được **trên Modal theo nhu cầu**, **Windows/Docker cục bộ** hoặc **Vercel** để chuyển **PDF / PDF scan / hình ảnh / Excel** thành file **Word `.docx` có thể chỉnh sửa**. Giao diện và thông báo đều bằng tiếng Việt; OCR mặc định ưu tiên model tiếng Việt chất lượng cao để giữ dấu tốt hơn. Tùy chọn **Hỗ trợ với AI** có thể hiệu đính lỗi OCR sau chuyển đổi mà vẫn giữ cấu trúc Word.

## Tính năng

- Kéo-thả tối đa 20 file trong một lượt; một file trả về DOCX, nhiều file trả về ZIP.
- Hiển thị thanh tiến độ phần trăm cho tải lên, OCR/chuyển đổi và đóng gói kết quả.
- Tùy chọn **Hỗ trợ với AI miễn phí** dùng Google Gemini Free Tier cho cả local và Vercel; bảng, đoạn, ngắt trang và kiểu chữ hiện có không bị dựng lại. Nếu Gemini tạm hết quota hoặc lỗi mạng, ứng dụng tự thử lại rồi trả bản OCR thường thay vì làm hỏng cả lượt chuyển đổi.
- PDF scan dài được tự phân tích, chia thành nhiều đoạn OCR chạy qua Queue, ghép lại đúng thứ tự rồi mới hiệu đính AI; tiến độ tổng luôn tăng và file tạm được dọn sau khi hoàn tất.
- OCR trực tiếp ảnh PNG, JPG/JPEG, TIFF, BMP và WEBP thành chữ Word có thể sửa.
- Nhận diện bảng có đường kẻ trong hình ảnh và PDF scan, đưa nội dung OCR vào từng ô của bảng Word thật.
- Xóa đường kẻ khỏi vùng bảng rồi OCR lần hai để giảm ký tự rác; chọn kết quả rõ hơn cho từng ô.
- Mở rộng các lưới bị lưu đồ/khối vẽ cắt rời để giữ đủ cột, ô gộp và số bước trong bảng Word.
- Chuẩn hóa nội dung OCR sang Unicode và font Times New Roman, lọc ký tự nhiễu và sửa một số lỗi dính từ tiếng Việt thường gặp.
- Đọc cỡ chữ do Tesseract ước lượng từ hOCR thay vì lấy ký tự có dấu cao nhất, nhờ đó cỡ chữ trong cùng một đoạn ổn định hơn.
- Phân tích độ dày nét, độ nghiêng và đường gạch dưới trên ảnh scan để khôi phục in đậm, in nghiêng, gạch chân theo từng từ/cụm từ.
- Ghép các dòng OCR liên tiếp thành đoạn Word thật, giữ căn đều và thụt đầu dòng; PDF đầu ra dùng lại kích thước trang nguồn.
- Chế độ **Giữ bố cục PDF chữ** bảo toàn trang bìa, nền, logo, font, bảng và mục lục tốt hơn đối với PDF có lớp chữ.
- Tự khôi phục khoảng trắng bị thiếu trong lớp chữ PDF dựa trên khoảng cách hình học giữa các ký tự.
- Tự nhận biết từng trang PDF có lớp chữ hay cần OCR.
- Tự xoay lại trang scan bị ngang/ngược khi lần OCR đúng chiều có độ tin cậy cao hơn rõ rệt.
- Ba chế độ PDF: tự động, OCR toàn bộ, hoặc chỉ trích xuất lớp chữ.
- Giữ đoạn văn, cỡ chữ, in đậm/nghiêng, hình ảnh và bảng PDF ở mức tốt nhất.
- Chuyển từng sheet Excel thành bảng Word thật; hỗ trợ XLSX, XLSM, XLS, ô gộp, màu nền, kiểu chữ và định dạng số phổ biến.
- Tự chia sheet quá rộng/thấp thành nhiều bảng có hàng tiêu đề lặp lại để tránh tràn lề.
- Bản Modal có thể chạy công khai hoặc có mật khẩu gia đình, tự bật container khi có người dùng và tự tắt khi rảnh; không cần giữ máy cá nhân hoạt động.
- Bản cục bộ xử lý dữ liệu ngay trên máy. Bản Vercel tải thẳng vào Blob riêng tư, xóa file nguồn sau chuyển đổi và tự dọn kết quả sau 24 giờ.
- Có image Modal, Dockerfile cục bộ và `Dockerfile.vercel` cài sẵn OCR Việt/Anh, không cần cấu hình Tesseract thủ công trên máy chủ.

## Deploy miễn phí theo nhu cầu lên Modal — khuyến nghị cho 1–2 người

Đây là phương án phù hợp nhất khi người nhà chỉ thỉnh thoảng sử dụng và máy cá nhân không chạy 24/7:

```text
Trình duyệt -> FastAPI web trên Modal
             -> job OCR bất đồng bộ (1 container CPU)
             -> DOCX/ZIP tải về từ kết quả job
               Modal Dict chỉ giữ trạng thái và phần trăm tiến độ
               Google Gemini chỉ được gọi khi bật "Hỗ trợ với AI"
```

Modal tự khởi động container khi có request và tự scale về 0 khi không còn việc, nên không có máy ảo chạy nền. Mã nguồn không đặt giới hạn số lượt theo giờ; `max_containers=1` chỉ xếp các lượt OCR vào hàng đợi tuần tự để tránh một đợt truy cập làm tăng chi phí đột ngột.

### 1. Cài CLI, đăng nhập và tạo bí mật

Nếu chưa có `.venv`, chạy `./setup.ps1` một lần. Sau đó:

Website công khai như bản đang triển khai:

```powershell
.\setup-modal.ps1 -PublicSite
```

Nếu muốn yêu cầu tên đăng nhập/mật khẩu, chạy `.\setup-modal.ps1` không có `-PublicSite`.

Script sẽ:

1. cài Modal CLI trong `.venv` với chế độ UTF-8 an toàn cho đường dẫn OneDrive có dấu;
2. mở trình duyệt để đăng nhập/tạo tài khoản Modal;
3. tạo Modal Secret `ocr2word-secrets`; chế độ riêng tư sẽ hỏi tên đăng nhập và mật khẩu nhưng không in mật khẩu ra terminal;
4. lấy `GEMINI_API_KEY` và model từ `.env`, đưa vào Secret mà không đưa vào Git.

Modal yêu cầu có phương thức thanh toán ngay cả với Starter. Không nhập thông tin thẻ vào script hoặc mã nguồn; chỉ thêm tại trang **Usage & Billing** chính thức của Modal.

### 2. Khóa chi phí ở mức 0 đồng trước khi deploy

Starter hiện có **30 USD compute credit mỗi tháng**. Trong **Modal Dashboard → Usage & Billing**, đặt **Workspace Budget = 20 USD/tháng**. Workspace Budget là hard cap trên tổng usage **trước khi trừ credit**, nên mức 20 USD nằm dưới 30 USD credit; khi chạm trần, workload dừng thay vì tiếp tục tạo usage. Dự án không dùng GPU, region trả thêm hay Volume lưu trữ lâu dài.

Thiết lập này là lớp bảo vệ bắt buộc của bộ deploy. `deploy-modal.ps1` sẽ từ chối chạy nếu chưa truyền xác nhận `-BudgetConfirmed`. Người dùng vẫn phải tự kiểm tra Dashboard vì Modal không cho script này đọc hoặc thay đổi phương thức thanh toán.

### 3. Deploy và lấy URL

Sau khi đã nhìn thấy Budget 20 USD trong Dashboard:

```powershell
.\deploy-modal.ps1 -BudgetConfirmed
```

Lần đầu Modal sẽ build image Linux có Tesseract `vie` + `eng`, nên lâu hơn các lần cập nhật sau. URL `https://...modal.run` được lưu cục bộ trong `.modal-url` (đã bị Git bỏ qua). Với chế độ công khai, chỉ cần gửi URL cho người nhà; không gửi Modal token hoặc Gemini API key.

### 4. Kiểm tra website thật

Kiểm tra health và trạng thái OCR của website công khai:

```powershell
.\verify-modal.ps1 -PublicSite
```

Kiểm tra trọn luồng upload → hàng đợi → OCR → tải DOCX bằng một file thật:

```powershell
.\verify-modal.ps1 -PublicSite -TestFile "C:\duong-dan\tai-lieu.pdf"
```

Kết quả test được lưu vào `output/modal-web/`. Khi cập nhật mã nguồn, chỉ cần chạy lại lệnh deploy ở bước 3.

### Giới hạn và dữ liệu trên Modal

- Mỗi lượt nhận tối đa 20 file, tối đa 80 MB/file, tổng input tối đa 90 MB và output tối đa 90 MB. Mức đệm này nằm dưới giới hạn payload 100 MB hiện tại của Modal.
- Job OCR có timeout 35 phút; giao diện theo dõi phần trăm tối đa 40 phút. Tài liệu quá dài nên chia nhỏ.
- Input/output Function được Modal mã hóa khi lưu và xóa trong tối đa 7 ngày; trạng thái trong Modal Dict cũng hết hạn sau 7 ngày không hoạt động. Nên tải DOCX ngay sau khi chuyển đổi.
- Lượt đầu sau thời gian rảnh có cold start để bật web/worker container. Đây là đánh đổi để không tốn tài nguyên lúc không dùng.
- Gemini có quota Free Tier riêng. Nếu không bật AI thì không gọi Gemini; nếu quota AI lỗi, converter vẫn trả bản OCR thường.
- Có thể xem usage bất kỳ lúc nào bằng `./.venv/Scripts/modal.exe billing summary` hoặc trong Dashboard.

## Deploy toàn bộ lên Vercel

Kiến trúc deploy gồm hai Vercel Services trong cùng một project:

```text
Trình duyệt -> Vercel Blob riêng tư -> Vercel Queue
                                      -> container FastAPI + Tesseract
                                      -> DOCX/ZIP trong Blob
                 Redis giữ trạng thái và phần trăm tiến độ
                 Google Gemini chỉ được gọi khi bật "Hỗ trợ với AI"
```

Cách này tránh giới hạn request/response 4,5 MB của Vercel Functions: PDF, ảnh, Excel và DOCX/ZIP đi thẳng giữa trình duyệt/container với Blob bằng URL ký có thời hạn; API chỉ truyền JSON nhỏ.

Ứng dụng không áp đặt rate limit theo IP/người dùng. Vercel, Queue, Blob, Redis và Gemini vẫn có quota/chống lạm dụng riêng của nhà cung cấp; khi chạm quota, yêu cầu có thể bị chậm, retry hoặc phải chờ quota được cấp lại.

Chế độ **Tự động thông minh** không tái OCR cả một PDF hỗn hợp chỉ vì vài trang là ảnh scan. Bìa, mục lục, logo và bố cục của các trang PDF có chữ được giữ lại; các trang scan/phức tạp trong cùng tài liệu được giữ nguyên như ảnh. Chọn **Ưu tiên chỉnh sửa** chỉ khi chấp nhận mất nền, dấu mộc hoặc chữ ký để OCR lại mọi trang. Chọn **Ưu tiên giống bản gốc** khi hình thức quan trọng hơn khả năng sửa chữ trên trang scan.

### 1. Tạo project

1. Đưa thư mục này lên GitHub/GitLab/Bitbucket và import repository trong Vercel.
2. Trong **Project Settings → Build and Deployment → Framework Preset**, chọn **Services**. Đây là điều bắt buộc để Vercel đọc khối `services` trong `vercel.json`.
3. Giữ **Fluid Compute** được bật. Nên đặt Function Region cùng vùng với Blob/Redis; với người dùng Việt Nam có thể chọn Singapore (`sin1`).

### 2. Gắn tài nguyên trong Vercel

Trong đúng project vừa tạo:

1. Tạo một **Vercel Blob store loại Private** rồi Connect vào project. Vercel tự thêm token Blob.
2. Cài **Upstash Redis** từ Vercel Marketplace, tạo database và Connect vào project. Vercel sẽ tự thêm `KV_REST_API_URL`/`KV_REST_API_TOKEN` (hoặc tên Upstash cũ `UPSTASH_REDIS_REST_URL`/`UPSTASH_REDIS_REST_TOKEN`); ứng dụng hỗ trợ cả hai cặp.
3. Tạo Gemini API key tại [Google AI Studio](https://aistudio.google.com/apikey), sau đó chạy `./setup-gemini.ps1`. Script lưu khóa local trong file bị Git bỏ qua và thêm khóa mã hóa vào Vercel; không dán khóa vào mã nguồn hay chat.

Không tạo Blob store loại Public vì tài liệu người dùng không nên có URL công khai lâu dài.

### 3. Thêm biến môi trường

Tạo hai chuỗi bí mật khác nhau, tối thiểu 32 byte. Có thể sinh trong PowerShell:

```powershell
$secretBytes = New-Object byte[] 32
$random = [Security.Cryptography.RandomNumberGenerator]::Create()
$random.GetBytes($secretBytes)
[Convert]::ToBase64String($secretBytes)
$random.Dispose()
```

Thêm vào **Project Settings → Environment Variables** cho Production, Preview và Development:

| Biến | Giá trị |
|---|---|
| `OCR2WORD_INTERNAL_SECRET` | chuỗi bí mật thứ nhất; bảo vệ liên lạc web ↔ worker |
| `CRON_SECRET` | chuỗi bí mật thứ hai; bảo vệ tác vụ dọn Blob |
| `GEMINI_API_KEY` | khóa bí mật tạo tại Google AI Studio; dùng `setup-gemini.ps1` để cấu hình an toàn |
| `OCR2WORD_AI_MODEL` | `gemini-3.7-flash` |
| `OCR2WORD_QUEUE_REGION` | `sin1` để local, Queue và project dùng cùng vùng Singapore |

Nếu project được tạo trước ngày 30/06/2026 và build báo function vượt 250 MB, thêm `VERCEL_SUPPORT_LARGE_FUNCTIONS=1` rồi redeploy. Project mới được Vercel tự đưa vào Large Functions beta.

### 4. Deploy và kiểm tra

Nhấn **Deploy** trong Dashboard. `vercel.json` sẽ:

- build `frontend/` thành Next.js service công khai;
- build `Dockerfile.vercel` thành container worker nội bộ có Tesseract `vie` + `eng`;
- tạo service bindings riêng tư hai chiều;
- gắn Queue consumer, tự retry tối đa ba lần;
- chạy cron mỗi ngày để xóa Blob đã hết hạn.

Sau deploy, mở `/api/status`. Trường `deployment.available` phải là `true`, OCR phải báo sẵn sàng và giao diện phải hiển thị **Blob riêng tư**. Hãy thử trước với một ảnh/PDF ngắn, sau đó mới bật AI.

### Chi phí và giới hạn Vercel

- Đây không phải hạ tầng miễn phí vô hạn. Vercel Functions, Blob, Queue và Upstash Redis đều tính theo hạn mức của plan; Hobby có quota miễn phí nhưng tác vụ OCR dùng CPU khá nhiều.
- Gemini Free Tier có quota và giới hạn tốc độ, không phải miễn phí vô hạn. Google có thể thay đổi model/quota; khi đó đổi `OCR2WORD_AI_MODEL` rồi redeploy.
- Vercel Hobby giới hạn mỗi lần chạy worker theo thời lượng tối đa của Hobby. `maxDuration: "max"` trong dự án tự dùng mức tối đa mà plan cho phép; tài liệu scan rất dài vẫn có thể cần Pro hoặc chia thành các file nhỏ hơn.
- Ứng dụng đã bỏ giới hạn “số lượt mỗi giờ” do chính mã nguồn đặt ra. Điều này không biến các dịch vụ bên ngoài thành miễn phí vô hạn và cũng không ngăn Vercel/Google áp quota của họ.
- File nguồn tự xóa ngay sau khi chuyển đổi thành công. Kết quả và job hết hạn sau 24 giờ; liên kết tải xuống được ký lại mỗi lần người dùng bấm tải.

## Kiểm thử local giống Vercel

Không dùng `run.ps1` cho bài kiểm thử này vì lệnh đó đi thẳng vào FastAPI cục bộ và bỏ qua Next.js, Blob, Queue, Redis cùng service bindings. Chế độ **local parity** dùng chính `vercel.json` và `Dockerfile.vercel`:

```text
Trình duyệt local -> Next.js chạy bởi vercel dev
                  -> Private Blob Development thật
                  -> Vercel Queue Development thật
                  -> container Linux từ Dockerfile.vercel
                  -> Upstash Redis Development thật
                  -> Google Gemini API thật khi bật AI
```

Nhờ vậy, mã API, signed URL, Queue retry, tiến độ Redis, Tesseract Linux và liên lạc web ↔ worker đều đi qua cùng đường chạy với production. Nên tạo Blob/Redis riêng cho **Development**, không dùng kho production để tránh trộn tài liệu test với dữ liệu thật.

### 1. Yêu cầu

- Node.js **24.x**. Dự án đã ghim phiên bản trong `.node-version`, `.nvmrc` và `frontend/package.json` vì Vercel hiện dùng Node 24 mặc định.
- Docker Desktop đang chạy; Vercel CLI cần Docker để dựng container worker.
- Một Vercel project đã chọn Framework Preset **Services**.
- Development environment đã gắn Private Blob, Upstash Redis và có đủ các biến ở bảng phía trên.

Nếu đang dùng NVM for Windows:

```powershell
nvm install 24
nvm use 24
node --version
```

### 2. Liên kết project và kéo cấu hình Development

Chạy một lần:

```powershell
.\setup-vercel-local.ps1
```

Script cài dependency từ lockfile, kiểm tra Node/Docker, đăng nhập Vercel, liên kết project, lấy OIDC token mới và chỉ kiểm tra **tên** biến môi trường; không in giá trị bí mật ra terminal.

### 3. Build và chạy toàn bộ Services local

```powershell
.\run-vercel-local.ps1 -BuildFirst
```

`-BuildFirst` chạy type-check rồi kiểm tra cả Next.js production build và image worker từ `Dockerfile.vercel` trước khi mở server. Trên Windows, script build riêng hai artifact để tránh lỗi quyền tạo symbolic link (`EPERM`) của `vercel build`; `vercel dev` ngay sau đó vẫn chạy toàn bộ Services graph và bindings. Script cũng tự dọn cache watcher `.next/dev` an toàn và thử khởi động lại một lần nếu lần đầu gặp `EPERM` trong thư mục OneDrive. Sau lần đầu có thể chạy nhanh hơn bằng `.\run-vercel-local.ps1`. Mở [http://127.0.0.1:3000](http://127.0.0.1:3000).

Không thêm cờ `-L`: chế độ này cố ý kết nối tới Blob, Queue, Redis và Google Gemini API của Development để kiểm tra đúng dịch vụ sẽ dùng khi deploy.

### 4. Kiểm thử tự động bằng file thật

Giữ cửa sổ server đang chạy, mở PowerShell thứ hai và chạy:

```powershell
.\verify-vercel-local.ps1 -FilePath "C:\Users\vinhv\Downloads\THU CAM KET - KGH.pdf"
```

Muốn kiểm tra cả Google Gemini:

```powershell
.\verify-vercel-local.ps1 -FilePath "C:\Users\vinhv\Downloads\THU CAM KET - KGH.pdf" -AiAssist
```

Bài test xác nhận Node 24, worker Linux/container, Tesseract `vie` + `eng`, Queue `sin1`, Blob/Redis, tiến độ không đi lùi, tải DOCX bằng signed URL và kiểm tra bên trong kết quả thật sự có `word/document.xml`. DOCX cùng báo cáo JSON được lưu trong `output/vercel-parity/`.

### 5. Phần chỉ Vercel Preview mới kiểm tra tuyệt đối được

Local parity không thể tái tạo phần cứng Fluid Compute, cold start, CDN/Firewall, giới hạn thời gian của plan và lịch Cron tự kích hoạt. Vì vậy cổng cuối trước production vẫn là một Preview Deployment:

```powershell
npx --yes vercel@59.4.0 deploy
```

Chỉ đưa lên production sau khi local parity, `vercel build` và một lượt chuyển đổi trên Preview đều đạt.

## Chạy nhanh bằng Docker

Yêu cầu: Docker Desktop hoặc Docker Engine có Compose.

```powershell
docker compose up --build
```

Mở [http://localhost:8000](http://localhost:8000). Image Docker cài sẵn Tesseract cùng dữ liệu `vie` và `eng`.

## Chạy trực tiếp trên Windows

### 1. Cài Python và thư viện

Yêu cầu Python 3.11 trở lên. Tại thư mục dự án:

```powershell
.\setup.ps1
```

Nếu Python không nằm trong `PATH`, trỏ trực tiếp tới file thực thi trước khi chạy:

```powershell
$env:OCR2WORD_PYTHON = "C:\duong-dan\python.exe"
.\setup.ps1
```

### 2. Cài OCR cho PDF scan

Cách tự động (dùng `winget`):

```powershell
.\setup.ps1 -InstallTesseract
```

Lệnh này cài Tesseract và mặc định tải các model `vie`, `eng`, `osd` chất lượng cao từ kho `tesseract-ocr/tessdata_best` vào `%LOCALAPPDATA%\OCR2Word\tessdata`. Đường dẫn không dấu này tránh lỗi Tesseract trên Windows khi dự án nằm trong thư mục OneDrive có ký tự Unicode.

Nếu đã cài Tesseract nhưng thiếu ngôn ngữ:

```powershell
.\scripts\download_languages.ps1 -Model best
```

Nếu cần giảm thời gian OCR và chấp nhận độ chính xác thấp hơn, dùng `-Model fast`.

Ứng dụng tự tìm Tesseract trong `PATH`, `C:\Program Files\Tesseract-OCR`, thư mục cài theo người dùng, hoặc đường dẫn được chỉ định bằng biến `TESSERACT_CMD`. Có thể đổi nơi lưu model bằng biến `OCR2WORD_TESSDATA_DIR`.

### 3. Chạy ứng dụng

```powershell
.\run.ps1
```

Sau đó mở [http://127.0.0.1:8000](http://127.0.0.1:8000). Có thể đổi cổng bằng `./run.ps1 -Port 8080`.

### 4. Dùng tùy chọn Hỗ trợ với AI

Tạo API key trong [Google AI Studio](https://aistudio.google.com/apikey), sau đó chạy lệnh sau một lần. Khi nhập, khóa được ẩn và không xuất hiện trong terminal:

```powershell
.\setup-gemini.ps1
```

Script cấu hình model `gemini-3.7-flash` cho local, Vercel Development, Preview và Production. Sau khi hoàn tất, mở lại ứng dụng và bật **Hỗ trợ với AI miễn phí**.

AI chỉ sửa nội dung chữ sau khi DOCX đã được tạo. Ứng dụng yêu cầu AI giữ nguyên số liệu và tự từ chối các đề xuất làm thay đổi quá nhiều nội dung. Gemini Free Tier có quota; theo chính sách của Google, dữ liệu gửi ở Free Tier có thể được dùng để cải thiện sản phẩm.

Có thể đổi model trước khi chạy ứng dụng:

```powershell
$env:OCR2WORD_AI_MODEL = "gemini-3.7-flash"
.\run.ps1
```

Docker Compose đọc `GEMINI_API_KEY` và `OCR2WORD_AI_MODEL` từ file `.env` do `setup-gemini.ps1` tạo.

## Chạy kiểm thử

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m pytest
Set-Location frontend
npm install
npm run typecheck
npm run build
```

Kiểm thử OCR scan được tự bỏ qua nếu máy chưa có Tesseract; chuyển PDF có lớp chữ, ảnh mô phỏng và Excel vẫn được kiểm tra đầy đủ.

## Cách chuyển đổi

### PDF

1. Với **Giữ bố cục PDF chữ**, đọc ký tự, hình ảnh, đường kẻ và tọa độ bằng PyMuPDF/pdf2docx rồi dựng từng trang Word theo đúng kích thước nguồn.
2. Ánh xạ các font PDF/TeX phổ biến sang font Word tương thích và suy luận khoảng trắng bị thiếu từ khoảng cách glyph.
3. Trong chế độ tự động, trang scan được raster hóa ở DPI đã chọn, OCR chữ và phân tích các đường kẻ ngang/dọc để tạo bảng Word theo từng ô; trang không bị chèn lại như một ảnh toàn trang.
4. Với **OCR toàn bộ**, cả trang PDF có lớp chữ cũng được nhận dạng lại để ưu tiên khả năng chỉnh sửa.
5. Khi tắt **Giữ bố cục PDF chữ**, ứng dụng dùng chế độ dựng nội dung theo đoạn văn/bảng để dễ tái cấu trúc hơn.

### Excel

1. Mở workbook ở cả chế độ công thức và giá trị đã lưu.
2. Mỗi sheet có dữ liệu trở thành một mục riêng trong Word.
3. Cột rộng được chia thành nhóm tối đa 10 cột; bảng dài được chia theo hàng và lặp tiêu đề.
4. Giá trị, công thức chưa có giá trị cache, ngày tháng, phần trăm, tiền tệ, ô gộp và một phần định dạng được giữ lại.

### Hình ảnh

1. Ảnh được tự xoay theo EXIF và chuyển về không gian màu phù hợp trước khi OCR.
2. Tesseract ưu tiên model tiếng Việt, sau đó ứng dụng chuẩn hóa Unicode, phục hồi khoảng trắng, đọc cỡ chữ hOCR và loại ký tự nhiễu có độ tin cậy thấp.
3. Nét mực trong từng vùng chữ được đo để suy ra in đậm, in nghiêng và gạch chân; các dòng thân bài liên tiếp được ghép lại thành đoạn Word có căn lề/thụt dòng phù hợp.
4. Khi bật **Nhận diện bảng**, OpenCV tìm lưới kẻ, kể cả một số ô gộp; vùng bảng được xóa đường kẻ và OCR lại trước khi đưa chữ vào từng ô Word.
5. TIFF nhiều trang được chuyển thành nhiều trang Word khi bật **Ngắt trang**.
6. Kết quả ưu tiên chữ và bảng có thể chỉnh sửa; nền ảnh gốc không được dùng làm nền Word.

## API

### Bản chạy cục bộ

- `GET /api/status`: phiên bản, định dạng, giới hạn và trạng thái Tesseract/ngôn ngữ.
- `GET /api/progress/{job_id}`: phần trăm và giai đoạn của một yêu cầu đang chạy.
- `POST /api/convert`: multipart form với một hoặc nhiều trường `files`.

Các tùy chọn form:

| Trường | Giá trị | Mặc định |
|---|---|---|
| `language` | `vie`, `vie+eng`, `eng` | `vie` |
| `ocr_mode` | `auto`, `force`, `text` | `auto` |
| `dpi` | `200`, `300`, `400` | `300` |
| `page_breaks` | boolean | `true` |
| `preserve_images` | boolean | `true` |
| `preserve_layout` | boolean | `true` |
| `detect_tables` | boolean | `true` |
| `ai_assist` | boolean; hiệu đính phần chữ bằng AI đã cấu hình | `false` |
| `job_id` | chuỗi 8-80 ký tự chữ/số/`_`/`-`, tùy chọn | trống |

Tài liệu tương tác của API có tại `/api/docs` khi ứng dụng đang chạy.

### Bản Vercel

- `POST /api/jobs`: tạo job, nhận các URL ký để trình duyệt PUT file thẳng vào Blob.
- `POST /api/jobs/{job_id}/start`: đưa job đã upload vào Vercel Queue.
- `GET /api/jobs/{job_id}`: phần trăm, trạng thái, báo cáo và URL tải kết quả có thời hạn.
- `/api/queue/convert` và `/api/internal/progress`: endpoint hệ thống, được Queue/service binding và secret bảo vệ.
- `GET /api/cron/cleanup`: cron Vercel gọi để xóa tài liệu hết hạn.

### Bản Modal

- `GET /healthz`: health check công khai, không đọc tài liệu hay tạo job OCR.
- `GET /api/status`: trạng thái OCR/AI và giới hạn của deployment; cần mật khẩu website.
- `POST /api/modal/jobs`: upload multipart và tạo job OCR bất đồng bộ; cần mật khẩu website.
- `GET /api/modal/jobs/{job_id}`: phần trăm, giai đoạn, báo cáo và URL tải khi hoàn tất.
- `GET /api/modal/jobs/{job_id}/download`: tải DOCX/ZIP từ kết quả Function; tối đa 7 ngày.

## Giới hạn thực tế

- PDF là định dạng trình bày cố định, còn Word là định dạng dòng chảy; vì vậy không thể đồng thời bảo đảm giống tuyệt đối và sửa tự do trong mọi tài liệu.
- AI hiệu đính văn bản, không thể tự biết chắc font, đậm/nghiêng hay nền gốc đã mất trong ảnh scan. Các thuộc tính này vẫn dựa trên phân tích pixel/OCR của ứng dụng.
- Ở local, Modal và Vercel, chỉ khi bật AI thì phần chữ OCR mới được gửi tới Google Gemini; file gốc không được gửi nguyên khối cho model.
- Chế độ giữ bố cục ưu tiên hình thức cho PDF có lớp chữ. Với trang scan, ứng dụng ưu tiên nội dung có thể sửa nên nền, dấu viết tay, font và vị trí tuyệt đối có thể thay đổi.
- Chế độ OCR toàn bộ ưu tiên chữ có thể sửa và có thể làm thay đổi nền, font, bảng hoặc vị trí đối tượng.
- Với tài liệu scan, tên font gốc không tồn tại dưới dạng dữ liệu nên ứng dụng dùng Times New Roman; đậm/nghiêng/gạch chân và cỡ chữ được suy ra từ pixel, có thể cần chỉnh nhẹ nếu ảnh quá mờ hoặc méo.
- Chuyển từ hình ảnh luôn cần Tesseract; nền ảnh không thể trở thành nền Word mà vẫn cho phép sửa độc lập toàn bộ nội dung.
- Bộ nhận diện bảng hoạt động tốt nhất với bảng có đường kẻ rõ, tương đối thẳng. Bảng không viền, đường mờ/đứt, ảnh chụp xiên hoặc lưu đồ phức tạp có thể cần chỉnh lại ô trong Word.
- Biểu đồ và hình nhúng trong Excel hiện được báo cảnh báo, không chuyển sang Word.
- Workbook được bảo vệ bằng mật khẩu cần được mở khóa trước.
- Local/Vercel nhận mỗi file tối đa 150 MB và tổng mỗi lượt tối đa 300 MB. Modal nhận 80 MB/file, 90 MB tổng input và 90 MB output do giới hạn truyền Function 100 MB. Sheet tối đa 10.000 dòng, 256 cột và 750.000 ô trong phạm vi có dữ liệu.

## Cấu trúc dự án

```text
app/
  main.py                 API FastAPI cục bộ + container worker Vercel
  converters/
    pdf.py                PDF chữ + OCR scan -> DOCX
    image.py              PNG/JPG/TIFF/BMP/WEBP -> DOCX bằng OCR
    ocr_layout.py         nhóm từ, phục hồi khoảng trắng và dựng dòng OCR
    raster_tables.py      dò lưới ảnh/PDF scan và tạo bảng Word theo từng ô
    layout_pdf.py         dựng PDF độ trung thực cao + sửa khoảng trắng/font
    excel.py              XLSX/XLSM/XLS -> DOCX
    base.py               style, bảng và cấu trúc DOCX dùng chung
  services/ocr.py         tìm Tesseract và chọn ngôn ngữ
  services/ai_assist.py   Google Gemini API, giữ style DOCX
  services/progress.py    lưu trạng thái tiến độ cho giao diện thăm dò
  services/serverless_jobs.py  chuyển đổi độc lập trong worker serverless
  static/                 giao diện web
tests/                    kiểm thử converter và API
scripts/                  tiện ích tải model OCR
modal_app.py              web + job queue + Tesseract image trên Modal
requirements-modal.txt    thư viện cài/deploy Modal
setup-modal.ps1           đăng nhập và tạo Modal Secret an toàn
deploy-modal.ps1          kiểm tra budget rồi deploy
verify-modal.ps1          kiểm thử URL và chuyển file thật
frontend/                 Next.js điều phối Blob, Queue, Redis và tiến độ
Dockerfile.vercel         image worker có Tesseract cho Vercel Functions
vercel.json               Services, bindings, Queue trigger và cron
setup-vercel-local.ps1    liên kết project và kiểm tra tài nguyên Development
run-vercel-local.ps1      chạy đúng Services graph bằng vercel dev
verify-vercel-local.ps1   kiểm thử E2E Blob -> Queue -> worker -> DOCX
```

## Nguồn công nghệ

- [Tesseract OCR](https://tesseract-ocr.github.io/tessdoc/), [tessdata_best](https://github.com/tesseract-ocr/tessdata_best) và model `tessdata_fast` tùy chọn
- [PyMuPDF](https://pymupdf.readthedocs.io/)
- [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx)
- [python-docx](https://python-docx.readthedocs.io/)
- [FastAPI](https://fastapi.tiangolo.com/)
- [openpyxl](https://openpyxl.readthedocs.io/)
- [Gemini API OpenAI compatibility](https://ai.google.dev/gemini-api/docs/openai), [Gemini structured outputs](https://ai.google.dev/gemini-api/docs/structured-output) và [Gemini API pricing](https://ai.google.dev/gemini-api/docs/pricing)
- [Modal pricing](https://modal.com/pricing), [Functions tự scale về 0](https://modal.com/docs/guide/functions), [job queue](https://modal.com/docs/guide/job-queue), [budgets](https://modal.com/docs/guide/budgets) và [data retention](https://modal.com/docs/guide/security)
- [Vercel Services](https://vercel.com/kb/guide/vercel-services), [Vercel Blob](https://vercel.com/docs/vercel-blob) và [Vercel Queues](https://vercel.com/docs/queues)
