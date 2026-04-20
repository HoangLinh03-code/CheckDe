<!-- @format -->

# TÀI LIỆU KỸ THUẬT: HỆ THỐNG SOÁT LỖI ĐÁP ÁN ĐỀ TRỘN TỰ ĐỘNG (CHECK ĐỀ)

## 1. TỔNG QUAN DỰ ÁN

**Check Đề** là một ứng dụng Desktop (phát triển bằng PyQt5) tích hợp trí tuệ nhân tạo (Google Vertex AI - Gemini 3.1 Pro Preview).

Mục tiêu cốt lõi của phần mềm là giải quyết bài toán kiểm thử dữ liệu thủ công trong giáo dục: **Đối chiếu và phát hiện lỗi giữa Đề thi gốc, Bảng đáp án Excel và các mã Đề đã bị xáo trộn (Đề trộn)**. Hệ thống có khả năng vượt qua rào cản về việc đảo vị trí câu hỏi, đảo thứ tự đáp án (A, B, C, D), và đặc biệt xử lý cực tốt các công thức Toán học/Hóa học phức tạp (OMML, MathType/WMF).

### Đầu vào (Inputs):

1. **File Đề Gốc (.docx hoặc .pdf)**.
2. **File Đề Trộn (.docx hoặc .pdf)** (Hỗ trợ quét hàng loạt mã đề).
3. **Bảng đáp án Excel (.xlsx)** chứa keys của đề gốc và các đề trộn.

### Đầu ra (Outputs):

- Báo cáo **Excel (.xlsx)** phân loại rõ ràng:
  - Lỗi sai đáp án (Sau khi đã map ngược câu hỏi từ đề trộn về đề gốc).
  - Lỗi trùng đáp án / trùng ký hiệu (VD: Hai câu A, hoặc nội dung A và B giống hệt nhau).
  - Lỗi vỡ định dạng Layout (Lệch Tab, lệch Indent).
  - Lỗi in ấn (Cắt đôi câu hỏi qua trang mới, Đề quá số trang quy định).
  - Lỗi nội dung (Thiếu câu, thừa câu so với đề gốc).

---

## 2. KIẾN TRÚC MÃ NGUỒN VÀ MODULES

Hệ thống được thiết kế linh hoạt, tách biệt giữa giao diện (GUI), Logic kết nối AI (API), và các bộ xử lý đặc thù cho từng môn học (Handlers).

```text
CheckDe/
├── callAPI.py                # Core API: Quản lý xác thực Google Vertex AI, File API Caching và Fallback
├── check_de.py               # Core Logic: parse text (docx/pdf), matching AI, detect lỗi cấu trúc
├── gui_app.py                # Giao diện người dùng (PyQt5), quản lý Multi-threading, Auto PDF convert, Prompt Manager
├── math_exam_handler.py      # Module chuyên biệt cho môn Toán/KHTN (Xử lý OMML, WMF, Cấu trúc 3 phần I, II, III)
├── prompt_loader.py          # Trình đọc và biên dịch các file template prompt (.txt)
├── prompts/                  # Thư mục chứa các System Prompt cấu hình cho AI
│   ├── prompt_math.txt       # Prompt cho Toán (Có trắc nghiệm, Đúng/Sai, Điền số)
│   ├── prompt_english.txt    # Prompt cho Tiếng Anh (Reading passage, trắc nghiệm)
│   └── prompt_other.txt      # Prompt cho các môn chung
├── .env                      # File cấu hình biến môi trường (Chứa keys GCP)
├── .gitignore                # Bỏ qua các file tạm, cache, venv khi push code
└── requirements.txt          # Danh sách thư viện Python phụ thuộc
```

### Chức năng chi tiết các Module lõi:

#### 1. `gui_app.py` (Giao diện & Quản lý luồng)

- Giao diện thân thiện xây dựng bằng PyQt5.
- Cung cấp **Prompt Manager** tích hợp (Chia đôi màn hình Splitter) cho phép tạo, sửa, lưu prompt trực tiếp trên UI.
- Quản lý **Multi-threading (`QThread`)**: Tách biệt luồng gọi AI và luồng UI để ứng dụng không bị "Not Responding". Sử dụng `pythoncom.CoInitialize()` để luồng nền có thể gọi Microsoft Word an toàn.

#### 2. `math_exam_handler.py` (Trái tim của hệ thống xử lý KHTN)

- **Nhận diện cấu trúc hiện đại:** Tự động cắt đề thành các phần dựa theo Regex (PHẦN I: Trắc nghiệm, PHẦN II: Đúng/Sai, PHẦN III: Điền số). Bỏ qua PHẦN IV (Tự luận).
- **Phát hiện Công thức WMF (Equation 3.0):** Khi phát hiện OLE embed WMF, hệ thống ép UI phải chuyển sang `Vision Mode` (Convert DOCX sang PDF và gửi ảnh cho AI đọc) vì text thuần không thể hiển thị được WMF.
- **Chuẩn hóa đáp án (`normalize_math_answer`):** Đồng nhất định dạng số (VD: `2,4` và `2.4` được coi là giống nhau), xử lý đúng sai `ĐSSS` vs `DSSS`.

#### 3. `callAPI.py` (Kết nối Google Vertex AI)

- Sử dụng thư viện `google.genai` mới nhất.
- **Cơ chế Cache File Thông Minh:** Tránh việc upload lại cùng một file đề gốc nhiều lần. Tool sử dụng File API để upload PDF lên GCP, lưu cache `uri` trong session. Nếu tài khoản GCP không hỗ trợ File API, tự động fallback về truyền inline (base64/text).

#### 4. `check_de.py` (Trình kiểm tra cấu trúc & Matching)

- **Parse XML (`lxml`)**: Quét cấu trúc Word để tìm lỗi vỡ Layout (lệch tab/indent), lỗi ngắt trang (`lastRenderedPageBreak`).
- Tính toán và ánh xạ: Sau khi AI trả về file JSON dạng `[{"shuffled_q": 1, "original_q": 15, "option_mapping": {"A":"C"...}}]`, đối chiếu chéo với bảng Excel để kết luận đúng sai.

---

## 3. LUỒNG XỬ LÝ (WORKFLOW)

1. **Khởi tạo (Init):** Giao diện nạp các cấu hình môn học từ thư mục `prompts/`. Người dùng nạp file Đề Gốc, Bảng Đáp Án, và list các Đề Trộn.
2. **Xử lý tiền kỳ (Pre-processing):** Nếu môn học là Toán/KHTN (quy định trong `[AUTO_PDF] = True`), `gui_app.py` gọi `docx2pdf` ngầm để render DOCX thành PDF nhằm bảo toàn hiển thị công thức WMF/MathType.
3. **Phân rã dữ liệu (Parsing):** Hệ thống dùng hàm extract text/OMML để bóc tách đề gốc thành các khối câu hỏi (`MathQuestion` objects).
4. **AI Matching:** Chuyển đổi dữ liệu thành Prompt + File PDF đính kèm gửi lên Gemini 3.1 Pro. AI có nhiệm vụ đọc, dò tìm nội dung và ghép cặp câu hỏi giữa đề trộn và đề gốc. Trả về JSON Option Mapping.
5. **So khớp Đáp án (Verification):** Hệ thống dùng mapping của AI để dò ngược đáp án trong file Excel.
6. **Kiểm tra cấu trúc (Structural Check):** Quét raw docx/pdf để tìm lỗi layout, trùng đáp án, quá số trang.
7. **Báo cáo (Export):** Xuất toàn bộ vào file `KetQua_KiemTra_Môn_TênĐề_ThờiGian.xlsx` thông qua thư viện `openpyxl`.

---

## 4. HƯỚNG DẪN CÀI ĐẶT MÔI TRƯỜNG & CHẠY DỰ ÁN

### Yêu cầu hệ thống:

- Hệ điều hành: Windows (Bắt buộc để thư viện `docx2pdf` có thể dùng COM Interop kết nối với MS Word).
- Đã cài đặt Microsoft Office Word.
- Python 3.10 trở lên.

### Bước 1: Clone và Cài đặt thư viện

Tạo môi trường ảo và cài đặt các thư viện phụ thuộc:

```bash
python -m venv venv
venv\Scripts\activate

# Cài đặt các thư viện cốt lõi
pip install PyQt5 docx2pdf python-docx openpyxl PyMuPDF google-genai google-auth python-dotenv lxml requests pywin32
```

_(Lưu ý: Bạn cần cài đặt đúng `PyQt5` và `docx2pdf` vì chúng là thành phần chạy GUI)._

### Bước 2: Khởi chạy

Chạy ứng dụng từ Terminal:

```bash
python gui_app.py
```

---

## 5. HƯỚNG DẪN DÀNH CHO DEVELOPER (MAINTAINERS)

### 5.1. Quản lý Đóng gói PyInstaller (.exe)

Hệ thống này được thiết kế để compile ra file `.exe` cho End-User.
Trong `gui_app.py`, hàm `get_app_dir()` và `setup_environment()` sẽ xử lý đường dẫn ảo `sys._MEIPASS`.
Khi đóng gói bằng PyInstaller, hãy đảm bảo bao gồm thư mục `prompts` vào file build bằng tham số `--add-data`:

```bash
pyinstaller --noconsole --add-data "prompts;prompts" gui_app.py
```

App sẽ tự động giải nén thư mục `prompts` ra cùng cấp với file `.exe` để người dùng có thể tùy chỉnh.

### 5.2. Chỉnh sửa Prompt Môn Học (Prompt Loader)

Cấu trúc các file `.txt` trong thư mục `prompts/` hoạt động như một hệ cơ sở dữ liệu nhỏ. Mỗi file gồm các TAG:

- `[SUBJECT]`, `[SUBJECT_LABEL]`: Khai báo mã và tên môn.
- `[AUTO_PDF]`: `True`/`False`. Nếu `True`, tool sẽ ép convert docx -> pdf trước khi gửi AI (Cần cho Toán/KHTN).
- `[MAX_PAGES]`: Giới hạn số trang để cảnh báo lỗi in ấn.
- `[PARTS_CONFIG]`: Định nghĩa quy tắc cắt đề (VD: `1|mc|Trắc nghiệm|nhiều phương án,trắc nghiệm`).
- `[PROMPT_HEADER]` và `[VISION_PROMPT]`: Nội dung lệnh System gửi cho AI.

Bạn có thể thêm môn học mới chỉ bằng cách tạo file `prompt_tenmon.txt` (hoặc tạo trực tiếp qua nút **Quản Lý & Sửa Prompt** trên Giao diện).

### 5.3. Xử lý Lỗi Tốc độ AI (Rate Limiting)

Trong `math_exam_handler.py`, ở vòng lặp `max_retries = 3`, luồng xử lý có lệnh `time.sleep(20 * attempt)`. Việc này giúp hệ thống tự động nhả rate limit nếu API của GCP trả về lỗi (429 Too Many Requests hoặc Safety Filter Block). Nếu bạn có tài khoản GCP tier cao hơn, có thể giảm thời gian sleep này xuống.
