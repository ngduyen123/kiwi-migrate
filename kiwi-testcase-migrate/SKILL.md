---
name: kiwi-testcase-migrate
description: Migrate and keep in sync test cases from a spreadsheet (Excel .xlsx, .csv, or a Google Sheet link) into Kiwi TCMS through its JSON-RPC API. Scans the sheet first and reports duplicates, empty data, wrong column names, invalid values and what will be created / updated / disabled for the user to confirm. First push creates every case (problem cases as PROPOSED with the reason in Notes) and writes the Kiwi case ID back into a "Kiwi ID" column; pushing the edited file again updates those cases in place (same ID, test results kept), creates new rows and disables cases removed from the file. Use when someone says "migrate testcase lên kiwi", "import test case vào Kiwi TCMS", "đẩy sheet testcase lên kiwi", "chuyển testcase từ excel sang kiwi", "đẩy lại file testcase đã sửa", "cập nhật testcase trên kiwi theo file", "kiểm tra file testcase trước khi đẩy kiwi", "upload test cases to Kiwi", "kiwi tcms import". Not for running tests or recording test execution results.
---

# Kiwi TCMS – Migrate & đồng bộ sheet test case

Skill đưa test case từ sheet (xlsx / csv / Google Sheet) lên Kiwi TCMS bằng
`scripts/migrate.py` (Python 3.8+, `requests`, `openpyxl`). Chạy script từ thư mục có file `.env`.

Nguyên tắc:
- **Quét trước – user xác nhận – rồi mới đẩy.**
- **File là nguồn chuẩn.** Đẩy lại file đã sửa thì Kiwi được đồng bộ theo file, nhưng **không xoá case
  nào**: case cũ được sửa tại chỗ (giữ ID nên giữ kết quả test), case bị xoá khỏi file chuyển DISABLED.
- Mọi case đều được đẩy, không bỏ case nào; case có vấn đề gắn status **PROPOSED** và ghi lý do vào Notes.
- **Không bao giờ sửa case do người khác tạo.**

## Quy trình

1. **Lấy input** – chỉ cần 2 thứ:
   - **File test case**: đường dẫn file .xlsx/.csv ở **bất kỳ đâu** trên máy (không cần copy vào thư mục
     dự án; đường dẫn có dấu cách/tiếng Việt thì đặt trong ngoặc kép), hoặc **link Google Sheet**. Link Google
     Sheet phải được chia sẻ "Anyone with the link – Viewer"; link có `#gid=...` sẽ đọc đúng tab đó (script
     tải bản .xlsx để giữ ô gộp và tự tìm tab khớp gid – file lớn có thể mất 1–2 phút). Nếu
     sheet không chia sẻ công khai được, nhờ user tải về (File → Download → .xlsx) và dùng file đó.
     File xlsx nhiều tab → hỏi tên tab (`--sheet`).
   - **Test Plan ID** (lấy từ URL `.../plan/<ID>/`). **Product được lấy tự động từ plan** – không cần sửa
     mapping. Không có plan thì hỏi tên Product và truyền `--product "<tên>"`.
   - Kiểm tra có `.env` với `KIWI_URL`, `KIWI_USERNAME`, `KIWI_PASSWORD` (mẫu: `templates/.env.example`).
     **Không bao giờ tự nhập hay in mật khẩu ra chat** – để user tự điền.
   - Nếu file đang mở trong Excel, nhắc user đóng lại (script cần ghi cột Kiwi ID vào file).

2. **Xem cấu trúc sheet**: `python scripts/migrate.py inspect --input <file> [--sheet <tab>]`

3. **Mapping** (mapping chỉ mô tả *cột nào là trường nào*, không chứa Product/Plan):
   - Sheet theo mẫu **"Kịch bản kiểm thử" của UpBase** (tiêu đề ở dòng 9 gồm "Mã trường hợp kiểm thử",
     "Mục đích kiểm thử", "Các bước thực hiện", "Kết quả mong muốn"; có dòng nhóm "I. Check …") → dùng
     `--mapping <skill>/templates/mapping.kich-ban-kiem-thu.json`. Mapping này: bỏ dòng tiêu đề phụ, lấy dòng
     nhóm làm Category `<tên tab>><nhóm>`, và ghép bước đầu tiên (hoặc kết quả) vào tên các case dùng chung
     một ô "Mục đích kiểm thử" bị gộp.
   - Mẫu khác: script tự dùng `./mapping.json` nếu có, không thì dùng mẫu chuẩn `templates/mapping.example.json`.
   - Không chắc mẫu nào: chạy `inspect` (và nếu tiêu đề không nằm ở dòng 1 thì xem dữ liệu thô) trước.
   Nếu tên cột của file khác mẫu (xem kết quả `inspect`), copy mẫu thành `mapping.json`, sửa `columns`
   (xem `references/column-mapping.md`), cho user xem và xác nhận. Team dùng chung một mẫu sheet thì
   chỉ cần làm một lần.

4. **Quét file – BẮT BUỘC trước khi đẩy**
   `python scripts/migrate.py validate --input "<file hoặc link>" --plan <ID>`
   (có `.env` thì đối chiếu với Kiwi, chỉ đọc; `--offline` để bỏ qua).
   Script in ra dòng tổng: *Tạo mới / Cập nhật / Giữ nguyên / DISABLED*, phần "ĐIỂM CẦN XÁC NHẬN" và ghi
   `ket-qua/validation-report-<thời gian>.csv`. Trình bày cho user theo nhóm, ngắn gọn, kèm số dòng trong sheet:
   - **Cảnh báo cấp file** – cột trong mapping không có trong sheet (có gợi ý tên đúng). Nên sửa mapping
     trước, vì nếu không, cả cột đó sẽ trống ở mọi case.
   - **Case đã xoá khỏi file** – case do user tạo từ file này trước đây, nay không còn → sẽ chuyển
     DISABLED (kết quả test vẫn giữ). Luôn nêu rõ danh sách này cho user.
   - **Sẽ cập nhật** – case nào, đổi gì (Title, Steps…, Priority: Low → High). Case đổi nội dung sẽ chuyển
     **NEED_UPDATE** để QA review lại.
   - **Có thể là case cũ đã đổi tên** – dòng chưa có Kiwi ID nhưng nội dung giống một case cũ. Nếu đúng,
     user điền ID đó vào cột Kiwi ID rồi quét lại; nếu không, chạy tiếp sẽ tạo case mới.
   - **Thiếu dữ liệu bắt buộc** (mặc định Summary, Steps, Expected result).
   - **Trùng trong sheet**: trùng TC ID; trùng tên (không phân biệt hoa thường) – trùng hoàn toàn
     hoặc khác nội dung; có ghi rõ dòng nào khác chữ hoa/thường. Trùng Kiwi ID (dòng bị copy).
   - **Dòng nghi bị gộp nhầm**: dòng không có ID/Title nhưng có Priority/Category khác case phía trên.
   - **Priority / Status không hợp lệ** (sẽ dùng giá trị mặc định).
   - **Trùng với case của người khác** (cùng Product + Category, tên không phân biệt hoa thường):
     không đụng vào case của họ, vẫn tạo case của user (PROPOSED; Notes trên Kiwi có dòng
     "Trùng với case: TC-x <link> (tạo bởi …)"; report có cột `trung_voi_case`).
     `"on_existing_other": "link"` để gắn case của người kia vào plan thay vì tạo mới.
   Sau đó **hỏi user**: sửa file rồi quét lại, hay vẫn đẩy như hiện tại. Không tự quyết thay user.

5. **Đẩy / đồng bộ** – chỉ khi user đồng ý:
   `python scripts/migrate.py run --input "<file hoặc link>" --plan <ID> --accept-issues`
   - Không có `--accept-issues` mà còn điểm cần xác nhận thì script dừng, không ghi gì.
   - Thử trước với `--limit 2` nếu là lần đầu dùng với Product/Kiwi mới (khi có `--limit`, script không
     tắt case đã xoá khỏi file).
   - Sau khi đẩy, **chỉ khi cột `Kiwi ID` trong file còn thiếu/sai** (có case mới, lần đầu đẩy…), script
     ghi cột này vào file (tạo bản sao lưu `<tên>.bak-<thời gian>.<đuôi>` trước, giữ 3 bản gần nhất). Cột
     đã đúng thì in "Cột Kiwi ID trong file đã đúng – không cần ghi lại" và không tạo file nào.
     File có hình ảnh/biểu đồ/macro hoặc đang mở trong Excel → ghi ra file mới `<tên>-kiwi.xlsx`;
     nhắc user dùng file đó cho lần sau. Google Sheet → script tạo/ghi đè `ket-qua/kiwi-id-column.txt`
     (cột Kiwi ID xếp đúng thứ tự dòng, bắt đầu từ dòng tiêu đề); hướng dẫn user copy toàn bộ và dán vào ô
     tiêu đề cột Kiwi ID (hoặc ô tiêu đề cột trống đầu tiên bên phải) trên sheet – cần làm để lần đẩy sau
     nhận diện được case. `--output <path>` để ghi ra file khác, `--no-write-back` để không ghi.
   - Report nằm trong thư mục `ket-qua/` (cạnh nơi chạy script), giữ 20 bản gần nhất mỗi loại.

   - **Kết quả test** (khi mapping có `results.enabled`, như mẫu Kịch bản kiểm thử): mỗi cột kết quả
     **có dữ liệu** (vd. "Chrome / Lần 1") → một Test Run trong plan; cột trống (Lần 2, Lần 3…) bị bỏ qua.
     P → PASSED, F → FAILED, PE → PAUSED; "Kết quả hiện tại"/"Mã lỗi" thành comment ở lần chạy gần nhất.
     - User đã tạo sẵn Test Run trên Kiwi → `--run <ID>` (ID trong URL `.../runs/<ID>/`). Script ghi dấu
       nhận diện vào Notes của run, các lần sau tự tìm lại, không cần `--run`.
     - Chưa có run → script tạo run mới, cần `--build <id>`; tài khoản thường không tạo được Build, nên
       chọn trong danh sách Build có sẵn mà bước quét in ra. Hỏi user chọn Build, không tự chọn.
     - Kiwi **chỉ cho thêm case CONFIRMED vào Test Run**: case PROPOSED / NEED_UPDATE bị bỏ qua, kết quả của
       chúng chưa ghi (bước quét liệt kê trước). Sửa case, chuyển CONFIRMED rồi đẩy lại để ghi nốt.
     - Người test trên Kiwi là tài khoản đang đăng nhập. `--no-results` để chỉ đẩy test case.

6. **Báo cáo**: số case tạo mới / cập nhật / giữ nguyên / DISABLED / lỗi, đường dẫn
   `ket-qua/migrate-report-*.csv`, file đã được ghi Kiwi ID (hoặc file cần dán cho Google Sheet),
   các case PROPOSED / NEED_UPDATE để QA review.

## Cách nhận diện case khi đẩy lại

1. Dòng có **Kiwi ID** → đúng case đó (đổi Title, đánh lại số TC ID, chuyển dòng đều không ảnh hưởng).
2. Dòng chưa có Kiwi ID → case **do user tạo** cùng Product + Category, Title giống (không phân biệt hoa
   thường) → coi là cùng case.
3. Không khớp → tạo mới (có cảnh báo nếu nội dung giống ≥ 80% một case cũ).
4. Case mang dấu `[Nguồn migrate: <tên file>]` trong Notes mà không còn dòng nào trong file → DISABLED.
   Đổi tên file, hoặc chuyển từ file sang Google Sheet (và ngược lại), thì đặt `"source_name"` trong
   mapping bằng định danh cũ (xem dòng `[Nguồn migrate: …]` trong Notes của case) để vẫn nhận ra.
TC ID trong sheet **không** dùng để nhận diện (team hay đánh lại số).

## Ghi chú

- Tài khoản không phải admin thường không đọc được danh sách Priority/Component, không tạo được Tag mới
  và không gỡ được case khỏi plan: dùng `priority_ids` trong mapping (Kiwi UpBase: High = 6, Medium = 7,
  Low = 8); tag/component lỗi chỉ ghi cảnh báo; case bị xoá khỏi file vẫn chuyển DISABLED nhưng còn nằm
  trong plan (report ghi rõ).
- Tuỳ chỉnh trong mapping: `required_fields`, `on_existing`, `on_existing_other`, `status_on_update`,
  `on_removed`, `rename_similarity`, `annotate_notes`, `valid_statuses`, `source_name` – xem
  `references/column-mapping.md`.
- Đổi format nội dung case trong hàm `build_text()`.
- Lỗi SSL với Kiwi tự host (cert tự ký): `"verify_ssl": false` trong mapping.
