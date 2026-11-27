# Mapping cột sheet → trường Kiwi TCMS

| Khoá trong `columns` | Trường Kiwi | Bắt buộc | Tên cột hay gặp trong sheet |
|---|---|---|---|
| `id` | (chỉ dùng để gộp dòng & ghi report, **không** dùng nhận diện case) | Không | TC ID, ID, No, STT, Mã TC |
| `kiwi_id` | ID case trên Kiwi – script **tự thêm và điền** sau mỗi lần đẩy | Không | Kiwi ID (mặc định) |
| `summary` | `summary` | **Có** | Title, Test Case, Tên test case, Mô tả, Summary |
| `precondition` | phần **Precondition** trong `text` | Không | Precondition, Điều kiện tiên quyết, Tiền điều kiện |
| `steps` | phần **Steps** trong `text` | Không | Steps, Test Steps, Các bước thực hiện |
| `expected` | phần **Expected result** trong `text` | Không | Expected Result, Kết quả mong đợi |
| `test_data` | phần **Test data** trong `text` | Không | Test Data, Dữ liệu test |
| `notes` | `notes` | Không | Notes, Ghi chú, Remark |
| `priority` | `priority` (P1..P5) | Không | Priority, Mức độ ưu tiên, Severity |
| `category` | `category` | Không | Module, Feature, Chức năng, Màn hình |
| `status` | `case_status` | Không | Status (PROPOSED / CONFIRMED / DISABLED / NEED_UPDATE) |
| `tags` | tag (phân tách bằng `,` hoặc `;`) | Không | Tags, Label |
| `component` | component | Không | Component |
| `is_automated` | `is_automated` | Không | Automated, Auto (yes/no, x, true, 1) |
| `requirement` | `requirement` | Không | Requirement, Story, Jira |

## Giá trị mặc định (`defaults`)

Dùng khi ô trống hoặc sheet không có cột đó: `category` (Kiwi luôn có `--default--`), `priority`, `status`.
`product` không cần điền: script lấy Product từ `--plan`, hoặc từ `--product`; `defaults.product` chỉ là
phương án cuối khi không truyền cả hai.

## Chuẩn hoá giá trị

- `priority_map`: đổi giá trị trong sheet sang P1..P5, ví dụ `{"High": "P1", "Cao": "P1"}`.
- `status_map`: đổi sang tên status trong Kiwi, ví dụ `{"Done": "CONFIRMED", "Draft": "PROPOSED"}`.
- So khớp không phân biệt hoa thường, đã cắt khoảng trắng.
- `priority_ids`: ID priority trên Kiwi khi tài khoản không đọc được danh sách Priority,
  ví dụ `{"High": 6, "Medium": 7, "Low": 8}` (Kiwi UpBase).

## Cấu trúc sheet

| Khoá | Mặc định | Ý nghĩa |
|---|---|---|
| `header_row` | `1` | Dòng chứa tiêu đề cột (mẫu Kịch bản kiểm thử: `9`) |
| `continuation_rows` | `false` | Dòng không có ID (và Title) được nối vào case phía trên |
| `id_required` | `false` | Mỗi case phải có ID: dòng không có ID mà chỉ có chữ ở cột Title (hoặc Title = Steps = Expected do gộp ô) là **dòng tiêu đề nhóm**; dòng không ID khác là dòng nối tiếp hoặc bị bỏ qua |
| `category_from_section` | `false` | Lấy tiêu đề nhóm (bỏ "I.", "1."…) làm Category khi sheet không có cột Category |
| `category_prefix` | `""` | Tiền tố Category; `{tab}` = tên tab. Ví dụ `"{tab}>"` → `TB Cập nhật_TB Hệ thống>Check Giao diện` |
| `disambiguate_shared_summary` | `false` | Các case liền nhau có cùng Title (ô gộp) được ghép thêm bước đầu tiên, hoặc kết quả mong muốn nếu bước cũng giống, vào Title |

Dòng lặp lại tên cột (tiêu đề gộp 2 dòng) luôn được bỏ qua.
`sub_header_row`: dòng tiêu đề phụ (mẫu Kịch bản kiểm thử: `10`) – tên cột được ghép thành
`"<tiêu đề> / <tiêu đề phụ>"`, vd. `Chrome / Lần 1`; dữ liệu bắt đầu sau dòng này.

## Kết quả test → Test Run (`results`)

```json
"results": {
  "enabled": true,
  "column_pattern": "^.+ / Lần \\d+$",
  "run_name": "{tab} – {column}",
  "status_map": {"P": "PASSED", "F": "FAILED", "PE": "PAUSED"},
  "comment_columns": ["Kết quả hiện tại", "Mã lỗi"]
}
```

- Mỗi cột khớp `column_pattern` **có ít nhất một kết quả** → một Test Run (tên theo `run_name`). Cột trống bỏ qua.
- `status_map`: giá trị trong ô → trạng thái Kiwi (UNTEST, PAUSED, RUNNING, PASSED, FAILED, BLOCKED, ERROR,
  WAIVED). Giá trị lạ → cảnh báo, không đẩy kết quả đó.
- `comment_columns`: ghép thành comment của kết quả ở lần chạy gần nhất của case.
- Dòng lệnh: `--run <ID>` ghi vào run có sẵn; `--build <id>` khi cần tạo run mới; `--no-results` để bỏ qua.

## Kiểm tra dữ liệu & trùng lặp

| Khoá | Mặc định | Ý nghĩa |
|---|---|---|
| `required_fields` | `["summary", "steps", "expected"]` | Trường bắt buộc; trống → cảnh báo + PROPOSED |
| `on_existing` | `"skip"` | Case trùng tên (không phân biệt hoa thường) trong cùng Product + Category đã có trên Kiwi: `skip` bỏ qua, `create` vẫn tạo mới (PROPOSED) |
| `on_existing_other` | `"create"` | Case trùng đã có trên Kiwi nhưng **do người khác tạo**: `create` tạo case mới của bạn trong plan (PROPOSED), `link` gắn case có sẵn của người kia vào plan |
| `annotate_notes` | `true` | Ghi lý do vào Notes của case trên Kiwi |
| `status_on_update` | `"NEED_UPDATE"` | Status của case khi đẩy lại mà nội dung thay đổi; `"keep"` để giữ nguyên status |
| `on_removed` | `"disable"` | Case không còn trong file: `disable` chuyển DISABLED + gỡ khỏi plan, `keep` chỉ báo |
| `rename_similarity` | `0.8` | Ngưỡng giống nội dung để cảnh báo "có thể là case cũ đã đổi tên" |
| `source_name` | tên file (+ `#tab`) | Định danh file trên Kiwi (dòng `[Nguồn migrate: …]` trong Notes). Đặt cố định nếu hay đổi tên file |
| `valid_statuses` | `[]` | Thêm status tuỳ chỉnh ngoài PROPOSED / CONFIRMED / DISABLED / NEED_UPDATE |

Case bị gắn **PROPOSED** khi: thiếu trường bắt buộc, trùng TC ID, trùng tên trong sheet (giống hoặc khác
nội dung), dòng nối tiếp đáng ngờ, Priority/Status không hợp lệ.
