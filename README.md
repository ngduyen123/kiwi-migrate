# Kiwi TCMS – skill migrate sheet test case

Skill cho Claude Code để đưa test case từ Excel / CSV / Google Sheet lên Kiwi TCMS.

```
kiwi-testcase-migrate/
├── SKILL.md                     # Hướng dẫn cho Claude (quy trình 6 bước)
├── scripts/
│   ├── migrate.py               # inspect / run (--dry-run)
│   └── requirements.txt
├── references/column-mapping.md # Bảng mapping cột sheet -> trường Kiwi
└── templates/
    ├── mapping.example.json
    ├── .env.example
    └── sample-testcases.csv
```

## Cài đặt

```powershell
git clone https://github.com/ngduyen123/kiwi-migrate.git
pip install -r kiwi-migrate/kiwi-testcase-migrate/scripts/requirements.txt
# Cài skill cho Claude Code (dùng cho mọi project):
Copy-Item -Recurse kiwi-migrate/kiwi-testcase-migrate "$HOME/.claude/skills/"
```

Sau đó trong Claude Code gõ, ví dụ:
- *"migrate testcase D:\QA\login.xlsx lên kiwi plan 30"*
- *"đẩy https://docs.google.com/spreadsheets/d/.../edit#gid=0 lên kiwi plan 30"*

hoặc `/kiwi-testcase-migrate`. Chỉ cần file (hoặc link) và Plan ID – Product được lấy tự động từ plan.

## Dùng script trực tiếp (không qua Claude)

```powershell
copy kiwi-testcase-migrate/templates/.env.example .env   # điền URL, user, password
# (tuỳ chọn) chỉ khi tên cột khác mẫu chuẩn:
copy kiwi-testcase-migrate/templates/mapping.example.json mapping.json   # sửa phần columns
python kiwi-testcase-migrate/scripts/migrate.py inspect --input "D:\QA\login.xlsx"
# Quét file: báo trùng lặp, ô trống, sai tên cột, case sẽ cập nhật/tắt... -> validation-report-*.csv
python kiwi-testcase-migrate/scripts/migrate.py validate --input "D:\QA\login.xlsx" --plan 30
# Đẩy lên Kiwi sau khi đã xem report (case có vấn đề sẽ gắn PROPOSED)
python kiwi-testcase-migrate/scripts/migrate.py run --input "D:\QA\login.xlsx" --plan 30 --accept-issues
```

`--input` nhận đường dẫn file ở bất kỳ đâu hoặc link Google Sheet (chia sẻ "Anyone with the link").
Sheet theo mẫu "Kịch bản kiểm thử" của UpBase: thêm
`--mapping kiwi-testcase-migrate/templates/mapping.kich-ban-kiem-thu.json`. Mẫu này đẩy cả **kết quả test**
(cột "Chrome / Lần 1"…) vào Test Run: thêm `--run <ID>` để ghi vào run đã tạo sẵn, hoặc `--build <id>` để
script tạo run mới; `--no-results` để chỉ đẩy test case.
Với Google Sheet, khi có case mới script tạo `ket-qua/kiwi-id-column.txt` (ghi đè mỗi lần) để bạn dán cột
Kiwi ID vào sheet; nếu cột Kiwi ID trên sheet đã đúng thì không tạo gì. Report nằm trong `ket-qua/`
(giữ 20 bản gần nhất mỗi loại); bản sao lưu file Excel/CSV chỉ giữ 3 bản gần nhất.

Sau lần đẩy đầu, script thêm cột **Kiwi ID** vào file (có bản sao lưu). Sửa file rồi chạy lại 2 lệnh trên
để đồng bộ: case đã có được cập nhật tại chỗ (giữ ID và kết quả test, chuyển NEED_UPDATE), dòng mới được
tạo, case đã xoá khỏi file chuyển DISABLED. Không case nào bị xoá khỏi Kiwi.

`.env`, `mapping.json` và report CSV đã nằm trong `.gitignore`.
