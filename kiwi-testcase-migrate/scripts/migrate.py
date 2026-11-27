#!/usr/bin/env python3
"""Migrate / sync test cases from a spreadsheet (xlsx / csv / Google Sheet) into Kiwi TCMS.

Usage:
    python migrate.py inspect  --input cases.xlsx [--sheet TestCases]
    python migrate.py validate --input cases.xlsx --mapping mapping.json [--offline]
    python migrate.py run      --input cases.xlsx --mapping mapping.json [--plan ID] [--accept-issues]

Flow: validate -> user reviews validation-report-*.csv -> run --accept-issues.

The sheet is the source of truth. Pushing the same file again syncs Kiwi to it:
  * rows that already have a Kiwi case are UPDATED in place (same case id, so test results are kept);
  * new rows are created;
  * cases created from this file earlier but no longer in it are removed from the plan and DISABLED.
After a push the script writes the Kiwi case id back into a "Kiwi ID" column of the file (with a backup),
so rows are recognised even when TC IDs are renumbered, rows are moved or titles change.
Problem cases (empty data, duplicates, invalid values) are created as PROPOSED with the reason in Notes.

Credentials come from env vars KIWI_URL, KIWI_USERNAME, KIWI_PASSWORD (or a .env file in the
current directory). Requires: requests, openpyxl (only for .xlsx).
"""
import argparse
import csv
import difflib
import html
import io
import json
import os
import re
import shutil
import sys
import zipfile
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import requests

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8")


FIELDS = ["id", "kiwi_id", "summary", "precondition", "steps", "expected", "test_data", "priority",
          "category", "status", "tags", "component", "is_automated", "requirement", "notes"]
MULTI_FIELDS = ["precondition", "steps", "expected", "test_data", "notes"]
CONTENT_FIELDS = ["precondition", "test_data", "steps", "expected"]
FIELD_LABELS = {
    "id": "TC ID", "kiwi_id": "Kiwi ID", "summary": "Summary/Title", "precondition": "Precondition",
    "test_data": "Test data", "steps": "Steps", "expected": "Expected result",
    "priority": "Priority", "category": "Category", "status": "Status", "tags": "Tags",
    "component": "Component", "is_automated": "Automated", "requirement": "Requirement",
    "notes": "Notes",
}
DEFAULT_REQUIRED = ["summary", "steps", "expected"]
DEFAULT_KIWI_ID_HEADER = "Kiwi ID"
KIWI_STATUSES = ["PROPOSED", "CONFIRMED", "DISABLED", "NEED_UPDATE"]
EMPTY_MARK = "_(trống)_"
TS = datetime.now().strftime("%Y%m%d-%H%M%S")
DEFAULT_RESULT_PATTERN = r"^.+ / Lần \d+$"   # "<browser/env> / Lần N" (two-level header)
DEFAULT_RESULT_STATUS = {"p": "PASSED", "pass": "PASSED", "passed": "PASSED", "đạt": "PASSED",
                         "f": "FAILED", "fail": "FAILED", "failed": "FAILED", "không đạt": "FAILED",
                         "pe": "PAUSED", "pending": "PAUSED", "b": "BLOCKED", "blocked": "BLOCKED",
                         "n/a": "WAIVED", "na": "WAIVED"}
OUT_DIR = "ket-qua"
KEEP_REPORTS = 20


def out_path(name: str) -> str:
    """Reports and the paste file go to one folder; only the newest KEEP_REPORTS of each kind are kept."""
    os.makedirs(OUT_DIR, exist_ok=True)
    m = re.match(r"(.+-report)-\d{8}-\d{6}\.csv$", name)
    if m:
        old = sorted(f for f in os.listdir(OUT_DIR) if re.match(re.escape(m.group(1)) + r"-\d{8}-\d{6}\.csv$", f))
        for f in old[: max(0, len(old) - (KEEP_REPORTS - 1))]:
            try:
                os.remove(os.path.join(OUT_DIR, f))
            except OSError:
                pass
    return os.path.join(OUT_DIR, name)


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #
def load_dotenv(path: str = ".env") -> None:
    if not os.path.isfile(path):
        return
    with open(path, encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def norm(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def norm_key(value: str) -> str:
    """Case-insensitive, whitespace-insensitive comparison key (HTML entities decoded)."""
    return re.sub(r"\s+", " ", html.unescape(value or "")).strip().lower()


def issue(level: str, code: str, message: str, proposed: bool = False) -> Dict[str, Any]:
    return {"level": level, "code": code, "message": message, "proposed": proposed}


def map_value(value: str, table: Optional[Dict[str, str]]) -> str:
    if not value:
        return value
    for k, v in (table or {}).items():
        if k.strip().lower() == value.strip().lower():
            return v
    return value


def truthy(value: str) -> bool:
    return value.strip().lower() in ("1", "x", "y", "yes", "true", "auto", "automated", "có", "co")


def split_tags(value: str) -> List[str]:
    return [t.strip() for t in re.split(r"[,;\n]", value or "") if t.strip()]


def as_list(parts: List[str], numbered: bool) -> str:
    if not parts:
        return ""
    if len(parts) == 1 or not numbered:
        return "\n".join(parts)
    return "\n".join("%d. %s" % (i, p) for i, p in enumerate(parts, start=1))


def rows_label(rows: List[str]) -> str:
    return ", ".join(rows)


def header_for(mapping: Dict[str, Any], field: str) -> Optional[str]:
    cols = mapping.get("columns", {})
    if field == "kiwi_id":
        return cols.get("kiwi_id") or DEFAULT_KIWI_ID_HEADER
    return cols.get(field)


def source_name(args: argparse.Namespace, mapping: Dict[str, Any]) -> str:
    """Identifies 'this file' on Kiwi, so cases removed from the file can be found later."""
    if mapping.get("source_name"):
        return mapping["source_name"]
    m = GSHEET_RE.match(args.input)
    if m:
        gid = re.search(r"[#&?]gid=(\d+)", args.input)
        return "gsheet:" + m.group(1) + ("#gid=" + gid.group(1) if gid else "")
    name = os.path.basename(args.input)
    sheet = args.sheet or mapping.get("sheet")
    if sheet and os.path.splitext(name)[1].lower() in (".xlsx", ".xlsm"):
        name += "#" + sheet
    return name


# --------------------------------------------------------------------------- #
# Reading the sheet
# --------------------------------------------------------------------------- #
GSHEET_RE = re.compile(r"https://docs\.google\.com/spreadsheets/d/([A-Za-z0-9_-]+)")


def _gs_get(url: str) -> requests.Response:
    try:
        resp = requests.get(url, timeout=180)
    except requests.RequestException as exc:
        sys.exit("Không kết nối được Google Sheet: %s" % exc)
    if resp.status_code != 200 or "text/html" in resp.headers.get("Content-Type", ""):
        sys.exit("Không tải được Google Sheet (HTTP %s). Kiểm tra: link đúng chưa, sheet đã chia sẻ "
                 "'Anyone with the link – Viewer' chưa. Nếu không chia sẻ công khai được, tải về bằng "
                 "File → Download → .xlsx rồi dùng đường dẫn file đó." % resp.status_code)
    return resp


def _ws_matrix(ws: Any) -> List[List[Any]]:
    """Worksheet values with merged cells filled with their top-left value."""
    merged_values = {}
    for rng in ws.merged_cells.ranges:
        top_left = ws.cell(rng.min_row, rng.min_col).value
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                merged_values[(r, c)] = top_left
    matrix = []
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        matrix.append([merged_values.get((r_idx, c_idx), v) for c_idx, v in enumerate(row, start=1)])
    return matrix


def _match_score(ws: Any, matrix: List[List[str]], n: int = 40) -> float:
    """Share of non-empty CSV cells (first n rows) that equal the worksheet cell."""
    total = hits = 0
    ws_rows = list(ws.iter_rows(min_row=1, max_row=n, values_only=True))
    for r_idx, crow in enumerate(matrix[:n]):
        wrow = ws_rows[r_idx] if r_idx < len(ws_rows) else ()
        for c_idx, b in enumerate(crow):
            b = b.strip()
            if not b:
                continue
            total += 1
            if c_idx < len(wrow) and norm(wrow[c_idx]) == b:
                hits += 1
    return hits / total if total else 0.0


def read_rows(src: str, sheet: Optional[str], header_row: int, sub_header_row: Optional[int] = None) -> Tuple[List[str], List[Dict[str, str]], str]:
    """Return (headers, rows, tab title). Each row is {header: value} plus '_row' (1-based sheet row)."""
    m = GSHEET_RE.match(src)
    if m:
        gid = re.search(r"[#&?]gid=(\d+)", src)
        base = "https://docs.google.com/spreadsheets/d/%s/export?format=" % m.group(1)
        csv_matrix = list(csv.reader(io.StringIO(
            _gs_get(base + "csv" + ("&gid=" + gid.group(1) if gid else "")).content.decode("utf-8-sig"))))
        # The CSV export loses merged cells, so read the xlsx export (whole workbook) and pick the
        # tab that matches the gid's CSV; fall back to the CSV if that fails.
        try:
            import openpyxl
            print("Đang tải bản .xlsx của Google Sheet để giữ ô gộp...")
            wb = openpyxl.load_workbook(io.BytesIO(_gs_get(base + "xlsx").content), data_only=True)
            if sheet and sheet in wb.sheetnames:
                ws = wb[sheet]
            else:
                scored = sorted(((_match_score(w, csv_matrix), w) for w in wb.worksheets), key=lambda x: -x[0])
                ws = scored[0][1] if scored and scored[0][0] >= 0.8 else None
            if ws is not None:
                print("Tab: %s" % ws.title)
                return _rows_from_matrix(_ws_matrix(ws), header_row, sub_header_row) + (ws.title,)
            print("! Không xác định được tab trong bản .xlsx – đọc bản CSV (ô gộp sẽ bị mất)")
        except Exception as exc:  # pragma: no cover - network / parsing issues
            print("! Không đọc được bản .xlsx (%s) – đọc bản CSV (ô gộp sẽ bị mất)" % exc)
        return _rows_from_matrix(csv_matrix, header_row, sub_header_row) + ("",)

    ext = os.path.splitext(src)[1].lower()
    if ext == ".csv":
        with open(src, encoding="utf-8-sig", newline="") as fh:
            return _rows_from_matrix(list(csv.reader(fh)), header_row, sub_header_row) + (
                os.path.splitext(os.path.basename(src))[0],)
    if ext in (".xlsx", ".xlsm"):
        try:
            import openpyxl
        except ImportError:
            sys.exit("Thiếu thư viện openpyxl: pip install openpyxl")
        wb = openpyxl.load_workbook(src, data_only=True, read_only=False)
        ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.worksheets[0]
        if sheet and sheet not in wb.sheetnames:
            print("! Không thấy tab '%s', dùng tab đầu tiên '%s'. Các tab: %s"
                  % (sheet, ws.title, ", ".join(wb.sheetnames)))
        return _rows_from_matrix(_ws_matrix(ws), header_row, sub_header_row) + (ws.title,)
    sys.exit("Định dạng không hỗ trợ: %s (dùng .xlsx, .csv hoặc link Google Sheet)" % ext)


def _rows_from_matrix(matrix: List[List[Any]], header_row: int,
                      sub_header_row: Optional[int] = None) -> Tuple[List[str], List[Dict[str, str]]]:
    if len(matrix) < header_row:
        return [], []
    headers = [norm(h) for h in matrix[header_row - 1]]
    first_data = header_row
    if sub_header_row and len(matrix) >= sub_header_row:
        # Two-level header, e.g. "Chrome" merged over sub-headers "Lần 1 / Lần 2 / Lần 3".
        subs = [norm(s) for s in matrix[sub_header_row - 1]]
        headers = ["%s / %s" % (h, s) if s and s != h else h
                   for h, s in zip(headers + [""] * (len(subs) - len(headers)), subs + [""] * len(headers))]
        first_data = max(header_row, sub_header_row)
    rows = []
    for offset, raw in enumerate(matrix[first_data:], start=first_data + 1):
        values = [norm(v) for v in raw]
        if not any(values):
            continue
        row = {h: (values[i] if i < len(values) else "") for i, h in enumerate(headers) if h}
        row["_row"] = str(offset)
        rows.append(row)
    return [h for h in headers if h], rows


def lookup_column(row: Dict[str, str], header: Optional[str]) -> str:
    if not header:
        return ""
    if header in row:
        return row[header]
    lowered = header.lower()
    for key, value in row.items():
        if key.lower() == lowered:
            return value
    return ""


# --------------------------------------------------------------------------- #
# Rows -> cases -> prepared items (with issues)
# --------------------------------------------------------------------------- #
SECTION_PREFIX_RE = re.compile(r"^\s*([IVXLC]+|\d+|[A-Za-z])[\.\)]\s*")


def first_line(parts: List[str], limit: int = 80) -> str:
    text = parts[0].strip().splitlines()[0] if parts and parts[0].strip() else ""
    text = re.sub(r"^\s*\d+[\.\)]\s*", "", text).strip()
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


def build_cases(rows: List[Dict[str, str]], mapping: Dict[str, Any]) -> List[Dict[str, Any]]:
    continuation = mapping.get("continuation_rows", False)
    id_required = mapping.get("id_required", False)
    prefix = (mapping.get("category_prefix") or "").replace("{tab}", mapping.get("_tab", ""))
    section = ""
    cases: List[Dict[str, Any]] = []
    for row in rows:
        # Sub-header rows (e.g. a header merged over two rows) repeat the column names.
        if sum(1 for h, v in row.items() if h != "_row" and v and v.strip().lower() == h.strip().lower()) >= 2:
            continue
        data = {f: lookup_column(row, header_for(mapping, f)) for f in FIELDS}
        cur = cases[-1] if cases else None

        if id_required and not data["id"]:
            s = data["summary"]
            if s and data["steps"] in ("", s) and data["expected"] in ("", s):
                # Group header such as "I. Check Giao diện" (often merged across Title/Steps/Expected).
                section = SECTION_PREFIX_RE.sub("", s).strip()
                continue
            if not cur or not continuation or not any(data[f] for f in MULTI_FIELDS):
                continue  # stray row (totals, notes outside the table...)

        is_cont = False
        if continuation and cur:
            blank_key = not data["id"] if id_required else (not data["id"] and not data["summary"])
            # Merged ID/Title cells repeat the same value on every row of the case.
            merged_key = bool(data["id"]) and data["id"] == cur["id"] and data["summary"] in ("", cur["summary"])
            is_cont = blank_key or merged_key
        if is_cont:
            for f in MULTI_FIELDS:
                if data[f] and not (cur[f] and cur[f][-1] == data[f] and f in ("precondition", "test_data")):
                    cur[f].append(data[f])
            extra = [f for f in FIELDS if f not in MULTI_FIELDS and f != "id"
                     and data[f] and data[f] != cur[f]]
            cur["_cont"].append({"row": row["_row"], "extra": extra})
            continue
        case = {f: data[f] for f in FIELDS if f not in MULTI_FIELDS}
        for f in MULTI_FIELDS:
            case[f] = [data[f]] if data[f] else []
        if not case["category"] and section and mapping.get("category_from_section"):
            case["category"] = prefix + section
        case["_row"] = row["_row"]
        case["_cont"] = []
        case["_notes"] = []
        rcfg = mapping.get("results") or {}
        if rcfg.get("enabled"):
            pattern = re.compile(rcfg.get("column_pattern", DEFAULT_RESULT_PATTERN), re.I)
            case["_results"] = {h: v.strip() for h, v in row.items() if h != "_row" and v.strip() and pattern.match(h)}
            case["_rcomment"] = ["%s: %s" % (h, lookup_column(row, h)) for h in rcfg.get("comment_columns", [])
                                 if lookup_column(row, h)]
        cases.append(case)

    if mapping.get("disambiguate_shared_summary"):
        disambiguate(cases)
    return cases


def disambiguate(cases: List[Dict[str, Any]]) -> None:
    """Consecutive cases sharing one (merged) purpose cell get their first step / expected result
    appended to the title, so they do not all look like duplicates of each other."""
    i = 0
    while i < len(cases):
        j = i
        key = norm_key(cases[i]["summary"])
        while key and j + 1 < len(cases) and norm_key(cases[j + 1]["summary"]) == key:
            j += 1
        run = cases[i: j + 1]
        if len(run) > 1:
            steps = [first_line(c["steps"]) for c in run]
            use = steps if len(set(steps)) > 1 else [first_line(c["expected"]) for c in run]
            rows = [c["_row"] for c in run]
            for c, suffix in zip(run, use):
                if not suffix:
                    suffix = first_line(c["expected"]) or first_line(c["steps"])
                if suffix:
                    others = ", ".join(r for r in rows if r != c["_row"])
                    c["summary"] = "%s – %s" % (c["summary"], suffix)
                    c["_notes"].append("Tên case được ghép thêm '%s' vì dùng chung Mục đích kiểm thử với dòng %s"
                                       % (suffix, others))
        i = j + 1


SECTION_TITLES = {"precondition": "Precondition", "test_data": "Test data",
                  "steps": "Steps", "expected": "Expected result"}


def build_text(case: Dict[str, Any], mapping: Dict[str, Any]) -> str:
    """Markdown body. Every mapped content column is shown, empty ones as '(trống)'."""
    cols = mapping.get("columns", {})
    parts = []
    for f in CONTENT_FIELDS:
        if not cols.get(f):
            continue
        body = as_list(case[f], numbered=f in ("steps", "expected")) or EMPTY_MARK
        parts.append("**%s**\n\n%s" % (SECTION_TITLES[f], body))
    return "\n\n".join(parts)


def text_sections(text: str) -> Dict[str, str]:
    parts = re.split(r"\*\*(.+?)\*\*\s*\n", html.unescape(text or ""))
    return {parts[i].strip(): norm_key(parts[i + 1]) for i in range(1, len(parts) - 1, 2)}


def check_columns(headers: List[str], mapping: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], set]:
    cols = mapping.get("columns", {})
    lower_headers = {h.lower() for h in headers}
    issues, missing = [], set()
    for field, header in cols.items():
        if not header or header.lower() in lower_headers or field == "kiwi_id":
            continue
        missing.add(field)
        msg = "Cột '%s' (%s) không có trong sheet → toàn bộ case sẽ trống %s" % (
            header, FIELD_LABELS.get(field, field), FIELD_LABELS.get(field, field))
        sugg = difflib.get_close_matches(header, headers, n=1, cutoff=0.5)
        if sugg:
            msg += ". Có phải cột '%s'? Sửa trong mapping.json" % sugg[0]
        issues.append(issue("WARN", "MISSING_COLUMN", msg))
    used = {h.lower() for h in cols.values() if h} | {header_for(mapping, "kiwi_id").lower()}
    rcfg = mapping.get("results") or {}
    if rcfg.get("enabled"):
        pattern = re.compile(rcfg.get("column_pattern", DEFAULT_RESULT_PATTERN), re.I)
        used |= {h.lower() for h in headers if pattern.match(h)}
        used |= {h.lower() for h in rcfg.get("comment_columns", [])}
    unused = [h for h in headers if h.lower() not in used]
    if unused:
        issues.append(issue("INFO", "UNUSED_COLUMN",
                            "Các cột sau không được migrate: " + ", ".join(unused)))
    return issues, missing


def prepare(cases: List[Dict[str, Any]], mapping: Dict[str, Any], missing: set) -> List[Dict[str, Any]]:
    cols = mapping.get("columns", {})
    defaults = mapping["defaults"]
    required = mapping.get("required_fields", DEFAULT_REQUIRED)
    priority_ids = {k.lower() for k in (mapping.get("priority_ids") or {})}
    valid_status = {s.lower(): s for s in KIWI_STATUSES + list(mapping.get("valid_statuses", []))}

    items = []
    for c in cases:
        iss: List[Dict[str, Any]] = [issue("INFO", "SUMMARY_EXTENDED", n) for n in c.get("_notes", [])]
        label = c["id"] or ("dòng " + c["_row"])

        # --- Kiwi ID written back by a previous push
        kiwi_id = None
        if c["kiwi_id"]:
            m = re.search(r"\d+", c["kiwi_id"])
            if m:
                kiwi_id = int(m.group())
            else:
                iss.append(issue("WARN", "KIWI_ID_INVALID", "Kiwi ID '%s' không hợp lệ → bỏ qua, coi như case mới"
                                 % c["kiwi_id"]))

        # --- empty data
        for f in required:
            empty = not any(c[f]) if f in MULTI_FIELDS else not c[f]
            if empty:
                why = " (cột không có trong sheet)" if f in missing else ""
                iss.append(issue("WARN", "EMPTY_REQUIRED", "Trống %s%s" % (FIELD_LABELS[f], why), True))
        for f in CONTENT_FIELDS:
            if cols.get(f) and f not in required and not c[f]:
                iss.append(issue("INFO", "EMPTY_OPTIONAL", "Trống %s" % FIELD_LABELS[f]))

        summary = c["summary"] or "[Thiếu summary] %s" % label
        if len(summary) > 255:
            iss.append(issue("WARN", "SUMMARY_TOO_LONG", "Summary dài %d ký tự, bị cắt còn 255" % len(summary)))
            summary = summary[:255]

        # --- category / priority / status
        category = c["category"] or defaults.get("category", "--default--")
        if not c["category"] and cols.get("category"):
            iss.append(issue("INFO", "DEFAULT_USED", "Trống Category → dùng '%s'" % category))

        priority = map_value(c["priority"], mapping.get("priority_map"))
        default_p = defaults.get("priority", "P3")
        if not priority:
            priority = default_p
            if cols.get("priority"):
                iss.append(issue("INFO", "DEFAULT_USED", "Trống Priority → dùng '%s'" % priority))
        elif priority.lower() not in priority_ids and not re.match(r"^P[1-5]$", priority, re.I):
            iss.append(issue("WARN", "BAD_PRIORITY", "Priority '%s' không hợp lệ → dùng '%s'"
                             % (c["priority"], default_p), True))
            priority = default_p

        status = map_value(c["status"], mapping.get("status_map"))
        default_s = defaults.get("status", "CONFIRMED")
        if not status:
            status = default_s
        elif status.lower() not in valid_status:
            iss.append(issue("WARN", "BAD_STATUS", "Status '%s' không hợp lệ" % c["status"], True))
            status = default_s
        else:
            status = valid_status[status.lower()]

        # --- continuation rows
        if c["_cont"]:
            iss.append(issue("INFO", "MERGED_ROWS", "Gộp dòng %s vào case này (dòng nối tiếp)"
                             % rows_label([x["row"] for x in c["_cont"]])))
            for x in c["_cont"]:
                if x["extra"]:
                    names = ", ".join(FIELD_LABELS[f] for f in x["extra"])
                    iss.append(issue(
                        "WARN", "SUSPICIOUS_ROW",
                        "Dòng %s không có ID/Title nhưng có %s khác case này → đang bị gộp vào case này; nếu là "
                        "case mới thì cần điền Title (giá trị %s của dòng đó không được dùng)"
                        % (x["row"], names, names), True))

        items.append({
            "row": c["_row"], "id": c["id"], "kiwi_id": kiwi_id, "sheet_kiwi_id": kiwi_id,
            "raw_summary": c["summary"],
            "summary": summary, "category": category, "priority": priority, "status": status,
            "text": build_text(c, mapping), "notes": "\n".join(c["notes"]),
            "tags": split_tags(c["tags"]) + list(mapping.get("extra_tags", [])),
            "components": split_tags(c["component"]),
            "is_automated": truthy(c["is_automated"]), "requirement": c["requirement"],
            "sig": "|".join(norm_key("\n".join(c[f])) for f in CONTENT_FIELDS),
            "issues": iss,
            "existing": None,   # Kiwi case to leave as is (unchanged / someone else's, linked)
            "match": None,      # own Kiwi case this row corresponds to (candidate for update)
            "update": None,     # computed changes for that case
            "dup_of": None,     # someone else's case with the same title
            "results": c.get("_results", {}),          # {"Chrome / Lần 1": "P", ...}
            "result_comment": "\n".join(c.get("_rcomment", [])),
        })

    find_duplicates(items)
    return items


def find_duplicates(items: List[Dict[str, Any]]) -> None:
    by_id: Dict[str, List[Dict[str, Any]]] = {}
    by_summary: Dict[str, List[Dict[str, Any]]] = {}
    for it in items:
        if it["id"]:
            by_id.setdefault(norm_key(it["id"]), []).append(it)
        if it["raw_summary"]:
            by_summary.setdefault(norm_key(it["raw_summary"]), []).append(it)

    for group in by_id.values():
        if len(group) > 1:
            for it in group:
                others = [o["row"] for o in group if o is not it]
                it["issues"].append(issue("WARN", "DUP_ID", "Trùng TC ID '%s' với dòng %s"
                                          % (it["id"], rows_label(others)), True))

    for group in by_summary.values():
        if len(group) < 2:
            continue
        for it in group:
            others = [o for o in group if o is not it]
            exact = [o for o in others if o["sig"] == it["sig"]]
            diff = [o for o in others if o["sig"] != it["sig"]]

            def case_note(sub: List[Dict[str, Any]]) -> str:
                variants = [o for o in sub if o["raw_summary"] != it["raw_summary"]]
                if not variants:
                    return ""
                return " (khác chữ hoa/thường: %s)" % "; ".join(
                    "dòng %s ghi '%s'" % (o["row"], o["raw_summary"]) for o in variants)

            if exact:
                it["issues"].append(issue("WARN", "DUP_EXACT", "Trùng hoàn toàn (tên + nội dung) với dòng %s%s"
                                          % (rows_label([o["row"] for o in exact]), case_note(exact)), True))
            if diff:
                it["issues"].append(issue("WARN", "DUP_SUMMARY", "Trùng tên với dòng %s nhưng nội dung khác%s"
                                          % (rows_label([o["row"] for o in diff]), case_note(diff)), True))


def has_proposed(it: Dict[str, Any]) -> bool:
    return any(i["proposed"] for i in it["issues"])


def finalize(items: List[Dict[str, Any]]) -> None:
    for it in items:
        it["final_status"] = "PROPOSED" if has_proposed(it) else it["status"]


def build_notes(it: Dict[str, Any], source: str, mapping: Dict[str, Any]) -> str:
    parts = []
    if it["notes"]:
        parts.append(it["notes"])
    if mapping.get("annotate_notes", True):
        msgs = [i["message"] for i in it["issues"] if i["proposed"] and i["code"] != "EXISTS_OTHER"]
        if msgs:
            parts.append("[Migrate – vấn đề] " + "; ".join(msgs))
    if it["dup_of"]:
        d = it["dup_of"]
        parts.append("Trùng với case: TC-%s %s%s" % (d["id"], d["url"],
                                                    " (tạo bởi %s)" % d["author"] if d["author"] else ""))
    if source:
        parts.append("[Nguồn migrate: %s]" % source)
    return "\n\n".join(parts)


# --------------------------------------------------------------------------- #
# Kiwi JSON-RPC client
# --------------------------------------------------------------------------- #
class KiwiError(Exception):
    pass


class Kiwi:
    def __init__(self, url: str, username: str, password: str, verify_ssl: bool = True):
        self.base = url.rstrip("/")
        for suffix in ("/json-rpc", "/xml-rpc"):
            if self.base.endswith(suffix):
                self.base = self.base[: -len(suffix)]
        self.endpoint = self.base + "/json-rpc/"
        self.session = requests.Session()
        self.session.verify = verify_ssl
        self.username = username
        self._id = 0
        self.call("Auth.login", username, password)

    def call(self, method: str, *params: Any) -> Any:
        self._id += 1
        resp = self.session.post(
            self.endpoint,
            json={"jsonrpc": "2.0", "method": method, "params": list(params), "id": self._id},
            headers={"Referer": self.base + "/"},
            timeout=120,
        )
        resp.raise_for_status()
        body = resp.json()
        if body.get("error"):
            raise KiwiError("%s: %s" % (method, body["error"].get("message", body["error"])))
        return body.get("result")

    def case_url(self, case_id: int) -> str:
        return "%s/case/%s/" % (self.base, case_id)


def connect(mapping: Dict[str, Any], required: bool) -> Optional[Kiwi]:
    load_dotenv()
    url, user, pwd = (os.environ.get(k) for k in ("KIWI_URL", "KIWI_USERNAME", "KIWI_PASSWORD"))
    if not (url and user and pwd):
        if required:
            sys.exit("Thiếu KIWI_URL / KIWI_USERNAME / KIWI_PASSWORD (đặt trong biến môi trường hoặc file .env)")
        print("! Không có thông tin đăng nhập Kiwi – bỏ qua bước kiểm tra với Kiwi")
        return None
    return Kiwi(url, user, pwd, verify_ssl=mapping.get("verify_ssl", True))


class Resolver:
    """Name -> id lookups and access to the product's existing cases."""

    def __init__(self, kiwi: Kiwi, mapping: Dict[str, Any], product_id: int):
        self.kiwi = kiwi
        self.mapping = mapping
        self.product_id = product_id
        self.categories = {html.unescape(c["name"]).lower(): c
                           for c in kiwi.call("Category.filter", {"product": self.product_id})}
        me = kiwi.call("User.filter", {"username": kiwi.username})
        self.my_id = me[0]["id"] if me else None
        self.status_cache: Dict[str, int] = {}
        self.priority_cache: Dict[str, int] = {}
        self.disabled_id = self.status("DISABLED")

    def cases_in_categories(self, category_names: List[str]) -> List[Dict[str, Any]]:
        found = []
        for name in sorted(set(n.lower() for n in category_names)):
            cat = self.categories.get(name)
            if cat:
                found.extend(self.kiwi.call("TestCase.filter", {"category": cat["id"]}))
        return found

    def category(self, name: str) -> int:
        c = self.categories.get(name.lower())
        if not c:
            if not self.mapping.get("create_missing_category", True):
                raise KiwiError("Không có category '%s'" % name)
            c = self.kiwi.call("Category.create", {"product": self.product_id, "name": name})
            self.categories[name.lower()] = c
            print("  + Đã tạo category '%s'" % name)
        return c["id"]

    def priority(self, value: str) -> int:
        ids = {k.lower(): v for k, v in (self.mapping.get("priority_ids") or {}).items()}
        if value.lower() in ids:
            return int(ids[value.lower()])
        if value not in self.priority_cache:
            try:
                res = self.kiwi.call("Priority.filter", {"value": value})
            except KiwiError:
                # Non-admin accounts often cannot read Priority; Kiwi's fixture has P1..P5 = 1..5.
                m = re.match(r"^P([1-5])$", value.strip(), re.I)
                if not m:
                    raise
                res = [{"id": int(m.group(1))}]
            if not res:
                raise KiwiError("Priority '%s' không tồn tại trên Kiwi" % value)
            self.priority_cache[value] = res[0]["id"]
        return self.priority_cache[value]

    def status(self, name: str) -> int:
        if name not in self.status_cache:
            res = self.kiwi.call("TestCaseStatus.filter", {"name": name})
            if not res:
                raise KiwiError("Case status '%s' không tồn tại trên Kiwi" % name)
            self.status_cache[name] = res[0]["id"]
        return self.status_cache[name]


def check_kiwi(resolver: Resolver, items: List[Dict[str, Any]], mapping: Dict[str, Any],
               source: str, full_file: bool) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """Match sheet rows to Kiwi cases. Returns (global issues, cases to disable)."""
    kiwi, me = resolver.kiwi, resolver.my_id
    global_issues: List[Dict[str, Any]] = []
    on_existing = mapping.get("on_existing", "skip")

    seen_new = set()
    for it in items:
        key = it["category"].lower()
        if key not in resolver.categories and key not in seen_new:
            seen_new.add(key)
            global_issues.append(issue("INFO", "NEW_CATEGORY", "Category '%s' chưa có trên Kiwi → sẽ tạo mới"
                                       % it["category"]))

    def describe(ex: Dict[str, Any], it: Dict[str, Any]) -> str:
        mine = ex.get("author") == me
        who = "do bạn tạo" if mine else "do %s tạo" % (ex.get("author__username") or "người khác")
        what = "nội dung giống" if norm_key(ex.get("text")) == norm_key(it["text"]) else "nội dung khác sheet"
        return "TC-%s %s (%s, %s)" % (ex["id"], kiwi.case_url(ex["id"]), who, what)

    # Existing cases in the categories used by the sheet. Own cases first, active only.
    cases = resolver.cases_in_categories([it["category"] for it in items])
    pool: Dict[Tuple[int, str], List[Dict[str, Any]]] = {}
    for c in sorted(cases, key=lambda x: (x.get("author") != me, x["id"])):
        if c.get("case_status") == resolver.disabled_id:
            continue
        pool.setdefault((c["category"], norm_key(c["summary"])), []).append(c)
    used = set()

    def pool_key(it: Dict[str, Any]) -> Optional[Tuple[int, str]]:
        cat = resolver.categories.get(it["category"].lower())
        return (cat["id"], norm_key(it["summary"])) if cat else None

    def note_other_duplicate(it: Dict[str, Any]) -> None:
        key = pool_key(it)
        others = [c for c in pool.get(key, []) if c.get("author") != me] if key else []
        if others:
            ex = others[0]
            it["dup_of"] = {"id": ex["id"], "url": kiwi.case_url(ex["id"]),
                            "author": ex.get("author__username") or ""}
            it["issues"].append(issue("WARN", "EXISTS_OTHER", "Trùng tên với case của người khác: %s"
                                      % describe(ex, it), True))

    # 1) Rows that carry a Kiwi ID from an earlier push.
    ids = [it["kiwi_id"] for it in items if it["kiwi_id"]]
    by_id = {c["id"]: c for c in kiwi.call("TestCase.filter", {"pk__in": ids})} if ids else {}
    for it in items:
        kid = it["kiwi_id"]
        if not kid:
            continue
        ex = by_id.get(kid)
        if kid in used:
            it["issues"].append(issue("WARN", "DUP_KIWI_ID", "Kiwi ID TC-%s trùng với dòng khác (dòng copy?) → "
                                      "dòng này được tạo thành case mới" % kid))
        elif not ex or ex.get("category") not in {c["id"] for c in resolver.categories.values()}:
            it["issues"].append(issue("WARN", "KIWI_ID_NOT_FOUND", "Kiwi ID TC-%s không có trong Product này → "
                                      "tạo case mới" % kid))
        elif ex.get("author") != me:
            it["issues"].append(issue("WARN", "KIWI_ID_OTHER", "Kiwi ID TC-%s là case do %s tạo → không sửa, "
                                      "tạo case mới của bạn" % (kid, ex.get("author__username") or "người khác")))
        else:
            used.add(kid)
            it["match"] = ex
            note_other_duplicate(it)
            continue
        it["kiwi_id"] = None

    # 2) Rows without a (usable) Kiwi ID: same title in the same product + category.
    for it in items:
        if it["match"] or it["kiwi_id"]:
            continue
        key = pool_key(it)
        candidates = [c for c in pool.get(key, []) if c["id"] not in used] if key else []
        if not candidates:
            continue
        ex = candidates[0]
        mine = ex.get("author") == me
        if on_existing == "create":
            it["dup_of"] = {"id": ex["id"], "url": kiwi.case_url(ex["id"]), "author": ex.get("author__username") or ""}
            it["issues"].append(issue("WARN", "EXISTS_ON_KIWI", "Trùng tên với %s → vẫn tạo mới"
                                      % describe(ex, it), True))
        elif mine:
            used.add(ex["id"])
            it["match"] = ex
            note_other_duplicate(it)
        elif mapping.get("on_existing_other", "create") == "link":
            used.add(ex["id"])
            it["existing"] = {"id": ex["id"], "link": True}
            it["issues"].append(issue("WARN", "EXISTS_OTHER", "Trùng tên với %s → không tạo mới, gắn case có sẵn "
                                      "đó vào plan" % describe(ex, it)))
        else:
            # Someone else's case: never touch it, but the sheet case still lands in our plan,
            # with a link to their case in the report and in Notes.
            note_other_duplicate(it)

    # 3) Still unmatched: maybe an own case whose title was changed in the sheet.
    leftovers = [c for lst in pool.values() for c in lst if c.get("author") == me and c["id"] not in used]
    threshold = float(mapping.get("rename_similarity", 0.8))
    for it in items:
        if it["match"] or it["existing"] or not leftovers:
            continue
        best, best_ratio = None, 0.0
        a = norm_key(it["text"])
        for c in leftovers:
            sm = difflib.SequenceMatcher(None, a, norm_key(c.get("text")))
            if sm.quick_ratio() < threshold:
                continue
            ratio = sm.ratio()
            if ratio > best_ratio:
                best, best_ratio = c, ratio
        if best and best_ratio >= threshold:
            it["issues"].append(issue(
                "WARN", "MAYBE_RENAMED",
                "Nội dung giống %d%% TC-%s '%s' %s → có thể là case cũ đã đổi tên. Nếu đúng, điền TC-%s vào cột "
                "Kiwi ID rồi quét lại; nếu chạy tiếp sẽ tạo case mới" % (
                    round(best_ratio * 100), best["id"], html.unescape(best["summary"]),
                    kiwi.case_url(best["id"]), best["id"])))

    # 4) Own cases created from this file earlier that are no longer in it.
    to_disable: List[Dict[str, Any]] = []
    if source and full_file:
        matched = used | {it["existing"]["id"] for it in items if it["existing"]}
        marker = "[Nguồn migrate: %s]" % source
        for c in kiwi.call("TestCase.filter", {"category__product": resolver.product_id, "notes__contains": marker}):
            if c.get("author") == me and c.get("case_status") != resolver.disabled_id and c["id"] not in matched:
                to_disable.append(c)
        action = ("gỡ khỏi plan và chuyển DISABLED (kết quả test vẫn giữ)"
                  if mapping.get("on_removed", "disable") == "disable" else "chỉ báo, không thay đổi")
        for c in to_disable:
            global_issues.append(issue("WARN", "REMOVED", "TC-%s '%s' %s không còn trong file → %s"
                                       % (c["id"], html.unescape(c["summary"]), kiwi.case_url(c["id"]), action)))
    elif source:
        global_issues.append(issue("INFO", "REMOVED_SKIPPED",
                                   "Đang chạy với --limit nên không kiểm tra case đã bị xoá khỏi file"))
    return global_issues, to_disable


def compute_updates(resolver: Resolver, items: List[Dict[str, Any]], mapping: Dict[str, Any], source: str) -> None:
    """For rows matched to an own Kiwi case, work out which fields differ from the sheet."""
    for it in items:
        ex = it["match"]
        if not ex:
            continue
        changes, values, content = [], {}, False
        if html.unescape(ex["summary"]) != it["summary"]:
            changes.append("Title: '%s' → '%s'" % (html.unescape(ex["summary"]), it["summary"]))
            values["summary"] = it["summary"]
            content = True
        if norm_key(ex.get("text")) != norm_key(it["text"]):
            old, new = text_sections(ex.get("text")), text_sections(it["text"])
            names = [t for t in list(new) + [t for t in old if t not in new] if old.get(t) != new.get(t)]
            changes.append("%s thay đổi" % (", ".join(names) if names else "Nội dung"))
            values["text"] = it["text"]
            content = True
        if ex.get("priority") != resolver.priority(it["priority"]):
            changes.append("Priority: %s → %s" % (ex.get("priority__value"), it["priority"]))
            values["priority"] = it["priority"]
            content = True
        if html.unescape(ex.get("category__name") or "").lower() != it["category"].lower():
            changes.append("Category: %s → %s" % (html.unescape(ex.get("category__name") or ""), it["category"]))
            values["category"] = it["category"]
            content = True
        if bool(ex.get("is_automated")) != it["is_automated"]:
            changes.append("Automated: %s → %s" % (bool(ex.get("is_automated")), it["is_automated"]))
            values["is_automated"] = it["is_automated"]
            content = True
        if (ex.get("requirement") or "") != (it["requirement"] or ""):
            changes.append("Requirement: '%s' → '%s'" % (ex.get("requirement") or "", it["requirement"]))
            values["requirement"] = it["requirement"]
            content = True
        notes = build_notes(it, source, mapping)
        if (ex.get("notes") or "").strip() != notes.strip():
            changes.append("Notes")
            values["notes"] = notes

        if not values:
            it["existing"] = {"id": ex["id"], "link": True}
            it["issues"].append(issue("INFO", "UNCHANGED", "Không thay đổi so với TC-%s" % ex["id"]))
        else:
            it["update"] = {"id": ex["id"], "values": values, "changes": changes, "content": content}
            level = "WARN" if content else "INFO"
            it["issues"].append(issue(level, "UPDATE", "Sẽ cập nhật TC-%s %s: %s"
                                      % (ex["id"], resolver.kiwi.case_url(ex["id"]), "; ".join(changes))))


# --------------------------------------------------------------------------- #
# Test results -> Test Runs
# --------------------------------------------------------------------------- #
def plan_results(items: List[Dict[str, Any]], mapping: Dict[str, Any], source: str) -> Dict[str, Dict[str, Any]]:
    """Group the result cells by result column; each column becomes one Test Run."""
    rcfg = mapping.get("results") or {}
    if not rcfg.get("enabled"):
        return {}
    smap = dict(DEFAULT_RESULT_STATUS)
    smap.update({k.strip().lower(): v.upper() for k, v in (rcfg.get("status_map") or {}).items()})
    template = rcfg.get("run_name", "{tab} – {column}")
    tab = mapping.get("_tab") or source
    runs: Dict[str, Dict[str, Any]] = {}
    for it in items:
        cols = list(it["results"])
        for col in cols:
            val = it["results"][col]
            status = smap.get(val.lower())
            if not status:
                it["issues"].append(issue("WARN", "BAD_RESULT", "Kết quả '%s' ở cột '%s' không hiểu → không đẩy kết "
                                          "quả này (khai báo trong results.status_map)" % (val, col)))
                continue
            run = runs.setdefault(col, {
                "name": template.format(tab=tab, column=col.replace(" / ", " – "))[:255],
                "marker": "[Nguồn migrate: %s | %s]" % (source, col), "entries": []})
            # "Kết quả hiện tại" / "Mã lỗi" belong to the latest round only.
            run["entries"].append((it, status, it["result_comment"] if col == cols[-1] else ""))
            after = status_after_push(it, mapping)
            if after and after.upper() != "CONFIRMED":
                run.setdefault("unconfirmed", []).append((it, after))
    return runs


def status_after_push(it: Dict[str, Any], mapping: Dict[str, Any]) -> str:
    """Case status on Kiwi once this push is done (only CONFIRMED cases can be added to a Test Run)."""
    if it["update"] and it["update"]["content"]:
        status = "PROPOSED" if has_proposed(it) else mapping.get("status_on_update", "NEED_UPDATE")
        if status.lower() != "keep":
            return status
    if it["match"]:
        return it["match"].get("case_status__name") or ""
    return it.get("final_status", "")


def known_builds(kiwi: Kiwi, product_id: int) -> Dict[int, Tuple[str, str]]:
    """Builds already used by Test Runs of the product (normal accounts cannot list Build directly)."""
    builds: Dict[int, Tuple[str, str]] = {}
    for r in kiwi.call("TestRun.filter", {"plan__product": product_id}):
        builds[r["build"]] = (html.unescape(r.get("build__name") or ""), r.get("build__version__value") or "")
    return builds


def builds_text(builds: Dict[int, Tuple[str, str]]) -> str:
    return "\n".join("    --build %-4s → %s (version %s)" % (bid, n, v) for bid, (n, v) in sorted(builds.items()))


def resolve_build(kiwi: Kiwi, product_id: int, build_arg: Optional[str]) -> int:
    builds = known_builds(kiwi, product_id)
    if not build_arg:
        raise KiwiError("Cần chọn Build cho Test Run mới – thêm --build <id>. Build đang có:\n" + builds_text(builds))
    if str(build_arg).isdigit():
        return int(build_arg)
    hits = [bid for bid, (n, _) in builds.items() if n.lower() == build_arg.strip().lower()]
    if len(hits) == 1:
        return hits[0]
    if hits:
        raise KiwiError("Có %d Build cùng tên '%s' – dùng --build <id>:\n%s"
                        % (len(hits), build_arg, builds_text({b: builds[b] for b in hits})))
    raise KiwiError("Không thấy Build '%s' (tài khoản không tạo được Build mới) – chọn một Build có sẵn:\n%s"
                    % (build_arg, builds_text(builds)))


def run_url(kiwi: Kiwi, run_id: int) -> str:
    return "%s/runs/%s/" % (kiwi.base, run_id)


def find_run(kiwi: Kiwi, plan_id: int, run: Dict[str, Any], run_arg: Optional[int],
             n_runs: int) -> Optional[Dict[str, Any]]:
    """The Test Run for one result column: --run if given, else the run carrying this column's marker."""
    if run_arg:
        if n_runs != 1:
            raise KiwiError("Sheet có %d cột kết quả có dữ liệu – --run chỉ dùng được khi có đúng 1 cột" % n_runs)
        res = kiwi.call("TestRun.filter", {"pk": int(run_arg)})
        if not res:
            raise KiwiError("Không tìm thấy Test Run %s" % run_arg)
        if res[0]["plan"] != int(plan_id):
            raise KiwiError("Test Run %s thuộc plan %s, không phải plan %s" % (run_arg, res[0]["plan"], plan_id))
        return res[0]
    found = kiwi.call("TestRun.filter", {"plan": int(plan_id), "notes__contains": run["marker"]})
    return found[0] if found else None


def print_results_plan(runs: Dict[str, Dict[str, Any]], kiwi: Optional[Kiwi], plan_id: Optional[int],
                       product_id: Optional[int], build_arg: Optional[str], run_arg: Optional[int] = None) -> int:
    """Show what will happen to test results. Returns the number of blocking problems."""
    if not runs:
        return 0
    print("\n=== KẾT QUẢ TEST → TEST RUN%s ===" % (" (plan %s)" % plan_id if plan_id else ""))
    need_build = False
    for col, r in runs.items():
        counts: Dict[str, int] = {}
        for _, st, _ in r["entries"]:
            counts[st] = counts.get(st, 0) + 1
        action = "tạo Test Run mới"
        if kiwi and plan_id:
            try:
                found = find_run(kiwi, int(plan_id), r, run_arg, len(runs))
            except KiwiError as exc:
                print("  ! %s" % exc)
                return 1
            if found:
                action = "ghi vào Test Run TR-%s '%s' %s" % (found["id"], html.unescape(found["summary"]),
                                                           run_url(kiwi, found["id"]))
            else:
                need_build = True
        comments = sum(1 for _, _, c in r["entries"] if c)
        print("  - '%s': %d kết quả (%s)%s → %s" % (
            r["name"], len(r["entries"]), ", ".join("%s %d" % kv for kv in counts.items()),
            "; %d comment" % comments if comments else "", action))
        if r.get("unconfirmed"):
            print("    ! %d case không phải CONFIRMED trên Kiwi nên Kiwi không cho thêm vào Test Run → chưa ghi kết quả: %s"
                  % (len(r["unconfirmed"]), "; ".join("dòng %s %s (%s)" % (it["row"], it["id"], s)
                                                      for it, s in r["unconfirmed"])))
    if not plan_id:
        print("  ! Cần --plan để đẩy kết quả test.")
        return 1
    if need_build and kiwi and product_id:
        try:
            bid = resolve_build(kiwi, product_id, build_arg)
            print("  Build cho Test Run mới: id %s" % bid)
        except KiwiError as exc:
            print("  ! %s" % exc)
            return 1
    return 0


def push_results(kiwi: Kiwi, resolver: Resolver, runs: Dict[str, Dict[str, Any]], plan_id: int,
                 build_arg: Optional[str], written: Dict[str, str], run_arg: Optional[int] = None) -> None:
    statuses = {s["name"].upper(): s["id"] for s in kiwi.call("TestExecutionStatus.filter", {})}
    print("\n=== ĐẨY KẾT QUẢ TEST ===")
    for col, r in runs.items():
        found = find_run(kiwi, plan_id, r, run_arg, len(runs))
        if found:
            run = found
            notes = run.get("notes") or ""
            if r["marker"] not in notes:
                # Remember the link sheet column -> run, so later pushes find it without --run.
                try:
                    kiwi.call("TestRun.update", run["id"], {"notes": (notes + "\n" if notes else "") + r["marker"]})
                except KiwiError as exc:
                    print("  ! Không ghi được dấu nhận diện vào Notes của TR-%s (%s) – lần sau nhớ truyền lại "
                          "--run %s" % (run["id"], exc, run["id"]))
        else:
            run = kiwi.call("TestRun.create", {
                "summary": r["name"], "plan": int(plan_id), "manager": resolver.my_id,
                "build": resolve_build(kiwi, resolver.product_id, build_arg),
                "notes": "Tạo từ sheet test case (kiwi-testcase-migrate).\n%s" % r["marker"]})
            print("  + Đã tạo Test Run TR-%s '%s'" % (run["id"], r["name"]))
        rid = run["id"]
        entries = [(int(written[it["row"]][3:]), st, cm) for it, st, cm in r["entries"] if it["row"] in written]
        execs = {e["case"]: e for e in kiwi.call("TestExecution.filter", {"run": rid})}
        missing = [cid for cid, _, _ in entries if cid not in execs]
        added, skipped = 0, {}
        for cid in missing:
            try:
                kiwi.call("TestRun.add_case", rid, cid)
                added += 1
            except KiwiError as exc:
                # Kiwi only accepts CONFIRMED cases in a run.
                skipped[cid] = "chưa CONFIRMED" if "not confirmed" in str(exc).lower() else str(exc)
        if missing:
            execs = {e["case"]: e for e in kiwi.call("TestExecution.filter", {"run": rid})}
        n_set = n_same = n_err = n_comment = 0
        for cid, st, cm in entries:
            if cid in skipped:
                print("  - bỏ qua TC-%s (%s) – kết quả %s chưa ghi" % (cid, skipped[cid], st))
                continue
            try:
                e = execs.get(cid)
                if not e:
                    raise KiwiError("TC-%s không có trong Test Run" % cid)
                if statuses.get(st) is None:
                    raise KiwiError("Trạng thái %s không có trên Kiwi" % st)
                if e["status"] != statuses[st]:
                    kiwi.call("TestExecution.update", e["id"], {"status": statuses[st]})
                    n_set += 1
                else:
                    n_same += 1
                if cm:
                    old = kiwi.call("TestExecution.get_comments", e["id"]) or []
                    if not any((c.get("comment") or "").strip() == cm.strip() for c in old):
                        kiwi.call("TestExecution.add_comment", e["id"], cm)
                        n_comment += 1
            except (KiwiError, requests.RequestException) as exc:
                n_err += 1
                print("  ! TC-%s: %s" % (cid, exc))
        print("  TR-%s '%s' %s: thêm %d case, ghi %d kết quả, %d không đổi, %d comment, %d bỏ qua (chưa CONFIRMED), "
              "%d lỗi" % (rid, html.unescape(run.get("summary") or r["name"]), run_url(kiwi, rid), added, n_set, n_same, n_comment, len(skipped), n_err))
        if skipped:
            print("  → Sửa các case bị bỏ qua (thường do thiếu dữ liệu nên đang PROPOSED), chuyển CONFIRMED rồi đẩy lại "
                  "để ghi nốt kết quả.")


# --------------------------------------------------------------------------- #
# Writing the Kiwi ID back into the file
# --------------------------------------------------------------------------- #
RISKY_XLSX_PARTS = re.compile(r"xl/(drawings/drawing\d+\.xml|charts/|media/|pivotTables/|pivotCache/|slicers/|"
                              r"timelines/|vbaProject\.bin)")


def prune_backups(base: str, ext: str, keep: int = 3) -> None:
    """Keep only the newest `keep` '<file>.bak-<timestamp><ext>' copies next to the test case file."""
    folder, stem = os.path.split(base)
    pattern = re.compile(re.escape(stem) + r"\.bak-\d{8}-\d{6}" + re.escape(ext) + "$", re.I)
    old = sorted(f for f in os.listdir(folder or ".") if pattern.match(f))
    for f in old[: max(0, len(old) - keep)]:
        try:
            os.remove(os.path.join(folder, f))
        except OSError:
            pass


def write_back(src: str, mapping: Dict[str, Any], sheet: Optional[str], results: Dict[str, str],
               output: Optional[str]) -> str:
    """Put each row's Kiwi case id into the 'Kiwi ID' column. Returns a human readable status line."""
    if not results:
        return "Không có gì để ghi vào file."
    header = header_for(mapping, "kiwi_id")
    header_row = mapping.get("header_row", 1)
    if GSHEET_RE.match(src):
        # The sheet cannot be written through the public export link: produce a ready-to-paste column,
        # one line per sheet row starting at the header row.
        path = os.path.abspath(out_path("kiwi-id-column.txt"))
        last = max(int(r) for r in results)
        lines = [header] + [results.get(str(r), "") for r in range(header_row + 1, last + 1)]
        with open(path, "w", encoding="utf-8") as fh:
            fh.write("\n".join(lines) + "\n")
        return ("Google Sheet: không ghi thẳng vào sheet được. Mở file %s, copy toàn bộ rồi dán vào ô tiêu đề cột "
                "'%s' (dòng %d) trên sheet – nếu chưa có cột này thì dán vào ô tiêu đề của cột trống đầu tiên bên "
                "phải. Lần đẩy sau script sẽ dùng cột này để nhận diện case." % (path, header, header_row))
    base, ext = os.path.splitext(src)
    ext = ext.lower()
    target = output or src
    notes = []

    if ext == ".csv":
        with open(src, encoding="utf-8-sig", newline="") as fh:
            matrix = list(csv.reader(fh))
        hdr = matrix[header_row - 1]
        lower = [h.strip().lower() for h in hdr]
        col = lower.index(header.lower()) if header.lower() in lower else len(hdr)
        if col == len(hdr):
            hdr.append(header)
        for r, val in results.items():
            row = matrix[int(r) - 1]
            row.extend([""] * (col + 1 - len(row)))
            row[col] = val
        if target == src:
            backup = "%s.bak-%s%s" % (base, TS, ext)
            shutil.copy2(src, backup)
            prune_backups(base, ext)
            notes.append("bản sao lưu: %s" % os.path.abspath(backup))
        try:
            with open(target, "w", encoding="utf-8-sig", newline="") as fh:
                csv.writer(fh).writerows(matrix)
        except PermissionError:
            target = "%s-kiwi-%s%s" % (base, TS, ext)
            with open(target, "w", encoding="utf-8-sig", newline="") as fh:
                csv.writer(fh).writerows(matrix)
            notes.append("file gốc đang mở nên đã ghi ra file mới")
        return "Đã ghi cột '%s' vào %s (%s)" % (header, os.path.abspath(target), "; ".join(notes))

    if ext not in (".xlsx", ".xlsm"):
        return "Không ghi ngược được định dạng %s" % ext
    import openpyxl
    from copy import copy
    from openpyxl.cell.cell import MergedCell

    if target == src:
        with zipfile.ZipFile(src) as z:
            risky = sorted({m.group(1).split("/")[0] for n in z.namelist() for m in [RISKY_XLSX_PARTS.match(n)] if m})
        if risky:
            target = "%s-kiwi%s" % (base, ext)
            notes.append("file có hình ảnh/biểu đồ/macro (%s) nên KHÔNG ghi đè, ghi ra file mới" % ", ".join(risky))
    wb = openpyxl.load_workbook(src, keep_vba=(ext == ".xlsm"))  # keep formulas, not data_only
    ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.worksheets[0]
    col = None
    for cell in ws[header_row]:
        if cell.value is not None and str(cell.value).strip().lower() == header.lower():
            col = cell.column
    if col is None:
        col = ws.max_column + 1
        new = ws.cell(row=header_row, column=col, value=header)
        prev = ws.cell(row=header_row, column=col - 1)
        if prev.has_style:
            new._style = copy(prev._style)
    for r, val in results.items():
        cell = ws.cell(row=int(r), column=col)
        if not isinstance(cell, MergedCell):
            cell.value = val
    if target == src:
        backup = "%s.bak-%s%s" % (base, TS, ext)
        shutil.copy2(src, backup)
        prune_backups(base, ext)
        notes.append("bản sao lưu: %s" % os.path.abspath(backup))
    try:
        wb.save(target)
    except PermissionError:
        target = "%s-kiwi-%s%s" % (base, TS, ext)
        wb.save(target)
        notes.append("file gốc đang mở trong Excel nên đã ghi ra file mới – hãy dùng file này cho lần sau")
    return "Đã ghi cột '%s' vào %s (%s)" % (header, os.path.abspath(target), "; ".join(notes))


# --------------------------------------------------------------------------- #
# Reporting
# --------------------------------------------------------------------------- #
ISSUE_TITLES = {
    "EMPTY_REQUIRED": "Thiếu dữ liệu bắt buộc", "DUP_EXACT": "Trùng hoàn toàn trong sheet",
    "DUP_SUMMARY": "Trùng tên, khác nội dung", "DUP_ID": "Trùng TC ID",
    "SUSPICIOUS_ROW": "Dòng nghi bị gộp nhầm", "BAD_PRIORITY": "Priority không hợp lệ",
    "BAD_STATUS": "Status không hợp lệ", "EXISTS_ON_KIWI": "Đã có trên Kiwi",
    "EXISTS_OTHER": "Trùng với case của người khác trên Kiwi", "SUMMARY_TOO_LONG": "Summary quá dài",
    "UPDATE": "Sẽ CẬP NHẬT case trên Kiwi (giữ nguyên ID và kết quả test)",
    "MAYBE_RENAMED": "Có thể là case cũ đã đổi tên", "DUP_KIWI_ID": "Trùng Kiwi ID",
    "KIWI_ID_NOT_FOUND": "Kiwi ID không tồn tại", "KIWI_ID_OTHER": "Kiwi ID là case của người khác",
    "KIWI_ID_INVALID": "Kiwi ID sai định dạng",
}


def action_of(it: Dict[str, Any]) -> str:
    if it["update"]:
        return "cập nhật TC-%s" % it["update"]["id"]
    if it["existing"]:
        return "giữ nguyên TC-%s" % it["existing"]["id"]
    return "tạo mới"


def print_summary(items: List[Dict[str, Any]], global_issues: List[Dict[str, Any]], rows_count: int,
                  to_disable: List[Dict[str, Any]]) -> int:
    to_create = [it for it in items if not it["existing"] and not it["update"]]
    to_update = [it for it in items if it["update"]]
    unchanged = [it for it in items if it["existing"]]
    proposed = [it for it in to_create if has_proposed(it)]
    print("\n=== KẾT QUẢ QUÉT ===")
    print("Đọc %d dòng → %d test case" % (rows_count, len(items)))
    print("Tạo mới: %d (trong đó %d gắn PROPOSED do có vấn đề) | Cập nhật: %d | Giữ nguyên: %d | "
          "Bỏ khỏi plan (DISABLED): %d" % (len(to_create), len(proposed), len(to_update), len(unchanged),
                                         len(to_disable)))

    warn_count = 0
    file_warns = [g for g in global_issues if g["level"] == "WARN" and g["code"] != "REMOVED"]
    if file_warns:
        print("\n!! CẢNH BÁO CẤP FILE:")
        for g in file_warns:
            print("  - " + g["message"])
        warn_count += len(file_warns)

    removed = [g for g in global_issues if g["code"] == "REMOVED"]
    if removed:
        print("\n!! CASE ĐÃ XOÁ KHỎI FILE (%d):" % len(removed))
        for g in removed:
            print("  - " + g["message"])
        warn_count += len(removed)

    groups: Dict[str, List[str]] = {}
    for it in items:
        for i in it["issues"]:
            if i["level"] == "WARN":
                groups.setdefault(i["code"], []).append(
                    "dòng %s%s: %s" % (it["row"], " (%s)" % it["id"] if it["id"] else "", i["message"]))
    if groups:
        print("\n?? ĐIỂM CẦN XÁC NHẬN:")
        for code, lines in groups.items():
            print("\n  [%s] %d mục" % (ISSUE_TITLES.get(code, code), len(lines)))
            for line in lines[:15]:
                print("    - " + line)
            if len(lines) > 15:
                print("    ... và %d mục khác (xem file report)" % (len(lines) - 15))
            warn_count += len(lines)

    for g in global_issues:
        if g["level"] == "INFO":
            print("\n  i " + g["message"])
    infos: Dict[str, int] = {}
    for it in items:
        for i in it["issues"]:
            if i["level"] == "INFO":
                infos[i["code"]] = infos.get(i["code"], 0) + 1
    if infos:
        info_titles = {"EMPTY_OPTIONAL": "case trống Precondition/Test data",
                       "DEFAULT_USED": "ô dùng giá trị mặc định", "MERGED_ROWS": "case gộp nhiều dòng",
                       "UNCHANGED": "case không thay đổi", "UPDATE": "case chỉ cập nhật Notes",
                       "SUMMARY_EXTENDED": "case được ghép thêm bước/kết quả vào tên (dùng chung Mục đích)"}
        print("\n  Thông tin: " + "; ".join("%d %s" % (n, info_titles.get(c, c)) for c, n in infos.items()))
    if not warn_count:
        print("\nKhông có điểm nào cần xác nhận.")
    return warn_count


def write_validation_report(items: List[Dict[str, Any]], global_issues: List[Dict[str, Any]]) -> str:
    path = out_path("validation-report-%s.csv" % TS)
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["sheet_row", "tc_id", "summary", "level", "issue", "message", "status_tren_kiwi", "hanh_dong",
                    "trung_voi_case"])
        for g in global_issues:
            w.writerow(["", "", "", g["level"], g["code"], g["message"], "", "", ""])
        for it in items:
            dup = it["dup_of"]["url"] if it["dup_of"] else ""
            rows = it["issues"] or [issue("OK", "", "")]
            for i in rows:
                w.writerow([it["row"], it["id"], it["summary"], i["level"], i["code"], i["message"],
                            it["final_status"], action_of(it), dup])
    return os.path.abspath(path)


def find_mapping(path: Optional[str]) -> str:
    """--mapping if given, else ./mapping.json, else the skill's standard template."""
    if path:
        return path
    if os.path.isfile("mapping.json"):
        return "mapping.json"
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "templates", "mapping.example.json")


def resolve_product(kiwi: Kiwi, args: argparse.Namespace, mapping: Dict[str, Any]) -> Tuple[int, str, str]:
    """Product from --product, else from the Test Plan, else from mapping.defaults.product."""
    name = getattr(args, "product", None)
    plan = getattr(args, "plan", None) or mapping.get("test_plan_id")
    plan_row = None
    if plan:
        res = kiwi.call("TestPlan.filter", {"pk": int(plan)})
        if not res:
            raise KiwiError("Không tìm thấy Test Plan %s (hoặc tài khoản không có quyền xem)" % plan)
        plan_row = res[0]
    if name:
        res = kiwi.call("Product.filter", {"name": name})
        if not res:
            raise KiwiError("Không tìm thấy Product '%s' trong Kiwi" % name)
        if plan_row and plan_row["product"] != res[0]["id"]:
            raise KiwiError("Test Plan %s thuộc Product khác, không phải '%s'" % (plan, name))
        return res[0]["id"], res[0]["name"], "--product"
    if plan_row:
        prod = kiwi.call("Product.filter", {"pk": plan_row["product"]})[0]
        return prod["id"], prod["name"], "lấy từ Test Plan %s '%s'" % (plan, html.unescape(plan_row["name"]))
    name = mapping.get("defaults", {}).get("product")
    if name:
        res = kiwi.call("Product.filter", {"name": name})
        if not res:
            raise KiwiError("Không tìm thấy Product '%s' (defaults.product trong mapping)" % name)
        return res[0]["id"], res[0]["name"], "mapping.json"
    sys.exit("Cần cho biết Test Plan (--plan <ID>) hoặc Product (--product \"<tên>\")")


def load_items(args: argparse.Namespace) -> Tuple[Dict[str, Any], List[Dict[str, Any]], List[Dict[str, Any]], int]:
    args.mapping = find_mapping(args.mapping)
    with open(args.mapping, encoding="utf-8") as fh:
        mapping = json.load(fh)
    mapping.setdefault("defaults", {})
    print("Mapping: %s" % os.path.abspath(args.mapping))
    headers, rows, tab = read_rows(args.input, args.sheet or mapping.get("sheet"), mapping.get("header_row", 1),
                                 mapping.get("sub_header_row"))
    mapping["_tab"] = tab
    if not rows:
        sys.exit("Sheet trống hoặc header_row sai.")
    global_issues, missing = check_columns(headers, mapping)
    items = prepare(build_cases(rows, mapping), mapping, missing)
    if getattr(args, "limit", None):
        items = items[: args.limit]
    return mapping, items, global_issues, len(rows)


def analyse(args: argparse.Namespace, mapping: Dict[str, Any], items: List[Dict[str, Any]],
            global_issues: List[Dict[str, Any]], kiwi: Kiwi) -> Tuple[Resolver, List[Dict[str, Any]], str]:
    product_id, product_name, how = resolve_product(kiwi, args, mapping)
    print("Product: %s (%s)" % (html.unescape(product_name), how))
    resolver = Resolver(kiwi, mapping, product_id)
    source = source_name(args, mapping)
    extra, to_disable = check_kiwi(resolver, items, mapping, source, full_file=not args.limit)
    global_issues += extra
    compute_updates(resolver, items, mapping, source)
    finalize(items)
    return resolver, to_disable, source


def preview(items: List[Dict[str, Any]], n: int) -> None:
    for it in items[:n]:
        print("\n--- dòng %s %s" % (it["row"], it["id"]))
        print("summary : %s" % it["summary"])
        print("category: %s | priority: %s | status: %s | auto: %s"
              % (it["category"], it["priority"], it["final_status"], it["is_automated"]))
        print("tags    : %s | components: %s" % (it["tags"], it["components"]))
        print(it["text"])


# --------------------------------------------------------------------------- #
# Commands
# --------------------------------------------------------------------------- #
def cmd_inspect(args: argparse.Namespace) -> None:
    headers, rows, _ = read_rows(args.input, args.sheet, args.header_row)
    if not rows:
        sys.exit("Sheet trống hoặc header_row sai.")
    print("Số dòng dữ liệu: %d" % len(rows))
    print("Các cột:")
    for h in headers:
        filled = sum(1 for r in rows if r.get(h))
        print("  - %-30s (%d ô có dữ liệu)" % (h, filled))
    print("\n%d dòng đầu:" % min(3, len(rows)))
    for r in rows[:3]:
        print(json.dumps(r, ensure_ascii=False, indent=2))


def cmd_validate(args: argparse.Namespace) -> None:
    mapping, items, global_issues, rows_count = load_items(args)
    kiwi = None if args.offline else connect(mapping, required=False)
    to_disable: List[Dict[str, Any]] = []
    product_id = None
    if kiwi:
        resolver, to_disable, _ = analyse(args, mapping, items, global_issues, kiwi)
        product_id = resolver.product_id
    else:
        finalize(items)
    runs = plan_results(items, mapping, source_name(args, mapping))
    if args.preview:
        preview(items, args.preview)
    print_summary(items, global_issues, rows_count, to_disable)
    print_results_plan(runs, kiwi, args.plan or mapping.get("test_plan_id"), product_id, args.build, args.run)
    print("\nReport chi tiết: %s" % write_validation_report(items, global_issues))


def cmd_run(args: argparse.Namespace) -> None:
    mapping, items, global_issues, rows_count = load_items(args)

    if args.dry_run:
        finalize(items)
        print("=== DRY RUN – không ghi gì lên Kiwi, không kiểm tra với Kiwi ===")
        preview(items, args.preview)
        print_summary(items, global_issues, rows_count, [])
        return

    kiwi = connect(mapping, required=True)
    resolver, to_disable, source = analyse(args, mapping, items, global_issues, kiwi)
    plan_id = args.plan or mapping.get("test_plan_id")
    runs = {} if args.no_results else plan_results(items, mapping, source)
    warn_count = print_summary(items, global_issues, rows_count, to_disable)
    results_blocked = print_results_plan(runs, kiwi, plan_id, resolver.product_id, args.build, args.run)
    if warn_count and not args.accept_issues:
        print("\nReport chi tiết: %s" % write_validation_report(items, global_issues))
        sys.exit("\nDỪNG: còn %d điểm cần xác nhận. Sửa sheet rồi chạy lại, hoặc thêm --accept-issues "
                 "để vẫn chạy (case có vấn đề sẽ gắn PROPOSED)." % warn_count)
    if results_blocked:
        sys.exit("\nDỪNG: chưa đủ thông tin để đẩy kết quả test (xem phía trên). Thêm --build <id>, hoặc "
                 "--no-results để chỉ đẩy test case. Chưa ghi gì lên Kiwi.")

    status_on_update = mapping.get("status_on_update", "NEED_UPDATE")
    report_path = out_path("migrate-report-%s.csv" % TS)
    counts = {"created": 0, "updated": 0, "unchanged": 0, "disabled": 0, "error": 0}
    written: Dict[str, str] = {}
    print("\n=== BẮT ĐẦU ĐỒNG BỘ ===")
    with open(report_path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["sheet_row", "sheet_id", "summary", "kiwi_case_id", "url", "result",
                         "status_tren_kiwi", "thay_doi", "van_de", "canh_bao_khac", "trung_voi_case"])

        def add_extras(case_id: int, it: Dict[str, Any]) -> List[str]:
            warnings = []
            if plan_id:
                kiwi.call("TestPlan.add_case", int(plan_id), case_id)
            # Tags/components are best effort: a missing tag needs "add tag" permission
            # to be created, and a component must already exist in the product.
            for tag in it["tags"]:
                try:
                    kiwi.call("TestCase.add_tag", case_id, tag)
                except KiwiError as exc:
                    warnings.append("tag '%s': %s" % (tag, exc))
            for comp in it["components"]:
                try:
                    kiwi.call("TestCase.add_component", case_id, comp)
                except KiwiError as exc:
                    warnings.append("component '%s': %s" % (comp, exc))
            return warnings

        for n, it in enumerate(items, start=1):
            prefix = "[%d/%d] dòng %s" % (n, len(items), it["row"])
            problems = "; ".join(i["message"] for i in it["issues"]
                                 if i["level"] == "WARN" and i["code"] != "UPDATE")
            dup_url = it["dup_of"]["url"] if it["dup_of"] else ""
            try:
                if it["existing"]:
                    ex = it["existing"]
                    if plan_id and ex["link"]:
                        kiwi.call("TestPlan.add_case", int(plan_id), ex["id"])
                    counts["unchanged"] += 1
                    written[it["row"]] = "TC-%s" % ex["id"]
                    writer.writerow([it["row"], it["id"], it["summary"], ex["id"], kiwi.case_url(ex["id"]),
                                     "unchanged", "", "", problems, "", dup_url])
                    print("%s: giữ nguyên TC-%s" % (prefix, ex["id"]))
                    continue

                if it["update"]:
                    u = it["update"]
                    values = dict(u["values"])
                    if "priority" in values:
                        values["priority"] = resolver.priority(values["priority"])
                    if "category" in values:
                        values["category"] = resolver.category(values["category"])
                    new_status = ""
                    if u["content"]:
                        new_status = "PROPOSED" if has_proposed(it) else status_on_update
                        if new_status.lower() != "keep":
                            values["case_status"] = resolver.status(new_status)
                        else:
                            new_status = ""
                    kiwi.call("TestCase.update", u["id"], values)
                    warnings = add_extras(u["id"], it)
                    counts["updated"] += 1
                    written[it["row"]] = "TC-%s" % u["id"]
                    writer.writerow([it["row"], it["id"], it["summary"], u["id"], kiwi.case_url(u["id"]), "updated",
                                     new_status, "; ".join(u["changes"]), problems, "; ".join(warnings), dup_url])
                    print("%s: cập nhật TC-%s%s – %s" % (prefix, u["id"], " [%s]" % new_status if new_status else "",
                                                       "; ".join(u["changes"])))
                    continue

                values = {
                    "product": resolver.product_id,
                    "category": resolver.category(it["category"]),
                    "priority": resolver.priority(it["priority"]),
                    "case_status": resolver.status(it["final_status"]),
                    "summary": it["summary"],
                    "text": it["text"],
                    "notes": build_notes(it, source, mapping),
                    "is_automated": it["is_automated"],
                    "requirement": it["requirement"],
                }
                case_id = kiwi.call("TestCase.create", values)["id"]
                warnings = add_extras(case_id, it)
                counts["created"] += 1
                written[it["row"]] = "TC-%s" % case_id
                writer.writerow([it["row"], it["id"], it["summary"], case_id, kiwi.case_url(case_id),
                                 "created", it["final_status"], "", problems, "; ".join(warnings), dup_url])
                print("%s: tạo TC-%s [%s]%s" % (prefix, case_id, it["final_status"],
                                                " (! lỗi tag/component)" if warnings else ""))
            except (KiwiError, requests.RequestException) as exc:
                counts["error"] += 1
                writer.writerow([it["row"], it["id"], it["summary"], "", "", "error", "", "", problems, str(exc),
                                 dup_url])
                print("%s: LỖI %s" % (prefix, exc))

        if mapping.get("on_removed", "disable") == "disable":
            for c in to_disable:
                try:
                    kiwi.call("TestCase.update", c["id"], {"case_status": resolver.disabled_id})
                    note = "chuyển DISABLED"
                    warn = ""
                    if plan_id:
                        try:
                            kiwi.call("TestPlan.remove_case", int(plan_id), c["id"])
                            note += ", gỡ khỏi plan"
                        except KiwiError as exc:
                            # Removing from a plan needs a permission normal accounts often lack.
                            warn = "không gỡ được khỏi plan (thiếu quyền): %s" % exc
                            note += ", vẫn nằm trong plan (thiếu quyền gỡ)"
                    counts["disabled"] += 1
                    writer.writerow(["", "", html.unescape(c["summary"]), c["id"], kiwi.case_url(c["id"]),
                                     "disabled", "DISABLED", "không còn trong file", "", warn, ""])
                    print("[xoá khỏi file] TC-%s: %s" % (c["id"], note))
                except (KiwiError, requests.RequestException) as exc:
                    counts["error"] += 1
                    writer.writerow(["", "", c["summary"], c["id"], "", "error", "", "", "", str(exc), ""])
                    print("[xoá khỏi file] TC-%s: LỖI %s" % (c["id"], exc))

    print("\nXong: tạo mới %(created)d, cập nhật %(updated)d, giữ nguyên %(unchanged)d, "
          "DISABLED %(disabled)d, lỗi %(error)d" % counts)
    print("Report: %s" % os.path.abspath(report_path))
    if runs and plan_id:
        try:
            push_results(kiwi, resolver, runs, int(plan_id), args.build, written, args.run)
        except (KiwiError, requests.RequestException) as exc:
            print("! Không đẩy được kết quả test: %s\n  (test case đã đẩy xong; nếu lỗi quyền, nhờ admin Kiwi cấp "
                  "quyền tạo Test Run / cập nhật Test Execution, hoặc chạy lại với --no-results)" % exc)
    # Only touch the file (or produce the paste file) when its Kiwi ID column is missing or wrong.
    sheet_ids = {it["row"]: "TC-%s" % it["sheet_kiwi_id"] for it in items if it["sheet_kiwi_id"]}
    changed = [r for r, v in written.items() if sheet_ids.get(r) != v]
    if args.no_write_back:
        print("Không ghi Kiwi ID vào file (--no-write-back).")
    elif not changed:
        print("Cột Kiwi ID trong file đã đúng – không cần ghi lại.")
    else:
        try:
            print(write_back(args.input, mapping, args.sheet or mapping.get("sheet"), written, args.output))
        except Exception as exc:  # never lose the push result because of the write-back
            print("! Không ghi được Kiwi ID vào file: %s. Lấy cột kiwi_case_id trong report để điền tay." % exc)


def main() -> None:
    ap = argparse.ArgumentParser(description="Migrate / sync test cases from a sheet into Kiwi TCMS")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_inspect = sub.add_parser("inspect", help="Xem các cột và vài dòng đầu của sheet")
    p_inspect.add_argument("--input", required=True)
    p_inspect.add_argument("--sheet")
    p_inspect.add_argument("--header-row", type=int, default=1)

    for name, help_text in (("validate", "Quét sheet, xuất report các điểm cần xác nhận (không ghi lên Kiwi)"),
                            ("run", "Đẩy / đồng bộ test case lên Kiwi")):
        p = sub.add_parser(name, help=help_text)
        p.add_argument("--input", required=True,
                       help="Đường dẫn file .xlsx/.csv ở bất kỳ đâu, hoặc link Google Sheet")
        p.add_argument("--mapping", help="File mapping (mặc định: ./mapping.json, nếu không có thì dùng "
                                         "templates/mapping.example.json của skill)")
        p.add_argument("--plan", type=int, help="Test Plan ID – Product được lấy tự động từ plan")
        p.add_argument("--product", help="Tên Product trên Kiwi (không cần nếu đã có --plan)")
        p.add_argument("--sheet", help="Tên tab (xlsx); ghi đè 'sheet' trong mapping")
        p.add_argument("--limit", type=int, help="Chỉ xử lý N case đầu (để thử; không tắt case đã xoá khỏi file)")
        p.add_argument("--build", help="Build (id hoặc tên) cho Test Run mới khi đẩy kết quả test")
        p.add_argument("--run", type=int, help="Ghi kết quả test vào Test Run có sẵn (ID trong URL .../runs/<ID>/)")
        if name == "validate":
            p.add_argument("--offline", action="store_true", help="Không kiểm tra với Kiwi")
            p.add_argument("--preview", type=int, default=0, help="In nội dung N case đầu")
        else:
            p.add_argument("--dry-run", action="store_true")
            p.add_argument("--preview", type=int, default=5, help="Số case in ra khi dry-run")
            p.add_argument("--accept-issues", action="store_true",
                           help="Vẫn chạy khi còn điểm cần xác nhận (case có vấn đề gắn PROPOSED)")
            p.add_argument("--no-write-back", action="store_true", help="Không ghi cột Kiwi ID vào file")
            p.add_argument("--no-results", action="store_true", help="Không đẩy kết quả test (chỉ đẩy test case)")
            p.add_argument("--output", help="Ghi file có cột Kiwi ID ra đường dẫn này thay vì file gốc")

    args = ap.parse_args()
    try:
        {"inspect": cmd_inspect, "validate": cmd_validate, "run": cmd_run}[args.cmd](args)
    except KiwiError as exc:
        sys.exit("LỖI Kiwi: %s" % exc)


if __name__ == "__main__":
    main()
