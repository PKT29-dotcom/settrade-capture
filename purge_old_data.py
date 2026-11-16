"""
Purge Old Data  v1.1.0  (purge_old_data.py)
-------------------------------------------
ย้ายแถวเก่าออกจากแท็บที่โตเร็ว เพื่อไม่ให้ Google Sheet ใหญ่เกินไป โดย "สำรองไปชีท Backup
ก่อน แล้วค่อยลบ" จากชีทหลัก

กติกา "เดือนนี้ + ย้อนหลัง N เดือน": เก็บข้อมูลตั้งแต่วันที่ 1 ของเดือนที่ย้อนหลังไป N
เดือนถึงปัจจุบัน เช่น รันวันที่ 12 พ.ย. 2026 (N=3) เก็บตั้งแต่ 1 ส.ค. 2026 (ส.ค.+ก.ย.+ต.ค.+พ.ย.)

ขั้นตอนต่อแท็บ:
  1) หาแถวที่วันที่เก่ากว่าวันตัด (แถวที่อ่านวันที่ไม่ได้จะถูกเก็บไว้ ไม่แตะ)
  2) คัดลอกแถวเหล่านั้นไปแท็บชื่อเดียวกันในชีท Backup (BACKUP_GSHEET_ID) ทีละวัน
     (วันที่อยู่ในแบ็กอัปแล้วจะข้าม จึงรันซ้ำได้โดยไม่ซ้ำ) วันที่ใส่เป็นข้อความ dd/mm/yyyy
  3) สำรองสำเร็จครบแล้วเท่านั้น จึงลบแถวนั้นออกจากชีทหลัก
  ถ้าสำรองไม่สำเร็จ หรือไม่ได้ตั้งค่าชีท Backup จะ "ไม่ลบ" (ยกเว้นตั้ง PURGE_WITHOUT_BACKUP=true)

ความปลอดภัยอื่น ๆ:
  - DRY_RUN เป็นค่าเริ่มต้น (ไม่ลบจริง แค่รายงาน) จะลบจริงเมื่อตั้ง DRY_RUN=false เท่านั้น
  - ถ้าจะลบเกิน 95% ของแถวที่อ่านวันที่ได้ จะยกเลิกแท็บนั้น (กันกรณีอ่านวันที่ผิด)
  - ไม่แตะแถวหัวตาราง ผลการรันบันทึกลงแท็บ Log
  - แท็บที่ "ไม่" Purge โดยตั้งใจ: SectorIndex และ SetDatabase (โตช้า ใช้วิเคราะห์ย้อนหลังยาว)
  - ชีทปลายทางที่ sync (เขียนทับ) จะสั้นลงเองหลัง sync รอบถัดไป

Environment variables:
    GSHEET_ID, GOOGLE_APPLICATION_CREDENTIALS
    BACKUP_GSHEET_ID      ID ของชีท Backup (แชร์สิทธิ์ Editor ให้ service account เดียวกัน)
    DRY_RUN               ค่าเริ่มต้น true
    PURGE_WITHOUT_BACKUP  ค่าเริ่มต้น false (true = ยอมลบโดยไม่สำรอง ไม่แนะนำ)

Changelog (เพิ่มเวอร์ชันทุกครั้งที่แก้ไฟล์นี้ และเวอร์ชันจะแสดงใน Log คอลัมน์ Detail):
  1.1.0  สำรองแถวที่จะลบไปชีท Backup ก่อนลบ (ข้ามวันที่ที่สำรองแล้ว), ไม่ลบถ้าไม่มี/สำรองไม่สำเร็จ,
         ลองใหม่อัตโนมัติเมื่อเชื่อมต่อ Google Sheets หลุด, แสดงเวอร์ชันใน Log
  1.0.0  ลบแถวเก่าอย่างเดียว (ไม่มี Backup)
"""

import os
import re
import sys
import time
import traceback
import datetime as dt
from zoneinfo import ZoneInfo

import gspread
import requests
from google.oauth2.service_account import Credentials

SCRIPT_NAME = "Purge Old Data"
__version__ = "1.1.0"

BANGKOK_TZ = ZoneInfo("Asia/Bangkok")

# ชื่อแท็บ -> จำนวนเดือนที่เก็บย้อนหลัง (ไม่นับเดือนนี้)
PURGE_RULES = {
    "TopDatabase": 3,
    "StockDailyDetail": 3,
    "NVDRTradingByStock": 3,
    "Log": 3,
}
HEADER_ROWS = {"NVDRTradingByStock": 2}      # แท็บที่หัวตารางมี 2 แถว (ที่เหลือมี 1 แถว)

MAX_DELETE_FRACTION = 0.95
ARCHIVE_CHUNK_ROWS = 4000
DATE_FORMATS = ("%d/%m/%Y", "%Y-%m-%d")
LOG_SHEET_NAME = "Log"
LOG_HEADERS = ["Date", "Time", "Workflow", "Trigger", "Status", "RowsSent", "Detail"]

RETRY_ATTEMPTS = 5
RETRY_BASE_DELAY = 2


# ---------- ลองใหม่เมื่อเชื่อมต่อหลุดชั่วคราว ----------
def _is_transient(e) -> bool:
    if isinstance(e, (requests.exceptions.ConnectionError, requests.exceptions.Timeout,
                      requests.exceptions.ChunkedEncodingError, ConnectionError, TimeoutError)):
        return True
    if isinstance(e, gspread.exceptions.APIError):
        code = getattr(getattr(e, "response", None), "status_code", None)
        return code in (429, 500, 502, 503, 504)
    return False


def _backoff(attempt: int, e) -> None:
    delay = min(RETRY_BASE_DELAY * 2 ** (attempt - 1), 30)
    print(f"  คำเตือน: เชื่อมต่อ Google Sheets ไม่สำเร็จ ({type(e).__name__}) "
          f"ลองใหม่ครั้งที่ {attempt + 1}/{RETRY_ATTEMPTS} ใน {delay} วินาที")
    time.sleep(delay)


def call_with_retry(fn, *args, **kwargs):
    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            if not _is_transient(e) or attempt == RETRY_ATTEMPTS:
                raise
            _backoff(attempt, e)


# ---------- วันที่ / แผนการลบ ----------
def parse_date(s):
    s = str(s).strip().lstrip("'")
    for fmt in DATE_FORMATS:
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def cutoff_date(today: dt.date, months: int) -> dt.date:
    y, m = today.year, today.month - months
    while m <= 0:
        m += 12
        y -= 1
    return dt.date(y, m, 1)


def plan_deletion(date_cells, cutoff, header_rows=1):
    """date_cells = ค่าทั้งคอลัมน์วันที่ (รวมแถวหัวตาราง)
    คืนค่า (แถวที่จะลบ เลขแถวแบบ 1-based, จำนวนแถวที่อ่านวันที่ได้, จำนวนที่อ่านไม่ได้)"""
    to_delete, parsed, unparsed = [], 0, 0
    for idx, cell in enumerate(date_cells[header_rows:], start=header_rows + 1):
        d = parse_date(cell)
        if d is None:
            unparsed += 1
            continue
        parsed += 1
        if d < cutoff:
            to_delete.append(idx)
    return to_delete, parsed, unparsed


def group_ranges(rows):
    ranges = []
    for r in sorted(rows):
        if ranges and r == ranges[-1][1] + 1:
            ranges[-1] = (ranges[-1][0], r)
        else:
            ranges.append((r, r))
    return ranges


def build_delete_requests(sheet_id, ranges):
    """ลบจากล่างขึ้นบน เพื่อให้เลขแถวของช่วงที่เหลือไม่เลื่อน"""
    return [
        {"deleteDimension": {"range": {
            "sheetId": sheet_id, "dimension": "ROWS",
            "startIndex": start - 1, "endIndex": end,
        }}}
        for start, end in sorted(ranges, reverse=True)
    ]


def _norm(s):
    return re.sub(r"[^a-zA-Z0-9]", "", s).lower()


def find_worksheet(sh, title):
    for ws in call_with_retry(sh.worksheets):
        if _norm(ws.title) == _norm(title):
            return ws
    return None


# ---------- สำรองไปชีท Backup ----------
def archive_to_backup(backup_sh, title, header_vals, rows, date_idx):
    """คัดลอก rows ไปแท็บชื่อเดียวกันในชีท Backup ทีละวัน (วันที่สำรองแล้วจะข้าม)
    คืนค่า (จำนวนแถวที่เขียนใหม่, จำนวนวันที่ข้ามเพราะมีแล้ว)"""
    bws = find_worksheet(backup_sh, title)
    if bws is None:
        bws = call_with_retry(backup_sh.add_worksheet, title=title,
                              rows=max(1000, len(rows) + 50), cols=max(15, len(header_vals[0]) if header_vals else 15))
        if header_vals:
            call_with_retry(bws.append_rows, header_vals, value_input_option="USER_ENTERED",
                            table_range="A1")

    def dates_in_backup():
        cells = call_with_retry(bws.col_values, date_idx + 1)
        return {d for d in (parse_date(c) for c in cells[len(header_vals):]) if d}

    done = dates_in_backup()
    groups = {}
    for r in rows:
        groups.setdefault(parse_date(r[date_idx]), []).append(r)
    pending = [(d, rs) for d, rs in sorted(groups.items()) if d not in done]
    skipped = len(groups) - len(pending)

    written, chunk, chunk_days = 0, [], []

    def flush():
        nonlocal written, chunk, chunk_days
        if not chunk:
            return
        for attempt in range(1, RETRY_ATTEMPTS + 1):
            try:
                call_args = dict(value_input_option="USER_ENTERED", table_range="A1")
                bws.append_rows(chunk, **call_args)
                break
            except Exception as e:
                if not _is_transient(e) or attempt == RETRY_ATTEMPTS:
                    raise
                _backoff(attempt, e)
                try:   # เขียนสำเร็จแล้วแต่คำตอบหาย? เช็คก่อนเขียนซ้ำ
                    if all(d in dates_in_backup() for d in chunk_days):
                        break
                except Exception:
                    pass
        written += len(chunk)
        chunk, chunk_days = [], []

    for d, rs in pending:
        for r in rs:
            row = list(r)
            row[date_idx] = "'" + str(row[date_idx]).lstrip("'")     # เก็บวันที่เป็นข้อความ
            chunk.append(row)
        chunk_days.append(d)
        if len(chunk) >= ARCHIVE_CHUNK_ROWS:       # ตัดก้อนเฉพาะระหว่างวัน ไม่ตัดกลางวัน
            flush()
    flush()
    return written, skipped


def purge_sheet(sh, backup_sh, backup_error, allow_no_backup, title, months, today, dry_run):
    ws = find_worksheet(sh, title)
    if ws is None:
        return {"tab": title, "status": "NoTab", "delete": 0}

    values = call_with_retry(ws.get_all_values)
    header_rows = HEADER_ROWS.get(title, 1)
    header = values[0] if values else []
    date_idx = header.index("Date") if "Date" in header else 0
    cells = [(r[date_idx] if len(r) > date_idx else "") for r in values]
    cutoff = cutoff_date(today, months)
    to_delete, parsed, unparsed = plan_deletion(cells, cutoff, header_rows)
    info = {"tab": title, "cutoff": cutoff.strftime("%d/%m/%Y"), "rows": max(len(cells) - header_rows, 0),
            "delete": len(to_delete), "unparsed": unparsed}

    if not to_delete:
        info["status"] = "NothingToDelete"
        return info
    if parsed and len(to_delete) >= MAX_DELETE_FRACTION * parsed:
        info["status"] = "Aborted"
        info["note"] = (f"จะลบ {len(to_delete)} จาก {parsed} แถวที่อ่านวันที่ได้ "
                        f"เกิน {int(MAX_DELETE_FRACTION * 100)}% จึงไม่ลบ")
        return info
    if dry_run:
        info["status"] = "DryRun"
        info["note"] = ("จะสำรองไปชีท Backup ก่อนลบ" if backup_sh is not None else
                        "ยังไม่มีชีท Backup จะไม่ลบจริง" if not allow_no_backup else "ลบโดยไม่สำรอง")
        return info

    if backup_sh is None and not allow_no_backup:
        info["status"] = "BackupFailed" if backup_error else "NoBackup"
        info["note"] = backup_error or "ไม่ได้ตั้ง BACKUP_GSHEET_ID จึงไม่ลบ"
        return info

    if backup_sh is not None:
        try:
            written, skipped_days = archive_to_backup(
                backup_sh, title, values[:header_rows], [values[r - 1] for r in to_delete], date_idx)
        except Exception as e:
            info["status"] = "BackupFailed"
            info["note"] = f"สำรองไม่สำเร็จ ไม่ลบ: {type(e).__name__}: {str(e)[:120]}"
            return info
        info["archived"] = written
        info["archived_skipped_days"] = skipped_days

    requests_body = build_delete_requests(ws.id, group_ranges(to_delete))
    for i in range(0, len(requests_body), 50):
        call_with_retry(sh.batch_update, {"requests": requests_body[i:i + 50]})
    info["status"] = "Deleted"
    return info


# ---------- เชื่อมต่อ ----------
def _client():
    creds_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if not creds_path:
        raise RuntimeError("ไม่พบ GOOGLE_APPLICATION_CREDENTIALS")
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    return gspread.authorize(Credentials.from_service_account_file(creds_path, scopes=scopes))


def get_open_spreadsheet():
    sheet_id = os.environ.get("GSHEET_ID")
    if not sheet_id:
        raise RuntimeError("ไม่พบ GSHEET_ID")
    return call_with_retry(_client().open_by_key, sheet_id)


def get_backup_spreadsheet():
    """คืนค่า (ชีท Backup หรือ None, ข้อความ error ถ้าเปิดไม่ได้)"""
    backup_id = (os.environ.get("BACKUP_GSHEET_ID") or "").strip()
    if not backup_id:
        return None, None
    try:
        return call_with_retry(_client().open_by_key, backup_id), None
    except Exception as e:
        return None, f"เปิดชีท Backup ไม่ได้: {type(e).__name__}: {str(e)[:120]}"


def push_to_log(sh, date_str, time_str, workflow, trigger, status, rows, detail):
    detail = f"[{SCRIPT_NAME} v{__version__}] {detail}".strip()
    ws = find_worksheet(sh, LOG_SHEET_NAME)
    if ws is None:
        ws = call_with_retry(sh.add_worksheet, title=LOG_SHEET_NAME, rows=2000, cols=10)
        call_with_retry(ws.append_row, LOG_HEADERS, value_input_option="USER_ENTERED", table_range="A1")
    call_with_retry(ws.append_row, [date_str, time_str, workflow, trigger, status, rows, detail],
                    value_input_option="USER_ENTERED", table_range="A1")


def main():
    now = dt.datetime.now(BANGKOK_TZ)
    dry_run = os.environ.get("DRY_RUN", "true").strip().lower() != "false"
    allow_no_backup = os.environ.get("PURGE_WITHOUT_BACKUP", "false").strip().lower() == "true"
    ev = os.environ.get("GITHUB_EVENT_NAME", "local")
    trigger = {"schedule": "Scheduled", "workflow_dispatch": "Manual"}.get(ev, ev)
    workflow = os.environ.get("GITHUB_WORKFLOW", "local")
    print(f"[{now:%d/%m/%Y %H:%M}] {SCRIPT_NAME} v{__version__} "
          f"({'DRY RUN - ไม่ลบจริง' if dry_run else 'ลบจริง'})")

    sh = get_open_spreadsheet()
    backup_sh, backup_error = get_backup_spreadsheet()
    print("  ชีท Backup:", "พร้อมใช้งาน" if backup_sh is not None else (backup_error or "ไม่ได้ตั้งค่า"))

    results = []
    try:
        for title, months in PURGE_RULES.items():
            info = purge_sheet(sh, backup_sh, backup_error, allow_no_backup, title, months, now.date(), dry_run)
            results.append(info)
            print(f"  {title}: {info['status']} ลบ {info['delete']} แถว"
                  f" สำรอง {info.get('archived', '-')}"
                  f" (เก็บตั้งแต่ {info.get('cutoff', '-')}, อ่านวันที่ไม่ได้ {info.get('unparsed', 0)})"
                  f" {info.get('note', '')}")
    except Exception as e:
        push_to_log(sh, f"{now:%d/%m/%Y}", f"{now:%H:%M}", workflow, trigger, "Failed", 0,
                    f"{type(e).__name__}: {str(e)[:300]}")
        raise

    deleted = sum(r["delete"] for r in results if r["status"] == "Deleted")
    would = sum(r["delete"] for r in results)
    detail = "; ".join(f"{r['tab']}:{r['status']}({r['delete']}"
                       + (f",สำรอง{r['archived']}" if "archived" in r else "") + ")" for r in results)
    status = "PurgeDryRun" if dry_run else "PurgeDone"
    push_to_log(sh, f"{now:%d/%m/%Y}", f"{now:%H:%M}", workflow, trigger, status,
                would if dry_run else deleted, detail)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
