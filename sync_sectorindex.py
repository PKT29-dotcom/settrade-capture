"""
Sync SectorIndex  v1.1.0  (sync_sectorindex.py)
---------------------
Copy ข้อมูลจากแท็บ SectorIndex ในชีท "Settrade Capture Log" (ต้นทาง) ไปวางที่
แท็บ "SetDatabase schedule" ในชีทของอีก Google Account (ปลายทาง)
ต้นทางและปลายทาง

"เขียนทับทั้งหมดทุกครั้ง" (ไม่ append) - ลบข้อมูลเดิมในแท็บปลายทางออกก่อน
แล้วเขียนข้อมูลปัจจุบันทั้งหมดจากต้นทางลงไปใหม่ (แบบเดียวกับ
sync_topdatabase.py) ทำให้ปลายทางไม่มีวันมีแถวซ้ำซ้อนสะสม เพราะเขียนทับ
ทั้งชุดใหม่ทุกรอบเสมอ

ต่างจาก sync_topdatabase.py ตรงที่ไม่ต้องปัดเวลา (SetDatabase capture วันละ
1 แถวเท่านั้น ไม่มี 4 ช่วงเวลาให้ปัดเหมือน TopDatabase)

Environment variables ที่ต้องตั้ง (ใช้ secrets ชุดเดียวกับ sync_topdatabase.py):
    GSHEET_ID                      -> ID ของ Google Sheet ต้นทาง (ของเราเอง)
    TARGET_GSHEET_ID                -> ID ของ Google Sheet ปลายทาง (อีก Account)
    GOOGLE_APPLICATION_CREDENTIALS -> path ของไฟล์ service-account JSON
      (ต้องแชร์สิทธิ์ Editor ให้ service account นี้ ทั้งในชีทต้นทางและ
      ชีทปลายทาง - ใช้ตัวเดียวกับที่แชร์ไว้แล้วสำหรับ sync_topdatabase.py)

การใช้งาน (ทดสอบในเครื่องตัวเอง):
    export GSHEET_ID="....."
    export TARGET_GSHEET_ID="....."
    export GOOGLE_APPLICATION_CREDENTIALS="./gcp-credentials.json"
    python sync_sectorindex.py

Changelog (เพิ่มเวอร์ชันทุกครั้งที่แก้ไฟล์นี้ และเวอร์ชันจะแสดงใน Log คอลัมน์ Detail):
  1.1.0  ลองใหม่อัตโนมัติเมื่อเชื่อมต่อ Google Sheets หลุดชั่วคราว, ถ้าแท็บต้นทางว่างจะไม่ sync
         (กันล้างแท็บปลายทางทิ้ง) และบันทึก NoData, แสดงเวอร์ชันใน Log, เขียนหัวตาราง Log ให้
         ถ้าแท็บ Log ว่าง
  1.0.0  เวอร์ชันเดิม (ไม่มีเลขเวอร์ชัน)
"""

import os
import time
import re
import sys
import traceback
import datetime as dt
from zoneinfo import ZoneInfo

import gspread
import requests
from google.oauth2.service_account import Credentials

BANGKOK_TZ = ZoneInfo("Asia/Bangkok")

SOURCE_SHEET_NAME = "SectorIndex"
TARGET_SHEET_NAME = "SectorIndex Schedule"

LOG_SHEET_NAME = "Log"
LOG_HEADERS = ["Date", "Time", "Workflow", "Trigger", "Status", "RowsSent", "Detail"]


SCRIPT_NAME = "Sync SectorIndex"
__version__ = "1.1.0"

# ---------- ลองใหม่อัตโนมัติเมื่อเชื่อมต่อ Google Sheets หลุดชั่วคราว ----------
RETRY_ATTEMPTS = 5
RETRY_BASE_DELAY = 2


def _is_transient(e) -> bool:
    """ข้อผิดพลาดชั่วคราวที่ควรลองใหม่ (เชื่อมต่อหลุด, หมดเวลา, 429, 5xx)"""
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


def append_safe(ws, rows, already_done):
    """เพิ่มแถวแบบกันซ้ำ: ถ้าเชื่อมต่อหลุด จะอ่านชีทดูก่อนว่ารอบที่แล้วเขียนเข้าไปแล้วหรือยัง
    (already_done = ฟังก์ชันรับค่าที่อ่านจากชีท คืน True ถ้าข้อมูลชุดนี้อยู่ในชีทแล้ว)"""
    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            ws.append_rows(rows, value_input_option="USER_ENTERED", table_range="A1")
            return
        except Exception as e:
            if not _is_transient(e) or attempt == RETRY_ATTEMPTS:
                raise
            _backoff(attempt, e)
            try:
                if already_done(call_with_retry(ws.get_all_values)):
                    print("  ตรวจพบว่าข้อมูลถูกเขียนสำเร็จไปแล้วก่อนเชื่อมต่อหลุด ไม่เขียนซ้ำ")
                    return
            except Exception:
                pass


def _is_sheet_like(x) -> bool:
    return hasattr(x, "get_all_values") or hasattr(x, "worksheets")


class _Retrying:
    """ครอบ Spreadsheet/Worksheet ของ gspread: ทุกเมธอดที่ติดต่อ Google จะลองใหม่เมื่อหลุดชั่วคราว
    ยกเว้น append_row/append_rows ที่ลองซ้ำแบบสุ่มสี่สุ่มห้าไม่ได้ (เสี่ยงข้อมูลซ้ำ) ให้ใช้ append_safe"""
    _NO_BLIND_RETRY = {"append_row", "append_rows"}

    def __init__(self, obj):
        object.__setattr__(self, "_obj", obj)

    def __getattr__(self, name):
        attr = getattr(self._obj, name)
        if not callable(attr):
            return attr

        def call(*args, **kwargs):
            res = attr(*args, **kwargs) if name in self._NO_BLIND_RETRY else call_with_retry(attr, *args, **kwargs)
            if _is_sheet_like(res):
                return _Retrying(res)
            if isinstance(res, list) and res and all(_is_sheet_like(w) for w in res):
                return [_Retrying(w) for w in res]
            return res
        return call


def _is_blank(values) -> bool:
    """แท็บว่างจริงหรือไม่ (gspread คืน [[]] เมื่อแท็บว่าง ดังนั้นห้ามเช็คจาก len หรือ not)"""
    return not any((c or "").strip() for row in values for c in row)


def get_open_spreadsheet(sheet_id_env: str):
    sheet_id = os.environ.get(sheet_id_env)
    creds_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if not sheet_id or not creds_path:
        raise RuntimeError(
            f"ไม่พบ {sheet_id_env} หรือ GOOGLE_APPLICATION_CREDENTIALS "
            "ใน environment variables กรุณาตั้งค่าก่อนรัน (ดู README.md)"
        )
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    gc = gspread.authorize(creds)
    return _Retrying(call_with_retry(gc.open_by_key, sheet_id))


def get_trigger_label() -> str:
    event_name = os.environ.get("GITHUB_EVENT_NAME", "")
    if event_name == "schedule":
        return "Scheduled"
    elif event_name == "workflow_dispatch":
        return "Manual"
    elif event_name:
        return event_name
    return "local"


def get_workflow_name() -> str:
    return os.environ.get("GITHUB_WORKFLOW", "local")


def get_or_create_worksheet(sh, title: str, rows: int, cols: int):
    """
    หาแท็บด้วยชื่อ (normalize แบบเข้มงวด ตัดทุกอักขระที่ไม่ใช่ตัวอักษร/ตัวเลข
    ออกก่อนเทียบ) ถ้าไม่เจอค่อยสร้างใหม่ พร้อม fallback ดักจับ error กรณี
    Google บอกว่ามีแท็บชื่อนี้อยู่แล้วจริง แต่ list ที่ดึงมาหาไม่เจอ (บั๊กที่
    เคยเจอกับ sync_topdatabase.py - ดูรายละเอียดในไฟล์นั้น)
    """
    def _normalize(s: str) -> str:
        return re.sub(r"[^a-zA-Z0-9]", "", s).lower()

    target_key = _normalize(title)

    def _find_match():
        for ws in sh.worksheets():
            if _normalize(ws.title) == target_key:
                return ws
        return None

    match = _find_match()
    if match:
        return match

    try:
        return sh.add_worksheet(title=title, rows=rows, cols=cols)
    except gspread.exceptions.APIError as e:
        if "already exists" not in str(e):
            raise
        match = _find_match()
        if match:
            return match
        raise


def push_to_log(sh, date_str: str, time_str: str, workflow_name: str, trigger_label: str,
                 status: str, rows_sent, detail: str = ""):
    detail = f"[{SCRIPT_NAME} v{__version__}] {detail}".strip()
    ws = get_or_create_worksheet(sh, LOG_SHEET_NAME, rows=2000, cols=10)
    if _is_blank(call_with_retry(ws.get_all_values)):
        call_with_retry(ws.append_row, LOG_HEADERS, value_input_option="USER_ENTERED", table_range="A1")

    call_with_retry(
        ws.append_row,
        [date_str, time_str, workflow_name, trigger_label, status, rows_sent, detail],
        value_input_option="USER_ENTERED",
        table_range="A1",
    )


def sync_once():
    now = dt.datetime.now(BANGKOK_TZ)
    date_str = now.strftime("%d/%m/%Y")
    time_str = now.strftime("%H:%M")
    trigger_label = get_trigger_label()
    workflow_name = get_workflow_name()
    print(f"[{date_str} {time_str} เวลาไทย] {SCRIPT_NAME} v{__version__} เริ่ม sync SectorIndex ไปยังชีทปลายทาง "
          f"(workflow: {workflow_name}, trigger: {trigger_label})")

    source_sh = get_open_spreadsheet("GSHEET_ID")

    try:
        source_ws = source_sh.worksheet(SOURCE_SHEET_NAME)
        all_values = source_ws.get_all_values()
    except Exception as e:
        detail = f"{type(e).__name__}: {str(e)[:200]}"
        push_to_log(source_sh, date_str, time_str, workflow_name, trigger_label, "Failed", 0, detail)
        raise

    if _is_blank(all_values):
        print("  ไม่พบข้อมูลในแท็บ SectorIndex ต้นทาง -> ข้าม sync รอบนี้")
        push_to_log(source_sh, date_str, time_str, workflow_name, trigger_label, "NoData", 0,
                    "ไม่พบข้อมูลในแท็บ SectorIndex ต้นทาง")
        return

    header = all_values[0]
    data_rows = all_values[1:]

    print(f"  พบ {len(data_rows)} แถวจากต้นทาง กำลังส่งไปยังชีทปลายทาง ...")

    try:
        target_sh = get_open_spreadsheet("TARGET_GSHEET_ID")
        target_ws = get_or_create_worksheet(
            target_sh, TARGET_SHEET_NAME,
            rows=len(data_rows) + 100, cols=len(header) + 2,
        )
        target_ws.clear()

        all_rows_to_write = [header] + data_rows
        target_ws.update(all_rows_to_write, value_input_option="USER_ENTERED")
    except Exception as e:
        detail = f"{type(e).__name__}: {str(e)[:200]}"
        push_to_log(source_sh, date_str, time_str, workflow_name, trigger_label, "Failed", 0, detail)
        raise

    print(f"  เขียนทับข้อมูล {len(data_rows)} แถว เข้าแท็บ '{TARGET_SHEET_NAME}' ในชีทปลายทางสำเร็จ")
    push_to_log(source_sh, date_str, time_str, workflow_name, trigger_label,
                "Success", len(data_rows), "")


if __name__ == "__main__":
    try:
        sync_once()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่าง sync:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
