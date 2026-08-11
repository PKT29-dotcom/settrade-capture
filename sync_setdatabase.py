"""
sync_setdatabase.py
---------------------
Copy ข้อมูลจากแท็บ SetDatabase ในชีท "Settrade Capture Log" (ต้นทาง) ไปวางที่
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
    python sync_setdatabase.py
"""

import os
import re
import sys
import traceback
import datetime as dt
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials

BANGKOK_TZ = ZoneInfo("Asia/Bangkok")

SOURCE_SHEET_NAME = "SetDatabase"
TARGET_SHEET_NAME = "SetDatabase Schedule"

LOG_SHEET_NAME = "Log"
LOG_HEADERS = ["Date", "Time", "Workflow", "Trigger", "Status", "RowsSent", "Detail"]


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
    return gc.open_by_key(sheet_id)


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
    ws = get_or_create_worksheet(sh, LOG_SHEET_NAME, rows=2000, cols=10)
    if len(ws.get_all_values()) == 0:
        ws.append_row(LOG_HEADERS, value_input_option="USER_ENTERED", table_range="A1")

    ws.append_row(
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
    print(f"[{date_str} {time_str} เวลาไทย] เริ่ม sync SetDatabase ไปยังชีทปลายทาง "
          f"(workflow: {workflow_name}, trigger: {trigger_label})")

    source_sh = get_open_spreadsheet("GSHEET_ID")

    try:
        source_ws = source_sh.worksheet(SOURCE_SHEET_NAME)
        all_values = source_ws.get_all_values()
    except Exception as e:
        detail = f"{type(e).__name__}: {str(e)[:200]}"
        push_to_log(source_sh, date_str, time_str, workflow_name, trigger_label, "Failed", 0, detail)
        raise

    if not all_values:
        print("  ไม่พบข้อมูลในแท็บ SetDatabase ต้นทาง -> ข้าม sync รอบนี้")
        push_to_log(source_sh, date_str, time_str, workflow_name, trigger_label, "NoData", 0,
                    "ไม่พบข้อมูลในแท็บ SetDatabase ต้นทาง")
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
