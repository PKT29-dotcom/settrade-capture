"""
sync_topdatabase.py
---------------------
Copy ข้อมูลจากแท็บ TopDatabase ในชีท "Settrade Capture Log" (ต้นทาง) ไปวางที่
แท็บ "TopDatabase schedule" ในชีทของอีก Google Account (ปลายทาง) โดย:

1. "เขียนทับทั้งหมดทุกครั้ง" (ไม่ append) - ลบข้อมูลเดิมในแท็บปลายทางออกก่อน
   แล้วเขียนข้อมูลปัจจุบันทั้งหมดจากต้นทางลงไปใหม่
2. "ปัดเวลา" คอลัมน์ Time ให้เข้าใกล้ 1 ใน 4 ช่วงเวลาที่กำหนด (10:30, 12:30,
   15:00, 16:30) ที่ใกล้ที่สุด เพราะเวลาจริงที่ capture ได้อาจคลาดเคลื่อนจาก
   ที่ตั้งไว้ (เช่น GitHub Actions delay ทำให้ได้ 10:50 แทนที่จะเป็น 10:30)
   แต่ชีทปลายทางมี dropdown/filter ที่ต้องการค่าตรงกับ 4 ช่วงนี้เป๊ะ

   หมายเหตุ: การปัดเวลานี้มีผลเฉพาะข้อมูลที่ส่งไปให้ชีทปลายทางเท่านั้น
   ไม่ได้แก้ไขค่า Time จริงในชีทต้นทาง (Settrade Capture Log) เพื่อให้ยังคง
   เห็นเวลา capture จริงไว้สำหรับตรวจสอบ/วิเคราะห์ (เหมือนที่เคยทำเรื่อง
   Log ไว้ก่อนหน้านี้)

Environment variables ที่ต้องตั้ง (เพิ่มจากที่มีอยู่เดิม):
    GSHEET_ID                      -> ID ของ Google Sheet ต้นทาง (ของเราเอง)
    TARGET_GSHEET_ID                -> ID ของ Google Sheet ปลายทาง (อีก Account)
    GOOGLE_APPLICATION_CREDENTIALS -> path ของไฟล์ service-account JSON
      (ต้องแชร์สิทธิ์ Editor ให้ service account นี้ ทั้งในชีทต้นทางและ
      ชีทปลายทาง มิฉะนั้นจะเขียน/อ่านชีทปลายทางไม่ได้)

การใช้งาน (ทดสอบในเครื่องตัวเอง):
    export GSHEET_ID="....."
    export TARGET_GSHEET_ID="....."
    export GOOGLE_APPLICATION_CREDENTIALS="./gcp-credentials.json"
    python sync_topdatabase.py
"""

import os
import sys
import traceback
import datetime as dt
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials

BANGKOK_TZ = ZoneInfo("Asia/Bangkok")

SOURCE_SHEET_NAME = "TopDatabase"
TARGET_SHEET_NAME = "TopDatabase schedule"

LOG_SHEET_NAME = "Log"
LOG_HEADERS = ["Date", "Time", "Workflow", "Trigger", "Status", "RowsSent", "Detail"]

# 4 ช่วงเวลาที่ต้องการให้ผลลัพธ์ตรงเป๊ะ (นาทีนับจากเที่ยงคืน ไว้คำนวณระยะห่าง)
CANONICAL_SLOTS = ["10:30", "12:30", "15:00", "16:30"]


def _time_to_minutes(time_str: str):
    """แปลง 'HH:MM' เป็นจำนวนนาทีนับจากเที่ยงคืน คืน None ถ้า parse ไม่ได้"""
    try:
        hh, mm = time_str.strip().split(":")
        return int(hh) * 60 + int(mm)
    except (ValueError, AttributeError):
        return None


_CANONICAL_MINUTES = [(_time_to_minutes(s), s) for s in CANONICAL_SLOTS]


def snap_to_nearest_slot(time_str: str) -> str:
    """
    ปัดเวลาที่ได้จริงให้เข้าใกล้ 1 ใน 4 ช่วงเวลาที่กำหนดที่ใกล้ที่สุด
    เช่น "10:50" -> "10:30", "12:35" -> "12:30", "14:58" -> "15:00"
    ถ้า parse เวลาไม่ได้ (ค่าว่าง/ผิดรูปแบบ) คืนค่าเดิมกลับไปโดยไม่แก้
    """
    minutes = _time_to_minutes(time_str)
    if minutes is None:
        return time_str

    best_slot = min(_CANONICAL_MINUTES, key=lambda pair: abs(pair[0] - minutes))
    return best_slot[1]


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


def get_or_create_worksheet(sh, title: str, rows: int, cols: int):
    """
    หาแท็บด้วยชื่อ (ตัดช่องว่างหัวท้ายออกก่อนเทียบ) ถ้าไม่เจอค่อยสร้างใหม่

    ใช้แทน sh.worksheet(title) ตรง ๆ เพราะเจอบั๊กที่ gspread หาแท็บที่มีอยู่
    จริงไม่เจอ (อาจเกิดจากช่องว่างแฝง/อักขระที่มองไม่เห็นต่างกันเล็กน้อยใน
    ชื่อแท็บจริงบน Google Sheets) แล้วพยายามสร้างแท็บใหม่ซ้ำ ทำให้ชนกับของเดิม
    และเกิด APIError "sheet with the name already exists" ทั้งที่จริง ๆ มีแท็บ
    นั้นอยู่แล้ว วิธีนี้ทนทานกว่าเพราะดึงรายชื่อแท็บทั้งหมดมาเทียบเองแบบ
    normalize ก่อน ไม่พึ่งพา exact-match lookup ของ gspread ตรง ๆ
    """
    for ws in sh.worksheets():
        if ws.title.strip() == title.strip():
            return ws
    return sh.add_worksheet(title=title, rows=rows, cols=cols)


def sync_once():
    now = dt.datetime.now(BANGKOK_TZ)
    date_str = now.strftime("%d/%m/%Y")
    time_str = now.strftime("%H:%M")
    trigger_label = get_trigger_label()
    workflow_name = get_workflow_name()
    print(f"[{date_str} {time_str} เวลาไทย] เริ่ม sync TopDatabase ไปยังชีทปลายทาง "
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
        print("  ไม่พบข้อมูลในแท็บ TopDatabase ต้นทาง -> ข้าม sync รอบนี้")
        push_to_log(source_sh, date_str, time_str, workflow_name, trigger_label, "NoData", 0,
                    "ไม่พบข้อมูลในแท็บ TopDatabase ต้นทาง")
        return

    header = all_values[0]
    data_rows = all_values[1:]

    try:
        time_col_index = header.index("Time")
    except ValueError:
        detail = "ไม่พบคอลัมน์ 'Time' ในหัวตาราง TopDatabase ต้นทาง"
        push_to_log(source_sh, date_str, time_str, workflow_name, trigger_label, "Failed", 0, detail)
        raise RuntimeError(detail)

    # ปัดค่า Time ของทุกแถวให้เข้าใกล้ 1 ใน 4 ช่วงเวลาที่กำหนดที่ใกล้ที่สุด
    # (มีผลเฉพาะข้อมูลที่ส่งไปชีทปลายทาง ไม่แก้ค่าจริงในชีทต้นทาง)
    transformed_rows = []
    for row in data_rows:
        row = list(row)
        if len(row) > time_col_index:
            row[time_col_index] = snap_to_nearest_slot(row[time_col_index])
        transformed_rows.append(row)

    print(f"  พบ {len(transformed_rows)} แถวจากต้นทาง กำลังส่งไปยังชีทปลายทาง ...")

    try:
        target_sh = get_open_spreadsheet("TARGET_GSHEET_ID")
        target_ws = get_or_create_worksheet(
            target_sh, TARGET_SHEET_NAME,
            rows=len(transformed_rows) + 100, cols=len(header) + 2,
        )
        target_ws.clear()

        all_rows_to_write = [header] + transformed_rows
        target_ws.update(all_rows_to_write, value_input_option="USER_ENTERED")
    except Exception as e:
        detail = f"{type(e).__name__}: {str(e)[:200]}"
        push_to_log(source_sh, date_str, time_str, workflow_name, trigger_label, "Failed", 0, detail)
        raise

    print(f"  เขียนทับข้อมูล {len(transformed_rows)} แถว เข้าแท็บ '{TARGET_SHEET_NAME}' ในชีทปลายทางสำเร็จ")
    push_to_log(source_sh, date_str, time_str, workflow_name, trigger_label,
                "Success", len(transformed_rows), "")


if __name__ == "__main__":
    try:
        sync_once()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่าง sync:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
