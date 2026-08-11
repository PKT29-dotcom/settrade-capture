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
3. "ตัดข้อมูลซ้ำ" หลังปัดเวลาแล้ว ถ้าวันเดียวกัน มีมากกว่า 1 ครั้งที่ปัดตกลง
   slot เดียวกัน (เช่น รอบ Scheduled ที่ล่าช้า + รอบ Manual ที่กดช่วยเสริม
   ในวันเดียวกัน ทั้งคู่ปัดเข้า 16:30 เหมือนกัน) จะเก็บไว้แค่ชุดล่าสุด ไม่ให้
   ข้อมูลซ้ำซ้อนไปที่ปลายทาง (คีย์เทียบ: Date+Time+Index+TopType+Rank+Symbol)

   หมายเหตุ: การปัดเวลา/ตัดซ้ำนี้มีผลเฉพาะข้อมูลที่ส่งไปให้ชีทปลายทางเท่านั้น
   ไม่ได้แก้ไข/ลบข้อมูลจริงในชีทต้นทาง (Settrade Capture Log) เพื่อให้ยังคง
   เห็นประวัติการรันทุกครั้งไว้สำหรับตรวจสอบ/วิเคราะห์

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
import re
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

# คีย์ที่ใช้เทียบว่าแถวไหนคือ "สล็อตเดียวกัน" ของวันเดียวกัน (ไว้ตัดซ้ำ)
DEDUP_KEY_COLUMNS = ["Date", "Time", "Index", "TopType", "Rank", "Symbol"]


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


def dedup_rows(header, rows):
    """
    ตัดแถวที่ซ้ำกันหลังปัดเวลาแล้ว โดยเทียบคีย์ DEDUP_KEY_COLUMNS เก็บไว้แค่
    แถวล่าสุด (ตัวท้ายสุดที่เจอ) ต่อ 1 คีย์ เพราะ TopDatabase ต้นทางเรียงตาม
    ลำดับเวลาที่ capture จริง แถวที่มาทีหลังคือข้อมูลใหม่กว่าเสมอ

    ถ้าหาคอลัมน์ที่ต้องใช้ทำคีย์ไม่ครบ (โครงสร้างเปลี่ยนไปจากที่คาด) จะข้าม
    การตัดซ้ำไปเลย คืนค่า rows เดิมกลับไปแทน เพื่อไม่ให้ sync ทั้งหมดล้มเหลว
    """
    try:
        col_indices = [header.index(c) for c in DEDUP_KEY_COLUMNS]
    except ValueError as e:
        print(f"  คำเตือน: หาคอลัมน์สำหรับตัดซ้ำไม่ครบ ({e}) -> ข้ามการตัดซ้ำรอบนี้")
        return rows

    deduped = {}
    for row in rows:
        key = tuple(row[i] if i < len(row) else "" for i in col_indices)
        deduped[key] = row  # แถวหลังทับแถวก่อน -> เก็บแถวล่าสุดของแต่ละคีย์ไว้

    removed = len(rows) - len(deduped)
    if removed > 0:
        print(f"  ตัดแถวซ้ำออก {removed} แถว (สล็อตเวลาเดียวกันถูก capture มากกว่า 1 ครั้ง)")

    return list(deduped.values())


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
    Google บอกว่ามีแท็บชื่อนี้อยู่แล้วจริง แต่ list ที่ดึงมาหาไม่เจอ
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

    # ตัดแถวซ้ำที่เกิดจากหลายรอบ capture ปัดตกลง slot เดียวกันในวันเดียวกัน
    # (เช่น Scheduled ล่าช้า + Manual ที่กดช่วยเสริม)
    transformed_rows = dedup_rows(header, transformed_rows)

    print(f"  พบ {len(transformed_rows)} แถวหลังตัดซ้ำ กำลังส่งไปยังชีทปลายทาง ...")

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
