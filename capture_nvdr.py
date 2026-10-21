"""
capture_nvdr.py
-----------------
Capture NVDR Trading by Stock (รายหุ้น) จาก set.or.th แล้ว append เข้าแท็บ
NVDRTradingByStock ใน Google Sheet "Settrade Capture Log"

วิธีทำงาน:
1. เปิดหน้า NVDR Trading by Stock แล้วกดปุ่ม "Export Excel" ดาวน์โหลดไฟล์
   (ไฟล์มีทั้ง Volume (Share) และ Value (Baht) ของทุกหุ้นในไฟล์เดียว)
2. อ่านวันที่จากข้อความ "As of dd Mon yyyy" ในไฟล์ แล้วใช้เป็นคอลัมน์ Date
   ของทุกแถว (เป็นวันที่ของข้อมูลจริง ไม่ใช่วันที่ script รัน)
3. ถ้าแท็บ NVDRTradingByStock มีข้อมูลของวันที่นั้นอยู่แล้ว จะไม่ append ซ้ำ
   (บันทึก Log ว่า AlreadyCaptured) จึงรันซ้ำกี่รอบก็ได้ และวันหยุดตลาดก็จะ
   ข้ามเองโดยปริยาย เพราะหน้าเว็บยังโชว์ข้อมูลของวันทำการล่าสุดเดิม
4. ถ้าหัวตารางในไฟล์ไม่ตรงกับที่คาด (เว็บเปลี่ยนรูปแบบ) จะหยุดและบันทึก Log
   ว่า Failed แทนที่จะเดาตำแหน่งคอลัมน์ผิดแล้วบันทึกข้อมูลเพี้ยน

Environment variables:
    GSHEET_ID
    GOOGLE_APPLICATION_CREDENTIALS
"""

import os
import io
import re
import sys
import numbers
import traceback
import datetime as dt
from zoneinfo import ZoneInfo

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

BANGKOK_TZ = ZoneInfo("Asia/Bangkok")

# URL นี้เดาจากหน้าข้างเคียง (.../nvdr/trading-by-sector) ยังไม่ได้ยืนยันกับ
# หน้าจริง ถ้าเปิดแล้วไม่ใช่หน้านี้ ให้แก้ค่านี้ตัวเดียว
NVDR_URL = "https://www.set.or.th/th/market/statistics/nvdr/trading-by-stock"

NVDR_SHEET_NAME = "NVDRTradingByStock"
NVDR_HEADERS = [
    "Date", "Symbol",
    "Vol.Buy", "Vol.Sell", "Vol.Total", "Vol.Net", "Vol%",
    "Val.Buy", "Val.Sell", "Val.Total", "Val.Net", "Val%",
    "Trigger",
]

LOG_SHEET_NAME = "Log"
LOG_HEADERS = ["Date", "Time", "Workflow", "Trigger", "Status", "RowsSent", "Detail"]

# หัวตารางแถวย่อยที่คาดไว้ (Volume 5 ช่อง ตามด้วย Value 5 ช่อง) ตามไฟล์ตัวอย่าง
EXPECTED_SUBHEADERS = ["buy", "sell", "total", "net", "%"] * 2

ASOF_RE = re.compile(r"As of\s+(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})")
EN_MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}
SYMBOL_RE = re.compile(r"^[A-Z0-9][A-Z0-9&.\-]{0,19}$")
NOT_SYMBOLS = {"TOTAL", "SUM", "REMARK", "NOTE"}


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


def _to_number(v):
    """ช่องว่าง/NaN/'-' -> 0 ตัวเลขเต็ม -> int ที่เหลือคงทศนิยมเดิม"""
    if v is None:
        return 0
    if isinstance(v, numbers.Real):
        f = float(v)
        if f != f:  # NaN
            return 0
        return int(f) if f.is_integer() else f
    s = str(v).replace(",", "").strip()
    if s in ("", "-"):
        return 0
    try:
        f = float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return 0


def parse_nvdr_excel(content: bytes):
    """คืนค่า (date_str 'dd/mm/yyyy', rows) โดย rows = [[symbol, 10 ตัวเลข], ...]
    ถ้าโครงสร้างไฟล์ไม่ตรงที่คาด จะ raise ValueError พร้อมบอกว่าเจออะไร"""
    df = pd.read_excel(io.BytesIO(content), header=None, dtype=object)

    # 1) วันที่ของข้อมูล
    as_of = None
    for r in range(min(len(df), 12)):
        for c in range(df.shape[1]):
            v = df.iat[r, c]
            if isinstance(v, str):
                m = ASOF_RE.search(v)
                if m:
                    month = EN_MONTHS.get(m.group(2)[:3].lower())
                    if month:
                        as_of = dt.date(int(m.group(3)), month, int(m.group(1)))
                        break
        if as_of:
            break
    if as_of is None:
        raise ValueError("หาข้อความ 'As of dd Mon yyyy' ในไฟล์ Excel ไม่เจอ")

    # 2) แถวหัวตาราง
    sym_idx = next((i for i in range(len(df))
                    if isinstance(df.iat[i, 0], str) and df.iat[i, 0].strip() == "Symbol"), None)
    if sym_idx is None:
        raise ValueError("หาหัวคอลัมน์ 'Symbol' ในไฟล์ Excel ไม่เจอ")
    if df.shape[1] < 11:
        raise ValueError(f"ไฟล์มี {df.shape[1]} คอลัมน์ ต้องการอย่างน้อย 11")

    sub_idx = None
    for i in range(sym_idx, min(sym_idx + 4, len(df))):
        cells = [str(df.iat[i, c]).replace("*", "").strip().lower() for c in range(1, 11)]
        if cells == EXPECTED_SUBHEADERS:
            sub_idx = i
            break
    if sub_idx is None:
        seen = [[str(df.iat[i, c]).strip() for c in range(1, 11)]
                for i in range(sym_idx, min(sym_idx + 3, len(df)))]
        raise ValueError(f"หัวตารางย่อยไม่ตรงที่คาด (Buy,Sell,Total,Net,% x2) เจอ: {seen}")

    # 3) แถวข้อมูล
    rows = []
    for i in range(sub_idx + 1, len(df)):
        sym = df.iat[i, 0]
        if not isinstance(sym, str):
            continue
        s = sym.strip()
        if s in NOT_SYMBOLS or not SYMBOL_RE.match(s):
            continue
        rows.append([s] + [_to_number(df.iat[i, c]) for c in range(1, 11)])

    return as_of.strftime("%d/%m/%Y"), rows


def download_excel_bytes():
    """เปิดหน้า NVDR Trading by Stock แล้วกด Export Excel คืนค่าเนื้อไฟล์"""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 2400})
        try:
            page.goto(NVDR_URL, wait_until="domcontentloaded", timeout=60000)
            try:
                page.wait_for_selector("table", timeout=30000)
            except PlaywrightTimeoutError:
                print("  คำเตือน: รอตาราง (<table>) ไม่เจอภายในเวลาที่กำหนด")
            page.wait_for_timeout(3000)

            with page.expect_download(timeout=60000) as dl_info:
                try:
                    page.get_by_role("button", name=re.compile("Export Excel", re.I)).first.click(timeout=10000)
                except Exception:
                    page.get_by_text("Export Excel").first.click(timeout=10000)
            download = dl_info.value
            with open(download.path(), "rb") as f:
                return f.read()
        finally:
            browser.close()


def get_open_spreadsheet():
    sheet_id = os.environ.get("GSHEET_ID")
    creds_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if not sheet_id or not creds_path:
        raise RuntimeError("ไม่พบ GSHEET_ID หรือ GOOGLE_APPLICATION_CREDENTIALS ใน environment variables")
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    return gspread.authorize(creds).open_by_key(sheet_id)


def get_or_create_worksheet(sh, title: str, rows: int, cols: int):
    """หาแท็บแบบ normalize ชื่อ (กันบั๊กชื่อแท็บที่เคยเจอ) ไม่เจอค่อยสร้างใหม่"""
    def norm(s):
        return re.sub(r"[^a-zA-Z0-9]", "", s).lower()

    def find():
        for ws in sh.worksheets():
            if norm(ws.title) == norm(title):
                return ws
        return None

    ws = find()
    if ws:
        return ws
    try:
        return sh.add_worksheet(title=title, rows=rows, cols=cols)
    except gspread.exceptions.APIError as e:
        if "already exists" not in str(e):
            raise
        ws = find()
        if ws:
            return ws
        raise


def date_already_captured(ws, date_str: str) -> bool:
    try:
        return date_str in {c.lstrip("'").strip() for c in ws.col_values(1)}
    except Exception:
        return False


def push_to_log(sh, date_str, time_str, workflow_name, trigger_label, status, rows_sent, detail=""):
    ws = get_or_create_worksheet(sh, LOG_SHEET_NAME, rows=2000, cols=10)
    if len(ws.get_all_values()) == 0:
        ws.append_row(LOG_HEADERS, value_input_option="USER_ENTERED", table_range="A1")
    ws.append_row(
        [date_str, time_str, workflow_name, trigger_label, status, rows_sent, detail],
        value_input_option="USER_ENTERED",
        table_range="A1",
    )


def capture_once():
    now = dt.datetime.now(BANGKOK_TZ)
    date_str = now.strftime("%d/%m/%Y")
    time_str = now.strftime("%H:%M")
    trigger_label = get_trigger_label()
    workflow_name = get_workflow_name()
    print(f"[{date_str} {time_str} เวลาไทย] เริ่ม capture NVDR Trading by Stock "
          f"(workflow: {workflow_name}, trigger: {trigger_label})")

    sh = get_open_spreadsheet()

    try:
        content = download_excel_bytes()
        data_date, rows = parse_nvdr_excel(content)
    except Exception as e:
        detail = f"{type(e).__name__}: {str(e)[:300]}"
        push_to_log(sh, date_str, time_str, workflow_name, trigger_label, "Failed", 0, detail)
        raise

    print(f"  ไฟล์ Excel เป็นข้อมูลวันที่ {data_date} จำนวน {len(rows)} หุ้น")

    if not rows:
        push_to_log(sh, date_str, time_str, workflow_name, trigger_label, "NoData", 0,
                    f"ไฟล์ข้อมูลวันที่ {data_date} ไม่มีแถวหุ้น")
        return

    ws = get_or_create_worksheet(sh, NVDR_SHEET_NAME, rows=len(rows) + 1000, cols=15)
    if len(ws.get_all_values()) == 0:
        ws.append_row(NVDR_HEADERS, value_input_option="USER_ENTERED", table_range="A1")

    if date_already_captured(ws, data_date):
        print(f"  มีข้อมูลวันที่ {data_date} อยู่แล้ว -> ไม่บันทึกซ้ำ")
        push_to_log(sh, date_str, time_str, workflow_name, trigger_label, "AlreadyCaptured", 0,
                    f"มีข้อมูล NVDR ของวันที่ {data_date} อยู่แล้ว")
        return

    # ขึ้นต้นด้วย ' ให้ Sheets เก็บวันที่เป็นข้อความล้วน (เหมือน SectorIndex)
    out = [[f"'{data_date}"] + r + [trigger_label] for r in rows]
    try:
        ws.append_rows(out, value_input_option="USER_ENTERED", table_range="A1")
    except Exception as e:
        detail = f"{type(e).__name__}: {str(e)[:300]}"
        push_to_log(sh, date_str, time_str, workflow_name, trigger_label, "Failed", 0, detail)
        raise

    print(f"  ส่ง {len(out)} แถว เข้า worksheet '{NVDR_SHEET_NAME}'")
    push_to_log(sh, date_str, time_str, workflow_name, trigger_label, "Success", len(out),
                f"ข้อมูลวันที่ {data_date}")


if __name__ == "__main__":
    try:
        capture_once()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่าง capture:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
