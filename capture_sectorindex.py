"""
capture_sectorindex.py
------------------------
Capture ดัชนีราคากลุ่มอุตสาหกรรมและหมวดธุรกิจ (Industry Group / Sector Index)
จากหน้า https://www.settrade.com/th/equities/market-data/overview?category=Industry-Sector
ทั้ง 2 ตลาด (SET และ mai) แล้ว append เข้าแท็บ SectorIndex ใน Google Sheet
"Settrade Capture Log" เก็บไว้ทำ RRG (Relative Rotation Graph) / Sector
Rotation ในระบบ Dashboard หลักของผู้ใช้

capture วันละครั้ง (เหมือน SetDatabase) รันหลังตลาดปิดสนิทแล้ว (18:30 น. ไทย)

วิธีทำงาน (จากผลสำรวจ diagnostic):
1. เปิด URL แบบไม่มี query param ใด ๆ -> ได้ข้อมูลตลาด SET อัตโนมัติ (เพราะ
   แท็บ "SET" เป็นค่า default ของหน้านี้อยู่แล้ว) - ห้ามเติม &market=SET/mai
   ต่อท้าย URL เพราะทำให้หน้าพังแสดงข้อมูลผิด (ทดสอบแล้วจากรอบ diagnostic)
2. ดึงตาราง 36 แถว (8 กลุ่มอุตสาหกรรมใหญ่ + 28 หมวดธุรกิจย่อย) ของตลาด SET
3. คลิกปุ่ม "mai" (element จริงบนหน้า ไม่ใช่ URL param) รอโหลดข้อมูลใหม่
4. ดึงตารางเดิมซ้ำอีกรอบ ได้ข้อมูลตลาด mai
5. ส่งเข้า Google Sheet แท็บ SectorIndex พร้อมคอลัมน์ Level (Group/Sector)
   แยกว่าแถวไหนเป็นกลุ่มใหญ่ (8 กลุ่ม) หรือหมวดย่อย (28 หมวด)
6. บันทึกผลการรันลงแท็บ Log เดียวกับสคริปต์อื่น (Scheduled/Manual, Success/Failed)

Environment variables ที่ต้องตั้ง (เหมือนสคริปต์อื่นในชุดนี้):
    GSHEET_ID
    GOOGLE_APPLICATION_CREDENTIALS
"""

import os
import io
import re
import sys
import traceback
import datetime as dt
from zoneinfo import ZoneInfo

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

BANGKOK_TZ = ZoneInfo("Asia/Bangkok")

BASE_URL = "https://www.settrade.com/th/equities/market-data/overview?category=Industry-Sector"

# 8 รหัสกลุ่มอุตสาหกรรมใหญ่ (Industry Group) ที่เหลือทั้งหมดในตารางคือหมวด
# ธุรกิจย่อย (Sector)
GROUP_CODES = {
    "AGRO", "CONSUMP", "FINCIAL", "INDUS",
    "PROPCON", "RESOURC", "SERVICE", "TECH",
}

# แผนที่ว่าแต่ละหมวดธุรกิจย่อย (Sector) อยู่ในกลุ่มอุตสาหกรรมใหญ่ (Group) ไหน
# อ้างอิงจากโครงสร้างจริงของหน้าเว็บ settrade.com (คงที่ ไม่เปลี่ยนบ่อย)
SECTOR_TO_GROUP = {
    "AGRI": "AGRO", "FOOD": "AGRO",
    "FASHION": "CONSUMP", "HOME": "CONSUMP", "PERSON": "CONSUMP",
    "BANK": "FINCIAL", "FIN": "FINCIAL", "INSUR": "FINCIAL",
    "AUTO": "INDUS", "IMM": "INDUS", "PAPER": "INDUS",
    "PETRO": "INDUS", "PKG": "INDUS", "STEEL": "INDUS",
    "CONMAT": "PROPCON", "PROP": "PROPCON", "PF&REIT": "PROPCON", "CONS": "PROPCON",
    "ENERG": "RESOURC", "MINE": "RESOURC",
    "COMM": "SERVICE", "HELTH": "SERVICE", "MEDIA": "SERVICE",
    "PROF": "SERVICE", "TOURISM": "SERVICE", "TRANS": "SERVICE",
    "ETRON": "TECH", "ICT": "TECH",
}

SECTORINDEX_SHEET_NAME = "SectorIndex"
SECTORINDEX_HEADERS = [
    "Date", "Time", "Market", "Group", "Sector", "Name",
    "Last", "Chg", "Chg%", "Volume('000 Shares)", "Value(MB)", "Trigger",
]

LOG_SHEET_NAME = "Log"
LOG_HEADERS = ["Date", "Time", "Trigger", "Status", "RowsSent", "Detail"]

NAME_CODE_RE = re.compile(r"^(.*?)\s*\(([A-Z&]+)\)\s*$")

# จับข้อความ "ข้อมูลล่าสุด 22 ก.ค. 2569 14:36:41" ที่โชว์บนหน้าเว็บ แยกเป็น
# กลุ่มวันที่ กับ กลุ่มเวลา คนละส่วนกัน (เวลาข้อมูลจริงจากฝั่ง settrade.com
# ต่างจากเวลาที่ script รันเอง)
SOURCE_TIME_RE = re.compile(
    r"ข้อมูลล่าสุด\s+(\d{1,2}\s+\S+\.?\s+\d{4})\s+(\d{1,2}:\d{2}:\d{2})"
)


def get_source_datetime(page):
    """ดึงวันที่/เวลาจากข้อความ 'ข้อมูลล่าสุด ...' บนหน้าเว็บ คืนค่า (date, time)
    เป็น string คู่ หรือ ('', '') ถ้าหาไม่เจอ"""
    body_text = page.evaluate("() => document.body.innerText || ''")
    m = SOURCE_TIME_RE.search(body_text)
    if not m:
        return "", ""
    return m.group(1), m.group(2)


def get_trigger_label() -> str:
    event_name = os.environ.get("GITHUB_EVENT_NAME", "")
    if event_name == "schedule":
        return "Scheduled"
    elif event_name == "workflow_dispatch":
        return "Manual"
    elif event_name:
        return event_name
    return "local"


def get_visible_table_html(page):
    """คืนค่า outerHTML ของตารางแรกที่มองเห็นอยู่จริงบนหน้าปัจจุบัน (ตารางดัชนีกลุ่มอุตสาหกรรม)"""
    htmls = page.evaluate(
        """
        () => Array.from(document.querySelectorAll('table'))
            .filter(t => t.offsetParent !== null)
            .map(t => t.outerHTML)
        """
    )
    return htmls[0] if htmls else None


def parse_sector_table(html):
    """แปลง HTML ตารางเป็น DataFrame แล้วแยกชื่อ/รหัส ออกจากกัน เช่น
    'แฟชั่น (FASHION)' -> name='แฟชั่น', code='FASHION'"""
    parsed = pd.read_html(io.StringIO(html))
    if not parsed:
        return []
    df = parsed[0]

    name_col = df.columns[0]
    rows = []
    for _, row in df.iterrows():
        raw_name = str(row[name_col]).strip()
        m = NAME_CODE_RE.match(raw_name)
        if not m:
            continue  # ข้ามแถวที่ parse ชื่อ/รหัสไม่ได้ (กันโครงสร้างแปลก ๆ)
        name, code = m.group(1).strip(), m.group(2).strip()
        if code in GROUP_CODES:
            # แถวสรุปของกลุ่มใหญ่เอง -> Group = รหัสกลุ่ม, Sector เว้นว่างไว้
            group_code, sector_code = code, ""
        else:
            # แถวหมวดธุรกิจย่อย -> Group = กลุ่มใหญ่ที่สังกัด, Sector = รหัสตัวเอง
            group_code, sector_code = SECTOR_TO_GROUP.get(code, ""), code
        rows.append({
            "Group": group_code,
            "Sector": sector_code,
            "Name": name,
            "Last": row.get("ล่าสุด", ""),
            "Chg": row.get("เปลี่ยนแปลง", ""),
            "Chg%": row.get("เปลี่ยนแปลง (%)", ""),
            "Volume": row.get("ปริมาณ ('000 หุ้น)", ""),
            "Value": row.get("มูลค่า (ล้านบาท)", ""),
        })
    return rows


def wait_for_table_ready(page, timeout_s=25):
    """รอจนกว่าตารางที่มองเห็นอยู่จะมีข้อมูลจริง (ไม่ใช่ placeholder ว่างเปล่า)"""
    for _ in range(timeout_s):
        html = get_visible_table_html(page)
        if html:
            try:
                parsed = pd.read_html(io.StringIO(html))
                if parsed and not parsed[0].dropna(how="all").empty:
                    return
            except Exception:
                pass
        page.wait_for_timeout(1000)
    print("  คำเตือน: รอตารางข้อมูลพร้อมไม่สำเร็จภายในเวลาที่กำหนด")


def fetch_sector_index_data():
    """เปิดหน้า Industry-Sector ดึงข้อมูล SET ก่อน แล้วคลิกปุ่ม mai ดึงซ้ำ"""
    all_rows = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 3000})
        try:
            print("  กำลังเปิดหน้า Industry-Sector (ค่า default = ตลาด SET) ...")
            page.goto(BASE_URL, wait_until="domcontentloaded", timeout=60000)
            try:
                page.wait_for_selector("table", timeout=30000)
            except PlaywrightTimeoutError:
                print("  คำเตือน: รอตาราง (<table>) ไม่เจอภายในเวลาที่กำหนด")
            wait_for_table_ready(page)

            html = get_visible_table_html(page)
            if html:
                set_rows = parse_sector_table(html)
                src_date, src_time = get_source_datetime(page)
                print(f"    ได้ {len(set_rows)} แถวจากตลาด SET (ข้อมูลล่าสุด: {src_date} {src_time})")
                for r in set_rows:
                    r["Market"] = "SET"
                    r["Date"] = src_date
                    r["Time"] = src_time
                all_rows.extend(set_rows)
            else:
                print("  คำเตือน: ไม่พบตารางข้อมูลของตลาด SET")

            print("  กำลังคลิกปุ่ม 'mai' เพื่อสลับตลาด ...")
            try:
                page.get_by_role("button", name="mai", exact=True).first.click(timeout=10000)
                page.wait_for_timeout(1000)
                wait_for_table_ready(page)

                html = get_visible_table_html(page)
                if html:
                    mai_rows = parse_sector_table(html)
                    src_date, src_time = get_source_datetime(page)
                    print(f"    ได้ {len(mai_rows)} แถวจากตลาด mai (ข้อมูลล่าสุด: {src_date} {src_time})")
                    for r in mai_rows:
                        r["Market"] = "mai"
                        r["Date"] = src_date
                        r["Time"] = src_time
                    all_rows.extend(mai_rows)
                else:
                    print("  คำเตือน: ไม่พบตารางข้อมูลของตลาด mai")
            except Exception as e:
                print(f"  คำเตือน: คลิกปุ่ม 'mai' ไม่สำเร็จ -> {type(e).__name__}: {str(e)[:150]}")
        finally:
            browser.close()

    return all_rows


def get_open_spreadsheet():
    sheet_id = os.environ.get("GSHEET_ID")
    creds_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if not sheet_id or not creds_path:
        raise RuntimeError(
            "ไม่พบ GSHEET_ID หรือ GOOGLE_APPLICATION_CREDENTIALS ใน environment variables"
        )
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open_by_key(sheet_id)


def push_to_sectorindex(sh, all_rows, trigger_label: str):
    try:
        ws = sh.worksheet(SECTORINDEX_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SECTORINDEX_SHEET_NAME, rows=3000, cols=15)
        ws.append_row(SECTORINDEX_HEADERS, value_input_option="USER_ENTERED")

    rows_to_append = []
    for r in all_rows:
        rows_to_append.append([
            r.get("Date", ""), r.get("Time", ""),
            r["Market"], r["Group"], r["Sector"], r["Name"],
            r["Last"], r["Chg"], r["Chg%"], r["Volume"], r["Value"],
            trigger_label,
        ])

    if not rows_to_append:
        print("  ไม่มีแถวข้อมูลจะส่งเข้า SectorIndex")
        return 0

    ws.append_rows(rows_to_append, value_input_option="USER_ENTERED")
    print(f"  ส่ง {len(rows_to_append)} แถว เข้า worksheet '{SECTORINDEX_SHEET_NAME}'")
    return len(rows_to_append)


def push_to_log(sh, date_str: str, time_str: str, trigger_label: str,
                 status: str, rows_sent, detail: str = ""):
    try:
        ws = sh.worksheet(LOG_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=LOG_SHEET_NAME, rows=2000, cols=10)
        ws.append_row(LOG_HEADERS, value_input_option="USER_ENTERED")

    ws.append_row(
        [date_str, time_str, trigger_label, status, rows_sent, detail],
        value_input_option="USER_ENTERED",
    )


def capture_once():
    now = dt.datetime.now(BANGKOK_TZ)
    date_str = now.strftime("%d/%m/%Y")
    time_str = now.strftime("%H:%M")
    trigger_label = get_trigger_label()
    print(f"[{date_str} {time_str} เวลาไทย] เริ่ม capture ดัชนีกลุ่มอุตสาหกรรมจาก {BASE_URL} "
          f"(trigger: {trigger_label})")

    sh = get_open_spreadsheet()

    try:
        all_rows = fetch_sector_index_data()
    except Exception as e:
        detail = f"{type(e).__name__}: {str(e)[:200]}"
        push_to_log(sh, date_str, time_str, trigger_label, "Failed", 0, detail)
        raise

    if not all_rows:
        print("  ไม่พบข้อมูลดัชนีกลุ่มอุตสาหกรรมเลย -> ข้ามการส่งเข้า Google Sheet รอบนี้")
        push_to_log(sh, date_str, time_str, trigger_label, "NoData", 0,
                    "ไม่พบข้อมูลดัชนีกลุ่มอุตสาหกรรม")
        return

    try:
        rows_sent = push_to_sectorindex(sh, all_rows, trigger_label)
    except Exception as e:
        detail = f"{type(e).__name__}: {str(e)[:200]}"
        push_to_log(sh, date_str, time_str, trigger_label, "Failed", 0, detail)
        raise

    push_to_log(sh, date_str, time_str, trigger_label, "Success", rows_sent, "")


if __name__ == "__main__":
    try:
        capture_once()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่าง capture:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
