"""
capture_settrade.py
--------------------
Capture ข้อมูล 4 ตารางในแท็บ "ภาพรวม Top 20" ของหน้า
https://www.settrade.com/th/equities/market-summary/top-ranking/overview
  - มูลค่าการซื้อขายสูงสุด (Most Active Value)
  - ปริมาณการซื้อขายสูงสุด (Most Active Volume)
  - ราคาเพิ่มขึ้นสูงสุด (Top Gainer)
  - ราคาลดลงสูงสุด (Top Loser)
ทั้ง 2 ตลาด (SET และ mai) รวมเป็น 8 ตาราง แล้วส่งเข้า Google Sheet โดยตรง
(ไม่บันทึกไฟล์ถาวรบนเครื่อง เพราะออกแบบมาให้รันบน GitHub Actions ซึ่งเป็นเครื่อง
ชั่วคราว)

วิธีทำงาน:
1. เปิดหน้าเว็บด้วย headless browser (Playwright) เพราะข้อมูลตารางถูกโหลด
   ด้วย JavaScript (client-side rendering)
2. รอจนกว่า placeholder (โครงตารางเปล่า) ที่ "มองเห็นอยู่จริง" จะหายไปหมด
   (หน้านี้ฝังตารางของทุกแท็บไว้ใน HTML พร้อมกัน แต่โหลดข้อมูลจริงเฉพาะแท็บ/
   ตลาดที่กำลังเปิดดูอยู่เท่านั้น จึงต้องดูเฉพาะของที่มองเห็นได้จริงบนจอ)
3. ดึงเฉพาะตารางที่ "มองเห็นอยู่จริง" (4 ตารางของแท็บ ภาพรวม Top 20) ตอนอยู่
   โหมด SET แล้วคลิกปุ่มสลับเป็น mai รอโหลดแล้ว capture ซ้ำอีกรอบ
4. ส่งแต่ละตารางเข้า Google Sheet คนละ worksheet ชื่อสื่อความหมาย เช่น
   SET_Most_Active_Value, mai_Top_Loser โดย "append" แถวใหม่พร้อม timestamp
   ต่อท้ายของเดิม -> เก็บเป็นประวัติสะสมทุกครั้งที่ capture

Environment variables ที่ต้องตั้ง (ตั้งผ่าน GitHub Actions secrets):
    GSHEET_ID                      -> ID ของ Google Sheet ปลายทาง
    GOOGLE_APPLICATION_CREDENTIALS -> path ของไฟล์ service-account JSON

การใช้งาน (ทดสอบในเครื่องตัวเอง):
    export GSHEET_ID="....."
    export GOOGLE_APPLICATION_CREDENTIALS="./gcp-credentials.json"
    python capture_settrade.py
"""

import os
import sys
import traceback
import datetime as dt

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

URL = "https://www.settrade.com/th/equities/market-summary/top-ranking/overview"

# ลำดับตารางในแท็บ "ภาพรวม Top 20" ตามที่ปรากฏบนหน้าเว็บ (ซ้าย->ขวา, บน->ล่าง)
TABLE_NAMES = ["Most_Active_Value", "Most_Active_Volume", "Top_Gainer", "Top_Loser"]

MARKETS = ["SET", "mai"]


def wait_for_visible_data_ready(page):
    """รอจนกว่า placeholder (โครงตารางเปล่า) ที่มองเห็นอยู่จริงจะหายไปหมด"""
    for _ in range(20):
        if page.locator(".placeholder:visible").count() == 0:
            return
        page.wait_for_timeout(1000)
    print("  คำเตือน: ยังมี placeholder ที่มองเห็นอยู่ หลังรอ 20 วินาที "
          "(ข้อมูลบางส่วนอาจว่างเปล่า)")


def switch_market(page, market: str):
    """คลิกปุ่มสลับตลาด SET/mai ที่มุมขวาบนของหน้า"""
    try:
        page.get_by_role("button", name=market, exact=True).first.click(timeout=10000)
    except Exception:
        try:
            page.get_by_text(market, exact=True).first.click(timeout=10000)
        except Exception:
            print(f"  คำเตือน: หาไม่ปุ่ม/ลิงก์ '{market}' เพื่อสลับตลาดไม่เจอ "
                  "(อาจต้องปรับ selector)")


def capture_visible_tables(page, market_label: str):
    """ดึง 4 ตารางที่มองเห็นอยู่จริงบนจอตอนนี้ คืนค่าเป็น list ของ (ชื่อ, DataFrame)"""
    visible_tables = page.locator("table:visible")
    count = visible_tables.count()
    results = []
    for i in range(count):
        try:
            outer_html = visible_tables.nth(i).evaluate("el => el.outerHTML")
        except Exception:
            continue
        try:
            parsed = pd.read_html(outer_html)
        except Exception:
            # ข้ามตารางที่ parse ไม่ได้ ไม่ให้ทั้งสคริปต์ล้มเพราะตารางเดียว
            continue
        for t in parsed:
            if t.dropna(how="all").empty:
                continue
            table_key = TABLE_NAMES[i] if i < len(TABLE_NAMES) else f"table_{i:02d}"
            results.append((f"{market_label}_{table_key}", t))
    return results


def fetch_all_tables():
    """เปิดหน้าเว็บ, capture ตาราง 4 อันของแท็บ Top 20 ทั้ง 2 ตลาด (SET, mai)"""
    all_results = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 2200})
        try:
            # ไม่ใช้ wait_until="networkidle" เพราะหน้านี้มีสคริปต์โฆษณา/ตัวติดตาม
            # ผู้ใช้เชื่อมต่อเน็ตต่อเนื่องตลอดเวลา ทำให้ไม่มีวันถึงสถานะ idle จริง ๆ
            page.goto(URL, wait_until="domcontentloaded", timeout=60000)
            try:
                page.wait_for_selector("table", timeout=30000)
            except PlaywrightTimeoutError:
                print("  คำเตือน: รอตาราง (<table>) ไม่เจอภายในเวลาที่กำหนด")

            for market in MARKETS:
                print(f"  กำลัง capture ตลาด {market} ...")
                switch_market(page, market)
                wait_for_visible_data_ready(page)
                results = capture_visible_tables(page, market)
                print(f"    ได้ {len(results)} ตารางจากตลาด {market}")
                all_results.extend(results)
        finally:
            browser.close()
    return all_results


def push_to_gsheet(named_tables, timestamp: str):
    """เขียนแต่ละตารางเข้า Google Sheet คนละ worksheet (ตั้งชื่อตาม market+table)"""
    import gspread
    from google.oauth2.service_account import Credentials

    sheet_id = os.environ.get("GSHEET_ID")
    creds_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")

    if not sheet_id or not creds_path:
        raise RuntimeError(
            "ไม่พบ GSHEET_ID หรือ GOOGLE_APPLICATION_CREDENTIALS ใน environment "
            "variables กรุณาตั้งค่าก่อนรัน (ดู README.md)"
        )

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(sheet_id)

    for ws_name, table in named_tables:
        ws_name = ws_name[:100]  # กัน worksheet name ยาวเกิน limit ของ Sheets
        try:
            ws = sh.worksheet(ws_name)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=ws_name, rows=2000, cols=30)
            header = ["capture_time"] + [str(c) for c in table.columns]
            ws.append_row(header, value_input_option="USER_ENTERED")

        rows = table.fillna("").astype(str).values.tolist()
        rows_with_ts = [[timestamp] + row for row in rows]
        ws.append_rows(rows_with_ts, value_input_option="USER_ENTERED")
        print(f"  ส่ง {len(rows_with_ts)} แถว เข้า worksheet '{ws_name}'")

    print(f"  เสร็จสิ้น: ส่งข้อมูลเข้า Google Sheet (id: {sheet_id}) เรียบร้อย")


def capture_once():
    now = dt.datetime.now()
    timestamp_str = now.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp_str}] เริ่ม capture ข้อมูลจาก {URL}")

    named_tables = fetch_all_tables()

    if not named_tables:
        print("  ไม่พบตารางข้อมูลใด ๆ ในหน้า (อาจต้องปรับ selector หรือ wait time) "
              "-> ข้ามการส่งเข้า Google Sheet รอบนี้")
        return

    print(f"  พบตารางที่มีข้อมูลรวม {len(named_tables)} ตาราง กำลังส่งเข้า Google Sheet...")
    push_to_gsheet(named_tables, timestamp_str)


if __name__ == "__main__":
    try:
        capture_once()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่าง capture:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
