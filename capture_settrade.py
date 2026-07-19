"""
capture_settrade.py
--------------------
Capture ข้อมูล Top Ranking / สรุปภาวะตลาด SET, mai จากหน้า
https://www.settrade.com/th/equities/market-summary/top-ranking/overview
แล้วส่งผลลัพธ์เข้า Google Sheet โดยตรง (ไม่บันทึกไฟล์ถาวรบนเครื่อง เพราะออกแบบ
มาให้รันบน GitHub Actions ซึ่งเป็นเครื่องชั่วคราว)

วิธีทำงาน:
1. เปิดหน้าเว็บด้วย headless browser (Playwright) เพราะข้อมูลตารางถูกโหลด
   ด้วย JavaScript (client-side rendering)
2. รอให้ตารางข้อมูลโหลดเสร็จ แล้วดึง HTML ที่ render แล้วออกมา
3. ใช้ pandas.read_html() แกะตาราง <table> ทั้งหมดออกมาอัตโนมัติ
4. ส่งแต่ละตาราง (ที่มีข้อมูลจริง ไม่ว่างเปล่า) เข้า Google Sheet คนละ worksheet
   โดย "append" แถวใหม่พร้อม timestamp ต่อท้ายของเดิม -> เก็บเป็นประวัติสะสม
   ทุกครั้งที่ capture

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
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

URL = "https://www.settrade.com/th/equities/market-summary/top-ranking/overview"


def fetch_rendered_html() -> str:
    """เปิดหน้าเว็บด้วย headless browser แล้วคืนค่า HTML ที่ render สมบูรณ์แล้ว"""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 2200})
        try:
            page.goto(URL, wait_until="networkidle", timeout=60000)
            try:
                page.wait_for_selector("table", timeout=20000)
            except PlaywrightTimeoutError:
                print("  คำเตือน: รอตาราง (<table>) ไม่เจอภายในเวลาที่กำหนด")
            page.wait_for_timeout(3000)
            html = page.content()
        finally:
            browser.close()
    return html


def extract_tables(html: str):
    """
    แกะเฉพาะแท็ก <table> ออกจาก HTML ก่อน (ด้วย BeautifulSoup) แล้วค่อยแปลงทีละ
    ตารางด้วย pandas วิธีนี้ทนทานกว่าการยัด HTML ทั้งหน้าเข้า pd.read_html()
    รวดเดียว เพราะหน้า settrade.com มี HTML จากสคริปต์โฆษณา/ตัวติดตามผู้ใช้
    ปนอยู่เยอะมาก ถ้า parser สะดุดตรงไหนตรงหนึ่งจะทำให้ตารางทั้งหมดพังไปด้วย
    ตัดปัญหานี้โดยแยก parse ทีละตารางเล็ก ๆ แทน
    """
    soup = BeautifulSoup(html, "lxml")
    table_tags = soup.find_all("table")

    tables = []
    for tag in table_tags:
        try:
            parsed = pd.read_html(str(tag))
        except (ValueError, ImportError):
            continue
        for t in parsed:
            if not t.dropna(how="all").empty:
                tables.append(t)
    return tables


def push_to_gsheet(tables, timestamp: str):
    """เขียนแต่ละตารางเข้า Google Sheet คนละ worksheet, append แถวใหม่ต่อท้าย"""
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

    for i, table in enumerate(tables):
        ws_name = f"table_{i:02d}"
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

    html = fetch_rendered_html()
    tables = extract_tables(html)

    if not tables:
        print("  ไม่พบตารางข้อมูลใด ๆ ในหน้า (อาจต้องปรับ selector หรือ wait time) "
              "-> ข้ามการส่งเข้า Google Sheet รอบนี้")
        return

    print(f"  พบตารางที่มีข้อมูล {len(tables)} ตาราง กำลังส่งเข้า Google Sheet...")
    push_to_gsheet(tables, timestamp_str)


if __name__ == "__main__":
    try:
        capture_once()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่าง capture:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
