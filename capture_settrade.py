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
3. ดึงเฉพาะตารางที่ "มองเห็นอยู่จริง" (4 ตารางของแท็บ ภาพรวม Top 20) โดยเปิด
   คนละ URL แยกกันสำหรับตลาด SET และ mai (ใช้ query parameter ?market=...
   ของเว็บ แทนการคลิกปุ่มสลับ ซึ่งเจอปัญหา race condition)
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
import io
import sys
import traceback
import datetime as dt
from zoneinfo import ZoneInfo

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

BANGKOK_TZ = ZoneInfo("Asia/Bangkok")

BASE_URL = "https://www.settrade.com/th/equities/market-summary/top-ranking/overview"

# เปิดคนละ URL แยกกันสำหรับแต่ละตลาดโดยตรง (แทนการคลิกปุ่มสลับตลาดบนหน้าเว็บ
# ซึ่งเจอปัญหา race condition ตอนหน้าเว็บถอด-สร้างตารางใหม่) เว็บ settrade.com
# รองรับ query parameter "market" อยู่แล้ว เชื่อถือได้กว่ามาก
MARKET_URLS = {
    "SET": f"{BASE_URL}?market=SET&securityType=Common+Stock",
    "mai": f"{BASE_URL}?market=mai&securityType=Common+Stock",
}

# ลำดับตารางในแท็บ "ภาพรวม Top 20" ตามที่ปรากฏบนหน้าเว็บ (ซ้าย->ขวา, บน->ล่าง)
TABLE_NAMES = ["Most_Active_Value", "Most_Active_Volume", "Top_Gainer", "Top_Loser"]


def get_visibility_status(page):
    """
    เช็คสถานะการมองเห็นด้วย JavaScript โดยตรง (เชื่อถือได้กว่าการพึ่งพา
    CSS pseudo-selector ':visible' ของ Playwright เพียงอย่างเดียว) คืนค่า
    จำนวนตารางที่มองเห็นอยู่จริง และจำนวน placeholder ที่ยังมองเห็นอยู่
    """
    return page.evaluate(
        """
        () => {
            const isVisible = (el) => el.offsetParent !== null;
            const tables = Array.from(document.querySelectorAll('table')).filter(isVisible);
            const placeholders = Array.from(document.querySelectorAll('.placeholder')).filter(isVisible);
            return {tableCount: tables.length, placeholderCount: placeholders.length};
        }
        """
    )


def wait_for_tables_ready(page, min_tables=1, timeout_s=25):
    """
    รอจนกว่า (1) มีตารางที่มองเห็นอยู่จริงอย่างน้อย min_tables ตาราง และ
    (2) ไม่มี placeholder ที่มองเห็นอยู่เหลือแล้ว ก่อนหน้านี้เช็คแค่เงื่อนไข (2)
    อย่างเดียว ทำให้เกิด race condition ตอนคลิกสลับตลาด: เว็บถอด-สร้างตาราง
    ใหม่ชั่วขณะ ทำให้ตอนเช็คไม่มีตารางอยู่เลย (placeholder=0 เพราะไม่มี element
    ให้เช็คเลย ไม่ใช่เพราะโหลดเสร็จ) เลยผ่านเงื่อนไขไปแบบผิด ๆ ทันที
    """
    status = {"tableCount": 0, "placeholderCount": 0}
    for _ in range(timeout_s):
        status = get_visibility_status(page)
        if status["tableCount"] >= min_tables and status["placeholderCount"] == 0:
            return
        page.wait_for_timeout(1000)
    print(f"  คำเตือน: หลังรอ {timeout_s} วินาที "
          f"ตารางที่มองเห็น={status['tableCount']}, "
          f"placeholder ที่ยังค้าง={status['placeholderCount']}")


def capture_visible_tables(page, market_label: str):
    """ดึงตารางที่มองเห็นอยู่จริงบนจอตอนนี้ (ผ่าน JS) คืนค่าเป็น list ของ (ชื่อ, DataFrame)"""
    diag = page.evaluate(
        """
        () => {
            const all = Array.from(document.querySelectorAll('table'));
            const visible = all.filter(t => t.offsetParent !== null);
            return {
                totalCount: all.length,
                visibleCount: visible.length,
                htmls: visible.map(t => t.outerHTML),
            };
        }
        """
    )
    print(f"    [diag] เจอ <table> ทั้งหมด {diag['totalCount']} ตัว, "
          f"มองเห็นได้ {diag['visibleCount']} ตัว")

    results = []
    for i, outer_html in enumerate(diag["htmls"]):
        print(f"    [diag] ตาราง #{i}: ยาว {len(outer_html)} ตัวอักษร")
        try:
            parsed = pd.read_html(io.StringIO(outer_html))
        except Exception as e:
            print(f"    [diag] ตาราง #{i}: parse ไม่ผ่าน -> {type(e).__name__}: {str(e)[:150]}")
            continue
        for t in parsed:
            print(f"    [diag] ตาราง #{i}: ขนาด {t.shape[0]} แถว x {t.shape[1]} คอลัมน์"
                  f"{' (ว่างเปล่า)' if t.dropna(how='all').empty else ''}")
            if t.dropna(how="all").empty:
                continue
            table_key = TABLE_NAMES[i] if i < len(TABLE_NAMES) else f"table_{i:02d}"
            results.append((f"{market_label}_{table_key}", t))
    return results


def fetch_all_tables():
    """เปิดหน้าเว็บทีละตลาด (คนละ URL) แล้ว capture ตาราง 4 อันของแท็บ Top 20"""
    all_results = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 2200})
        try:
            for market, market_url in MARKET_URLS.items():
                print(f"  กำลัง capture ตลาด {market} ...")
                # ไม่ใช้ wait_until="networkidle" เพราะหน้านี้มีสคริปต์โฆษณา/
                # ตัวติดตามผู้ใช้เชื่อมต่อเน็ตต่อเนื่องตลอดเวลา ทำให้ไม่มีวันถึง
                # สถานะ idle จริง ๆ
                page.goto(market_url, wait_until="domcontentloaded", timeout=60000)
                try:
                    page.wait_for_selector("table", timeout=30000)
                except PlaywrightTimeoutError:
                    print("  คำเตือน: รอตาราง (<table>) ไม่เจอภายในเวลาที่กำหนด")

                wait_for_tables_ready(page, min_tables=1, timeout_s=25)
                results = capture_visible_tables(page, market)
                print(f"    ได้ {len(results)} ตารางจากตลาด {market}")
                all_results.extend(results)
        finally:
            browser.close()
    return all_results


def clean_column_name(col) -> str:
    """
    ตัดข้อความหัวตารางที่ซ้ำกันออก เช่น "ราคาล่าสุด ราคาล่าสุด" -> "ราคาล่าสุด"
    (เว็บ settrade.com ใส่ข้อความหัวตารางซ้ำ 2 ชุดใน HTML เดียวกัน ชุดหนึ่งไว้
    แสดงผล อีกชุดซ่อนไว้สำหรับโปรแกรมอ่านหน้าจอ/screen reader)
    """
    if isinstance(col, tuple):
        parts = [str(p) for p in col if str(p) not in ("", "nan")]
        deduped = []
        for p in parts:
            if not deduped or deduped[-1] != p:
                deduped.append(p)
        col_str = " ".join(deduped)
    else:
        col_str = str(col)

    stripped = col_str.replace(" ", "")
    half = len(stripped) // 2
    if half > 0 and len(stripped) % 2 == 0 and stripped[:half] == stripped[half:]:
        # หาตำแหน่งใน col_str (ที่ยังมีช่องว่างเดิมอยู่) ที่ตรงกับความยาวครึ่งแรก
        # แบบไม่นับช่องว่าง เพื่อคงช่องว่างที่ตั้งใจไว้จริงของข้อความส่วนแรก
        count = 0
        cut_index = len(col_str)
        for idx, ch in enumerate(col_str):
            if ch != " ":
                count += 1
            if count == half:
                cut_index = idx + 1
                break
        col_str = col_str[:cut_index]
    return col_str.strip()


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
            header = ["capture_time"] + [clean_column_name(c) for c in table.columns]
            ws.append_row(header, value_input_option="USER_ENTERED")

        rows = table.fillna("").astype(str).values.tolist()
        rows_with_ts = [[timestamp] + row for row in rows]
        ws.append_rows(rows_with_ts, value_input_option="USER_ENTERED")
        print(f"  ส่ง {len(rows_with_ts)} แถว เข้า worksheet '{ws_name}'")

    print(f"  เสร็จสิ้น: ส่งข้อมูลเข้า Google Sheet (id: {sheet_id}) เรียบร้อย")


def capture_once():
    now = dt.datetime.now(BANGKOK_TZ)
    timestamp_str = now.strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp_str} เวลาไทย] เริ่ม capture ข้อมูลจาก {BASE_URL}")

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
