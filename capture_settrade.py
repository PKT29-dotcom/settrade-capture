"""
capture_settrade.py
--------------------
Capture ข้อมูล 4 ตารางในแท็บ "ภาพรวม Top 20" ของหน้า
https://www.settrade.com/th/equities/market-summary/top-ranking/overview
  - มูลค่าการซื้อขายสูงสุด (Most Active Value)
  - ปริมาณการซื้อขายสูงสุด (Most Active Volume)
  - ราคาเพิ่มขึ้นสูงสุด (Top Gainer)
  - ราคาลดลงสูงสุด (Top Loser)
ทั้ง 2 ตลาด (SET และ mai) แล้วส่งเข้า Google Sheet ของเราเอง (คนละไฟล์กับ
Dashboard หลักที่ผู้ใช้ทำเอง) ในรูปแบบ "Master" ตาราง long format แถวต่อแถว
เดียวกับ schema ของ TopDatabase ในไฟล์ Dashboard หลัก (Date, Time, Index,
TopType, Rank, Symbol, Sector, Volume, Value, Last, Chg, Chg%) เพื่อให้ผู้ใช้
copy ไปวางใน TopDatabase จริงได้สะดวกเวลามีเวลาว่าง

ระบบนี้ทำหน้าที่เป็น "ตัวสำรอง" อัตโนมัติ สำหรับวันที่ผู้ใช้ติดธุระไม่ได้อยู่
หน้าจอเพื่อ capture ข้อมูลเองด้วยมือตามปกติ

(ไม่บันทึกไฟล์ถาวรบนเครื่อง เพราะออกแบบมาให้รันบน GitHub Actions ซึ่งเป็นเครื่อง
ชั่วคราว)

วิธีทำงาน:
1. เปิดหน้าเว็บด้วย headless browser (Playwright) เพราะข้อมูลตารางถูกโหลด
   ด้วย JavaScript (client-side rendering)
2. รอจนกว่า placeholder (โครงตารางเปล่า) ที่ "มองเห็นอยู่จริง" จะหายไปหมด
3. ดึงเฉพาะตารางที่ "มองเห็นอยู่จริง" (4 ตารางของแท็บ ภาพรวม Top 20) โดยเปิด
   คนละ URL แยกกันสำหรับตลาด SET และ mai (ใช้ query parameter ?market=...
   ของเว็บ แทนการคลิกปุ่มสลับ ซึ่งเจอปัญหา race condition)
4. ดาวน์โหลดรายชื่อบริษัทจดทะเบียนจาก SET (Symbol + กลุ่มอุตสาหกรรม +
   หมวดธุรกิจ) อัปเดตแท็บ Sector_Map ในไฟล์เดียวกัน (ทุกครั้งที่ capture)
5. แปลงข้อมูลทั้งหมดเป็นแถว Master format แล้ว append เข้าแท็บ Master
   คอลัมน์ Sector เป็นสูตร VLOOKUP อ้างอิง Sector_Map ให้อัตโนมัติ
6. เติมช่อง Value/Volume ที่ขาดหายไปในแต่ละแถว โดยจับคู่ Symbol ข้ามตาราง
   ภายใน batch capture เดียวกัน (เช่น หุ้นที่โผล่ทั้งในตาราง Most Active Value
   และ Most Active Volume พร้อมกัน จะได้ค่าครบทั้งคู่)

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

import requests
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
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

# ป้ายชื่อประเภทตาราง (TopType) ตามลำดับที่ปรากฏบนหน้าเว็บ (ซ้าย->ขวา, บน->ล่าง)
# ใช้ข้อความเดียวกับที่ปรากฏในคอลัมน์ TopType ของ TopDatabase (มีช่องว่าง ไม่ใช่ _)
TABLE_TYPE_LABELS = ["Most Active Value", "Most Active Volume", "Top Gainer", "Top Loser"]

# ไฟล์รายชื่อบริษัทจดทะเบียนทางการจาก SET (Symbol + กลุ่มอุตสาหกรรม + หมวดธุรกิจ)
SET_LISTED_COMPANIES_XLS_URL = (
    "https://www.set.or.th/dat/eod/listedcompany/static/listedCompanies_th_TH.xls"
)
SECTOR_MAP_SHEET_NAME = "Sector_Map"

MASTER_SHEET_NAME = "Master"
MASTER_HEADERS = [
    "Date", "Time", "Index", "TopType", "Rank", "Symbol",
    "Sector", "Volume", "Value", "Last", "Chg", "Chg%",
]


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
    (2) ไม่มี placeholder ที่มองเห็นอยู่เหลือแล้ว
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
    """ดึงตารางที่มองเห็นอยู่จริงบนจอตอนนี้ (ผ่าน JS)
    คืนค่าเป็น list ของ (market_label, table_type_label, DataFrame)"""
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
        try:
            parsed = pd.read_html(io.StringIO(outer_html))
        except Exception as e:
            print(f"    [diag] ตาราง #{i}: parse ไม่ผ่าน -> {type(e).__name__}: {str(e)[:150]}")
            continue
        for t in parsed:
            if t.dropna(how="all").empty:
                continue
            table_type = TABLE_TYPE_LABELS[i] if i < len(TABLE_TYPE_LABELS) else f"table_{i:02d}"
            results.append((market_label, table_type, t))
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


def extract_row_fields(cleaned_cols, row):
    """
    แปลงแถวดิบ (cleaned_cols คู่กับค่าจริงใน row) ให้เป็น dict ตาม schema ของ
    Master: Rank, Symbol, Value, Volume, Last, Chg, ChgPct
    ตารางแต่ละประเภทมีแค่ Value หรือ Volume อย่างใดอย่างหนึ่ง (ไม่ใช่ทั้งคู่) เพราะ
    ดึงมาจากหน้า Top Ranking โดยตรง อีกช่องจะเว้นว่างไว้
    """
    fields = {"Rank": "", "Symbol": "", "Value": "", "Volume": "", "Last": "", "Chg": "", "ChgPct": ""}
    for col_name, val in zip(cleaned_cols, row):
        if "อันดับ" in col_name:
            fields["Rank"] = val
        elif "ชื่อย่อ" in col_name or "หลักทรัพย์" in col_name:
            fields["Symbol"] = val
        elif "มูลค่า" in col_name:
            fields["Value"] = val
        elif "ปริมาณ" in col_name:
            fields["Volume"] = val
        elif "เปลี่ยนแปลง" in col_name and "%" in col_name:
            fields["ChgPct"] = val
        elif "เปลี่ยนแปลง" in col_name:
            fields["Chg"] = val
        elif "ราคา" in col_name and "ล่าสุด" in col_name:
            fields["Last"] = val
    return fields


def _find_column_by_keywords(columns, keywords):
    """หาชื่อคอลัมน์ที่มีคำใน keywords ปนอยู่ (ไม่สนตัวพิมพ์เล็ก-ใหญ่) คืนชื่อคอลัมน์จริง หรือ None"""
    for col in columns:
        col_str = str(col)
        for kw in keywords:
            if kw.lower() in col_str.lower():
                return col
    return None


def update_sector_map(sh):
    """
    ดาวน์โหลดรายชื่อบริษัทจดทะเบียนทั้งหมดจากเว็บทางการของ SET (Symbol +
    กลุ่มอุตสาหกรรม + หมวดธุรกิจ) แล้วอัปเดตแท็บ Sector_Map ทั้งหมด (เขียนทับ
    ของเดิม เพราะเป็นตารางอ้างอิงล่าสุด ไม่ใช่ประวัติสะสม)

    ออกแบบให้ "ไม่ทำให้การ capture หลักล้ม" หากขั้นตอนนี้ผิดพลาด จะแค่ print
    คำเตือนแล้วข้ามไป ไม่ raise exception ออกไปนอกฟังก์ชันนี้
    """
    print("  กำลังอัปเดต Sector_Map จากรายชื่อบริษัทจดทะเบียนของ SET ...")
    try:
        resp = requests.get(
            SET_LISTED_COMPANIES_XLS_URL,
            headers={"User-Agent": "Mozilla/5.0 (compatible; SettradeCaptureBot/1.0)"},
            timeout=30,
        )
        resp.raise_for_status()
        df = pd.read_excel(io.BytesIO(resp.content))
    except Exception as e:
        print(f"    คำเตือน: ดาวน์โหลด/อ่านไฟล์รายชื่อบริษัทจาก SET ไม่สำเร็จ -> "
              f"{type(e).__name__}: {str(e)[:150]} (ข้ามการอัปเดต Sector_Map รอบนี้)")
        return

    print(f"    [diag] คอลัมน์ที่พบในไฟล์: {list(df.columns)}")

    symbol_col = _find_column_by_keywords(df.columns, ["symbol", "หลักทรัพย์", "ย่อ"])
    industry_col = _find_column_by_keywords(df.columns, ["กลุ่มอุตสาหกรรม", "industry"])
    sector_col = _find_column_by_keywords(df.columns, ["หมวดธุรกิจ", "sector"])

    if symbol_col is None:
        print("    คำเตือน: หาคอลัมน์ชื่อย่อหลักทรัพย์ในไฟล์ไม่เจอ "
              "(ข้ามการอัปเดต Sector_Map รอบนี้)")
        return

    out = pd.DataFrame({
        "Symbol": df[symbol_col].astype(str).str.strip(),
        "IndustryGroup": df[industry_col].astype(str).str.strip() if industry_col is not None else "",
        "Sector": df[sector_col].astype(str).str.strip() if sector_col is not None else "",
    }).drop_duplicates(subset=["Symbol"])

    try:
        ws = sh.worksheet(SECTOR_MAP_SHEET_NAME)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SECTOR_MAP_SHEET_NAME, rows=len(out) + 10, cols=5)

    rows = [out.columns.tolist()] + out.astype(str).values.tolist()
    ws.update(rows, value_input_option="USER_ENTERED")
    print(f"    อัปเดต Sector_Map สำเร็จ ({len(out)} บริษัท)")


def build_cross_reference_lookups(results):
    """
    สร้าง lookup dict ต่อตลาด สำหรับเติม Value/Volume ที่ขาดหายไป โดยจับคู่ Symbol
    ข้ามตารางภายใน batch capture เดียวกัน (เช่น หุ้นที่โผล่ทั้งในตาราง Most
    Active Value และ Most Active Volume พร้อมกัน จะเอาค่าที่ขาดจากอีกตาราง
    มาเติมให้ครบทั้งคู่)

    คืนค่า: {market: {"value": {symbol: value}, "volume": {symbol: volume}}}
    """
    lookups = {}
    for market_label, table_type, table in results:
        if market_label not in lookups:
            lookups[market_label] = {"value": {}, "volume": {}}
        cleaned_cols = [clean_column_name(c) for c in table.columns]
        raw_rows = table.fillna("").astype(str).values.tolist()
        for raw_row in raw_rows:
            f = extract_row_fields(cleaned_cols, raw_row)
            symbol = f["Symbol"]
            if not symbol:
                continue
            if f["Value"] and symbol not in lookups[market_label]["value"]:
                lookups[market_label]["value"][symbol] = f["Value"]
            if f["Volume"] and symbol not in lookups[market_label]["volume"]:
                lookups[market_label]["volume"][symbol] = f["Volume"]
    return lookups


def push_to_master(sh, results, date_str: str, time_str: str):
    """
    เขียนข้อมูลทั้งหมดลงแท็บ Master เดียว แบบแถวต่อแถว (long format) ตรงกับ
    schema ของ TopDatabase (Date, Time, Index, TopType, Rank, Symbol, Sector,
    Volume, Value, Last, Chg, Chg%) คอลัมน์ Sector เป็นสูตร VLOOKUP อ้างอิง
    แท็บ Sector_Map ในไฟล์เดียวกันให้อัตโนมัติ
    """
    try:
        ws = sh.worksheet(MASTER_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=MASTER_SHEET_NAME, rows=2000, cols=15)
        ws.append_row(MASTER_HEADERS, value_input_option="USER_ENTERED")

    existing_row_count = len(ws.get_all_values())
    start_row = existing_row_count + 1

    symbol_col_index = MASTER_HEADERS.index("Symbol") + 1  # 1-based
    symbol_col_letter = gspread.utils.rowcol_to_a1(1, symbol_col_index).rstrip("0123456789")

    lookups = build_cross_reference_lookups(results)

    rows_to_append = []
    for market_label, table_type, table in results:
        cleaned_cols = [clean_column_name(c) for c in table.columns]
        raw_rows = table.fillna("").astype(str).values.tolist()
        for raw_row in raw_rows:
            row_num = start_row + len(rows_to_append)
            f = extract_row_fields(cleaned_cols, raw_row)
            symbol = f["Symbol"]
            market_lookup = lookups.get(market_label, {"value": {}, "volume": {}})
            if not f["Value"]:
                f["Value"] = market_lookup["value"].get(symbol, "")
            if not f["Volume"]:
                f["Volume"] = market_lookup["volume"].get(symbol, "")
            sector_formula = (
                f'=IFERROR(VLOOKUP({symbol_col_letter}{row_num},'
                f'{SECTOR_MAP_SHEET_NAME}!A:C,3,FALSE),"")'
            )
            rows_to_append.append([
                date_str, time_str, market_label, table_type,
                f["Rank"], f["Symbol"], sector_formula,
                f["Volume"], f["Value"], f["Last"], f["Chg"], f["ChgPct"],
            ])

    if not rows_to_append:
        print("  ไม่มีแถวข้อมูลจะส่งเข้า Master")
        return

    ws.append_rows(rows_to_append, value_input_option="USER_ENTERED")
    print(f"  ส่ง {len(rows_to_append)} แถว เข้า worksheet '{MASTER_SHEET_NAME}'")


def get_open_spreadsheet():
    """สร้าง gspread client และเปิด Google Sheet ปลายทาง (ใช้ร่วมกันทั้งสองขั้นตอน)"""
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
    return gc.open_by_key(sheet_id)


def capture_once():
    now = dt.datetime.now(BANGKOK_TZ)
    date_str = now.strftime("%d/%m/%Y")
    time_str = now.strftime("%H:%M")
    print(f"[{date_str} {time_str} เวลาไทย] เริ่ม capture ข้อมูลจาก {BASE_URL}")

    sh = get_open_spreadsheet()

    # อัปเดต Sector_Map ก่อน (ถ้าพลาดจะแค่ print คำเตือน ไม่กระทบงานหลักด้านล่าง)
    try:
        update_sector_map(sh)
    except Exception:
        print("  คำเตือน: อัปเดต Sector_Map ไม่สำเร็จ (ไม่กระทบการ capture หลัก):")
        traceback.print_exc(limit=3, file=sys.stdout)

    results = fetch_all_tables()

    if not results:
        print("  ไม่พบตารางข้อมูลใด ๆ ในหน้า (อาจต้องปรับ selector หรือ wait time) "
              "-> ข้ามการส่งเข้า Google Sheet รอบนี้")
        return

    print(f"  พบตารางที่มีข้อมูลรวม {len(results)} ตาราง กำลังส่งเข้า Google Sheet...")
    push_to_master(sh, results, date_str, time_str)


if __name__ == "__main__":
    try:
        capture_once()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่าง capture:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
