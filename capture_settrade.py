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
Dashboard หลักที่ผู้ใช้ทำเอง) ในรูปแบบแท็บ "TopDatabase" ตาราง long format แถวต่อแถว
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
4. ตัดตัวย่อท้ายชื่อหุ้น (เช่น "TRITN CB" -> "TRITN") ให้ Symbol ตรงกับที่เก็บ
   ในแท็บ StockList ของผู้ใช้เอง (คัดลอกมาไว้ในไฟล์เดียวกันแล้ว)
5. แปลงข้อมูลทั้งหมดเป็นแถว TopDatabase format แล้ว append เข้าแท็บ TopDatabase
   คอลัมน์ Sector เป็นสูตร VLOOKUP อ้างอิงแท็บ StockList ให้อัตโนมัติ
6. เติมช่อง Value/Volume ที่ขาดหายไปในแต่ละแถว โดยจับคู่ Symbol ข้ามตาราง
   ภายใน batch capture เดียวกัน (เช่น หุ้นที่โผล่ทั้งในตาราง Most Active Value
   และ Most Active Volume พร้อมกัน จะได้ค่าครบทั้งคู่)
7. บันทึกประวัติการรันทุกครั้ง (สำเร็จ/ไม่พบข้อมูล/ล้มเหลว) ลงแท็บ Log แยก
   ต่างหาก พร้อมระบุว่าเป็นการรันแบบ Scheduled (อัตโนมัติ) หรือ Manual (กด
   Run workflow เอง) เพื่อช่วยตรวจสอบว่า schedule ทำงานตรงเวลาไหม

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
    "MAI": f"{BASE_URL}?market=mai&securityType=Common+Stock",  # URL param ต้องเป็น mai ตัวเล็กตามที่เว็บ settrade.com กำหนด แต่ label แสดงผลใช้ MAI ตัวใหญ่ให้เข้าชุดกับ SET
}

# ป้ายชื่อประเภทตาราง (TopType) ตามลำดับที่ปรากฏบนหน้าเว็บ (ซ้าย->ขวา, บน->ล่าง)
# ใช้ข้อความเดียวกับที่ปรากฏในคอลัมน์ TopType ของ TopDatabase (มีช่องว่าง ไม่ใช่ _)
TABLE_TYPE_LABELS = ["Most Active Value", "Most Active Volume", "Top Gainer", "Top Loser"]

# ตารางอ้างอิง Symbol -> Sector ของผู้ใช้เอง (คัดลอกมาไว้ในไฟล์เดียวกันแล้ว)
# คอลัมน์ A = Symbol, คอลัมน์ B = Sector
STOCKLIST_SHEET_NAME = "StockList"

# ตัวย่อท้ายชื่อหุ้นที่ต้องตัดออก เพื่อให้ Symbol ตรงกับที่เก็บใน StockList
# (เช่น "TRITN CB" -> "TRITN") ใช้ชุดตัวย่อเดียวกับสูตรที่ผู้ใช้ใช้อยู่แล้ว
SYMBOL_SUFFIX_TOKENS = {
    "CF", "CB", "CC", "CS", "SP", "ST", "XD", "XR", "XM", "XT",
    "XA", "XW", "NP", "NC", "NR",
}

MASTER_SHEET_NAME = "TopDatabase"
MASTER_HEADERS = [
    "Date", "Time", "Index", "TopType", "Rank", "Symbol",
    "Sector", "Volume", "Value", "Last", "Chg", "Chg%", "Trigger",
]

# แท็บเก็บประวัติการรันทุกครั้ง (แยกจากข้อมูล Master) ไว้ตรวจสอบว่า schedule
# ทำงานตรงเวลาไหม รอบไหน fail บ้าง เพื่อวิเคราะห์ปัญหาความล่าช้าของ
# GitHub Actions schedule ได้ง่ายขึ้น
LOG_SHEET_NAME = "Log"
LOG_HEADERS = ["Date", "Time", "Trigger", "Status", "RowsSent", "Detail"]


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


def clean_symbol(raw_symbol: str, known_symbols=None) -> str:
    """
    ตัดตัวย่อท้ายชื่อหุ้นออก เช่น "TGPROCB" -> "TGPRO" ให้ตรงกับรูปแบบ Symbol
    ที่เก็บใน StockList โดยยึด known_symbols (ชุด Symbol จริงจาก StockList)
    เป็นหลักฐานยืนยันก่อนตัดเสมอ เพื่อไม่ให้ตัดผิดกับชื่อหุ้นที่ลงท้ายด้วย
    ตัวอักษรกลุ่มเดียวกับ badge พอดีอยู่แล้ว (เช่น SCB, TACC, TBSP)

    รองรับทั้งกรณี badge ติดกับชื่อไม่มีช่องว่างคั่น (เช่น "TGPROCB") และกรณี
    มีช่องว่างคั่น (เช่น "TRITN CB") เพราะ normalize เอาช่องว่างออกก่อนเช็คเสมอ
    รองรับหลาย badge ต่อกัน (เช่น "GRANDCBCSCC" -> "GRAND") ด้วยการลองตัด
    หลายรอบ (สูงสุด 3 badge) แล้วเลือกคำตอบที่ตัดน้อยที่สุดที่ยังตรงกับ
    StockList จริง

    ถ้าไม่มี known_symbols ให้ตรวจสอบ (เช่น โหลด StockList ไม่สำเร็จ) จะไม่ตัด
    อะไรเลย ปล่อยให้ VLOOKUP ฝั่ง Google Sheets หา Sector ไม่เจอแทน (ปลอดภัยกว่า
    การเดาตัดผิด)
    """
    s = raw_symbol.strip().upper().replace(" ", "")

    if not known_symbols:
        return s

    if s in known_symbols:
        return s

    best_candidate = None

    def try_strip(remaining, stripped_count):
        nonlocal best_candidate
        if remaining in known_symbols and len(remaining) < len(s):
            if best_candidate is None or len(remaining) > len(best_candidate):
                best_candidate = remaining
        if stripped_count >= 3 or len(remaining) <= 2:
            return
        for token in SYMBOL_SUFFIX_TOKENS:
            if remaining.endswith(token):
                try_strip(remaining[: -len(token)], stripped_count + 1)

    try_strip(s, 0)
    return best_candidate if best_candidate else s


def _dash_to_zero(val):
    """เว็บ settrade.com แสดง '-' แทน 0.00 เวลาราคาไม่มีการเปลี่ยนแปลง
    แปลงเป็นเลข 0 ก่อนส่งเข้า Google Sheet เพื่อให้เป็นตัวเลขที่คำนวณต่อได้"""
    s = str(val).strip()
    return "0" if s in ("-", "", "nan", "NaN") else val


def extract_row_fields(cleaned_cols, row, known_symbols=None):
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
            fields["Symbol"] = clean_symbol(val, known_symbols)
        elif "มูลค่า" in col_name:
            fields["Value"] = val
        elif "ปริมาณ" in col_name:
            fields["Volume"] = val
        elif "เปลี่ยนแปลง" in col_name and "%" in col_name:
            fields["ChgPct"] = _dash_to_zero(val)
        elif "เปลี่ยนแปลง" in col_name:
            fields["Chg"] = _dash_to_zero(val)
        elif "ราคา" in col_name and "ล่าสุด" in col_name:
            fields["Last"] = val
    return fields


def build_cross_reference_lookups(results, known_symbols=None):
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
            f = extract_row_fields(cleaned_cols, raw_row, known_symbols)
            symbol = f["Symbol"]
            if not symbol:
                continue
            if f["Value"] and symbol not in lookups[market_label]["value"]:
                lookups[market_label]["value"][symbol] = f["Value"]
            if f["Volume"] and symbol not in lookups[market_label]["volume"]:
                lookups[market_label]["volume"][symbol] = f["Volume"]
    return lookups


def load_stocklist_symbols(sh):
    """
    โหลดชุด Symbol ที่ถูกต้องทั้งหมดจากแท็บ StockList (คอลัมน์ A) เพื่อใช้เป็น
    หลักฐานยืนยันตอนตัด badge ท้ายชื่อหุ้น (เช่น "TGPROCB" -> "TGPRO")
    ถ้าโหลดไม่สำเร็จจะคืนค่า None (clean_symbol จะไม่ตัดอะไรเลยแทน ปลอดภัยกว่า
    การเดาตัดผิด)
    """
    try:
        ws = sh.worksheet(STOCKLIST_SHEET_NAME)
        col_a = ws.col_values(1)
    except Exception as e:
        print(f"  คำเตือน: โหลด StockList ไม่สำเร็จ -> {type(e).__name__}: {str(e)[:150]} "
              "(จะไม่ตัด badge ท้ายชื่อหุ้น ปล่อยให้ VLOOKUP หา Sector ไม่เจอแทน)")
        return None
    symbols = {s.strip().upper() for s in col_a[1:] if s.strip()}  # ข้าม header แถวแรก
    print(f"  โหลด StockList สำเร็จ ({len(symbols)} Symbol) สำหรับตรวจสอบการตัด badge")
    return symbols


def push_to_master(sh, results, date_str: str, time_str: str, trigger_label: str):
    """
    เขียนข้อมูลทั้งหมดลงแท็บ Master เดียว แบบแถวต่อแถว (long format) ตรงกับ
    schema ของ TopDatabase (Date, Time, Index, TopType, Rank, Symbol, Sector,
    Volume, Value, Last, Chg, Chg%) คอลัมน์ Sector เป็นสูตร VLOOKUP อ้างอิง
    แท็บ StockList ในไฟล์เดียวกันให้อัตโนมัติ คอลัมน์ Trigger บอกว่าแถวนี้มาจาก
    schedule อัตโนมัติ หรือมาจากการกด Run workflow เอง (ไว้ช่วยวิเคราะห์ปัญหา
    เรื่องความล่าช้า/การข้ามรอบของ GitHub Actions schedule)
    """
    known_symbols = load_stocklist_symbols(sh)

    try:
        ws = sh.worksheet(MASTER_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=MASTER_SHEET_NAME, rows=2000, cols=15)
        ws.append_row(MASTER_HEADERS, value_input_option="USER_ENTERED")

    existing_row_count = len(ws.get_all_values())
    start_row = existing_row_count + 1

    symbol_col_index = MASTER_HEADERS.index("Symbol") + 1  # 1-based
    symbol_col_letter = gspread.utils.rowcol_to_a1(1, symbol_col_index).rstrip("0123456789")

    lookups = build_cross_reference_lookups(results, known_symbols)

    rows_to_append = []
    for market_label, table_type, table in results:
        cleaned_cols = [clean_column_name(c) for c in table.columns]
        raw_rows = table.fillna("").astype(str).values.tolist()
        for raw_row in raw_rows:
            row_num = start_row + len(rows_to_append)
            f = extract_row_fields(cleaned_cols, raw_row, known_symbols)
            symbol = f["Symbol"]
            market_lookup = lookups.get(market_label, {"value": {}, "volume": {}})
            if not f["Value"]:
                f["Value"] = market_lookup["value"].get(symbol, "")
            if not f["Volume"]:
                f["Volume"] = market_lookup["volume"].get(symbol, "")
            sector_formula = (
                f'=IFERROR(VLOOKUP({symbol_col_letter}{row_num},'
                f'{STOCKLIST_SHEET_NAME}!A:B,2,FALSE),"")'
            )
            rows_to_append.append([
                date_str, time_str, market_label, table_type,
                f["Rank"], f["Symbol"], sector_formula,
                f["Volume"], f["Value"], f["Last"], f["Chg"], f["ChgPct"],
                trigger_label,
            ])

    if not rows_to_append:
        print("  ไม่มีแถวข้อมูลจะส่งเข้า Master")
        return

    ws.append_rows(rows_to_append, value_input_option="USER_ENTERED")
    print(f"  ส่ง {len(rows_to_append)} แถว เข้า worksheet '{MASTER_SHEET_NAME}'")


def push_to_log(sh, date_str: str, time_str: str, trigger_label: str,
                 status: str, rows_sent, detail: str = ""):
    """
    บันทึกประวัติการรัน 1 แถวต่อ 1 ครั้งที่รัน ลงแท็บ Log แยกจาก Master
    เพื่อให้ตรวจสอบย้อนหลังได้ว่า schedule ทำงานตรงเวลาไหม รอบไหนหายไป/fail
    บันทึกทุกครั้งไม่ว่าจะสำเร็จ ล้มเหลว หรือไม่พบข้อมูล เพื่อให้เห็นภาพครบ
    """
    try:
        ws = sh.worksheet(LOG_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=LOG_SHEET_NAME, rows=2000, cols=10)
        ws.append_row(LOG_HEADERS, value_input_option="USER_ENTERED")

    ws.append_row(
        [date_str, time_str, trigger_label, status, rows_sent, detail],
        value_input_option="USER_ENTERED",
    )


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


def get_trigger_label() -> str:
    """
    อ่านประเภทของ trigger จาก environment variable ที่ GitHub Actions ตั้งให้
    อัตโนมัติ (GITHUB_EVENT_NAME) เพื่อบอกว่าแถวนี้มาจาก schedule อัตโนมัติ
    หรือมาจากการกด Run workflow เอง ไว้ช่วยวิเคราะห์ปัญหาความล่าช้า/การข้าม
    รอบของ schedule ได้ง่ายขึ้น (ถ้ารันในเครื่องตัวเอง ไม่มีตัวแปรนี้ จะขึ้น
    เป็น "local" แทน)
    """
    event_name = os.environ.get("GITHUB_EVENT_NAME", "")
    if event_name == "schedule":
        return "Scheduled"
    elif event_name == "workflow_dispatch":
        return "Manual"
    elif event_name:
        return event_name
    return "local"


def capture_once():
    now = dt.datetime.now(BANGKOK_TZ)
    date_str = now.strftime("%d/%m/%Y")
    time_str = now.strftime("%H:%M")
    trigger_label = get_trigger_label()
    print(f"[{date_str} {time_str} เวลาไทย] เริ่ม capture ข้อมูลจาก {BASE_URL} "
          f"(trigger: {trigger_label})")

    # เปิด Google Sheet ก่อน เพื่อให้บันทึก Log ได้แม้ขั้นตอนถัดไปจะล้มเหลว
    sh = get_open_spreadsheet()

    try:
        results = fetch_all_tables()
    except Exception as e:
        detail = f"{type(e).__name__}: {str(e)[:200]}"
        push_to_log(sh, date_str, time_str, trigger_label, "Failed", 0, detail)
        raise

    if not results:
        print("  ไม่พบตารางข้อมูลใด ๆ ในหน้า (อาจต้องปรับ selector หรือ wait time) "
              "-> ข้ามการส่งเข้า Google Sheet รอบนี้")
        push_to_log(sh, date_str, time_str, trigger_label, "NoData", 0,
                    "ไม่พบตารางข้อมูลในหน้าเว็บ")
        return

    print(f"  พบตารางที่มีข้อมูลรวม {len(results)} ตาราง กำลังส่งเข้า Google Sheet...")
    total_rows = sum(len(t) for _, _, t in results)
    try:
        push_to_master(sh, results, date_str, time_str, trigger_label)
    except Exception as e:
        detail = f"{type(e).__name__}: {str(e)[:200]}"
        push_to_log(sh, date_str, time_str, trigger_label, "Failed", 0, detail)
        raise

    push_to_log(sh, date_str, time_str, trigger_label, "Success", total_rows, "")


if __name__ == "__main__":
    try:
        capture_once()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่าง capture:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
