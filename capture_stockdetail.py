"""
capture_stockdetail.py
------------------------
Capture ข้อมูลหุ้นรายตัวของแต่ละหมวดธุรกิจย่อย (Sector) จาก set.or.th
ทั้ง 28 หมวดย่อย x 2 ตลาด (SET, mai) = สูงสุด 56 หน้า แล้ว append เข้าแท็บ
StockDetailDaily ใน Google Sheet "Settrade Capture Log"

capture วันละครั้ง (เหมือน SetDatabase/SectorIndex) รันหลังตลาดปิดสนิทแล้ว

URL pattern (ยืนยันจากการทดสอบจริงแล้ว - โครงสร้าง SET กับ mai ต่างกัน):
    SET (3 ระดับ, ต้องระบุทั้ง Group และ Sector):
        https://www.set.or.th/th/market/index/set/{group}/{sector}
        เช่น .../set/consump/fashion, .../set/service/comm
    mai (2 ระดับ, ไม่มี Sector ย่อย ใช้แค่ Group):
        https://www.set.or.th/th/market/index/mai/{group}
        เช่น .../mai/agro, .../mai/consump

ออกแบบให้ทนทานต่อความผิดพลาดรายหน้า: ถ้าหมวดย่อยไหนของตลาดไหนไม่มีข้อมูล
(เช่น mai ไม่มีหุ้นในหมวดนั้นเลย) หรือโหลดไม่สำเร็จ จะข้ามไปหน้าถัดไปแทนที่จะ
ทำให้ทั้งรอบ capture ล้มเหลว

Environment variables ที่ต้องตั้ง (เหมือนสคริปต์อื่นในชุดนี้):
    GSHEET_ID
    GOOGLE_APPLICATION_CREDENTIALS
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

BASE_URL = "https://www.set.or.th/th/market/index"

# แผนที่ว่าแต่ละหมวดธุรกิจย่อย (Sector) อยู่ในกลุ่มอุตสาหกรรมใหญ่ (Group) ไหน
# (ชุดเดียวกับที่ใช้ใน capture_sectorindex.py)
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

# ตัวย่อท้ายชื่อหุ้นที่ต้องตัดออก (เหมือนที่ใช้ใน capture_settrade.py)
SYMBOL_SUFFIX_TOKENS = {
    "CF", "CB", "CC", "CS", "SP", "ST", "XD", "XR", "XM", "XT",
    "XA", "XW", "NP", "NC", "NR",
}

STOCKLIST_SHEET_NAME = "StockList"
STOCKDETAIL_SHEET_NAME = "StockDetailDaily"
STOCKDETAIL_HEADERS = [
    "Date", "Market", "Group", "Sector", "Symbol",
    "Open", "High", "Low", "Last", "Chg", "Chg%", "Bid", "Ask",
    "Volume", "Value", "Trigger",
]


def slug_candidates(code: str):
    """
    สร้าง URL slug ที่เป็นไปได้จากรหัสหมวดธุรกิจ ส่วนใหญ่แค่ตัวพิมพ์เล็ก
    แต่ "PF&REIT" มีอักขระพิเศษ ไม่แน่ใจ 100% ว่าเว็บใช้ slug แบบไหน จึงลอง
    หลายแบบไล่ไปจนกว่าจะโหลดตารางหุ้นได้สำเร็จ
    """
    base = code.lower()
    if "&" in code:
        return [base.replace("&", ""), base.replace("&", "-"), base]
    return [base]


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


def clean_column_name(col) -> str:
    """ตัดข้อความหัวตารางที่ซ้ำกันออก (เหมือนที่ใช้ใน capture_settrade.py)"""
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
    """ตัดตัวย่อท้ายชื่อหุ้นออก เทียบกับ StockList ก่อนตัด กันตัดผิด (เช่น SCB, TACC)"""
    s = str(raw_symbol).strip().upper()
    if known_symbols and s in known_symbols:
        return s
    parts = s.split()
    while len(parts) > 1 and parts[-1] in SYMBOL_SUFFIX_TOKENS:
        candidate = " ".join(parts[:-1])
        if not known_symbols or candidate in known_symbols:
            parts = parts[:-1]
        else:
            break
    return " ".join(parts).strip()


def _dash_to_zero(val):
    s = str(val).strip()
    return "0" if s in ("-", "", "nan", "NaN") else val


def _dash_to_blank(val):
    s = str(val).strip()
    return "" if s in ("-", "nan", "NaN") else val


def load_stocklist_symbols(sh):
    """โหลดรายชื่อ Symbol ทั้งหมดจากแท็บ StockList (คอลัมน์ A) ไว้เทียบกันตัด
    ตัวย่อท้ายชื่อหุ้นผิด คืนค่า set ว่างถ้าหาแท็บไม่เจอ (ไม่ทำให้ capture ล้ม)"""
    try:
        ws = sh.worksheet(STOCKLIST_SHEET_NAME)
        col_a = ws.col_values(1)[1:]  # ข้ามหัวตาราง
        return set(s.strip().upper() for s in col_a if s.strip())
    except Exception as e:
        print(f"  คำเตือน: โหลด StockList ไม่สำเร็จ -> {type(e).__name__}: {str(e)[:150]} "
              "(จะไม่ตัดตัวย่อท้ายชื่อหุ้นเลย เพื่อความปลอดภัย)")
        return set()


def extract_stock_rows(page, market_label, group_code, sector_code, known_symbols):
    """ดึงตารางหุ้นรายตัวที่มองเห็นอยู่จริงบนหน้าปัจจุบัน คืนค่า list ของ dict"""
    htmls = page.evaluate(
        """
        () => Array.from(document.querySelectorAll('table'))
            .filter(t => t.offsetParent !== null)
            .map(t => t.outerHTML)
        """
    )

    rows_out = []
    for html in htmls:
        try:
            parsed = pd.read_html(io.StringIO(html))
        except Exception:
            continue
        for t in parsed:
            cleaned_cols = [clean_column_name(c) for c in t.columns]
            # หาตารางที่มีคอลัมน์ "หลักทรัพย์" (Symbol) เท่านั้น ตัวอื่นข้ามไป
            if not any("หลักทรัพย์" in c for c in cleaned_cols):
                continue
            raw_rows = t.fillna("").astype(str).values.tolist()
            for raw_row in raw_rows:
                fields = {
                    "Symbol": "", "Open": "", "High": "", "Low": "", "Last": "",
                    "Chg": "", "ChgPct": "", "Bid": "", "Ask": "", "Volume": "", "Value": "",
                }
                for col_name, val in zip(cleaned_cols, raw_row):
                    if "หลักทรัพย์" in col_name:
                        fields["Symbol"] = clean_symbol(val, known_symbols)
                    elif "เปิด" in col_name:
                        fields["Open"] = _dash_to_blank(val)
                    elif "สูงสุด" in col_name:
                        fields["High"] = _dash_to_blank(val)
                    elif "ต่ำสุด" in col_name:
                        fields["Low"] = _dash_to_blank(val)
                    elif "ล่าสุด" in col_name:
                        fields["Last"] = _dash_to_blank(val)
                    elif "เปลี่ยนแปลง" in col_name and "%" in col_name:
                        fields["ChgPct"] = _dash_to_zero(val)
                    elif "เปลี่ยนแปลง" in col_name:
                        fields["Chg"] = _dash_to_zero(val)
                    elif "เสนอซื้อ" in col_name:
                        fields["Bid"] = _dash_to_blank(val)
                    elif "เสนอขาย" in col_name:
                        fields["Ask"] = _dash_to_blank(val)
                    elif "ปริมาณ" in col_name:
                        fields["Volume"] = _dash_to_blank(val)
                    elif "มูลค่า" in col_name:
                        fields["Value"] = _dash_to_blank(val)

                if not fields["Symbol"]:
                    continue
                fields["Market"] = market_label.upper()
                fields["Group"] = group_code
                fields["Sector"] = sector_code
                rows_out.append(fields)
    return rows_out


GROUP_CODES = sorted(set(SECTOR_TO_GROUP.values()))  # 8 กลุ่มใหญ่ (ไม่ซ้ำ)


def fetch_all_stock_details(known_symbols):
    """
    เปิดทีละหน้าดึงข้อมูลหุ้นรายตัวทั้งหมด

    โครงสร้างหน้าเว็บของ SET กับ mai ต่างกัน:
      - SET: มี 3 ระดับ Group -> Sector -> หุ้นรายตัว ต้องเข้า URL ระดับ Sector
        (เช่น set/consump/fashion) ถึงจะเห็นตารางหุ้นแบบไม่ย่อ (28 หมวดย่อย)
      - mai: มีแค่ 2 ระดับ Group -> หุ้นรายตัว ไม่มีระดับ Sector ย่อยเลย
        (เช่น mai/agro เจอตารางหุ้นเต็มทันที ไม่ต้องมี sub-path ต่อท้าย)
        ใช้แค่ 8 กลุ่มใหญ่ ไม่ใช่ 28 หมวดย่อยเหมือน SET
    """
    all_rows = []
    sector_items = list(SECTOR_TO_GROUP.items())

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 3000})
        try:
            # ----- SET: 28 หมวดย่อย (Group/Sector) -----
            for sector_code, group_code in sector_items:
                success = False
                for slug in slug_candidates(sector_code):
                    url = f"{BASE_URL}/set/{group_code.lower()}/{slug}"
                    try:
                        page.goto(url, wait_until="domcontentloaded", timeout=30000)
                        page.wait_for_selector("table", timeout=10000)
                        page.wait_for_timeout(1500)
                    except PlaywrightTimeoutError:
                        continue

                    rows = extract_stock_rows(page, "set", group_code, sector_code, known_symbols)
                    if rows:
                        all_rows.extend(rows)
                        print(f"    [SET] {group_code}/{sector_code}: ได้ {len(rows)} หุ้น")
                        success = True
                        break

                if not success:
                    print(f"    [SET] {group_code}/{sector_code}: ไม่พบข้อมูลหุ้น (ข้ามไป)")

            # ----- mai: แค่ 8 กลุ่มใหญ่ (ไม่มีระดับ Sector ย่อย) -----
            for group_code in GROUP_CODES:
                url = f"{BASE_URL}/mai/{group_code.lower()}"
                try:
                    page.goto(url, wait_until="domcontentloaded", timeout=30000)
                    page.wait_for_selector("table", timeout=10000)
                    page.wait_for_timeout(1500)
                except PlaywrightTimeoutError:
                    print(f"    [MAI] {group_code}: โหลดหน้าไม่สำเร็จ (ข้ามไป)")
                    continue

                rows = extract_stock_rows(page, "mai", group_code, "", known_symbols)
                if rows:
                    all_rows.extend(rows)
                    print(f"    [MAI] {group_code}: ได้ {len(rows)} หุ้น")
                else:
                    print(f"    [MAI] {group_code}: ไม่พบข้อมูลหุ้น (ข้ามไป)")
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


def push_to_stockdetail(sh, all_rows, date_str: str, trigger_label: str):
    try:
        ws = sh.worksheet(STOCKDETAIL_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=STOCKDETAIL_SHEET_NAME, rows=len(all_rows) + 500, cols=20)
        ws.append_row(STOCKDETAIL_HEADERS, value_input_option="USER_ENTERED", table_range="A1")

    rows_to_append = []
    for r in all_rows:
        rows_to_append.append([
            date_str, r["Market"], r["Group"], r["Sector"], r["Symbol"],
            r["Open"], r["High"], r["Low"], r["Last"], r["Chg"], r["ChgPct"],
            r["Bid"], r["Ask"], r["Volume"], r["Value"],
            trigger_label,
        ])

    if not rows_to_append:
        print("  ไม่มีแถวข้อมูลจะส่งเข้า StockDetailDaily")
        return 0

    ws.append_rows(rows_to_append, value_input_option="USER_ENTERED", table_range="A1")
    print(f"  ส่ง {len(rows_to_append)} แถว เข้า worksheet '{STOCKDETAIL_SHEET_NAME}'")
    return len(rows_to_append)


LOG_SHEET_NAME = "Log"
LOG_HEADERS = ["Date", "Time", "Workflow", "Trigger", "Status", "RowsSent", "Detail"]


def push_to_log(sh, date_str: str, time_str: str, workflow_name: str, trigger_label: str,
                 status: str, rows_sent, detail: str = ""):
    try:
        ws = sh.worksheet(LOG_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=LOG_SHEET_NAME, rows=2000, cols=10)
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
    print(f"[{date_str} {time_str} เวลาไทย] เริ่ม capture ข้อมูลหุ้นรายตัวจาก set.or.th "
          f"(workflow: {workflow_name}, trigger: {trigger_label})")

    sh = get_open_spreadsheet()
    known_symbols = load_stocklist_symbols(sh)

    try:
        all_rows = fetch_all_stock_details(known_symbols)
    except Exception as e:
        detail = f"{type(e).__name__}: {str(e)[:200]}"
        push_to_log(sh, date_str, time_str, workflow_name, trigger_label, "Failed", 0, detail)
        raise

    if not all_rows:
        print("  ไม่พบข้อมูลหุ้นรายตัวเลย -> ข้ามการส่งเข้า Google Sheet รอบนี้")
        push_to_log(sh, date_str, time_str, workflow_name, trigger_label, "NoData", 0,
                    "ไม่พบข้อมูลหุ้นรายตัว")
        return

    print(f"  พบข้อมูลหุ้นรวม {len(all_rows)} แถว กำลังส่งเข้า Google Sheet...")
    try:
        rows_sent = push_to_stockdetail(sh, all_rows, date_str, trigger_label)
    except Exception as e:
        detail = f"{type(e).__name__}: {str(e)[:200]}"
        push_to_log(sh, date_str, time_str, workflow_name, trigger_label, "Failed", 0, detail)
        raise

    push_to_log(sh, date_str, time_str, workflow_name, trigger_label, "Success", rows_sent, "")


if __name__ == "__main__":
    try:
        capture_once()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่าง capture:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
