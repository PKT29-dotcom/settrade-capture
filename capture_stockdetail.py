"""
Capture StockDailyDetail  v1.1.0  (capture_stockdetail.py)
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

วันที่ของข้อมูล (v1.1.0): ใช้วันที่จากแถบ "ข้อมูลล่าสุด dd ม.ค. yyyy HH:MM:SS" ด้านบนของหน้า
set.or.th (ปี พ.ศ. แปลงเป็น ค.ศ. ให้) เป็นคอลัมน์ Date แทนวันที่ที่ script รัน ถ้ารันช้าข้าม
เที่ยงคืนหรือรันวันหยุด ข้อมูลจะไม่ถูกลงเป็นวันใหม่ผิด ๆ แล้วตรวจก่อนบันทึก 2 ขั้น:
  1) ถ้าแท็บมีแถววันที่นั้นอยู่แล้ว -> ไม่บันทึกซ้ำ (Log: AlreadyCaptured)
  2) ถ้าตัวเลขหุ้นเหมือนวันก่อนหน้าเกือบทุกตัว -> ไม่บันทึก (Log: DuplicateOfPrevious)
ถ้าหาวันที่บนหน้าเว็บไม่เจอเลย จะไม่บันทึก (Log: Failed) แทนการใช้วันที่ผิด

Environment variables ที่ต้องตั้ง (เหมือนสคริปต์อื่นในชุดนี้):
    GSHEET_ID
    GOOGLE_APPLICATION_CREDENTIALS

Changelog (เพิ่มเวอร์ชันทุกครั้งที่แก้ไฟล์นี้ และเวอร์ชันจะแสดงใน Log คอลัมน์ Detail):
  1.1.0  ใช้วันที่จากเว็บเป็น Date, ตรวจซ้ำวันเดิม/ซ้ำวันก่อน, ลองใหม่อัตโนมัติเมื่อเชื่อมต่อ
         Google Sheets หลุด, แสดงเวอร์ชันใน Log
  1.0.0  เวอร์ชันเดิม ใช้วันที่ที่ script รันเป็น Date
"""

import os
import io
import re
import sys
import time
import traceback
from collections import Counter
import datetime as dt
from zoneinfo import ZoneInfo

import pandas as pd
import gspread
import requests
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

SCRIPT_NAME = "Capture StockDailyDetail"
__version__ = "1.1.0"

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
STOCKDETAIL_SHEET_NAME = "StockDailyDetail"
STOCKDETAIL_HEADERS = [
    "Date", "Time", "Market", "Group", "Sector", "Symbol",
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


# ---------- ลองใหม่อัตโนมัติเมื่อเชื่อมต่อ Google Sheets หลุดชั่วคราว ----------
RETRY_ATTEMPTS = 5
RETRY_BASE_DELAY = 2


def _is_transient(e) -> bool:
    """ข้อผิดพลาดชั่วคราวที่ควรลองใหม่ (เชื่อมต่อหลุด, หมดเวลา, 429, 5xx)"""
    if isinstance(e, (requests.exceptions.ConnectionError, requests.exceptions.Timeout,
                      requests.exceptions.ChunkedEncodingError, ConnectionError, TimeoutError)):
        return True
    if isinstance(e, gspread.exceptions.APIError):
        code = getattr(getattr(e, "response", None), "status_code", None)
        return code in (429, 500, 502, 503, 504)
    return False


def _backoff(attempt: int, e) -> None:
    delay = min(RETRY_BASE_DELAY * 2 ** (attempt - 1), 30)
    print(f"  คำเตือน: เชื่อมต่อ Google Sheets ไม่สำเร็จ ({type(e).__name__}) "
          f"ลองใหม่ครั้งที่ {attempt + 1}/{RETRY_ATTEMPTS} ใน {delay} วินาที")
    time.sleep(delay)


def call_with_retry(fn, *args, **kwargs):
    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            if not _is_transient(e) or attempt == RETRY_ATTEMPTS:
                raise
            _backoff(attempt, e)


def append_safe(ws, rows, already_done):
    """เพิ่มแถวแบบกันซ้ำ: ถ้าเชื่อมต่อหลุด จะอ่านชีทดูก่อนว่ารอบที่แล้วเขียนเข้าไปแล้วหรือยัง
    (already_done = ฟังก์ชันรับค่าที่อ่านจากชีท คืน True ถ้าข้อมูลชุดนี้อยู่ในชีทแล้ว)"""
    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            ws.append_rows(rows, value_input_option="USER_ENTERED", table_range="A1")
            return
        except Exception as e:
            if not _is_transient(e) or attempt == RETRY_ATTEMPTS:
                raise
            _backoff(attempt, e)
            try:
                if already_done(call_with_retry(ws.get_all_values)):
                    print("  ตรวจพบว่าข้อมูลถูกเขียนสำเร็จไปแล้วก่อนเชื่อมต่อหลุด ไม่เขียนซ้ำ")
                    return
            except Exception:
                pass


def _is_sheet_like(x) -> bool:
    return hasattr(x, "get_all_values") or hasattr(x, "worksheets")


class _Retrying:
    """ครอบ Spreadsheet/Worksheet ของ gspread: ทุกเมธอดที่ติดต่อ Google จะลองใหม่เมื่อหลุดชั่วคราว
    ยกเว้น append_row/append_rows ที่ลองซ้ำแบบสุ่มสี่สุ่มห้าไม่ได้ (เสี่ยงข้อมูลซ้ำ) ให้ใช้ append_safe"""
    _NO_BLIND_RETRY = {"append_row", "append_rows"}

    def __init__(self, obj):
        object.__setattr__(self, "_obj", obj)

    def __getattr__(self, name):
        attr = getattr(self._obj, name)
        if not callable(attr):
            return attr

        def call(*args, **kwargs):
            res = attr(*args, **kwargs) if name in self._NO_BLIND_RETRY else call_with_retry(attr, *args, **kwargs)
            if _is_sheet_like(res):
                return _Retrying(res)
            if isinstance(res, list) and res and all(_is_sheet_like(w) for w in res):
                return [_Retrying(w) for w in res]
            return res
        return call


# ---------- วันที่ของข้อมูลตามที่หน้าเว็บ SET ระบุ ----------
SOURCE_TIME_RE = re.compile(r"ข้อมูลล่าสุด\s+(\d{1,2})\s+(\S+?)\.?\s+(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})")
LAST_UPDATE_EN_RE = re.compile(r"Last Update\s+(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})")
THAI_MONTH_ABBR = {
    "ม.ค": 1, "ก.พ": 2, "มี.ค": 3, "เม.ย": 4, "พ.ค": 5, "มิ.ย": 6,
    "ก.ค": 7, "ส.ค": 8, "ก.ย": 9, "ต.ค": 10, "พ.ย": 11, "ธ.ค": 12,
}
EN_MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
             "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}
DATE_CELL_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d %b %Y", "%d-%m-%Y")
DUPLICATE_FRACTION = 0.98       # หุ้นที่ตัวเลขเหมือนวันก่อนเกินสัดส่วนนี้ = ข้อมูลซ้ำ
MIN_ROWS_FOR_DUPLICATE_CHECK = 100

_PAGE_DATES = []                # วันที่ที่อ่านได้จากแต่ละหน้าที่โหลดสำเร็จ


def parse_source_date(text: str):
    """อ่านวันที่จากข้อความ 'ข้อมูลล่าสุด 11 พ.ย. 2569 18:33:20' (หรือ 'Last Update 11 Nov 2026')"""
    m = SOURCE_TIME_RE.search(text or "")
    if m:
        month = THAI_MONTH_ABBR.get(m.group(2).strip().rstrip("."))
        if month:
            year = int(m.group(3))
            year = year - 543 if year > 2400 else year
            try:
                return dt.date(year, month, int(m.group(1)))
            except ValueError:
                pass
    m = LAST_UPDATE_EN_RE.search(text or "")
    if m:
        month = EN_MONTHS.get(m.group(2)[:3].lower())
        if month:
            try:
                return dt.date(int(m.group(3)), month, int(m.group(1)))
            except ValueError:
                pass
    return None


def record_page_date(page):
    try:
        d = parse_source_date(page.evaluate("() => document.body.innerText || ''"))
    except Exception:
        d = None
    if d:
        _PAGE_DATES.append(d)


def choose_data_date(dates):
    """คืนค่า (วันที่ที่พบบ่อยที่สุด, หมายเหตุถ้าแต่ละหน้าไม่ตรงกัน)"""
    if not dates:
        return None, ""
    counts = Counter(dates)
    best, _ = counts.most_common(1)[0]
    note = ""
    if len(counts) > 1:
        note = " | เตือน: แต่ละหน้าแสดงวันที่ไม่ตรงกัน " + ", ".join(
            f"{d:%d/%m/%Y}x{n}" for d, n in sorted(counts.items()))
    return best, note


def parse_date_cell(s):
    s = str(s or "").strip().lstrip("'")
    for fmt in DATE_CELL_FORMATS:
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def existing_dates(values) -> set:
    return {d for d in (parse_date_cell(r[0]) for r in values[1:] if r) if d}


def _norm_value(v):
    s = str(v if v is not None else "").replace(",", "").strip()
    if s == "":
        return ""
    try:
        return round(float(s), 4)
    except ValueError:
        return s.lower()


def _snapshot_from_sheet(values):
    """ข้อมูลของวันล่าสุดที่มีในแท็บ: {(market, symbol): (last, chg, volume, value)} และวันที่นั้น"""
    idx = {h: i for i, h in enumerate(STOCKDETAIL_HEADERS)}
    last_date, rows = None, []
    for r in values[1:]:
        d = parse_date_cell(r[0]) if r else None
        if d is None:
            continue
        if last_date is None or d > last_date:
            last_date, rows = d, []
        if d == last_date and len(r) > idx["Value"]:
            rows.append(r)
    snap = {(str(r[idx["Market"]]).strip().lower(), str(r[idx["Symbol"]]).strip().upper()):
            tuple(_norm_value(r[idx[c]]) for c in ("Last", "Chg", "Volume", "Value")) for r in rows}
    return last_date, snap


def _snapshot_from_new(all_rows):
    return {(str(r["Market"]).strip().lower(), str(r["Symbol"]).strip().upper()):
            tuple(_norm_value(r[c]) for c in ("Last", "Chg", "Volume", "Value")) for r in all_rows}


def same_as_previous(values, all_rows):
    """คืนค่า (ซ้ำกับวันก่อนไหม, วันที่ของข้อมูลเดิมที่เทียบ)"""
    prev_date, prev = _snapshot_from_sheet(values)
    new = _snapshot_from_new(all_rows)
    if prev_date is None or len(new) < MIN_ROWS_FOR_DUPLICATE_CHECK or len(prev) < MIN_ROWS_FOR_DUPLICATE_CHECK:
        return False, prev_date
    same = sum(1 for k, v in new.items() if prev.get(k) == v)
    return same / len(new) >= DUPLICATE_FRACTION, prev_date


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
    """
    ตัดตัวย่อท้ายชื่อหุ้นออก รองรับกรณีมีหลายตัวย่อต่อกัน เช่น "SQ CB CS" -> "SQ"

    ปลอดภัยจากการตัดผิดกรณีหุ้นชื่อลงท้ายด้วยตัวอักษรพวกนี้พอดี (เช่น SCB,
    TACC, TBSP) เพราะเช็คว่าต้องเป็น "คำแยกต่างหาก" ที่มีช่องว่างคั่นเท่านั้น
    ถึงจะตัด ไม่ได้เช็คแค่ตัวอักษรท้ายสตริงเฉย ๆ - SCB เป็นคำเดียวไม่มีช่องว่าง
    เลยไม่ถูกตัดเลย ในขณะที่ "SQ CB CS" แยกเป็น 3 คำ ตัดวนไปทีละคำจนกว่าจะ
    ไม่เจอ badge ต่อท้ายแล้ว (รองรับ badge ซ้อนกันกี่ตัวก็ได้ ไม่จำกัด)
    """
    s = str(raw_symbol).strip().upper()
    parts = s.split()
    while len(parts) > 1 and parts[-1] in SYMBOL_SUFFIX_TOKENS:
        parts = parts[:-1]
    return " ".join(parts).strip()


def _dash_to_zero(val):
    s = str(val).strip()
    return "0" if s in ("-", "", "nan", "NaN") else val


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
                        fields["Open"] = _dash_to_zero(val)
                    elif "สูงสุด" in col_name:
                        fields["High"] = _dash_to_zero(val)
                    elif "ต่ำสุด" in col_name:
                        fields["Low"] = _dash_to_zero(val)
                    elif "ล่าสุด" in col_name:
                        fields["Last"] = _dash_to_zero(val)
                    elif "เปลี่ยนแปลง" in col_name and "%" in col_name:
                        fields["ChgPct"] = _dash_to_zero(val)
                    elif "เปลี่ยนแปลง" in col_name:
                        fields["Chg"] = _dash_to_zero(val)
                    elif "เสนอซื้อ" in col_name:
                        fields["Bid"] = _dash_to_zero(val)
                    elif "เสนอขาย" in col_name:
                        fields["Ask"] = _dash_to_zero(val)
                    elif "ปริมาณ" in col_name:
                        fields["Volume"] = _dash_to_zero(val)
                    elif "มูลค่า" in col_name:
                        fields["Value"] = _dash_to_zero(val)

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

                    record_page_date(page)
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

                record_page_date(page)
                rows = extract_stock_rows(page, "mai", group_code, "N/A", known_symbols)
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
    return _Retrying(call_with_retry(gc.open_by_key, sheet_id))


def load_stockdetail(sh):
    """คืนค่า (worksheet, ค่าทั้งหมดในแท็บ) สร้างแท็บ+หัวตารางให้ถ้ายังไม่มี"""
    try:
        ws = sh.worksheet(STOCKDETAIL_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=STOCKDETAIL_SHEET_NAME, rows=3000, cols=20)
        call_with_retry(ws.append_row, STOCKDETAIL_HEADERS, value_input_option="USER_ENTERED",
                        table_range="A1")
    return ws, ws.get_all_values()


def push_to_stockdetail(ws, values, all_rows, date_str: str, time_str: str, trigger_label: str, data_date):
    """เพิ่มแถวหุ้นทั้งหมดต่อท้ายแท็บ (Date = วันที่ของข้อมูลจากเว็บ)"""
    non_empty = [r for r in values if any((c or "").strip() for c in r)]
    start_row = len(non_empty) + 1

    symbol_col_index = STOCKDETAIL_HEADERS.index("Symbol") + 1  # 1-based
    symbol_col_letter = gspread.utils.rowcol_to_a1(1, symbol_col_index).rstrip("0123456789")

    rows_to_append = []
    for r in all_rows:
        row_num = start_row + len(rows_to_append)
        # ใช้ VLOOKUP อ้างอิงแท็บ StockList เพื่อให้ Sector มาจากแหล่งเดียวกันทั้งไฟล์
        sector_formula = (
            f'=IFERROR(VLOOKUP({symbol_col_letter}{row_num},'
            f'{STOCKLIST_SHEET_NAME}!A:B,2,FALSE),"")'
        )
        rows_to_append.append([
            date_str, time_str, r["Market"], r["Group"], sector_formula, r["Symbol"],
            r["Open"], r["High"], r["Low"], r["Last"], r["Chg"], r["ChgPct"],
            r["Bid"], r["Ask"], r["Volume"], r["Value"],
            trigger_label,
        ])

    if not rows_to_append:
        print("  ไม่มีแถวข้อมูลจะส่งเข้า StockDailyDetail")
        return 0

    append_safe(ws, rows_to_append, lambda vals: data_date in existing_dates(vals))
    print(f"  ส่ง {len(rows_to_append)} แถว เข้า worksheet '{STOCKDETAIL_SHEET_NAME}'")
    return len(rows_to_append)


LOG_SHEET_NAME = "Log"
LOG_HEADERS = ["Date", "Time", "Workflow", "Trigger", "Status", "RowsSent", "Detail"]


def push_to_log(sh, date_str: str, time_str: str, workflow_name: str, trigger_label: str,
                 status: str, rows_sent, detail: str = ""):
    """บันทึกประวัติการรันลงแท็บ Log (Detail ขึ้นต้นด้วยเวอร์ชัน)"""
    detail = f"[{SCRIPT_NAME} v{__version__}] {detail}".strip()
    try:
        ws = sh.worksheet(LOG_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=LOG_SHEET_NAME, rows=2000, cols=10)
        call_with_retry(ws.append_row, LOG_HEADERS, value_input_option="USER_ENTERED", table_range="A1")

    call_with_retry(
        ws.append_row,
        [date_str, time_str, workflow_name, trigger_label, status, rows_sent, detail],
        value_input_option="USER_ENTERED",
        table_range="A1",
    )


def capture_once():
    now = dt.datetime.now(BANGKOK_TZ)
    date_str = now.strftime("%d/%m/%Y")      # วันที่ที่ script รัน (ใช้กับ Log และคอลัมน์ Time)
    time_str = now.strftime("%H:%M")
    trigger_label = get_trigger_label()
    workflow_name = get_workflow_name()
    print(f"[{date_str} {time_str} เวลาไทย] เริ่ม {SCRIPT_NAME} v{__version__} "
          f"(workflow: {workflow_name}, trigger: {trigger_label})")

    sh = get_open_spreadsheet()
    known_symbols = load_stocklist_symbols(sh)

    def log(status, rows_sent, detail=""):
        push_to_log(sh, date_str, time_str, workflow_name, trigger_label, status, rows_sent, detail)

    _PAGE_DATES.clear()
    try:
        all_rows = fetch_all_stock_details(known_symbols)
    except Exception as e:
        log("Failed", 0, f"{type(e).__name__}: {str(e)[:200]}")
        raise

    if not all_rows:
        print("  ไม่พบข้อมูลหุ้นรายตัวเลย -> ข้ามการส่งเข้า Google Sheet รอบนี้")
        log("NoData", 0, "ไม่พบข้อมูลหุ้นรายตัว")
        return

    data_date, note = choose_data_date(_PAGE_DATES)
    if data_date is None:
        log("Failed", 0, "หาวันที่ 'ข้อมูลล่าสุด' บนหน้าเว็บ SET ไม่เจอเลย ไม่บันทึก เพื่อไม่ให้ใช้วันที่ผิด")
        raise ValueError("หาวันที่ 'ข้อมูลล่าสุด' บนหน้าเว็บ SET ไม่เจอเลย")
    data_date_str = data_date.strftime("%d/%m/%Y")
    print(f"  พบข้อมูลหุ้นรวม {len(all_rows)} แถว วันที่ของข้อมูล {data_date_str}{note}")

    try:
        ws, values = load_stockdetail(sh)
        if data_date in existing_dates(values):
            print(f"  มีข้อมูลวันที่ {data_date_str} อยู่แล้ว -> ไม่บันทึกซ้ำ")
            log("AlreadyCaptured", 0, f"มีข้อมูลวันที่ {data_date_str} อยู่แล้ว")
            return
        is_dup, prev_date = same_as_previous(values, all_rows)
        if is_dup:
            print(f"  ตัวเลขเหมือนข้อมูลวันที่ {prev_date:%d/%m/%Y} -> ไม่บันทึก")
            log("DuplicateOfPrevious", 0,
                f"ข้อมูลวันที่ {data_date_str} เหมือนวันที่ {prev_date:%d/%m/%Y} (ตัวเลขหุ้นซ้ำ)")
            return
        rows_sent = push_to_stockdetail(ws, values, all_rows, data_date_str, time_str, trigger_label, data_date)
    except Exception as e:
        log("Failed", 0, f"{type(e).__name__}: {str(e)[:200]}")
        raise

    log("Success", rows_sent, f"ข้อมูลวันที่ {data_date_str}{note}")


if __name__ == "__main__":
    try:
        capture_once()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่าง capture:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
