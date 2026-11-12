"""
Capture SET Overview  v1.1.0  (capture_setoverview.py)
-----------------------
Capture สรุปภาพรวมตลาดจากหน้า set.or.th (SET, SET50, mai, Institution/
Proprietary/Foreign/Individual, TFEX) แล้ว append เข้าแท็บ SetDatabase ใน
Google Sheet "Settrade Capture Log" (คนละแท็บกับ Master ของ Top Ranking)

capture วันละครั้ง (ไม่ใช่ 4 รอบ/วันเหมือน Top Ranking) เพราะเป็นข้อมูลสรุป
รวมทั้งวัน รันหลังตลาดปิดสนิทแล้ว (18:30 น. เวลาไทย)

หมายเหตุ: คอลัมน์ TFEX Prem/Disc งดไว้ก่อน (สูตรคำนวณซับซ้อน ต้องใช้ราคา
Futures เทียบ SET50 spot) จะเว้นว่างไว้ในแถวที่ capture ทุกครั้ง

แหล่งข้อมูล:
  - https://www.set.or.th/en/home
      -> ตาราง Index summary (SET, SET50, ... ) : Last, Change, Volume, Value
      -> ตาราง Institution/Proprietary/Foreign/Individual : Buy, Sell, Net
      -> ข้อความสรุปด้านบน (SET %Change, mai %Change, TFEX Volume/OI)
  - https://www.set.or.th/en/market/index/mai/overview
      -> ข้อความ "Value (M.Baht)" ของ mai (ไม่มีในหน้า home)

วันที่ของข้อมูล (v1.1.0): ใช้วันที่ "As of dd Mon yyyy" ที่หน้าเว็บ SET ระบุไว้ในส่วน Trading
Summary เป็นคอลัมน์ Date (ไม่ใช่วันที่ script รัน) แล้วตรวจก่อนบันทึก 2 ขั้น:
  1) ถ้าแท็บ SetDatabase มีแถววันที่นั้นอยู่แล้ว -> ไม่บันทึกซ้ำ (Log: AlreadyCaptured)
  2) ถ้าตัวเลขเหมือนแถวล่าสุดของวันก่อนทุกช่อง -> ไม่บันทึก (Log: DuplicateOfPrevious)
ถ้าหาวันที่บนหน้าเว็บไม่เจอ จะไม่บันทึก (Log: Failed) แทนการใช้วันที่ผิด
วันหยุดตลาดจึงถูกข้ามเอง เพราะหน้าเว็บยังโชว์วันทำการล่าสุดเดิม

Environment variables ที่ต้องตั้ง (เหมือนกับ capture_settrade.py):
    GSHEET_ID
    GOOGLE_APPLICATION_CREDENTIALS

Changelog (เพิ่มเวอร์ชันทุกครั้งที่แก้ไฟล์นี้ และเวอร์ชันจะแสดงใน Log คอลัมน์ Detail):
  1.1.0  ใช้วันที่ As of จากหน้าเว็บ SET เป็น Date, ตรวจซ้ำวันเดิมและซ้ำกับแถวล่าสุด,
         ลองใหม่อัตโนมัติเมื่อเชื่อมต่อ Google Sheets หลุด, แสดงเวอร์ชันใน Log
  1.0.0  เวอร์ชันเดิม ใช้วันที่ที่ script รันเป็น Date
"""

import os
import io
import re
import sys
import time
import traceback
import datetime as dt
from zoneinfo import ZoneInfo

import pandas as pd
import gspread
import requests
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

SCRIPT_NAME = "Capture SET Overview"
__version__ = "1.1.0"

BANGKOK_TZ = ZoneInfo("Asia/Bangkok")

SET_HOME_URL = "https://www.set.or.th/en/home"
MAI_OVERVIEW_URL = "https://www.set.or.th/en/market/index/mai/overview"

SETDATABASE_SHEET_NAME = "SetDatabase"
SETDATABASE_HEADERS = [
    "Date", "Set Index", "Chg", "Chg%", "Value(MB)",
    "Inst.Buy", "Inst.Sell", "Inst.Net",
    "Prop.Buy", "Prop.Sell", "Prop.Net",
    "Foreign.Buy", "Foreign.Sell", "Foreign.Net",
    "Indiv.Buy", "Indiv.Sell", "Indiv.Net",
    "SET50 Idx", "SET50 Chg%", "SET50 Val(MB)",
    "mai Idx", "mai Chg%", "mai Val(MB)",
    "TFEX Vol", "TFEX OI", "TFEX Prem/Disc",
]

# แท็บเก็บประวัติการรันทุกครั้ง (ใช้ร่วมกับ capture_settrade.py และ
# capture_sectorindex.py) คอลัมน์ Workflow บอกว่าแถวนี้มาจาก workflow ไหน
LOG_SHEET_NAME = "Log"
LOG_HEADERS = ["Date", "Time", "Workflow", "Trigger", "Status", "RowsSent", "Detail"]


EN_MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}
# "As of dd Mon yyyy" ที่อยู่ถัดจากหัวข้อ Trading Summary (ข้อมูลที่ใช้บันทึกจริง)
# ต้องเริ่มค้นหลังคำว่า Trading Summary เพราะก่อนหน้านั้นมี As of ของดัชนี SETTRI (วันก่อน)
ASOF_AFTER_SUMMARY_RE = re.compile(
    r"Trading Summary.*?As of\s+(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})", re.S)
# แบบแม่นที่สุด: วันที่อยู่ติดแท็บ "Accumulated Trading Summary" ของกล่องเดียวกันเลย
ASOF_ACCUM_RE = re.compile(
    r"Accumulated Trading Summary\s+As of\s+(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})")
LAST_UPDATE_RE = re.compile(r"Last Update\s+(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})")
DATE_CELL_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d %b %Y", "%d-%m-%Y")


def _make_date(day, mon, year):
    month = EN_MONTHS.get(str(mon)[:3].lower())
    if not month:
        return None
    try:
        return dt.date(int(year), month, int(day))
    except ValueError:
        return None


def extract_as_of(body_text: str):
    """วันที่ของข้อมูลบนหน้า SET: ใช้ As of ใต้ Trading Summary ก่อน ถ้าไม่เจอใช้ Last Update"""
    for pattern in (ASOF_ACCUM_RE, ASOF_AFTER_SUMMARY_RE):
        m = pattern.search(body_text or "")
        if m:
            d = _make_date(*m.groups())
            if d:
                return d
    m = LAST_UPDATE_RE.search(body_text or "")
    return _make_date(*m.groups()) if m else None


def extract_last_update(text: str):
    m = LAST_UPDATE_RE.search(text or "")
    return _make_date(*m.groups()) if m else None


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


def last_data_row(values):
    for r in reversed(values):
        if r and parse_date_cell(r[0]):
            return r
    return None


def _norm_value(v):
    s = str(v if v is not None else "").replace(",", "").strip()
    if s == "":
        return ""
    try:
        return round(float(s), 4)
    except ValueError:
        return s.lower()


def same_numbers(prev_row, new_row) -> bool:
    """ตัวเลขทุกช่อง (ยกเว้น Date) เหมือนกัน และมีข้อมูลจริงอย่างน้อย 5 ช่อง"""
    n = max(len(prev_row), len(new_row))
    a = [_norm_value(prev_row[i]) if i < len(prev_row) else "" for i in range(1, n)]
    b = [_norm_value(new_row[i]) if i < len(new_row) else "" for i in range(1, n)]
    return a == b and sum(1 for x in b if x != "") >= 5


RETRY_ATTEMPTS = 5
RETRY_BASE_DELAY = 2


def _is_transient(e) -> bool:
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
    """เพิ่มแถวแบบกันซ้ำ ถ้าเชื่อมต่อหลุด จะอ่านชีทดูก่อนว่ารอบที่แล้วเขียนเข้าไปแล้วหรือยัง"""
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


def _num(x):
    """แปลงข้อความตัวเลข (มี comma/เว้นวรรค) เป็น float แบบปลอดภัย คืน '' ถ้าแปลงไม่ได้"""
    try:
        s = str(x).replace(",", "").strip()
        if s in ("", "-", "nan", "NaN"):
            return ""
        return float(s)
    except (ValueError, TypeError):
        return ""


def _pct_change(last, chg):
    """คำนวณ %เปลี่ยนแปลง จาก Last และ Change (กรณีหน้าเว็บไม่ได้โชว์ % ตรง ๆ)"""
    last = _num(last)
    chg = _num(chg)
    if last == "" or chg == "":
        return ""
    prior = last - chg
    if prior == 0:
        return ""
    return round(chg / prior * 100, 2)


def get_visible_tables(page):
    """คืนค่า list ของ DataFrame จากตารางที่มองเห็นอยู่จริงบนหน้าปัจจุบัน"""
    htmls = page.evaluate(
        """
        () => Array.from(document.querySelectorAll('table'))
            .filter(t => t.offsetParent !== null)
            .map(t => t.outerHTML)
        """
    )
    tables = []
    for html in htmls:
        try:
            parsed = pd.read_html(io.StringIO(html))
        except Exception:
            continue
        tables.extend(parsed)
    return tables


def fetch_set_home_data(page):
    """เปิดหน้า SET Home แล้วดึงข้อมูล SET, SET50, Institution/Prop/Foreign/Indiv, TFEX"""
    page.goto(SET_HOME_URL, wait_until="domcontentloaded", timeout=60000)
    try:
        page.wait_for_selector("table", timeout=30000)
    except PlaywrightTimeoutError:
        print("  คำเตือน: รอตาราง (<table>) ไม่เจอภายในเวลาที่กำหนด (SET Home)")
    page.wait_for_timeout(4000)

    tables = get_visible_tables(page)
    body_text = page.evaluate("() => document.body.innerText || ''")

    data = {}

    # ----- ตาราง Index summary: หา SET กับ SET50 -----
    index_table = None
    for t in tables:
        cols = [str(c) for c in t.columns]
        if "Index" in cols and "Last" in cols and "Change" in cols:
            index_table = t
            break

    if index_table is not None:
        for _, row in index_table.iterrows():
            idx_name = str(row.get("Index", "")).strip()
            if idx_name == "SET":
                data["set_last"] = row.get("Last", "")
                data["set_chg"] = row.get("Change", "")
                data["set_value"] = row.get("Value (M.Baht)", "")
            elif idx_name == "SET50":
                data["set50_last"] = row.get("Last", "")
                data["set50_chg"] = row.get("Change", "")
                data["set50_value"] = row.get("Value (M.Baht)", "")
    else:
        print("  คำเตือน: หาตาราง Index summary (SET/SET50) ไม่เจอ")

    # ----- ตาราง Institution/Proprietary/Foreign/Individual -----
    trading_table = None
    for t in tables:
        cols = [str(c) for c in t.columns]
        if "Type" in cols and "Net" in cols:
            trading_table = t
            break

    trading = {}
    if trading_table is not None:
        for _, row in trading_table.iterrows():
            t_type = str(row.get("Type", "")).strip()
            buy_col = next((c for c in trading_table.columns if "Buy" in str(c)), None)
            sell_col = next((c for c in trading_table.columns if "Sell" in str(c)), None)
            trading[t_type] = {
                "Buy": row.get(buy_col, "") if buy_col else "",
                "Sell": row.get(sell_col, "") if sell_col else "",
                "Net": row.get("Net", ""),
            }
    else:
        print("  คำเตือน: หาตาราง Institution/Proprietary/Foreign/Individual ไม่เจอ")
    data["trading"] = trading

    # ----- %Change ของ SET และ mai จากข้อความบนหน้า (ตัวเลขแรก 2 ค่าที่เจอ) -----
    pct_matches = re.findall(r"\(([+-]?[\d.]+)%\)", body_text)
    data["set_chg_pct_text"] = pct_matches[0] if len(pct_matches) > 0 else ""
    data["mai_chg_pct_text"] = pct_matches[1] if len(pct_matches) > 1 else ""

    # ----- mai Last (เผื่อใช้คำนวณสำรอง) -----
    mai_last_match = re.search(r"mai\s*\n?\s*([\d,]+\.\d+)", body_text)
    data["mai_last"] = mai_last_match.group(1) if mai_last_match else ""

    # ----- TFEX Volume / OI -----
    tfex_match = re.search(r"TFEX\s*\nVolume\s+([\d,.\-]+)\s*\nOI\s+([\d,.\-]+)", body_text)
    if tfex_match:
        data["tfex_vol"] = tfex_match.group(1)
        data["tfex_oi"] = tfex_match.group(2)
    else:
        print("  คำเตือน: หาตัวเลข TFEX Volume/OI ในข้อความหน้าเว็บไม่เจอ")
        data["tfex_vol"] = ""
        data["tfex_oi"] = ""

    # ----- วันที่ของข้อมูลตามที่หน้าเว็บ SET ระบุ (As of ...) -----
    data["as_of"] = extract_as_of(body_text)
    if data["as_of"] is None:
        print("  คำเตือน: หาวันที่ 'As of' บนหน้า SET Home ไม่เจอ")

    return data


def fetch_mai_overview_data(page):
    """
    เปิดหน้า mai Overview แล้วดึง Value (M.Baht) ของ mai

    หน้านี้มีคำว่า "Value (M.Baht)" โผล่ 2 จุด: จุดแรกเป็นของ SET (ค้างมาจาก
    กล่องสรุปด้านบนสุดของหน้า ก่อนถึงส่วน mai Index) จุดที่สองถึงจะเป็นของ mai
    จริง ๆ (อยู่หลังคำว่า "mai Index") จึงต้องตัดข้อความให้เหลือเฉพาะส่วนหลัง
    "mai Index" ก่อนค้นหา ไม่เช่นนั้นจะได้ค่าของ SET ผิดตัวมาแทน
    """
    page.goto(MAI_OVERVIEW_URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(4000)

    body_text = page.evaluate("() => document.body.innerText || ''")

    mai_index_pos = body_text.find("mai Index")
    if mai_index_pos == -1:
        print("  คำเตือน: หาจุดเริ่มต้นส่วน 'mai Index' ในหน้า mai Overview ไม่เจอ")
        mai_section = body_text
    else:
        mai_section = body_text[mai_index_pos:]

    value_match = re.search(r"Value \(M\.Baht\)\s+([\d,.\-]+)", mai_section)
    mai_value = value_match.group(1) if value_match else ""
    if not mai_value:
        print("  คำเตือน: หา mai Value (M.Baht) ในหน้า mai Overview ไม่เจอ")

    return {"mai_value": mai_value, "mai_date": extract_last_update(mai_section)}


def fetch_all_setoverview_data():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 2400})
        try:
            print("  กำลังดึงข้อมูลจาก SET Home ...")
            home_data = fetch_set_home_data(page)
            print("  กำลังดึงข้อมูลจาก mai Overview ...")
            mai_data = fetch_mai_overview_data(page)
        finally:
            browser.close()
    home_data.update(mai_data)
    return home_data


def build_setdatabase_row(data, date_str):
    trading = data.get("trading", {})

    def t(type_name, field):
        return trading.get(type_name, {}).get(field, "")

    set_chg_pct = data.get("set_chg_pct_text") or _pct_change(data.get("set_last"), data.get("set_chg"))
    set50_chg_pct = _pct_change(data.get("set50_last"), data.get("set50_chg"))
    mai_chg_pct = data.get("mai_chg_pct_text")

    row = [
        date_str,
        data.get("set_last", ""),
        data.get("set_chg", ""),
        set_chg_pct,
        data.get("set_value", ""),
        t("Institution", "Buy"), t("Institution", "Sell"), t("Institution", "Net"),
        t("Proprietary", "Buy"), t("Proprietary", "Sell"), t("Proprietary", "Net"),
        t("Foreign", "Buy"), t("Foreign", "Sell"), t("Foreign", "Net"),
        t("Individual", "Buy"), t("Individual", "Sell"), t("Individual", "Net"),
        data.get("set50_last", ""),
        set50_chg_pct,
        data.get("set50_value", ""),
        data.get("mai_last", ""),
        mai_chg_pct,
        data.get("mai_value", ""),
        data.get("tfex_vol", ""),
        data.get("tfex_oi", ""),
        "",  # TFEX Prem/Disc -> งดไว้ก่อนตามที่ตกลงกัน
    ]
    return row


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
    return call_with_retry(gc.open_by_key, sheet_id)


def load_setdatabase(sh):
    """คืนค่า (worksheet, ค่าทั้งหมดในแท็บ) สร้างแท็บ+หัวตารางให้ถ้ายังไม่มี"""
    try:
        ws = call_with_retry(sh.worksheet, SETDATABASE_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = call_with_retry(sh.add_worksheet, title=SETDATABASE_SHEET_NAME, rows=1000, cols=30)
        call_with_retry(ws.append_row, SETDATABASE_HEADERS, value_input_option="USER_ENTERED",
                        table_range="A1")
    return ws, call_with_retry(ws.get_all_values)


def get_trigger_label() -> str:
    """อ่านประเภทของ trigger จาก GITHUB_EVENT_NAME (Scheduled/Manual/local)"""
    event_name = os.environ.get("GITHUB_EVENT_NAME", "")
    if event_name == "schedule":
        return "Scheduled"
    elif event_name == "workflow_dispatch":
        return "Manual"
    elif event_name:
        return event_name
    return "local"


def get_workflow_name() -> str:
    """อ่านชื่อ workflow จาก GITHUB_WORKFLOW (ตรงกับ name: ในไฟล์ .yml เอง)"""
    return os.environ.get("GITHUB_WORKFLOW", "local")


def push_to_log(sh, date_str: str, time_str: str, workflow_name: str, trigger_label: str,
                 status: str, rows_sent, detail: str = ""):
    """บันทึกประวัติการรันลงแท็บ Log เดียวกับสคริปต์อื่นในชุดนี้ (Detail ขึ้นต้นด้วยเวอร์ชัน)"""
    detail = f"[{SCRIPT_NAME} v{__version__}] {detail}".strip()
    try:
        ws = call_with_retry(sh.worksheet, LOG_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = call_with_retry(sh.add_worksheet, title=LOG_SHEET_NAME, rows=2000, cols=10)
        call_with_retry(ws.append_row, LOG_HEADERS, value_input_option="USER_ENTERED")

    call_with_retry(
        ws.append_row,
        [date_str, time_str, workflow_name, trigger_label, status, rows_sent, detail],
        value_input_option="USER_ENTERED",
    )


def capture_once():
    now = dt.datetime.now(BANGKOK_TZ)
    date_str = now.strftime("%d/%m/%Y")      # วันที่ที่ script รัน (ใช้กับ Log เท่านั้น)
    time_str = now.strftime("%H:%M")
    trigger_label = get_trigger_label()
    workflow_name = get_workflow_name()
    print(f"[{date_str} {time_str} เวลาไทย] เริ่ม {SCRIPT_NAME} v{__version__} "
          f"(workflow: {workflow_name}, trigger: {trigger_label})")

    # เปิด Google Sheet ก่อน เพื่อให้บันทึก Log ได้แม้ขั้นตอนถัดไปจะล้มเหลว
    sh = get_open_spreadsheet()

    def log(status, rows_sent, detail=""):
        push_to_log(sh, date_str, time_str, workflow_name, trigger_label, status, rows_sent, detail)

    try:
        data = fetch_all_setoverview_data()
        data_date = data.get("as_of")
        if data_date is None:
            raise ValueError("หาวันที่ 'As of dd Mon yyyy' ของข้อมูลบนหน้า SET ไม่เจอ "
                             "ไม่บันทึก เพื่อไม่ให้ใช้วันที่ผิด")
        data_date_str = data_date.strftime("%d/%m/%Y")
        row = build_setdatabase_row(data, data_date_str)
    except Exception as e:
        log("Failed", 0, f"{type(e).__name__}: {str(e)[:300]}")
        raise

    note = ""
    mai_date = data.get("mai_date")
    if mai_date and mai_date != data_date:
        note = f" | เตือน: หน้า mai Overview เป็นวันที่ {mai_date:%d/%m/%Y} ไม่ตรงกับ SET"
    print(f"  ข้อมูลวันที่ {data_date_str} (As of บนหน้า SET){note}")
    print(f"  แถวที่จะส่ง: {row}")

    try:
        ws, values = load_setdatabase(sh)
        if data_date in existing_dates(values):
            print(f"  มีข้อมูลวันที่ {data_date_str} อยู่แล้ว -> ไม่บันทึกซ้ำ")
            log("AlreadyCaptured", 0, f"มีข้อมูลวันที่ {data_date_str} อยู่แล้ว")
            return
        prev = last_data_row(values)
        if prev is not None and same_numbers(prev, row):
            print(f"  ตัวเลขเหมือนแถวล่าสุด ({prev[0]}) ทุกช่อง -> ไม่บันทึก")
            log("DuplicateOfPrevious", 0,
                f"ข้อมูลวันที่ {data_date_str} เหมือนแถววันที่ {prev[0]} ทุกช่อง{note}")
            return
        append_safe(ws, [row], lambda vals: data_date in existing_dates(vals))
    except Exception as e:
        log("Failed", 0, f"{type(e).__name__}: {str(e)[:300]}")
        raise

    print(f"  ส่ง 1 แถว เข้า worksheet '{SETDATABASE_SHEET_NAME}'")
    log("Success", 1, f"ข้อมูลวันที่ {data_date_str}{note}")


if __name__ == "__main__":
    try:
        capture_once()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่าง capture:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
