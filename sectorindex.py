"""
Capture Sector Index  v1.1.0  (sectorindex.py)
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
5. ส่งเข้า Google Sheet แท็บ SectorIndex พร้อมคอลัมน์ Group/Sector
6. บันทึกผลการรันลงแท็บ Log เดียวกับสคริปต์อื่น (Scheduled/Manual, Success/Failed)

วันที่ของข้อมูลใช้วันที่ "ข้อมูลล่าสุด" บนหน้าเว็บ (มีอยู่แล้วตั้งแต่เวอร์ชันแรก) v1.1.0 เพิ่มการตรวจซ้ำ
แยกตามตลาด (SET / mai) ก่อนบันทึก:
  1) ถ้าแท็บมีแถวของวันที่นั้น+ตลาดนั้นอยู่แล้ว -> ข้ามตลาดนั้น (Log: AlreadyCaptured)
  2) ถ้าตัวเลขเหมือนข้อมูลวันก่อนหน้าของตลาดนั้นเกือบทุกแถว -> ข้าม (Log: DuplicateOfPrevious)
     (เช่น วันหยุดตลาดที่หน้าเว็บยังโชว์ข้อมูลวันทำการล่าสุด)
ตลาดที่หาวันที่ไม่เจอจะไม่ถูกบันทึก (ระบุใน Log) แทนการบันทึกแถวที่ไม่มีวันที่

Environment variables ที่ต้องตั้ง (เหมือนสคริปต์อื่นในชุดนี้):
    GSHEET_ID
    GOOGLE_APPLICATION_CREDENTIALS

Changelog (เพิ่มเวอร์ชันทุกครั้งที่แก้ไฟล์นี้ และเวอร์ชันจะแสดงใน Log คอลัมน์ Detail):
  1.1.0  ตรวจซ้ำวันเดิม/ซ้ำวันก่อนแยกตามตลาด, ลองใหม่อัตโนมัติเมื่อเชื่อมต่อ Google Sheets หลุด,
         แสดงเวอร์ชันใน Log
  1.0.0  เวอร์ชันเดิม (แปลง - เป็น 0 ในคอลัมน์ตัวเลข, ใช้วันที่ข้อมูลจากเว็บ)
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

SCRIPT_NAME = "Capture Sector Index"
__version__ = "1.1.0"

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
LOG_HEADERS = ["Date", "Time", "Workflow", "Trigger", "Status", "RowsSent", "Detail"]

NAME_CODE_RE = re.compile(r"^(.*?)\s*\(([A-Z&]+)\)\s*$")

# จับข้อความ "ข้อมูลล่าสุด 22 ก.ค. 2569 14:36:41" ที่โชว์บนหน้าเว็บ แยกเป็น
# กลุ่มวันที่ กับ กลุ่มเวลา คนละส่วนกัน (เวลาข้อมูลจริงจากฝั่ง settrade.com
# ต่างจากเวลาที่ script รันเอง)
SOURCE_TIME_RE = re.compile(
    r"ข้อมูลล่าสุด\s+(\d{1,2})\s+(\S+?)\.?\s+(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})"
)

THAI_MONTH_ABBR = {
    "ม.ค": 1, "ก.พ": 2, "มี.ค": 3, "เม.ย": 4, "พ.ค": 5, "มิ.ย": 6,
    "ก.ค": 7, "ส.ค": 8, "ก.ย": 9, "ต.ค": 10, "พ.ย": 11, "ธ.ค": 12,
}


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


DATE_CELL_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d %b %Y", "%d-%m-%Y")
DUPLICATE_FRACTION = 0.98
MIN_ROWS_FOR_DUPLICATE_CHECK = 20
MARKETS = ("SET", "mai")


def parse_date_cell(s):
    s = str(s or "").strip().lstrip("'")
    for fmt in DATE_CELL_FORMATS:
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def existing_keys(values) -> set:
    """ชุด (วันที่, ตลาด) ที่มีอยู่ในแท็บแล้ว"""
    idx = {h: i for i, h in enumerate(SECTORINDEX_HEADERS)}
    out = set()
    for r in values[1:]:
        if len(r) > idx["Market"]:
            d = parse_date_cell(r[idx["Date"]])
            if d:
                out.add((d, str(r[idx["Market"]]).strip()))
    return out


def _norm_value(v):
    s = str(v if v is not None else "").replace(",", "").strip()
    if s == "":
        return ""
    try:
        return round(float(s), 4)
    except ValueError:
        return s.lower()


def _sig(group, sector, last, chg, vol, val):
    return ((str(group).strip(), str(sector).strip()), tuple(_norm_value(x) for x in (last, chg, vol, val)))


def previous_snapshot(values, market, before_date):
    """ข้อมูลของตลาดนี้ ณ วันล่าสุดที่เก่ากว่า before_date: (วันที่, {(group,sector): (last,chg,vol,val)})"""
    idx = {h: i for i, h in enumerate(SECTORINDEX_HEADERS)}
    best, rows = None, []
    for r in values[1:]:
        if len(r) <= idx["Value(MB)"] or str(r[idx["Market"]]).strip() != market:
            continue
        d = parse_date_cell(r[idx["Date"]])
        if d is None or d >= before_date:
            continue
        if best is None or d > best:
            best, rows = d, []
        if d == best:
            rows.append(r)
    snap = dict(_sig(r[idx["Group"]], r[idx["Sector"]], r[idx["Last"]], r[idx["Chg"]],
                     r[idx["Volume('000 Shares)"]], r[idx["Value(MB)"]]) for r in rows)
    return best, snap


def same_as_previous(values, market, data_date, market_rows):
    prev_date, prev = previous_snapshot(values, market, data_date)
    new = dict(_sig(r["Group"], r["Sector"], r["Last"], r["Chg"], r["Volume"], r["Value"]) for r in market_rows)
    if prev_date is None or len(new) < MIN_ROWS_FOR_DUPLICATE_CHECK or len(prev) < MIN_ROWS_FOR_DUPLICATE_CHECK:
        return False, prev_date
    same = sum(1 for k, v in new.items() if prev.get(k) == v)
    return same / len(new) >= DUPLICATE_FRACTION, prev_date


def _dash_to_zero(val):
    """เว็บ settrade.com แสดง '-' หรือปล่อยว่างไว้ เวลาหุ้น/ดัชนีไม่มีการ
    เปลี่ยนแปลง (เช่น ไม่มีการซื้อขายเลยในวันนั้น) แปลงเป็นเลข 0 ก่อนส่งเข้า
    Google Sheet เพื่อให้เป็นตัวเลขที่คำนวณ/กรองต่อได้ทันที"""
    s = str(val).strip()
    return "0" if s in ("-", "", "nan", "NaN") else val


def get_source_datetime(page):
    """
    ดึงวันที่/เวลาจากข้อความ 'ข้อมูลล่าสุด ...' บนหน้าเว็บ (เป็นวันที่แบบไทย
    เช่น '22 ก.ค. 2569') แล้วแปลงเป็น:
      - date: string รูปแบบ "DD/MM/YYYY" ปฏิทินสากล (แปลง พ.ศ. -> ค.ศ. โดย
        ลบ 543 ให้เอง)
      - time: string รูปแบบ "HH:MM" (ตัดวินาทีทิ้ง)
    คืนค่า ('', '') ถ้าหาข้อความหรือแปลงเดือนไทยไม่สำเร็จ

    ไม่ปล่อยให้ Google Sheets เดารูปแบบวันที่เอง เพราะเคยเจอปัญหาแสดงผลไม่
    สม่ำเสมอ (บางแถวโชว์ "24/07/2026" บางแถวโชว์ "24 ก.ค. 2569" ปนกัน) ต้อง
    ควบคุมรูปแบบให้แน่นอนจากฝั่งโค้ดเอง แล้วบังคับเก็บเป็นข้อความล้วนตอนส่งเข้า
    Sheet (ดู push_to_sectorindex)
    """
    body_text = page.evaluate("() => document.body.innerText || ''")
    m = SOURCE_TIME_RE.search(body_text)
    if not m:
        return "", ""

    day, month_abbr, buddhist_year, hh, mm, _ss = m.groups()
    month_num = THAI_MONTH_ABBR.get(month_abbr.strip())
    if month_num is None:
        return "", ""

    gregorian_year = int(buddhist_year) - 543
    date_str = f"{int(day):02d}/{month_num:02d}/{gregorian_year}"
    time_str = f"{int(hh):02d}:{mm}"
    return date_str, time_str


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
    """อ่านชื่อ workflow จาก GITHUB_WORKFLOW (ตรงกับ name: ในไฟล์ .yml เอง)"""
    return os.environ.get("GITHUB_WORKFLOW", "local")


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
            "Last": _dash_to_zero(row.get("ล่าสุด", "")),
            "Chg": _dash_to_zero(row.get("เปลี่ยนแปลง", "")),
            "Chg%": _dash_to_zero(row.get("เปลี่ยนแปลง (%)", "")),
            "Volume": _dash_to_zero(row.get("ปริมาณ ('000 หุ้น)", "")),
            "Value": _dash_to_zero(row.get("มูลค่า (ล้านบาท)", "")),
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
    return _Retrying(call_with_retry(gc.open_by_key, sheet_id))


def push_to_sectorindex(sh, all_rows, trigger_label: str):
    """บันทึกข้อมูลแยกตามตลาด พร้อมตรวจซ้ำ คืนค่า (จำนวนแถวที่ส่ง, รายการที่ข้ามพร้อมเหตุผล)"""
    try:
        ws = sh.worksheet(SECTORINDEX_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SECTORINDEX_SHEET_NAME, rows=3000, cols=15)
        call_with_retry(ws.append_row, SECTORINDEX_HEADERS, value_input_option="USER_ENTERED", table_range="A1")
    values = ws.get_all_values()
    have = existing_keys(values)

    sent, skipped = 0, []
    for market in MARKETS:
        market_rows = [r for r in all_rows if r["Market"] == market]
        if not market_rows:
            continue
        data_date = parse_date_cell(market_rows[0].get("Date", ""))
        if data_date is None:
            skipped.append((market, "NoDate", f"{market}: หาวันที่ข้อมูลบนหน้าเว็บไม่เจอ ไม่บันทึก"))
            continue
        label = data_date.strftime("%d/%m/%Y")
        if (data_date, market) in have:
            skipped.append((market, "AlreadyCaptured", f"มีข้อมูล {market} วันที่ {label} อยู่แล้ว"))
            continue
        is_dup, prev_date = same_as_previous(values, market, data_date, market_rows)
        if is_dup:
            skipped.append((market, "DuplicateOfPrevious",
                            f"{market} วันที่ {label} เหมือนวันที่ {prev_date:%d/%m/%Y} (หน้าเว็บอาจยังไม่อัปเดต/วันหยุด)"))
            continue

        rows_to_append = []
        for r in market_rows:
            # ขึ้นต้นด้วย ' เพื่อบังคับให้ Google Sheets เก็บเป็นข้อความล้วน ไม่ตีความเป็นวันที่เอง
            date_val = f"'{r['Date']}" if r.get("Date") else ""
            time_val = f"'{r['Time']}" if r.get("Time") else ""
            rows_to_append.append([
                date_val, time_val, r["Market"], r["Group"], r["Sector"], r["Name"],
                r["Last"], r["Chg"], r["Chg%"], r["Volume"], r["Value"], trigger_label,
            ])
        append_safe(ws, rows_to_append, lambda vals, d=data_date, m=market: (d, m) in existing_keys(vals))
        sent += len(rows_to_append)
        print(f"  ส่ง {len(rows_to_append)} แถว ({market} วันที่ {label}) เข้า worksheet '{SECTORINDEX_SHEET_NAME}'")
    return sent, skipped


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
    date_str = now.strftime("%d/%m/%Y")
    time_str = now.strftime("%H:%M")
    trigger_label = get_trigger_label()
    workflow_name = get_workflow_name()
    print(f"[{date_str} {time_str} เวลาไทย] เริ่ม {SCRIPT_NAME} v{__version__} "
          f"(workflow: {workflow_name}, trigger: {trigger_label})")

    sh = get_open_spreadsheet()

    def log(status, rows_sent, detail=""):
        push_to_log(sh, date_str, time_str, workflow_name, trigger_label, status, rows_sent, detail)

    try:
        all_rows = fetch_sector_index_data()
    except Exception as e:
        log("Failed", 0, f"{type(e).__name__}: {str(e)[:200]}")
        raise

    if not all_rows:
        print("  ไม่พบข้อมูลดัชนีกลุ่มอุตสาหกรรมเลย -> ข้ามการส่งเข้า Google Sheet รอบนี้")
        log("NoData", 0, "ไม่พบข้อมูลดัชนีกลุ่มอุตสาหกรรม")
        return

    try:
        rows_sent, skipped = push_to_sectorindex(sh, all_rows, trigger_label)
    except Exception as e:
        log("Failed", 0, f"{type(e).__name__}: {str(e)[:200]}")
        raise

    detail = "; ".join(msg for _, _, msg in skipped)
    for _, _, msg in skipped:
        print(f"  ข้าม: {msg}")
    if rows_sent > 0:
        log("Success", rows_sent, detail)
    elif skipped and all(reason == "NoDate" for _, reason, _ in skipped):
        log("Failed", 0, detail)
        raise ValueError(detail)
    elif any(reason == "DuplicateOfPrevious" for _, reason, _ in skipped):
        log("DuplicateOfPrevious", 0, detail)
    else:
        log("AlreadyCaptured", 0, detail)


if __name__ == "__main__":
    try:
        capture_once()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่าง capture:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
