"""
capture_setoverview.py
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

Environment variables ที่ต้องตั้ง (เหมือนกับ capture_settrade.py):
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

    return {"mai_value": mai_value}


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
    return gc.open_by_key(sheet_id)


def push_row_to_setdatabase(sh, row):
    try:
        ws = sh.worksheet(SETDATABASE_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SETDATABASE_SHEET_NAME, rows=1000, cols=30)
        ws.append_row(SETDATABASE_HEADERS, value_input_option="USER_ENTERED")

    ws.append_row(row, value_input_option="USER_ENTERED")
    print(f"  ส่ง 1 แถว เข้า worksheet '{SETDATABASE_SHEET_NAME}'")


def capture_once():
    now = dt.datetime.now(BANGKOK_TZ)
    date_str = now.strftime("%d/%m/%Y")
    print(f"[{date_str} {now.strftime('%H:%M')} เวลาไทย] เริ่ม capture ข้อมูลจาก set.or.th")

    data = fetch_all_setoverview_data()
    row = build_setdatabase_row(data, date_str)

    print(f"  แถวที่จะส่ง: {row}")

    sh = get_open_spreadsheet()
    push_row_to_setdatabase(sh, row)


if __name__ == "__main__":
    try:
        capture_once()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่าง capture:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
