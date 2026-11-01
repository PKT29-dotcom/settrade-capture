"""
capture_nvdr.py  (เวอร์ชันอ่านตารางจากหน้าเว็บ ไม่ใช้ Export Excel)
------------------------------------------------------------------
Capture NVDR Trading by Stock (รายหุ้น) จาก set.or.th แล้ว append เข้าแท็บ
NVDRTradingByStock ใน Google Sheet "Settrade Capture Log"

วิธีทำงาน (แนวเดียวกับ capture_settrade.py):
1. เปิดหน้า NVDR Trading by Stock ด้วย Playwright รอตารางโหลด
2. หน้าเว็บแสดง Volume กับ Value แยกกันด้วยปุ่มสลับ "Categorized by" จึงอ่าน
   2 รอบ (Volume แล้วกด Value) แล้วรวมเป็นแถวเดียวต่อหุ้น
3. ตารางยาวและอาจโหลดแถวทีละช่วงตอนเลื่อน จึงเลื่อนหน้า/ตารางไปเรื่อย ๆ
   แล้วเก็บแถวสะสมตาม Symbol จนจำนวนหุ้นไม่เพิ่มอีก
4. ใช้วันที่จากข้อความ "As of dd Mon yyyy" บนหน้า เป็นคอลัมน์ Date ของทุกแถว
   (วันที่ของข้อมูลจริง ไม่ใช่วันที่ที่ script รัน)
5. ถ้าแท็บมีข้อมูลของวันที่นั้นอยู่แล้วจะไม่ append ซ้ำ (Log: AlreadyCaptured)
6. ตรวจความน่าเชื่อถือก่อนบันทึก: ลำดับหัวคอลัมน์ต้องเป็น Buy,Sell,Total,Net,
   จำนวนหุ้นต้องไม่น้อยผิดปกติ และจำนวนหุ้นของ Volume กับ Value ต้องใกล้เคียงกัน
   ถ้าไม่ผ่านจะไม่บันทึก และ Log เป็น Failed พร้อมเหตุผล
7. หน่วยของตัวเลขอ่านจากหัวตาราง (เช่น '000 บาท) และแปลงเป็น หุ้น/บาท ให้อัตโนมัติ

Environment variables:
    GSHEET_ID
    GOOGLE_APPLICATION_CREDENTIALS
"""

import os
import re
import sys
import numbers
import traceback
import datetime as dt
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

BANGKOK_TZ = ZoneInfo("Asia/Bangkok")

NVDR_URL = "https://www.set.or.th/th/market/statistics/nvdr/trading-by-stock"

NVDR_SHEET_NAME = "NVDRTradingByStock"
# รูปแบบแท็บให้เหมือนไฟล์ Export จากเว็บ SET (หัว 2 แถว) แต่มี Date และ Symbol นำหน้า
NVDR_GROUP_ROW = ["", "", "Volume (Share)", "", "", "", "", "Value (Baht)", "", "", "", ""]
NVDR_COL_ROW = ["Date", "Symbol", "Buy", "Sell", "Total", "Net", "%*",
                "Buy", "Sell", "Total", "Net", "%*"]

# ช่องที่เว็บแสดง "-" (ไม่มีรายการ): False = เว้นว่างเหมือนไฟล์ Export, True = ใส่ 0
FILL_ZERO_FOR_BLANK = False

LOG_SHEET_NAME = "Log"
LOG_HEADERS = ["Date", "Time", "Workflow", "Trigger", "Status", "RowsSent", "Detail"]

MIN_EXPECTED_ROWS = 30      # น้อยกว่านี้ถือว่าเก็บได้ไม่ครบ ไม่บันทึก
MAX_SCROLL_STEPS = 600
STAGNANT_LIMIT = 6          # เลื่อนต่อเนื่องแล้วจำนวนหุ้นไม่เพิ่มกี่รอบถึงหยุด

ASOF_RE = re.compile(r"As of\s+(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})")
EN_MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}
SYMBOL_RE = re.compile(r"^[A-Z0-9][A-Z0-9&.\-]{0,19}$")
NOT_SYMBOLS = {"TOTAL", "SUM", "REMARK", "NOTE"}

# ---------- JavaScript ที่ใช้บนหน้าเว็บ ----------
_JS_FIND = """
const findTable = () => {
  const ts = Array.from(document.querySelectorAll('table')).filter(t => t.offsetParent !== null);
  ts.sort((a, b) => b.querySelectorAll('tr').length - a.querySelectorAll('tr').length);
  return ts[0] || null;
};
const findScroller = (t) => {
  let el = t.parentElement;
  while (el && el !== document.body && el !== document.documentElement) {
    const cs = getComputedStyle(el);
    if (/(auto|scroll)/.test(cs.overflowY) && el.scrollHeight > el.clientHeight + 5) return el;
    el = el.parentElement;
  }
  return null;
};
"""

JS_EXTRACT = "() => {" + _JS_FIND + """
  const t = findTable();
  if (!t) return null;
  const rows = Array.from(t.querySelectorAll('tr')).map(tr =>
    Array.from(tr.querySelectorAll('th,td')).map(c => (c.innerText || '').trim()));
  const head = t.tHead ? (t.tHead.innerText || '')
                       : rows.slice(0, 2).map(r => r.join(' ')).join(' ');
  return {rows: rows, head: head};
}"""

JS_HEAD = "() => {" + _JS_FIND + """
  const t = findTable();
  if (!t) return '';
  return t.tHead ? (t.tHead.innerText || '') : '';
}"""

JS_SCROLL_TOP = "() => {" + _JS_FIND + """
  const t = findTable();
  if (t) { const s = findScroller(t); if (s) s.scrollTop = 0; }
  window.scrollTo(0, 0);
}"""

JS_SCROLL_STEP = "() => {" + _JS_FIND + """
  const t = findTable();
  if (!t) return {moved: false, inner: false};
  const s = findScroller(t);
  let moved = false;
  if (s) {
    const b = s.scrollTop;
    s.scrollBy(0, Math.max(200, s.clientHeight * 0.8));
    moved = s.scrollTop !== b;
  }
  if (!moved) {
    const b = window.scrollY;
    window.scrollBy(0, Math.max(300, window.innerHeight * 0.8));
    moved = window.scrollY !== b;
  }
  return {moved: moved, inner: !!s};
}"""


# ---------- ฟังก์ชันประมวลผลข้อมูล (ไม่พึ่งเบราว์เซอร์ ทดสอบแยกได้) ----------
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


_NUMERIC_CELL_RE = re.compile(r"^-?[\d,]*\.?\d+$")


def _is_numeric_cell(c) -> bool:
    s = (c or "").strip()
    return s in ("", "-") or bool(_NUMERIC_CELL_RE.match(s))


def _num_or_none(c):
    """'-' หรือช่องว่าง -> None, ตัวเลข -> int/float"""
    s = str(c if c is not None else "").replace(",", "").strip()
    if s in ("", "-"):
        return None
    try:
        f = float(s)
    except ValueError:
        return None
    return int(f) if f.is_integer() else f


def _identity_ok(w) -> bool:
    """ข้อมูลจริงต้องเป็น Total = Buy + Sell และ Net = Buy - Sell (None นับเป็น 0)"""
    buy, sell, total, net = [(x or 0) for x in w[:4]]
    tol = max(1, 0.001 * abs(total))
    return abs(total - (buy + sell)) <= tol and abs(net - (buy - sell)) <= tol


def parse_as_of(body_text: str):
    m = ASOF_RE.search(body_text or "")
    if not m:
        return None
    month = EN_MONTHS.get(m.group(2)[:3].lower())
    if not month:
        return None
    return dt.date(int(m.group(3)), month, int(m.group(1)))


def unit_multiplier(head_text: str) -> int:
    """อ่านหน่วยจากหัวตาราง เช่น ('000 บาท) -> 1000, (ล้านบาท) -> 1,000,000"""
    h = head_text or ""
    if re.search(r"million|ล้าน", h, re.I):
        return 1_000_000
    if re.search(r"['’`]000|thousand|พัน", h, re.I):
        return 1000
    return 1


def check_head(head: str, view: str):
    """หัวตารางต้องเป็นมุมมองที่ต้องการ และคอลัมน์เรียง Buy,Sell,Total,Net"""
    if not re.search(rf"\b{view}\b", head or "", re.I):
        raise ValueError(f"หัวตารางไม่ใช่มุมมอง {view}: {head[:80]!r}")
    if not re.search(r"Buy.*Sell.*Total.*Net", head or "", re.S | re.I):
        raise ValueError(f"ลำดับหัวคอลัมน์ไม่ใช่ Buy,Sell,Total,Net: {head[:120]!r}")


def parse_view_rows(raw_rows) -> dict:
    """แปลงแถวดิบของตาราง -> {symbol: ([Buy, Sell, Total, Net, %], ผ่านสูตรตรวจ?)}
    ไม่ฟิกตำแหน่งคอลัมน์ เพราะเว็บมีช่องแทรกหลัง Symbol: ลองช่วง 5 ช่องตัวเลข
    จากขวาไปซ้าย แล้วเลือกช่วงแรกที่ Total=Buy+Sell และ Net=Buy-Sell"""
    out = {}
    for cells in raw_rows:
        cells = [(c or "").strip() for c in cells]
        if len(cells) < 6:
            continue
        sym = None
        for c in cells[:-5]:
            sym = next((tok for tok in c.split() if SYMBOL_RE.match(tok)), None)
            if sym:
                break
        if sym is None or sym in NOT_SYMBOLS:
            continue
        chosen = None
        for start in range(len(cells) - 5, 0, -1):
            window = cells[start:start + 5]
            if not all(_is_numeric_cell(c) for c in window):
                continue
            nums = [_num_or_none(c) for c in window]
            if _identity_ok(nums):
                chosen = nums
                break
        if chosen is not None:
            out[sym] = (chosen, True)
        else:
            out[sym] = ([_num_or_none(c) for c in cells[-5:]], False)
    return out


def _scale(vals, mult):
    out = []
    for i, v in enumerate(vals):
        if v is not None and i < 4 and mult != 1:
            v = v * mult
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        out.append(v)
    return out


def merge_views(vol: dict, vol_mult: int, val: dict, val_mult: int):
    blank = [None] * 5
    return [[s] + _scale(vol.get(s, blank), vol_mult) + _scale(val.get(s, blank), val_mult)
            for s in sorted(set(vol) | set(val))]


def check_alignment(vol: dict, val: dict) -> int:
    """ถ้าหลายแถวไม่เข้าสูตร Total=Buy+Sell, Net=Buy-Sell แปลว่าคอลัมน์อาจเลื่อน -> ไม่บันทึก"""
    bad = [("Volume", s) for s, (_, ok) in vol.items() if not ok]
    bad += [("Value", s) for s, (_, ok) in val.items() if not ok]
    total = len(vol) + len(val)
    if len(bad) > max(2, 0.01 * total):
        raise ValueError(f"ตรวจ Total=Buy+Sell และ Net=Buy-Sell ไม่ผ่าน {len(bad)} จาก {total} แถว "
                         f"(คอลัมน์อาจเลื่อน) ตัวอย่าง: {bad[:5]}")
    return len(bad)


def sanity_check(vol: dict, val: dict):
    n = max(len(vol), len(val))
    if n < MIN_EXPECTED_ROWS:
        raise ValueError(f"เก็บได้เพียง Volume {len(vol)} หุ้น / Value {len(val)} หุ้น "
                         f"น้อยผิดปกติ (ขั้นต่ำ {MIN_EXPECTED_ROWS}) ไม่บันทึก")
    if min(len(vol), len(val)) < 0.9 * n:
        raise ValueError(f"จำนวนหุ้นของ Volume ({len(vol)}) กับ Value ({len(val)}) "
                         "ต่างกันมาก อาจโหลดไม่ครบ ไม่บันทึก")


# ---------- ส่วนที่ทำงานกับเบราว์เซอร์ ----------
def _head_text(page) -> str:
    return page.evaluate(JS_HEAD) or ""


TOGGLE_ALIASES = {"Volume": ["volume", "ปริมาณ"], "Value": ["value", "มูลค่า"]}

# คลิกด้วย DOM event โดยตรง (ใช้เป็นทางสุดท้าย เผื่อปุ่มถูกบัง/ไม่ใช่ <button> ปกติ)
# เลือกองค์ประกอบ "ใบ" ที่ข้อความตรงชื่อมุมมอง ใกล้ข้อความ 'Categorized by' ก่อน
JS_CLICK_TOGGLE = r"""(names) => {
  const norm = s => (s || '').replace(/\s+/g, ' ').trim().toLowerCase();
  const wanted = names.map(norm);
  const isTarget = e => wanted.includes(norm(e.textContent)) && e.children.length <= 2;
  const all = Array.from(document.querySelectorAll('body *'));
  const fire = el => {
    for (const t of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
      el.dispatchEvent(new MouseEvent(t, {bubbles: true, cancelable: true, view: window}));
    }
  };
  const leaves = list => list.filter(e => !Array.from(e.querySelectorAll('*')).some(isTarget));
  const labels = all.filter(e => norm(e.textContent).includes('categorized by') && norm(e.textContent).length < 160)
                    .sort((a, b) => norm(a.textContent).length - norm(b.textContent).length);
  for (const l of labels) {
    const scope = l.parentElement || l;
    const hits = leaves(Array.from(scope.querySelectorAll('*')).filter(isTarget));
    if (hits.length) { fire(hits[0]); return {clicked: true, where: 'near-label', tag: hits[0].tagName}; }
  }
  const hits = leaves(all.filter(isTarget));
  if (hits.length) { fire(hits[0]); return {clicked: true, where: 'page', tag: hits[0].tagName}; }
  const sc = labels.length ? (labels[0].parentElement || labels[0]) : null;
  let html = null;
  if (sc) {
    const h = sc.outerHTML.replace(/\s+/g, ' ');
    const i = Math.max(0, h.toLowerCase().indexOf('categorized by'));
    html = h.slice(Math.max(0, i - 120), i + 600);
  }
  return {clicked: false, html: html};
}"""


def _view_active(page, view: str) -> bool:
    return bool(re.search(rf"\b{view}\b", _head_text(page), re.I))


def _wait_view(page, view: str, timeout_s: float) -> bool:
    for _ in range(int(timeout_s / 0.5)):
        if _view_active(page, view):
            return True
        page.wait_for_timeout(500)
    return _view_active(page, view)


def ensure_view(page, view: str):
    """ให้ตารางอยู่ในมุมมอง Volume หรือ Value ถ้ายังไม่ใช่จะลองกดปุ่มสลับหลายวิธี
    ตามลำดับ (คลิกจริง -> คลิกแบบบังคับ -> ส่ง DOM event) และตรวจทุกครั้งว่าหัวตาราง
    เปลี่ยนจริงหรือไม่ ถ้าไม่สำเร็จจะแสดงโครงสร้าง HTML รอบปุ่มไว้ใน error"""
    if _view_active(page, view):
        return
    page.evaluate(JS_SCROLL_TOP)
    page.wait_for_timeout(600)

    rx = re.compile(r"^\s*(?:" + "|".join(re.escape(a) for a in TOGGLE_ALIASES[view]) + r")\s*$", re.I)
    locators = [
        page.locator("xpath=//*[contains(normalize-space(text()),'Categorized by')]/..").get_by_text(rx),
        page.get_by_role("button", name=rx),
        page.get_by_role("tab", name=rx),
        page.get_by_role("radio", name=rx),
        page.get_by_role("link", name=rx),
        page.get_by_text(rx),
    ]
    for loc in locators:
        try:
            if loc.count() == 0:
                continue
        except Exception:
            continue
        for force in (False, True):
            try:
                loc.first.click(timeout=3000, force=force)
            except Exception:
                continue
            if _wait_view(page, view, 6):
                page.wait_for_timeout(1500)
                return

    info = page.evaluate(JS_CLICK_TOGGLE, TOGGLE_ALIASES[view]) or {}
    if info.get("clicked") and _wait_view(page, view, 8):
        page.wait_for_timeout(1500)
        return

    snippet = info.get("html") or "ไม่พบข้อความ 'Categorized by' บนหน้า"
    raise RuntimeError(f"สลับเป็นมุมมอง '{view}' ไม่สำเร็จ | HTML ใกล้ปุ่ม: {snippet}")


def collect_view(page):
    """เลื่อนตารางไปเรื่อย ๆ เก็บแถวสะสมตาม Symbol จนจำนวนหุ้นไม่เพิ่ม
    คืนค่า (dict ของหุ้น, ข้อความหัวตาราง, จำนวนรอบที่เลื่อน)"""
    page.evaluate(JS_SCROLL_TOP)
    page.wait_for_timeout(800)
    seen, head, stagnant, steps = {}, "", 0, 0
    for steps in range(1, MAX_SCROLL_STEPS + 1):
        data = page.evaluate(JS_EXTRACT)
        if not data:
            raise RuntimeError("ไม่พบตารางบนหน้าเว็บ")
        head = data.get("head") or head
        before = len(seen)
        seen.update(parse_view_rows(data["rows"]))
        page.evaluate(JS_SCROLL_STEP)
        page.wait_for_timeout(450)
        stagnant = 0 if len(seen) > before else stagnant + 1
        if stagnant >= STAGNANT_LIMIT:
            break
    return seen, head, steps


def fetch_nvdr_data():
    """คืนค่า (date_str 'dd/mm/yyyy', rows, info)"""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 1400})
        try:
            page.goto(NVDR_URL, wait_until="domcontentloaded", timeout=60000)
            try:
                page.wait_for_selector("table", timeout=30000)
            except PlaywrightTimeoutError:
                raise RuntimeError("รอตาราง (<table>) บนหน้า NVDR ไม่เจอ (URL อาจไม่ถูกต้อง)")
            page.wait_for_timeout(3000)
            as_of = parse_as_of(page.evaluate("() => document.body.innerText || ''"))

            ensure_view(page, "Volume")
            vol, vol_head, vol_steps = collect_view(page)
            check_head(vol_head, "Volume")

            ensure_view(page, "Value")
            val, val_head, val_steps = collect_view(page)
            check_head(val_head, "Value")

            if as_of is None:
                as_of = parse_as_of(page.evaluate("() => document.body.innerText || ''"))
        finally:
            browser.close()

    if as_of is None:
        raise ValueError("หาข้อความ 'As of dd Mon yyyy' บนหน้าเว็บไม่เจอ")
    sanity_check(vol, val)
    bad = check_alignment(vol, val)
    vol_n = {s: v[0] for s, v in vol.items()}
    val_n = {s: v[0] for s, v in val.items()}

    vm, am = unit_multiplier(vol_head), unit_multiplier(val_head)
    rows = merge_views(vol_n, vm, val_n, am)
    info = (f"{len(rows)} หุ้น (Volume {len(vol_n)}/Value {len(val_n)}) "
            f"แรก={rows[0][0]} ท้าย={rows[-1][0]} หน่วย x{vm}/x{am} "
            f"ไม่เข้าสูตร {bad} แถว เลื่อน {vol_steps}+{val_steps} รอบ")
    return as_of.strftime("%d/%m/%Y"), rows, info


# ---------- Google Sheets ----------
def get_open_spreadsheet():
    sheet_id = os.environ.get("GSHEET_ID")
    creds_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if not sheet_id or not creds_path:
        raise RuntimeError("ไม่พบ GSHEET_ID หรือ GOOGLE_APPLICATION_CREDENTIALS ใน environment variables")
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    return gspread.authorize(creds).open_by_key(sheet_id)


def get_or_create_worksheet(sh, title: str, rows: int, cols: int):
    def norm(s):
        return re.sub(r"[^a-zA-Z0-9]", "", s).lower()

    def find():
        for ws in sh.worksheets():
            if norm(ws.title) == norm(title):
                return ws
        return None

    ws = find()
    if ws:
        return ws
    try:
        return sh.add_worksheet(title=title, rows=rows, cols=cols)
    except gspread.exceptions.APIError as e:
        if "already exists" not in str(e):
            raise
        ws = find()
        if ws:
            return ws
        raise


def date_already_captured(values, date_str: str) -> bool:
    return date_str in {(r[0] if r else "").lstrip("'").strip() for r in values}


def style_headers(ws):
    """จัดหัวตาราง 2 แถวให้หน้าตาใกล้ไฟล์ Export (ผสานเซลล์ ตัวหนา สีส้ม ตรึง 2 แถว)
    ถ้าตกแต่งไม่สำเร็จไม่กระทบการบันทึกข้อมูล"""
    try:
        ws.merge_cells("C1:G1")
        ws.merge_cells("H1:L1")
        ws.format("A1:L2", {
            "textFormat": {"bold": True},
            "horizontalAlignment": "CENTER",
            "backgroundColor": {"red": 0.98, "green": 0.69, "blue": 0.23},
        })
        ws.freeze(rows=2)
    except Exception as e:
        print(f"  คำเตือน: ตกแต่งหัวตารางไม่สำเร็จ ({type(e).__name__}) ข้ามไป")


def push_to_log(sh, date_str, time_str, workflow_name, trigger_label, status, rows_sent, detail=""):
    ws = get_or_create_worksheet(sh, LOG_SHEET_NAME, rows=2000, cols=10)
    if len(ws.get_all_values()) == 0:
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
    print(f"[{date_str} {time_str} เวลาไทย] เริ่ม capture NVDR Trading by Stock "
          f"(workflow: {workflow_name}, trigger: {trigger_label})")

    sh = get_open_spreadsheet()

    try:
        data_date, rows, info = fetch_nvdr_data()
    except Exception as e:
        detail = f"{type(e).__name__}: {str(e)[:900]}"
        push_to_log(sh, date_str, time_str, workflow_name, trigger_label, "Failed", 0, detail)
        raise

    print(f"  ข้อมูลวันที่ {data_date}: {info}")

    ws = get_or_create_worksheet(sh, NVDR_SHEET_NAME, rows=len(rows) + 1000, cols=15)
    existing = ws.get_all_values()
    try:
        if not existing:
            ws.append_rows([NVDR_GROUP_ROW, NVDR_COL_ROW], value_input_option="USER_ENTERED",
                           table_range="A1")
            style_headers(ws)
        elif [c.strip() for c in (existing[1] if len(existing) > 1 else [])][:12] != NVDR_COL_ROW:
            raise ValueError(f"แท็บ {NVDR_SHEET_NAME} มีข้อมูลรูปแบบเก่า (ไม่มีหัวตาราง 2 แถวตามรูปแบบใหม่) "
                             "กรุณาลบข้อมูลทั้งหมดในแท็บนี้ก่อน แล้วรันใหม่")
    except Exception as e:
        detail = f"{type(e).__name__}: {str(e)[:900]}"
        push_to_log(sh, date_str, time_str, workflow_name, trigger_label, "Failed", 0, detail)
        raise

    if date_already_captured(existing, data_date):
        print(f"  มีข้อมูลวันที่ {data_date} อยู่แล้ว -> ไม่บันทึกซ้ำ")
        push_to_log(sh, date_str, time_str, workflow_name, trigger_label, "AlreadyCaptured", 0,
                    f"มีข้อมูล NVDR ของวันที่ {data_date} อยู่แล้ว")
        return

    # ขึ้นต้นด้วย ' ให้ Sheets เก็บวันที่เป็นข้อความล้วน (เหมือน SectorIndex)
    def _cell(v):
        return (0 if FILL_ZERO_FOR_BLANK else "") if v is None else v
    out = [[f"'{data_date}", r[0]] + [_cell(v) for v in r[1:]] for r in rows]
    try:
        ws.append_rows(out, value_input_option="USER_ENTERED", table_range="A1")
    except Exception as e:
        detail = f"{type(e).__name__}: {str(e)[:900]}"
        push_to_log(sh, date_str, time_str, workflow_name, trigger_label, "Failed", 0, detail)
        raise

    print(f"  ส่ง {len(out)} แถว เข้า worksheet '{NVDR_SHEET_NAME}'")
    push_to_log(sh, date_str, time_str, workflow_name, trigger_label, "Success", len(out),
                f"วันที่ {data_date} | {info}")


if __name__ == "__main__":
    try:
        capture_once()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่าง capture:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
