"""
market_holidays.py
--------------------
โมดูลกลางสำหรับเช็คว่าวันนี้เป็นวันหยุดทำการซื้อขายของตลาดหลักทรัพย์ไทย
หรือไม่ (รวมเสาร์-อาทิตย์ และวันหยุดนักขัตฤกษ์ที่ตรงกับวันธรรมดา) ดึงข้อมูล
จากปฏิทินวันหยุดทางการของ SET โดยตรง:
    https://www.set.or.th/th/about/event-calendar/holiday?year={ปีค.ศ.}

ใช้ร่วมกันได้กับทุกสคริปต์ capture ในชุดนี้ ดึงจากปฏิทินสดทุกครั้ง ไม่
hardcode วันที่ไว้ในโค้ด จึงตามปีใหม่/วันหยุดพิเศษที่ประกาศเพิ่มได้เอง
อัตโนมัติ ไม่ต้องแก้โค้ดในอนาคต

วิธีใช้ในสคริปต์อื่น:
    from market_holidays import is_market_holiday
    if is_market_holiday(today):
        ... ข้าม capture รอบนี้ ...
"""

import re
import datetime as dt

from playwright.sync_api import sync_playwright

HOLIDAY_URL_TEMPLATE = "https://www.set.or.th/th/about/event-calendar/holiday?year={year}"

THAI_MONTHS_FULL = {
    "มกราคม": 1, "กุมภาพันธ์": 2, "มีนาคม": 3, "เมษายน": 4,
    "พฤษภาคม": 5, "มิถุนายน": 6, "กรกฎาคม": 7, "สิงหาคม": 8,
    "กันยายน": 9, "ตุลาคม": 10, "พฤศจิกายน": 11, "ธันวาคม": 12,
}

_DAY_OF_WEEK_ALT = "|".join([
    "วันจันทร์", "วันอังคาร", "วันพุธ", "วันพฤหัสบดี",
    "วันศุกร์", "วันเสาร์", "วันอาทิตย์",
])

_ENTRY_RE = re.compile(rf"(\d{{1,2}})\s*\n(.+?)\n(?:{_DAY_OF_WEEK_ALT})", re.MULTILINE)

_holiday_cache = {}


def _extract_holidays_from_text(body_text: str, year: int) -> set:
    month_positions = []
    for name, num in THAI_MONTHS_FULL.items():
        idx = 0
        while True:
            idx = body_text.find(name, idx)
            if idx == -1:
                break
            month_positions.append((idx, num))
            idx += len(name)
    month_positions.sort()

    holidays = set()
    for i, (pos, month_num) in enumerate(month_positions):
        end = month_positions[i + 1][0] if i + 1 < len(month_positions) else len(body_text)
        segment = body_text[pos:end]
        for m in _ENTRY_RE.finditer(segment):
            day = int(m.group(1))
            if 1 <= day <= 31:
                try:
                    holidays.add(dt.date(year, month_num, day))
                except ValueError:
                    pass
    return holidays


def get_holidays_for_year(year: int) -> set:
    if year in _holiday_cache:
        return _holiday_cache[year]

    url = HOLIDAY_URL_TEMPLATE.format(year=year)
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=60000)
                page.wait_for_timeout(3000)
                body_text = page.evaluate("() => document.body.innerText || ''")
            finally:
                browser.close()
        holidays = _extract_holidays_from_text(body_text, year)
    except Exception as e:
        print(f"คำเตือน: โหลดปฏิทินวันหยุดปี {year} ไม่สำเร็จ -> "
              f"{type(e).__name__}: {str(e)[:150]} (ถือว่าไม่มีวันหยุดในปีนี้)")
        holidays = set()

    _holiday_cache[year] = holidays
    return holidays


def is_market_holiday(check_date: dt.date) -> bool:
    if check_date.weekday() >= 5:
        return True
    holidays = get_holidays_for_year(check_date.year)
    return check_date in holidays


if __name__ == "__main__":
    today = dt.date.today()
    holidays = get_holidays_for_year(today.year)
    print(f"ปี {today.year}: เจอวันหยุด {len(holidays)} วัน")
    for d in sorted(holidays):
        print(" ", d)
    print(f"\nวันนี้ ({today}) เป็นวันหยุดตลาดหรือไม่: {is_market_holiday(today)}")
