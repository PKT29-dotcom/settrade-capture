"""
capture_sectorindex_explore.py
--------------------------------
[เวอร์ชันสำรวจโครงสร้าง - diagnostic] สำรวจหน้า "ดัชนีราคากลุ่มอุตสาหกรรมและ
หมวดธุรกิจ" ของ settrade.com เพื่อเตรียมเขียน capture_sectorindex.py เวอร์ชัน
จริง (เก็บข้อมูลไว้ทำ RRG / Sector Rotation)

URL: https://www.settrade.com/th/equities/market-data/overview?category=Industry-Sector

หน้านี้มีตารางแบบ nested (กลุ่มอุตสาหกรรมใหญ่ เช่น AGRO, CONSUMP, ... และ
หมวดธุรกิจย่อยข้างใน เช่น AGRI, FOOD, ...) และมีปุ่มสลับตลาด SET/mai คล้ายกับ
หน้า Top Ranking ที่เคยทำมาก่อน

สคริปต์นี้แค่พิมพ์โครงสร้างข้อมูลออกมาดูก่อน ยังไม่ส่งเข้า Google Sheet จริง

การใช้งาน (ทดสอบในเครื่องตัวเอง หรือรันผ่าน workflow_dispatch):
    python capture_sectorindex_explore.py
"""

import io
import sys
import traceback

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

BASE_URL = "https://www.settrade.com/th/equities/market-data/overview?category=Industry-Sector"

# ลองทั้ง 2 แบบ: URL ที่มี query param ระบุตลาดตรง ๆ (เผื่อรองรับแบบเดียวกับ
# หน้า Top Ranking) และ URL เปล่า (จะดูว่า default เป็นตลาดไหน)
CANDIDATE_URLS = {
    "default (no param)": BASE_URL,
    "market=SET": f"{BASE_URL}&market=SET",
    "market=mai": f"{BASE_URL}&market=mai",
}


def explore_url(page, url, label):
    print("\n" + "#" * 70)
    print(f"# สำรวจ: {label}")
    print(f"# URL: {url}")
    print("#" * 70)

    page.goto(url, wait_until="domcontentloaded", timeout=60000)
    try:
        page.wait_for_selector("table", timeout=20000)
    except PlaywrightTimeoutError:
        print("คำเตือน: รอตาราง (<table>) ไม่เจอภายในเวลาที่กำหนด")
    page.wait_for_timeout(4000)

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
    print(f"[diag] เจอ <table> ทั้งหมด {diag['totalCount']} ตัว, มองเห็นได้ {diag['visibleCount']} ตัว")

    for i, html in enumerate(diag["htmls"]):
        print(f"\n----- ตาราง #{i} (ยาว {len(html)} ตัวอักษร) -----")
        try:
            parsed = pd.read_html(io.StringIO(html))
        except Exception as e:
            print(f"  parse ไม่ผ่าน -> {type(e).__name__}: {str(e)[:200]}")
            continue
        for j, t in enumerate(parsed):
            print(f"  [ตาราง #{i}.{j}] shape={t.shape}")
            print(f"  columns={list(t.columns)}")
            print(t.head(40).to_string())
            print()

    # ดึงข้อความหน้าเว็บช่วงต้น เผื่อมีตัวเลขนอกตาราง หรือปุ่มสลับตลาดที่ต้องคลิก
    top_text = page.evaluate("() => (document.body.innerText || '').slice(0, 1500)")
    print(f"\n[diag] ข้อความ 1500 ตัวอักษรแรกของหน้า {label}:")
    print(top_text)

    # เช็คว่ามีปุ่ม/แท็บ SET กับ mai ให้คลิกไหม (เผื่อ URL param ใช้ไม่ได้)
    tab_info = page.evaluate(
        """
        () => {
            const candidates = Array.from(document.querySelectorAll('button, a, div, span'))
                .filter(el => {
                    const txt = (el.textContent || '').trim();
                    return (txt === 'SET' || txt === 'mai') && el.offsetParent !== null;
                })
                .slice(0, 10)
                .map(el => ({tag: el.tagName, text: el.textContent.trim(), class: el.className}));
            return candidates;
        }
        """
    )
    print(f"\n[diag] element ที่มีข้อความ 'SET' หรือ 'mai' พอดี (มองเห็นได้) ที่เจอในหน้า:")
    for t in tab_info:
        print(f"  {t}")


def explore_page():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 3000})
        try:
            for label, url in CANDIDATE_URLS.items():
                explore_url(page, url, label)
        finally:
            browser.close()


if __name__ == "__main__":
    try:
        explore_page()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่างสำรวจหน้า:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
