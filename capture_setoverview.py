"""
capture_setoverview.py
-----------------------
[เวอร์ชันสำรวจโครงสร้าง - diagnostic] Capture สรุปภาพรวมตลาดจากหน้า
https://www.set.or.th/en/home เพื่อเติมข้อมูลลงแท็บ SetDatabase
(Date, Set Index, Chg, Chg%, Value(MB), Inst.Buy/Sell/Net,
Prop.Buy/Sell/Net, Foreign.Buy/Sell/Net, Indiv.Buy/Sell/Net,
SET50 Idx/Chg%/Val(MB), mai Idx/Chg%/Val(MB), TFEX Vol/OI/Prem/Disc)

เนื่องจากยังไม่เคยสำรวจโครงสร้างหน้านี้มาก่อน สคริปต์รอบนี้จะ "แค่พิมพ์ข้อมูล
วินิจฉัย" (diagnostic) ออกมาให้ดูก่อน ยังไม่ส่งข้อมูลเข้า Google Sheet จริง
เพื่อให้แน่ใจว่า field mapping ถูกต้องก่อนค่อยเปิดใช้งานจริงในรอบถัดไป

วิธีทำงาน:
1. เปิดหน้า https://www.set.or.th/en/home ด้วย headless browser (Playwright)
2. รอให้ตารางข้อมูลโหลดเสร็จ
3. ดึงตารางที่มองเห็นอยู่จริงทั้งหมด (เหมือนเทคนิคที่ใช้กับ settrade.com)
   แล้วพิมพ์โครงสร้าง (จำนวนแถว/คอลัมน์, ชื่อคอลัมน์, ตัวอย่างข้อมูล) ออกมา
4. ดึงข้อความในกล่องสรุป SET/mai/TFEX ด้านบนสุดของหน้า (ตัวเลข Last/Change/
   %Change ที่อาจไม่ได้อยู่ใน <table>) มาพิมพ์ดูโครงสร้างเช่นกัน

การใช้งาน (ทดสอบในเครื่องตัวเอง):
    python capture_setoverview.py
"""

import io
import sys
import traceback

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

SET_HOME_URL = "https://www.set.or.th/en/home"
MAI_OVERVIEW_URL = "https://www.set.or.th/en/market/index/mai/overview"


def explore_url(page, url, label):
    print("\n" + "#" * 70)
    print(f"# สำรวจหน้า: {label} ({url})")
    print("#" * 70)

    page.goto(url, wait_until="domcontentloaded", timeout=60000)
    try:
        page.wait_for_selector("table", timeout=15000)
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
    top_text = page.evaluate("() => (document.body.innerText || '').slice(0, 2500)")

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
            print(t.head(15).to_string())
            print()

    print(f"\n[diag] ข้อความ 2500 ตัวอักษรแรกของหน้า {label} (body.innerText):")
    print(top_text)


def explore_page():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 2400})
        try:
            explore_url(page, SET_HOME_URL, "SET Home")
            explore_url(page, MAI_OVERVIEW_URL, "mai Overview")
        finally:
            browser.close()


if __name__ == "__main__":
    try:
        explore_page()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่างสำรวจหน้า:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
