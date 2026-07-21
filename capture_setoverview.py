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


def explore_page():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 2400})
        try:
            page.goto(SET_HOME_URL, wait_until="domcontentloaded", timeout=60000)
            try:
                page.wait_for_selector("table", timeout=30000)
            except PlaywrightTimeoutError:
                print("คำเตือน: รอตาราง (<table>) ไม่เจอภายในเวลาที่กำหนด")

            # เผื่อข้อมูลบางส่วนโหลดช้ากว่านั้นอีกเล็กน้อย
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

            # ดึงข้อความในกล่องสรุปด้านบนสุด (SET / mai / TFEX) เผื่อไม่ได้อยู่ใน table
            # ลองดึงข้อความของ element ที่มีคำว่า "SET" "mai" "TFEX" ปนอยู่ในบริเวณ
            # ต้นหน้า (1500px แรกจากบน) มาดูโครงสร้างคร่าว ๆ
            top_text = page.evaluate(
                """
                () => {
                    const body = document.body.innerText || "";
                    return body.slice(0, 3000);
                }
                """
            )
        finally:
            browser.close()

    print("=" * 70)
    print(f"[diag] เจอ <table> ทั้งหมด {diag['totalCount']} ตัว, มองเห็นได้ {diag['visibleCount']} ตัว")
    print("=" * 70)

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

    print("=" * 70)
    print("[diag] ข้อความ 3000 ตัวอักษรแรกของหน้า (body.innerText) "
          "สำหรับหาตัวเลข Last/Change/%Change ที่อาจไม่ได้อยู่ใน <table>:")
    print("=" * 70)
    print(top_text)


if __name__ == "__main__":
    try:
        explore_page()
    except Exception:
        print("เกิดข้อผิดพลาดระหว่างสำรวจหน้า:")
        traceback.print_exc(limit=5, file=sys.stdout)
        sys.exit(1)
