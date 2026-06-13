import asyncio
from playwright.async_api import async_playwright
import openpyxl
from openpyxl.styles import PatternFill

# Excel uchun ranglar (HEX formatda)
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")     # Case A uchun qizil
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")   # Case C1 uchun yashil

EXCEL_FILE = "15_08_2025_ERP.xlsx"
JSHSHIR_COLUMN = 5  # JSHSHIR (ПИНФЛ) joylashgan ustun indeksi (E ustuni = 5)

async def main():
    # Excel faylni yuklash (faqat RAMga o'qiladi)
    wb = openpyxl.load_workbook(EXCEL_FILE)
    sheet = wb.active

    async with async_playwright() as p:
        # Tezlikni kuzatish uchun headless=False (orqa fonda ishlatish uchun True qiling)
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        # Tizimga kirish va avtorizatsiya URL
        await page.goto("TIZIM_URL_MANZILI_SHU_YERGA")
        
        # TODO: Bu yerda login/parol kiritish mantig'ini yozishingiz mumkin (agar kerak bo'lsa)
        # await page.fill("input#login", "username")
        # await page.click("button#submit")
        
        modal_open = False  # Modal oyna holatini kuzatish

        for row in range(2, sheet.max_row + 1):
            jshshir_value = sheet.cell(row=row, column=JSHSHIR_COLUMN).value
            if not jshshir_value:
                continue
            
            jshshir = str(jshshir_value).strip()
            print(f"Ishorlanmoqda: Qator {row} -> JSHSHIR: {jshshir}")

            # 1-QADAM: Agar modal yopiq bo'lsa (Boshlanish yoki C1 dan keyin)
            if not modal_open:
                await page.click("SELECTOR_PLUS_TUGMASI")
                await page.click("SELECTOR_FUQAROLIK_PASSPORTI_BO`LIMI")
                # JSHSHIR va Tug'ilgan sana opsiyasini tanlash mantig'i
                await page.click("SELECTOR_JSHSHIR_OPTION")
                modal_open = True

            # 2-QADAM: JSHSHIR kiritish va Qidirish
            await page.fill("SELECTOR_JSHSHIR_INPUT", jshshir)
            await page.click("SELECTOR_QIDIRISH_TUGMASI")

            # Server javobini va UI o'zgarishini qisqa kutish (Tezlikni oshirish uchun kichik timeout)
            await page.wait_for_timeout(1500) 

            # QIDIRUV NATIJASI TAHLILI (Qaror nuqtasi)
            
            # Case A: Sariq ogohlantirish (Boshqa MFY)
            yellow_toast = page.locator("SELECTOR_SARIQ_TOAST_NOTIFICATION")
            if await yellow_toast.is_visible() and "бошқа МФЙда" in await yellow_toast.inner_text():
                print(f"-> Case A: Boshqa MFY. Excel QIZIL rangga bo'yaldi.")
                sheet.cell(row=row, column=JSHSHIR_COLUMN).fill = RED_FILL
                # Inputni tozalab keyingisiga o'tamiz, oyna ochiq qoladi
                await page.fill("SELECTOR_JSHSHIR_INPUT", "")
                continue

            # Case B: Dialog oynasi (Fuqaro mahallalarda)
            dialog_box = page.locator("SELECTOR_DIALOG_OYNASI")
            if await dialog_box.is_visible():
                print(f"-> Case B: Dialog chiqdi. 'Davom etish' bosildi.")
                await page.click("SELECTOR_DAVOM_ETISH_KO`K_TUGMASI")
                # Inputni tozalab keyingisiga o'tamiz, oyna ochiq qoladi
                await page.fill("SELECTOR_JSHSHIR_INPUT", "")
                continue

            # Case C: Ma'lumotlar avto-to'ldi (F.I.Sh inputi bo'sh emasligini tekshirish)
            fish_input = page.locator("SELECTOR_FISH_INPUT")
            if await fish_input.get_attribute("value"):
                print(f"-> Case C: Ma'lumotlar topildi. Xonadon to'ldirilmoqda...")
                
                # Xonadon dropdown qismini ochish
                await page.click("SELECTOR_XONADON_DROPDOWN")
                await page.wait_for_timeout(500)
                
                # Random yoki birinchi xonadonni tanlash (Dropdown elementlari ro'yxatidan)
                await page.click("SELECTOR_XONADON_BIRINCHI_ELEMENT")
                
                # Oila bo'limidan 'Boshqa'ni tanlash
                await page.select_option("SELECTOR_OILA_BOLIMI_SELECT", value="Boshqa") # yoki tegishli selector
                
                # SAQLASH
                await page.click("SELECTOR_SAQLASH_TUGMASI")
                await page.wait_for_timeout(1500)

                # SAQLASH NATIJASI TAHLILI
                success_toast = page.locator("SELECTOR_YASHIL_SUCCESS_TOAST")
                
                # C1: Muvaffaqiyatli saqlandi
                if await success_toast.is_visible() and "муваффақиятли" in await success_toast.inner_text():
                    print(f"-> Case C1: Muvaffaqiyatli saqlandi. Excel YASHIL rangga bo'yaldi.")
                    sheet.cell(row=row, column=JSHSHIR_COLUMN).fill = GREEN_FILL
                    modal_open = False  # Tizim modalni o'zi yopdi, holat 0 dan boshlanadi
                    continue
                
                # C2: Xonadon allaqachon mavjud (Sariq ogohlantirish)
                if await yellow_toast.is_visible() and "тизимга киритилган" in await yellow_toast.inner_text():
                    print(f"-> Case C2: Xonadon allaqachon tizimda bor. Oyna ochiq holda o'tildi.")
                    # Oyna yopilmagan, faqat keyingi JSHSHIRga o'tiladi
                    await page.fill("SELECTOR_JSHSHIR_INPUT", "")
                    continue

        # Barcha sikllar tugagach, Excel faylni yozish (Batch Update)
        wb.save(EXCEL_FILE)
        print("!!! JARAYON YAKUNLANDI. Excel fayl saqlandi !!!")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())