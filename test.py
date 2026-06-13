"""
ERP avtomatlashtirish — mahalla.ijro.uz "Оила қўшиш" oynasi orqali o'quvchilarni
ПИНФЛ (ЖШШИР) bo'yicha ro'yxatga olish.

Sayt: Angular + PrimeNG. Shuning uchun:
  - Element id'lari (pn_id_xxx) HAR SAFAR o'zgaradi -> ularga tayanmaymiz.
  - Toast'lar PrimeNG severity klasslari bilan ajratiladi:
        .p-toast-message-warn     -> SARIQ ogohlantirish
        .p-toast-message-success  -> YASHIL muvaffaqiyat
        .p-toast-message-error    -> QIZIL xato
  - Boshqaruv elementlari label matni va formcontrolname bo'yicha topiladi.

ALGORITM:
  0. Modal yopiq bo'lsa: "+" bosiladi (Ҳужжат тури sukut bo'yicha "Фуқаролик пасспорти").
  1. ЖШШИР Excel'dan olinadi -> inputga yoziladi -> "Қидириш..." bosiladi.
  2. Qidiruv natijasi:
       A) SARIQ toast ("...бошқа МФЙда...")  -> Excel QIZIL, keyingisiga (modal ochiq).
       B) "Давом этиш" tugmali dialog        -> tugma bosiladi, keyingisiga (modal ochiq).
       C) Ф.И.Ш avto-to'ldi (toast yo'q)     -> Хонадон random + Оила аъзо тури "Бошқа" -> "Сақлаш".
            C1) YASHIL toast  -> Excel YASHIL, modal O'ZI YOPILADI -> 0-qadam.
            C2) SARIQ toast   -> Excel SARIQ, modal ochiq, keyingi ЖШШИР ga (qayta urinmaymiz).

ISHGA TUSHIRISH:
    pip install playwright openpyxl
    playwright install chromium
    python test.py
  Birinchi marta brauzer ochilganda QO'LDA login qiling va kerakli sahifani oching,
  keyin terminalda ENTER bosing. Profil saqlanadi (.pw_profile), keyingi safar login kerak emas.
"""

import asyncio
import random
import re
import sys

import openpyxl
from openpyxl.styles import PatternFill
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

# ============================================================================
#  SOZLAMALAR
# ============================================================================

EXCEL_FILE = "15.08.2025 ERP.xlsx"   # nuqta va probel bilan — haqiqiy nom
JSHSHIR_COLUMN = 6                    # ПИНФЛ ustuni = F (6).  DIQQAT: 5 = "Пол"!
START_ROW = 3                         # 1=sarlavha, 2=ustun nomlari, data 3-qatordan
END_ROW = None                        # None = oxirigacha. Sinov uchun masalan 7.

TARGET_URL = ("https://mahalla.ijro.uz/dashboard/list/family"
              "?region_id=00s0eed0000region000008"
              "&district_id=00s0eed0000region000049"
              "&mahalla_id=66016a6ae237a52f91961c45"
              "&returnPath=%2Fdashboard%2Fassistance&page=1&limit=20&offset=0")

PROFILE_DIR = ".pw_profile"           # login shu yerda saqlanadi (qayta-qayta login kerak emas)
HEADLESS = False
SLOW_MO = 0                           # ms; sekinlashtirib kuzatish uchun masalan 200
SAVE_EVERY = 1                        # har necha qatorda Excel saqlash

FAMILY_TYPE_LABEL = "Бошқа"           # "Оила аъзо тури" dan tanlanadigan qiymat

# Kutish vaqtlari
SEARCH_TIMEOUT_MS = 9000              # qidiruv natijasini kutish
SAVE_TIMEOUT_MS = 9000               # saqlash natijasini kutish
POLL_MS = 250

# ============================================================================
#  SELEKTORLAR  (HTML'dan olingan; faqat BTN_ADD taxminiy)
# ============================================================================

DIALOG = "app-create-family-dialog"

# "+" tugmasi — PrimeNG icon-only button (pi pi-plus, rounded-full)
BTN_ADD = "button.p-button-icon-only.rounded-full:has(.pi-plus)"

# ЖШШИР input — "ЖШШИР" labeldan keyingi input
JSHSHIR_INPUT = "xpath=//label[contains(normalize-space(.),'ЖШШИР')]/following-sibling::input[1]"
SEARCH_BTN = f"{DIALOG} button:has-text('Қидириш')"
# Ф.И.Ш readonly input — "Ф.И.Ш" labeldan keyingi input
FISH_INPUT = "xpath=//label[contains(normalize-space(.),'Ф.И.Ш')]/following-sibling::input[1]"

HOUSE_DROPDOWN = "p-dropdown[formcontrolname='id']"        # Хонадон
FAMILY_TYPE_DROPDOWN = "p-dropdown[formcontrolname='type']"  # Оила аъзо тури
SAVE_BTN = "app-create-family-dialog-footer button"

CONTINUE_BTN = "button:has-text('Давом')"   # Case B "Давом этиш"

TOAST_WARN = ".p-toast-message-warn"
TOAST_SUCCESS = ".p-toast-message-success"
TOAST_ERROR = ".p-toast-message-error"
TOAST_ANY = ".p-toast-message"
TOAST_CLOSE = ".p-toast-icon-close"

# Excel ranglari
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")


# ============================================================================
#  YORDAMCHI FUNKSIYALAR
# ============================================================================

def log(msg):
    print(msg, flush=True)


def already_processed(cell):
    """Resume: qator allaqachon bo'yalgan bo'lsa o'tkazib yuboramiz."""
    fill = cell.fill
    if fill is None or fill.fill_type != "solid":
        return False
    color = (fill.start_color.rgb or "").upper()
    return any(c in color for c in ("FFC7CE", "C6EFCE", "FFEB9C"))


async def visible(page, selector, timeout=600):
    try:
        await page.locator(selector).first.wait_for(state="visible", timeout=timeout)
        return True
    except PWTimeout:
        return False


async def dismiss_toasts(page):
    """Eski toast'larni yopamiz — stale o'qishning oldini olish uchun."""
    try:
        closers = page.locator(TOAST_CLOSE)
        for i in range(await closers.count()):
            try:
                await closers.nth(i).click(timeout=300)
            except Exception:
                pass
    except Exception:
        pass


async def open_modal(page):
    """0-qadam: 'Оила қўшиш' modalini ochish va ЖШШИР input chiqishini kutish."""
    await page.locator(BTN_ADD).first.click()
    await page.locator(JSHSHIR_INPUT).first.wait_for(state="visible", timeout=8000)


async def open_primeng_dropdown(page, selector):
    """PrimeNG p-dropdown'ni ochib, panel ko'rinishini kutadi."""
    await page.locator(selector).first.click()
    await page.locator(".p-dropdown-panel").last.wait_for(state="visible", timeout=5000)


async def pick_random_house(page):
    """Хонадон dropdown'dan random element tanlaydi. True/False qaytaradi."""
    await open_primeng_dropdown(page, HOUSE_DROPDOWN)
    options = page.get_by_role("option")
    # ro'yxat yuklanishini biroz kutamiz
    try:
        await options.first.wait_for(state="visible", timeout=4000)
    except PWTimeout:
        # bo'sh bo'lsa panelni yopamiz
        await page.keyboard.press("Escape")
        return False
    count = await options.count()
    if count == 0:
        await page.keyboard.press("Escape")
        return False
    await options.nth(random.randint(0, count - 1)).click()
    return True


async def pick_family_type(page, label):
    """Оила аъзо тури dropdown'dan berilgan label'ni tanlaydi."""
    await open_primeng_dropdown(page, FAMILY_TYPE_DROPDOWN)
    opt = page.get_by_role("option", name=label, exact=True)
    if await opt.count() == 0:
        opt = page.get_by_role("option", name=label)  # contains fallback
    await opt.first.click()


async def wait_search_result(page):
    """
    Qidiruv natijasini aniqlaydi:
      'warn'   -> Case A (sariq, boshqa MFY)
      'dialog' -> Case B (Давом этиш)
      'filled' -> Case C (Ф.И.Ш to'ldi)
      'error'  -> qizil xato
      'none'   -> hech narsa (timeout)
    """
    deadline = SEARCH_TIMEOUT_MS
    elapsed = 0
    while elapsed < deadline:
        if await page.locator(CONTINUE_BTN).first.is_visible():
            return "dialog"
        if await page.locator(TOAST_WARN).first.is_visible():
            return "warn"
        if await page.locator(TOAST_ERROR).first.is_visible():
            return "error"
        try:
            val = await page.locator(FISH_INPUT).first.input_value(timeout=300)
            if val and val.strip():
                return "filled"
        except Exception:
            pass
        await page.wait_for_timeout(POLL_MS)
        elapsed += POLL_MS
    return "none"


async def wait_save_result(page):
    """
    Saqlash natijasini aniqlaydi:
      'success' -> C1 (yashil)
      'warn'    -> C2 (sariq)
      'error'   -> qizil xato
      'none'    -> timeout
    """
    deadline = SAVE_TIMEOUT_MS
    elapsed = 0
    while elapsed < deadline:
        if await page.locator(TOAST_SUCCESS).first.is_visible():
            return "success"
        if await page.locator(TOAST_WARN).first.is_visible():
            return "warn"
        if await page.locator(TOAST_ERROR).first.is_visible():
            return "error"
        # success belgisi sifatida modal yopilishini ham tekshiramiz
        if not await page.locator(DIALOG).first.is_visible():
            return "success"
        await page.wait_for_timeout(POLL_MS)
        elapsed += POLL_MS
    return "none"


# ============================================================================
#  ASOSIY OQIM
# ============================================================================

async def main():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    sheet = wb.active
    last_row = END_ROW or sheet.max_row

    stats = {"A": 0, "B": 0, "C1": 0, "C2": 0, "skip": 0, "err": 0}

    async with async_playwright() as p:
        context = await p.chromium.launch_persistent_context(
            PROFILE_DIR, headless=HEADLESS, slow_mo=SLOW_MO
        )
        page = context.pages[0] if context.pages else await context.new_page()
        await page.goto(TARGET_URL)

        await asyncio.to_thread(
            input, ">>> Login qiling va kerakli sahifani oching, keyin ENTER bosing... "
        )

        modal_open = False

        for row in range(START_ROW, last_row + 1):
            cell = sheet.cell(row=row, column=JSHSHIR_COLUMN)
            raw = cell.value
            if not raw:
                continue
            if already_processed(cell):
                stats["skip"] += 1
                continue

            jshshir = str(raw).strip()
            if not jshshir.isdigit() or len(jshshir) != 14:
                log(f"[{row}] NOTO'G'RI ЖШШИР: {jshshir!r} — o'tkazildi")
                stats["err"] += 1
                continue

            log(f"[{row}] ЖШШИР: {jshshir}")
            try:
                if not modal_open:
                    await open_modal(page)
                    modal_open = True

                await dismiss_toasts(page)
                await page.locator(JSHSHIR_INPUT).first.fill("")
                await page.locator(JSHSHIR_INPUT).first.fill(jshshir)
                await page.locator(SEARCH_BTN).first.click()

                result = await wait_search_result(page)

                if result == "warn":
                    log("  -> Case A: boshqa MFY. Excel QIZIL.")
                    cell.fill = RED_FILL
                    stats["A"] += 1

                elif result == "dialog":
                    log("  -> Case B: 'Давом этиш' bosildi.")
                    await page.locator(CONTINUE_BTN).first.click()
                    await page.wait_for_timeout(500)
                    stats["B"] += 1

                elif result == "filled":
                    log("  -> Case C: ma'lumot topildi. Xonadon + Оила аъзо тури to'ldirilmoqda.")
                    if not await pick_random_house(page):
                        log("     !! Xonadon ro'yxati bo'sh — o'tkazildi.")
                        stats["err"] += 1
                        continue
                    await pick_family_type(page, FAMILY_TYPE_LABEL)
                    await dismiss_toasts(page)
                    await page.locator(SAVE_BTN).first.click()

                    save_res = await wait_save_result(page)
                    if save_res == "success":
                        log("  -> C1: MUVAFFAQIYATLI. Excel YASHIL. Modal yopildi.")
                        cell.fill = GREEN_FILL
                        modal_open = False
                        stats["C1"] += 1
                    elif save_res == "warn":
                        log("  -> C2: xonadon allaqachon tizimda. Excel SARIQ.")
                        cell.fill = YELLOW_FILL
                        stats["C2"] += 1
                    else:
                        log(f"  -> ?: saqlash natijasi noma'lum ({save_res}).")
                        stats["err"] += 1

                else:
                    log(f"  -> ?: qidiruv natijasi aniqlanmadi ({result}).")
                    stats["err"] += 1

            except Exception as e:
                log(f"  !! XATO [{row}]: {e}")
                stats["err"] += 1
                modal_open = False  # xavfsizlik uchun modalni qayta ochamiz

            if (row - START_ROW + 1) % SAVE_EVERY == 0:
                wb.save(EXCEL_FILE)

        wb.save(EXCEL_FILE)
        log("\n==== YAKUNLANDI ====")
        log(f"A  (qizil / boshqa MFY)      : {stats['A']}")
        log(f"B  (dialog / davom etish)    : {stats['B']}")
        log(f"C1 (yashil / muvaffaqiyatli) : {stats['C1']}")
        log(f"C2 (sariq / xonadon dublikat): {stats['C2']}")
        log(f"O'tkazib yuborilgan          : {stats['skip']}")
        log(f"Xato / noma'lum              : {stats['err']}")
        await context.close()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        sys.exit("\nFoydalanuvchi to'xtatdi.")
