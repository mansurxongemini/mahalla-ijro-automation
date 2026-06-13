"""
ERP avtomatlashtirish — mahalla.ijro.uz "Оила қўшиш" oynasi orqali o'quvchilarni
ПИНФЛ (ЖШШИР) bo'yicha ro'yxatga olish.

Sayt: Angular + PrimeNG.
  - Element id'lari (pn_id_xxx) HAR SAFAR o'zgaradi -> ularga tayanmaymiz.
  - Ogohlantirish (sariq) toast'i <exclamationtriangleicon> ikonkasi bilan keladi.
  - HOLAT IZOLYATSIYASI: har bir o'quvchidan keyin modal YOPILIB, keyingisi uchun
    QAYTADAN OCHILADI. Shu sabab har bir qidiruv toza holatdan boshlanadi.
    (Avval modalni ochiq qoldirish Case B dan keyin "none" xatosini keltirib chiqargan.)

ALGORITM:
  Har bir ЖШШИР uchun:
    0. Modal ochiladi ("+"; Ҳужжат тури sukut bo'yicha "Фуқаролик пасспорти").
    1. ЖШШИР kiritiladi -> "Қидириш..." bosiladi.
    2. Natija:
       A) SARIQ ogohlantirish toast  -> Excel QIZIL.
       B) "Давом этиш" dialogi        -> tugma bosiladi (o'tkaziladi).
       C) Ф.И.Ш to'ldi (toast yo'q)   -> Хонадон random + Оила аъзо тури "Бошқа" -> "Сақлаш":
            C1) modal O'ZI YOPILADI   -> Excel YASHIL (muvaffaqiyat).
            C2) SARIQ toast (хонадон) -> Excel SARIQ (dublikat).
    3. Modal ochiq qolgan bo'lsa yopiladi (keyingisi toza boshlanishi uchun).

ISHGA TUSHIRISH:
    pip install playwright openpyxl
    playwright install chromium
    python test.py
  Birinchi marta brauzer ochilganda QO'LDA login qiling va kerakli sahifani oching,
  keyin terminalda ENTER bosing. Profil saqlanadi (.pw_profile).
"""

import asyncio
import random
import sys

import openpyxl
from openpyxl.styles import PatternFill
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

# ============================================================================
#  SOZLAMALAR
# ============================================================================

EXCEL_FILE = "15.08.2025 ERP.xlsx"
JSHSHIR_COLUMN = 6           # ПИНФЛ ustuni = F (6).  DIQQAT: 5 = "Пол"!
START_ROW = 3                # data 3-qatordan
END_ROW = None               # None = oxirigacha. Sinov uchun masalan 10.

SKIP_PROCESSED = True        # allaqachon bo'yalgan qatorlarni o'tkazib yuborish (resume)
DEBUG_SCREENSHOTS = True     # noaniq holatda screenshot saqlash (debug_row_N.png)

TARGET_URL = ("https://mahalla.ijro.uz/dashboard/list/family"
              "?region_id=00s0eed0000region000008"
              "&district_id=00s0eed0000region000049"
              "&mahalla_id=66016a6ae237a52f91961c45"
              "&returnPath=%2Fdashboard%2Fassistance&page=1&limit=20&offset=0")

PROFILE_DIR = ".pw_profile"
HEADLESS = False
SLOW_MO = 0
SAVE_EVERY = 1

FAMILY_TYPE_LABEL = "Бошқа"

SEARCH_TIMEOUT_MS = 9000
SAVE_TIMEOUT_MS = 9000
POLL_MS = 250

# ============================================================================
#  SELEKTORLAR
# ============================================================================

DIALOG = "app-create-family-dialog"
DIALOG_WRAP = "div.p-dialog:has(app-create-family-dialog)"          # asosiy modal konteyneri
DIALOG_CLOSE = f"{DIALOG_WRAP} button[aria-label='Close']"          # X tugma

BTN_ADD = "button.p-button-icon-only.rounded-full:has(.pi-plus)"

JSHSHIR_INPUT = "xpath=//label[contains(normalize-space(.),'ЖШШИР')]/following-sibling::input[1]"
SEARCH_BTN = f"{DIALOG} button:has-text('Қидириш')"
FISH_INPUT = "xpath=//label[contains(normalize-space(.),'Ф.И.Ш')]/following-sibling::input[1]"

HOUSE_DROPDOWN = "p-dropdown[formcontrolname='id']"
FAMILY_TYPE_DROPDOWN = "p-dropdown[formcontrolname='type']"
SAVE_BTN = "app-create-family-dialog-footer button"

CASE_B_DIALOG = "app-citizen-family-info-dialog"
CONTINUE_BTN = f"{CASE_B_DIALOG} button:has-text('Давом этиш')"

# Ogohlantirish (sariq) toast — exclamationtriangle ikonkasi yoki warn klass bo'yicha
WARN_TOAST = ".p-toast-message-warn, p-toast .p-toast-message:has(exclamationtriangleicon)"
ANY_DIALOG_CLOSE = "div.p-dialog button[aria-label='Close']"
TOAST_DETAIL = ".p-toast-detail"
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
    fill = cell.fill
    if fill is None or fill.fill_type != "solid":
        return False
    color = (fill.start_color.rgb or "").upper()
    return any(c in color for c in ("FFC7CE", "C6EFCE", "FFEB9C"))


async def is_visible(page, selector):
    try:
        return await page.locator(selector).first.is_visible()
    except Exception:
        return False


async def dismiss_toasts(page):
    """Mavjud toast'larni yopamiz — stale o'qishning oldini olish uchun."""
    try:
        closers = page.locator(TOAST_CLOSE)
        for i in range(await closers.count()):
            try:
                await closers.nth(i).click(timeout=300)
            except Exception:
                pass
    except Exception:
        pass


async def toast_detail_text(page):
    loc = page.locator(TOAST_DETAIL)
    try:
        if await loc.count() == 0:
            return None
        el = loc.last
        if await el.is_visible():
            return (await el.inner_text()).strip()
    except Exception:
        pass
    return None


async def open_modal(page):
    """'Оила қўшиш' modalini ochib, ЖШШИР input chiqishini kutadi."""
    await page.locator(BTN_ADD).first.click()
    await page.locator(JSHSHIR_INPUT).first.wait_for(state="visible", timeout=10000)
    await page.wait_for_timeout(300)  # forma settle bo'lishi uchun


async def close_all_dialogs(page):
    """Ochiq barcha p-dialog'larni X orqali yopadi (Case B + asosiy modal)."""
    for _ in range(4):
        btns = page.locator(ANY_DIALOG_CLOSE)
        try:
            cnt = await btns.count()
        except Exception:
            cnt = 0
        if cnt == 0:
            break
        try:
            await btns.last.click(timeout=600)
        except Exception:
            try:
                await page.keyboard.press("Escape")
            except Exception:
                pass
        await page.wait_for_timeout(250)


async def open_primeng_dropdown(page, selector):
    await page.locator(selector).first.click()
    await page.locator(".p-dropdown-panel").last.wait_for(state="visible", timeout=5000)


async def pick_random_house(page):
    await open_primeng_dropdown(page, HOUSE_DROPDOWN)
    options = page.get_by_role("option")
    try:
        await options.first.wait_for(state="visible", timeout=4000)
    except PWTimeout:
        await page.keyboard.press("Escape")
        return False
    count = await options.count()
    if count == 0:
        await page.keyboard.press("Escape")
        return False
    await options.nth(random.randint(0, count - 1)).click()
    return True


async def pick_family_type(page, label):
    await open_primeng_dropdown(page, FAMILY_TYPE_DROPDOWN)
    opt = page.get_by_role("option", name=label, exact=True)
    if await opt.count() == 0:
        opt = page.get_by_role("option", name=label)
    await opt.first.click()


async def wait_search_result(page):
    """
    'warn'   -> Case A (sariq ogohlantirish)
    'dialog' -> Case B
    'filled' -> Case C (Ф.И.Ш to'ldi)
    'none'   -> hech narsa
    """
    elapsed = 0
    while elapsed < SEARCH_TIMEOUT_MS:
        if await is_visible(page, CASE_B_DIALOG):
            return "dialog"
        if await is_visible(page, WARN_TOAST):
            return "warn"
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
    'success' -> C1: modal O'ZI YOPILADI
    'warn'    -> C2: sariq toast (хонадон дубликат), modal ochiq
    'none'    -> timeout
    """
    elapsed = 0
    while elapsed < SAVE_TIMEOUT_MS:
        if await is_visible(page, WARN_TOAST):
            return "warn"
        if not await is_visible(page, DIALOG):
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

        for row in range(START_ROW, last_row + 1):
            cell = sheet.cell(row=row, column=JSHSHIR_COLUMN)
            raw = cell.value
            if not raw:
                continue
            if SKIP_PROCESSED and already_processed(cell):
                stats["skip"] += 1
                continue

            jshshir = str(raw).strip()
            if not jshshir.isdigit() or len(jshshir) != 14:
                log(f"[{row}] NOTO'G'RI ЖШШИР: {jshshir!r}")
                stats["err"] += 1
                continue

            log(f"[{row}] ЖШШИР: {jshshir}")
            try:
                # Har doim toza holatdan: eski toast/dialoglarni tozalab, modalni ochamiz
                await dismiss_toasts(page)
                await close_all_dialogs(page)
                await open_modal(page)

                await page.locator(JSHSHIR_INPUT).first.fill(jshshir)
                await page.locator(SEARCH_BTN).first.click()

                result = await wait_search_result(page)

                if result == "warn":
                    detail = await toast_detail_text(page)
                    log(f"  -> Case A: ogohlantirish ({detail!r}). Excel QIZIL.")
                    cell.fill = RED_FILL
                    stats["A"] += 1
                    await close_all_dialogs(page)

                elif result == "dialog":
                    log("  -> Case B: 'Давом этиш' bosildi, o'tkazildi.")
                    try:
                        await page.locator(CONTINUE_BTN).first.click(timeout=3000)
                    except Exception:
                        pass
                    await page.wait_for_timeout(500)
                    stats["B"] += 1
                    await close_all_dialogs(page)

                elif result == "filled":
                    log("  -> Case C: ma'lumot topildi. Xonadon + Оила аъзо тури to'ldirilmoqda.")
                    if not await pick_random_house(page):
                        log("     !! Xonadon ro'yxati bo'sh — o'tkazildi.")
                        stats["err"] += 1
                        await close_all_dialogs(page)
                        if (row - START_ROW + 1) % SAVE_EVERY == 0:
                            wb.save(EXCEL_FILE)
                        continue
                    await pick_family_type(page, FAMILY_TYPE_LABEL)
                    await dismiss_toasts(page)
                    await page.locator(SAVE_BTN).first.click()

                    save_res = await wait_save_result(page)
                    if save_res == "success":
                        log("  -> C1: MUVAFFAQIYATLI. Excel YASHIL.")
                        cell.fill = GREEN_FILL
                        stats["C1"] += 1
                    elif save_res == "warn":
                        detail = await toast_detail_text(page)
                        log(f"  -> C2: xonadon dublikat ({detail!r}). Excel SARIQ.")
                        cell.fill = YELLOW_FILL
                        stats["C2"] += 1
                        await close_all_dialogs(page)
                    else:
                        log(f"  -> ?: saqlash natijasi noma'lum.")
                        stats["err"] += 1
                        if DEBUG_SCREENSHOTS:
                            await page.screenshot(path=f"debug_row_{row}_save.png")
                        await close_all_dialogs(page)

                else:  # none
                    detail = await toast_detail_text(page)
                    log(f"  -> ?: qidiruv natijasi aniqlanmadi. Toast={detail!r}")
                    if DEBUG_SCREENSHOTS:
                        await page.screenshot(path=f"debug_row_{row}_search.png")
                    stats["err"] += 1
                    await close_all_dialogs(page)

            except Exception as e:
                log(f"  !! XATO [{row}]: {e}")
                stats["err"] += 1
                if DEBUG_SCREENSHOTS:
                    try:
                        await page.screenshot(path=f"debug_row_{row}_error.png")
                    except Exception:
                        pass
                await close_all_dialogs(page)

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
