"""
ERP avtomatlashtirish — mahalla.ijro.uz "Оила қўшиш" oynasi orqali o'quvchilarni
ПИНФЛ (ЖШШИР) bo'yicha ro'yxatga olish.

Sayt: Angular + PrimeNG.
  - Element id'lari (pn_id_xxx) HAR SAFAR o'zgaradi -> ularga tayanmaymiz.
  - Ogohlantirish (sariq) toast'i <exclamationtriangleicon> ikonkasi bilan keladi.
  - HOLAT IZOLYATSIYASI: har bir o'quvchidan keyin modal YOPILIB, keyingisi uchun
    QAYTADAN OCHILADI.

BU VERSIYADA:
  - Modal ochilganda "Ҳужжат тури = Фуқаролик пасспорти" ANIQ tanlanadi.
  - BATAFSIL QADAM LOGI: STEP_LOG=True bo'lsa, har bir mayda amal vaqt belgisi bilan
    loglanadi (". [HH:MM:SS.mmm] ..."). Shu loglardan ketma-ketlikni tahlil qilamiz.

ISHGA TUSHIRISH:
    pip install playwright openpyxl
    playwright install chromium
    python test.py
"""

import asyncio
import datetime
import random
import sys

import openpyxl
from openpyxl.styles import PatternFill
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

# ============================================================================
#  SOZLAMALAR
# ============================================================================

EXCEL_FILE = "15.08.2025 ERP.xlsx"
JSHSHIR_COLUMN = 6
START_ROW = 3
END_ROW = None               # None = oxirigacha. Sinov uchun masalan 10.

SKIP_PROCESSED = True
DEBUG_SCREENSHOTS = True
STEP_LOG = True              # batafsil qadam-baqadam log

SELECT_DOC_TYPE = True
DOC_TYPE_VALUE = "Фуқаролик пасспорти"   # Ҳужжат тури qiymati

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
DIALOG_WRAP = "div.p-dialog:has(app-create-family-dialog)"
ANY_DIALOG_CLOSE = "div.p-dialog button[aria-label='Close']"

BTN_ADD = "button.p-button-icon-only.rounded-full:has(.pi-plus)"

DOC_TYPE_DROPDOWN = "xpath=//label[contains(normalize-space(.),'Ҳужжат тури')]/following-sibling::p-dropdown[1]"
JSHSHIR_INPUT = "xpath=//label[contains(normalize-space(.),'ЖШШИР')]/following-sibling::input[1]"
SEARCH_BTN = f"{DIALOG} button:has-text('Қидириш')"
FISH_INPUT = "xpath=//label[contains(normalize-space(.),'Ф.И.Ш')]/following-sibling::input[1]"

HOUSE_DROPDOWN = "p-dropdown[formcontrolname='id']"
FAMILY_TYPE_DROPDOWN = "p-dropdown[formcontrolname='type']"
SAVE_BTN = "app-create-family-dialog-footer button"

CASE_B_DIALOG = "app-citizen-family-info-dialog"
CONTINUE_BTN = f"{CASE_B_DIALOG} button:has-text('Давом этиш')"

WARN_TOAST = ".p-toast-message-warn, p-toast .p-toast-message:has(exclamationtriangleicon)"
TOAST_DETAIL = ".p-toast-detail"
TOAST_CLOSE = ".p-toast-icon-close"

RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")


# ============================================================================
#  LOG
# ============================================================================

def log(msg):
    print(msg, flush=True)


def step(msg):
    """Batafsil qadam logi (vaqt belgisi bilan)."""
    if STEP_LOG:
        t = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        print(f"   . [{t}] {msg}", flush=True)


# ============================================================================
#  YORDAMCHI FUNKSIYALAR
# ============================================================================

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
    try:
        closers = page.locator(TOAST_CLOSE)
        cnt = await closers.count()
        if cnt:
            step(f"{cnt} ta eski toast yopilmoqda")
        for i in range(cnt):
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


async def select_from_dropdown(page, dropdown_selector, label, what):
    """PrimeNG p-dropdown'ni ochib, matn bo'yicha variant tanlaydi."""
    step(f"{what}: dropdown ochilmoqda")
    await page.locator(dropdown_selector).first.click()
    await page.locator(".p-dropdown-panel").last.wait_for(state="visible", timeout=5000)
    opt = page.get_by_role("option", name=label, exact=True)
    if await opt.count() == 0:
        opt = page.get_by_role("option", name=label)
    step(f"{what}: '{label}' tanlanmoqda")
    await opt.first.click()
    step(f"{what}: tanlandi")


async def open_modal(page):
    step("'+' (Оила қўшиш) tugmasi bosilmoqda")
    await page.locator(BTN_ADD).first.click()
    step("modal ochilishini kutyapman (ЖШШИР input ko'rinishi)")
    await page.locator(JSHSHIR_INPUT).first.wait_for(state="visible", timeout=10000)
    await page.wait_for_timeout(300)
    step("modal ochildi")
    if SELECT_DOC_TYPE:
        try:
            await select_from_dropdown(page, DOC_TYPE_DROPDOWN, DOC_TYPE_VALUE, "Ҳужжат тури")
        except Exception as e:
            step(f"Ҳужжат тури tanlanmadi (ehtimol allaqachon tanlangan): {e}")


async def close_all_dialogs(page):
    step("ochiq dialoglar yopilmoqda")
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


async def pick_random_house(page):
    step("Хонадон: dropdown ochilmoqda")
    await page.locator(HOUSE_DROPDOWN).first.click()
    await page.locator(".p-dropdown-panel").last.wait_for(state="visible", timeout=5000)
    options = page.get_by_role("option")
    try:
        await options.first.wait_for(state="visible", timeout=4000)
    except PWTimeout:
        step("Хонадон: ro'yxat bo'sh")
        await page.keyboard.press("Escape")
        return False
    count = await options.count()
    if count == 0:
        step("Хонадон: 0 ta variant")
        await page.keyboard.press("Escape")
        return False
    idx = random.randint(0, count - 1)
    step(f"Хонадон: {count} ta variant, #{idx} tanlanmoqda")
    await options.nth(idx).click()
    step("Хонадон: tanlandi")
    return True


async def wait_search_result(page):
    step("qidiruv natijasi kutilmoqda...")
    elapsed = 0
    while elapsed < SEARCH_TIMEOUT_MS:
        if await is_visible(page, CASE_B_DIALOG):
            return "dialog"
        if await is_visible(page, WARN_TOAST):
            return "warn"
        try:
            val = await page.locator(FISH_INPUT).first.input_value(timeout=300)
            if val and val.strip():
                step(f"Ф.И.Ш to'ldi: {val.strip()!r}")
                return "filled"
        except Exception:
            pass
        await page.wait_for_timeout(POLL_MS)
        elapsed += POLL_MS
    return "none"


async def wait_save_result(page):
    step("saqlash natijasi kutilmoqda...")
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

            log(f"[{row}] ЖШШИР: {jshshir}  ----------------------------------")
            try:
                step("toza holatga keltirish: toast/dialoglar tozalanmoqda")
                await dismiss_toasts(page)
                await close_all_dialogs(page)

                await open_modal(page)

                step(f"ЖШШИР kiritilmoqda: {jshshir}")
                await page.locator(JSHSHIR_INPUT).first.fill(jshshir)
                step("Қидириш tugmasi bosilmoqda")
                await page.locator(SEARCH_BTN).first.click()

                result = await wait_search_result(page)
                step(f"qidiruv natijasi = {result}")

                if result == "warn":
                    detail = await toast_detail_text(page)
                    log(f"  -> Case A: ogohlantirish ({detail!r}). Excel QIZIL.")
                    cell.fill = RED_FILL
                    stats["A"] += 1
                    await close_all_dialogs(page)

                elif result == "dialog":
                    log("  -> Case B: boshqa mahallada ro'yxatda. 'Давом этиш' bosilmoqda.")
                    try:
                        await page.locator(CONTINUE_BTN).first.click(timeout=3000)
                        step("'Давом этиш' bosildi")
                    except Exception as e:
                        step(f"'Давом этиш' bosilmadi: {e}")
                    await page.wait_for_timeout(500)
                    stats["B"] += 1
                    await close_all_dialogs(page)

                elif result == "filled":
                    log("  -> Case C: ma'lumot topildi.")
                    if not await pick_random_house(page):
                        log("     !! Xonadon ro'yxati bo'sh — o'tkazildi.")
                        stats["err"] += 1
                        await close_all_dialogs(page)
                        if (row - START_ROW + 1) % SAVE_EVERY == 0:
                            wb.save(EXCEL_FILE)
                        continue
                    await select_from_dropdown(page, FAMILY_TYPE_DROPDOWN,
                                               FAMILY_TYPE_LABEL, "Оила аъзо тури")
                    await dismiss_toasts(page)
                    step("Сақлаш tugmasi bosilmoqda")
                    await page.locator(SAVE_BTN).first.click()

                    save_res = await wait_save_result(page)
                    step(f"saqlash natijasi = {save_res}")
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
                        log("  -> ?: saqlash natijasi noma'lum.")
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
