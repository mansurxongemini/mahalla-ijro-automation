"""
ERP avtomatlashtirish — mahalla.ijro.uz "Оила қўшиш" oynasi orqali o'quvchilarni
ПИНФЛ (ЖШШИР) bo'yicha ro'yxatga olish.

Sayt: Angular + PrimeNG.
  - Element id'lari (pn_id_xxx) HAR SAFAR o'zgaradi -> ularga tayanmaymiz.
  - Ogohlantirish (sariq) toast'i <exclamationtriangleicon> ikonkasi bilan keladi.
  - HOLAT IZOLYATSIYASI: har bir o'quvchidan keyin modal YOPILIB, keyingisi uchun
    QAYTADAN OCHILADI.

BU VERSIYADA:
  - Modal ochilganda "Ҳужжат тури" dan ЖШШИР-li opsiya ("ЖШШИР ва туғилган сана") tanlanadi.
  - Case C readonly natija inputlari (Ф.И.Ш / Туғилган сана) qiymati bo'yicha aniqlanadi.
  - Natija ALOHIDA faylga yoziladi ("..._natija.xlsx") -> manba fayl Excelda ochiq qolsa
    ham PermissionError bo'lmaydi.
  - BATAFSIL QADAM LOGI (step) va debug screenshot mavjud.

ISHGA TUSHIRISH:
    pip install playwright openpyxl
    playwright install chromium
    python test.py
"""

import asyncio
import datetime
import os
import random
import sys

import openpyxl
from openpyxl.styles import PatternFill
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

# ============================================================================
#  SOZLAMALAR
# ============================================================================

EXCEL_FILE = "15.08.2025 ERP.xlsx"   # MANBA fayl (o'qish uchun; ochiq qolsa ham bo'ladi)
# Natija (rangli) alohida faylga yoziladi -> manba fayl Excelda ochiq bo'lsa ham xato bermaydi.
# Avtomatik nom: "<manba>_natija.xlsx". Xohlasangiz o'zingiz belgilang.
_base, _ext = os.path.splitext(EXCEL_FILE)
OUTPUT_FILE = f"{_base}_natija{_ext}"
JSHSHIR_COLUMN = 6
START_ROW = 3
END_ROW = None               # None = oxirigacha. Sinov uchun masalan 10.

SKIP_PROCESSED = True
DEBUG_SCREENSHOTS = True
STEP_LOG = True              # batafsil qadam-baqadam log

SELECT_DOC_TYPE = True
# Ҳужжат тури dropdown'dan TANLANADIGAN opsiyada shu so'z bo'lishi kerak.
# "ЖШШИР ва туғилган сана" opsiyasini tanlash uchun "ЖШШИР" yetarli.
DOC_TYPE_MATCH = "ЖШШИР"

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

SEARCH_TIMEOUT_MS = 9000     # ogohlantirish/dialog kutishning MAKS vaqti (shundan keyin Case C)
SAVE_TIMEOUT_MS = 9000
CASE_C_CONFIRM_MS = 1000     # readonly to'lgach, warning kechikmaganini tasdiqlash oynasi
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
# Natija inputlari (Ф.И.Ш, Туғилган санаси) — readonly. To'lganda Case C demakdir.
READONLY_INPUTS = f"{DIALOG} input[readonly]"

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

def safe_save(wb):
    """Natijani OUTPUT_FILE'ga saqlaydi. Fayl band bo'lsa (Excelda ochiq) qulamaydi."""
    try:
        wb.save(OUTPUT_FILE)
        return True
    except PermissionError:
        log(f"   !! '{OUTPUT_FILE}' band (Excelda ochiqmi?) — keyinroq qayta urinaman.")
        return False
    except Exception as e:
        log(f"   !! Saqlash xatosi: {e}")
        return False


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


async def select_document_type(page):
    """Ҳужжат тури dropdown'dan DOC_TYPE_MATCH so'zli opsiyani tanlaydi.
    Topilmasa, mavjud opsiyalarni logga chiqaradi (debug uchun)."""
    step("Ҳужжат тури: dropdown ochilmoqda")
    await page.locator(DOC_TYPE_DROPDOWN).first.click()
    await page.locator(".p-dropdown-panel").last.wait_for(state="visible", timeout=5000)
    all_opts = page.locator(".p-dropdown-panel [role='option']")
    n = await all_opts.count()
    texts = []
    for i in range(n):
        try:
            texts.append((await all_opts.nth(i).inner_text()).strip())
        except Exception:
            texts.append("?")
    step(f"Ҳужжат тури opsiyalari: {texts}")
    target = page.locator(".p-dropdown-panel [role='option']", has_text=DOC_TYPE_MATCH)
    if await target.count() == 0:
        await page.keyboard.press("Escape")
        raise RuntimeError(f"'{DOC_TYPE_MATCH}' so'zli opsiya topilmadi. Mavjud: {texts}")
    chosen = (await target.first.inner_text()).strip()
    step(f"Ҳужжат тури: '{chosen}' tanlanmoqda")
    await target.first.click()
    step("Ҳужжат тури: tanlandi")


async def open_modal(page):
    step("'+' (Оила қўшиш) tugmasi bosilmoqda")
    await page.locator(BTN_ADD).first.click()
    step("modal ochilishini kutyapman")
    await page.locator(DIALOG).first.wait_for(state="visible", timeout=10000)
    await page.wait_for_timeout(300)
    step("modal ochildi")
    if SELECT_DOC_TYPE:
        await select_document_type(page)
    # Hujjat turi tanlangach ЖШШИР input chiqishini kutamiz
    step("ЖШШИР input ko'rinishini kutyapman")
    await page.locator(JSHSHIR_INPUT).first.wait_for(state="visible", timeout=8000)


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


async def readonly_values(page):
    """Modaldagi readonly inputlar (Ф.И.Ш, Туғилган санаси) qiymatlarini qaytaradi."""
    vals = []
    try:
        ro = page.locator(READONLY_INPUTS)
        for i in range(await ro.count()):
            try:
                v = await ro.nth(i).evaluate("el => el.value")
            except Exception:
                v = None
            vals.append((v or "").strip())
    except Exception:
        pass
    return vals


async def wait_search_result(page):
    """
    Qidiruv natijasi MANTIG'I:
      - Dialog chiqsa            -> 'dialog' (Case B)
      - Ogohlantirish toast chiqsa -> 'warn'   (Case A)
      - Ikkalasi HAM chiqmasa    -> 'filled' (Case C) — natija topilgani KAFOLATLANADI.
    Ya'ni "none" yo'q: ogohlantirish/dialog bo'lmasa, saqlash bosqichiga o'tamiz.
    Tezlashtirish uchun: readonly natija inputi to'lsa, qisqa tasdiqdan keyin darrov 'filled'.
    """
    step("qidiruv natijasi kutilmoqda (ogohlantirish/dialog tekshirilmoqda)...")
    elapsed = 0
    data_seen_at = None
    while elapsed < SEARCH_TIMEOUT_MS:
        if await is_visible(page, CASE_B_DIALOG):
            return "dialog"
        if await is_visible(page, WARN_TOAST):
            return "warn"
        # Ma'lumot to'ldimi? (tezkor yo'l) — lekin warning kechikishi mumkin, qisqa tasdiqlaymiz
        vals = await readonly_values(page)
        if any(vals):
            if data_seen_at is None:
                data_seen_at = elapsed
                step(f"natija to'ldi (readonly): {vals} — tasdiqlanmoqda")
            elif elapsed - data_seen_at >= CASE_C_CONFIRM_MS:
                return "filled"
        await page.wait_for_timeout(POLL_MS)
        elapsed += POLL_MS
    # Belgilangan vaqt o'tdi, ogohlantirish/dialog chiqmadi -> Case C (natija bor)
    step("ogohlantirish/dialog chiqmadi -> Case C deb qabul qilinadi")
    return "filled"


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
    load_path = OUTPUT_FILE if os.path.exists(OUTPUT_FILE) else EXCEL_FILE
    log(f"Yuklanmoqda: {load_path}  ->  natija: {OUTPUT_FILE}")
    wb = openpyxl.load_workbook(load_path)
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
                            safe_save(wb)
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
                    vals = await readonly_values(page)
                    log(f"  -> ?: qidiruv natijasi aniqlanmadi. Toast={detail!r} | readonly={vals}")
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
                safe_save(wb)

        # Yakuniy saqlash — band bo'lsa foydalanuvchidan faylni yopishni so'raymiz
        while not safe_save(wb):
            ans = await asyncio.to_thread(
                input, f">>> '{OUTPUT_FILE}' band. Faylni yopib ENTER bosing (yoki q+ENTER chiqish): "
            )
            if ans.strip().lower() == "q":
                break
        log(f"\nNatija saqlandi: {OUTPUT_FILE}")
        log("==== YAKUNLANDI ====")
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
