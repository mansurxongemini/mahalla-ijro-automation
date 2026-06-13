"""
============================================================================
 ERP AVTOMATLASHTIRISH  —  mahalla.ijro.uz "Оила қўшиш"
============================================================================
O'quvchilarni ПИНФЛ (ЖШШИР) bo'yicha avtomatik ro'yxatga olish.

SOZLAMALAR: config.json faylida (Bloknotda ochib o'zgartiring).
Agar config.json yo'q bo'lsa — birinchi ishga tushirishda o'zi yaratiladi.

ISHGA TUSHIRISH:
    python test.py (yoki MahallaBot.exe)

QANDAY ISHLAYDI (har bir ЖШШИР uchun):
    1. "+" bosilib modal ochiladi, "Ҳужжат тури = ЖШШИР ва туғилган сана" tanlanadi.
    2. ЖШШИР kiritilib "Қидириш" bosiladi, SERVER JAVOBI kutiladi (tez).
    3. Natija:
        A) Ogohlantirish chiqsa     -> Excel QIZIL, keyingisiga.
        B) Dialog chiqsa            -> "Давом этиш" bosiladi, keyingisiga.
        C) Hech narsa chiqmasa      -> natija topilgan: Хонадон random +
           Оила аъзо тури "Бошқа" + Сақлаш:
              C1) modal yopiladi    -> Excel YASHIL (muvaffaqiyat)
              C2) sariq toast       -> Excel SARIQ (dublikat)
"""

import asyncio
import datetime
import json
import os
import random
import sys

import openpyxl
from openpyxl.styles import PatternFill
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

# ============================================================================
#  CONFIG.JSON DAN O'QISH
# ============================================================================

CONFIG_FILE = "config.json"

DEFAULTS = {
    "EXCEL_FILE": "1.xlsx",
    "JSHSHIR_COLUMN": 6,
    "START_ROW": 3,
    "END_ROW": None,
    "conJsh": True,
    "CACHE_FILE": "jshshir_cache.txt",
    "SKIP_PROCESSED": True,
    "TARGET_URL": "https://mahalla.ijro.uz/dashboard/list/family?region_id=00s0eed0000region000008&district_id=00s0eed0000region000049&mahalla_id=66016a6ae237a52f91961c45&returnPath=%2Fdashboard%2Fassistance&page=1&limit=20&offset=0",
    "HEADLESS": False,
    "SLOW_MO": 0,
    "SELECT_DOC_TYPE": True,
    "DOC_TYPE_MATCH": "ЖШШИР",
    "FAMILY_TYPE_LABEL": "Бошқа",
    "SERVER_WAIT_MS": 12000,
    "SETTLE_MS": 300,
    "UI_CHECK_MS": 1500,
    "SAVE_EVERY": 1,
    "POLL_MS": 200,
    "SEARCH_API_HINT": "",
    "STEP_LOG": True,
    "DEBUG_SCREENSHOTS": True,
}


def load_config():
    """config.json dan sozlamalarni o'qiydi. Yo'q bo'lsa DEFAULT + fayl yaratiladi."""
    cfg = dict(DEFAULTS)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, encoding="utf-8") as f:
                user = json.load(f)
            for k, v in user.items():
                if k.startswith("_"):
                    continue
                if k in cfg:
                    cfg[k] = v
        except Exception as e:
            print(f"!! config.json o'qilmadi: {e}. Default sozlamalar ishlatiladi.")
    else:
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                out = {
                    "_IZOH_": "Bu faylni Bloknotda ochib sozlamalarni o'zgartirishingiz mumkin.",
                    "_IZOH2_": "O'zgartirgandan keyin saqlang va dasturni qayta ishga tushiring.",
                }
                out.update(DEFAULTS)
                json.dump(out, f, ensure_ascii=False, indent=4)
            print(f"config.json yaratildi. Sozlamalarni o'zgartirish uchun uni Bloknotda oching.")
        except Exception:
            pass
    return cfg


CFG = load_config()

EXCEL_FILE = CFG["EXCEL_FILE"]
JSHSHIR_COLUMN = CFG["JSHSHIR_COLUMN"]
START_ROW = CFG["START_ROW"]
END_ROW = CFG["END_ROW"]
conJsh = CFG["conJsh"]
CACHE_FILE = CFG["CACHE_FILE"]
SKIP_PROCESSED = CFG["SKIP_PROCESSED"]
TARGET_URL = CFG["TARGET_URL"]
HEADLESS = CFG["HEADLESS"]
SLOW_MO = CFG["SLOW_MO"]
SELECT_DOC_TYPE = CFG["SELECT_DOC_TYPE"]
DOC_TYPE_MATCH = CFG["DOC_TYPE_MATCH"]
FAMILY_TYPE_LABEL = CFG["FAMILY_TYPE_LABEL"]
SERVER_WAIT_MS = CFG["SERVER_WAIT_MS"]
SETTLE_MS = CFG["SETTLE_MS"]
UI_CHECK_MS = CFG["UI_CHECK_MS"]
SAVE_EVERY = CFG["SAVE_EVERY"]
POLL_MS = CFG["POLL_MS"]
SEARCH_API_HINT = CFG["SEARCH_API_HINT"]
STEP_LOG = CFG["STEP_LOG"]
DEBUG_SCREENSHOTS = CFG["DEBUG_SCREENSHOTS"]

PROFILE_DIR = ".pw_profile"
_base, _ext = os.path.splitext(EXCEL_FILE)
OUTPUT_FILE = f"{_base}_natija{_ext}"

# ============================================================================
#  SELEKTORLAR  (sayt o'zgarmasa, tegmasangiz ham bo'ladi)
# ============================================================================

DIALOG = "app-create-family-dialog"
ANY_DIALOG_CLOSE = "div.p-dialog button[aria-label='Close']"
BTN_ADD = "button.p-button-icon-only.rounded-full:has(.pi-plus)"

DOC_TYPE_DROPDOWN = "xpath=//label[contains(normalize-space(.),'Ҳужжат тури')]/following-sibling::p-dropdown[1]"
JSHSHIR_INPUT = "xpath=//label[contains(normalize-space(.),'ЖШШИР')]/following-sibling::input[1]"
SEARCH_BTN = f"{DIALOG} button:has-text('Қидириш')"
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
    if STEP_LOG:
        t = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
        print(f"   . [{t}] {msg}", flush=True)


# ============================================================================
#  LOKAL KESH (conJsh)
# ============================================================================

def load_cache():
    s = set()
    if conJsh and os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, encoding="utf-8") as f:
                for line in f:
                    tok = line.strip().split("\t")[0].strip()
                    if tok.isdigit():
                        s.add(tok)
        except Exception as e:
            log(f"!! Kesh o'qilmadi: {e}")
    return s


def cache_add(jshshir, result):
    if not conJsh:
        return
    try:
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(CACHE_FILE, "a", encoding="utf-8") as f:
            f.write(f"{jshshir}\t{result}\t{ts}\n")
    except Exception as e:
        log(f"   !! Keshga yozilmadi: {e}")


def safe_save(wb):
    try:
        wb.save(OUTPUT_FILE)
        return True
    except PermissionError:
        log(f"   !! '{OUTPUT_FILE}' band (Excelda ochiqmi?)")
        return False
    except Exception as e:
        log(f"   !! Saqlash xatosi: {e}")
        return False


def already_processed(cell):
    fill = cell.fill
    if fill is None or fill.fill_type != "solid":
        return False
    try:
        color = str(fill.start_color.rgb).upper()
    except Exception:
        return False
    return any(c in color for c in ("FFC7CE", "C6EFCE", "FFEB9C"))


# ============================================================================
#  UI YORDAMCHILARI
# ============================================================================

async def is_visible(page, selector):
    try:
        return await page.locator(selector).first.is_visible()
    except Exception:
        return False


async def dismiss_toasts(page):
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


async def readonly_values(page):
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


def _resp_pred(resp):
    """Server javobini aniqlash."""
    try:
        if SEARCH_API_HINT:
            return SEARCH_API_HINT in resp.url
        return resp.request.resource_type in ("xhr", "fetch")
    except Exception:
        return False


async def click_and_wait_server(page, click_selector, what):
    """Tugmani bosib server javobini (XHR) kutadi — natija tayyor bo'lguncha."""
    step(f"{what} bosilmoqda + server javobi kutilmoqda")
    try:
        async with page.expect_response(_resp_pred, timeout=SERVER_WAIT_MS) as ri:
            await page.locator(click_selector).first.click()
        resp = await ri.value
        step(f"server javobi: {resp.status} {resp.url[:80]}")
    except PWTimeout:
        step(f"{what}: server javobi kutilmadi (timeout)")
    await page.wait_for_timeout(SETTLE_MS)


async def select_document_type(page):
    step("Ҳужжат тури: dropdown ochilmoqda")
    await page.locator(DOC_TYPE_DROPDOWN).first.click()
    await page.locator(".p-dropdown-panel").last.wait_for(state="visible", timeout=5000)
    all_opts = page.locator(".p-dropdown-panel [role='option']")
    texts = []
    for i in range(await all_opts.count()):
        try:
            texts.append((await all_opts.nth(i).inner_text()).strip())
        except Exception:
            texts.append("?")
    step(f"Ҳужжат тури opsiyalari: {texts}")
    target = page.locator(".p-dropdown-panel [role='option']", has_text=DOC_TYPE_MATCH)
    if await target.count() == 0:
        await page.keyboard.press("Escape")
        raise RuntimeError(f"'{DOC_TYPE_MATCH}' topilmadi. Mavjud: {texts}")
    chosen = (await target.first.inner_text()).strip()
    step(f"Ҳужжат тури: '{chosen}' tanlanmoqda")
    await target.first.click()
    step("Ҳужжат тури: tanlandi")


async def open_modal(page):
    step("'+' tugmasi bosilmoqda")
    await page.locator(BTN_ADD).first.click()
    await page.locator(DIALOG).first.wait_for(state="visible", timeout=10000)
    step("modal ochildi")
    if SELECT_DOC_TYPE:
        await select_document_type(page)
    await page.locator(JSHSHIR_INPUT).first.wait_for(state="visible", timeout=8000)


async def close_all_dialogs(page):
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
        await page.wait_for_timeout(150)


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
        await page.keyboard.press("Escape")
        return False
    idx = random.randint(0, count - 1)
    step(f"Хонадон: {count} ta variant, #{idx} tanlanmoqda")
    await options.nth(idx).click()
    return True


async def pick_family_type(page):
    step("Оила аъзо тури: dropdown ochilmoqda")
    await page.locator(FAMILY_TYPE_DROPDOWN).first.click()
    await page.locator(".p-dropdown-panel").last.wait_for(state="visible", timeout=5000)
    opt = page.get_by_role("option", name=FAMILY_TYPE_LABEL, exact=True)
    if await opt.count() == 0:
        opt = page.get_by_role("option", name=FAMILY_TYPE_LABEL)
    step(f"Оила аъзо тури: '{FAMILY_TYPE_LABEL}' tanlanmoqda")
    await opt.first.click()


async def classify_search(page):
    """Server javobidan keyin: dialog / warn / filled (Case C)."""
    elapsed = 0
    while elapsed < UI_CHECK_MS:
        if await is_visible(page, CASE_B_DIALOG):
            return "dialog"
        if await is_visible(page, WARN_TOAST):
            return "warn"
        if any(await readonly_values(page)):
            return "filled"
        await page.wait_for_timeout(POLL_MS)
        elapsed += POLL_MS
    # Ogohlantirish/dialog chiqmadi -> Case C
    return "filled"


async def classify_save(page):
    """Saqlash javobidan keyin: warn (C2) / success (C1)."""
    elapsed = 0
    while elapsed < UI_CHECK_MS:
        if await is_visible(page, WARN_TOAST):
            return "warn"
        if not await is_visible(page, DIALOG):
            return "success"
        await page.wait_for_timeout(POLL_MS)
        elapsed += POLL_MS
    return "success" if not await is_visible(page, DIALOG) else "warn"


# ============================================================================
#  ASOSIY OQIM
# ============================================================================

async def main():
    load_path = OUTPUT_FILE if os.path.exists(OUTPUT_FILE) else EXCEL_FILE
    log(f"Yuklanmoqda: {load_path}  ->  natija: {OUTPUT_FILE}")
    wb = openpyxl.load_workbook(load_path)
    sheet = wb.active
    last_row = END_ROW or sheet.max_row

    cache_set = load_cache()
    if conJsh:
        log(f"Lokal kesh: {len(cache_set)} ta ЖШШИР (takrorlar o'tkaziladi)")

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

            jshshir = str(raw).strip()
            if not jshshir.isdigit() or len(jshshir) != 14:
                log(f"[{row}] NOTO'G'RI ЖШШИР: {jshshir!r}")
                stats["err"] += 1
                continue

            if conJsh and jshshir in cache_set:
                stats["skip"] += 1
                continue
            if SKIP_PROCESSED and already_processed(cell):
                stats["skip"] += 1
                continue

            log(f"[{row}] ЖШШИР: {jshshir}  ----------------------------------")
            outcome = None
            try:
                await dismiss_toasts(page)
                await close_all_dialogs(page)
                await open_modal(page)

                step(f"ЖШШИР kiritilmoqda: {jshshir}")
                await page.locator(JSHSHIR_INPUT).first.fill(jshshir)
                await click_and_wait_server(page, SEARCH_BTN, "Қидириш")

                result = await classify_search(page)
                step(f"qidiruv natijasi = {result}")

                if result == "warn":
                    detail = await toast_detail_text(page)
                    log(f"  -> Case A: ogohlantirish ({detail!r}). Excel QIZIL.")
                    cell.fill = RED_FILL
                    stats["A"] += 1
                    outcome = "A"
                    await close_all_dialogs(page)

                elif result == "dialog":
                    log("  -> Case B: 'Давом этиш' bosilmoqda.")
                    try:
                        await page.locator(CONTINUE_BTN).first.click(timeout=3000)
                    except Exception:
                        pass
                    stats["B"] += 1
                    outcome = "B"
                    await close_all_dialogs(page)

                else:  # filled -> Case C
                    log("  -> Case C: Хонадон + Оила аъзо тури + Сақлаш")
                    if not await pick_random_house(page):
                        log("     !! Xonadon ro'yxati bo'sh")
                        stats["err"] += 1
                        await close_all_dialogs(page)
                        continue
                    await pick_family_type(page)
                    await dismiss_toasts(page)
                    await click_and_wait_server(page, SAVE_BTN, "Сақлаш")

                    save_res = await classify_save(page)
                    step(f"saqlash natijasi = {save_res}")
                    if save_res == "success":
                        log("  -> C1: MUVAFFAQIYATLI. Excel YASHIL.")
                        cell.fill = GREEN_FILL
                        stats["C1"] += 1
                        outcome = "C1"
                    else:
                        detail = await toast_detail_text(page)
                        log(f"  -> C2: xonadon dublikat ({detail!r}). Excel SARIQ.")
                        cell.fill = YELLOW_FILL
                        stats["C2"] += 1
                        outcome = "C2"
                        await close_all_dialogs(page)

            except Exception as e:
                log(f"  !! XATO [{row}]: {e}")
                stats["err"] += 1
                outcome = "ERR"
                if DEBUG_SCREENSHOTS:
                    try:
                        await page.screenshot(path=f"debug_row_{row}.png")
                    except Exception:
                        pass
                await close_all_dialogs(page)

            if outcome and outcome != "ERR":
                cache_set.add(jshshir)
                cache_add(jshshir, outcome)

            if (row - START_ROW + 1) % SAVE_EVERY == 0:
                safe_save(wb)

        while not safe_save(wb):
            ans = await asyncio.to_thread(
                input, f">>> '{OUTPUT_FILE}' band. Faylni yopib ENTER bosing (yoki q): "
            )
            if ans.strip().lower() == "q":
                break

        log(f"\nNatija: {OUTPUT_FILE}")
        log("==== YAKUNLANDI ====")
        log(f"A  (qizil / boshqa MFY)      : {stats['A']}")
        log(f"B  (dialog / davom etish)    : {stats['B']}")
        log(f"C1 (yashil / muvaffaqiyatli) : {stats['C1']}")
        log(f"C2 (sariq / xonadon dublikat): {stats['C2']}")
        log(f"O'tkazib yuborilgan          : {stats['skip']}")
        log(f"Xato                         : {stats['err']}")
        await context.close()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nFoydalanuvchi to'xtatdi.")
    except Exception as e:
        print(f"\n!!! KUTILMAGAN XATO: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("\n")
        input(">>> Tugatish uchun ENTER bosing... ")
