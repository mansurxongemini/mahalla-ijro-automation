"""
============================================================================
 ERP AVTOMATLASHTIRISH  —  mahalla.ijro.uz "Оила қўшиш"
============================================================================
O'quvchilarni ПИНФЛ (ЖШШИР) bo'yicha avtomatik ro'yxatga olish.

ISHGA TUSHIRISH (bir marta):
    pip install playwright openpyxl
    playwright install chromium

ISHLATISH:
    python test.py
    -> Brauzer ochiladi. QO'LDA login qilib, kerakli sahifani oching, ENTER bosing.
       (Login .pw_profile papkasida saqlanadi — keyingi safar login shart emas.)

QANDAY ISHLAYDI (har bir ЖШШИР uchun):
    1. "+" bosilib modal ochiladi, "Ҳужжат тури = ЖШШИР ва туғилган сана" tanlanadi.
    2. ЖШШИР kiritilib "Қидириш" bosiladi, SERVER JAVOBI kutiladi (tez).
    3. Natija:
        - Dialog chiqsa            -> "Давом этиш" bosilib o'tkaziladi (Case B).
        - Ogohlantirish chiqsa     -> Excel QIZIL, o'tkaziladi (Case A).
        - Hech narsa chiqmasa      -> natija topilgan (Case C): Хонадон random +
                                       Оила аъзо тури "Бошқа" + Сақлаш:
              * muvaffaqiyat (modal yopiladi) -> Excel YASHIL (C1)
              * "хонадон ... киритилган" sariq -> Excel SARIQ (C2)

Hamma o'zgartiriladigan sozlamalar PASTDA, "SOZLAMALAR" bo'limida.
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
#  SOZLAMALAR  (asosan SHU YERNI o'zgartirasiz)
# ============================================================================

# --- Excel fayl ---
EXCEL_FILE = "1.xlsx"          # MANBA Excel fayl nomi (o'qish uchun).
                               #   Har xil fayl ishlatsangiz shu nomni o'zgartiring.
JSHSHIR_COLUMN = 6             # ЖШШИР (ПИНФЛ) qaysi ustunda. A=1, B=2 ... F=6.
START_ROW = 3                  # Ma'lumot nechanchi qatordan boshlanadi.
END_ROW = None                 # Qaysi qatorgacha. None = oxirigacha. Sinov: masalan 15.
# Natija (ranglar) ALOHIDA faylga yoziladi -> manba faylni Excelda ochiq qoldirsangiz ham
# xato bo'lmaydi. Nomi avtomatik: "<manba>_natija.xlsx".
_base, _ext = os.path.splitext(EXCEL_FILE)
OUTPUT_FILE = f"{_base}_natija{_ext}"

# --- Lokal ЖШШИР keshi (takrorlanishni o'tkazib yuborish) ---
conJsh = True                  # True  -> avval kiritilgan ЖШШИРlarni (cache faylidan) o'tkazib yuboradi.
                               # False -> keshga qaramay, hammasini boshidan tekshiradi.
CACHE_FILE = "jshshir_cache.txt"   # Kiritilgan ЖШШИРlar shu faylga yozib boriladi.

# --- Excel rangi bo'yicha resume (shu fayl ichida) ---
SKIP_PROCESSED = True          # True -> natija faylida allaqachon bo'yalgan qatorlarni o'tkazadi.

# --- Sayt ---
TARGET_URL = ("https://mahalla.ijro.uz/dashboard/list/family"
              "?region_id=00s0eed0000region000008"
              "&district_id=00s0eed0000region000049"
              "&mahalla_id=66016a6ae237a52f91961c45"
              "&returnPath=%2Fdashboard%2Fassistance&page=1&limit=20&offset=0")

PROFILE_DIR = ".pw_profile"    # Brauzer profili (login shu yerda saqlanadi).
HEADLESS = False               # True -> brauzer ko'rinmaydi (tezroq, lekin kuzata olmaysiz).
SLOW_MO = 0                    # Har amaldan keyin ms kechikish (kuzatish uchun masalan 100).

# --- Forma qiymatlari ---
SELECT_DOC_TYPE = True         # Ҳужжат тури dropdown'ini tanlash kerakmi.
DOC_TYPE_MATCH = "ЖШШИР"       # Ҳужжат тури dropdown'da SHU so'zli opsiya tanlanadi
                               #   (masalan "ЖШШИР ва туғилган санаси").
FAMILY_TYPE_LABEL = "Бошқа"    # "Оила аъзо тури" dan tanlanadigan qiymat.

# --- Tezlik / kutish (ms) ---
# Qidirish/saqlashda dastur SERVER JAVOBINI (XHR) kutadi — bu natija tayyor bo'lgan lahza.
SERVER_WAIT_MS = 12000         # Server javobini kutishning MAKS vaqti.
SETTLE_MS = 300                # Javobdan keyin UI chizilishi uchun qisqa pauza.
UI_CHECK_MS = 1500             # Javobdan so'ng ogohlantirish/dialog chiqishini kuzatish oynasi.
                               #   (shundan keyin "hech narsa yo'q" -> Case C deb hisoblaydi)
SAVE_EVERY = 1                 # Har necha qatorda natija faylini saqlash.
POLL_MS = 200                  # UI tekshiruv qadami.
# Agar qidirish API manzilini bilsangiz (debug logda ko'rinadi), aniqlik uchun shu yerga
# uning bir qismini yozing (masalan "/citizen" yoki "/search"). Bo'sh bo'lsa har qanday XHR kutiladi.
SEARCH_API_HINT = ""

# --- Log / debug ---
STEP_LOG = True                # Har bir mayda amalni vaqt belgisi bilan loglash.
DEBUG_SCREENSHOTS = True       # Noaniq holatda debug_row_N.png saqlash.

# ============================================================================
#  SELEKTORLAR  (sayt o'zgarmasa, tegmasangiz ham bo'ladi)
# ============================================================================

DIALOG = "app-create-family-dialog"
ANY_DIALOG_CLOSE = "div.p-dialog button[aria-label='Close']"
BTN_ADD = "button.p-button-icon-only.rounded-full:has(.pi-plus)"

DOC_TYPE_DROPDOWN = "xpath=//label[contains(normalize-space(.),'Ҳужжат тури')]/following-sibling::p-dropdown[1]"
JSHSHIR_INPUT = "xpath=//label[contains(normalize-space(.),'ЖШШИР')]/following-sibling::input[1]"
SEARCH_BTN = f"{DIALOG} button:has-text('Қидириш')"
READONLY_INPUTS = f"{DIALOG} input[readonly]"   # Ф.И.Ш / Туғилган сана (to'lsa Case C)

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
        log(f"   !! '{OUTPUT_FILE}' band (Excelda ochiqmi?) — keyinroq qayta urinaman.")
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
    """Qaysi network javobini 'natija' deb hisoblaymiz."""
    try:
        if SEARCH_API_HINT:
            return SEARCH_API_HINT in resp.url
        return resp.request.resource_type in ("xhr", "fetch")
    except Exception:
        return False


async def click_and_wait_server(page, click_selector, what):
    """Tugmani bosadi va server javobini (XHR) kutadi — natija tayyor bo'lguncha."""
    step(f"{what} bosilmoqda + server javobi kutilmoqda")
    try:
        async with page.expect_response(_resp_pred, timeout=SERVER_WAIT_MS) as ri:
            await page.locator(click_selector).first.click()
        resp = await ri.value
        step(f"server javobi: {resp.status}  {resp.url}")
    except PWTimeout:
        step(f"{what}: server javobi kutilmadi (timeout) — UI bo'yicha davom etamiz")
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
        raise RuntimeError(f"'{DOC_TYPE_MATCH}' so'zli opsiya topilmadi. Mavjud: {texts}")
    chosen = (await target.first.inner_text()).strip()
    step(f"Ҳужжат тури: '{chosen}' tanlanmoqda")
    await target.first.click()
    step("Ҳужжат тури: tanlandi")


async def open_modal(page):
    step("'+' (Оила қўшиш) tugmasi bosilmoqda")
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
    """Server javobidan keyin: 'dialog' / 'warn' / 'filled' (Case C)."""
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
    # Ogohlantirish ham, dialog ham chiqmadi -> natija topilgan (Case C)
    return "filled"


async def classify_save(page):
    """Saqlash javobidan keyin: 'warn' (C2) / 'success' (C1)."""
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

            # O'tkazib yuborish: lokal kesh (conJsh) yoki Excel rangi
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
                    log("  -> Case B: 'Давом этиш' bosilmoqda, o'tkaziladi.")
                    try:
                        await page.locator(CONTINUE_BTN).first.click(timeout=3000)
                    except Exception:
                        pass
                    stats["B"] += 1
                    outcome = "B"
                    await close_all_dialogs(page)

                else:  # filled -> Case C
                    log("  -> Case C: natija topildi. Хонадон + Оила аъзо тури + Сақлаш.")
                    if not await pick_random_house(page):
                        log("     !! Xonadon ro'yxati bo'sh — o'tkazildi.")
                        stats["err"] += 1
                        await close_all_dialogs(page)
                        if (row - START_ROW + 1) % SAVE_EVERY == 0:
                            safe_save(wb)
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

            # Lokalga yozib qo'yamiz (xato bo'lmasa) — keyingi safar o'tkaziladi
            if outcome and outcome != "ERR":
                cache_set.add(jshshir)
                cache_add(jshshir, outcome)

            if (row - START_ROW + 1) % SAVE_EVERY == 0:
                safe_save(wb)

        # Yakuniy saqlash (band bo'lsa foydalanuvchidan so'raymiz)
        while not safe_save(wb):
            ans = await asyncio.to_thread(
                input, f">>> '{OUTPUT_FILE}' band. Faylni yopib ENTER bosing (yoki q+ENTER): "
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
        log(f"Xato                         : {stats['err']}")
        await context.close()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        sys.exit("\nFoydalanuvchi to'xtatdi.")
