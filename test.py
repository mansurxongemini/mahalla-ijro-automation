"""
============================================================================
 ERP AVTOMATLASHTIRISH  —  mahalla.ijro.uz  (AHOLI / population)
============================================================================
O'quvchilarni ПИНФЛ (ЖШШИР) bo'yicha joriy mahallaga ko'chirish/qo'shish.

SOZLAMALAR: config.json faylida (Bloknotda ochib o'zgartiring).

OQIM (har bir ЖШШИР uchun):
  1. "+Қўшиш" bosiladi -> ЖШШИР kiritiladi -> "Қидириш" bosiladi (server javobi kutiladi).
  2. Natija:
     A) "Фуқаро ушбу ҳудудда топилмади" (xato, tashqi mahalla ko'rsatiladi)
        -> NOTFOUND faylga yoziladi:  A ustun = tashqi mahalla,  B+ = asl qatorning
           to'liq nusxasi.  So'ng "Бекор қилиш" bosilib keyingisiga o'tiladi.
     B) "...аллақачон рўйхатга олинган" modal -> "Давом этиш":
           - telefon raqami (random, ruxsat etilgan prefiks bilan)
           - Хонадон: random
           - Оила: random
           - Оила аъзо тури: "Бошқа"
           - "Сақлаш"  -> keyingisiga.
     BOSHQA holatlar -> "Бекор қилиш" -> keyingisiga.

ISHGA TUSHIRISH:  python test.py   (yoki MahallaBot.exe)
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
#  CONFIG.JSON
# ============================================================================

CONFIG_FILE = "config.json"

DEFAULTS = {
    "EXCEL_FILE": "1.xlsx",
    "JSHSHIR_COLUMN": 6,
    "START_ROW": 3,
    "END_ROW": None,
    "NOTFOUND_FILE": "topilmadi.xlsx",
    "conJsh": True,
    "CACHE_FILE": "jshshir_cache.txt",
    "SKIP_PROCESSED": True,
    "TARGET_URL": "https://mahalla.ijro.uz/dashboard/list/population?region_id=00s0eed0000region000008&district_id=00s0eed0000region000049&mahalla_id=66016a6ae237a52f91961c45&returnPath=%2Fdashboard%2Fassistance",
    "HEADLESS": False,
    "SLOW_MO": 0,
    "PHONE_PREFIXES": [20, 33, 50, 55, 77, 78, 71, 88, 80, 90, 91, 92, 94, 93, 95, 97, 99],
    "FAMILY_TYPE_LABEL": "Бошқа",
    "SERVER_WAIT_MS": 12000,
    "SETTLE_MS": 350,
    "UI_CHECK_MS": 5000,
    "SAVE_EVERY": 1,
    "POLL_MS": 200,
    "SEARCH_API_HINT": "",
    "STEP_LOG": True,
    "DEBUG_SCREENSHOTS": False,
}


def load_config():
    cfg = dict(DEFAULTS)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, encoding="utf-8") as f:
                user = json.load(f)
            for k, v in user.items():
                if not k.startswith("_") and k in cfg:
                    cfg[k] = v
        except Exception as e:
            print(f"!! config.json o'qilmadi: {e}. Default ishlatiladi.")
    else:
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                out = {"_IZOH_": "Bloknotda ochib sozlamalarni o'zgartiring, so'ng qayta ishga tushiring."}
                out.update(DEFAULTS)
                json.dump(out, f, ensure_ascii=False, indent=4)
        except Exception:
            pass
    return cfg


CFG = load_config()

EXCEL_FILE = CFG["EXCEL_FILE"]
JSHSHIR_COLUMN = CFG["JSHSHIR_COLUMN"]
START_ROW = CFG["START_ROW"]
END_ROW = CFG["END_ROW"]
NOTFOUND_FILE = CFG["NOTFOUND_FILE"]
conJsh = CFG["conJsh"]
CACHE_FILE = CFG["CACHE_FILE"]
SKIP_PROCESSED = CFG["SKIP_PROCESSED"]
TARGET_URL = CFG["TARGET_URL"]
HEADLESS = CFG["HEADLESS"]
SLOW_MO = CFG["SLOW_MO"]
PHONE_PREFIXES = CFG["PHONE_PREFIXES"]
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
#  SELEKTORLAR
# ============================================================================

# 1-qadam
BTN_ADD = "p-button[label='Қўшиш'] button, button:has(.pi-plus):has-text('Қўшиш')"

# 2-qadam
JSHSHIR_INPUT = "xpath=//label[contains(normalize-space(.),'ЖШШИР')]/following-sibling::input[1]"
# Qidirish tugmasi — ЖШШИР inputi bilan BIR formada bo'lgani (sahifa filtri emas).
# ЖШШИР label'ning eng yaqin ajdodi ichidagi "Қидириш" tugmasini topadi.
SEARCH_BTN = ("xpath=//label[contains(normalize-space(.),'ЖШШИР')]"
              "/ancestor::*[.//button[.//span[contains(normalize-space(.),'Қидириш')]]][1]"
              "//button[.//span[contains(normalize-space(.),'Қидириш')]]")
# Zaxira: umumiy (debug uchun)
SEARCH_BTN_ANY = "button[label='Қидириш'], button:has(.pi-search):has-text('Қидириш')"

# Case A — fuqaro topilmadi (tashqi mahalla MA'LUM)
CASE_A_MSG = "app-state-message[severity='error']"
# Xato/ko'chirib bo'lmaydigan holatlar (yakuniy, to'ldirilmaydi)
CASE_ERR = ("app-state-message[severity='error'], "
            "app-state-message:has-text('кўчириб бўлмайди'), "
            "app-state-message:has-text('мавжуд эмас')")
TERR_ROW = "app-state-message .terr-row"

# Case B — boshqa mahallada ro'yxatda (modal)
CASE_B_DIALOG = "app-citizen-family-info-dialog"
CONTINUE_BTN = f"{CASE_B_DIALOG} button:has-text('Давом этиш')"

# B dan keyingi forma
PHONE_INPUT = "input[formcontrolname='phone_number']"
HOUSE_DD = "div[formgroupname='house'] p-dropdown[formcontrolname='id']"
FAMILY_DD = "div[formgroupname='family'] p-dropdown[formcontrolname='id']"
MEMBER_TYPE_DD = "p-dropdown[formcontrolname='type']"
SAVE_BTN = "button:has-text('Сақлаш')"

# Holatni 0 ga qaytarish uchun (sahifa qayta yuklanadi — Бекор tugmasi ishlatilmaydi)
ANY_DIALOG_CLOSE = "div.p-dialog button[aria-label='Close']"

DROPDOWN_PANEL = ".p-dropdown-panel"

GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")


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
#  KESH / EXCEL
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


def safe_save(wb, path):
    try:
        wb.save(path)
        return True
    except PermissionError:
        log(f"   !! '{path}' band (Excelda ochiqmi?)")
        return False
    except Exception as e:
        log(f"   !! Saqlash xatosi ({path}): {e}")
        return False


def already_processed(cell):
    fill = cell.fill
    if fill is None or fill.fill_type != "solid":
        return False
    try:
        color = str(fill.start_color.rgb).upper()
    except Exception:
        return False
    return any(c in color for c in ("FFC7CE", "C6EFCE"))


def random_phone():
    """Ruxsat etilgan prefiksli 9 xonali raqam (prefiks 2 + 7 random)."""
    pref = random.choice(PHONE_PREFIXES)
    rest = random.randint(0, 9999999)
    return f"{pref:02d}{rest:07d}"


def next_empty_row(sheet):
    """Sheetdagi birinchi bo'sh qator (ketma-ket yozish uchun)."""
    r = sheet.max_row
    while r >= 1:
        maxc = sheet.max_column or 1
        if all((sheet.cell(row=r, column=c).value in (None, "")) for c in range(1, maxc + 1)):
            r -= 1
        else:
            break
    return r + 1


# ============================================================================
#  UI YORDAMCHILARI
# ============================================================================

async def is_visible(page, selector):
    try:
        return await page.locator(selector).first.is_visible()
    except Exception:
        return False


async def first_actionable(page, selector):
    """Selektorga mos KO'RINADIGAN va FAOL (enabled) birinchi elementni qaytaradi."""
    loc = page.locator(selector)
    try:
        n = await loc.count()
    except Exception:
        n = 0
    for i in range(n):
        b = loc.nth(i)
        try:
            if await b.is_visible() and await b.is_enabled():
                return b
        except Exception:
            pass
    return loc.first


async def _click_robust(page, selector, what):
    """Tugmani 3 usulda bosishga urinadi: oddiy -> force -> JavaScript."""
    btn = await first_actionable(page, selector)
    try:
        await btn.scroll_into_view_if_needed(timeout=3000)
    except Exception:
        pass
    attempts = (
        ("oddiy", lambda: btn.click(timeout=6000)),
        ("force", lambda: btn.click(timeout=4000, force=True)),
        ("js", lambda: btn.evaluate("el => el.click()")),
    )
    for how, fn in attempts:
        try:
            await fn()
            step(f"{what}: bosildi ({how})")
            return True
        except Exception as e:
            step(f"{what}: {how} click bo'lmadi: {str(e)[:70]}")
    return False


async def click_and_wait_server(page, click_selector, what):
    step(f"{what} bosilmoqda + server javobi kutilmoqda")
    # Debug: nechta 'Қидириш' tugmasi bor
    if what.startswith("Қидириш"):
        try:
            total = await page.locator(SEARCH_BTN_ANY).count()
            scoped = await page.locator(click_selector).count()
            step(f"'Қидириш' tugmalari: sahifada {total} ta, formada {scoped} ta")
        except Exception:
            pass
    def _pred(resp):
        try:
            if SEARCH_API_HINT:
                return SEARCH_API_HINT in resp.url
            return resp.request.resource_type in ("xhr", "fetch")
        except Exception:
            return False
    try:
        async with page.expect_response(_pred, timeout=SERVER_WAIT_MS):
            await _click_robust(page, click_selector, what)
        step(f"{what}: server javobi keldi")
    except PWTimeout:
        step(f"{what}: server javobi kutilmadi (timeout)")
    await page.wait_for_timeout(SETTLE_MS)


async def type_jshshir(page, jshshir):
    """ЖШШИРni harf-harf yozadi (Angular validatsiyasi uchun) va 'Қидириш' faollashuvini kutadi."""
    inp = page.locator(JSHSHIR_INPUT).first
    await inp.click()
    await inp.fill("")
    try:
        await inp.type(jshshir, delay=25)
    except Exception:
        await inp.fill(jshshir)
    await page.wait_for_timeout(300)
    for _ in range(15):
        btn = await first_actionable(page, SEARCH_BTN)
        try:
            if await btn.is_visible() and await btn.is_enabled():
                return True
        except Exception:
            pass
        await page.wait_for_timeout(200)
    step("OGOHLANTIRISH: 'Қидириш' tugmasi faollashmadi/topilmadi")
    return False


async def dd_options_ready(page, selector, what, total_ms=15000):
    """Dropdownni ochib, optionlar (serverdan) yuklanishini kutadi.
    Bo'sh bo'lsa yopib-ochib qayta urinadi. Tayyor bo'lsa ochiq qoldiradi."""
    waited = 0
    while waited < total_ms:
        if not await is_visible(page, DROPDOWN_PANEL):
            try:
                await page.locator(selector).first.click(timeout=4000)
            except Exception:
                pass
        try:
            await page.locator(DROPDOWN_PANEL).last.wait_for(state="visible", timeout=2000)
        except PWTimeout:
            await page.wait_for_timeout(500)
            waited += 500
            continue
        opts = page.get_by_role("option")
        try:
            cnt = await opts.count()
        except Exception:
            cnt = 0
        if cnt > 0:
            return True
        # optionlar hali yo'q -> yopib, server uchun kutib, qayta urinamiz
        try:
            await page.keyboard.press("Escape")
        except Exception:
            pass
        await page.wait_for_timeout(700)
        waited += 700
    step(f"{what}: optionlar yuklanmadi ({total_ms} ms)")
    return False


async def dd_pick_random(page, selector, what):
    """p-dropdown'ni ochib (optionlar yuklanishini kutib) random variant tanlaydi."""
    step(f"{what}: dropdown ochilmoqda (optionlar kutilmoqda)")
    if not await dd_options_ready(page, selector, what):
        return False
    opts = page.get_by_role("option")
    n = await opts.count()
    if n == 0:
        await page.keyboard.press("Escape")
        return False
    idx = random.randint(0, n - 1)
    txt = ""
    try:
        txt = (await opts.nth(idx).inner_text()).strip()[:40]
    except Exception:
        pass
    step(f"{what}: {n} ta variant, #{idx} ('{txt}') tanlanmoqda")
    await opts.nth(idx).click()
    await page.wait_for_timeout(SETTLE_MS)  # keyingi dropdown serverdan yuklanishi uchun
    return True


async def dd_pick_text(page, selector, label, what):
    """p-dropdown'dan (optionlar yuklanishini kutib) matn bo'yicha variant tanlaydi."""
    step(f"{what}: dropdown ochilmoqda (optionlar kutilmoqda)")
    if not await dd_options_ready(page, selector, what):
        return False
    opt = page.get_by_role("option", name=label, exact=True)
    if await opt.count() == 0:
        opt = page.get_by_role("option", name=label)
    if await opt.count() == 0:
        step(f"{what}: '{label}' topilmadi")
        await page.keyboard.press("Escape")
        return False
    step(f"{what}: '{label}' tanlanmoqda")
    await opt.first.click()
    return True


async def ensure_clean(page):
    """Holatni 0 ga qaytaradi. Qo'shish formasi/dialog/xato ochiq bo'lsa sahifani
    qayta yuklaydi (Бекор tugmasi ishlatilmaydi). Toza bo'lsa hech narsa qilmaydi."""
    busy = (await is_visible(page, CASE_B_DIALOG)
            or await is_visible(page, JSHSHIR_INPUT)
            or await is_visible(page, PHONE_INPUT)
            or await is_visible(page, CASE_A_MSG))
    if busy:
        step("jarayon 0 dan boshlanmoqda (sahifa qayta yuklanmoqda)")
        try:
            await page.goto(TARGET_URL)
        except Exception:
            await page.reload()
        await page.locator(BTN_ADD).first.wait_for(state="visible", timeout=20000)
        await page.wait_for_timeout(300)


async def get_external_mahalla(page):
    """Case A da tashqi manbadagi mahalla nomini oladi."""
    try:
        rows = page.locator(TERR_ROW)
        for i in range(await rows.count()):
            k = (await rows.nth(i).locator(".k").inner_text()).strip()
            if "Маҳалла" in k or "МФЙ" in k:
                return (await rows.nth(i).locator(".v").inner_text()).strip()
    except Exception:
        pass
    return ""


async def classify_search(page):
    """
    Natija aniqlash. MUHIM: 'Фуқаро рўйҳатга олинган маҳаллалар' dialogi
    HAM Case B, HAM Case A (topilmadi) uchun ishlatiladi. Shuning uchun:
      - 'Давом этиш' tugmasi BOR bo'lsa            -> 'B' (ko'chirish mumkin).
      - Xato/topilmadi/'мавжуд эмас' xabari bo'lsa -> 'error' (yakuniy).
      - Forma barqaror chiqsa                      -> 'addable'.
    Forma birinchi chiqib keyin xatoga o'zgarishi mumkin — shuning uchun kuzatib turamiz.
    """
    elapsed = 0
    while elapsed < UI_CHECK_MS:
        if await is_visible(page, CONTINUE_BTN):
            return "B"
        if await is_visible(page, CASE_ERR):
            return "error"
        # forma ko'rinsa ham darrov xulosa qilmaymiz (xato kelishi mumkin)
        await page.wait_for_timeout(POLL_MS)
        elapsed += POLL_MS
    # Oyna tugadi — barqaror holat
    if await is_visible(page, CONTINUE_BTN):
        return "B"
    if await is_visible(page, CASE_ERR):
        return "error"
    if await is_visible(page, PHONE_INPUT) or await is_visible(page, HOUSE_DD):
        return "addable"
    return "other"


async def open_add(page):
    step("'+Қўшиш' bosilmoqda")
    await page.locator(BTN_ADD).first.click()
    await page.locator(JSHSHIR_INPUT).first.wait_for(state="visible", timeout=10000)
    await page.wait_for_timeout(200)


async def fill_form_and_save(page):
    """B dan keyingi forma: telefon + Хонадон + Оила + Оила аъзо тури + Сақлаш."""
    # Telefon
    phone = random_phone()
    step(f"Telefon kiritilmoqda: +998 {phone}")
    ph = page.locator(PHONE_INPUT).first
    await ph.wait_for(state="visible", timeout=8000)
    await ph.click()
    try:
        await ph.type(phone, delay=30)
    except Exception:
        await ph.fill(phone)

    # Хонадон (random)
    await dd_pick_random(page, HOUSE_DD, "Хонадон")
    # Оила (random)
    await dd_pick_random(page, FAMILY_DD, "Оила")
    # Оила аъзо тури = Бошқа
    await dd_pick_text(page, MEMBER_TYPE_DD, FAMILY_TYPE_LABEL, "Оила аъзо тури")

    # Сақлаш
    await click_and_wait_server(page, SAVE_BTN, "Сақлаш")


# ============================================================================
#  ASOSIY OQIM
# ============================================================================

async def main():
    load_path = OUTPUT_FILE if os.path.exists(OUTPUT_FILE) else EXCEL_FILE
    log(f"Yuklanmoqda: {load_path}  ->  natija: {OUTPUT_FILE}")
    wb = openpyxl.load_workbook(load_path)
    sheet = wb.active
    src_maxcol = sheet.max_column
    last_row = END_ROW or sheet.max_row

    # NOTFOUND fayl (Case A uchun) — boshidayoq yaratamiz, doim mavjud bo'lsin
    if os.path.exists(NOTFOUND_FILE):
        nf_wb = openpyxl.load_workbook(NOTFOUND_FILE)
    else:
        nf_wb = openpyxl.Workbook()
    nf_sheet = nf_wb.active
    nf_row = next_empty_row(nf_sheet)   # ketma-ket yozish uchun keyingi bo'sh qator
    if safe_save(nf_wb, NOTFOUND_FILE):
        log(f"NOTFOUND fayl tayyor: {NOTFOUND_FILE} (keyingi qator: {nf_row})")

    cache_set = load_cache()
    if conJsh:
        log(f"Lokal kesh: {len(cache_set)} ta ЖШШИР (takrorlar o'tkaziladi)")

    stats = {"B": 0, "addable": 0, "notadded": 0, "skip": 0, "err": 0}

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
                await ensure_clean(page)
                await open_add(page)

                step(f"ЖШШИР kiritilmoqda: {jshshir}")
                await type_jshshir(page, jshshir)
                await click_and_wait_server(page, SEARCH_BTN, "Қидириш")

                result = await classify_search(page)
                step(f"natija = {result}")

                if result == "B":
                    log("  -> Case B: 'Давом этиш' -> forma to'ldirilmoqda.")
                    cont = page.locator(CONTINUE_BTN).first
                    try:
                        await cont.wait_for(state="visible", timeout=5000)
                    except Exception:
                        pass
                    await _click_robust(page, CONTINUE_BTN, "Давом этиш")
                    await page.wait_for_timeout(SETTLE_MS)
                    await fill_form_and_save(page)
                    log("  -> B: SAQLANDI. Excel YASHIL.")
                    cell.fill = GREEN_FILL
                    stats["B"] += 1
                    outcome = "B"

                elif result == "addable":
                    log("  -> Qo'shish imkoni bor: forma to'ldirilmoqda.")
                    await fill_form_and_save(page)
                    log("  -> SAQLANDI. Excel YASHIL.")
                    cell.fill = GREEN_FILL
                    stats["addable"] += 1
                    outcome = "addable"

                else:  # error yoki other -> qo'shilmadi, EXCELGA yoziladi
                    mahalla = await get_external_mahalla(page)
                    if not mahalla:
                        mahalla = "None"
                    log(f"  -> Qo'shilmadi (davom etib bo'lmadi). Mahalla: {mahalla}. NOTFOUND[{nf_row}] ga.")
                    nf_sheet.cell(row=nf_row, column=1, value=mahalla)
                    for c in range(1, src_maxcol + 1):
                        nf_sheet.cell(row=nf_row, column=c + 1, value=sheet.cell(row=row, column=c).value)
                    nf_row += 1
                    safe_save(nf_wb, NOTFOUND_FILE)
                    cell.fill = RED_FILL
                    stats["notadded"] += 1
                    outcome = "notadded"

            except Exception as e:
                log(f"  !! XATO [{row}]: {e}")
                stats["err"] += 1
                outcome = "ERR"
                if DEBUG_SCREENSHOTS:
                    try:
                        await page.screenshot(path=f"debug_row_{row}.png")
                    except Exception:
                        pass

            if outcome and outcome != "ERR":
                cache_set.add(jshshir)
                cache_add(jshshir, outcome)

            if (row - START_ROW + 1) % SAVE_EVERY == 0:
                safe_save(wb, OUTPUT_FILE)

        safe_save(wb, OUTPUT_FILE)
        safe_save(nf_wb, NOTFOUND_FILE)
        log("\n==== YAKUNLANDI ====")
        log(f"B        (boshqa MFY -> qo'shildi): {stats['B']}")
        log(f"Qo'shildi (to'g'ridan forma)     : {stats['addable']}")
        log(f"Qo'shilmadi (-> NOTFOUND excel)  : {stats['notadded']}")
        log(f"O'tkazib yuborilgan              : {stats['skip']}")
        log(f"Xato                             : {stats['err']}")
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
