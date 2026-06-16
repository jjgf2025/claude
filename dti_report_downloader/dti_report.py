"""
DTI Portal - Custom Purchases Report Downloader
------------------------------------------------
PONEDELJAK : skida Cetvrtak - Subota prethodne sedmice
CETVRTAK   : skida Ponedeljak - Sreda tekuce sedmice

Koristi postojeci Opera GX profil (ne treba login).
"""

import os
import sys
import time
import glob
import shutil
from datetime import date, timedelta
from pathlib import Path

# ── openpyxl za highlight ──────────────────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font
except ImportError:
    os.system('pip install openpyxl -q')
    import openpyxl
    from openpyxl.styles import PatternFill, Font

# ── selenium ───────────────────────────────────────────────────────
try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait, Select
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException, NoSuchElementException
except ImportError:
    os.system('pip install selenium -q')
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait, Select
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ══════════════════════════════════════════════════════════════════
#  KONFIGURACIJA  –  prilagodi po potrebi
# ══════════════════════════════════════════════════════════════════

PORTAL_URL = "https://www.dtiportal.com"

# Desktop folder za cuvanje reporta
SAVE_FOLDER = str(Path.home() / "Desktop" / "DTI_Reports")

# Opera GX putanje  (auto-detect, ali mozes rucno podesiti)
OPERA_BINARY_CANDIDATES = [
    r"C:\Users\{user}\AppData\Local\Programs\Opera GX\launcher.exe",
    r"C:\Users\{user}\AppData\Local\Programs\Opera GX\opera.exe",
    r"C:\Program Files\Opera GX\opera.exe",
    r"C:\Program Files (x86)\Opera GX\opera.exe",
]
OPERA_PROFILE_CANDIDATES = [
    r"C:\Users\{user}\AppData\Roaming\Opera Software\Opera GX Stable",
    r"C:\Users\{user}\AppData\Local\Opera Software\Opera GX Stable",
]

# ChromeDriver 147 - stavi chromedriver.exe u isti folder kao ovaj script
# Skini sa: googlechromelabs.github.io/chrome-for-testing/ → 147.x → win64
CHROMEDRIVER_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chromedriver.exe")

# ══════════════════════════════════════════════════════════════════

RED_FILL = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
WHITE_FONT = Font(color="FFFFFF", bold=True)


def get_date_range():
    """Vraca (start_date, end_date) na osnovu dana u sedmici."""
    today = date.today()
    wd = today.weekday()  # Mon=0, Thu=3

    if wd == 0:  # Ponedeljak
        # Prethodna srijeda je -4 cetvrtak, -2 subota
        start = today - timedelta(days=4)   # cetvrtak
        end   = today - timedelta(days=2)   # subota
    elif wd == 3:  # Cetvrtak
        start = today - timedelta(days=3)   # ponedeljak
        end   = today - timedelta(days=1)   # srijeda
    else:
        day_name = today.strftime("%A")
        msg = f"Danas je {day_name}. Script se pokrece samo Ponedeljkom i Cetvrto."
        print(msg)
        input("\nPritisni Enter za izlaz...")
        sys.exit(0)

    return start, end


def find_opera():
    user = os.environ.get("USERNAME", os.environ.get("USER", ""))
    for c in OPERA_BINARY_CANDIDATES:
        p = c.replace("{user}", user)
        if os.path.exists(p):
            return p
    return None


def find_opera_profile():
    user = os.environ.get("USERNAME", os.environ.get("USER", ""))
    for c in OPERA_PROFILE_CANDIDATES:
        p = c.replace("{user}", user)
        if os.path.exists(p):
            return p
    return None


def launch_browser(download_dir):
    opera_bin = find_opera()
    opera_profile = find_opera_profile()

    if not opera_bin:
        print("[GRESKA] Opera GX nije pronadjena. Provjeri putanju u skripti.")
        input("Enter za izlaz...")
        sys.exit(1)

    print(f"[OK] Opera pronadjena: {opera_bin}")
    if opera_profile:
        print(f"[OK] Opera profil: {opera_profile}")

    options = Options()
    options.binary_location = opera_bin

    # Koristi postojeci profil (da bude vec ulogovan)
    if opera_profile:
        options.add_argument(f"--user-data-dir={opera_profile}")

    # Download postavke
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    # options.add_argument("--headless")  # Uncomment za invisible mod

    # Koristi chromedriver.exe iz istog foldera
    if os.path.exists(CHROMEDRIVER_PATH):
        service = Service(executable_path=CHROMEDRIVER_PATH)
    else:
        print("[GRESKA] chromedriver.exe nije pronadjen u folderu skripte!")
        print(f"  Ocekivana lokacija: {CHROMEDRIVER_PATH}")
        print("  Skini ChromeDriver 147 sa: googlechromelabs.github.io/chrome-for-testing/")
        input("Pritisni Enter za izlaz...")
        sys.exit(1)

    driver = webdriver.Chrome(service=service, options=options)
    return driver


def wait_for_download(download_dir, timeout=60):
    """Ceka da se .xlsx / .xls fajl pojavi u folderu."""
    end_time = time.time() + timeout
    while time.time() < end_time:
        # Ignorisi .crdownload (Chrome privremeni fajl)
        files = [f for f in glob.glob(os.path.join(download_dir, "*.xls*"))
                 if not f.endswith(".crdownload")]
        if files:
            return max(files, key=os.path.getctime)
        time.sleep(1)
    return None


def highlight_blank_order_ids(filepath):
    """Highlightuje crveno sve redove gdje je Order ID prazan."""
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    # Nadji kolonu 'Order ID' / 'Order #' u prvom redu (header)
    header_row = 1
    order_id_col = None
    for cell in ws[header_row]:
        if cell.value and "order" in str(cell.value).lower():
            order_id_col = cell.column
            print(f"[OK] Pronadjena kolona '{cell.value}' na poziciji {cell.column}")
            break

    if order_id_col is None:
        print("[UPOZORENJE] Kolona 'Order ID' nije pronadjena. Provjeri naziv kolone u reportu.")
        wb.save(filepath)
        return 0

    count = 0
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
        order_cell = row[order_id_col - 1]
        val = order_cell.value
        if val is None or str(val).strip() == "":
            for cell in row:
                cell.fill = RED_FILL
                cell.font = WHITE_FONT
            count += 1

    wb.save(filepath)
    print(f"[OK] Highlightovano {count} redova sa praznim Order ID.")
    return count


def set_date_input(driver, selector, date_val):
    """Unosi datum u input polje (format MM/DD/YYYY)."""
    formatted = date_val.strftime("%m/%d/%Y")
    try:
        el = driver.find_element(By.CSS_SELECTOR, selector)
        driver.execute_script("arguments[0].value = '';", el)
        el.clear()
        el.send_keys(formatted)
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change', {bubbles:true}))", el
        )
        return True
    except NoSuchElementException:
        return False


def run():
    start_date, end_date = get_date_range()
    print(f"\n{'='*50}")
    print(f"  DTI Portal - Custom Purchases Report")
    print(f"  Period: {start_date.strftime('%d.%m.%Y')} - {end_date.strftime('%d.%m.%Y')}")
    print(f"{'='*50}\n")

    # Kreiraj folder za cuvanje
    os.makedirs(SAVE_FOLDER, exist_ok=True)

    # Privremeni download folder
    temp_dl = os.path.join(SAVE_FOLDER, "_temp_download")
    os.makedirs(temp_dl, exist_ok=True)

    driver = launch_browser(temp_dl)
    wait = WebDriverWait(driver, 20)

    try:
        print("[1/5] Otvaranje DTI portala...")
        driver.get(PORTAL_URL)
        time.sleep(3)

        # Provjeri da li smo ulogovani (ako nije, cekaj)
        if "login" in driver.current_url.lower() or "signin" in driver.current_url.lower():
            print("\n[!] Nisi ulogovan/a. Uloguj se u browser pa pritisni Enter ovdje...")
            input()
            time.sleep(2)

        print("[2/5] Navigacija na Reports sekciju...")
        try:
            # Pokusaj kliknuti na Reports u navigaciji
            reports_link = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//a[contains(translate(text(),'REPORTS','reports'),'report')]")
                )
            )
            reports_link.click()
            time.sleep(2)
        except TimeoutException:
            # Ako nema direktnog linka, idi na /reports URL
            driver.get(PORTAL_URL + "/reports")
            time.sleep(2)

        print("[3/5] Podesavanje datuma i tipa reporta...")

        # ── Odabir Custom Purchases iz dropdown menija ──
        try:
            report_select = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//select[.//option[contains(text(),'Custom Purchases')]]")
                )
            )
            Select(report_select).select_by_visible_text("Custom Purchases")
            time.sleep(1)
        except TimeoutException:
            print("[GRESKA] Nije moguce naci dropdown za tip reporta.")
            print("Sacuvaj stranicu rucno pa pritisni Enter...")
            input()

        # ── Unos datuma ──
        # Probamo vise mogucih selektora za Date From/To
        date_from_selectors = [
            "input[name*='from']", "input[name*='start']",
            "input[placeholder*='From']", "input[id*='from']",
            "input[id*='start']", ".date-from input", "#dateFrom",
        ]
        date_to_selectors = [
            "input[name*='to']", "input[name*='end']",
            "input[placeholder*='To']", "input[id*='to']",
            "input[id*='end']", ".date-to input", "#dateTo",
        ]

        from_set = False
        for sel in date_from_selectors:
            if set_date_input(driver, sel, start_date):
                print(f"[OK] Date From: {start_date.strftime('%m/%d/%Y')}")
                from_set = True
                break

        to_set = False
        for sel in date_to_selectors:
            if set_date_input(driver, sel, end_date):
                print(f"[OK] Date To: {end_date.strftime('%m/%d/%Y')}")
                to_set = True
                break

        if not from_set or not to_set:
            print("[!] Datumi nisu automatski uneti. Unesi ih rucno i pritisni Enter...")
            input()

        # ── Klik na Export / Download dugme ──
        print("[4/5] Pokretanje downloada...")
        try:
            export_btn = driver.find_element(
                By.XPATH,
                "//button[contains(translate(text(),'EXPORTDOWNLOAD','exportdownload'),'export') or "
                "contains(translate(text(),'EXPORTDOWNLOAD','exportdownload'),'download') or "
                "contains(translate(text(),'EXPORTDOWNLOAD','exportdownload'),'excel')]"
            )
            export_btn.click()
        except NoSuchElementException:
            print("[!] Export dugme nije automatski pronadjeno. Klikni ga rucno pa pritisni Enter...")
            input()

        print("[...] Cekanje na download (max 60s)...")
        downloaded = wait_for_download(temp_dl, timeout=60)

        if not downloaded:
            print("[GRESKA] Fajl nije skinut u roku od 60 sekundi.")
            input("Pritisni Enter za izlaz...")
            return

        print(f"[OK] Fajl skinut: {os.path.basename(downloaded)}")

        # ── Premesti i preimenuj fajl ──
        day_label = "PON" if date.today().weekday() == 0 else "CET"
        new_name = f"CustomPurchases_{day_label}_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"
        final_path = os.path.join(SAVE_FOLDER, new_name)

        # Ako je .xls (stari format), sacuvaj kao xlsx
        if downloaded.endswith(".xls") and not downloaded.endswith(".xlsx"):
            shutil.copy2(downloaded, final_path.replace(".xlsx", ".xls"))
            final_path = final_path.replace(".xlsx", ".xls")
        else:
            shutil.copy2(downloaded, final_path)
        os.remove(downloaded)

        print(f"\n[5/5] Highlight praznih Order ID redova...")
        blank_count = highlight_blank_order_ids(final_path)

        print(f"\n{'='*50}")
        print(f"  GOTOVO!")
        print(f"  Fajl: {final_path}")
        print(f"  Prazni Order IDs (crveno): {blank_count}")
        print(f"{'='*50}\n")

    except Exception as e:
        print(f"\n[GRESKA] {e}")
        import traceback
        traceback.print_exc()
    finally:
        time.sleep(2)
        driver.quit()
        # Ocisti temp folder
        try:
            shutil.rmtree(temp_dl, ignore_errors=True)
        except Exception:
            pass
        input("\nPritisni Enter za zatvaranje...")


if __name__ == "__main__":
    run()
