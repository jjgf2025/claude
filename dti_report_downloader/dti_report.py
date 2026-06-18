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
import hashlib
import urllib.request
from datetime import date, timedelta
from pathlib import Path

# ── AUTO-UPDATE ────────────────────────────────────────────────────
SCRIPT_VERSION = "1.3"
UPDATE_URL = (
    "https://raw.githubusercontent.com/jjgf2025/claude/"
    "claude/relaxed-keller-6nkadt/dti_report_downloader/dti_report.py"
)

def auto_update():
    """Provjeri GitHub za novu verziju i zamijeni script ako postoji update."""
    print("[UPDATE] Provjera nove verzije...")
    try:
        req = urllib.request.urlopen(UPDATE_URL, timeout=8)
        latest = req.read()

        current_path = os.path.abspath(__file__)
        with open(current_path, "rb") as f:
            current = f.read()

        if hashlib.md5(latest).hexdigest() != hashlib.md5(current).hexdigest():
            print("[UPDATE] Nova verzija pronadjena! Azuriram...")
            backup = current_path + ".bak"
            shutil.copy2(current_path, backup)
            with open(current_path, "wb") as f:
                f.write(latest)
            print("[UPDATE] Script azuriran. Restartuj shortcut jednom.")
            input("Pritisni Enter za izlaz (pa pokreni ponovo)...")
            sys.exit(0)
        else:
            print("[UPDATE] Vec imas najnoviju verziju.")
    except Exception as e:
        print(f"[UPDATE] Preskacam update provjeru ({e})")
# ──────────────────────────────────────────────────────────────────

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

# Chrome for Testing binary - u folderu skripte
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CHROMEDRIVER_PATH = os.path.join(SCRIPT_DIR, "chromedriver.exe")
CHROME_BINARY_PATH = os.path.join(SCRIPT_DIR, "chrome-win64", "chrome.exe")

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


def launch_browser(download_dir):
    if not os.path.exists(CHROME_BINARY_PATH):
        print("[GRESKA] chrome-win64\\chrome.exe nije pronadjen!")
        print(f"  Ocekivana lokacija: {CHROME_BINARY_PATH}")
        print("\n  Pokreni ovu komandu u cmd da skines Chrome for Testing:")
        print(f'  powershell -Command "Invoke-WebRequest -Uri \'https://storage.googleapis.com/chrome-for-testing-public/149.0.7827.115/win64/chrome-win64.zip\' -OutFile \'$env:TEMP\\chrome-win64.zip\'; Expand-Archive \'$env:TEMP\\chrome-win64.zip\' -DestinationPath \'{SCRIPT_DIR}\' -Force"')
        input("\nPritisni Enter za izlaz...")
        sys.exit(1)

    print(f"[OK] Chrome for Testing: {CHROME_BINARY_PATH}")

    options = Options()
    options.binary_location = CHROME_BINARY_PATH

    # Dedicated profil - pamti login između pokretanja
    dti_profile = os.path.join(SCRIPT_DIR, "dti_browser_profile")
    os.makedirs(dti_profile, exist_ok=True)
    options.add_argument(f"--user-data-dir={dti_profile}")
    print(f"[OK] DTI profil: {dti_profile}")

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
    options.add_argument("--disable-gpu")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-extensions")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])

    # Koristi chromedriver.exe iz istog foldera
    if os.path.exists(CHROMEDRIVER_PATH):
        service = Service(executable_path=CHROMEDRIVER_PATH)
    else:
        print("[GRESKA] chromedriver.exe nije pronadjen u folderu skripte!")
        print(f"  Ocekivana lokacija: {CHROMEDRIVER_PATH}")
        input("Pritisni Enter za izlaz...")
        sys.exit(1)

    print("[INFO] Pokretanje browsera...")
    try:
        driver = webdriver.Chrome(service=service, options=options)
        print("[OK] Browser pokrenut uspjesno.")
        return driver
    except Exception as e:
        print(f"\n[GRESKA] Nije moguce pokrenuti browser:")
        print(f"  {e}")
        print("\nMoguca resenja:")
        print("  1. Zatvori Opera GX potpuno pa pokusaj ponovo")
        print("  2. Provjeri da je chromedriver.exe u istom folderu kao script")
        input("\nPritisni Enter za izlaz...")
        sys.exit(1)


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
    auto_update()
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
        driver.get(PORTAL_URL + "/reports/")
        time.sleep(3)

        # Provjeri login ponovo nakon navigacije
        if "login" in driver.current_url.lower() or "signin" in driver.current_url.lower():
            print("\n[!] Nisi ulogovan/a. Uloguj se u browser pa pritisni Enter ovdje...")
            input()
            driver.get(PORTAL_URL + "/reports/")
            time.sleep(3)

        print("[3/5] Podesavanje datuma i tipa reporta...")

        from datetime import datetime as dt

        def pick_date(placeholder, target_date):
            """Klikce kroz Bootstrap Datepicker i bira trazeni datum."""
            inp = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, f"input[placeholder='{placeholder}']")
            ))
            inp.click()
            time.sleep(0.8)

            for _ in range(24):
                try:
                    header = driver.find_element(By.CSS_SELECTOR, "th.datepicker-switch")
                except Exception:
                    break
                current_text = header.text.strip()   # "June 2026"
                target_text  = target_date.strftime("%B %Y")

                if current_text == target_text:
                    # Pronadjen pravi mjesec - klikni dan
                    days = driver.find_elements(
                        By.XPATH, "//td[contains(@class,'day') and not(contains(@class,'old')) and not(contains(@class,'new'))]"
                    )
                    for day_el in days:
                        if day_el.text.strip() == str(target_date.day):
                            day_el.click()
                            time.sleep(0.4)
                            val = inp.get_attribute("value")
                            print(f"[OK] {placeholder}: {val}")
                            return True
                    break

                # Navigacija naprijed/nazad
                try:
                    cur_dt = dt.strptime(current_text, "%B %Y")
                    tgt_dt = dt.strptime(target_text, "%B %Y")
                    if tgt_dt > cur_dt:
                        driver.find_element(By.CSS_SELECTOR, "th.next").click()
                    else:
                        driver.find_element(By.CSS_SELECTOR, "th.prev").click()
                    time.sleep(0.4)
                except Exception:
                    break

            print(f"[GRESKA] Nisam mogao odabrati {placeholder} automatski.")
            return False

        # ── Odabir Custom Purchases ──
        try:
            report_select = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//select[.//option[contains(text(),'Custom Purchases')]]")
            ))
            Select(report_select).select_by_visible_text("Custom Purchases")
            time.sleep(0.5)
            print("[OK] Custom Purchases selektovan.")
        except TimeoutException:
            print("[INFO] Dropdown nije nadjen ili je vec selektovan.")

        # ── Unos datuma ──
        pick_date("Start Date", start_date)
        time.sleep(0.3)
        pick_date("End Date", end_date)
        time.sleep(0.5)

        # ── Klik na Add dugme ──
        print("[4/5] Dodavanje reporta...")
        try:
            add_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[normalize-space(text())='Add' or normalize-space(text())='Submit' "
                           "or normalize-space(text())='Generate' or normalize-space(text())='Add Report']")
            ))
            add_btn.click()
            print(f"[OK] Kliknuto: '{add_btn.text}'")
        except Exception:
            # Pokusaj input[type=submit]
            try:
                sub = driver.find_element(By.CSS_SELECTOR, "input[type='submit'], button[type='submit']")
                sub.click()
                print("[OK] Submit kliknut.")
            except Exception:
                print("[!] Add dugme nije nadjen - klikni ga rucno pa pritisni Enter...")
                input()

        # ── Cekanje da report bude Ready i klik za download ──
        print("[...] Cekanje da report bude spreman (max 120s)...")
        expected_label = f"Purchases from {start_date.strftime('%m/%d/%Y')} to {end_date.strftime('%m/%d/%Y')}"
        downloaded = None
        deadline = time.time() + 120
        while time.time() < deadline:
            try:
                links = driver.find_elements(By.XPATH, "//a[contains(text(),'Purchases from')]")
                for lnk in links:
                    if start_date.strftime("%m/%d/%Y") in lnk.text or start_date.strftime("%-m/%-d/%Y") in lnk.text:
                        lnk.click()
                        print(f"[OK] Kliknuto: {lnk.text}")
                        time.sleep(3)
                        downloaded = wait_for_download(temp_dl, timeout=30)
                        break
            except Exception:
                pass
            if downloaded:
                break
            time.sleep(2)
            driver.refresh() if time.time() % 15 < 2 else None

        if not downloaded:
            # Pokusaj direktno cekati fajl (ako download vec pokrenut)
            downloaded = wait_for_download(temp_dl, timeout=30)

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
