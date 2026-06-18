"""
DTI Portal - Custom Purchases Report Downloader
------------------------------------------------
PONEDELJAK : skida Cetvrtak - Subota prethodne sedmice
CETVRTAK   : skida Ponedeljak - Sreda tekuce sedmice
"""

import os
import sys
import time
import glob
import shutil
import hashlib
import subprocess
import urllib.request
from datetime import date, timedelta, datetime
from pathlib import Path

# ── AUTO-UPDATE ────────────────────────────────────────────────────
SCRIPT_VERSION = "2.0"
UPDATE_URL = (
    "https://raw.githubusercontent.com/jjgf2025/claude/"
    "claude/relaxed-keller-6nkadt/dti_report_downloader/dti_report.py"
)

def auto_update():
    print("[UPDATE] Provjera nove verzije...")
    try:
        req = urllib.request.urlopen(UPDATE_URL, timeout=8)
        latest = req.read()
        current_path = os.path.abspath(__file__)
        with open(current_path, "rb") as f:
            current = f.read()
        if hashlib.md5(latest).hexdigest() != hashlib.md5(current).hexdigest():
            print("[UPDATE] Nova verzija pronadjena! Azuriram...")
            shutil.copy2(current_path, current_path + ".bak")
            with open(current_path, "wb") as f:
                f.write(latest)
            print("[UPDATE] Script azuriran. Restartuj shortcut jednom.")
            input("Pritisni Enter za izlaz (pa pokreni ponovo)...")
            sys.exit(0)
        else:
            print("[UPDATE] Vec imas najnoviju verziju.")
    except Exception as e:
        print(f"[UPDATE] Preskacam ({e})")

# ──────────────────────────────────────────────────────────────────

PORTAL_URL  = "https://www.dtiportal.com"
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
SAVE_FOLDER = str(Path.home() / "Desktop" / "DTI_Reports")

# ── openpyxl ──────────────────────────────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font
except ImportError:
    subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])
    import openpyxl
    from openpyxl.styles import PatternFill, Font

# ── playwright ────────────────────────────────────────────────────
try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    print("[INFO] Instaliram Playwright (jednom)...")
    subprocess.run([sys.executable, "-m", "pip", "install", "playwright", "-q"])
    print("[INFO] Instaliram Chromium browser (jednom, ~150MB)...")
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"])
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

RED_FILL   = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
WHITE_FONT = Font(color="FFFFFF", bold=True)


def get_date_range():
    today = date.today()
    wd = today.weekday()  # Mon=0, Thu=3
    if wd == 0:      # Ponedeljak → Cet-Sub prethodne sedmice
        start = today - timedelta(days=4)
        end   = today - timedelta(days=2)
    elif wd == 3:    # Cetvrtak → Pon-Sri tekuce sedmice
        start = today - timedelta(days=3)
        end   = today - timedelta(days=1)
    else:
        print(f"Danas je {today.strftime('%A')}. Script radi samo Pon i Cet.")
        input("Enter za izlaz...")
        sys.exit(0)
    return start, end


def pick_date(page, placeholder, target_date):
    """Klikce kroz Bootstrap Datepicker - trazi vidljivo polje."""
    # Cekaj da bude vidljivo i skroluj do njega
    locator = page.locator(f"input[placeholder='{placeholder}']:visible").first
    locator.scroll_into_view_if_needed()
    locator.click()
    page.wait_for_timeout(800)

    for _ in range(24):
        header = page.locator("th.datepicker-switch").first
        current = header.inner_text().strip()   # "June 2026"
        target  = target_date.strftime("%B %Y")

        if current == target:
            # Klikni pravi dan
            day_str = str(target_date.day)
            days = page.locator("td.day:not(.old):not(.new)").all()
            for d in days:
                if d.inner_text().strip() == day_str:
                    d.click()
                    page.wait_for_timeout(400)
                    val = page.input_value(f"input[placeholder='{placeholder}']")
                    print(f"[OK] {placeholder}: {val}")
                    return True
            break

        cur_dt = datetime.strptime(current, "%B %Y")
        tgt_dt = datetime.strptime(target,  "%B %Y")
        if tgt_dt > cur_dt:
            page.click("th.next")
        else:
            page.click("th.prev")
        page.wait_for_timeout(300)

    print(f"[!] Nisam mogao automatski postaviti {placeholder}")
    return False


def wait_for_download(download_dir, timeout=60):
    end = time.time() + timeout
    while time.time() < end:
        files = [f for f in glob.glob(os.path.join(download_dir, "*"))
                 if not f.endswith(".crdownload") and not f.endswith(".tmp")
                 and os.path.isfile(f)]
        if files:
            return max(files, key=os.path.getctime)
        time.sleep(1)
    return None


def highlight_blank_order_ids(filepath):
    try:
        wb = openpyxl.load_workbook(filepath)
    except Exception:
        print("[UPOZORENJE] Ne mogu otvoriti Excel za highlight.")
        return 0
    ws = wb.active
    order_id_col = None
    for cell in ws[1]:
        if cell.value and "order" in str(cell.value).lower():
            order_id_col = cell.column
            print(f"[OK] Kolona '{cell.value}' na poziciji {cell.column}")
            break
    if order_id_col is None:
        print("[UPOZORENJE] Kolona 'Order ID' nije pronadjena.")
        wb.save(filepath)
        return 0
    count = 0
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        val = row[order_id_col - 1].value
        if val is None or str(val).strip() == "":
            for cell in row:
                cell.fill = RED_FILL
                cell.font = WHITE_FONT
            count += 1
    wb.save(filepath)
    print(f"[OK] Highlightovano {count} redova sa praznim Order ID.")
    return count


def run():
    auto_update()
    start_date, end_date = get_date_range()
    print(f"\n{'='*50}")
    print(f"  DTI Portal - Custom Purchases Report")
    print(f"  Period: {start_date.strftime('%d.%m.%Y')} - {end_date.strftime('%d.%m.%Y')}")
    print(f"{'='*50}\n")

    os.makedirs(SAVE_FOLDER, exist_ok=True)
    temp_dl = os.path.join(SAVE_FOLDER, "_temp_download")
    os.makedirs(temp_dl, exist_ok=True)

    # Playwright profil - pamti login
    profile_dir = os.path.join(SCRIPT_DIR, "pw_profile")

    with sync_playwright() as pw:
        print("[INFO] Pokretanje Chromium browsera...")
        context = pw.chromium.launch_persistent_context(
            user_data_dir=profile_dir,
            headless=False,
            accept_downloads=True,
            downloads_path=temp_dl,
            args=["--no-sandbox", "--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        page = context.pages[0] if context.pages else context.new_page()
        print("[OK] Browser pokrenut.")

        try:
            print("[1/5] Otvaranje DTI portala...")
            page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(2000)

            if "login" in page.url.lower() or "signin" in page.url.lower():
                print("\n[!] Uloguj se u browser koji se otvorio, pa pritisni Enter ovdje...")
                input()
                page.wait_for_timeout(2000)

            print("[2/5] Navigacija na Reports...")
            page.goto(PORTAL_URL + "/reports/", wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(2000)

            if "login" in page.url.lower():
                print("[!] Potreban login - uloguj se pa pritisni Enter...")
                input()
                page.goto(PORTAL_URL + "/reports/", wait_until="domcontentloaded")
                page.wait_for_timeout(2000)

            print("[3/5] Podesavanje reporta...")

            # Odaberi Custom Purchases
            try:
                page.select_option("select", label="Custom Purchases")
                page.wait_for_timeout(1000)
                print("[OK] Custom Purchases selektovan.")
                # Skroluj do forme sa datumima
                page.locator("input[placeholder='Start Date']:visible").first.scroll_into_view_if_needed()
                page.wait_for_timeout(500)
            except Exception:
                print("[INFO] Dropdown nije nadjen ili vec selektovan.")

            # Unos datuma
            pick_date(page, "Start Date", start_date)
            page.wait_for_timeout(300)
            pick_date(page, "End Date", end_date)
            page.wait_for_timeout(500)

            # Klik na Add dugme
            print("[4/5] Dodavanje reporta...")
            try:
                page.click("button:has-text('Add'), button[type='submit'], input[type='submit']")
                print("[OK] Report dodat.")
            except Exception:
                print("[!] Add dugme nije nadjen automatski. Klikni ga rucno pa pritisni Enter...")
                input()

            # Cekanje da report postane Ready i download
            print("[...] Cekanje da report bude spreman...")
            start_str = start_date.strftime("%m/%d/%Y")
            end_str   = end_date.strftime("%m/%d/%Y")

            downloaded_file = None
            deadline = time.time() + 120
            while time.time() < deadline:
                try:
                    links = page.locator("a").all()
                    for lnk in links:
                        txt = lnk.inner_text()
                        if "Purchases" in txt and (
                            start_date.strftime("%-m/%-d/%Y") in txt or
                            start_str in txt or
                            start_date.strftime("%m/%d/%Y").lstrip("0").replace("/0", "/") in txt
                        ):
                            print(f"[OK] Report spreman: {txt.strip()}")
                            with page.expect_download(timeout=30000) as dl_info:
                                lnk.click()
                            download = dl_info.value
                            save_path = os.path.join(temp_dl, download.suggested_filename or "report.xlsx")
                            download.save_as(save_path)
                            downloaded_file = save_path
                            break
                except Exception:
                    pass
                if downloaded_file:
                    break
                time.sleep(3)
                page.reload(wait_until="domcontentloaded")
                page.wait_for_timeout(1000)

            if not downloaded_file:
                # Pokusaj direktno iz temp foldera
                downloaded_file = wait_for_download(temp_dl, timeout=15)

        except Exception as e:
            print(f"\n[GRESKA] {e}")
            import traceback
            traceback.print_exc()
            downloaded_file = None
        finally:
            page.wait_for_timeout(1000)
            context.close()

    if not downloaded_file or not os.path.exists(downloaded_file):
        print("[GRESKA] Fajl nije skinut.")
        input("Pritisni Enter za izlaz...")
        return

    print(f"[OK] Fajl skinut: {os.path.basename(downloaded_file)}")

    day_label = "PON" if date.today().weekday() == 0 else "CET"
    ext = os.path.splitext(downloaded_file)[1] or ".xlsx"
    new_name = f"CustomPurchases_{day_label}_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}{ext}"
    final_path = os.path.join(SAVE_FOLDER, new_name)
    shutil.copy2(downloaded_file, final_path)
    try:
        os.remove(downloaded_file)
    except Exception:
        pass

    print(f"\n[5/5] Highlight praznih Order ID redova...")
    blank_count = highlight_blank_order_ids(final_path)

    print(f"\n{'='*50}")
    print(f"  GOTOVO!")
    print(f"  Fajl: {final_path}")
    print(f"  Prazni Order IDs (crveno): {blank_count}")
    print(f"{'='*50}\n")

    try:
        shutil.rmtree(temp_dl, ignore_errors=True)
    except Exception:
        pass
    input("Pritisni Enter za zatvaranje...")


if __name__ == "__main__":
    run()
