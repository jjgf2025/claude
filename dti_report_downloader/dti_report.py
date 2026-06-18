"""
DTI Portal - Custom Purchases Report Downloader
PONEDELJAK : Cetvrtak - Subota prethodne sedmice
CETVRTAK   : Ponedeljak - Sreda tekuce sedmice
"""

import os, sys, time, glob, shutil, hashlib, subprocess, json, csv as csv_mod
import urllib.request
from datetime import date, timedelta, datetime
from pathlib import Path

# ── AUTO-UPDATE ───────────────────────────────────────────────────
SCRIPT_VERSION = "3.0"
UPDATE_URL = ("https://raw.githubusercontent.com/jjgf2025/claude/"
              "claude/relaxed-keller-6nkadt/dti_report_downloader/dti_report.py")

def auto_update():
    print("[UPDATE] Provjera nove verzije...")
    try:
        latest  = urllib.request.urlopen(UPDATE_URL, timeout=8).read()
        current_path = os.path.abspath(__file__)
        current = open(current_path, "rb").read()
        if hashlib.md5(latest).hexdigest() != hashlib.md5(current).hexdigest():
            print("[UPDATE] Nova verzija! Azuriram...")
            shutil.copy2(current_path, current_path + ".bak")
            open(current_path, "wb").write(latest)
            print("[UPDATE] Azurirano. Pokreni ponovo.")
            input("Enter..."); sys.exit(0)
        else:
            print("[UPDATE] Najnovija verzija.")
    except Exception as e:
        print(f"[UPDATE] Preskacam ({e})")

# ── KONFIGURACIJA ─────────────────────────────────────────────────
PORTAL_URL  = "https://www.dtiportal.com"
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
SAVE_FOLDER = str(Path.home() / "Desktop" / "DTI_Reports")
CREDS_FILE  = os.path.join(SCRIPT_DIR, "dti_creds.json")
PROFILE_DIR = os.path.join(SCRIPT_DIR, "pw_profile")

# ── BIBLIOTEKE ────────────────────────────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font
except ImportError:
    subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])
    import openpyxl
    from openpyxl.styles import PatternFill, Font

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("[INFO] Instaliram Playwright...")
    subprocess.run([sys.executable, "-m", "pip", "install", "playwright", "-q"])
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"])
    from playwright.sync_api import sync_playwright

RED_FILL   = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
WHITE_FONT = Font(color="FFFFFF", bold=True)

# ── KREDENCIJALI ──────────────────────────────────────────────────
def get_credentials():
    if os.path.exists(CREDS_FILE):
        with open(CREDS_FILE) as f:
            c = json.load(f)
        return c.get("email"), c.get("password")
    print("\n[LOGIN] Unesi kredencijale za DTI portal (cuvaju se lokalno, jednom):")
    email    = input("  Email: ").strip()
    password = input("  Lozinka: ").strip()
    with open(CREDS_FILE, "w") as f:
        json.dump({"email": email, "password": password}, f)
    print("[OK] Kredencijali sacuvani.\n")
    return email, password

def do_login(page, email, password):
    page.wait_for_timeout(1500)
    try:
        page.fill("input[type='email'], input[name='email'], input[name='username'], #email, #username", email)
        page.fill("input[type='password'], input[name='password'], #password", password)
        page.click("button[type='submit'], input[type='submit'], .btn-login, .login-btn")
        page.wait_for_timeout(2000)
        print("[OK] Ulogovan automatski.")
        return True
    except Exception as e:
        print(f"[!] Auto-login nije uspio: {e}")
        return False

# ── DATUM ─────────────────────────────────────────────────────────
def get_date_range():
    today = date.today()
    wd = today.weekday()
    if wd == 0:
        return today - timedelta(days=4), today - timedelta(days=2)
    elif wd == 3:
        return today - timedelta(days=3), today - timedelta(days=1)
    else:
        print(f"Danas je {today.strftime('%A')}. Script radi samo Pon i Cet.")
        input("Enter..."); sys.exit(0)

# ── CALENDAR PICKER ───────────────────────────────────────────────
def pick_date(page, placeholder, target_date):
    loc = page.locator(f"input[placeholder='{placeholder}']:visible").first
    loc.scroll_into_view_if_needed()
    loc.click()
    page.wait_for_timeout(700)
    for _ in range(24):
        try:
            header  = page.locator("th.datepicker-switch").first.inner_text(timeout=2000).strip()
            target  = target_date.strftime("%B %Y")
            if header == target:
                for d in page.locator("td.day:not(.old):not(.new)").all():
                    if d.inner_text(timeout=300).strip() == str(target_date.day):
                        d.click(); page.wait_for_timeout(400)
                        print(f"[OK] {placeholder}: {target_date.strftime('%m/%d/%Y')}")
                        return True
                break
            cur = datetime.strptime(header, "%B %Y")
            tgt = datetime.strptime(target,  "%B %Y")
            page.click("th.next" if tgt > cur else "th.prev")
            page.wait_for_timeout(300)
        except Exception:
            break
    print(f"[!] Nije moguce automatski postaviti {placeholder}")
    return False

# ── CSV → EXCEL ───────────────────────────────────────────────────
def csv_to_xlsx(csv_path, xlsx_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    order_col = None
    count = 0
    with open(csv_path, newline='', encoding='utf-8-sig', errors='replace') as f:
        for row_idx, row in enumerate(csv_mod.reader(f), start=1):
            for col_idx, val in enumerate(row, start=1):
                ws.cell(row=row_idx, column=col_idx, value=val)
            if row_idx == 1:
                for i, v in enumerate(row, start=1):
                    if "order" in v.lower():
                        order_col = i
                        print(f"[OK] Order kolona: '{v}'")
                        break
            elif order_col and row_idx > 1:
                cell_val = row[order_col-1] if len(row) >= order_col else ""
                if not cell_val.strip():
                    for c in range(1, len(row)+1):
                        ws.cell(row=row_idx, column=c).fill = RED_FILL
                        ws.cell(row=row_idx, column=c).font = WHITE_FONT
                    count += 1
    wb.save(xlsx_path)
    print(f"[OK] Highlightovano {count} redova sa praznim Order ID.")
    return count

# ── MAIN ──────────────────────────────────────────────────────────
def run():
    auto_update()
    start_date, end_date = get_date_range()
    email, password = get_credentials()

    print(f"\n{'='*50}")
    print(f"  DTI Portal - Custom Purchases Report")
    print(f"  Period: {start_date.strftime('%d.%m.%Y')} - {end_date.strftime('%d.%m.%Y')}")
    print(f"{'='*50}\n")

    os.makedirs(SAVE_FOLDER, exist_ok=True)
    os.makedirs(PROFILE_DIR, exist_ok=True)

    with sync_playwright() as pw:
        print("[INFO] Pokretanje browsera...")
        context = pw.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=False,
            args=["--no-sandbox", "--disable-blink-features=AutomationControlled"],
            ignore_default_args=["--enable-automation"],
        )
        page = context.pages[0] if context.pages else context.new_page()
        print("[OK] Browser pokrenut.")

        downloaded_file = None
        try:
            # 1. Otvori portal
            print("[1/5] Otvaranje portala...")
            page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(2000)

            # 2. Login ako je potrebno
            if "login" in page.url.lower() or "signin" in page.url.lower() or page.locator("input[type='password']").count() > 0:
                print("[2/5] Logovanje...")
                if not do_login(page, email, password):
                    print("[!] Uloguj se rucno u browser, pa pritisni Enter...")
                    input()
            else:
                print("[2/5] Vec ulogovan.")

            # 3. Idi na Reports
            print("[3/5] Reports stranica...")
            page.goto(PORTAL_URL + "/reports/", wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(2000)

            if "login" in page.url.lower():
                print("[!] Sesija istekla - logovanje ponovo...")
                do_login(page, email, password)
                page.goto(PORTAL_URL + "/reports/", wait_until="domcontentloaded")
                page.wait_for_timeout(2000)

            # 4. Podesi report
            print("[4/5] Podesavanje Custom Purchases...")
            try:
                page.select_option("select", label="Custom Purchases")
                page.wait_for_timeout(800)
                page.locator("input[placeholder='Start Date']:visible").first.scroll_into_view_if_needed()
                page.wait_for_timeout(300)
            except Exception:
                pass

            pick_date(page, "Start Date", start_date)
            page.wait_for_timeout(300)
            pick_date(page, "End Date", end_date)
            page.wait_for_timeout(500)

            # Provjeri da li report vec postoji za ovaj period
            already_exists = False
            try:
                for lnk in page.locator("a:visible").all():
                    try:
                        txt = lnk.inner_text(timeout=300).strip()
                        if "Purchases from" in txt and start_date.strftime("%m/%d") in txt:
                            already_exists = True
                            break
                    except Exception:
                        continue
            except Exception:
                pass

            if not already_exists:
                try:
                    page.click("button:has-text('Add'), button[type='submit']")
                    print("[OK] Report dodat.")
                    page.wait_for_timeout(2000)
                except Exception:
                    print("[!] Klikni Add dugme rucno pa Enter...")
                    input()
            else:
                print("[OK] Report vec postoji.")

            # 5. Cekaj Ready i skini
            print("[5/5] Cekanje da bude Ready i download...")
            deadline = time.time() + 120
            while time.time() < deadline:
                page.wait_for_timeout(3000)
                page.reload(wait_until="domcontentloaded")
                page.wait_for_timeout(1500)
                try:
                    links = page.locator("a:visible").all()
                    purchase_links = []
                    for lnk in links:
                        try:
                            txt = lnk.inner_text(timeout=300).strip()
                            if "Purchases from" in txt and start_date.strftime("%m/%d") in txt:
                                purchase_links.append(lnk)
                        except Exception:
                            continue
                    if purchase_links:
                        target_link = purchase_links[-1]
                        txt = target_link.inner_text(timeout=300).strip()
                        print(f"[OK] Klikcem: {txt}")
                        save_tmp = os.path.join(SAVE_FOLDER, "_dl_tmp")
                        os.makedirs(save_tmp, exist_ok=True)
                        with page.expect_download(timeout=30000) as dl_info:
                            target_link.click()
                        dl = dl_info.value
                        fname = dl.suggested_filename or "report.csv"
                        tmp_path = os.path.join(save_tmp, fname)
                        dl.save_as(tmp_path)
                        if os.path.exists(tmp_path):
                            downloaded_file = tmp_path
                            print(f"[OK] Preuzeto: {fname}")
                            break
                except Exception as ex:
                    print(f"[...] {ex}")
                print("[...] Cekam...")

        except Exception as e:
            print(f"\n[GRESKA] {e}")
            import traceback; traceback.print_exc()
        finally:
            page.wait_for_timeout(500)
            context.close()

    # Obrada fajla
    if not downloaded_file or not os.path.exists(downloaded_file):
        print("[GRESKA] Fajl nije pronadjen.")
        input("Enter za izlaz..."); return

    day_label  = "PON" if date.today().weekday() == 0 else "CET"
    final_path = os.path.join(SAVE_FOLDER,
        f"CustomPurchases_{day_label}_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx")

    print(f"[INFO] Konvertujem u Excel sa highlightom...")
    src_ext = os.path.splitext(downloaded_file)[1].lower()
    if src_ext in (".csv", ""):
        csv_to_xlsx(downloaded_file, final_path)
    else:
        shutil.copy2(downloaded_file, final_path)

    print(f"\n{'='*50}")
    print(f"  GOTOVO!")
    print(f"  Fajl: {final_path}")
    print(f"{'='*50}\n")

    print("[INFO] Otvaram Excel...")
    os.startfile(final_path)
    input("Enter za zatvaranje...")


if __name__ == "__main__":
    run()
