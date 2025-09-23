#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import json
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import shutil
import subprocess

# Configuration
GOOGLE_SHEET_URL = os.getenv('GOOGLE_SHEET_URL', 'https://docs.google.com/spreadsheets/d/1uEbsT3PZ8tdwiU1Xga_hS6uPve2H74xD5wUci0EcT0Q/edit?gid=0#gid=0')
GOOGLE_SHEET_NAME = os.getenv('GOOGLE_SHEET_NAME', 'ชีต1')
USERNAME = "01000566"
PASSWORD = "01000566"

def clean_html(cell):
    try:
        return BeautifulSoup(cell.get_attribute("innerHTML").strip(), "html.parser").get_text(strip=True)
    except Exception as e:
        print(f"⚠️ Error cleaning HTML: {e}")
        return ""

def parse_row(row):
    try:
        cols = row.find_elements(By.TAG_NAME, "td")
        if len(cols) < 8:
            return None
        return [clean_html(cols[i]) for i in range(1, 8)]
    except Exception as e:
        print(f"⚠️ Error parsing row: {e}")
        return None

def normalize_job_no(job_no):
    return job_no.split("/")[0].strip().lower()

def _detect_chrome_binary():
    # ลำดับการหา Chrome binary
    # 1) จาก env CHROME_BIN (มาจาก setup-chrome action)
    env_bin = os.getenv("CHROME_BIN")
    if env_bin and os.path.exists(env_bin):
        return env_bin

    # 2) ค่า default ที่พบบ่อย
    candidates = [
        "/usr/bin/google-chrome",
        "/usr/bin/chromium-browser",
        "/usr/bin/chromium",
        "/opt/google/chrome/chrome",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c

    # 3) which
    which = shutil.which("google-chrome") or shutil.which("chromium") or shutil.which("chromium-browser")
    return which

def _detect_chromedriver():
    # 1) จาก env CHROMEDRIVER (มาจาก setup-chrome action)
    env_drv = os.getenv("CHROMEDRIVER")
    if env_drv and os.path.exists(env_drv):
        return env_drv

    # 2) PATH
    which = shutil.which("chromedriver")
    if which:
        return which

    # 3) ตำแหน่งยอดฮิต
    for c in ["/usr/local/bin/chromedriver", "/usr/bin/chromedriver"]:
        if os.path.exists(c):
            return c
    return None

def setup_driver():
    """Setup Chrome WebDriver with multi-fallback; prefer Selenium Manager"""
    print("🔧 Setting up Chrome WebDriver...")
    options = Options()
    # headless เสถียรบน GHA
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-features=TranslateUI")
    options.add_argument("--disable-background-timer-throttling")
    options.add_argument("--disable-renderer-backgrounding")
    options.add_argument("--disable-ipc-flooding-protection")
    options.add_argument("--remote-debugging-port=9222")

    # ไม่ปิด JavaScript/Images เพราะเว็บส่วนใหญ่ต้องใช้ในการ login/render
    # options.add_argument("--disable-images")  # ถ้าจำเป็นค่อยเปิด
    # options.add_argument("--disable-javascript")

    chrome_binary = _detect_chrome_binary()
    if chrome_binary:
        print(f"🔎 Detected Chrome binary: {chrome_binary}")
        options.binary_location = chrome_binary
    else:
        print("⚠️ Chrome binary not found via known paths; relying on default.")

    # กลยุทธ์หลายชั้น:
    # A) ลองให้ Selenium Manager จัดการ (ไม่ระบุ path) — เวิร์กใน Selenium 4.6+ ส่วนใหญ่
    # B) ถ้าไม่ผ่าน ลองใช้ chromedriver จาก env/ระบบ
    last_error = None

    # A) Selenium Manager (ไม่ใส่ service.path)
    try:
        print("🔄 Try A: Selenium Manager (auto driver)")
        driver = webdriver.Chrome(options=options)
        driver.get("about:blank")
        print("✅ Selenium Manager pathless driver OK")
        return driver
    except Exception as e:
        print(f"⚠️ Selenium Manager failed: {e}")
        last_error = e

    # B) ใช้พาธ chromedriver ที่เจอ
    try:
        print("🔄 Try B: Use detected chromedriver path")
        chromedriver_path = _detect_chromedriver()
        if not chromedriver_path:
            raise RuntimeError("No chromedriver path found in env/PATH/candidates.")
        print(f"🔎 Using chromedriver at: {chromedriver_path}")
        service = Service(chromedriver_path)
        driver = webdriver.Chrome(service=service, options=options)
        driver.get("about:blank")
        print("✅ Explicit chromedriver OK")
        return driver
    except Exception as e:
        print(f"❌ All driver setups failed. Last: {e}")
        raise last_error or e

def login_to_system(driver):
    try:
        print("🔐 Logging in...")
        driver.get("https://jobm.edoclite.com/jobManagement/pages/login")

        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.NAME, "username"))
        )

        driver.find_element(By.NAME, "username").send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)

        # ปุ่ม login ระบุ selector ให้แน่นขึ้น (กัน DOM เปลี่ยน)
        # ถ้าปุ่มเป็น <button type="submit"> ให้กด submit ผ่าน form ตรง ๆ
        driver.find_element(By.NAME, "password").submit()

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(., 'งานใหม่')]"))
        )
        print("✅ Login successful")
        return True
    except Exception as e:
        print(f"❌ Login failed: {e}")
        return False

def fetch_new_jobs(driver):
    try:
        print("📥 Fetching new jobs...")
        driver.get("https://jobm.edoclite.com/jobManagement/pages/index?tab=13")

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
        )

        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        data = []
        for row in rows:
            parsed = parse_row(row)
            if parsed:
                data.append(parsed)

        print(f"📊 Found {len(data)} new jobs")
        return data
    except Exception as e:
        print(f"❌ Error fetching new jobs: {e}")
        return []

def fetch_closed_jobs(driver):
    try:
        print("📦 Fetching closed jobs...")
        driver.get("https://jobm.edoclite.com/jobManagement/pages/index?tab=15")

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
        )

        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        closed = set()
        for row in rows:
            try:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 2:
                    job_no = normalize_job_no(clean_html(cols[1]))
                    if job_no:
                        closed.add(job_no)
            except Exception as e:
                print(f"⚠️ Error parsing closed row: {e}")
                continue

        print(f"📊 Found {len(closed)} closed jobs")
        return closed
    except Exception as e:
        print(f"❌ Error fetching closed jobs: {e}")
        return set()

def setup_google_sheets():
    """Connect to Google Sheets using a Service Account (modern auth)."""
    import json, pathlib, os, re
    import gspread
    from google.oauth2.service_account import Credentials as GCreds
    from gspread.exceptions import APIError, SpreadsheetNotFound

    print("📄 Connecting to Google Sheets...")

    cred_path = pathlib.Path("credentials.json")
    if not cred_path.exists() or cred_path.stat().st_size == 0:
        raise RuntimeError("credentials.json missing or empty")

    # 1) ใช้ scopes ใหม่ (ต้องเปิดทั้ง Sheets API + Drive API)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    # 2) โหลด credential (raw JSON ไฟล์)
    data = json.loads(cred_path.read_text(encoding="utf-8"))
    client_email = data.get("client_email")
    print(f"🔐 Service Account: {client_email}")

    creds = GCreds.from_service_account_info(data, scopes=scopes)
    gc = gspread.authorize(creds)

    # 3) เปิดชีตด้วย URL (หรือจะสลับไปใช้ key ก็ได้)
    url = os.getenv(
        "GOOGLE_SHEET_URL",
        "https://docs.google.com/spreadsheets/d/1uEbsT3PZ8tdwiU1Xga_hS6uPve2H74xD5wUci0EcT0Q/edit?gid=0#gid=0",
    )
    sheet_name = os.getenv("GOOGLE_SHEET_NAME", "ชีต1")
    print(f"🔗 Target URL: {url}")
    print(f"📑 Worksheet: {sheet_name}")

    try:
        sh = gc.open_by_url(url)   # ถ้าอยากกันปัญหา URL → ใช้ open_by_key ได้ (ดูบล็อกด้านล่าง)
        ws = sh.worksheet(sheet_name)
        print("✅ Connected to Google Sheets")
        return ws
    except SpreadsheetNotFound as e:
        # สาเหตุหลัก ๆ: ยังไม่ได้แชร์ชีตให้ service account / URL ผิด / ใช้บัญชีคนละโดเมนที่บล็อกการแชร์
        msg = (
            "Spreadsheet not found or no access.\n"
            f"- Make sure the sheet is shared to: {client_email} (Editor)\n"
            "- Double-check the URL/key and that Drive API is enabled.\n"
        )
        print("❌ SpreadsheetNotFound:", e or "(no message)")
        raise RuntimeError(msg)
    except APIError as e:
        # มักเป็น 403: insufficient permissions / scope ไม่ตรง
        print("❌ Google APIError:", repr(e))
        raise
    except Exception as e:
        print("❌ Error connecting to Google Sheets:", repr(e))
        raise


    def _load_json_str_maybe_base64(s: str) -> dict:
        """รับสตริงที่อาจเป็น JSON ตรง ๆ หรือ base64-encoded JSON"""
        s = s.strip()
        # ลอง parse เป็น JSON ตรง ๆ ก่อน
        try:
            return json.loads(s)
        except Exception:
            pass
        # ถ้าไม่ใช่ JSON ตรง ๆ ลอง base64
        try:
            decoded = base64.b64decode(s).decode("utf-8")
            return json.loads(decoded)
        except Exception as e:
            raise RuntimeError(f"GOOGLE_SERVICE_ACCOUNT_JSON is neither JSON nor valid base64 JSON: {e}")

    try:
        if cred_path.exists() and cred_path.stat().st_size > 0:
            # จากไฟล์
            data = json.loads(cred_path.read_text(encoding="utf-8"))
            creds_obj = ServiceAccountCredentials.from_json_keyfile_dict(data, scope)
            print("🔐 Using credentials from credentials.json")
        else:
            # จาก ENV (raw หรือ base64)
            env_val = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
            if not env_val:
                raise RuntimeError(
                    "credentials.json not found and GOOGLE_SERVICE_ACCOUNT_JSON is empty."
                )
            data = _load_json_str_maybe_base64(env_val)
            creds_obj = ServiceAccountCredentials.from_json_keyfile_dict(data, scope)
            print("🔐 Using credentials from GOOGLE_SERVICE_ACCOUNT_JSON (env)")
    except Exception as e:
        print(f"❌ Error loading credentials: {e}")
        raise

    # 3) Authorize และเปิดชีต
    try:
        client = gspread.authorize(creds_obj)
        sheet = client.open_by_url(GOOGLE_SHEET_URL).worksheet(GOOGLE_SHEET_NAME)
        print("✅ Connected to Google Sheets")
        return sheet
    except Exception as e:
        print(f"❌ Error connecting to Google Sheets: {e}")
        raise


def update_google_sheets(sheet, new_jobs, closed_job_nos):
    try:
        print("✏️ Updating Google Sheets...")
        sheet_data = sheet.get_all_values()
        if not sheet_data:
            headers = ["Job No", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Status"]
            sheet.append_row(headers)
            sheet_data = [headers]

        existing = set()
        for row in sheet_data[1:]:
            if row and len(row) > 0:
                existing.add(normalize_job_no(row[0]))

        new_added = 0
        updated = 0

        for job in new_jobs:
            if not job or len(job) < 7:
                continue

            job_no = normalize_job_no(job[0])
            status = "ปิดงาน" if job_no in closed_job_nos else "รอแจ้ง"

            if job_no not in existing:
                try:
                    sheet.append_row(job + [status], value_input_option="USER_ENTERED")
                    print(f"✅ Added: {job_no}")
                    new_added += 1
                    time.sleep(0.5)
                except Exception as e:
                    print(f"❌ Error adding job {job_no}: {e}")
            elif status == "ปิดงาน":
                try:
                    # หาแถวเดิมเพื่ออัพเดตคอลัมน์สถานะ (คอลัมน์ที่ 8)
                    for i, row in enumerate(sheet_data[1:], start=2):
                        if row and len(row) > 0 and normalize_job_no(row[0]) == job_no:
                            if len(row) < 8 or row[7] != "ปิดงาน":
                                sheet.update_cell(i, 8, "ปิดงาน")
                                print(f"🔒 Updated status: {job_no}")
                                updated += 1
                                time.sleep(0.5)
                            break
                except Exception as e:
                    print(f"❌ Error updating job {job_no}: {e}")

        print(f"📊 Summary: {new_added} new jobs added, {updated} jobs updated")
        return {"new_added": new_added, "updated": updated}
    except Exception as e:
        print(f"❌ Error updating Google Sheets: {e}")
        return {"new_added": 0, "updated": 0, "error": str(e)}

def main():
    print(f"🚀 Starting job fetch process at {datetime.now()}")
    driver = None
    try:
        driver = setup_driver()

        if not login_to_system(driver):
            raise Exception("Login failed")

        new_jobs = fetch_new_jobs(driver)
        closed_job_nos = fetch_closed_jobs(driver)

        sheet = setup_google_sheets()
        result = update_google_sheets(sheet, new_jobs, closed_job_nos)

        print("✅ Process completed successfully!")
        print(f"📊 Results: {result}")
    except Exception as e:
        print(f"❌ Process failed: {e}")
        exit(1)
    finally:
        if driver:
            try:
                driver.quit()
                print("🔧 WebDriver closed")
            except Exception as e:
                print(f"⚠️ Error closing driver: {e}")

if __name__ == "__main__":
    main()
