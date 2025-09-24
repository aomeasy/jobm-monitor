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
import shutil
import subprocess
import re

# Configuration
GOOGLE_SHEET_URL = os.getenv('GOOGLE_SHEET_URL', 'https://docs.google.com/spreadsheets/d/1uEbsT3PZ8tdwiU1Xga_hS6uPve2H74xD5wUci0EcT0Q/edit?gid=0#gid=0')
GOOGLE_SHEET_NAME = os.getenv('GOOGLE_SHEET_NAME', 'ชีต1')
USERNAME = "01000566"
PASSWORD = "01000566"
JOBNO_PAT = re.compile(r"No\d+-\d+")  # จับ No68-0033, No123-4567 ฯลฯ (มีอักษรไทยค้ำหน้าก็เจอ)

def fetch_jobs_by_tab(driver, tab):
    """
    ดึงข้อมูลแถวงานจากหน้า index?tab=<tab>
    คืนค่าเป็น list ของแต่ละงาน [col1..col7] (ตาม parse_row)
    """
    try:
        url = f"https://jobm.edoclite.com/jobManagement/pages/index?tab={tab}"
        print(f"📥 Fetching jobs from tab={tab} ...")
        driver.get(url)

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
        )

        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        data = []
        for row in rows:
            parsed = parse_row(row)
            if parsed:
                data.append(parsed)

        print(f"📊 Found {len(data)} rows on tab={tab}")
        return data
    except Exception as e:
        print(f"❌ Error fetching tab={tab}: {e}")
        return []


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
def parse_row_by_tab(row, tab: int):
    """
    คืน list 7 ช่องเหมือน parse_row() แต่:
    - tab=16: ดักกรณีคอลัมน์ 'Job No.' กับ 'เรื่องที่แจ้ง' สลับกัน แล้วสลับกลับให้
               และทำความสะอาด Job No สำหรับ 'แสดง' (ตัดหลัง '/')
    """
    cols = row.find_elements(By.TAG_NAME, "td")
    if len(cols) < 8:
        return None
    # ค่าดิบตามหน้าเว็บ (ข้ามคอลัมน์ลำดับ)
    raw = [clean_html(cols[i]) for i in range(1, 8)]

    if tab == 16:
        # โดยปกติ raw[0] = Job No., raw[1] = เรื่องที่แจ้ง
        # แต่หน้าเว็บบางครั้งสลับกัน: ตรวจด้วย regex ถ้าเจอ jobno อยู่ที่ raw[1] แต่ไม่อยู่ raw[0] -> สลับกลับ
        has_job0 = bool(JOBNO_PAT.search(raw[0]))
        has_job1 = bool(JOBNO_PAT.search(raw[1]))
        if (not has_job0) and has_job1:
            raw[0], raw[1] = raw[1], raw[0]  # สลับกลับ

        # ทำความสะอาด Job No สำหรับ 'แสดง' (คงรูปเดิมแค่ตัดหลัง '/')
        raw[0] = clean_job_no_display(raw[0])

    return raw


def normalize_job_no(job_no: str) -> str:
    if not job_no:
        return ""
    return job_no.split("/")[0].strip().lower()

def clean_job_no_display(s: str) -> str:
    if not s:
        return ""
    return s.split("/")[0].strip()


def adjust_cols_for_sheet(job: list) -> list:
    """
    ใช้กับข้อมูลจาก tab=15 เท่านั้น
    - Job No: ตัดข้อความหลัง '/' ออก
    - ศูนย์ที่แจ้ง (col C) -> เว้นว่าง
    - ขยับค่าเดิมไปทางขวา 1 ช่อง
    """
    arr = (job or [])[:]
    while len(arr) < 7:
        arr.append("")
    # โครงใหม่: [A(JobNo), B, ''(C), D<-เดิมC, E<-เดิมD, F<-เดิมE, G<-เดิมF]
    return [clean_job_no_display(arr[0]), arr[1], "", arr[2], arr[3], arr[4], arr[5]]
    
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
    import re, json, pathlib, os
    import gspread
    from google.oauth2.service_account import Credentials as GCreds
    from gspread.exceptions import APIError, SpreadsheetNotFound

    print("📄 Connecting to Google Sheets...")
    print(f"🔍 Env has GOOGLE_SHEET_KEY? {bool(os.getenv('GOOGLE_SHEET_KEY'))}")
    print(f"🔍 Env has GOOGLE_SHEET_URL? {bool(os.getenv('GOOGLE_SHEET_URL'))}")


    cred_path = pathlib.Path("credentials.json")
    if not cred_path.exists() or cred_path.stat().st_size == 0:
        raise RuntimeError("credentials.json missing or empty")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    data = json.loads(cred_path.read_text(encoding="utf-8"))
    client_email = data.get("client_email")
    print(f"🔐 Service Account: {client_email}")

    creds = GCreds.from_service_account_info(data, scopes=scopes)
    gc = gspread.authorize(creds)

    # ----- รับค่า target -----
    sheet_name = os.getenv("GOOGLE_SHEET_NAME", "ชีต1")
    # ให้สิทธิ์ส่ง key ตรง ๆ มาก่อน
    key = os.getenv("GOOGLE_SHEET_KEY", "").strip()

    # ถ้าไม่ส่ง key ให้พยายามอ่านจาก URL (env → ค่าคงที่ด้านบน)
    url_env = os.getenv("GOOGLE_SHEET_URL", "").strip()
    url_fallback = GOOGLE_SHEET_URL  # ค่าคงที่บนสุดของไฟล์
    url = url_env or url_fallback

    if not key:
        # ดึง key ด้วย regex (ครอบคลุมหลายรูปแบบ URL)
        m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
        if m:
            key = m.group(1)

    if not key:
        raise RuntimeError("Cannot parse spreadsheet key from GOOGLE_SHEET_URL. "
                           "Set GOOGLE_SHEET_KEY explicitly or pass a standard Sheets URL.")

    print(f"🔗 Spreadsheet key: {key}")
    print(f"📑 Worksheet: {sheet_name}")

    try:
        sh = gc.open_by_key(key)
        ws = sh.worksheet(sheet_name)
        print("✅ Connected to Google Sheets")
        return ws
    except SpreadsheetNotFound as e:
        print("❌ SpreadsheetNotFound:", e or "(no message)")
        raise RuntimeError(
            "Spreadsheet not found or no access.\n"
            f"- Share the sheet to: {client_email} (Editor)\n"
            "- Check the key/URL and ensure Sheets API + Drive API are enabled."
        )
    except APIError as e:
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


def update_google_sheets(sheet, new_jobs, closed_job_nos,
                         waiting_jobs=None, closed_jobs_full=None,
                         closed_already_jobs=None):  # ⬅️ เพิ่ม param สำหรับ tab=16
    """
    - tab=13 : ถ้ายังไม่เจอ -> เพิ่ม พร้อมสถานะ 'รอแจ้ง' หรือ 'ปิดงาน' (ถ้าอยู่ใน closed_job_nos)
               ถ้าเจอแล้วและยังไม่ปิด -> อัปเดตสถานะเป็น 'ปิดงาน'
    - tab=14 : ถ้ายังไม่เจอ -> เพิ่ม พร้อมสถานะ 'รอแจ้ง'
    - tab=15 : ถ้ายังไม่เจอ -> เพิ่ม (ปรับคอลัมน์: C ว่าง + shift ขวา 1) พร้อมสถานะ 'ปิดงาน'
               ถ้าเจอแล้วและยังไม่ปิด -> อัปเดตสถานะเป็น 'ปิดงาน'
    - tab=16 : ถ้ายังไม่เจอ -> เพิ่ม พร้อมสถานะ 'งานที่ปิดแล้ว' (ไม่แก้รายการเดิม)
    """
    waiting_jobs = waiting_jobs or []
    closed_jobs_full = closed_jobs_full or []
    closed_already_jobs = closed_already_jobs or []  # ⬅️ tab=16

    try:
        print("✏️ Updating Google Sheets...")
        sheet_data = sheet.get_all_values()
        if not sheet_data:
            headers = ["Job No", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Status"]
            sheet.append_row(headers)
            sheet_data = [headers]

        # ทำดัชนีข้อมูลเดิมในชีต (ใช้ compare แบบ normalize)
        existing = set()
        for row in sheet_data[1:]:
            if row and len(row) > 0:
                existing.add(normalize_job_no(row[0]))

        new_added = 0
        updated = 0

        # ====== tab=13 ======
        for job in new_jobs:
            if not job or len(job) < 7:
                continue
            job_no = normalize_job_no(job[0])
            status = "ปิดงาน" if job_no in closed_job_nos else "รอแจ้ง"

            if job_no not in existing:
                try:
                    sheet.append_row(job + [status], value_input_option="USER_ENTERED")
                    print(f"✅ Added (tab13): {job_no} -> {status}")
                    new_added += 1
                    existing.add(job_no)
                    time.sleep(0.5)
                except Exception as e:
                    print(f"❌ Error adding job {job_no} from tab13: {e}")
            elif status == "ปิดงาน":
                try:
                    for i, row in enumerate(sheet_data[1:], start=2):
                        if row and len(row) > 0 and normalize_job_no(row[0]) == job_no:
                            if len(row) < 8 or row[7] != "ปิดงาน":
                                sheet.update_cell(i, 8, "ปิดงาน")
                                print(f"🔒 Updated status (tab13 closed): {job_no}")
                                updated += 1
                                time.sleep(0.5)
                            break
                except Exception as e:
                    print(f"❌ Error updating job {job_no} from tab13: {e}")

        # ====== tab=14 ======
        for job in waiting_jobs:
            if not job or len(job) < 7:
                continue
            job_no = normalize_job_no(job[0])
            if job_no not in existing:
                try:
                    sheet.append_row(job + ["รอแจ้ง"], value_input_option="USER_ENTERED")
                    print(f"✅ Added (tab14): {job_no} -> รอแจ้ง")
                    new_added += 1
                    existing.add(job_no)
                    time.sleep(0.5)
                except Exception as e:
                    print(f"❌ Error adding job {job_no} from tab14: {e}")

        # ====== tab=15 ======
        for job in closed_jobs_full:
            if not job or len(job) < 7:
                continue
            job_no = normalize_job_no(job[0])
            job_for_sheet = adjust_cols_for_sheet(job)  # ✅ ใช้เฉพาะ tab=15

            if job_no not in existing:
                try:
                    print("DEBUG (tab15) ->", job_for_sheet + ["ปิดงาน"])
                    sheet.append_row(job_for_sheet + ["ปิดงาน"], value_input_option="USER_ENTERED")
                    print(f"✅ Added (tab15): {job_no} -> ปิดงาน")
                    new_added += 1
                    existing.add(job_no)
                    time.sleep(0.5)
                except Exception as e:
                    print(f"❌ Error adding job {job_no} from tab15: {e}")
            else:
                try:
                    for i, row in enumerate(sheet_data[1:], start=2):
                        if row and len(row) > 0 and normalize_job_no(row[0]) == job_no:
                            if len(row) < 8 or row[7] != "ปิดงาน":
                                sheet.update_cell(i, 8, "ปิดงาน")
                                print(f"🔒 Updated status (tab15 exists): {job_no}")
                                updated += 1
                                time.sleep(0.5)
                            break
                except Exception as e:
                    print(f"❌ Error updating existing job {job_no} from tab15: {e}")

        # ====== tab=16 (งานที่ปิดแล้ว) ======
        for job in closed_already_jobs:
            if not job or len(job) < 7:
                continue
            job_no = normalize_job_no(job[0])  # parser ของ tab=16 ตัด '/' แล้วใน display
            if job_no not in existing:
                try:
                    print("DEBUG (tab16) ->", job + ["งานที่ปิดแล้ว"])
                    sheet.append_row(job + ["งานที่ปิดแล้ว"], value_input_option="USER_ENTERED")
                    print(f"✅ Added (tab16): {job_no} -> งานที่ปิดแล้ว")
                    new_added += 1
                    existing.add(job_no)
                    time.sleep(0.5)
                except Exception as e:
                    print(f"❌ Error adding job {job_no} from tab16: {e}")

        print(f"📊 Summary: {new_added} new rows added, {updated} rows updated")
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

        closed_already_jobs = fetch_jobs_by_tab(driver, 16)  # ⬅️ ใหม่: งานที่ปิดแล้ว

        # ของเดิม
        new_jobs = fetch_new_jobs(driver)            # tab=13 (เดิม)
        closed_job_nos = fetch_closed_jobs(driver)   # tab=15 (set of job_no for update status)

        # ใหม่: ดึงข้อมูลเต็มจาก tab=14 และ tab=15 (เพื่อ 'เติมแถว' ถ้ายังไม่เคยมี)
        waiting_jobs = fetch_jobs_by_tab(driver, 14)  # เพิ่มใหม่ถ้าไม่พบ → สถานะ 'รอแจ้ง'
        closed_jobs_full = fetch_jobs_by_tab(driver, 15)  # เพิ่มใหม่ถ้าไม่พบ → สถานะ 'ปิดงาน'

        sheet = setup_google_sheets()
        result = update_google_sheets(
            sheet,
            new_jobs=new_jobs,
            closed_job_nos=closed_job_nos,
            waiting_jobs=waiting_jobs,
            closed_jobs_full=closed_jobs_full,
            closed_already_jobs=closed_already_jobs  # ⬅️ ใหม่
        )

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
