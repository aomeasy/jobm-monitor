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
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Configuration
GOOGLE_SHEET_URL = os.getenv('GOOGLE_SHEET_URL', 'https://docs.google.com/spreadsheets/d/1uEbsT3PZ8tdwiU1Xga_hS6uPve2H74xD5wUci0EcT0Q/edit?gid=0#gid=0')
GOOGLE_SHEET_NAME = os.getenv('GOOGLE_SHEET_NAME', 'ชีต1')
USERNAME = "01000566"
PASSWORD = "01000566"

def clean_html(cell):
    """Clean HTML content from cell"""
    try:
        return BeautifulSoup(cell.get_attribute("innerHTML").strip(), "html.parser").get_text(strip=True)
    except Exception as e:
        print(f"⚠️ Error cleaning HTML: {e}")
        return ""

def parse_row(row):
    """Parse table row and extract job data"""
    try:
        cols = row.find_elements(By.TAG_NAME, "td")
        if len(cols) < 8:
            return None
        return [clean_html(cols[i]) for i in range(1, 8)]
    except Exception as e:
        print(f"⚠️ Error parsing row: {e}")
        return None

def normalize_job_no(job_no):
    """Normalize job number for comparison"""
    return job_no.split("/")[0].strip().lower()

def setup_driver():
    """Setup Chrome WebDriver with GitHub Actions compatible options"""
    print("🔧 Setting up Chrome WebDriver...")
    options = Options()
    options.add_argument("--headless=new")  # Use new headless mode
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-web-security")
    options.add_argument("--allow-running-insecure-content")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-plugins")
    options.add_argument("--disable-images")
    options.add_argument("--disable-javascript")
    options.add_argument("--remote-debugging-port=9222")
    options.add_argument("--disable-background-timer-throttling")
    options.add_argument("--disable-backgrounding-occluded-windows")
    options.add_argument("--disable-renderer-backgrounding")
    options.add_argument("--disable-features=TranslateUI,BlinkGenPropertyTrees")
    options.add_argument("--disable-ipc-flooding-protection")
    options.add_argument("--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    # List of possible chromedriver paths
    possible_paths = [
        "/usr/local/bin/chromedriver",
        "/usr/bin/chromedriver", 
        "/snap/bin/chromium.chromedriver",
        "chromedriver"  # Let the system find it
    ]
    
    # Try multiple approaches to setup the driver
    for attempt in range(4):
        try:
            print(f"🔄 Attempt {attempt + 1}/4 to setup ChromeDriver...")
            service = None
            
            if attempt == 0:
                # Method 1: Try system chromedriver in known paths
                import shutil
                chromedriver_path = shutil.which('chromedriver')
                if chromedriver_path:
                    print(f"🔍 Found chromedriver via which: {chromedriver_path}")
                    service = Service(chromedriver_path)
                else:
                    # Try specific paths
                    for path in possible_paths[:-1]:  # Exclude the last one (system search)
                        import os
                        if os.path.exists(path) and os.access(path, os.X_OK):
                            print(f"🔍 Found chromedriver at: {path}")
                            service = Service(path)
                            break
                    
                if not service:
                    print("⚠️ System chromedriver not found in standard paths")
                    continue
                    
            elif attempt == 1:
                # Method 2: Try webdriver-manager (disable if causing issues)
                print("🔍 Trying webdriver-manager...")
                try:
                    from webdriver_manager.chrome import ChromeDriverManager
                    driver_path = ChromeDriverManager(cache_valid_range=7).install()
                    service = Service(driver_path)
                except Exception as wdm_error:
                    print(f"⚠️ WebDriver Manager failed: {wdm_error}")
                    continue
                    
            elif attempt == 2:
                # Method 3: Try without service (let Selenium auto-detect)
                print("🔍 Trying Chrome without explicit service...")
                service = None
                
            else:
                # Method 4: Try chromium-chromedriver (Ubuntu package)
                print("🔍 Trying chromium-chromedriver (Ubuntu package)...")
                chromium_driver = "/usr/bin/chromedriver"
                import os
                if os.path.exists(chromium_driver):
                    service = Service(chromium_driver)
                else:
                    print("⚠️ Chromium chromedriver not found")
                    continue
            
            # Create the driver
            if service:
                driver = webdriver.Chrome(service=service, options=options)
            else:
                driver = webdriver.Chrome(options=options)
            
            # Test if driver works
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            driver.get("about:blank")
            print("✅ Chrome WebDriver setup successful")
            return driver
            
        except Exception as e:
            print(f"⚠️ Attempt {attempt + 1} failed: {e}")
            if attempt == 3:  # Last attempt
                print(f"❌ All attempts failed. Last error: {e}")
                # Additional debugging info
                import platform
                import subprocess
                import os
                print(f"🔍 Platform: {platform.platform()}")
                print(f"🔍 Architecture: {platform.architecture()}")
                
                # Check if Chrome is installed
                try:
                    chrome_version = subprocess.check_output(['google-chrome', '--version'], text=True)
                    print(f"🔍 Chrome installed: {chrome_version.strip()}")
                except:
                    print("🔍 Chrome not found in system")
                
                # Check all possible ChromeDriver locations
                print("🔍 Checking ChromeDriver paths:")
                for path in possible_paths[:-1]:
                    if os.path.exists(path):
                        try:
                            result = subprocess.check_output([path, '--version'], text=True)
                            print(f"  ✅ {path}: {result.strip()}")
                        except:
                            print(f"  ❌ {path}: exists but not executable")
                    else:
                        print(f"  ❌ {path}: not found")
                
                # Check PATH
                import shutil
                which_result = shutil.which('chromedriver')
                if which_result:
                    print(f"🔍 chromedriver in PATH: {which_result}")
                else:
                    print("🔍 chromedriver not in PATH")
                
                raise Exception(f"Cannot setup ChromeDriver after 4 attempts. Last error: {e}")
            
            # Wait a bit before next attempt
            time.sleep(1)

def login_to_system(driver):
    """Login to the job management system"""
    try:
        print("🔐 Logging in...")
        driver.get("https://jobm.edoclite.com/jobManagement/pages/login")
        
        # Wait for login form
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "username"))
        )
        
        # Fill login credentials
        driver.find_element(By.NAME, "username").send_keys(USERNAME)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.NAME, "login__username").click()
        
        # Wait for successful login
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(., 'งานใหม่')]"))
        )
        print("✅ Login successful")
        return True
        
    except Exception as e:
        print(f"❌ Login failed: {e}")
        return False

def fetch_new_jobs(driver):
    """Fetch new jobs from the system"""
    try:
        print("📥 Fetching new jobs...")
        driver.get("https://jobm.edoclite.com/jobManagement/pages/index?tab=13")
        
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
        )
        
        new_rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        new_jobs = []
        
        for row in new_rows:
            parsed_row = parse_row(row)
            if parsed_row:
                new_jobs.append(parsed_row)
        
        print(f"📊 Found {len(new_jobs)} new jobs")
        return new_jobs
        
    except Exception as e:
        print(f"❌ Error fetching new jobs: {e}")
        return []

def fetch_closed_jobs(driver):
    """Fetch closed job numbers"""
    try:
        print("📦 Fetching closed jobs...")
        driver.get("https://jobm.edoclite.com/jobManagement/pages/index?tab=15")
        
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
        )
        
        closed_rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        closed_job_nos = set()
        
        for row in closed_rows:
            try:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 2:
                    job_no = normalize_job_no(clean_html(cols[1]))
                    if job_no:
                        closed_job_nos.add(job_no)
            except Exception as e:
                print(f"⚠️ Error parsing closed job row: {e}")
                continue
        
        print(f"📊 Found {len(closed_job_nos)} closed jobs")
        return closed_job_nos
        
    except Exception as e:
        print(f"❌ Error fetching closed jobs: {e}")
        return set()

def setup_google_sheets():
    """Setup Google Sheets connection"""
    try:
        print("📄 Connecting to Google Sheets...")
        
        # Load credentials from file (created by GitHub Actions)
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
        
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            "credentials.json", scope
        )
        client = gspread.authorize(creds)
        
        # Open the specific worksheet
        sheet = client.open_by_url(GOOGLE_SHEET_URL).worksheet(GOOGLE_SHEET_NAME)
        print("✅ Connected to Google Sheets")
        return sheet
        
    except Exception as e:
        print(f"❌ Error connecting to Google Sheets: {e}")
        raise

def update_google_sheets(sheet, new_jobs, closed_job_nos):
    """Update Google Sheets with new jobs and status updates"""
    try:
        print("✏️ Updating Google Sheets...")
        
        # Get existing sheet data
        sheet_data = sheet.get_all_values()
        if not sheet_data:
            # If sheet is empty, add headers
            headers = ["Job No", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Status"]
            sheet.append_row(headers)
            sheet_data = [headers]
        
        # Get existing job numbers from sheet
        sheet_job_nos = set()
        for row in sheet_data[1:]:  # Skip header row
            if row and len(row) > 0:
                sheet_job_nos.add(normalize_job_no(row[0]))
        
        new_added = 0
        updated = 0
        
        # Process each new job
        for job in new_jobs:
            if not job or len(job) < 7:
                continue
                
            job_no = normalize_job_no(job[0])
            status = "ปิดงาน" if job_no in closed_job_nos else "รอแจ้ง"
            
            # Add new job if not exists
            if job_no not in sheet_job_nos:
                try:
                    sheet.append_row(job + [status], value_input_option="USER_ENTERED")
                    print(f"✅ Added: {job_no}")
                    new_added += 1
                    time.sleep(1)  # Rate limiting
                except Exception as e:
                    print(f"❌ Error adding job {job_no}: {e}")
            
            # Update status if job is closed
            elif status == "ปิดงาน":
                try:
                    # Find the row to update
                    for i, row in enumerate(sheet_data[1:], start=2):
                        if row and len(row) > 0 and normalize_job_no(row[0]) == job_no:
                            # Check if status column exists and is different
                            if len(row) >= 8 and row[7] != "ปิดงาน":
                                sheet.update_cell(i, 8, "ปิดงาน")
                                print(f"🔒 Updated status: {job_no}")
                                updated += 1
                                time.sleep(1)  # Rate limiting
                            elif len(row) < 8:
                                # Add status column if it doesn't exist
                                sheet.update_cell(i, 8, "ปิดงาน")
                                print(f"🔒 Added status: {job_no}")
                                updated += 1
                                time.sleep(1)
                            break
                except Exception as e:
                    print(f"❌ Error updating job {job_no}: {e}")
        
        print(f"📊 Summary: {new_added} new jobs added, {updated} jobs updated")
        return {"new_added": new_added, "updated": updated}
        
    except Exception as e:
        print(f"❌ Error updating Google Sheets: {e}")
        return {"new_added": 0, "updated": 0, "error": str(e)}

def main():
    """Main execution function"""
    print(f"🚀 Starting job fetch process at {datetime.now()}")
    
    driver = None
    try:
        # Setup WebDriver
        driver = setup_driver()
        
        # Login to system
        if not login_to_system(driver):
            raise Exception("Login failed")
        
        # Fetch data
        new_jobs = fetch_new_jobs(driver)
        closed_job_nos = fetch_closed_jobs(driver)
        
        # Setup Google Sheets
        sheet = setup_google_sheets()
        
        # Update Google Sheets
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
