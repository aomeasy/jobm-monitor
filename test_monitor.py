#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JobM Monitor Testing Script
ทดสอบความพร้อมของระบบก่อนรันโปรแกรมจริง
"""

import sys
import os
import pandas as pd
import gspread
import requests
from google.auth import default
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
from io import StringIO
import time
from datetime import datetime
import json

# --- Configuration ---
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/176UvX_WAHQvWtbqIdvZu4Wo91U_6YG3GuBSFjt8IDYk/edit#gid=687082847"
GOOGLE_SHEET_NAME = "ปิดงาน"
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")

# Test Configuration
TEST_TABS = [13, 14, 15]
TAB_NAMES = {
    13: "งานใหม่_แจ้งศูนย์อื่น",
    14: "อยู่ระหว่างดำเนินการ_แจ้งศูนย์อื่น", 
    15: "ปิดงานรอตรวจสอบ_แจ้งศูนย์อื่น"
}

USERNAME = "01000566"
PASSWORD = "01000566"
LOGIN_URL = "https://jobm.edoclite.com/jobManagement/pages/login"
BASE_INDEX_URL = "https://jobm.edoclite.com/jobManagement/pages/index"

class TestResult:
    """คลาสสำหรับเก็บผลการทดสอบ"""
    def __init__(self, name):
        self.name = name
        self.success = False
        self.details = {}
        self.error = None
        self.start_time = time.time()
        self.end_time = None
    
    def set_success(self, details=None):
        self.success = True
        self.details = details or {}
        self.end_time = time.time()
    
    def set_failure(self, error):
        self.success = False
        self.error = str(error)
        self.end_time = time.time()
    
    @property
    def duration(self):
        if self.end_time:
            return round(self.end_time - self.start_time, 2)
        return 0

def print_header():
    """แสดงหัวเรื่องการทดสอบ"""
    print("=" * 70)
    print("🧪 JobM Monitor System Testing")
    print("=" * 70)
    print(f"⏰ เริ่มการทดสอบ: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"🌍 Environment: {'GitHub Actions' if 'GITHUB_ACTIONS' in os.environ else 'Local'}")
    print("-" * 70)

def send_telegram_message(message):
    """ส่งข้อความ Telegram"""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        return False, "Token or Chat ID not configured"

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "parse_mode": "HTML"
    }
    try:
        response = requests.post(url, data=payload, timeout=10)
        response.raise_for_status()
        return True, "Success"
    except Exception as e:
        return False, str(e)

def test_environment():
    """ทดสอบ Environment และ Dependencies"""
    result = TestResult("Environment Check")
    
    try:
        print("🔧 ทดสอบ Environment...")
        
        # ตรวจสอบ Python version
        python_version = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
        
        # ตรวจสอบ required packages
        required_packages = ['pandas', 'gspread', 'selenium', 'requests']
        missing_packages = []
        
        for package in required_packages:
            try:
                __import__(package)
            except ImportError:
                missing_packages.append(package)
        
        # ตรวจสอบ environment variables
        env_vars = {
            'TELEGRAM_BOT_TOKEN': bool(TELEGRAM_BOT_TOKEN),
            'TELEGRAM_CHAT_ID': bool(TELEGRAM_CHAT_ID),
            'GOOGLE_APPLICATION_CREDENTIALS': bool(os.getenv('GOOGLE_APPLICATION_CREDENTIALS')),
            'GITHUB_ACTIONS': bool(os.getenv('GITHUB_ACTIONS'))
        }
        
        details = {
            'python_version': python_version,
            'missing_packages': missing_packages,
            'environment_variables': env_vars
        }
        
        if missing_packages:
            result.set_failure(f"Missing packages: {missing_packages}")
        else:
            result.set_success(details)
            print(f"   ✅ Python {python_version}")
            print(f"   ✅ All packages installed")
            print(f"   ✅ Environment variables: {sum(env_vars.values())}/4 configured")
        
    except Exception as e:
        result.set_failure(e)
        print(f"   ❌ Environment check failed: {e}")
    
    return result

def test_webdriver():
    """ทดสอบ WebDriver"""
    result = TestResult("WebDriver")
    driver = None
    
    try:
        print("🌐 ทดสอบ WebDriver...")
        
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        
        # ทดสอบเปิดหน้าเว็บง่ายๆ
        driver.get("https://www.google.com")
        
        result.set_success({
            'chrome_version': driver.capabilities.get('browserVersion', 'Unknown'),
            'chromedriver_version': driver.capabilities.get('chrome', {}).get('chromedriverVersion', 'Unknown')
        })
        print(f"   ✅ Chrome: {driver.capabilities.get('browserVersion', 'Unknown')}")
        print(f"   ✅ ChromeDriver: {driver.capabilities.get('chrome', {}).get('chromedriverVersion', 'Unknown')}")
        
    except Exception as e:
        result.set_failure(e)
        print(f"   ❌ WebDriver failed: {e}")
    finally:
        if driver:
            driver.quit()
    
    return result

def test_login(driver=None):
    """ทดสอบการ Login"""
    result = TestResult("Login")
    should_close_driver = False
    
    if driver is None:
        should_close_driver = True
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--headless")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
        except Exception as e:
            result.set_failure(f"Cannot create WebDriver: {e}")
            return result
    
    try:
        print("🔐 ทดสอบการ Login...")
        
        # เปิดหน้า login
        driver.get(LOGIN_URL)
        
        # รอ elements โหลด
        username_field = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.NAME, "username"))
        )
        password_field = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.NAME, "password"))
        )
        login_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))
        )

        # กรอกข้อมูลและ login
        username_field.clear()
        username_field.send_keys(USERNAME)
        password_field.clear()
        password_field.send_keys(PASSWORD)
        login_button.click()

        # รอ redirect
        WebDriverWait(driver, 20).until(
            EC.url_contains("jobManagement/pages/index")
        )
        
        current_url = driver.current_url
        result.set_success({
            'login_url': LOGIN_URL,
            'redirect_url': current_url,
            'username': USERNAME
        })
        print(f"   ✅ Login successful")
        print(f"   ✅ Redirected to: {current_url}")
        
    except TimeoutException:
        result.set_failure("Login timeout - elements not found or redirect failed")
        print(f"   ❌ Login timeout")
    except Exception as e:
        result.set_failure(e)
        print(f"   ❌ Login failed: {e}")
    finally:
        if should_close_driver and driver:
            driver.quit()
    
    return result

def test_tab_data_extraction():
    """ทดสอบการดึงข้อมูลจากแต่ละ tab"""
    results = {}
    driver = None
    
    try:
        # สร้าง WebDriver
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        
        # Login ก่อน
        login_result = test_login(driver)
        if not login_result.success:
            for tab_num in TEST_TABS:
                results[f"tab_{tab_num}"] = TestResult(f"Tab {tab_num}")
                results[f"tab_{tab_num}"].set_failure("Cannot login")
            return results
        
        # ทดสอบแต่ละ tab
        for tab_number in TEST_TABS:
            result = TestResult(f"Tab {tab_number}")
            tab_name = TAB_NAMES.get(tab_number, f"Tab_{tab_number}")
            
            try:
                print(f"🔍 ทดสอบ {tab_name} (tab={tab_number})...")
                
                data_url = f"{BASE_INDEX_URL}?tab={tab_number}"
                driver.get(data_url)
                time.sleep(3)
                
                # รอให้ตารางโหลด
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "table-bordered"))
                )
                
                page_source = driver.page_source
                tables = pd.read_html(StringIO(page_source), 
                                   attrs={'class': 'table table-bordered table-striped'})
                
                if tables:
                    df = tables[0]
                    df.columns = df.columns.str.strip()
                    
                    # ตรวจสอบคอลัมน์ที่สำคัญ
                    required_columns = ['Job No.']
                    missing_columns = [col for col in required_columns if col not in df.columns]
                    
                    # ดึงข้อมูลตัวอย่าง
                    sample_data = {}
                    if len(df) > 0:
                        for col in df.columns[:5]:  # แสดงแค่ 5 คอลัมน์แรก
                            if len(df[col].dropna()) > 0:
                                sample_data[col] = str(df[col].dropna().iloc[0])[:50]  # จำกัด 50 ตัวอักษร
                    
                    details = {
                        'row_count': len(df),
                        'column_count': len(df.columns),
                        'columns': list(df.columns),
                        'missing_columns': missing_columns,
                        'sample_data': sample_data,
                        'url': data_url
                    }
                    
                    if missing_columns:
                        result.set_failure(f"Missing required columns: {missing_columns}")
                        print(f"   ⚠️ Missing columns: {missing_columns}")
                    else:
                        result.set_success(details)
                        print(f"   ✅ Data found: {len(df)} rows, {len(df.columns)} columns")
                        
                        # แสดง Job No. ตัวอย่าง
                        if 'Job No.' in df.columns and len(df) > 0:
                            job_numbers = df['Job No.'].dropna().astype(str).head(3).tolist()
                            print(f"   📋 Sample Job No.: {job_numbers}")
                else:
                    result.set_failure("No table found")
                    print(f"   ❌ No table found")
                    
            except TimeoutException:
                result.set_failure("Timeout waiting for table")
                print(f"   ⏰ Timeout waiting for table")
            except Exception as e:
                result.set_failure(str(e))
                print(f"   ❌ Error: {e}")
            
            results[f"tab_{tab_number}"] = result
            time.sleep(2)  # หน่วงเวลาระหว่าง tab
            
    except Exception as e:
        print(f"❌ Critical error in tab testing: {e}")
        for tab_num in TEST_TABS:
            if f"tab_{tab_num}" not in results:
                result = TestResult(f"Tab {tab_num}")
                result.set_failure(f"Critical error: {e}")
                results[f"tab_{tab_num}"] = result
    finally:
        if driver:
            driver.quit()
    
    return results

def test_google_sheet_connection():
    """ทดสอบการเชื่อมต่อ Google Sheet"""
    result = TestResult("Google Sheets")
    
    try:
        print("📊 ทดสอบการเชื่อมต่อ Google Sheet...")
        
        scope = ["https://spreadsheets.google.com/feeds", 
                "https://www.googleapis.com/auth/drive"]
        creds, project = default(scopes=scope)
        
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(GOOGLE_SHEET_URL)
        worksheet = spreadsheet.worksheet(GOOGLE_SHEET_NAME)
        
        # ทดสอบอ่านข้อมูล
        all_values = worksheet.get_all_values()
        row_count = len(all_values)
        
        # ทดสอบเขียนข้อมูล (แถวทดสอบ)
        test_row = [f"TEST_{datetime.now().strftime('%Y%m%d_%H%M%S')}", "Test Data"]
        worksheet.append_row(test_row)
        
        # ลบแถวทดสอบ
        all_values_after = worksheet.get_all_values()
        if len(all_values_after) > row_count:
            worksheet.delete_rows(len(all_values_after))
        
        details = {
            'spreadsheet_title': spreadsheet.title,
            'worksheet_title': worksheet.title,
            'row_count': row_count,
            'column_count': len(all_values[0]) if all_values else 0,
            'spreadsheet_id': spreadsheet.id,
            'project_id': project
        }
        
        result.set_success(details)
        print(f"   ✅ Connected to: {spreadsheet.title}")
        print(f"   ✅ Worksheet: {worksheet.title}")
        print(f"   ✅ Data: {row_count} rows")
        print(f"   ✅ Read/Write test: Passed")
        
    except Exception as e:
        result.set_failure(e)
        print(f"   ❌ Google Sheets error: {e}")
    
    return result

def test_telegram_notification():
    """ทดสอบการส่งการแจ้งเตือน Telegram"""
    result = TestResult("Telegram")
    
    try:
        print("📱 ทดสอบการส่งข้อความ Telegram...")
        
        if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
            result.set_failure("Telegram credentials not configured")
            print(f"   ❌ Token or Chat ID not configured")
            return result
        
        test_message = (
            f"🧪 <b>JobM Monitor Test</b>\n"
            f"⏰ เวลา: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"🤖 ระบบทดสอบการส่งข้อความ\n"
            f"✅ หากเห็นข้อความนี้แสดงว่า Telegram ทำงานปกติ"
        )
        
        success, message = send_telegram_message(test_message)
        
        if success:
            result.set_success({
                'bot_token_length': len(TELEGRAM_BOT_TOKEN),
                'chat_id': TELEGRAM_CHAT_ID,
                'message_sent': True
            })
            print(f"   ✅ Message sent successfully")
            print(f"   ✅ Chat ID: {TELEGRAM_CHAT_ID}")
        else:
            result.set_failure(f"Failed to send message: {message}")
            print(f"   ❌ Failed to send: {message}")
        
    except Exception as e:
        result.set_failure(e)
        print(f"   ❌ Telegram error: {e}")
    
    return result

def generate_test_report(test_results):
    """สร้างรายงานผลการทดสอบ"""
    total_tests = len(test_results)
    passed_tests = sum(1 for result in test_results.values() if result.success)
    failed_tests = total_tests - passed_tests
    success_rate = (passed_tests / total_tests * 100) if total_tests > 0 else 0
    
    report = {
        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'overall_status': 'PASS' if failed_tests == 0 else 'FAIL',
        'summary': {
            'total_tests': total_tests,
            'passed_tests': passed_tests,
            'failed_tests': failed_tests,
            'success_rate': f"{success_rate:.1f}%"
        },
        'tests': {},
        'total_duration': sum(result.duration for result in test_results.values())
    }
    
    for name, result in test_results.items():
        report['tests'][name] = {
            'success': result.success,
            'duration': result.duration,
            'details': result.details,
            'error': result.error
        }
    
    return report

def print_test_report(report):
    """แสดงรายงานการทดสอบ"""
    print("\n" + "=" * 70)
    print("📋 รายงานผลการทดสอบ JobM Monitor")
    print("=" * 70)
    print(f"⏰ เวลา: {report['timestamp']}")
    print(f"📊 สถานะรวม: {'✅ PASS' if report['overall_status'] == 'PASS' else '❌ FAIL'}")
    print(f"📈 ผลสรุป: {report['summary']['passed_tests']}/{report['summary']['total_tests']} ({report['summary']['success_rate']})")
    print(f"⏱️  เวลารวม: {report['total_duration']:.2f} วินาทเ")
    
    print(f"\n🔍 รายละเอียดการทดสอบ:")
    print("-" * 70)
    
    for test_name, test_result in report['tests'].items():
        status_icon = "✅" if test_result['success'] else "❌"
        print(f"{status_icon} {test_name} ({test_result['duration']:.2f}s)")
        
        if test_result['success'] and test_result['details']:
            details = test_result['details']
            if 'row_count' in details:
                print(f"   📊 จำนวนแถว: {details['row_count']}")
            if 'column_count' in details:
                print(f"   📋 จำนวนคอลัมน์: {details['column_count']}")
            if 'missing_columns' in details and details['missing_columns']:
                print(f"   ⚠️ คอลัมน์ที่ขาด: {details['missing_columns']}")
        
        if not test_result['success'] and test_result['error']:
            print(f"   ❌ ข้อผิดพลาด: {test_result['error']}")
        
        print()

def main():
    """ฟังก์ชันหลักสำหรับการทดสอบ"""
    print_header()
    
    # รันการทดสอบทั้งหมด
    test_results = {}
    
    # 1. ทดสอบ Environment
    test_results['environment'] = test_environment()
    
    # 2. ทดสอบ WebDriver
    test_results['webdriver'] = test_webdriver()
    
    # 3. ทดสอบ Google Sheets
    test_results['google_sheets'] = test_google_sheet_connection()
    
    # 4. ทดสอบ Telegram
    test_results['telegram'] = test_telegram_notification()
    
    # 5. ทดสอบการดึงข้อมูลจาก tabs (ถ้า WebDriver ทำงาน)
    if test_results['webdriver'].success:
        tab_results = test_tab_data_extraction()
        test_results.update(tab_results)
    else:
        print("⚠️ ข้าม tab testing เนื่องจาก WebDriver ไม่ทำงาน")
        for tab_num in TEST_TABS:
            result = TestResult(f"Tab {tab_num}")
            result.set_failure("WebDriver not available")
            test_results[f"tab_{tab_num}"] = result
    
    # สร้างและแสดงรายงาน
    report = generate_test_report(test_results)
    print_test_report(report)
    
    # ส่งรายงานผ่าน Telegram (ถ้าทำงาน)
    if test_results['telegram'].success:
        summary_message = (
            f"🧪 <b>รายงานการทดสอบ JobM Monitor</b>\n"
            f"⏰ เวลา: {report['timestamp']}\n"
            f"📊 สถานะ: {'✅ PASS' if report['overall_status'] == 'PASS' else '❌ FAIL'}\n"
            f"📈 ผลสรุป: {report['summary']['passed_tests']}/{report['summary']['total_tests']} "
            f"({report['summary']['success_rate']})\n"
            f"⏱️ เวลารวม: {report['total_duration']:.2f}s\n\n"
            f"🔍 รายละเอียด:\n"
        )
        
        for test_name, test_result in test_results.items():
            status_icon = "✅" if test_result.success else "❌"
            summary_message += f"• {status_icon} {test_name}\n"
        
        send_telegram_message(summary_message)
        print("📤 ส่งรายงานผ่าน Telegram แล้ว")
    
    # บันทึกรายงานเป็นไฟล์ (ถ้าเป็น local)
    if not os.getenv('GITHUB_ACTIONS'):
        try:
            filename = f"test_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(report, f, ensure_ascii=False, indent=2)
            print(f"💾 บันทึกรายงานที่: {filename}")
        except Exception as e:
            print(f"⚠️ ไม่สามารถบันทึกไฟล์รายงาน: {e}")
    
    print("\n🏁 การทดสอบเสร็จสิ้น")
    
    # Exit code สำหรับ CI/CD
    if report['overall_status'] != 'PASS':
        print("❌ มีการทดสอบล้มเหลว")
        sys.exit(1)
    else:
        print("✅ การทดสอบทั้งหมดผ่าน")
        sys.exit(0)

if __name__ == "__main__":
    main()
