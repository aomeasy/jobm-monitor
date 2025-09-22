import sys
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import json
import os
from google.auth import default
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
import requests
from io import StringIO
import time
from datetime import datetime

# --- Configuration ---
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1uEbsT3PZ8tdwiU1Xga_hS6uPve2H74xD5wUci0EcT0Q/edit?gid=0#gid=0"
GOOGLE_SHEET_NAME = "ชีต1"
TELEGRAM_BOT_TOKEN = "7978005713:AAHoMsNl_cyT3SkKLDq139YuTzGAnStfl4"
TELEGRAM_CHAT_ID = "8028926248"

# ✅ Sequential Tab Processing Logic
PROCESSING_ORDER = [
    {
        'tab': 13,
        'name': 'งานใหม่_แจ้งศูนย์อื่น',
        'action': 'ADD_NEW',  # เพิ่มใหม่ถ้าไม่พบ
        'status': 'งานใหม่'
    },
    {
        'tab': 14,
        'name': 'อยู่ระหว่างดำเนินการ_แจ้งศูนย์อื่น',
        'action': 'ADD_NEW',  # เพิ่มใหม่ถ้าไม่พบ
        'status': 'อยู่ระหว่างดำเนินการ'
    },
    {
        'tab': 15,
        'name': 'ปิดงานรอตรวจสอบ_แจ้งศูนย์อื่น',
        'action': 'UPDATE_STATUS',  # อัปเดตสถานะเป็น "ปิดงาน"
        'status': 'ปิดงาน'
    }
]

def test_google_sheet_authentication():
    """
    ทดสอบการเชื่อมต่อ Google Sheet และแสดงข้อมูลการตรวจสอบ
    """
    print("🔍 กำลังทดสอบ Google Sheet Authentication...")
    
    try:
        # ตรวจสอบว่ามี environment variable หรือไม่
        if not os.environ.get('GOOGLE_APPLICATION_CREDENTIALS'):
            print("❌ GOOGLE_APPLICATION_CREDENTIALS environment variable ไม่ได้ตั้งค่า")
            return False
            
        creds_file = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS')
        print(f"✅ พบ credentials file: {creds_file}")
        
        # ตรวจสอบไฟล์มีอยู่หรือไม่
        if not os.path.exists(creds_file):
            print(f"❌ ไฟล์ credentials ไม่มีอยู่: {creds_file}")
            return False
            
        print("✅ ไฟล์ credentials พบแล้ว")
        
        # อ่านและตรวจสอบ JSON structure
        try:
            with open(creds_file, 'r') as f:
                creds_data = json.load(f)
                
            # ตรวจสอบ required fields
            required_fields = ['type', 'project_id', 'private_key_id', 'private_key', 'client_email', 'client_id']
            missing_fields = [field for field in required_fields if field not in creds_data]
            
            if missing_fields:
                print(f"❌ ขาดข้อมูลสำคัญใน JSON: {missing_fields}")
                return False
                
            print("✅ JSON structure ถูกต้อง")
            print(f"📧 Service Account Email: {creds_data['client_email']}")
            print(f"🆔 Project ID: {creds_data['project_id']}")
            print(f"🔑 Private Key ID: {creds_data['private_key_id'][:10]}...")
            
        except json.JSONDecodeError as e:
            print(f"❌ ไฟล์ JSON มีปัญหา: {e}")
            return False
        except Exception as e:
            print(f"❌ อ่านไฟล์ credentials ไม่ได้: {e}")
            return False
            
        # ทดสอบสร้าง credentials
        try:
            credentials = Credentials.from_service_account_file(creds_file)
            print("✅ สร้าง credentials สำเร็จ")
        except Exception as e:
            print(f"❌ สร้าง credentials ไม่ได้: {e}")
            return False
            
        # ทดสอบการเชื่อมต่อ gspread
        try:
            client = gspread.authorize(credentials)
            print("✅ เชื่อมต่อ gspread สำเร็จ")
        except Exception as e:
            print(f"❌ เชื่อมต่อ gspread ไม่ได้: {e}")
            print("💡 ตรวจสอบว่าเปิดใช้งาน Google Sheets API แล้วหรือไม่")
            return False
            
        # ทดสอบการเข้าถึง spreadsheet (ถ้ามี SHEET_ID)
        sheet_id = os.environ.get('GOOGLE_SHEET_ID', 'YOUR_SHEET_ID_HERE')  # แก้ไขเป็น ID ของคุณ
        if sheet_id and sheet_id != 'YOUR_SHEET_ID_HERE':
            try:
                spreadsheet = client.open_by_key(sheet_id)
                print(f"✅ เข้าถึง spreadsheet สำเร็จ: {spreadsheet.title}")
                
                # ลองดู worksheets
                worksheets = spreadsheet.worksheets()
                print(f"📊 พบ {len(worksheets)} worksheets:")
                for ws in worksheets:
                    print(f"   - {ws.title} ({ws.row_count}x{ws.col_count})")
                    
            except gspread.SpreadsheetNotFound:
                print(f"❌ ไม่พบ spreadsheet ID: {sheet_id}")
                print("💡 ตรวจสอบว่า spreadsheet แชร์ให้ service account แล้วหรือไม่")
                return False
            except gspread.APIError as e:
                print(f"❌ Google Sheets API Error: {e}")
                return False
            except Exception as e:
                print(f"❌ เข้าถึง spreadsheet ไม่ได้: {e}")
                return False
                
        print("🎉 Google Sheet Authentication ทำงานปกติ!")
        return True
        
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดไม่ทราบสาเหตุ: {e}")
        return False

def get_existing_data_from_sheet():
    """
    ดึงข้อมูลเดิมจาก Google Sheet พร้อมการตรวจสอบที่ดีขึ้น
    """
    print("📊 กำลังดึงข้อมูลเดิมจาก Google Sheet...")
    
    try:
        # ทดสอบ authentication ก่อน
        if not test_google_sheet_authentication():
            print("❌ Google Sheet Authentication ล้มเหลว - ข้ามการดึงข้อมูล")
            return set()
            
        # ดำเนินการดึงข้อมูลตามปกติ
        credentials = Credentials.from_service_account_file(
            os.environ.get('GOOGLE_APPLICATION_CREDENTIALS')
        )
        client = gspread.authorize(credentials)
        
        # แทนที่ด้วย Sheet ID ของคุณ
        sheet_id = os.environ.get('GOOGLE_SHEET_ID', 'YOUR_SHEET_ID_HERE')
        if sheet_id == 'YOUR_SHEET_ID_HERE':
            print("⚠️ ยังไม่ได้ตั้งค่า GOOGLE_SHEET_ID environment variable")
            return set()
            
        spreadsheet = client.open_by_key(sheet_id)
        worksheet = spreadsheet.sheet1  # หรือระบุชื่อ worksheet
        
        # ดึงข้อมูลทั้งหมด
        records = worksheet.get_all_records()
        existing_job_nos = set()
        
        for record in records:
            if 'Job No.' in record and record['Job No.']:
                existing_job_nos.add(record['Job No.'].strip())
                
        print(f"✅ ดึงข้อมูลจาก Google Sheet สำเร็จ: {len(existing_job_nos)} รายการ")
        return existing_job_nos
        
    except gspread.SpreadsheetNotFound:
        print("❌ ไม่พบ spreadsheet - ตรวจสอบ Sheet ID และการแชร์")
        return set()
    except gspread.APIError as e:
        print(f"❌ Google Sheets API Error: {e}")
        if 'INVALID_ARGUMENT' in str(e):
            print("💡 อาจเป็นปัญหาการ encoding - ตรวจสอบ base64 encoding")
        return set()
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการดึงข้อมูลจาก Google Sheet: {e}")
        return set()

# เพิ่มในส่วนต้นของ main() function
def main():
    print("🚀 เริ่มต้นโปรแกรม JobM Monitor")
    print(f"⏰ เวลาเริ่มต้น: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # ทดสอบ Google Sheet Authentication ก่อน
    print("\n" + "="*50)
    print("🔐 ทดสอบ Google Sheet Authentication")
    print("="*50)
    
    auth_success = test_google_sheet_authentication()
    
    print("\n" + "="*50)
    print("🔄 เริ่มการตรวจสอบ JobM")  
    print("="*50)
    
    if not auth_success:
        print("⚠️ Google Sheet Authentication ล้มเหลว - ระบบจะทำงานต่อแต่ไม่สามารถบันทึกข้อมูลได้")
        print("💡 ตรวจสอب:")
        print("   1. base64 encoding ของ service account JSON")
        print("   2. Google Sheets API เปิดใช้งานแล้วหรือไม่")
        print("   3. Google Sheet แชร์ให้ service account แล้วหรือไม่")

        
def send_telegram_message(message):
    """Sends a message to the specified Telegram chat."""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("ข้อผิดพลาด: ไม่ได้ตั้งค่า TELEGRAM_BOT_TOKEN หรือ TELEGRAM_CHAT_ID")
        return False

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "parse_mode": "HTML"
    }
    try:
        response = requests.post(url, data=payload, timeout=10)
        response.raise_for_status()
        print(f"✅ ส่งข้อความ Telegram สำเร็จ")
        return True
    except requests.exceptions.RequestException as e:
        print(f"❌ เกิดข้อผิดพลาดในการส่งข้อความ Telegram: {e}")
        return False

def init_webdriver():
    """Initialize Chrome WebDriver with proper configuration"""
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        print("✅ WebDriver เริ่มทำงานแล้ว")
        return driver
    except WebDriverException as e:
        print(f"❌ เกิดข้อผิดพลาดในการเริ่มต้น WebDriver: {e}")
        sys.exit(1)

def login_to_jobm(driver, username, password, login_url):
    """Login to JobM system"""
    print("🔐 กำลังพยายาม Login...")
    try:
        driver.get(login_url)
        print(f"📄 เข้าถึงหน้า Login: {driver.current_url}")

        username_field = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.NAME, "username")) 
        )
        password_field = WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.NAME, "password"))
        )
        login_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))
        )

        username_field.clear()
        username_field.send_keys(username)
        password_field.clear()
        password_field.send_keys(password)
        
        print("📝 กรอกข้อมูล Login แล้ว กำลังกดปุ่ม Login...")
        login_button.click()

        WebDriverWait(driver, 20).until(
            EC.url_contains("jobManagement/pages/index")
        )
        print("✅ Login สำเร็จ!")
        return True
    except (NoSuchElementException, TimeoutException) as e:
        print(f"❌ Login ล้มเหลว: {e}")
        driver.save_screenshot("login_failure.png")
        return False
    except Exception as e:
        print(f"❌ Login ล้มเหลว: ข้อผิดพลาดที่ไม่คาดคิด: {e}")
        driver.save_screenshot("login_failure.png")
        return False

def fetch_data_from_tab(driver, base_url, tab_config):
    """ดึงข้อมูลจาก tab ที่ระบุ"""
    tab_number = tab_config['tab']
    tab_name = tab_config['name']
    data_url = f"{base_url}?tab={tab_number}"
    
    print(f"\n🔍 กำลังดึงข้อมูลจาก {tab_name} (tab={tab_number})")
    print(f"🌐 URL: {data_url}")
    
    try:
        driver.get(data_url)
        time.sleep(3)  # รอให้หน้าโหลด
        
        # รอให้ตารางโหลดเสร็จ
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CLASS_NAME, "table-bordered"))
        )
        print(f"✅ ตารางโหลดเสร็จแล้ว")

        page_source = driver.page_source
        
        # ดึงข้อมูลจาก HTML table
        tables = pd.read_html(StringIO(page_source), attrs={'class': 'table table-bordered table-striped'})
        
        if tables:
            df = tables[0]
            df.columns = df.columns.str.strip()
            
            # เพิ่มข้อมูล metadata
            df['Source_Tab'] = tab_number
            df['Source_Tab_Name'] = tab_name
            df['Status'] = tab_config['status']
            df['Last_Updated'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            print(f"✅ ดึงข้อมูลได้ {len(df)} แถวจาก {tab_name}")
            print(f"📋 คอลัมน์: {list(df.columns)}")
            
            # แสดงตัวอย่างข้อมูล
            if len(df) > 0 and 'Job No.' in df.columns:
                job_numbers = df['Job No.'].dropna().astype(str).str.strip().tolist()
                print(f"🏷️ Job No. ที่พบ: {job_numbers[:5]}{'...' if len(job_numbers) > 5 else ''}")
            
            return df
        else:
            print(f"❌ ไม่พบตารางใน {tab_name}")
            return pd.DataFrame()
            
    except TimeoutException:
        print(f"⏰ Timeout: ตารางโหลดไม่เสร็จใน {tab_name}")
        driver.save_screenshot(f"tab_{tab_number}_timeout.png")
        return pd.DataFrame()
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการดึงข้อมูลจาก {tab_name}: {e}")
        driver.save_screenshot(f"tab_{tab_number}_failure.png")
        return pd.DataFrame()

def get_all_sheet_data(spreadsheet_url, sheet_name):
    """ดึงข้อมูลทั้งหมดจาก Google Sheet"""
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds, project = default(scopes=scope)
        
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(spreadsheet_url)
        worksheet = spreadsheet.worksheet(sheet_name)
        
        all_values = worksheet.get_all_values()
        if not all_values:
            print("📊 Google Sheet ว่างเปล่า")
            return pd.DataFrame()

        headers = all_values[0]
        data = all_values[1:]
        df = pd.DataFrame(data, columns=headers)
        df.columns = df.columns.str.strip()
        print(f"📊 ดึงข้อมูลจาก Google Sheet '{sheet_name}': {len(df)} แถว")
        return df
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"❌ ไม่พบ Spreadsheet ที่ URL นี้")
        return pd.DataFrame()
    except gspread.exceptions.WorksheetNotFound:
        print(f"❌ ไม่พบ Sheet ชื่อ '{sheet_name}'")
        return pd.DataFrame()
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการดึงข้อมูลจาก Google Sheet: {e}")
        return pd.DataFrame()

def append_rows_to_google_sheet(rows_to_append_df, spreadsheet_url, sheet_name):
    """เพิ่มแถวใหม่ลงใน Google Sheet"""
    if rows_to_append_df.empty:
        print("⚠️ ไม่มีข้อมูลใหม่ที่จะเพิ่มลง Google Sheet")
        return False

    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds, project = default(scopes=scope)

        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(spreadsheet_url)
        worksheet = spreadsheet.worksheet(sheet_name)
        
        data_to_append = rows_to_append_df.values.tolist()
        worksheet.append_rows(data_to_append)
        print(f"✅ เพิ่ม {len(rows_to_append_df)} แถวใหม่ลง Google Sheet สำเร็จ")
        return True
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการเพิ่มข้อมูลลง Google Sheet: {e}")
        return False

def update_job_status_in_sheet(job_no, new_status, spreadsheet_url, sheet_name):
    """อัปเดตสถานะของ Job No. ใน Google Sheet"""
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds, project = default(scopes=scope)

        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(spreadsheet_url)
        worksheet = spreadsheet.worksheet(sheet_name)
        
        # หาแถวที่มี Job No. นี้
        all_values = worksheet.get_all_values()
        if not all_values:
            return False
        
        headers = all_values[0]
        
        # หาคอลัมน์ที่เกี่ยวข้อง
        job_no_col = None
        status_col = None
        last_updated_col = None
        
        for i, header in enumerate(headers):
            if 'Job No.' in header:
                job_no_col = i + 1  # gspread ใช้ 1-based indexing
            elif 'Status' in header:
                status_col = i + 1
            elif 'Last_Updated' in header:
                last_updated_col = i + 1
        
        if job_no_col is None:
            print(f"❌ ไม่พบคอลัมน์ 'Job No.' ใน Sheet")
            return False
        
        # ค้นหา Job No.
        for row_idx, row in enumerate(all_values[1:], start=2):  # เริ่มจากแถว 2
            if row_idx - 2 < len(row) and str(row[job_no_col - 1]).strip() == str(job_no).strip():
                # พบ Job No. แล้ว อัปเดตสถานะ
                if status_col:
                    worksheet.update_cell(row_idx, status_col, new_status)
                if last_updated_col:
                    worksheet.update_cell(row_idx, last_updated_col, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
                
                print(f"✅ อัปเดตสถานะ Job No. {job_no} เป็น '{new_status}' สำเร็จ")
                return True
        
        print(f"⚠️ ไม่พบ Job No. {job_no} ใน Google Sheet")
        return False
        
    except Exception as e:
        print(f"❌ เกิดข้อผิดพลาดในการอัปเดตสถานะ: {e}")
        return False

def process_tab_data(tab_data, tab_config, existing_job_numbers):
    """ประมวลผลข้อมูลจาก tab ตาม logic ที่กำหนด"""
    if tab_data.empty:
        print(f"⚠️ ไม่มีข้อมูลจาก {tab_config['name']}")
        return [], []
    
    tab_name = tab_config['name']
    action = tab_config['action']
    status = tab_config['status']
    
    new_jobs = []
    updated_jobs = []
    
    print(f"\n🔄 ประมวลผล {tab_name} ({action})...")
    
    for index, row in tab_data.iterrows():
        if 'Job No.' not in row.index or pd.isna(row['Job No.']):
            continue
            
        job_no = str(row['Job No.']).strip()
        if not job_no:
            continue
        
        if action == 'ADD_NEW':
            # Tab 13 & 14: เพิ่มใหม่ถ้าไม่พบ
            if job_no not in existing_job_numbers:
                new_jobs.append(row)
                print(f"🆕 {tab_name}: พบงานใหม่ {job_no}")
                
        elif action == 'UPDATE_STATUS':
            # Tab 15: ตรวจสอบและดำเนินการ
            if job_no in existing_job_numbers:
                # มีอยู่แล้ว → อัปเดตสถานะ
                updated_jobs.append({'job_no': job_no, 'new_status': status})
                print(f"🔄 {tab_name}: อัปเดตสถานะ {job_no} เป็น '{status}'")
            else:
                # ไม่มี → เพิ่มใหม่
                new_jobs.append(row)
                print(f"🆕 {tab_name}: เพิ่มงานใหม่ {job_no} (จากปิดงาน)")
                # เพิ่มลงใน existing_job_numbers เพื่อป้องกันการซ้ำในรอบเดียวกัน
                existing_job_numbers.add(job_no)
    
    return new_jobs, updated_jobs 

def generate_summary_report(results):
    """สร้างรายงานสรุปผลการทำงาน"""
    total_new = sum(len(result['new_jobs']) for result in results.values())
    total_updated = sum(len(result['updated_jobs']) for result in results.values())
    
    report = [
        f"📊 <b>รายงานสรุปการทำงาน</b>",
        f"⏰ เวลา: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"",
        f"🔍 <b>ผลการตรวจสอบแต่ละ Tab:</b>"
    ]
    
    for tab_config in PROCESSING_ORDER:
        tab_num = tab_config['tab']
        tab_name = tab_config['name']
        
        if tab_num in results:
            new_count = len(results[tab_num]['new_jobs'])
            updated_count = len(results[tab_num]['updated_jobs'])
            
            if tab_config['action'] == 'ADD_NEW':
                report.append(f"• Tab {tab_num} ({tab_name}): ➕ {new_count} งานใหม่")
            else:
                report.append(f"• Tab {tab_num} ({tab_name}): 🔄 {updated_count} งานปิด")
        else:
            report.append(f"• Tab {tab_num} ({tab_name}): ❌ ไม่สามารถดึงข้อมูลได้")
    
    report.extend([
        f"",
        f"📈 <b>สรุปรวม:</b>",
        f"🆕 งานใหม่ทั้งหมด: {total_new} รายการ",
        f"✅ งานที่ปิด: {total_updated} รายการ"
    ])
    
    return "\n".join(report)

# --- Main Logic ---
if __name__ == "__main__":
    USERNAME = "01000566"
    PASSWORD = "01000566"
    LOGIN_URL = "https://jobm.edoclite.com/jobManagement/pages/login"
    BASE_INDEX_URL = "https://jobm.edoclite.com/jobManagement/pages/index"

    print(f"🚀 เริ่มต้นโปรแกรม JobM Monitor")
    print(f"⏰ เวลาเริ่มต้น: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"📋 Tab ที่จะตรวจสอบ: {[config['tab'] for config in PROCESSING_ORDER]}")
    
    # สำหรับ GitHub Actions - รันครั้งเดียวแล้วจบ
    # สำหรับ Local/Railway - รันแบบ loop
    is_github_actions = 'GITHUB_ACTIONS' in os.environ
    
    if is_github_actions:
        print("🐙 รันผ่าน GitHub Actions - ทำงานครั้งเดียว")
        loop_count = 1
    else:
        print("🔄 รันแบบ Local - ทำงานแบบ loop")
        loop_count = float('inf')  # วนลูปไม่รู้จบ
    
    current_loop = 0
    while current_loop < loop_count:
        driver = None
        processing_results = {}
        
        try:
            print(f"\n{'='*60}")
            print(f"🔄 เริ่มรอบการทำงานใหม่: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"{'='*60}")
            
            # 1. เริ่มต้น WebDriver
            driver = init_webdriver()
            
            # 2. Login
            if not login_to_jobm(driver, USERNAME, PASSWORD, LOGIN_URL):
                print("❌ Login ล้มเหลว ข้ามรอบนี้")
                send_telegram_message("❌ JobM Monitor: Login ล้มเหลว")
                continue
            
            # 3. ดึงข้อมูลเดิมจาก Google Sheet
            print("\n📊 กำลังดึงข้อมูลเดิมจาก Google Sheet...")
            existing_data = get_all_sheet_data(GOOGLE_SHEET_URL, GOOGLE_SHEET_NAME)
            
            # สร้าง set ของ Job No. ที่มีอยู่แล้ว
            existing_job_numbers = set()
            if not existing_data.empty and 'Job No.' in existing_data.columns:
                existing_job_numbers = set(existing_data['Job No.'].astype(str).str.strip().tolist())
            print(f"📋 Job No. ที่มีอยู่แล้ว: {len(existing_job_numbers)} รายการ")
            
            # 4. ประมวลผลแต่ละ tab ตามลำดับ
            for tab_config in PROCESSING_ORDER:
                print(f"\n{'─'*40}")
                print(f"🔄 กำลังประมวลผล Tab {tab_config['tab']}: {tab_config['name']}")
                print(f"{'─'*40}")
                
                # ดึงข้อมูลจาก tab
                tab_data = fetch_data_from_tab(driver, BASE_INDEX_URL, tab_config)
                
                if not tab_data.empty:
                    # ประมวลผลตาม logic
                    new_jobs, updated_jobs = process_tab_data(tab_data, tab_config, existing_job_numbers)
                    
                processing_results[tab_config['tab']] = {
                    'new_jobs': new_jobs,
                    'updated_jobs': updated_jobs,
                    'tab_config': tab_config
                }
                
                # บันทึกงานใหม่ลง Google Sheet (ทั้ง ADD_NEW และ UPDATE_STATUS ที่ไม่พบ)
                if new_jobs:
                    new_jobs_df = pd.DataFrame(new_jobs)
                    
                    # เรียงคอลัมน์ให้ตรงกับ Sheet เดิม
                    if not existing_data.empty:
                        try:
                            new_jobs_ordered = new_jobs_df.reindex(columns=existing_data.columns, fill_value='')
                            append_rows_to_google_sheet(new_jobs_ordered, GOOGLE_SHEET_URL, GOOGLE_SHEET_NAME)
                        except Exception as e:
                            print(f"⚠️ ปัญหาการเรียงคอลัมน์: {e}")
                            append_rows_to_google_sheet(new_jobs_df, GOOGLE_SHEET_URL, GOOGLE_SHEET_NAME)
                    else:
                        # ถ้า Sheet ว่าง
                        header_df = pd.DataFrame([new_jobs_df.columns.tolist()])
                        combined_df = pd.concat([header_df, new_jobs_df], ignore_index=True)
                        append_rows_to_google_sheet(combined_df, GOOGLE_SHEET_URL, GOOGLE_SHEET_NAME)
                    
                    # อัปเดต existing_job_numbers
                    for job in new_jobs:
                        if 'Job No.' in job.index and pd.notna(job['Job No.']):
                            existing_job_numbers.add(str(job['Job No.']).strip())
                
                # อัปเดตสถานะงานที่ปิด
                if updated_jobs:
                    for update_info in updated_jobs:
                        update_job_status_in_sheet(
                            update_info['job_no'], 
                            update_info['new_status'],
                            GOOGLE_SHEET_URL, 
                            GOOGLE_SHEET_NAME
                        )
                else:
                    print(f"❌ ไม่สามารถดึงข้อมูลจาก Tab {tab_config['tab']} ได้")
                    processing_results[tab_config['tab']] = {
                        'new_jobs': [],
                        'updated_jobs': [],
                        'tab_config': tab_config,
                        'error': True
                    }
                
                # หน่วงเวลาระหว่าง tab
                time.sleep(3)
            
            # 5. ส่งรายงานสรุป
            summary_report = generate_summary_report(processing_results)
            print(f"\n{summary_report}")
            send_telegram_message(summary_report)
            
            print(f"\n✅ รอบการทำงานเสร็จสิ้น: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
        except KeyboardInterrupt:
            print(f"\n⏹️ หยุดการทำงานโดยผู้ใช้")
            break
        except Exception as e:
            print(f"❌ เกิดข้อผิดพลาดในรอบการทำงาน: {e}")
            error_message = f"🚨 JobM Monitor Error: {str(e)[:200]}..."
            send_telegram_message(error_message)
        finally:
            if driver:
                print("🔌 ปิด WebDriver")
                driver.quit()
        
        current_loop += 1
        
        # 6. พักก่อนรอบถัดไป (เฉพาะ local run)
        if not is_github_actions and current_loop < loop_count:
            print(f"\n⏰ พัก 1 ชั่วโมงก่อนรอบถัดไป...")
            print(f"⏰ รอบถัดไปเวลา: {datetime.fromtimestamp(time.time() + 3600).strftime('%Y-%m-%d %H:%M:%S')}")
            time.sleep(3600)
    
    if is_github_actions:
        print(f"✅ GitHub Actions job เสร็จสิ้น - รอ schedule ถัดไป")
    else:
        print(f"🔚 โปรแกรมสิ้นสุดการทำงาน")

 


