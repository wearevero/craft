import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from dotenv import load_dotenv

load_dotenv()
WEB_URL = os.getenv("WEB_URL")
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
CHROMEDRIVER_PATH = os.getenv("CHROMEDRIVER_PATH")

def create_driver():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    service = Service(CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def login(driver):
    print("🔐 Login ke sistem...")
    driver.get(WEB_URL)
    time.sleep(2)

    try:
        driver.find_element(By.NAME, "email").send_keys(EMAIL)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.TAG_NAME, "form").submit()
        time.sleep(2)
        print("✅ Login berhasil.")
    except Exception as e:
        print(f"❌ Gagal login: {e}")

def highlight_rows(driver, rows, duration=1.5):
    """Highlight sekaligus semua baris <tr> yang ingin di-copy"""
    driver.execute_script("""
        const rows = arguments[0];
        for (const row of rows) {
            row.style.border = '2px solid red';
            row.style.backgroundColor = '
        }
    """, rows)

    time.sleep(duration)

    driver.execute_script("""
        const rows = arguments[0];
        for (const row of rows) {
            row.style.border = '';
            row.style.backgroundColor = '';
        }
    """, rows)

def extract_table_data(driver):
    """Mengambil data dari tabel utama dan highlight semua baris yang akan diambil sekaligus"""
    try:
        table = driver.find_element(By.CSS_SELECTOR, "td.judul > table > tbody")
        rows = table.find_elements(By.TAG_NAME, "tr")
        data_rows = rows[1:-1]  

        
        highlight_rows(driver, data_rows, duration=1.5)

        data = []
        for row in data_rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if cols:
                data.append([col.text.strip() for col in cols])
        return data
    except Exception as e:
        print(f"⚠️ Gagal ekstrak data: {e}")
        return []

def collect_data(driver, urls_dict, jenis):
    all_data = []
    original_tab = driver.current_window_handle  
    
    for i, (bagian, url) in enumerate(urls_dict.items()):
        print(f"📥 Mengambil data {jenis} - {bagian}...")

        
        if i == 0:
            driver.get(url)
        else:
            driver.execute_script("window.open('');")
            driver.switch_to.window(driver.window_handles[-1])
            driver.get(url)

        time.sleep(1.5)

        try:
            table = driver.find_element(By.CSS_SELECTOR, "td.judul > table > tbody")
        except Exception as e:
            print(f"⚠️ Tidak dapat menemukan tabel: {e}")

        data = extract_table_data(driver)
        for row in data:
            all_data.append([bagian] + row)

    driver.switch_to.window(original_tab)

    return all_data

def generate_urls(base_path, tanggal):
    bagian_mapping = {
        "cutting 2": 7,
        "tambah part": 125,
        "recasting": 101,
        "repair part": 126,
        "ilca": 103,
        "striping": 9,
        "pending": 15,
        "perbaikan": 16,
        "rangkai 1": 12,
        "segong repair": 2,
        "fillling1": 13,
        "filling 2": 17,
        "polishing 1": 100,
        "polishing 2": 11,
        "polishing cvd": 122,
    }

    urls = {}
    for bagian, bagian_id in bagian_mapping.items():
        urls[bagian] = (
            f"{WEB_URL}/{base_path}?d={tanggal}&s=&b={bagian_id}&m=all"
        )
    return urls

def get_tanggal_input():
    """Meminta input tanggal dari user, default ke hari ini"""
    input_tanggal = input("📅 Masukkan tanggal (format: YYYY-MM-DD) [default: hari ini]: ").strip()
    if input_tanggal == "":
        input_tanggal = datetime.today().strftime("%Y-%m-%d")
    return input_tanggal
