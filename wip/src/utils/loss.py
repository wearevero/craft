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
        raise


def highlight_rows(driver, rows, duration=1.5):
    """Highlight sekaligus semua baris <tr> yang ingin di-copy"""
    try:
        driver.execute_script(
            """
            const rows = arguments[0];
            for (const row of rows) {
                row.style.border = '2px solid red';
                row.style.backgroundColor = 'yellow';
            }
        """,
            rows,
        )

        time.sleep(duration)

        driver.execute_script(
            """
            const rows = arguments[0];
            for (const row of rows) {
                row.style.border = '';
                row.style.backgroundColor = '';
            }
        """,
            rows,
        )
    except Exception as e:
        print(f"⚠️ Gagal highlight rows: {e}")


def extract_table_data(driver):
    """Mengambil data dari tabel utama dan highlight semua baris yang akan diambil sekaligus"""
    print("🔍 Mencari tabel data...")

    selectors = [
        "td.judul > table > tbody",
        "table > tbody",
        ".judul table tbody",
        "table tbody",
    ]

    table = None
    for selector in selectors:
        try:
            table = driver.find_element(By.CSS_SELECTOR, selector)
            print(f"✅ Tabel ditemukan dengan selector: {selector}")
            break
        except:
            continue

    if not table:
        print("❌ Tidak dapat menemukan tabel dengan selector apapun")

        print("📄 Struktur HTML halaman:")
        print(driver.page_source[:1000])
        return []

    try:
        rows = table.find_elements(By.TAG_NAME, "tr")
        print(f"📊 Total baris ditemukan: {len(rows)}")

        if len(rows) <= 2:
            print("⚠️ Tabel kosong atau hanya ada header")
            return []

        data_rows = rows[1:-1]
        print(f"📈 Baris data yang akan diambil: {len(data_rows)}")

        if data_rows:
            highlight_rows(driver, data_rows, duration=1.5)

        data = []
        for i, row in enumerate(data_rows):
            cols = row.find_elements(By.TAG_NAME, "td")
            if cols:
                row_data = [col.text.strip() for col in cols]
                data.append(row_data)
                print(f"  Baris {i+1}: {row_data}")

        print(f"✅ Total data berhasil diambil: {len(data)} baris")
        return data

    except Exception as e:
        print(f"❌ Gagal ekstrak data: {e}")
        return []


def collect_data(driver, urls_dict, jenis):
    all_data = []
    original_tab = driver.current_window_handle

    print(f"🚀 Memulai pengumpulan data {jenis} dari {len(urls_dict)} bagian...")

    for i, (bagian, url) in enumerate(urls_dict.items()):
        print(f"\n📥 [{i+1}/{len(urls_dict)}] Mengambil data {jenis} - {bagian}...")
        print(f"🔗 URL: {url}")

        if i == 0:
            driver.get(url)
        else:
            driver.execute_script("window.open('');")
            driver.switch_to.window(driver.window_handles[-1])
            driver.get(url)

        time.sleep(2)

        try:

            time.sleep(1)
            current_url = driver.current_url
            print(f"📍 Current URL: {current_url}")

            if "error" in driver.page_source.lower() or "404" in driver.title:
                print(f"⚠️ Halaman error untuk bagian {bagian}")
                continue

        except Exception as e:
            print(f"⚠️ Gagal memuat halaman untuk {bagian}: {e}")
            continue

        data = extract_table_data(driver)

        if data:
            for row in data:
                all_data.append([bagian] + row)
            print(f"✅ Berhasil mengambil {len(data)} baris dari {bagian}")
        else:
            print(f"⚠️ Tidak ada data dari {bagian}")

    driver.switch_to.window(original_tab)
    print(f"\n🎯 Total data terkumpul: {len(all_data)} baris")
    return all_data


def generate_urls(base_path, tanggal):
    bagian_mapping = {
        "CUTTING 2": 7,
        "TAMBAH PART": 125,
        "RE-CASTING": 101,
        "REPAIR PART": 126,
        "ILCA": 103,
        "STRIPING": 9,
        "PENDING": 15,
        "PERBAIKAN": 16,
        "RANGKAI 1": 12,
        "SEGONG REPAIR": 2,
        "FILLLING1": 13,
        "FILLING 2": 17,
        "POLISHING 1": 100,
        "POLISHING 2": 11,
        "POLISHING CVD": 122,
    }

    urls = {}
    for bagian, bagian_id in bagian_mapping.items():
        urls[bagian] = f"{WEB_URL}/{base_path}?d={tanggal}&s=&b={bagian_id}&m=all"
    return urls


def get_tanggal_input():
    """Meminta input tanggal dari user, default ke hari ini"""
    input_tanggal = input(
        "📅 Masukkan tanggal (format: YYYY-MM-DD) [default: hari ini]: "
    ).strip()
    if input_tanggal == "":
        input_tanggal = datetime.today().strftime("%Y-%m-%d")
    return input_tanggal
