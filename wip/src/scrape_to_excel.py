import os
import time
from datetime import datetime
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

load_dotenv()
WEB_URL = os.getenv("WEB_URL")
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
CHROMEDRIVER_PATH = os.getenv("CHROMEDRIVER_PATH")


input_tanggal = input("📅 Masukkan tanggal (format: YYYY-MM-DD) [default: hari ini]: ").strip()
if input_tanggal == "":
    input_tanggal = datetime.today().strftime("%Y-%m-%d")


bagian_kode = {
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


def generate_urls(base_path, tanggal):
    urls = {}
    for nama_bagian, kode in bagian_kode.items():
        url = f"{WEB_URL}/{base_path}?d={tanggal}&s=&b={kode}&m=all"
        urls[nama_bagian] = url
    return urls


loss_urls = generate_urls("laporan/loss_bagian_cetak", input_tanggal)
komponen_urls = generate_urls("+/komponen_cetak", input_tanggal)


chrome_options = Options()
chrome_options.add_argument("--start-maximized")
service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service, options=chrome_options)


def login():
    print("🔐 Login ke sistem...")
    driver.get(WEB_URL)
    time.sleep(2)
    driver.find_element(By.NAME, "email").send_keys(EMAIL)
    driver.find_element(By.NAME, "password").send_keys(PASSWORD)
    driver.find_element(By.TAG_NAME, "form").submit()
    time.sleep(2)
    print("✅ Login berhasil.")


def extract_table_data():
    try:
        table = driver.find_element(By.CSS_SELECTOR, "td.judul > table > tbody")
        rows = table.find_elements(By.TAG_NAME, "tr")
        data_rows = rows[1:-1]  
        data = []
        for row in data_rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if cols:
                data.append([col.text.strip() for col in cols])
        return data
    except Exception as e:
        print(f"⚠️  Gagal ekstrak data: {e}")
        return []


def collect_data(urls_dict, jenis):
    all_data = []
    for bagian, url in urls_dict.items():
        print(f"📥 Mengambil data {jenis} - {bagian}...")
        driver.get(url)
        time.sleep(1.5)
        data = extract_table_data()
        for row in data:
            all_data.append([bagian] + row)  
    return all_data


try:
    login()

    
    loss_data = collect_data(loss_urls, "LOSS")
    komponen_data = collect_data(komponen_urls, "KOMPONEN")

    
    os.makedirs("data", exist_ok=True)

    df_loss = pd.DataFrame(loss_data)
    df_loss.to_excel(f"data/loss-bagian-{input_tanggal}.xlsx", index=False, header=False)

    df_komponen = pd.DataFrame(komponen_data)
    df_komponen.to_excel(f"data/komponen-bagian-{input_tanggal}.xlsx", index=False, header=False)

    print(f"✅ Semua data berhasil disimpan di folder 'data/' untuk tanggal {input_tanggal}.")

finally:
    driver.quit()
    print("🚪 Browser ditutup.")
