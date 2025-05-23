# komponen_scraper.py
import os
import pandas as pd
from utils import create_driver, login, generate_urls, collect_data, get_tanggal_input

input_tanggal = get_tanggal_input()
komponen_urls = generate_urls("laporan/komponen_cetak", input_tanggal)

driver = create_driver()

try:
    login(driver)
    komponen_data = collect_data(driver, komponen_urls, "KOMPONEN")

    os.makedirs("data", exist_ok=True)
    df_komponen = pd.DataFrame(komponen_data)
    df_komponen.to_excel(f"data/komponen-bagian-{input_tanggal}.xlsx", index=False, header=False)
    print(f"✅ Data KOMPONEN berhasil disimpan di 'data/komponen-bagian-{input_tanggal}.xlsx'")
finally:
    driver.quit()
    print("🚪 Browser ditutup.")
