# loss_scraper.py
import os
import pandas as pd
from utils import create_driver, login, generate_urls, collect_data, get_tanggal_input

input_tanggal = get_tanggal_input()
loss_urls = generate_urls("laporan/loss_bagian_cetak", input_tanggal)

driver = create_driver()

try:
    login(driver)
    loss_data = collect_data(driver, loss_urls, "LOSS")

    os.makedirs("data", exist_ok=True)
    df_loss = pd.DataFrame(loss_data)
    df_loss.to_excel(f"data/loss-bagian-{input_tanggal}.xlsx", index=False, header=False)
    print(f"✅ Data LOSS berhasil disimpan di 'data/loss-bagian-{input_tanggal}.xlsx'")
finally:
    driver.quit()
    print("🚪 Browser ditutup.")
