import os
import time
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
DATA_TABLE_ID = os.getenv("DATA_TABLE_ID")


chrome_options = Options()
chrome_options.add_argument("--start-maximized")
service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service, options=chrome_options)

try:
    driver.get(WEB_URL)
    time.sleep(2)
    
    driver.find_element(By.NAME, "email").send_keys(EMAIL)
    driver.find_element(By.NAME, "password").send_keys(PASSWORD)
    driver.find_element(By.TAG_NAME, "form").submit()

    time.sleep(3)  

    
    table = driver.find_element(By.ID, DATA_TABLE_ID)
    rows = table.find_elements(By.TAG_NAME, "tr")

    data = []
    for row in rows:
        cols = row.find_elements(By.TAG_NAME, "td")
        if cols:  
            data.append([col.text for col in cols])

    
    df = pd.DataFrame(data)
    output_path = os.path.join(os.path.dirname(__file__), "data", "data-harian.xlsx")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df.to_excel(output_path, index=False, header=False)

    print(f"✅ Data berhasil disimpan ke: {output_path}")

finally:
    driver.quit()
