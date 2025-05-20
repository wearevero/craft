import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
from dotenv import load_dotenv

load_dotenv()

TARGET_URL = os.getenv("TARGET_URL")
DATA_TABLE_ID = os.getenv("DATA_TABLE_ID")
CHROMEDRIVER_PATH = os.getenv("CHROMEDRIVER_PATH")

chrome_options = Options()
chrome_options.add_argument("--start-maximized")

service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service, options=chrome_options)

try:
    
    driver.get(TARGET_URL)
    time.sleep(2)

    table_element = driver.find_element(By.ID, DATA_TABLE_ID)
    table_text = table_element.text
    print(table_text)

    current_dir = os.path.dirname(__file__)            
    parent_dir = os.path.dirname(current_dir)          
    data_dir = os.path.join(parent_dir, "data")        

    os.makedirs(data_dir, exist_ok=True)

    output_file = os.path.join(data_dir, "data-harian.txt")
    with open(output_file, "w", encoding="utf-8") as file:
        file.write(table_text)

    print(f"Hasil disimpan di: {output_file}")

finally:
    driver.quit()
