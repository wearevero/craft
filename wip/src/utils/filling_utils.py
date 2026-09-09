import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from dotenv import load_dotenv
from rich.console import Console

load_dotenv()

console = Console()
WEB_URL = os.getenv("WEB_URL")
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
CHROMEDRIVER_PATH = os.getenv("CHROMEDRIVER_PATH")


def create_driver():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
    service = Service(CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver


def login(driver):
    console.print("[dim]› Logging in...[/dim]")
    driver.get(WEB_URL)
    time.sleep(2)

    try:
        driver.find_element(By.NAME, "email").send_keys(EMAIL)
        driver.find_element(By.NAME, "password").send_keys(PASSWORD)
        driver.find_element(By.TAG_NAME, "form").submit()
        time.sleep(2)
        console.print("[green]✓[/green] Logged in")
        console.print()
    except Exception as e:
        console.print("[red]✗[/red] Login failed")
        console.print(f"[dim]  {e}[/dim]")
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
        console.print(f"[dim]  (skip) highlight failed: {e}[/dim]")


def extract_table_data(driver):
    """Mengambil data dari tabel utama dan highlight semua baris yang akan diambil sekaligus"""
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
            break
        except:
            continue

    if not table:
        return []

    try:
        rows = table.find_elements(By.TAG_NAME, "tr")

        if len(rows) <= 2:
            return []

        data_rows = rows[1:-1]

        if data_rows:
            highlight_rows(driver, data_rows, duration=1.5)

        data = []
        for row in data_rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if cols:
                row_data = [col.text.strip() for col in cols]
                data.append(row_data)

        return data

    except Exception as e:
        console.print(f"[red]✗[/red] [dim]Failed to extract data: {e}[/dim]")
        return []


def collect_data(driver, urls_dict, jenis):
    all_data = []
    original_tab = driver.current_window_handle
    total_bagian = len(urls_dict)

    console.print(f"[bold white]{jenis}[/bold white] [dim]· {total_bagian} sections[/dim]")
    console.print()

    for i, (bagian, url) in enumerate(urls_dict.items()):
        if i == 0:
            driver.get(url)
        else:
            driver.execute_script("window.open('');")
            driver.switch_to.window(driver.window_handles[-1])
            driver.get(url)

        time.sleep(2)

        try:
            time.sleep(1)
            if "error" in driver.page_source.lower() or "404" in driver.title:
                console.print(f"  [yellow]‼[/yellow] {bagian:<20} [dim]page error[/dim]")
                continue
        except Exception:
            console.print(f"  [red]✗[/red] {bagian:<20} [dim]failed to load[/dim]")
            continue

        data = extract_table_data(driver)

        if data:
            for row in data:
                all_data.append([bagian] + row)
            console.print(f"  [green]✓[/green] {bagian:<20} [dim]{len(data)} rows[/dim]")
        else:
            console.print(f"  [dim]·[/dim] [dim]{bagian:<20} no data[/dim]")

    driver.switch_to.window(original_tab)
    console.print()

    if all_data:
        console.print(f"[dim]total[/dim]  [bold]{len(all_data)}[/bold] rows")
    else:
        console.print(f"[dim]no data collected for {jenis}[/dim]")
    console.print()

    return all_data


def generate_urls(base_path, tanggal):
    bagian_mapping = {
        "FILLING TEM": 8,
    }

    urls = {}
    for bagian, bagian_id in bagian_mapping.items():
        urls[bagian] = f"{WEB_URL}/{base_path}?d={tanggal}&s=&b={bagian_id}&m=all"
    return urls


def get_tanggal_input():
    """Meminta input tanggal dari user, default ke hari ini"""
    input_tanggal = input("date (YYYY-MM-DD, default: today) › ").strip()
    if input_tanggal == "":
        input_tanggal = datetime.today().strftime("%Y-%m-%d")
    return input_tanggal