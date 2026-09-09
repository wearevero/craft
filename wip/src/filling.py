import os
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import numbers
from rich.console import Console
from utils.filling_utils import create_driver, login, generate_urls, collect_data, get_tanggal_input

# Folder tujuan penyimpanan file. Bisa dioverride lewat variabel OUTPUT_DIR di .env
OUTPUT_DIR = os.getenv("OUTPUT_DIR", r"D:\Work\CORE\WIP\KRIDO\WIP TEMPLATE")
console = Console()


def save_to_excel_with_number_format(data, filename):
    """Menyimpan data ke Excel dengan format number yang benar sesuai format asli dari web"""
    if not data:
        return

    wb = Workbook()
    ws = wb.active

    for row in data:
        ws.append(row)

    for row_idx in range(1, ws.max_row + 1):
        for col_idx in range(2, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)

            if cell.value and str(cell.value).strip():
                cell_value_str = str(cell.value).strip()

                if is_numeric_value(cell_value_str):
                    cleaned_value = cell_value_str.replace(',', '')

                    try:
                        numeric_value = float(cleaned_value)
                        cell.value = numeric_value

                        if '.' in cell_value_str:
                            decimal_places = len(cell_value_str.split('.')[-1])

                            if decimal_places == 1:
                                cell.number_format = '0.0'
                            elif decimal_places == 2:
                                cell.number_format = '0.00'
                            elif decimal_places == 3:
                                cell.number_format = '0.000'
                            else:
                                cell.number_format = '0.' + '0' * decimal_places
                        else:
                            if numeric_value == int(numeric_value):
                                cell.number_format = '0'
                            else:
                                cell.number_format = '0.00'

                    except (ValueError, TypeError):
                        pass

    # Simpan file
    wb.save(filename)


def is_numeric_value(value_str):
    """Mengecek apakah string adalah nilai numerik"""
    cleaned = value_str.replace(',', '')

    pattern = r'^-?\d+\.?\d*$'

    return bool(re.match(pattern, cleaned))


def main():
    console.print()
    console.print("[bold white]Filling TEM[/bold white] [dim]· data scraper[/dim]")
    console.print()

    input_tanggal = get_tanggal_input()
    console.print(f"[dim]date[/dim]  {input_tanggal}")
    console.print()

    filling_urls = generate_urls("laporan/loss_bagian_cetak", input_tanggal)
    driver = create_driver()

    try:
        login(driver)
        filling_data = collect_data(driver, filling_urls, "FILLING TEM")

        if not filling_data:
            console.print("[yellow]✗[/yellow] No data found")
            console.print("[dim]  · date may have no records[/dim]")
            console.print("[dim]  · page structure may have changed[/dim]")
            console.print("[dim]  · connection timeout or failed to load[/dim]")
            console.print()
            return

        os.makedirs(OUTPUT_DIR, exist_ok=True)
        filename = os.path.join(OUTPUT_DIR, f"filling-{input_tanggal}.xlsx")

        save_to_excel_with_number_format(filling_data, filename)

        if os.path.exists(filename):
            file_size_kb = os.path.getsize(filename) / 1024
            df_verifikasi = pd.read_excel(filename, header=None)

            console.print("[green]✓[/green] Saved")
            console.print(f"[dim]  file[/dim]  {os.path.basename(filename)}")
            console.print(f"[dim]  rows[/dim]  {len(df_verifikasi)}")
            console.print(f"[dim]  size[/dim]  {file_size_kb:.1f} KB")
            console.print()
        else:
            console.print("[red]✗[/red] Failed to save file")
            console.print(f"[dim]  {filename}[/dim]")
            console.print()

    except Exception as e:
        console.print("[red]✗[/red] Something went wrong")
        console.print(f"[dim]  {e}[/dim]")
        console.print()
        import traceback
        traceback.print_exc()

    finally:
        console.print("[dim]Browser left open for manual inspection.[/dim]")


if __name__ == "__main__":
    main()