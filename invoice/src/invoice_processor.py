# invoice_processor.py
import pandas as pd
import re
from pathlib import Path
from import_data import InvoiceImporter

class InvoiceProcessor:
    def __init__(self, template_path="template/template.xls"):
        self.importer = InvoiceImporter()
        self.template_path = template_path

    def process_single_file(self, file_path):
        df = self.importer.read_specific_file(file_path)
        if df is None:
            return None

        df = df.fillna('').astype(str)

        # Temukan baris awal (berisi PO#)
        start_row, end_row = None, None
        for i in range(len(df)):
            if any("po#" in str(cell).lower() for cell in df.iloc[i]):
                start_row = i
            if any("all unpaid balance will be charged" in str(cell).lower() for cell in df.iloc[i]):
                end_row = i
                break

        if start_row is None or end_row is None:
            return None  # Tidak bisa diproses

        df_section = df.iloc[start_row:end_row].copy()

        # Hapus baris yang mengandung teks "Buyer No" hingga "Cust Ref"
        pattern1 = re.compile(r"buyer no|cust ref", re.IGNORECASE)
        df_section = df_section[~df_section.apply(lambda row: row.astype(str).str.contains(pattern1).any(), axis=1)]

        # Hapus baris dari "Dia w’t" sampai sebelum baris yang mengandung "maklon"
        drop_flag = False
        cleaned_rows = []
        for _, row in df_section.iterrows():
            text = " ".join(row.astype(str)).lower()
            if "dia w" in text:
                drop_flag = True
                continue
            if "maklon" in text:
                drop_flag = False
            if not drop_flag:
                cleaned_rows.append(row)

        df_clean = pd.DataFrame(cleaned_rows)

        # Ambil 7 kolom yang diinginkan (jika ditemukan berdasarkan header)
        header_row = df_clean[df_clean.apply(lambda row: row.astype(str).str.contains("po#|item|metal|qty|w't|maklon|total", case=False).any(), axis=1)].index.min()

        if header_row is None:
            return None

        df_data = df_clean.iloc[header_row + 1:].copy()
        df_data.columns = df_clean.iloc[header_row]

        # Hanya ambil kolom berikut
        expected_columns = ['PO#', 'Item', 'No.', 'Metal', "Q'ty", "Total w't", 'maklon', 'total']
        df_final = df_data[[col for col in expected_columns if col in df_data.columns]]

        # Simpan ke template
        output_path = Path("template") / f"processed_{Path(file_path).stem}.xlsx"
        df_final.to_excel(output_path, index=False)
        return str(output_path)

    def process_all_files(self):
        valid_files = self.importer.get_valid_files_list()
        results = []
        for file in valid_files:
            result = self.process_single_file(file)
            if result:
                print(f"Processed: {file} => {result}")
                results.append(result)
            else:
                print(f"Skipped: {file}")
        return results
