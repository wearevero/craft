# this file is used for read the file from master directory
# after read the file, it will be processed

import pandas as pd
import os
import glob
from pathlib import Path
import logging
import json
from datetime import datetime

class InvoiceImporter:
    def __init__(self, master_folder="master", log_folder="logs"):
        self.master_folder = master_folder
        self.log_folder = log_folder
        self.supported_formats = ['.xlsx', '.xls', '.xlsm']
        self.setup_logging()
        
    def setup_logging(self):
        """Setup logging untuk tracking import process"""
        os.makedirs(self.log_folder, exist_ok=True)
        log_file = os.path.join(self.log_folder, f"import_log_{datetime.now().strftime('%Y%m%d')}.log")
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def scan_master_directory(self):
        """Scan direktori master untuk menemukan semua file Excel"""
        self.logger.info(f"Scanning directory: {self.master_folder}")
        
        if not os.path.exists(self.master_folder):
            self.logger.error(f"Master folder '{self.master_folder}' tidak ditemukan!")
            return []
        
        found_files = []
        
        for ext in self.supported_formats:
            pattern = os.path.join(self.master_folder, f"*{ext}")
            files = glob.glob(pattern)
            found_files.extend(files)
        
        self.logger.info(f"Ditemukan {len(found_files)} file Excel")
        return found_files
    
    def get_file_info(self, file_path):
        """Mendapatkan informasi detail file"""
        try:
            stat = os.stat(file_path)
            return {
                'path': file_path,
                'name': os.path.basename(file_path),
                'size': stat.st_size,
                'modified': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                'extension': os.path.splitext(file_path)[1].lower()
            }
        except Exception as e:
            self.logger.error(f"Error getting file info for {file_path}: {e}")
            return None
    
    def validate_file_structure(self, file_path):
        """Validasi struktur file Excel"""
        try:
            # Baca file untuk pengecekan dasar
            if file_path.endswith('.xlsx') or file_path.endswith('.xlsm'):
                df = pd.read_excel(file_path, engine='openpyxl', header=None)
            else:
                df = pd.read_excel(file_path, engine='xlrd', header=None)
            
            # Check if file has minimum required data
            if df.empty:
                return False, "File kosong"
            
            # Convert to string for searching
            df_str = df.astype(str)
            
            # Check for required elements
            has_po = False
            has_unpaid_text = False
            
            for i in range(len(df_str)):
                for j in range(len(df_str.columns)):
                    cell_value = df_str.iloc[i, j].lower()
                    
                    if 'po#' in cell_value or 'po #' in cell_value:
                        has_po = True
                    
                    if 'all unpaid balance will be charged' in cell_value:
                        has_unpaid_text = True
            
            if not has_po:
                return False, "PO# tidak ditemukan dalam file"
            
            if not has_unpaid_text:
                return False, "Batas akhir data tidak ditemukan"
            
            return True, "File valid"
            
        except Exception as e:
            return False, f"Error validating file: {str(e)}"
    
    def create_file_catalog(self, files):
        """Membuat katalog file yang ditemukan"""
        catalog = {
            'scan_date': datetime.now().isoformat(),
            'total_files': len(files),
            'files': []
        }
        
        for file_path in files:
            file_info = self.get_file_info(file_path)
            if file_info:
                # Validate file
                is_valid, message = self.validate_file_structure(file_path)
                file_info['is_valid'] = is_valid
                file_info['validation_message'] = message
                
                catalog['files'].append(file_info)
                
                status = "VALID" if is_valid else "INVALID"
                self.logger.info(f"File: {file_info['name']} - {status} - {message}")
        
        return catalog
    
    def save_catalog(self, catalog):
        """Simpan katalog ke file JSON"""
        catalog_file = os.path.join(self.log_folder, f"file_catalog_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
        
        try:
            with open(catalog_file, 'w', encoding='utf-8') as f:
                json.dump(catalog, f, indent=2, ensure_ascii=False)
            
            self.logger.info(f"Katalog disimpan ke: {catalog_file}")
            return catalog_file
        except Exception as e:
            self.logger.error(f"Error saving catalog: {e}")
            return None
    
    def import_data(self):
        """Fungsi utama untuk import data"""
        self.logger.info("=== MEMULAI PROSES IMPORT ===")
        
        # 1. Scan directory
        files = self.scan_master_directory()
        
        if not files:
            self.logger.warning("Tidak ada file Excel ditemukan untuk diimport")
            return None
        
        # 2. Create catalog
        catalog = self.create_file_catalog(files)
        
        # 3. Save catalog
        catalog_file = self.save_catalog(catalog)
        
        # 4. Summary
        valid_files = [f for f in catalog['files'] if f['is_valid']]
        invalid_files = [f for f in catalog['files'] if not f['is_valid']]
        
        self.logger.info("=== RINGKASAN IMPORT ===")
        self.logger.info(f"Total file ditemukan: {len(catalog['files'])}")
        self.logger.info(f"File valid: {len(valid_files)}")
        self.logger.info(f"File invalid: {len(invalid_files)}")
        
        if invalid_files:
            self.logger.warning("File yang tidak valid:")
            for file in invalid_files:
                self.logger.warning(f"  - {file['name']}: {file['validation_message']}")
        
        return catalog
    
    def get_valid_files_list(self):
        """Mendapatkan daftar file yang valid untuk diproses"""
        files = self.scan_master_directory()
        valid_files = []
        
        for file_path in files:
            is_valid, _ = self.validate_file_structure(file_path)
            if is_valid:
                valid_files.append(file_path)
        
        return valid_files
    
    def read_specific_file(self, file_path):
        """Membaca file spesifik dan return dataframe"""
        try:
            self.logger.info(f"Reading file: {file_path}")
            
            if file_path.endswith('.xlsx') or file_path.endswith('.xlsm'):
                df = pd.read_excel(file_path, engine='openpyxl', header=None)
            else:
                df = pd.read_excel(file_path, engine='xlrd', header=None)
            
            self.logger.info(f"File berhasil dibaca: {df.shape[0]} baris, {df.shape[1]} kolom")
            return df
            
        except Exception as e:
            self.logger.error(f"Error reading file {file_path}: {e}")
            return None

def main():
    """Fungsi untuk testing import module"""
    importer = InvoiceImporter()
    
    print("Testing Invoice Importer...")
    catalog = importer.import_data()
    
    if catalog:
        print(f"\nHasil scan:")
        print(f"Total file: {catalog['total_files']}")
        valid_count = sum(1 for f in catalog['files'] if f['is_valid'])
        print(f"File valid: {valid_count}")
        print(f"File invalid: {catalog['total_files'] - valid_count}")

if __name__ == "__main__":
    main()