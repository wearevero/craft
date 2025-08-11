# import_data.py - Fixed version
# This file is used for reading files from master directory
# After reading the file, it will be processed

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
        self.ensure_directories()
        
    def ensure_directories(self):
        """Membuat direktori yang diperlukan jika belum ada"""
        os.makedirs(self.master_folder, exist_ok=True)
        os.makedirs(self.log_folder, exist_ok=True)
        
    def setup_logging(self):
        """Setup logging untuk tracking import process"""
        os.makedirs(self.log_folder, exist_ok=True)
        log_file = os.path.join(self.log_folder, f"import_log_{datetime.now().strftime('%Y%m%d')}.log")
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def scan_master_directory(self):
        """Scan direktori master untuk menemukan semua file Excel"""
        self.logger.info(f"Scanning directory: {self.master_folder}")
        
        if not os.path.exists(self.master_folder):
            self.logger.error(f"Master folder '{self.master_folder}' tidak ditemukan!")
            os.makedirs(self.master_folder, exist_ok=True)
            self.logger.info(f"Master folder '{self.master_folder}' telah dibuat")
            return []
        
        found_files = []
        
        for ext in self.supported_formats:
            pattern = os.path.join(self.master_folder, f"*{ext}")
            files = glob.glob(pattern)
            found_files.extend(files)
        
        # Sort files for consistent ordering
        found_files.sort()
        
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
                'size_mb': round(stat.st_size / (1024 * 1024), 2),
                'modified': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                'extension': os.path.splitext(file_path)[1].lower()
            }
        except Exception as e:
            self.logger.error(f"Error getting file info for {file_path}: {e}")
            return None
    
    def validate_file_structure(self, file_path):
        """Validasi struktur file Excel dengan pengecekan yang lebih robust"""
        try:
            self.logger.info(f"Validating file: {os.path.basename(file_path)}")
            
            # Coba baca file dengan berbagai engine
            df = None
            try:
                if file_path.endswith(('.xlsx', '.xlsm')):
                    df = pd.read_excel(file_path, engine='openpyxl', header=None)
                else:
                    df = pd.read_excel(file_path, engine='xlrd', header=None)
            except Exception as engine_error:
                self.logger.warning(f"Primary engine failed, trying alternative: {engine_error}")
                try:
                    df = pd.read_excel(file_path, header=None)
                except Exception as fallback_error:
                    return False, f"Cannot read file: {fallback_error}"
            
            # Check if file has minimum required data
            if df is None or df.empty:
                return False, "File kosong atau tidak dapat dibaca"
            
            # Check minimum dimensions
            if df.shape[0] < 5 or df.shape[1] < 3:
                return False, f"File terlalu kecil ({df.shape[0]} baris, {df.shape[1]} kolom)"
            
            # Convert to string for searching, handle NaN values
            df_str = df.fillna('').astype(str)
            
            # Check for required elements
            has_po = False
            has_data_indicators = False
            po_row = -1
            
            # Search for key indicators
            for i in range(min(20, len(df_str))):  # Only check first 20 rows for efficiency
                for j in range(len(df_str.columns)):
                    cell_value = str(df_str.iloc[i, j]).lower().strip()
                    
                    # Check for PO# header
                    if 'po#' in cell_value or 'po #' in cell_value:
                        has_po = True
                        po_row = i
                        self.logger.info(f"Found PO# at row {i}")
                    
                    # Check for other data indicators
                    if any(keyword in cell_value for keyword in ['item', 'metal', 'qty', 'quantity', 'weight']):
                        has_data_indicators = True
            
            # Additional validation for invoice-like structure
            if not has_po:
                return False, "PO# tidak ditemukan dalam file"
            
            if not has_data_indicators:
                return False, "Indikator data invoice tidak ditemukan"
            
            # Check if there's actual data after headers
            if po_row >= 0 and po_row < len(df_str) - 2:
                data_rows = 0
                for i in range(po_row + 1, min(po_row + 50, len(df_str))):
                    row_data = df_str.iloc[i].fillna('').astype(str)
                    if any(cell.strip() and cell.lower() not in ['nan', ''] for cell in row_data):
                        data_rows += 1
                
                if data_rows < 1:
                    return False, "Tidak ada data setelah header"
            
            return True, f"File valid - {df.shape[0]} baris, {df.shape[1]} kolom"
            
        except Exception as e:
            return False, f"Error validating file: {str(e)}"
    
    def create_file_catalog(self, files):
        """Membuat katalog file yang ditemukan"""
        catalog = {
            'scan_date': datetime.now().isoformat(),
            'total_files': len(files),
            'valid_files': 0,
            'invalid_files': 0,
            'total_size_mb': 0,
            'files': []
        }
        
        for file_path in files:
            file_info = self.get_file_info(file_path)
            if file_info:
                # Add to total size
                catalog['total_size_mb'] += file_info['size_mb']
                
                # Validate file
                is_valid, message = self.validate_file_structure(file_path)
                file_info['is_valid'] = is_valid
                file_info['validation_message'] = message
                
                if is_valid:
                    catalog['valid_files'] += 1
                else:
                    catalog['invalid_files'] += 1
                
                catalog['files'].append(file_info)
                
                status = "✓ VALID" if is_valid else "✗ INVALID"
                self.logger.info(f"File: {file_info['name']} - {status}")
                if not is_valid:
                    self.logger.warning(f"  Reason: {message}")
        
        catalog['total_size_mb'] = round(catalog['total_size_mb'], 2)
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
            self.logger.info(f"Letakkan file Excel (.xlsx, .xls, .xlsm) di folder: {os.path.abspath(self.master_folder)}")
            return None
        
        # 2. Create catalog
        catalog = self.create_file_catalog(files)
        
        # 3. Save catalog
        catalog_file = self.save_catalog(catalog)
        
        # 4. Summary
        self.logger.info("=== RINGKASAN IMPORT ===")
        self.logger.info(f"Total file ditemukan: {catalog['total_files']}")
        self.logger.info(f"File valid: {catalog['valid_files']}")
        self.logger.info(f"File invalid: {catalog['invalid_files']}")
        self.logger.info(f"Total ukuran file: {catalog['total_size_mb']} MB")
        
        if catalog['invalid_files'] > 0:
            self.logger.warning("File yang tidak valid:")
            for file in catalog['files']:
                if not file['is_valid']:
                    self.logger.warning(f"  - {file['name']}: {file['validation_message']}")
        
        return catalog
    
    def get_valid_files_list(self):
        """Mendapatkan daftar file yang valid untuk diproses"""
        files = self.scan_master_directory()
        valid_files = []
        
        for file_path in files:
            is_valid, message = self.validate_file_structure(file_path)
            if is_valid:
                valid_files.append(file_path)
                self.logger.info(f"Valid file added: {os.path.basename(file_path)}")
            else:
                self.logger.warning(f"Invalid file skipped: {os.path.basename(file_path)} - {message}")
        
        return valid_files
    
    def read_specific_file(self, file_path):
        """Membaca file spesifik dan return dataframe"""
        try:
            self.logger.info(f"Reading file: {os.path.basename(file_path)}")
            
            # Try different engines for compatibility
            df = None
            if file_path.endswith(('.xlsx', '.xlsm')):
                try:
                    df = pd.read_excel(file_path, engine='openpyxl', header=None)
                except Exception:
                    df = pd.read_excel(file_path, header=None)
            else:
                try:
                    df = pd.read_excel(file_path, engine='xlrd', header=None)
                except Exception:
                    df = pd.read_excel(file_path, header=None)
            
            if df is not None and not df.empty:
                self.logger.info(f"File berhasil dibaca: {df.shape[0]} baris, {df.shape[1]} kolom")
                return df
            else:
                self.logger.error(f"File kosong atau tidak dapat dibaca: {file_path}")
                return None
                
        except Exception as e:
            self.logger.error(f"Error reading file {file_path}: {e}")
            return None
    
    def get_file_preview(self, file_path, rows=10):
        """Mendapatkan preview file untuk debugging"""
        try:
            df = self.read_specific_file(file_path)
            if df is not None:
                preview = df.head(rows).fillna('').astype(str)
                return preview
            return None
        except Exception as e:
            self.logger.error(f"Error getting preview for {file_path}: {e}")
            return None

def main():
    """Fungsi untuk testing import module"""
    print("Testing Invoice Importer...")
    
    importer = InvoiceImporter()
    
    # Test import process
    catalog = importer.import_data()
    
    if catalog:
        print(f"\n=== HASIL SCAN ===")
        print(f"Total file: {catalog['total_files']}")
        print(f"File valid: {catalog['valid_files']}")
        print(f"File invalid: {catalog['invalid_files']}")
        print(f"Total ukuran: {catalog['total_size_mb']} MB")
        
        if catalog['files']:
            print(f"\nDetail file:")
            for file_info in catalog['files']:
                status = "✓" if file_info['is_valid'] else "✗"
                print(f"  {status} {file_info['name']} ({file_info['size_mb']} MB)")
        
        # Test getting valid files
        valid_files = importer.get_valid_files_list()
        print(f"\nFile yang dapat diproses: {len(valid_files)}")
    else:
        print("Tidak ada file ditemukan untuk diimport")
        print("Letakkan file Excel di folder 'master' dan coba lagi")

if __name__ == "__main__":
    main()