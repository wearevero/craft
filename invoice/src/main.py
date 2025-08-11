# main file for the invoice module
import sys
from pathlib import Path

# Import modul-modul yang diperlukan
try:
    from import_data import InvoiceImporter
    from export import InvoiceExporter
    from invoice_processor import InvoiceProcessor
except ImportError as e:
    print(f"Error importing modules: {e}")
    print("Pastikan semua file modul berada dalam direktori yang sama")
    sys.exit(1)

class InvoiceManager:
    def __init__(self):
        self.processor = InvoiceProcessor()
        self.importer = InvoiceImporter()
        self.exporter = InvoiceExporter()
        
    def setup_directories(self):
        """Membuat direktori yang diperlukan jika belum ada"""
        directories = ['master', 'template', 'output', 'logs']
        for directory in directories:
            Path(directory).mkdir(exist_ok=True)
        print("Direktori setup completed.")
    
    def show_menu(self):
        """Menampilkan menu pilihan"""
        print("\n" + "="*50)
        print("    INVOICE PROCESSING AUTOMATION")
        print("="*50)
        print("1. Process All Files (Otomatis)")
        print("2. Process Single File")
        print("3. Import Data from Master")
        print("4. Export Processed Data")
        print("5. Check Directory Status")
        print("6. Exit")
        print("="*50)
    
    def check_directory_status(self):
        """Mengecek status direktori dan file"""
        print("\n=== STATUS DIREKTORI ===")
        
        # Check master directory
        master_files = list(Path('master').glob('*.xls*'))
        print(f"Master Directory: {len(master_files)} file(s) ditemukan")
        for file in master_files[:5]:  # Show first 5 files
            print(f"  - {file.name}")
        if len(master_files) > 5:
            print(f"  ... dan {len(master_files) - 5} file lainnya")
        
        # Check template directory
        template_files = list(Path('template').glob('*'))
        print(f"Template Directory: {len(template_files)} file(s)")
        
        # Check output directory
        output_files = list(Path('output').glob('*'))
        print(f"Output Directory: {len(output_files)} file(s)")
    
    def process_all_files(self):
        """Memproses semua file secara otomatis"""
        print("\n=== MEMULAI PEMROSESAN OTOMATIS ===")
        try:
            self.processor.process_all_files()
        except Exception as e:
            print(f"Error during processing: {e}")
    
    def process_single_file(self):
        """Memproses file tunggal"""
        master_files = list(Path('master').glob('*.xls*'))
        
        if not master_files:
            print("Tidak ada file Excel ditemukan di folder master")
            return
        
        print("\n=== PILIH FILE UNTUK DIPROSES ===")
        for i, file in enumerate(master_files, 1):
            print(f"{i}. {file.name}")
        
        try:
            choice = int(input("Pilih nomor file: ")) - 1
            if 0 <= choice < len(master_files):
                selected_file = master_files[choice]
                print(f"\nMemproses file: {selected_file.name}")
                result = self.processor.process_single_file(str(selected_file))
                if result:
                    print(f"File berhasil diproses: {result}")
                else:
                    print("Gagal memproses file")
            else:
                print("Pilihan tidak valid")
        except ValueError:
            print("Input tidak valid. Masukkan nomor.")
        except Exception as e:
            print(f"Error: {e}")
    
    def run(self):
        """Menjalankan aplikasi utama"""
        self.setup_directories()
        
        while True:
            try:
                self.show_menu()
                choice = input("Pilih opsi (1-6): ").strip()
                
                if choice == '1':
                    self.process_all_files()
                elif choice == '2':
                    self.process_single_file()
                elif choice == '3':
                    print("Fitur import akan segera tersedia...")
                    self.importer.import_data()
                elif choice == '4':
                    print("Fitur export akan segera tersedia...")
                    self.exporter.export_data()
                elif choice == '5':
                    self.check_directory_status()
                elif choice == '6':
                    print("Terima kasih! Program selesai.")
                    break
                else:
                    print("Pilihan tidak valid. Pilih 1-6.")
                    
                input("\nTekan Enter untuk melanjutkan...")
                
            except KeyboardInterrupt:
                print("\n\nProgram dihentikan oleh user.")
                break
            except Exception as e:
                print(f"Error: {e}")
                input("Tekan Enter untuk melanjutkan...")

def main():
    """Fungsi utama"""
    print("Memulai Invoice Processing Automation...")
    
    
    try:
        manager = InvoiceManager()
        manager.run()
    except Exception as e:
        print(f"Error menjalankan aplikasi: {e}")

if __name__ == "__main__":
    main()