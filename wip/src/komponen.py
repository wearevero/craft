import os
import pandas as pd
from utils import create_driver, login, generate_urls, collect_data, get_tanggal_input


def main():
    print("🚀 Memulai scraping data KOMPONEN...")

    input_tanggal = get_tanggal_input()
    print(f"📅 Tanggal yang dipilih: {input_tanggal}")

    komponen_urls = generate_urls("laporan/komponen_cetak", input_tanggal)
    print(f"🔗 Generated {len(komponen_urls)} URLs untuk KOMPONEN")

    driver = create_driver()

    try:

        login(driver)

        komponen_data = collect_data(driver, komponen_urls, "KOMPONEN")

        print(f"\n📊 Data KOMPONEN yang terkumpul: {len(komponen_data)} baris")

        if not komponen_data:
            print("❌ Tidak ada data KOMPONEN yang berhasil dikumpulkan!")
            print("🔍 Kemungkinan penyebab:")
            print("  - Selector CSS tidak cocok dengan struktur HTML")
            print("  - Tanggal tidak memiliki data komponen")
            print("  - Halaman web berubah struktur")
            print("  - Koneksi timeout")
            print("  - URL komponen_cetak tidak valid")
            return

        os.makedirs("data", exist_ok=True)

        df_komponen = pd.DataFrame(komponen_data)
        filename = f"data/komponen-bagian-{input_tanggal}.xlsx"

        if len(komponen_data) > 0:

            print(f"📋 Struktur data KOMPONEN pertama: {komponen_data[0]}")
            print(f"📏 Jumlah kolom: {len(komponen_data[0]) if komponen_data else 0}")

            print("📝 Preview data KOMPONEN (5 baris pertama):")
            for i, row in enumerate(komponen_data[:5]):
                print(f"  Baris {i+1}: {row}")

            print("💡 Struktur kolom:")
            print("  Kolom terakhir: NAMA BAGIAN (UPPERCASE)")
            print("  Kolom sebelumnya: Jenis material (Diamond/Mounting/CVD, dll)")

            df_komponen.to_excel(filename, index=False, header=False)

            if os.path.exists(filename):
                file_size = os.path.getsize(filename)
                print(f"✅ Data KOMPONEN berhasil disimpan di '{filename}'")
                print(f"📁 Ukuran file: {file_size} bytes")

                df_test = pd.read_excel(filename, header=None)
                print(
                    f"✅ Verifikasi: File berisi {len(df_test)} baris dan {len(df_test.columns)} kolom"
                )

                print(f"📈 Statistik data:")
                print(f"  - Total baris: {len(df_test)}")
                print(f"  - Total kolom: {len(df_test.columns)}")
                print(
                    f"  - Bagian unik: {df_test[0].nunique() if len(df_test) > 0 else 0}"
                )

            else:
                print(f"❌ File tidak berhasil dibuat: {filename}")
        else:
            print("❌ Data kosong, file tidak dibuat")

    except Exception as e:
        print(f"❌ Terjadi kesalahan: {e}")
        import traceback

        traceback.print_exc()

        print("\n🔍 Debug info untuk KOMPONEN:")
        try:
            print(f"📍 Current URL: {driver.current_url}")
            print(f"📄 Page title: {driver.title}")

            with open("debug_komponen_page.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            print("📄 HTML halaman komponen disimpan ke debug_komponen_page.html")

        except Exception as debug_e:
            print(f"⚠️ Gagal mendapatkan debug info: {debug_e}")

    finally:

        print("\n🟢 Browser tetap terbuka untuk pemeriksaan manual.")
        print("💡 Tips debugging untuk KOMPONEN:")
        print("  1. Cek apakah URL laporan/komponen_cetak valid")
        print("  2. Bandingkan struktur tabel dengan laporan/loss_bagian_cetak")
        print("  3. Pastikan tanggal memiliki data komponen")
        print("  4. Inspect element untuk melihat struktur HTML tabel komponen")
        print("  5. Cek console browser untuk error JavaScript")
        print("  6. Verifikasi apakah bagian_mapping untuk komponen sama dengan loss")


if __name__ == "__main__":
    main()
