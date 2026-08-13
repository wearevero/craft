import streamlit as st
import pandas as pd
import openpyxl
import re
import os
import zipfile
from io import BytesIO

st.set_page_config(page_title="⚡ CI Automator", layout="wide")
st.title("Customs Invoice Automation")

# ==========================================
# 1. FUNGSI LOGIKA BISNIS & EKSTRAKSI
# ==========================================

def get_item_category(item_no):
    item_str = str(item_no).strip().upper()
    if "EARRING" in item_str: return "E"
    if "RING" in item_str: return "R"
    if "BRACELET" in item_str: return "B"
    if "NECKLACE" in item_str: return "N"
    if "PENDANT" in item_str: return "P"
    
    if len(item_str) >= 8:
        chunk = item_str[5:8]
        for char in chunk:
            if char in ["B", "E", "N", "P", "R"]: return char
            if char == "C": return "N" 
    return "N" 

def clean_num(val):
    if pd.isna(val) or str(val).strip() == "": return 0.0
    try:
        return float(re.sub(r'[^\d\.\-]', '', str(val)))
    except:
        return 0.0

def process_proforma(file_bytes):
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    sheet = wb.active
    
    # Ambil data langsung tembak ke H14, I14, atau J14 (Gw sisain baris 13 buat jaga-jaga kalau ada format lama)
    raw_to = str(sheet["J14"].value or sheet["I14"].value or sheet["H14"].value or sheet["J13"].value or sheet["I13"].value or "").strip()
    
    # Kalau teks "To:" gabung sama nama negara di satu sel, bersihin biar rapi
    if raw_to.lower().startswith("to:"):
        raw_to = raw_to[3:].strip()
    elif raw_to.lower().startswith("to :"):
        raw_to = raw_to[4:].strip()
        
    # --- A. EKSTRAK LETTERHEAD ---
    header_data = {
        'Consignee': [], 
        'Bill_To': [],
        'Invoice_No': '',
        'Departure_Date': '',
        'Vessel_Flight': '',
        'To': raw_to 
    }
    
    cons_r, cons_c = None, None
    bill_r, bill_c = None, None
    
    for r in range(1, 20):
        # Scan lebar sampai kolom 18 buat nangkap alamat yang kejauhan
        for c in range(1, 18):
            val_raw = sheet.cell(row=r, column=c).value
            val = str(val_raw or "").strip().lower()
            
            if "consignee:" in val or "consignee" == val:
                if not cons_r: 
                    cons_r = r
                    cons_c = c + 1
                    for check_c in range(c + 1, 18):
                        if sheet.cell(row=r, column=check_c).value:
                            cons_c = check_c
                            break
            elif "bill to" in val:
                if not bill_r: 
                    bill_r = r
                    bill_c = c + 1
                    for check_c in range(c + 1, 18):
                        if sheet.cell(row=r, column=check_c).value:
                            bill_c = check_c
                            break
            
            elif "invoice no" in val:
                inv_val = sheet.cell(row=r, column=c+1).value or sheet.cell(row=r, column=c+2).value or sheet.cell(row=r, column=c+3).value
                if not inv_val:
                    for check_c in range(c + 1, 18):
                        if sheet.cell(row=r, column=check_c).value:
                            inv_val = sheet.cell(row=r, column=check_c).value
                            break
                header_data['Invoice_No'] = inv_val
            elif "departure date" in val:
                dep_val = sheet.cell(row=r, column=c+1).value or sheet.cell(row=r, column=c+2).value
                if not dep_val:
                    for check_c in range(c + 1, 18):
                        if sheet.cell(row=r, column=check_c).value:
                            dep_val = sheet.cell(row=r, column=check_c).value
                            break
                header_data['Departure_Date'] = dep_val
            elif "vessel/flight" in val:
                found_c = None
                ves_val = sheet.cell(row=r, column=c+1).value
                if ves_val:
                    found_c = c + 1
                else:
                    ves_val = sheet.cell(row=r, column=c+2).value
                    if ves_val:
                        found_c = c + 2
                        
                if not ves_val:
                    for check_c in range(c + 1, 18):
                        if sheet.cell(row=r, column=check_c).value:
                            ves_val = sheet.cell(row=r, column=check_c).value
                            found_c = check_c
                            break
                            
                # --- TAMBAHAN LOGIKA GABUNG AKUN KURIR ---
                # Kalau dapet FedEx/Brinks, cek 1 sel di sebelah kanannya persis
                if ves_val and found_c:
                    acc_val = sheet.cell(row=r, column=found_c + 1).value
                    # Kalau ada isinya, gabungin pakai spasi
                    if acc_val:
                        ves_val = f"{str(ves_val).strip()} {str(acc_val).strip()}"
                        
                header_data['Vessel_Flight'] = str(ves_val).strip() if ves_val else ""
            elif val in ["to", "to:", "to :"]:
                to_val = sheet.cell(row=r, column=c+1).value or sheet.cell(row=r, column=c+2).value
                if not to_val:
                    for check_c in range(c + 1, 18):
                        if sheet.cell(row=r, column=check_c).value:
                            to_val = sheet.cell(row=r, column=check_c).value
                            break
                if to_val:
                    header_data['To'] = str(to_val).strip()

    # --- A2. EKSTRAKSI ALAMAT (SISTEM KUNCI KOLOM, REM, & TARUH BAWAH) ---
    
    if cons_r:
        cons_start_col = None
        main_lines = []
        extra_lines = []
        for i in range(6):
            r_scan = cons_r + i
            if not cons_start_col:
                # Cari titik awal Consignee di blok kiri
                for col_scan in range(2, 10): 
                    val_str = str(sheet.cell(row=r_scan, column=col_scan).value or "").strip()
                    if val_str and val_str.lower() not in ["consignee", "consignee:", "(buyer : if then consignee):"] and not any(x in val_str.lower() for x in ["departure", "vessel", "invoice", "bill to"]):
                        cons_start_col = col_scan # Kunci posisinya!
                        main_lines.append(val_str)
                        
                        # Cek kolom kanannya buat tambahan
                        for c_offset in range(1, 6):
                            next_val = str(sheet.cell(row=r_scan, column=col_scan + c_offset).value or "").strip()
                            if next_val and next_val.lower() not in ["bill to", "bill to:"]:
                                extra_lines.append(next_val)
                        break
            else:
                main_val = str(sheet.cell(row=r_scan, column=cons_start_col).value or "").strip()
                if main_val and any(stop in main_val.lower() for stop in ["terms", "po#", "departure", "vessel", "bill to", "invoice"]):
                    break # Nabrak label bawah, langsung rem!
                    
                # Cek kolom kanannya buat tambahan
                row_extras = []
                for c_offset in range(1, 6):
                    c_idx = cons_start_col + c_offset
                    next_val = str(sheet.cell(row=r_scan, column=c_idx).value or "").strip()
                    if next_val and next_val.lower() not in ["bill to", "bill to:"]:
                        row_extras.append(next_val)
                
                if not main_val and not row_extras: 
                    break # Kalau sejajar kosong semua, stop.
                    
                if main_val:
                    main_lines.append(main_val)
                extra_lines.extend(row_extras)
                
        # Gabungin vertikal: Teks utama dulu, baru teks tambahan di bawahnya
        header_data['Consignee'].extend(main_lines + extra_lines)
                
    if bill_r:
        bill_start_col = None
        main_lines_b = []
        extra_lines_b = []
        for i in range(6):
            r_scan = bill_r + i
            if not bill_start_col:
                # Cari titik awal Bill To di blok kanan
                for col_scan in range(10, 21):
                    val_str = str(sheet.cell(row=r_scan, column=col_scan).value or "").strip()
                    if val_str and val_str.lower() not in ["bill to", "bill to:", "bill to :", "consignee", "consignee:"]:
                        bill_start_col = col_scan
                        main_lines_b.append(val_str)
                        
                        for c_offset in range(1, 6):
                            next_val = str(sheet.cell(row=r_scan, column=col_scan + c_offset).value or "").strip()
                            if next_val: 
                                extra_lines_b.append(next_val)
                        break
            else:
                main_val = str(sheet.cell(row=r_scan, column=bill_start_col).value or "").strip()
                if main_val and any(stop in main_val.lower() for stop in ["terms", "po#", "departure", "vessel", "consignee"]):
                    break
                    
                row_extras = []
                for c_offset in range(1, 6):
                    c_idx = bill_start_col + c_offset
                    next_val = str(sheet.cell(row=r_scan, column=c_idx).value or "").strip()
                    if next_val: 
                        row_extras.append(next_val)
                
                if not main_val and not row_extras:
                    break
                    
                if main_val:
                    main_lines_b.append(main_val)
                extra_lines_b.extend(row_extras)

        # Gabungin vertikal: Teks utama dulu, baru teks tambahan di bawahnya
        header_data['Bill_To'].extend(main_lines_b + extra_lines_b)

    # --- TAMBAHAN LOGIKA BILL TO GHOIB ---
    # Kalau list Bill_To masih kosong, langsung clone (copy) dari Consignee
    if not header_data['Bill_To'] and header_data['Consignee']:
        header_data['Bill_To'] = header_data['Consignee'].copy()


    # --- B. SCAN KANAN KE KIRI UNTUK HEADER TABEL ---
    header_r = None
    for r in range(1, 50):
        row_vals = [re.sub(r'\s+', ' ', str(sheet.cell(row=r, column=c).value or "").strip().lower()) for c in range(1, 15)]
        if any("po#" in v or "item no." in v or "item no" == v for v in row_vals):
            header_r = r
            break
            
    if not header_r:
        raise ValueError("❌ Gagal menemukan baris Header Tabel (PO# / Item No.) di file PI ini.")
        
    col_map = {}
    for c in range(1, sheet.max_column + 1):
        raw_val = sheet.cell(row=header_r, column=c).value
        if not raw_val: continue
        
        h = re.sub(r'\s+', ' ', str(raw_val).strip().lower())
        
        if h == "po#": col_map["po"] = c
        elif h == "item no." or h == "item no": col_map["item"] = c
        elif h == "metal": col_map["metal"] = c
        elif h == "q'ty" or h == "qty": col_map["qty"] = c
        elif h == "total w't" or h == "total wt": col_map["net"] = c
        
        elif h == "labor" and c > 20: col_map["mfg"] = c
        elif "total (id customs)" in h: col_map["id_total"] = c
        elif (h == "dia" or h == "dia value") and c > 20: col_map["dia"] = c
        elif "gold (us customs)" in h: col_map["gold"] = c
        elif "total (us customs)" in h: col_map["us_total"] = c

    def safe_cell(r, col_key):
        c_idx = col_map.get(col_key)
        if not c_idx: return None 
        return sheet.cell(row=r, column=c_idx).value

    # --- C. EKSTRAK ITEM ---
    raw_items = []
    active_po = None
    cancellation_mode = False # <-- FLAG MODE BATAL
    
    for r in range(header_r + 1, sheet.max_row + 1):
        
        # RADAR RESET: Kalau ketemu ujung blok (Subtotal/Total), matiin mode batalnya biar item di blok bawah aman!
        first_col_val = str(sheet.cell(row=r, column=1).value or "").strip().lower()
        item_col_val = str(safe_cell(r, "item") or "").strip().lower()
        if "total" in first_col_val or "subtotal" in first_col_val or "total" in item_col_val:
            cancellation_mode = False
            
        # RADAR CANCELLATION: Cek 5 kolom awal, kalau nemu kata "cancellation", aktifkan mode!
        for check_c in range(1, 6):
            cell_val = str(sheet.cell(row=r, column=check_c).value or "").strip().lower()
            if "cancellation" in cell_val:
                cancellation_mode = True
                break

        po_raw = safe_cell(r, "po")
        po_str = str(po_raw or "").strip()
        
        if (not po_raw and not active_po) or "total" in item_col_val: 
            continue
        
        if "K" in po_str.upper() and len(po_str) > 5:
            active_po = po_str
        else:
            po_str = active_po 
            
        item_no = safe_cell(r, "item")
        if not item_no: continue
        
        # BLOKIR GHOST ITEM 14KN: Lewati baris kalau isinya ternyata Header berulang
        if str(item_no).strip().lower() in ["item no.", "item no", "item"]:
            continue
            
        metal = safe_cell(r, "metal")
        k_match = re.search(r'(14K|18K|9K|24K)', str(metal or "14K"), re.IGNORECASE)
        karat = k_match.group(0).upper() if k_match else "14K"
        cat = get_item_category(item_no)
        gen_id = f"{karat}{cat}"
        
        # LOGIKA CANCELLATION: Set QTY jadi 0 kalau mode aktif, selain itu normal
        final_qty = 0.0 if cancellation_mode else clean_num(safe_cell(r, "qty"))
        
        raw_items.append({
            "ID": gen_id,
            "Qty": final_qty,
            "Net": clean_num(safe_cell(r, "net")) / 1000.0,
            "ID_Total": clean_num(safe_cell(r, "id_total")),
            "Mfg": clean_num(safe_cell(r, "mfg")),
            "Dia": clean_num(safe_cell(r, "dia")),
            "Gold": clean_num(safe_cell(r, "gold")),
            "US_Total": clean_num(safe_cell(r, "us_total"))
        })
        
    if not raw_items:
         raise ValueError("❌ Tidak ada data item yang bisa dibaca. Pastikan layout kolom PI sesuai standar.")
         
    df = pd.DataFrame(raw_items)
    df_grouped = df.groupby('ID', as_index=False).sum()
    
    cols_to_round = ['Net', 'ID_Total', 'Mfg', 'Dia', 'Gold', 'US_Total']
    df_grouped[cols_to_round] = df_grouped[cols_to_round].round(2)
    
    # --- LOGIKA MINIMAL NETTO 0.01 ---
    # Paksa nilai Net jadi 0.01 khusus untuk yang hasil akhirnya 0.00
    df_grouped.loc[df_grouped['Net'] == 0.00, 'Net'] = 0.01
    
    return header_data, df_grouped

# ==========================================
# 2. INJEKSI KE TEMPLATE CUSTOMS INVOICE
# ==========================================

def inject_customs_invoice(template_bytes, header, df_items, version):
    wb = openpyxl.load_workbook(BytesIO(template_bytes))
    ws = wb["Customs Invoice"]
    import datetime # Import modul waktu
    
    def inject_multiline(start_row, start_col, data_list, max_lines):
        for i, text in enumerate(data_list):
            if i < max_lines: 
                # MURNI HANYA NGISI NILAI. GAK NYENTUH FONT ATAU ALIGNMENT TEMPLATE.
                ws.cell(row=start_row + i, column=start_col).value = text

    # --- EKSTRAK DATE, DEPARTURE DATE & PREFIX DARI INVOICE NO ---
    inv_no = header.get("Invoice_No", "")
    
    extracted_date = ""
    # Fallback: Ambil dari hasil scan kalau ternyata angka di invoice kurang panjang
    departure_date_val = header.get("Departure_Date", "") 
    
    # Deteksi 10 digit angka: 2 (YY), 2 (MM Inv), 2 (DD Inv), opsional 2 (MM Dep), opsional 2 (DD Dep)
    match_date = re.search(r'(\d{2})(\d{2})(\d{2})(?:(\d{2})(\d{2}))?', inv_no)
    if match_date:
        yy = match_date.group(1)
        mm_inv = match_date.group(2)
        dd_inv = match_date.group(3)
        extracted_date = f"20{yy}-{mm_inv}-{dd_inv}"
        
        # Cek kalau ada 4 digit ekstra di belakangnya buat Departure Date
        mm_dep = match_date.group(4)
        dd_dep = match_date.group(5)
        if mm_dep and dd_dep:
            # Ubah jadi objek Date biar Excel kenal
            departure_date_val = datetime.datetime(int(f"20{yy}"), int(mm_dep), int(dd_dep))
            
    inv_prefix = ""
    match_prefix = re.search(r'^([A-Za-z]+)', inv_no)
    if match_prefix:
        inv_prefix = match_prefix.group(1).upper()
    else:
        inv_prefix = "IDD"
    # ---------------------------------------------

    # --- LOGIKA BIAYA TAMBAHAN (ANTI ERROR JIKA KOSONG) ---
    dest = str(header.get("To", "") or "").strip().upper()
    vessel = str(header.get("Vessel_Flight", "") or "").strip().upper()
    
    doc_fee = 0
    handling_fee = 0
    
    if dest:
        eropa_list = ["UK", "UNITED KINGDOM", "ENGLAND", "ITALY", "ITALIA", "SWEDEN", "GERMANY", "FRANCE", "SPAIN", "NETHERLANDS", "BELGIUM", "SWITZERLAND", "POLAND", "DENMARK", "NORWAY", "FINLAND", "AUSTRIA", "IRELAND", "PORTUGAL", "GREECE", "EUROPE"]
        if re.search(r'\b(USA|U\.S\.A\.?|UNITED STATES)\b', dest) or any(er in dest for er in eropa_list):
            doc_fee = 30
            
    if vessel:
        if "FEDEX" in vessel:
            handling_fee = 40
    # --------------------------------------------------------

    if version == "US":
        inject_multiline(9, 2, header.get("Consignee", []), 6) 
        inject_multiline(6, 11, header.get("Bill_To", []), 6)  
        
        # Inject Departure Date dengan format Excel Date
        ws["B15"].value = departure_date_val
        if type(departure_date_val) is datetime.datetime:
            ws["B15"].number_format = 'd-mmm-yyyy'
            
        ws["B16"].value = header.get("Vessel_Flight", "")
        ws["B17"].value = header.get("To", "")
        if not ws["B17"].value: ws["B17"].value = "U.S.A" 
            
        ws["K4"].value = inv_no 
        ws["L4"].value = extracted_date  
        ws["L4"].number_format = '@'
        
        ws["M165"].value = handling_fee
        ws["M166"].value = doc_fee
        
        start_r = 21
        cols = {'ID': 6, 'Qty': 7, 'Net': 9, 'ID_Total': 12, 'Mfg': 13, 'Dia': 14, 'Gold': 15, 'US_Total': 16}
        id_col_letter = "E"
        
    elif version == "International":
        inject_multiline(10, 2, header.get("Consignee", []), 6) 
        inject_multiline(6, 10, header.get("Bill_To", []), 7)   
        
        # Inject Departure Date dengan format Excel Date
        ws["B16"].value = departure_date_val
        if type(departure_date_val) is datetime.datetime:
            ws["B16"].number_format = 'd-mmm-yyyy'
            
        ws["B17"].value = header.get("Vessel_Flight", "")
        ws["B18"].value = header.get("To", "")
        
        ws["J4"].value = inv_no 
        ws["K4"].value = extracted_date  
        ws["K4"].number_format = '@'
        
        ws["L70"].value = handling_fee
        ws["L71"].value = doc_fee
        
        start_r = 22
        cols = {'ID': 5, 'Qty': 6, 'Net': 8, 'ID_Total': 11, 'Mfg': 12, 'Dia': 13, 'Gold': 14, 'US_Total': 15}
        id_col_letter = "D" 
    
    # --- INJECT DATA ITEM & RUMUS ID CODE ---
    items_to_inject = df_items.head(20)
    num_items = len(items_to_inject)
    
    for i, row in items_to_inject.iterrows():
        r = start_r + i
        
        formula = f'=CONCATENATE("MFGG-",MID({id_col_letter}{r},1,3),MID({id_col_letter}{r},10,1),"-{inv_prefix}")'
        ws.cell(row=r, column=3).value = formula
        
        ws.cell(row=r, column=cols['ID']).value = row['ID']
        ws.cell(row=r, column=cols['Qty']).value = row['Qty']
        ws.cell(row=r, column=cols['Net']).value = row['Net']
        ws.cell(row=r, column=cols['ID_Total']).value = row['ID_Total']
        ws.cell(row=r, column=cols['Mfg']).value = row['Mfg']
        ws.cell(row=r, column=cols['Dia']).value = row['Dia']
        ws.cell(row=r, column=cols['Gold']).value = row['Gold']
        ws.cell(row=r, column=cols['US_Total']).value = row['US_Total']
        
    # --- HIDE BARIS KOSONG ---
    for i in range(num_items, 20):
        empty_r = start_r + i
        ws.row_dimensions[empty_r].hidden = True
        
    out_buffer = BytesIO()
    wb.save(out_buffer)
    return out_buffer.getvalue()

# ==========================================
# 3. ANTARMUKA WEB (BATCH UPLOAD & CUSTOM UI)
# ==========================================

st.markdown(
    """
    <style>
    [data-testid="stFileUploaderDropzone"] {
        border: 3px dashed #4285F4 !important;
        border-radius: 20px !important;
        padding: 60px !important;
        background-color: #1e1e1e !important;
        text-align: center !important;
        transition: all 0.3s ease !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        border-color: #34A853 !important;
        background-color: #2a2a2a !important;
    }
    
    [data-testid="stFileUploaderDropzone"] > div > svg {
        width: 70px !important;
        height: 70px !important;
        fill: #4285F4 !important;
    }

    [data-testid="stUploadedFile"] > div:first-child > svg {
        display: none !important;
    }
    
    [data-testid="stUploadedFile"] > div:first-child::before {
        content: "PI";
        display: flex;
        align-items: center;
        justify-content: center;
        background-color: #ffffff !important;
        color: #34A853 !important;
        font-weight: 900 !important;
        font-size: 15px !important;
        width: 32px !important;
        height: 40px !important;
        border-radius: 6px !important;
        margin-right: 12px !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2) !important;
        font-family: Arial, sans-serif !important;
    }
    
    [data-testid="stUploadedFile"] button svg {
        width: 14px !important;
        height: 14px !important;
        fill: #ffffff !important;
    }
    [data-testid="stUploadedFile"] button {
        background-color: transparent !important;
        border-radius: 50% !important;
        padding: 4px !important;
        transition: 0.2s !important;
    }
    [data-testid="stUploadedFile"] button:hover {
        background-color: #ff4b4b !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.write("### 📤 Upload Proforma Invoice")
st.markdown("Upload or Drag & Drop file PI nya di bawah ini.")

pi_files = st.file_uploader("", type=["xlsx"], accept_multiple_files=True)

if pi_files:
    st.info(f"📦 {len(pi_files)} PI files detected")
    
    if st.button("Generate Semua Customs Invoice"):
        with st.spinner("Merakit data..."):
            generated_files = [] # Tempat penampungan hasil file sementara
            
            for pi_file in pi_files:
                try:
                    # 1. BACA DALEMAN PI DULUAN
                    file_bytes = pi_file.read()
                    header_data, df_items = process_proforma(file_bytes)
                    
                    # --- OVERRIDE INVOICE NO DARI NAMA FILE ---
                    raw_filename = pi_file.name
                    name_no_ext = re.sub(r'(?i)\.xlsx?$', '', raw_filename)
                    inv_name = re.sub(r'(?i)^PI\s+', '', name_no_ext)
                    inv_name = re.sub(r'(?i)\s*R\d+$', '', inv_name).strip()
                    header_data['Invoice_No'] = inv_name
                    
                    # 2. LOGIKA RADAR "TO:" (Deteksi Destinasi)
                    dest = str(header_data.get('To', '')).upper()
                    if re.search(r'\b(USA|U\.S\.A\.?|UNITED STATES)\b', dest):
                        version = "US"
                        template_path = "templates/Template_US.xlsx"
                    else:
                        version = "International"
                        template_path = "templates/Template_Int.xlsx"
                        
                    # 3. LANJUT INJECT DATA KE TEMPLATE YG TEPAT
                    if not os.path.exists(template_path):
                        st.error(f"❌ Template '{template_path}' tidak ditemukan di folder!")
                        continue
                        
                    with open(template_path, "rb") as f:
                        template_bytes = f.read()
                        
                    final_file = inject_customs_invoice(template_bytes, header_data, df_items, version)
                    
                    # --- LOGIKA NON MAKLON ---
                    # Bandingkan total nilai ID_Total, Mfg (Labor), dan US_Total (dibulatkan ke 2 desimal biar aman)
                    sum_id = round(df_items['ID_Total'].sum(), 2)
                    sum_mfg = round(df_items['Mfg'].sum(), 2)
                    sum_us = round(df_items['US_Total'].sum(), 2)
                    
                    is_non_maklon = (sum_id == sum_mfg) and (sum_mfg == sum_us)
                    
                    # Bersihin karakter ilegal buat nama file
                    inv_clean = "".join(c for c in inv_name if c.isalnum() or c in " -_")
                    if not inv_clean: inv_clean = "FINAL"
                    
                    # Eksekusi penamaan file
                    if is_non_maklon:
                        out_name = f"{inv_clean} Indonesian Customs Invoice NON MAKLON.xlsx"
                    else:
                        out_name = f"{inv_clean} Indonesian Customs Invoice.xlsx"
                    
                    # Simpan data file ke list penampungan
                    generated_files.append({"name": out_name, "data": final_file})
                    st.success(f"✅ {out_name} berhasil dirakit!")
                    
                except Exception as e:
                    st.error(f"❌ Gagal memproses '{pi_file.name}': {e}")
            
            # --- LOGIKA DOWNLOAD (1 FILE = XLSX, >1 FILE = ZIP) ---
            if generated_files:
                st.write("---")
                
                if len(generated_files) == 1:
                    # Cuma 1 file? Langsung kasih Excel-nya!
                    single_file = generated_files[0]
                    st.download_button(
                        label=f"📥 Download Hasil ({single_file['name']})",
                        data=single_file['data'],
                        file_name=single_file['name'],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    # Lebih dari 1 file? Bungkus jadi ZIP!
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        for gf in generated_files:
                            zip_file.writestr(gf['name'], gf['data'])
                            
                    st.download_button(
                        label=f"📥 Download {len(generated_files)} Hasil (Format .ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name="Batch_Customs_Invoices.zip",
                        mime="application/zip"
                    )