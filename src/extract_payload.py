import pandas as pd
import os
import re
import json

def find_header_row(file_path, column_keyword, sheet_name):
    """Mencari baris header secara otomatis pada sheet tertentu"""
    try:
        df_temp = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=20)
        for index, row in df_temp.iterrows():
            if row.astype(str).str.contains(column_keyword, case=False, na=False).any():
                return index
    except Exception as e:
        print(f"[ERROR] Gagal memindai header di sheet '{sheet_name}': {e}")
    return None

def extract_payloads_from_excel(input_excel_path, target_column="Request", target_sheet=None):
    if not os.path.exists(input_excel_path):
        print(f"[ERROR] File tidak ditemukan: {input_excel_path}")
        return

    try:
        xl = pd.ExcelFile(input_excel_path)
        sheet_name_raw = target_sheet if target_sheet else xl.sheet_names[0]
        if sheet_name_raw not in xl.sheet_names:
            print(f"[ERROR] Sheet '{sheet_name_raw}' tidak ditemukan. Pilihan: {xl.sheet_names}")
            return
    except Exception as e:
        print(f"[ERROR] Gagal membaca Excel: {e}")
        return

    # Normalisasi Nama Sheet (Spasi -> Underscore)
    sheet_name_clean = sheet_name_raw.replace(" ", "_")
    input_dir_name = os.path.basename(os.path.dirname(os.path.abspath(input_excel_path))) or "root_data"

    # Struktur folder sesuai request: result/extract_payload/{nama_sheet}/
    final_output_dir = os.path.join("result", "extract_payload", sheet_name_clean)
    
    if not os.path.exists(final_output_dir):
        os.makedirs(final_output_dir)

    header_idx = find_header_row(input_excel_path, target_column, sheet_name_raw)
    if header_idx is None:
        print(f"[ERROR] Kolom '{target_column}' tidak ditemukan di sheet '{sheet_name_raw}'.")
        return

    print(f"[INFO] Memproses Sheet: '{sheet_name_raw}'")
    print(f"[INFO] Output Folder: {final_output_dir}")

    # Membaca data tanpa membuang baris kosong agar index tetap akurat
    df = pd.read_excel(input_excel_path, sheet_name=sheet_name_raw, header=header_idx)
    series_data = df[target_column]
    
    total_items = len(series_data)
    pad_width = max(2, len(str(total_items)))

    count_success = 0
    count_skipped = 0

    for i, content in enumerate(series_data, 1):
        text = str(content).strip()
        file_number = str(i).zfill(pad_width)
        file_name = f"{sheet_name_clean}_{file_number}.json"

        # Regex untuk mendeteksi struktur JSON {}
        match = re.search(r'\{.*\}', text, re.DOTALL)
        
        if not match or text in ["nan", "", "None"]:
            print(f"[SKIP] Baris {i}: Data kosong atau bukan JSON valid (Akan dilewati).")
            count_skipped += 1
            continue

        try:
            json_str = match.group(0).strip()
            parsed_json = json.loads(json_str)
            # Pretty Print dengan indentasi 4 spasi
            json_beauty = json.dumps(parsed_json, indent=4, ensure_ascii=False)
            
            save_path = os.path.join(final_output_dir, file_name)
            
            # Mode 'w' otomatis akan overwrite file jika sudah ada
            with open(save_path, "w", encoding="utf-8") as f:
                f.write(json_beauty)
            count_success += 1
        except Exception as e:
            print(f"[ERR] Baris {i}: Gagal format JSON. ({e})")
            count_skipped += 1

    print("-" * 65)
    print(f"PROSES EKSTRAKSI SELESAI")
    print(f"Total Baris Diproses : {total_items}")
    print(f"File Berhasil Dibuat : {count_success} (Overwrite mode aktif)")
    print(f"Baris Dilewati       : {count_skipped}")
    print(f"Lokasi Hasil         : {final_output_dir}")
    print("-" * 65)