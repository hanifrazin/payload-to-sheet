import os
import json
import pandas as pd
import re
from collections import defaultdict

def clean_header(text):
    """Mengubah path nested ke SNAKE_CASE_KAPITAL"""
    s1 = re.sub('(.)([A-Z][a-z]+)', r'\1_\2', text)
    s2 = re.sub('([a-z0-9])([A-Z])', r'\1_\2', s1)
    return re.sub(r'[^A-Z0-9]+', '_', s2.upper()).strip('_')

def flatten_json(y, parent_key='', sep='.'):
    """Membongkar struktur JSON nested"""
    items = []
    for k, v in y.items():
        new_key = parent_key + sep + k if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten_json(v, new_key, sep=sep).items())
        elif isinstance(v, list):
            items.append((new_key, str(v)))
        else:
            items.append((new_key, v))
    return dict(items)

def process_json_to_excel(folder_path, custom_output_name=None):
    output_dir = "result"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    folder_path = folder_path.rstrip(os.sep).rstrip('/')
    folder_name = os.path.basename(folder_path)
    output_filename = os.path.join(output_dir, f"{custom_output_name or folder_name}_result.xlsx")
    
    if not os.path.exists(folder_path):
        print(f"\n[ERROR] Path '{folder_path}' tidak ditemukan.")
        return

    # --- VALIDASI BARU ---
    # Hanya ambil file yang berakhiran .json DAN pastikan itu adalah FILE (bukan folder)
    json_files = [
        f for f in os.listdir(folder_path) 
        if f.endswith('.json') and os.path.isfile(os.path.join(folder_path, f))
    ]

    # Jika tidak ada file .json sama sekali di folder tersebut
    if not json_files:
        print("-" * 45)
        print(f"[ERROR] Tidak ditemukan file .json di: {folder_path}")
        print("Pastikan Anda memasukkan folder yang berisi file JSON.")
        print("-" * 45)
        return

    # Urutan file secara numerik
    json_files.sort(key=lambda f: int(re.sub(r'\D', '', f) if re.sub(r'\D', '', f) else 0))
    
    print(f"\nMembaca folder: {folder_path}")
    print(f"Memproses {len(json_files)} file JSON...")

    all_rows = []
    for file_name in json_files:
        file_path = os.path.join(folder_path, file_name)
        with open(file_path, 'r', encoding='utf-8') as f:
            try:
                data = json.load(f)
                records = data if isinstance(data, list) else [data]
                
                for item in records:
                    flat_data = flatten_json(item)
                    row_data = {"SOURCE": file_name}
                    
                    # --- BAGIAN YANG DIUBAH ---
                    # Langsung masukkan semua key tanpa cek duplikasi nilai
                    for full_key, value in flat_data.items():
                        header = clean_header(full_key)
                        row_data[header] = value
                    # ---------------------------
                    
                    all_rows.append(row_data)
                        
                       
            except Exception as e:
                print(f"[GAGAL] {file_name}: {e}")

    # Simpan ke Excel
    try:
        df = pd.DataFrame(all_rows).fillna('')
        writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
        df.to_excel(writer, index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        text_format = workbook.add_format({'num_format': '@', 'quote_prefix': True})

        for col_num, col_name in enumerate(df.columns):
            worksheet.set_column(col_num, col_num, 25)
            for row_num, value in enumerate(df[col_name]):
                worksheet.write_string(row_num + 1, col_num, str(value), text_format)

        writer.close()

        print("---------------------------------------------")
        print("SUKSES KONVERSI!")
        print(f"Lokasi File  : {output_filename}")
        print(f"Total Baris  : {len(df)} baris data")
        print(f"Total Kolom  : {len(df.columns) - 1} kolom data")
        print("---------------------------------------------")
    except PermissionError:
        print("\n" + "!"*50)
        print("[GAGAL] FILE EXCEL SEDANG TERBUKA!")
        print(f"File: {output_filename}")
        print("Silakan TUTUP file tersebut di Excel, lalu jalankan perintah ini lagi.")
        print("!"*50 + "\n")
    except Exception as e:
        print(f"\n[ERROR] Terjadi kesalahan saat menyimpan file: {e}")


def response_json_to_excel(folder_path, custom_output_name=None):
    output_dir = "result"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    folder_path = folder_path.rstrip(os.sep).rstrip('/')
    folder_name = os.path.basename(folder_path)
    output_filename = os.path.join(output_dir, f"{custom_output_name or folder_name}_response.xlsx")

    if not os.path.exists(folder_path):
        print(f"\n[ERROR] Path '{folder_path}' tidak ditemukan.")
        return

    # Filter file .json
    json_files = [
        f for f in os.listdir(folder_path)
        if f.endswith('.json') and os.path.isfile(os.path.join(folder_path, f))
    ]

    if not json_files:
        print(f"[ERROR] Tidak ditemukan file .json di: {folder_path}")
        return

    # Natural Sorting (1, 2, 10...)
    def natural_sort_key(s):
        return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]

    json_files.sort(key=natural_sort_key)

    rows = []
    for file_name in json_files:
        file_path = os.path.join(folder_path, file_name)
        try:
            final_content = ""  # Default untuk file kosong

            if os.path.getsize(file_path) > 0:
                with open(file_path, 'r', encoding='utf-8') as f:
                    raw_data = f.read().strip()
                    try:
                        parsed = json.loads(raw_data)
                        final_content = json.dumps(parsed, indent=4, ensure_ascii=False)
                    except json.JSONDecodeError:
                        final_content = raw_data

            # Jangan tambahkan "'" di sini agar tidak muncul double atau terlihat literal
            rows.append({
                "SOURCE": file_name,
                "EXPECTED_JSON": final_content
            })
            print(f"[OK] {file_name} -> Baris {len(rows) + 1}")

        except Exception as e:
            print(f"[ERROR] {file_name}: {e}")

    df = pd.DataFrame(rows)
    writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # SOLUSI: Menggunakan quote_prefix=True
    # Ini akan menambahkan tanda apostrof secara 'hidden' di Excel
    json_format = workbook.add_format({
        'num_format': '@',
        'text_wrap': True,
        'valign': 'top',
        'font_name': 'Consolas',
        'font_size': 10,
        'quote_prefix': True  # <-- Ini kuncinya!
    })

    # Terapkan format ke kolom B (EXPECTED_JSON)
    # Tulis ulang data dengan format yang mendukung quote_prefix
    for row_num, content in enumerate(df['EXPECTED_JSON']):
        worksheet.write_string(row_num + 1, 1, content, json_format)

    worksheet.set_column(0, 0, 30)  # Kolom SOURCE
    worksheet.set_column(1, 1, 120)  # Kolom EXPECTED_JSON

    writer.close()

    print("\n" + "=" * 45)
    print("KONVERSI SELESAI (Hidden Apostrophe Active)")
    print(f"File Output : {output_filename}")
    print("=" * 45)