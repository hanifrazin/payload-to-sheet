import os
import json
import pandas as pd
import re
from collections import defaultdict


def clean_header(text):
    """Mengubah camelCase atau nested path ke SNAKE_CASE_KAPITAL"""
    # Pisahkan berdasarkan titik (nested) dan ambil bagian terakhir jika ingin ringkas,
    # namun di sini kita gunakan full path agar unik, lalu diclean.
    s1 = re.sub('(.)([A-Z][a-z]+)', r'\1_\2', text)
    s2 = re.sub('([a-z0-9])([A-Z])', r'\1_\2', s1)
    return re.sub(r'[^A-Z0-9]+', '_', s2.upper()).strip('_')


def flatten_json(y, parent_key='', sep='.'):
    """Membongkar struktur JSON nested menjadi satu level (Flatten)"""
    items = []
    for k, v in y.items():
        new_key = parent_key + sep + k if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten_json(v, new_key, sep=sep).items())
        elif isinstance(v, list):
            # Jika list berisi objek, kita jadikan string untuk kolom ini
            items.append((new_key, str(v)))
        else:
            items.append((new_key, v))
    return dict(items)


def process_json_to_excel(folder_path, custom_output_name=None):
    # 1. Setup Folder Output
    output_dir = "result"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 2. Setup Nama File
    folder_path = folder_path.rstrip(os.sep).rstrip('/')
    folder_name = os.path.basename(folder_path)

    if custom_output_name:
        base_name = os.path.splitext(custom_output_name)[0]
        output_filename = f"{base_name}.xlsx"
    else:
        output_filename = f"{folder_name}_result.xlsx"

    full_output_path = os.path.join(output_dir, output_filename)

    if not os.path.exists(folder_path):
        print(f"\n[ERROR] Folder '{folder_path}' tidak ditemukan.")
        return

    # 3. Baca File
    json_files = [f for f in os.listdir(folder_path) if f.endswith('.json')]
    json_files.sort(key=lambda f: int(re.sub(r'\D', '', f) if re.sub(r'\D', '', f) else 0))

    print(f"\nMembaca folder: {folder_path}")
    print(f"Memproses {len(json_files)} file...")

    all_rows = []

    for file_name in json_files:
        file_path = os.path.join(folder_path, file_name)
        with open(file_path, 'r', encoding='utf-8') as f:
            try:
                data = json.load(f)

                # Cek apakah file berisi Array (seperti users.json) atau Single Object
                records = data if isinstance(data, list) else [data]

                for item in records:
                    flat_data = flatten_json(item)
                    row_data = {"SOURCE": file_name}

                    # Logic: Cek duplikasi value antara Parent vs Child
                    for full_key, value in flat_data.items():
                        val_str = str(value).strip()
                        header = clean_header(full_key)

                        # Cek apakah value ini sudah ada di kolom lain dalam baris yang sama
                        is_duplicate_val = False
                        for existing_val in row_data.values():
                            if str(existing_val).strip() == val_str and val_str != "":
                                is_duplicate_val = True
                                break

                        # Jika value unik, buat kolom baru. Jika sama, abaikan (Ekspektasi User)
                        if not is_duplicate_val:
                            row_data[header] = value

                    all_rows.append(row_data)

            except Exception as e:
                print(f"[GAGAL] {file_name}: {e}")

    if not all_rows:
        print("---------------------------------------------")
        print("GAGAL: Tidak ada data valid untuk dikonversi.")
        print("---------------------------------------------")
        return

    # 4. Simpan ke Excel
    df = pd.DataFrame(all_rows).fillna('')
    writer = pd.ExcelWriter(full_output_path, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    text_format = workbook.add_format({'num_format': '@', 'quote_prefix': True})

    for col_num, col_name in enumerate(df.columns):
        worksheet.set_column(col_num, col_num, 25)
        for row_num, value in enumerate(df[col_name]):
            worksheet.write_string(row_num + 1, col_num, str(value), text_format)

    writer.close()

    # 5. Tampilan Message Sesuai Permintaan
    total_data_rows = len(df)
    total_columns_data = len(df.columns) - 1  # Mengecualikan SOURCE

    print("---------------------------------------------")
    print("SUKSES KONVERSI!")
    print(f"Lokasi File  : {full_output_path}")
    print(f"Total Baris  : {total_data_rows} baris data")
    print(f"Total Kolom  : {total_columns_data} kolom data (Kolom SOURCE tidak dihitung)")
    print("---------------------------------------------")