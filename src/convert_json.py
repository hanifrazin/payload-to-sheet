import os
import json
import pandas as pd
import re
from collections import defaultdict


def clean_header(text):
    """Mengubah camelCase ke SNAKE_CASE_KAPITAL"""
    s1 = re.sub('(.)([A-Z][a-z]+)', r'\1_\2', text)
    s2 = re.sub('([a-z0-9])([A-Z])', r'\1_\2', s1)
    return re.sub(r'[^A-Z0-9]+', '_', s2.upper()).strip('_')


def flatten_json(y, parent_key='', sep='.'):
    """Fungsi helper untuk meratakan JSON (Flatten)"""
    items = []
    for k, v in y.items():
        new_key = parent_key + sep + k if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten_json(v, new_key, sep=sep).items())
        elif isinstance(v, list):
            # Jika list berisi objek, kita gabungkan sebagai string (atau sesuaikan kebutuhan)
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

    all_rows = []

    if not os.path.exists(folder_path):
        print(f"Error: Folder '{folder_path}' tidak ditemukan.")
        return

    json_files = [f for f in os.listdir(folder_path) if f.endswith('.json')]
    json_files.sort(key=lambda f: int(re.sub(r'\D', '', f) if re.sub(r'\D', '', f) else 0))

    for file_name in json_files:
        file_path = os.path.join(folder_path, file_name)
        with open(file_path, 'r', encoding='utf-8') as f:
            try:
                data = json.load(f)

                # JIKA DATA ADALAH LIST: Pecah menjadi baris-baris terpisah
                # Ini memperbaiki isu data 'users.json' yang jomplang
                records = data if isinstance(data, list) else [data]

                for item in records:
                    # Flatten data untuk mendapatkan path lengkap (misal: data.name)
                    flat_data = flatten_json(item)
                    row_data = {"SOURCE": file_name}

                    # Logic Pengecekan Duplikasi Value (Expectasi User)
                    processed_keys = {}
                    for full_key, value in flat_data.items():
                        base_key = full_key.split('.')[-1]  # Ambil 'name' dari 'data.name'
                        val_str = str(value)

                        header = clean_header(full_key)

                        # Cek apakah value ini sama dengan value di level parent yang sudah diproses
                        is_duplicate_val = False
                        for prev_key, prev_val in row_data.items():
                            if prev_key != "SOURCE" and str(prev_val) == val_str:
                                # Jika valuenya sama (misal 'Parent' di root dan 'Parent' di child)
                                # Kita cukup gunakan kolom yang sudah ada (Root)
                                is_duplicate_val = True
                                break

                        if not is_duplicate_val:
                            row_data[header] = value

                    all_rows.append(row_data)

            except Exception as e:
                print(f"Gagal memproses {file_name}: {e}")

    if not all_rows:
        print("Tidak ada data untuk diproses.")
        return

    df = pd.DataFrame(all_rows).fillna('')

    writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
    df.to_excel(writer, index=False)

    # Formatting (Auto-fit & Text format)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    text_format = workbook.add_format({'num_format': '@', 'quote_prefix': True})

    for col_num, col_name in enumerate(df.columns):
        worksheet.set_column(col_num, col_num, 25)
        for row_num, value in enumerate(df[col_name]):
            worksheet.write_string(row_num + 1, col_num, str(value), text_format)

    writer.close()
    print(f"SUKSES! Data tersimpan di: {output_filename} ({len(df)} baris)")