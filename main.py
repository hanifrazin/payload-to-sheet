import argparse
import os
import json
import shutil
from src.convert_json import process_json_to_excel, response_json_to_excel
from src.extract_payload import extract_payloads_from_excel

def load_config():
    config_path = 'config.json'
    example_path = 'config.json.mac.example'
    
    # 1. Otomatis buat config.json dari example jika belum ada
    if not os.path.exists(config_path) and os.path.exists(example_path):
        print("[INFO] config.json tidak ditemukan. Menyalin dari template example...")
        shutil.copyfile(example_path, config_path)
    
    # 2. Baca isi config
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r') as f:
                return json.load(f)
        except json.JSONDecodeError:
            print("[WARNING] Format config.json salah. Pastikan format JSON valid.")
    return {}

def main():
    config = load_config()
    parser = argparse.ArgumentParser(description="Payload-to-Sheet Tool")
    subparsers = parser.add_subparsers(dest='command', help='Pilih mode')

    # Perintah: response
    p_response = subparsers.add_parser('response', help='JSON ke 1 kolom Excel dengan hidden apostrophe')
    p_response.add_argument('-i', '--input', help='Path folder JSON')
    p_response.add_argument('-o', '--output', help='Nama file output')

    # Sub-command: convert
    p_convert = subparsers.add_parser('convert', help='Konversi JSON ke Excel')
    # Hilangkan required=True agar tidak error saat panggil 'python main.py convert' saja
    p_convert.add_argument('-i', '--input', help='Path folder JSON')
    p_convert.add_argument('-o', '--output', help='Nama file output')

    # Sub-command: extract
    p_extract = subparsers.add_parser('extract', help='Ekstrak JSON dari Excel')
    p_extract.add_argument('-i', '--input', help='Path file .xlsx')
    p_extract.add_argument('-c', '--col', help='Nama kolom')
    p_extract.add_argument('-s', '--sheet', help='Nama sheet')

    args = parser.parse_args()
    
    # Ambil data dari config (default ke dictionary kosong jika tidak ada)
    conf_convert = config.get('convert', {})
    conf_extract = config.get('extract', {})
    conf_response = config.get('response', {})

    if args.command == 'convert':
        # Prioritas: 1. CLI (-i), 2. Config.json, 3. None
        folder_input = args.input or conf_convert.get('input')
        output_name = args.output or conf_convert.get('output')

        if not folder_input:
            print("\n[ERROR] Input tidak ditemukan! Isi config.json atau gunakan -i <folder>")
            return

        process_json_to_excel(folder_input, output_name)
    elif args.command == 'response':
        folder_input = args.input or conf_response.get('input')
        output_name = args.output or conf_response.get('output')

        if not folder_input:
            print("\n[ERROR] Input tidak ditemukan! Isi config.json atau gunakan -i <folder>")
            return

        response_json_to_excel(folder_input, output_name)
    elif args.command == 'extract':
        file_input = args.input or conf_extract.get('input')
        column = args.col or conf_extract.get('column', 'Request')
        sheet = args.sheet or conf_extract.get('sheet')

        if not file_input:
            print("\n[ERROR] Input tidak ditemukan! Isi config.json atau gunakan -i <file>")
            return
            
        extract_payloads_from_excel(file_input, column, sheet)
    else:
        # Jika user hanya ngetik 'python main.py' tanpa perintah
        parser.print_help()

if __name__ == "__main__":
    main()