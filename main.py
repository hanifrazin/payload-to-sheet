import argparse
import sys
from src.convert_json import process_json_to_excel
from src.extract_payload import extract_payloads_from_excel

def main():
    parser = argparse.ArgumentParser(description="Payload-to-Sheet Tool")
    subparsers = parser.add_subparsers(dest='command', help='Pilih mode')

    # Sub-command: extract
    p_extract = subparsers.add_parser('extract', help='Ekstrak JSON dari Excel')
    p_extract.add_argument('-i', '--input', required=True, help='Path file .xlsx')
    p_extract.add_argument('-c', '--col', default='Request', help='Nama kolom')
    p_extract.add_argument('-s', '--sheet', help='Nama sheet (opsional)')

    # Sub-command: convert
    p_convert = subparsers.add_parser('convert', help='Konversi JSON ke Excel')
    p_convert.add_argument('-i', '--input', required=True, help='Path folder JSON')
    p_convert.add_argument('-o', '--output', help='Nama file output')

    args = parser.parse_args()

    if args.command == 'extract':
        extract_payloads_from_excel(args.input, args.col, args.sheet)
    elif args.command == 'convert':
        process_json_to_excel(args.input, args.output)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()