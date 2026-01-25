import argparse
from src.convert_json import process_json_to_excel


def main():
    # Inisialisasi parser utama
    parser = argparse.ArgumentParser(
        description="Payload-to-Sheet: CLI Tool untuk konversi JSON ke Excel"
    )

    # Langsung tambahkan argument di sini (bukan di subparser)
    parser.add_argument("-i", "--input", type=str, required=True,
                        help="Path ke folder input yang berisi file JSON")

    parser.add_argument("-o", "--output", type=str,
                        help="Nama file hasil konversi di folder result (opsional)")

    # Parsing argumen
    args = parser.parse_args()

    # Langsung jalankan fungsi karena tidak ada sub-command lagi
    process_json_to_excel(args.input, args.output)


if __name__ == "__main__":
    main()