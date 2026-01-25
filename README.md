# 📊 Payload-to-Sheet CLI

Tool berbasis Python untuk mengonversi data JSON menjadi file Excel (.xlsx) secara otomatis. Tool ini dirancang untuk mempermudah QA dalam mapping key dan value yang ada di payload API ke dalam sebuah tabel untuk kebutuhan automation test .

---

## 🚀 Persiapan (Installation)

### 1. Prasyarat
Pastikan Anda sudah menginstall Python (versi 3.12 atau lebih baru).

### 2. Clone Repositori
```bash
git clone https://gitlab.com/username/payload-to-sheet.git
```
```bash
cd payload-to-sheet
```

### 3. **Install Dependencies**
```bash
pip install -r requirements.txt 
```

### 4. Setup Virtual Environment (Opsional tapi Disarankan)
```bash
python -m venv .venv
```
```bash
source .venv/bin/activate  # Unix/macOS
```
```bash
.venv\Scripts\activate     # Windows 
``` 

---

## 🛠 Cara Penggunaan
Gunakan perintah python main.py dengan opsi berikut:

### 1. Bantuan (Help)
Untuk melihat panduan pengguna
```bash
python main.py --help
```

### 2. Konversi JSON ke Excel
Untuk mengonversi semua JSON di dalam folder:
```bash
python main.py -i sample-data/single
```

### 3. Output Custom
Untuk menentukan nama file hasil konversi:
```bash
python main.py -i sample-data/single -o hasil_report
```

___

## 📝 Catatan Penting
* Kolom SOURCE: Tool akan selalu menambahkan kolom SOURCE di awal Excel untuk melacak asal file JSON tiap baris data.
* Format Teks: Semua data di Excel diformat sebagai Text untuk menjaga integritas data (misalnya agar angka 0812 tidak berubah menjadi 812).