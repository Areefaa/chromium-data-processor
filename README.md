# 🧪 Chromium Experiment Data Processor

> Mengotomatiskan konversi dataset percobaan chromium dari format CSV ke laporan Excel (`.xlsx`) yang terformat profesional.

---

## 📋 Deskripsi

Script Python ini membaca **12 file CSV** hasil percobaan zat chromium, lalu menghasilkan satu file `.xlsx` yang berisi:

| Sheet | Isi |
|---|---|
| **Ringkasan** | Daftar semua file, jumlah baris/kolom, nama kolom |
| **Per-Eksperimen** | Setiap CSV mendapat sheet tersendiri |
| **Semua Data** | Gabungan seluruh CSV dalam satu tabel |

---

## 🗂️ Struktur Folder

```
chromium-data-processor/
├── data/
│   └── raw/               ← Letakkan 12 file CSV di sini
│       ├── eksperimen_01.csv
│       ├── eksperimen_02.csv
│       └── ...
├── output/                ← File .xlsx hasil generate (auto-dibuat)
├── chromium_to_excel.py   ← Script utama
├── requirements.txt
└── README.md
```

---

## 🚀 Cara Penggunaan

### 1. Clone repository
```bash
git clone https://github.com/username/chromium-data-processor.git
cd chromium-data-processor
```

### 2. Install dependensi
```bash
pip install -r requirements.txt
```

### 3. Letakkan CSV
Salin 12 file CSV ke dalam folder `data/raw/`.

### 4. Jalankan script
```bash
python chromium_to_excel.py
```

### 5. Ambil hasilnya
File tersimpan di `output/chromium_data.xlsx`.

---

## 🛠️ Teknologi

- **Python 3.10+**
- **pandas** – membaca dan menggabungkan CSV
- **openpyxl** – formatting dan penulisan Excel

---

## 📁 Branch

| Branch | Fungsi |
|---|---|
| `main` | Kode stabil, siap pakai |
| `dev` | Pengembangan fitur baru |
| `feature/chart` | Menambahkan visualisasi grafik ke Excel *(opsional)* |
| `feature/statistics` | Menambahkan sheet statistik deskriptif *(opsional)* |

---

## 📄 Lisensi

MIT License
