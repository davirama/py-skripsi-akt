# 📄 Generator Berita Acara & Nota Dinas Ujian Skripsi (S1)

Aplikasi desktop berbasis Python + PySide6 untuk menghasilkan:

- ✅ Berita Acara dan Nilai Ujian Skripsi
- ✅ Undangan Nota Dinas Ujian Skripsi
- ✅ Satu template untuk 1 atau 2 pembimbing (baris Pembimbing 2 hilang otomatis)
- ✅ Data dosen dari Google Sheets (sync online) dengan fallback Excel lokal
- ✅ Output `.docx` siap cetak
- ✅ Bisa dibuild menjadi `.exe` (tanpa perlu Python di komputer lain)

## 📁 Struktur Project

```
├─ main.py                 # entry point
├─ ui_main.py              # GUI (PySide6)
├─ core/
│  ├─ paths.py             # path resource + pemilihan template
│  ├─ word_generator.py    # render template .docx (docxtpl)
│  ├─ date_formatter.py    # tanggal/hari/urutan bahasa Indonesia
│  ├─ validator.py         # validasi form
│  └─ excel_loader.py      # load Excel + sync Google Sheets
├─ resources/
│  ├─ dosen.xlsx
│  ├─ template_berita_acara_dan_nilai.docx   # 1 file untuk 1 / 2 pembimbing
│  └─ template_undangan_nota_dinas.docx      # 1 file untuk 1 / 2 pembimbing
├─ tests/                  # pytest
├─ pytest.ini
├─ test-watch.bat          # test otomatis (watch mode)
├─ requirements.txt        # dependency aplikasi
└─ requirements-dev.txt    # + pytest, pytest-watcher
```

## ⚙️ Setup (sekali di komputer dev)

```powershell
# 1. Aktifkan virtual environment
.\.venv\Scripts\Activate.ps1

# 2. Install dependency (aplikasi + testing)
pip install -r requirements-dev.txt
```

Jika `.venv` belum ada: `python -m venv .venv` lalu ulangi langkah di atas.

## ▶️ Menjalankan Aplikasi

```powershell
python main.py
```

## 🧪 Testing

Test memakai **pytest** dan tidak membuka GUI. Pembuatan dokumen dites dengan
merender template asli lalu membaca ulang isinya, jadi perubahan template yang
merusak placeholder akan ketahuan.

### Sekali jalan

```powershell
# dengan venv aktif
pytest

# tanpa mengaktifkan venv
.\.venv\Scripts\python.exe -m pytest
```

Opsi yang sering dipakai:

| Perintah | Fungsi |
|---|---|
| `pytest tests/test_validator.py` | Jalankan satu file saja |
| `pytest -k "nota_dinas"` | Jalankan test yang namanya mengandung kata tertentu |
| `pytest -x` | Berhenti di kegagalan pertama |
| `pytest --lf` | Ulangi hanya test yang gagal terakhir kali |

### Otomatis (watch mode)

Klik dua kali **`test-watch.bat`**, atau jalankan di terminal:

```powershell
.\test-watch.bat
```

Test dijalankan lagi setiap kali ada file `.py`, `.docx`, atau `.xlsx` yang
disimpan, dan hasilnya tampil langsung di jendela. Tekan `f` di jendela itu
untuk hanya menjalankan test yang gagal, dan `Ctrl+C` untuk berhenti.

### Isi test

| File | Yang dites |
|---|---|
| `test_validator.py` | Aturan validasi form dan nota dinas |
| `test_date_formatter.py` | Format tanggal, nama hari, nomor urut |
| `test_paths.py` | Template ditemukan |
| `test_excel_loader.py` | Parsing data dosen dan sync online (request dipalsukan) |
| `test_word_generator.py` | Render 2 template asli untuk kasus 1 dan 2 pembimbing, serta penamaan file output |

## 📝 Mengedit Template Word

Hanya ada **satu template per dokumen**. Aplikasi mengisi `pembimbing_2` dengan
string kosong bila hanya ada 1 pembimbing, dan baris Pembimbing 2 otomatis hilang.

Hal yang perlu diperhatikan saat mengedit `resources/*.docx`:

- Jangan hapus baris kecil yang berisi `{%tr if pembimbing_2 %}` dan
  `{%tr endif %}` (di atas dan di bawah baris Pembimbing 2 pada tabel), serta
  paragraf `{%p if pembimbing_2 %}` / `{%p endif %}` di Nota Dinas. Baris-baris
  itu tidak ikut tercetak, tetapi menentukan baris mana yang kondisional.
- Teks Pembimbing 2 (`{{pembimbing_2}}`) harus tetap berada di antara kedua tag.
- Satu perubahan format cukup dilakukan sekali; berlaku untuk 1 maupun 2 pembimbing.
- Setelah mengedit, jalankan `pytest` (atau `test-watch.bat`). Test akan merender
  template untuk kedua kasus dan memastikan tidak ada placeholder yang tersisa.

## 🏗️ Build menjadi EXE (Windows)

Catatan:

- Pastikan EXE lama tidak sedang berjalan.
- Tutup Explorer yang membuka folder `dist/`.
- Jika muncul `Access is denied`, hapus folder `dist/` dan `build/`.

```powershell
.\.venv\Scripts\Activate.ps1

pyinstaller `
  --noconsole `
  --onefile `
  --name "BeritaAcaraSkripsi" `
  --clean `
  --hidden-import docxtpl `
  --hidden-import jinja2 `
  --hidden-import lxml `
  --hidden-import requests `
  --add-data "resources;resources" `
  main.py
```

Atau cukup `pyinstaller BeritaAcaraSkripsi.spec` karena konfigurasinya sudah
tersimpan di file tersebut.

Hasilnya: `dist/BeritaAcaraSkripsi.exe`. File ini bisa dijalankan langsung di
komputer lain tanpa Python atau library tambahan.

> Jalankan `pytest` sebelum build supaya `.exe` tidak membawa bug yang sudah
> tertangkap test.

## 📝 Format Nama File Output

Output disimpan di `output/<Nama>_<NPM>/`.

1. **Berita Acara**
   `Berita Acara dan Nilai Ujian Skripsi_<Nama>_<NPM>.docx`
2. **Nota Dinas**
   `YYYY-MM-DD_<ID ND>_Undangan_Ujian_Skripsi_S1_<Prodi>_<Nama>_<NPM>.docx`

   Contoh: `2025-12-23_123ND_Undangan_Ujian_Skripsi_S1_Matematika_Andi Wijaya_21120123.docx`

## 🛠️ Troubleshooting

**`PermissionError: [WinError 5] Access is denied` saat build**

```powershell
Remove-Item .\dist -Recurse -Force
Remove-Item .\build -Recurse -Force
```

Lalu build ulang. Jika masih terjadi, tambahkan folder project ke Windows Defender
Exclusion.

**Template tidak ditemukan saat EXE dijalankan**

Pastikan build memakai `--add-data "resources;resources"`.

**`pytest` tidak dikenali**

Aktifkan venv dulu, atau pakai `.\.venv\Scripts\python.exe -m pytest`. Jika modul
belum terpasang: `pip install -r requirements-dev.txt`.

**Sync dosen gagal**

Pastikan Google Sheets di-share dengan akses *Anyone with the link → Viewer* dan
nama tab sheet persis `Dosen Mat`. Saat sync gagal, aplikasi tetap memakai
`resources/dosen.xlsx`.
