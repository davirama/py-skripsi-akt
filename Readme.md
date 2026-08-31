📄 Generator Berita Acara & Nota Dinas Ujian Skripsi (S1)

Aplikasi desktop berbasis Python + PySide6 untuk menghasilkan:

✅ Berita Acara dan Nilai Ujian Skripsi

✅ Undangan Nota Dinas Ujian Skripsi

✅ Otomatis pilih template 1 atau 2 pembimbing

✅ Data dosen dari Excel

✅ Output .docx siap cetak

✅ Bisa dibuild menjadi file .exe (tanpa perlu Python di komputer lain)

📁 Struktur Project
project/
│
├─ main.py
├─ ui_main.py
├─ core/
│ ├─ paths.py
│ ├─ word_generator.py
│ ├─ date_formatter.py
│ ├─ validator.py
│ └─ excel_loader.py
│
├─ resources/
│ ├─ dosen.xlsx
│ ├─ template_berita_acara_dan_nilai_1pembimbing.docx
│ ├─ template_berita_acara_dan_nilai_2pembimbing.docx
│ ├─ template_undangan_nota_dinas_1pembimbing.docx
│ └─ template_undangan_nota_dinas_2pembimbing.docx
│
├─ .venv/
├─ README.md

⚙️ Setup Environment (Sekali di Komputer Dev)
1️⃣ Aktifkan Virtual Environment
.\.venv\Scripts\Activate.ps1

Pastikan prompt berubah menjadi:

(.venv) PS ...

2️⃣ Install Dependency
pip install -r requirements.txt

Atau manual:

pip install pyside6 pandas openpyxl docxtpl pyinstaller

▶️ Menjalankan Aplikasi (Mode Development)
python main.py

🏗️ Build Menjadi File EXE (Windows)

Catatan penting:

Pastikan EXE lama tidak sedang berjalan

Disarankan tutup Explorer di folder dist/

Jika error Access is denied, hapus folder dist/ dan build/

✅ Langkah Build yang BENAR
1️⃣ Aktifkan venv
.\.venv\Scripts\Activate.ps1

2️⃣ Jalankan PyInstaller
pyinstaller `  --noconsole`
--onefile `  --name "BeritaAcaraSkripsi"`
--clean `  --hidden-import docxtpl`
--hidden-import jinja2 `  --hidden-import lxml`
--hidden-import requests ` --hidden-import requests`
--add-data "resources;resources" `
main.py

📦 Hasil Build

Setelah sukses, file akan muncul di:

dist/
└─ BeritaAcaraSkripsi.exe

✅ File ini bisa dijalankan langsung di komputer lain
❌ Tidak perlu install Python / pip / library apa pun

📝 Format Nama File Output
1️⃣ Berita Acara
Berita Acara dan Nilai Ujian Skripsi_Nama Mahasiswa_NPM.docx

2️⃣ Nota Dinas
YYYY-MM-DD_NoID_Undangan_Ujian_Skripsi_S1_Prodi_Nama Mahasiswa_NPM.docx

Contoh:

2025-12-23_123ND_Undangan_Ujian_Skripsi_S1_Matematika_Andi Wijaya_21120123.docx

🛠️ Troubleshooting
❌ Error: PermissionError: [WinError 5] Access is denied

Solusi:

Remove-Item .\dist -Recurse -Force
Remove-Item .\build -Recurse -Force

Lalu build ulang.

Jika masih terjadi:

Tambahkan Windows Defender Exclusion untuk folder project.

❌ Template tidak ditemukan saat EXE dijalankan

Pastikan:

Folder resources/ ikut dibundle

Build pakai:

--add-data "resources;resources"

1. Aktifkan venv
   .\.venv\Scripts\Activate.ps1

2. Build EXE (onefile + tanpa console + include semua resources)
   pyinstaller `
   > > --noconsole `  --onefile`
   > > --name "BeritaAcaraSkripsi" `  --clean`
   > > --hidden-import docxtpl `  --hidden-import jinja2`
   > > --hidden-import lxml `  --hidden-import docxtpl`
   > > --hidden-import jinja2 `  --hidden-import lxml`
   > > --hidden-import jinja2 `  --hidden-import lxml`
   > > --add-data "resources;resources" `  --hidden-import lxml`
   > > --add-data "resources;resources" `
   > > main.py
