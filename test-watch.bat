@echo off
rem Jalankan test otomatis setiap ada file .py / .docx / .xlsx yang berubah.
rem Tutup dengan Ctrl+C. Tekan 'f' di jendela untuk hanya menjalankan test yang gagal.
cd /d "%~dp0"
".venv\Scripts\ptw.exe" . --now --clear --patterns "*.py,*.docx,*.xlsx" --ignore-patterns "output/*,build/*,dist/*,.venv/*" --runner ".venv\Scripts\python.exe -m pytest"
