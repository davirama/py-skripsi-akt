# Desain: Refactor Terstruktur + Fitur Baru (Generator Berita Acara & Nota Dinas)

Tanggal: 2026-10-08
Status: menunggu review

## Tujuan

Membuat aplikasi lebih efisien, kodenya rapi dan terstruktur, lalu menambah fitur
riwayat, cek bentrok, ekspor PDF, dan impor massal. Aplikasi dipakai **lokal** (satu
komputer), jadi tidak ada database bersama.

## Keputusan yang sudah disepakati

- Semua 4 perbaikan efisiensi dan 4 fitur baru dikerjakan.
- Penyimpanan lokal saja (SQLite di AppData).
- Kode distruktur ulang menjadi paket dengan tanggung jawab jelas.
- Dikerjakan dalam 3 fase; tiap fase selesai dan bisa dipakai sebelum fase berikutnya.

## Struktur target

```
main.py                      # entry point tipis
pyproject.toml / requirements.txt
resources/                   # template .docx + dosen.xlsx (fallback)
tests/
app/
  config.py                  # load/save config.json (AppData), default tertanam
  paths.py                   # app_root, resource_path, user_data_dir, pemilih template
  models.py                  # dataclass: Dosen, UjianInput, NotaDinasInput
  services/
    dosen_repo.py            # load xlsx, sync gviz, cache AppData, diff logging
    validation.py            # validate_ujian(), validate_nota_dinas(), validate_all()
    context_builder.py       # bangun context BA & ND dari satu UjianInput
    document_service.py      # render docxtpl, penamaan file, sanitize
    history_repo.py          # SQLite: simpan/cari/ubah riwayat, saran ID ND
    conflict_checker.py      # cek bentrok dosen+tanggal+slot
    pdf_export.py            # docx -> pdf (docx2pdf), deteksi Word
    bulk_import.py           # baca Excel massal, validasi per baris, template unduhan
    dates_id.py              # format tanggal/hari/urutan Indonesia (tanpa dependensi Qt)
  ui/
    main_window.py           # susun layout & sambungkan sinyal
    widgets.py               # NoWheelComboBox, NoWheelDateEdit, combo dosen
    form_panel.py            # form ujian & nota dinas -> UjianInput
    history_panel.py         # tab riwayat
    bulk_dialog.py           # dialog pratinjau impor massal
    workers.py               # QThread: sync dosen, impor massal
    styles.py                # stylesheet
```

Aturan dependensi: `services/` tidak boleh mengimpor PySide6 (`dates_id` memakai
`datetime.date`, bukan `QDate`). UI hanya memanggil services. Ini membuat logika bisa
diuji tanpa GUI.

## Fase 1: Fondasi dan efisiensi

1. **Struktur ulang** sesuai pohon di atas; `core/` dan `ui_main.py` dipindah/dipecah,
   `core/init.py` dihapus (diganti `__init__.py` per paket). Perilaku tidak berubah.
2. **Config** (`config.json` di `%APPDATA%\BeritaAcaraSkripsi\`): ID spreadsheet, nama
   sheet, slot jam, daftar prodi, folder output terakhir. Default tertanam; file rusak
   atau hilang -> pakai default dan log warning.
3. **Cache dosen** di `%APPDATA%\BeritaAcaraSkripsi\dosen.xlsx`. Urutan baca: cache ->
   `resources/dosen.xlsx` (fallback). Sync menulis ke cache, bukan ke `_MEIPASS`.
   Validasi struktur sebelum menimpa cache tetap dipertahankan.
4. **Satu `UjianInput`** (dataclass) dari form; `context_builder` membuat context BA dan
   ND darinya. Duplikasi di UI dihapus.
5. **`validate_all()`** dijalankan sebelum dokumen apa pun dibuat. "Generate Semua" tidak
   lagi meninggalkan file setengah jadi. Nota Dinas ikut divalidasi penuh. Validasi jam
   yang tidak pernah gagal dihapus (jam berasal dari slot tetap).
6. **Housekeeping**: hapus `__pycache__` dari git (`git rm --cached`), rapikan Readme
   (perintah build duplikat), perbarui `.spec`, tambah `pytest` ke dev dependency.

## Fase 2: Riwayat, bentrok, PDF

- **Skema SQLite** (`history.db`, AppData): tabel `ujian` berisi semua field input form
  (mahasiswa, tanggal, slot, id dosen pb1/pb2/pj1/pj2, ID ND, prodi, lokasi), `created_at`,
  `updated_at`, `path_ba`, `path_nd`. Versi skema disimpan di `PRAGMA user_version`.
- **Riwayat**: tab dengan tabel yang bisa dicari (nama/NPM). Pilih baris -> isi ulang form
  untuk diedit/generate ulang (menimpa baris yang sama, bukan menambah duplikat).
  Setiap generate sukses menyimpan/memperbarui baris.
- **Saran ID ND**: ID terakhir di riwayat + 1 jika polanya numerik; kalau tidak, kosong.
  Selalu bisa diubah manual.
- **Cek bentrok**: sebelum generate, cari ujian lain dengan tanggal dan slot sama yang
  melibatkan dosen yang sama (peran apa pun). Hasil berupa **peringatan** dengan tombol
  "Lanjut" / "Batal", tidak memblokir. Ujian yang sedang diedit dikecualikan.
- **Ekspor PDF**: tombol aktif hanya jika MS Word terdeteksi (docx2pdf). Berjalan di
  background thread; kegagalan menampilkan pesan, `.docx` tetap ada.

## Fase 3: Impor massal

- **Template Excel** diunduh dari aplikasi (kolom: nama, npm, judul, tanggal, slot,
  nama/id dosen pb1, pb2, pj1, pj2, id_nd, prodi, lokasi, urutan, tanggal_nd).
- **Dialog pratinjau**: tiap baris menampilkan status (valid / error validasi / bentrok).
  Baris bermasalah dilewati, sisanya diproses di background thread dengan progress bar
  dan bisa dibatalkan. Dosen dicocokkan lewat ID (NIP/NUP) atau nama persis; tidak
  cocok -> error pada baris itu.
- Hasil per baris dilaporkan di akhir (sukses/gagal + alasan) dan masuk riwayat.

## Penanganan error

- Sync, ekspor PDF, dan impor massal memakai worker thread; error dikirim lewat sinyal
  sebagai pesan teks, UI tidak pernah crash karena itu.
- File riwayat/config rusak: backup dengan akhiran `.bak`, buat baru, beri tahu pengguna.
- Template tidak ditemukan: error jelas yang menyebut path.

## Pengujian

`pytest` tanpa GUI untuk: `validation`, `config`, `dosen_repo` (parsing + diff),
`context_builder`, `history_repo` (SQLite sementara), `conflict_checker`, `bulk_import`.
`document_service` diuji dengan merender template asli ke folder sementara lalu membaca
ulang teks `.docx` untuk memastikan placeholder terisi. Verifikasi UI dilakukan manual
dengan menjalankan aplikasi.

## Risiko dan batasan

- Ekspor PDF membutuhkan MS Word terpasang.
- Riwayat lokal per komputer; tidak ada sinkronisasi antar-komputer.
- Placeholder di 4 template `.docx` belum diaudit; diaudit di awal fase 1 agar
  `context_builder` mencakup semuanya.
- Refactor besar pada `ui_main.py` (838 baris); fase 1 harus mempertahankan perilaku
  lama dan diverifikasi dengan membuat dokumen dari input yang sama sebelum/sesudah.

## Di luar cakupan (YAGNI)

Database bersama/multi-pengguna, login, pembuatan template baru dari dalam aplikasi,
dan pengiriman email.
