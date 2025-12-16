from dataclasses import dataclass


@dataclass
class FormData:
    nama_mahasiswa: str
    npm: str
    judul_skripsi: str
    urutan: int
    hari: str
    jam_mulai: str
    jam_selesai: str
    pembimbing_1: str
    pembimbing_2: str
    penguji_1: str
    penguji_2: str


def validate_form(d: FormData) -> tuple[bool, str]:
    if not d.nama_mahasiswa.strip():
        return False, "Nama mahasiswa wajib diisi."
    if not d.npm.strip():
        return False, "NPM wajib diisi."
    if not d.judul_skripsi.strip():
        return False, "Judul skripsi wajib diisi."
    if not d.hari.strip():
        return False, "Hari wajib terisi (akan auto dari tanggal, tapi jangan kosong)."
    if not d.pembimbing_1.strip():
        return False, "Pembimbing 1 wajib dipilih."
    if not d.penguji_1.strip() or not d.penguji_2.strip():
        return False, "Penguji 1 dan Penguji 2 wajib dipilih."
    if d.penguji_1.strip() == d.penguji_2.strip():
        return False, "Penguji 1 dan Penguji 2 tidak boleh orang yang sama."
    if d.pembimbing_2.strip() and d.pembimbing_2.strip() == d.pembimbing_1.strip():
        return False, "Pembimbing 2 tidak boleh sama dengan Pembimbing 1."
    if d.jam_selesai <= d.jam_mulai:
        return False, "Jam selesai harus lebih besar dari jam mulai."
    return True, ""
