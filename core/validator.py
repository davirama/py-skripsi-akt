from __future__ import annotations

from dataclasses import dataclass


@dataclass
class FormData:
    nama_mahasiswa: str
    npm: str
    judul_skripsi: str
    urutan: int
    hari: str
    jam_mulai: str  # "HH:mm"
    jam_selesai: str  # "HH:mm"
    pembimbing_1: str
    pembimbing_2: str
    penguji_1: str
    penguji_2: str


def _parse_hhmm(t: str) -> tuple[int, int] | None:
    """
    Parse 'HH:mm' -> (HH, mm). Return None kalau format tidak valid.
    """
    try:
        hh, mm = t.split(":")
        return int(hh), int(mm)
    except Exception:
        return None


def validate_form(d: FormData) -> tuple[bool, str]:
    # basic
    if not d.nama_mahasiswa.strip():
        return False, "Nama mahasiswa wajib diisi."
    if not d.npm.strip():
        return False, "NPM wajib diisi."
    if not d.judul_skripsi.strip():
        return False, "Judul skripsi wajib diisi."
    if not d.hari.strip():
        return False, "Hari wajib terisi (auto dari tanggal, tapi jangan kosong)."

    # dosen
    if not d.pembimbing_1.strip():
        return False, "Pembimbing 1 wajib dipilih."
    if not d.penguji_1.strip() or not d.penguji_2.strip():
        return False, "Penguji 1 dan Penguji 2 wajib dipilih."
    if d.penguji_1.strip() == d.penguji_2.strip():
        return False, "Penguji 1 dan Penguji 2 tidak boleh orang yang sama."
    if d.pembimbing_2.strip() and d.pembimbing_2.strip() == d.pembimbing_1.strip():
        return False, "Pembimbing 2 tidak boleh sama dengan Pembimbing 1."

    # waktu
    a = _parse_hhmm(d.jam_mulai.strip())
    b = _parse_hhmm(d.jam_selesai.strip())
    if not a or not b:
        return False, "Format jam harus HH:mm (contoh 09:30)."

    if (b[0], b[1]) <= (a[0], a[1]):
        return False, "Jam selesai harus lebih besar dari jam mulai."

    return True, ""


def validate_nota_dinas_inputs(
    id_nd: str, lokasi_ujian: str, prodi: str
) -> tuple[bool, str]:
    if not id_nd.strip():
        return False, "ID ND wajib diisi."
    if not lokasi_ujian.strip():
        return False, "Lokasi ujian wajib diisi."
    if not prodi.strip():
        return False, "Prodi wajib diisi."
    return True, ""
