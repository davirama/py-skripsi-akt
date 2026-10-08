import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


@pytest.fixture
def context_ba() -> dict:
    return {
        "hari": "Selasa",
        "tanggal_bulan_tahun": "23 Desember 2025",
        "jam_mulai": "08:00",
        "jam_selesai": "10:00",
        "urutan": "Pertama",
        "nama_mahasiswa": "Andi Wijaya",
        "npm": "21120123",
        "judul_skripsi": "Analisis Risiko Klaim Asuransi Jiwa",
        "pembimbing_1": "Dr. Budi Santoso",
        "pembimbing_2": "Dr. Citra Dewi",
        "penguji_1": "Prof. Dedi Kurnia",
        "penguji_2": "Dr. Eka Putri",
        "nipnup_penguji1": "NIP",
        "nomor_nipnup_penguji1": "198001012005011001",
        "nipnup_penguji2": "NUP",
        "nomor_nipnup_penguji2": "9990001112",
    }


@pytest.fixture
def context_nd(context_ba) -> dict:
    return {
        **context_ba,
        "id_nd": "123ND",
        "prodi": "Ilmu Aktuaria",
        "lokasi_ujian": "Ruang Sidang Gedung A",
        "tanggal_hari_ini": "20 Desember 2025",
        "tahun_nd": "2025",
    }
