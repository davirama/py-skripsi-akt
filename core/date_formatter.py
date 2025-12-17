from PySide6.QtCore import QDate
from datetime import date


# =========================
# Konstanta Bulan Indonesia
# =========================
_BULAN_ID = {
    1: "Januari",
    2: "Februari",
    3: "Maret",
    4: "April",
    5: "Mei",
    6: "Juni",
    7: "Juli",
    8: "Agustus",
    9: "September",
    10: "Oktober",
    11: "November",
    12: "Desember",
}


# =========================
# Format tanggal dari QDate
# Contoh: 12 Januari 2025
# =========================
def format_tanggal_indonesia(qdate: QDate) -> str:
    return f"{qdate.day()} {_BULAN_ID[qdate.month()]} {qdate.year()}"


# =========================
# Nama hari dari QDate
# =========================
def nama_hari_indonesia(qdate: QDate) -> str:
    # Qt: 1=Senin ... 7=Minggu
    hari = {
        1: "Senin",
        2: "Selasa",
        3: "Rabu",
        4: "Kamis",
        5: "Jumat",
        6: "Sabtu",
        7: "Minggu",
    }
    return hari.get(qdate.dayOfWeek(), "")


# =========================
# Urutan ujian → kata
# =========================
def urutan_ke_kata(n: int) -> str:
    mapping = {
        1: "Pertama",
        2: "Kedua",
        3: "Ketiga",
        4: "Keempat",
        5: "Kelima",
        6: "Keenam",
        7: "Ketujuh",
        8: "Kedelapan",
        9: "Kesembilan",
        10: "Kesepuluh",
    }

    if n in mapping:
        return mapping[n]

    # fallback aman untuk > 10
    return f"ke-{n}"


# =========================
# Tanggal hari ini (auto)
# Contoh: 15 Desember 2025
# =========================
def format_tanggal_hari_ini_indonesia() -> str:
    today = date.today()
    return f"{today.day} {_BULAN_ID[today.month]} {today.year}"
