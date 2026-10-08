import pytest
from PySide6.QtCore import QDate

from core.date_formatter import (
    format_tanggal_indonesia,
    format_tanggal_iso,
    nama_hari_indonesia,
    urutan_ke_kata,
)


def test_format_tanggal_indonesia():
    assert format_tanggal_indonesia(QDate(2025, 12, 23)) == "23 Desember 2025"
    assert format_tanggal_indonesia(QDate(2026, 1, 5)) == "5 Januari 2026"


def test_format_tanggal_iso_nol_di_depan():
    assert format_tanggal_iso(QDate(2026, 1, 5)) == "2026-01-05"


@pytest.mark.parametrize(
    "tanggal,hari",
    [
        (QDate(2025, 12, 22), "Senin"),
        (QDate(2025, 12, 23), "Selasa"),
        (QDate(2025, 12, 24), "Rabu"),
        (QDate(2025, 12, 25), "Kamis"),
        (QDate(2025, 12, 26), "Jumat"),
        (QDate(2025, 12, 27), "Sabtu"),
        (QDate(2025, 12, 28), "Minggu"),
    ],
)
def test_nama_hari(tanggal, hari):
    assert nama_hari_indonesia(tanggal) == hari


@pytest.mark.parametrize(
    "n,kata", [(1, "Pertama"), (2, "Kedua"), (10, "Kesepuluh"), (11, "ke-11")]
)
def test_urutan_ke_kata(n, kata):
    assert urutan_ke_kata(n) == kata
