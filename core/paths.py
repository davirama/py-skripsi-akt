# core/paths.py
from __future__ import annotations

import sys
from pathlib import Path


def app_root() -> Path:
    # .../project_root/core/paths.py -> project_root
    return Path(__file__).resolve().parents[1]


def resource_path(relative: str) -> Path:
    """
    Support dev mode + PyInstaller onefile mode.
    """
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative
    return app_root() / relative


def _pilih_template(jumlah_pembimbing: int, file_1: str, file_2: str) -> Path:
    if jumlah_pembimbing == 1:
        return resource_path(f"resources/{file_1}")
    if jumlah_pembimbing == 2:
        return resource_path(f"resources/{file_2}")
    raise ValueError("Jumlah pembimbing tidak valid (harus 1 atau 2)")


def pilih_template_berdasarkan_pembimbing(jumlah_pembimbing: int) -> Path:
    """
    Template Berita Acara + Nilai.
    """
    return _pilih_template(
        jumlah_pembimbing,
        "template_berita_acara_dan_nilai_1pembimbing.docx",
        "template_berita_acara_dan_nilai_2pembimbing.docx",
    )


def pilih_template_nota_dinas_berdasarkan_pembimbing(jumlah_pembimbing: int) -> Path:
    """
    Template Undangan Nota Dinas.
    """
    return _pilih_template(
        jumlah_pembimbing,
        "template_undangan_nota_dinas_1pembimbing.docx",
        "template_undangan_nota_dinas_2pembimbing.docx",
    )
