# core/paths.py
import sys
from pathlib import Path


def app_root() -> Path:
    return Path(__file__).resolve().parents[1]


def resource_path(relative: str) -> Path:
    """
    Support dev mode + PyInstaller onefile mode.
    """
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative
    return app_root() / relative


def pilih_template_berdasarkan_pembimbing(jumlah_pembimbing: int) -> Path:
    if jumlah_pembimbing == 1:
        return resource_path("resources/template_berita_acara_dan_nilai_1pembimbing.docx")
    elif jumlah_pembimbing == 2:
        return resource_path("resources/template_berita_acara_dan_nilai_2pembimbing.docx")
    else:
        raise ValueError("Jumlah pembimbing tidak valid")
