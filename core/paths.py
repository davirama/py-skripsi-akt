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


def template_berita_acara() -> Path:
    """
    Template Berita Acara + Nilai. Satu file untuk 1 atau 2 pembimbing:
    baris Pembimbing 2 hanya muncul jika `pembimbing_2` terisi.
    """
    return resource_path("resources/template_berita_acara_dan_nilai.docx")


def template_nota_dinas() -> Path:
    """
    Template Undangan Nota Dinas. Satu file untuk 1 atau 2 pembimbing.
    """
    return resource_path("resources/template_undangan_nota_dinas.docx")
