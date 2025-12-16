import sys
from pathlib import Path


def app_root() -> Path:
    # folder tempat main.py berada (saat dev)
    return Path(__file__).resolve().parents[1]


def resource_path(relative: str) -> Path:
    """
    Support dev mode + PyInstaller onefile mode.
    """
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative
    return app_root() / relative
