from dataclasses import dataclass
from pathlib import Path
import pandas as pd


@dataclass(frozen=True)
class Dosen:
    nama: str
    jenis_id: str  # NIP / NUP
    id: str        # nomor


def load_dosen_excel(path: Path) -> tuple[dict[str, Dosen], dict[str, str]]:
    """
    Return:
      - dosen_by_id: {id: Dosen}
      - display_to_id: {"Nama — NIP: 123": "123"}
    """
    df = pd.read_excel(path, dtype=str).fillna("")
    df.columns = [c.strip().lower() for c in df.columns]

    required = {"nama", "jenis_id", "id"}
    if not required.issubset(set(df.columns)):
        raise ValueError("Kolom Excel harus ada: nama, jenis_id, id")

    dosen_by_id: dict[str, Dosen] = {}
    display_to_id: dict[str, str] = {}

    for _, row in df.iterrows():
        nama = str(row["nama"]).strip()
        jenis = str(row["jenis_id"]).strip()
        idnum = str(row["id"]).strip()

        if not nama or not idnum:
            continue

        d = Dosen(nama=nama, jenis_id=jenis, id=idnum)
        dosen_by_id[idnum] = d

        display = f"{nama} — {jenis}: {idnum}"
        display_to_id[display] = idnum

    return dosen_by_id, display_to_id
