from dataclasses import dataclass
from pathlib import Path
from io import BytesIO
from urllib.parse import quote
import logging

import pandas as pd
import requests

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class Dosen:
    nama: str
    jenis_id: str  # NIP / NUP
    id: str  # nomor


# ID spreadsheet diambil dari URL edit:
# https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit
DOSEN_SPREADSHEET_ID = "1d_2SjqC6WofA7qMA1wbSaWAtrY9Nm5DP"
DOSEN_SHEET_NAME = "Dosen Mat"


def _build_gviz_csv_url(spreadsheet_id: str, sheet_name: str) -> str:
    """
    Bangun URL endpoint Google Visualization API (gviz) yang membaca
    data LANGSUNG dari spreadsheet (live), bukan dari snapshot hasil
    "Publish to web" yang di-cache Google selama beberapa menit.

    Syarat: spreadsheet sudah di-share dengan akses
    "Anyone with the link" -> Viewer (tidak perlu Publish to web).
    """
    encoded_sheet = quote(sheet_name)
    return (
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
        f"/gviz/tq?tqx=out:csv&sheet={encoded_sheet}"
    )


DOSEN_SHEET_CSV_URL = _build_gviz_csv_url(DOSEN_SPREADSHEET_ID, DOSEN_SHEET_NAME)


def _parse_dosen_dataframe(df: pd.DataFrame) -> tuple[dict[str, Dosen], dict[str, str]]:
    """
    Parsing inti dari DataFrame (dipakai bareng oleh load_dosen_excel
    dan sync_dosen_from_online supaya logikanya konsisten).
    """
    df = df.fillna("")
    df.columns = [c.strip().lower() for c in df.columns]

    required = {"nama", "jenis_id", "id"}
    if not required.issubset(set(df.columns)):
        raise ValueError("Kolom harus ada: nama, jenis_id, id")

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


def load_dosen_excel(path: Path) -> tuple[dict[str, Dosen], dict[str, str]]:
    """
    Load data dosen dari file Excel lokal.

    Return:
      - dosen_by_id: {id: Dosen}
      - display_to_id: {"Nama — NIP: 123": "123"}
    """
    df = pd.read_excel(path, dtype=str)
    return _parse_dosen_dataframe(df)


def sync_dosen_from_online(
    cache_path: Path,
    url: str = DOSEN_SHEET_CSV_URL,
    timeout: int = 10,
) -> Path:
    """
    Ambil data dosen terbaru dari Google Sheets lewat endpoint gviz
    (Google Visualization API), yang membaca data LANGSUNG dari
    spreadsheet (near real-time), bukan dari snapshot "Publish to
    web" yang bisa nyangkut di cache Google selama beberapa menit.

    Lalu simpan sebagai file .xlsx ke cache_path (menimpa file lama).

    Kalau gagal (offline, sheet tidak bisa diakses / tidak public,
    atau kolom tidak sesuai), exception dilempar ke caller supaya
    file lokal lama tetap dipakai dan tidak ditimpa dengan data rusak.

    Catatan: spreadsheet harus di-share dengan akses
    "Anyone with the link" -> Viewer. Tidak perlu "Publish to web"
    sama sekali untuk endpoint ini.
    """
    logger.info("Sync dosen: fetching %s", url)

    try:
        resp = requests.get(url, timeout=timeout)
        resp.raise_for_status()
    except requests.exceptions.HTTPError as e:
        raise ConnectionError(
            "Gagal mengakses spreadsheet online. Pastikan sharing setting "
            "sudah 'Anyone with the link -> Viewer' dan nama sheet "
            f"('{DOSEN_SHEET_NAME}') sudah benar. Detail: {e}"
        ) from e

    # gviz kadang mengembalikan halaman HTML error (bukan CSV) kalau
    # sheet tidak ditemukan / akses ditolak -> deteksi dini di sini
    # supaya errornya jelas, bukan error parsing CSV yang membingungkan.
    content_type = resp.headers.get("Content-Type", "")
    if "text/csv" not in content_type and "text/plain" not in content_type:
        raise ConnectionError(
            "Response bukan CSV (kemungkinan sheet tidak ditemukan atau "
            "akses ditolak). Pastikan nama sheet "
            f"('{DOSEN_SHEET_NAME}') persis sama dengan tab di spreadsheet, "
            "dan sharing sudah 'Anyone with the link -> Viewer'."
        )

    # dtype=str penting supaya kolom id (NIP/NUP) tidak berubah jadi
    # float (mis. "123456" jadi "123456.0") saat dibaca dari CSV.
    df = pd.read_csv(BytesIO(resp.content), dtype=str)

    if df.empty:
        raise ValueError("Data dosen dari online kosong, sync dibatalkan.")

    # Validasi struktur & isi sebelum menimpa file cache lokal,
    # supaya kalau sheet online berantakan, file lama tidak rusak.
    dosen_by_id, _ = _parse_dosen_dataframe(df.copy())
    logger.info(
        "Sync dosen: berhasil parse %d baris dosen dari online", len(dosen_by_id)
    )

    # Bandingkan dengan data lokal lama (kalau ada) untuk logging,
    # supaya kelihatan jelas apakah datanya memang berubah atau tidak.
    if cache_path.exists():
        try:
            old_by_id, _ = load_dosen_excel(cache_path)
            if old_by_id == dosen_by_id:
                logger.info(
                    "Sync dosen: tidak ada perubahan data dibanding lokal terakhir."
                )
            else:
                added = set(dosen_by_id) - set(old_by_id)
                removed = set(old_by_id) - set(dosen_by_id)
                changed = {
                    i
                    for i in (set(dosen_by_id) & set(old_by_id))
                    if dosen_by_id[i] != old_by_id[i]
                }
                logger.info(
                    "Sync dosen: perubahan terdeteksi -> "
                    "%d baru, %d dihapus, %d diubah",
                    len(added),
                    len(removed),
                    len(changed),
                )
        except Exception as e:
            logger.warning(
                "Sync dosen: gagal baca data lokal lama untuk dibandingkan: %s", e
            )

    cache_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(cache_path, index=False)
    logger.info("Sync dosen: file cache ditulis ke %s", cache_path)

    return cache_path
