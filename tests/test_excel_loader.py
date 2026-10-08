import pandas as pd
import pytest
import requests

from core import excel_loader
from core.excel_loader import (
    Dosen,
    _build_gviz_csv_url,
    _parse_dosen_dataframe,
    load_dosen_excel,
    sync_dosen_from_online,
)

CSV = "nama,jenis_id,id\nDr. Budi,NIP,111\nDr. Citra,NUP,222\n"


def test_build_gviz_url_meng_encode_nama_sheet():
    url = _build_gviz_csv_url("SID", "Dosen Mat")
    assert url.startswith("https://docs.google.com/spreadsheets/d/SID/gviz/tq")
    assert "sheet=Dosen%20Mat" in url and "out:csv" in url


def test_parse_dataframe_normal():
    df = pd.DataFrame({"Nama": ["A"], "Jenis_ID": ["NIP"], "ID": ["1"]})
    by_id, display = _parse_dosen_dataframe(df)
    assert by_id == {"1": Dosen("A", "NIP", "1")}
    assert display == {"A — NIP: 1": "1"}


def test_parse_dataframe_lewati_baris_tanpa_nama_atau_id():
    df = pd.DataFrame(
        {"nama": ["A", "", "C"], "jenis_id": ["NIP"] * 3, "id": ["1", "2", ""]}
    )
    by_id, _ = _parse_dosen_dataframe(df)
    assert list(by_id) == ["1"]


def test_parse_dataframe_kolom_kurang():
    with pytest.raises(ValueError, match="Kolom harus ada"):
        _parse_dosen_dataframe(pd.DataFrame({"nama": ["A"]}))


def test_load_excel_menjaga_id_sebagai_string(tmp_path):
    f = tmp_path / "d.xlsx"
    pd.DataFrame(
        {"nama": ["A"], "jenis_id": ["NIP"], "id": ["0198001"]}
    ).to_excel(f, index=False)
    by_id, _ = load_dosen_excel(f)
    assert "0198001" in by_id


def test_dosen_xlsx_bawaan_valid():
    from core.paths import resource_path

    by_id, display = load_dosen_excel(resource_path("resources/dosen.xlsx"))
    assert len(by_id) > 0
    assert len(display) == len(by_id)


class FakeResp:
    def __init__(self, content=CSV.encode(), ctype="text/csv; charset=utf-8", err=None):
        self.content = content
        self.headers = {"Content-Type": ctype}
        self._err = err

    def raise_for_status(self):
        if self._err:
            raise self._err


def test_sync_sukses_menulis_cache(tmp_path, monkeypatch):
    monkeypatch.setattr(excel_loader.requests, "get", lambda *a, **k: FakeResp())
    cache = tmp_path / "sub" / "dosen.xlsx"
    assert sync_dosen_from_online(cache) == cache
    by_id, _ = load_dosen_excel(cache)
    assert set(by_id) == {"111", "222"}


def test_sync_response_html_ditolak_dan_cache_tidak_tertimpa(tmp_path, monkeypatch):
    cache = tmp_path / "dosen.xlsx"
    pd.DataFrame({"nama": ["Lama"], "jenis_id": ["NIP"], "id": ["9"]}).to_excel(
        cache, index=False
    )
    monkeypatch.setattr(
        excel_loader.requests,
        "get",
        lambda *a, **k: FakeResp(b"<html>", "text/html"),
    )
    with pytest.raises(ConnectionError, match="bukan CSV"):
        sync_dosen_from_online(cache)
    assert set(load_dosen_excel(cache)[0]) == {"9"}


def test_sync_http_error(tmp_path, monkeypatch):
    err = requests.exceptions.HTTPError("403")
    monkeypatch.setattr(
        excel_loader.requests, "get", lambda *a, **k: FakeResp(err=err)
    )
    with pytest.raises(ConnectionError, match="Gagal mengakses"):
        sync_dosen_from_online(tmp_path / "d.xlsx")


def test_sync_data_kosong_atau_kolom_salah_tidak_menimpa(tmp_path, monkeypatch):
    monkeypatch.setattr(
        excel_loader.requests, "get", lambda *a, **k: FakeResp(b"nama,jenis_id,id\n")
    )
    with pytest.raises(ValueError, match="kosong"):
        sync_dosen_from_online(tmp_path / "a.xlsx")

    monkeypatch.setattr(
        excel_loader.requests, "get", lambda *a, **k: FakeResp(b"x,y\n1,2\n")
    )
    with pytest.raises(ValueError, match="Kolom harus ada"):
        sync_dosen_from_online(tmp_path / "b.xlsx")
    assert not (tmp_path / "b.xlsx").exists()
