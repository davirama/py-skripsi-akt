import re

import pytest
from docx import Document

from core.paths import template_berita_acara, template_nota_dinas
from core.word_generator import generate_docx, sanitize_filename


def teks_docx(path) -> str:
    d = Document(str(path))
    bagian = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            for c in row.cells:
                bagian.append(c.text)
    return "\n".join(bagian)


def jumlah_baris_dengan(path, teks: str) -> int:
    """Hitung baris tabel yang memuat `teks` (untuk memastikan baris Pembimbing 2 hilang)."""
    d = Document(str(path))
    return sum(
        1
        for t in d.tables
        for row in t.rows
        if teks in "".join(c.text for c in row.cells)
    )


def tanpa_pembimbing_2(ctx: dict) -> dict:
    return {**ctx, "pembimbing_2": ""}


@pytest.mark.parametrize(
    "raw,hasil",
    [
        ('a/b\\c:d*e?f"g<h>i|j', "abcdefghij"),
        ("  Andi    Wijaya  ", "Andi Wijaya"),
        ("Aman_123", "Aman_123"),
    ],
)
def test_sanitize_filename(raw, hasil):
    assert sanitize_filename(raw) == hasil


def test_template_tidak_ada(tmp_path):
    with pytest.raises(FileNotFoundError):
        generate_docx(tmp_path / "x.docx", tmp_path, "A", "1", {})


def test_nama_dan_folder_output_default(tmp_path):
    out = generate_docx(template_berita_acara(), tmp_path, "Andi/Wijaya", "21:12", {})
    assert out.parent == tmp_path / "AndiWijaya_2112"
    assert out.name == "Berita Acara dan Nilai Ujian Skripsi_AndiWijaya_2112.docx"
    assert out.exists()


def test_nama_file_custom(tmp_path):
    out = generate_docx(
        template_berita_acara(), tmp_path, "A", "1", {}, output_filename="Kustom/1"
    )
    assert out.name == "Kustom1.docx"


# ---------- Berita Acara: satu template untuk 1 atau 2 pembimbing ----------
KUNCI_BA = ("nama_mahasiswa", "npm", "judul_skripsi", "pembimbing_1", "penguji_1",
            "penguji_2", "hari", "tanggal_bulan_tahun")


def buat_ba(tmp_path, ctx):
    return generate_docx(
        template_berita_acara(), tmp_path, ctx["nama_mahasiswa"], ctx["npm"], ctx
    )


def test_berita_acara_dua_pembimbing(tmp_path, context_ba):
    out = buat_ba(tmp_path, context_ba)
    teks = teks_docx(out)
    assert "{{" not in teks and "{%" not in teks
    for kunci in KUNCI_BA:
        assert context_ba[kunci] in teks, f"{kunci} tidak muncul di Berita Acara"
    # Pembimbing 2 muncul di tabel tanda tangan DAN tabel nilai
    assert jumlah_baris_dengan(out, "Dr. Citra Dewi") == 2


def test_berita_acara_satu_pembimbing(tmp_path, context_ba):
    ctx = tanpa_pembimbing_2(context_ba)
    out = buat_ba(tmp_path, ctx)
    teks = teks_docx(out)
    assert "{{" not in teks and "{%" not in teks
    for kunci in KUNCI_BA:
        assert ctx[kunci] in teks
    assert "Dr. Citra Dewi" not in teks
    # tabel tanda tangan anggota: header + pembimbing 1 saja (bukan 4 baris)
    assert len(Document(str(out)).tables[2].rows) == 3


def test_berita_acara_satu_pembimbing_baris_lebih_sedikit(tmp_path, context_ba):
    dua = Document(str(buat_ba(tmp_path / "dua", context_ba)))
    satu = Document(str(buat_ba(tmp_path / "satu", tanpa_pembimbing_2(context_ba))))
    # tabel tanda tangan anggota (3 vs 2 baris data), nilai, dan penutup
    assert [len(t.rows) for t in satu.tables] != [len(t.rows) for t in dua.tables]
    assert sum(len(t.rows) for t in satu.tables) == sum(
        len(t.rows) for t in dua.tables
    ) - 3


# ---------- Nota Dinas ----------
KUNCI_ND = ("id_nd", "prodi", "lokasi_ujian", "tanggal_hari_ini", "tahun_nd",
            "nama_mahasiswa", "npm", "judul_skripsi")


def buat_nd(tmp_path, ctx):
    return generate_docx(
        template_nota_dinas(), tmp_path, ctx["nama_mahasiswa"], ctx["npm"], ctx,
        output_filename="nd",
    )


def test_nota_dinas_dua_pembimbing(tmp_path, context_nd):
    out = buat_nd(tmp_path, context_nd)
    teks = teks_docx(out)
    assert "{{" not in teks and "{%" not in teks
    for kunci in KUNCI_ND:
        assert context_nd[kunci] in teks, f"{kunci} tidak muncul di Nota Dinas"
    assert re.search(r"^\s*4\.", teks, re.M)
    assert teks.count("Dr. Citra Dewi") == 2  # daftar undangan + tabel tim penguji
    assert "( Pembimbing II )" in teks


def test_nota_dinas_satu_pembimbing(tmp_path, context_nd):
    ctx = tanpa_pembimbing_2(context_nd)
    out = buat_nd(tmp_path, ctx)
    teks = teks_docx(out)
    assert "{{" not in teks and "{%" not in teks
    for kunci in KUNCI_ND:
        assert ctx[kunci] in teks
    assert "( Pembimbing I )" in teks
    assert "( Pembimbing II )" not in teks
    assert not re.search(r"^\s*4\.", teks, re.M)  # nomor "4." milik Pembimbing 2 hilang
