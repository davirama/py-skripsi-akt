import pytest

from core.validator import FormData, validate_form, validate_nota_dinas_inputs


def make(**over) -> FormData:
    base = dict(
        nama_mahasiswa="Andi",
        npm="123",
        judul_skripsi="Judul",
        urutan=1,
        hari="Senin",
        jam_mulai="08:00",
        jam_selesai="10:00",
        pembimbing_1="PB1",
        pembimbing_2="",
        penguji_1="PJ1",
        penguji_2="PJ2",
    )
    base.update(over)
    return FormData(**base)


def test_form_valid():
    assert validate_form(make()) == (True, "")


def test_form_valid_dengan_dua_pembimbing():
    assert validate_form(make(pembimbing_2="PB2"))[0] is True


@pytest.mark.parametrize(
    "field,pesan",
    [
        ("nama_mahasiswa", "Nama mahasiswa"),
        ("npm", "NPM"),
        ("judul_skripsi", "Judul"),
        ("hari", "Hari"),
        ("pembimbing_1", "Pembimbing 1"),
        ("penguji_1", "Penguji"),
        ("penguji_2", "Penguji"),
    ],
)
def test_field_wajib_kosong_ditolak(field, pesan):
    ok, msg = validate_form(make(**{field: "   "}))
    assert ok is False
    assert pesan in msg


def test_penguji_tidak_boleh_sama():
    ok, msg = validate_form(make(penguji_1="X", penguji_2="X"))
    assert not ok and "tidak boleh orang yang sama" in msg


def test_pembimbing_2_tidak_boleh_sama_dengan_pembimbing_1():
    ok, msg = validate_form(make(pembimbing_1="A", pembimbing_2="A"))
    assert not ok and "Pembimbing 2" in msg


@pytest.mark.parametrize("mulai,selesai", [("0800", "10:00"), ("08:00", "abc"), ("", "")])
def test_format_jam_salah(mulai, selesai):
    ok, msg = validate_form(make(jam_mulai=mulai, jam_selesai=selesai))
    assert not ok and "HH:mm" in msg


@pytest.mark.parametrize("mulai,selesai", [("10:00", "10:00"), ("10:00", "09:59")])
def test_jam_selesai_harus_lebih_besar(mulai, selesai):
    ok, msg = validate_form(make(jam_mulai=mulai, jam_selesai=selesai))
    assert not ok and "lebih besar" in msg


def test_nota_dinas_valid():
    assert validate_nota_dinas_inputs("1ND", "Ruang A", "Matematika") == (True, "")


@pytest.mark.parametrize(
    "args,pesan",
    [
        (("", "Ruang", "Mat"), "ID ND"),
        (("1", " ", "Mat"), "Lokasi"),
        (("1", "Ruang", ""), "Prodi"),
    ],
)
def test_nota_dinas_field_kosong(args, pesan):
    ok, msg = validate_nota_dinas_inputs(*args)
    assert not ok and pesan in msg
