import pytest

from core.paths import (
    pilih_template_berdasarkan_pembimbing,
    pilih_template_nota_dinas_berdasarkan_pembimbing,
)


@pytest.mark.parametrize("n", [1, 2])
def test_template_berita_acara_ada(n):
    p = pilih_template_berdasarkan_pembimbing(n)
    assert p.exists()
    assert f"{n}pembimbing" in p.name


@pytest.mark.parametrize("n", [1, 2])
def test_template_nota_dinas_ada(n):
    p = pilih_template_nota_dinas_berdasarkan_pembimbing(n)
    assert p.exists()
    assert f"{n}pembimbing" in p.name


@pytest.mark.parametrize("n", [0, 3, -1])
def test_jumlah_pembimbing_tidak_valid(n):
    with pytest.raises(ValueError):
        pilih_template_berdasarkan_pembimbing(n)
    with pytest.raises(ValueError):
        pilih_template_nota_dinas_berdasarkan_pembimbing(n)
