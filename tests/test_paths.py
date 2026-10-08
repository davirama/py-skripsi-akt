from core.paths import template_berita_acara, template_nota_dinas


def test_template_berita_acara_ada():
    assert template_berita_acara().exists()


def test_template_nota_dinas_ada():
    assert template_nota_dinas().exists()
