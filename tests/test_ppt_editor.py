import zipfile

from dispatch_sna.services.ppt_editor import modifier_ppt
from tests.fichiers import pptx, ppt_babayaga, ppt_bp, run, sp, textes


def test_modifier_ppt_template_bp(tmp_path):
    p = ppt_bp(tmp_path / "a.pptx")
    res = modifier_ppt(str(p), "Ligne 1 & <ok>\nLigne 2", "https://sv/new")
    assert res == {"analyse": True, "lien": True}
    t = textes(p)
    assert "Ligne 1 &amp; &lt;ok&gt;" in t and "Ligne 2" in t
    assert "Ancienne analyse" not in t
    assert "Lien Street View : https://sv/new" in t
    assert "Analyses:" in t and "SNA nombre de logement: 2" in t
    zipfile.ZipFile(p).testzip()


def test_modifier_ppt_babayaga_hyperlien(tmp_path):
    p = ppt_babayaga(tmp_path / "b.pptx")
    res = modifier_ppt(str(p), "Nouvelle analyse", "https://sv/new?a=1&b=2")
    assert res == {"analyse": True, "lien": True}
    t = textes(p)
    assert "Nouvelle analyse" in t and "Justification IA" not in t
    assert "Lien Street View : https://sv/new?a=1&amp;b=2" in t
    rels = zipfile.ZipFile(p).read("ppt/slides/_rels/slide1.xml.rels").decode()
    assert 'Target="https://sv/new?a=1&amp;b=2"' in rels
    assert "(2 - Maison): adresse au 3 rue Y (Etat IPE: OK)" in t


def test_modifier_ppt_analyse_vide_reprend_la_zone_avant_le_lien(tmp_path):
    p = pptx(tmp_path / "c.pptx", [
        sp("ZoneTexte 5", [run("Analyses:")]),
        sp("ZoneTexte 7", ['<a:endParaRPr lang="fr-FR" sz="1000"/>']),
        sp("ZoneTexte 7", [run("Lien Street View : x")]),
    ])
    assert modifier_ppt(str(p), "Remplie", None) == {"analyse": True, "lien": False}
    assert "Remplie" in textes(p) and "Lien Street View : x" in textes(p)


def test_modifier_ppt_none_ne_touche_pas(tmp_path):
    p = ppt_bp(tmp_path / "a.pptx")
    avant = open(p, "rb").read()
    assert modifier_ppt(str(p), None, None) == {"analyse": False, "lien": False}
    assert open(p, "rb").read() == avant
