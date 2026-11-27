"""Co-édition PowerPoint (ppt_com) avec le VRAI PowerPoint, sur de vrais .pptx locaux."""
import os
import re
import zipfile

import pytest

from dispatch_sna.services.ppt_com import SessionPowerPoint, powerpoint_disponible
from dispatch_sna.services.ppt_editor import zone_analyse

pytestmark = pytest.mark.skipif(not powerpoint_disponible(), reason="PowerPoint absent")


def _fabriquer(chemin, diapos):
    import pythoncom
    import win32com.client as wc
    pythoncom.CoInitialize()
    app = wc.DispatchEx("PowerPoint.Application")
    p = app.Presentations.Add(0)
    for formes in diapos:
        s = p.Slides.Add(p.Slides.Count + 1, 12)          # ppLayoutBlank
        for i, (nom, texte, lien) in enumerate(formes):
            f = s.Shapes.AddTextbox(1, 20, 20 + 60 * i, 600, 50)
            f.Name = nom
            f.TextFrame.TextRange.Text = texte
            if lien:
                f.TextFrame.TextRange.Characters(texte.index(lien) + 1, len(lien)).ActionSettings(1).Hyperlink.Address = lien
    p.SaveAs(str(chemin)); p.Close()
    if app.Presentations.Count == 0:
        app.Quit()
    pythoncom.CoUninitialize()
    return str(chemin)


def _lire(chemin):
    with zipfile.ZipFile(chemin) as z:
        diapos = []
        for n in sorted(x for x in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", x)):
            xml = z.read(n).decode()
            diapos.append(["".join(re.findall(r"<a:t>(.*?)</a:t>", p)) for p in re.findall(r"<a:p>.*?</a:p>", xml, re.S)])
        rels = " ".join(z.read(n).decode() for n in z.namelist() if "slides/_rels" in n)
    return diapos, re.findall(r'Target="(https?://[^"]+)"', rels)


CAS = [[("Titre 2", "38068 : adresse au 1 rue X", None), ("ZoneTexte 5", "Analyses:", None),
        ("Rectangle 8", "Ancienne analyse", None),
        ("ZoneTexte 7", "Lien Street View : https://old.example", "https://old.example")]]


@pytest.fixture(scope="module")
def pp():
    with SessionPowerPoint(delai_s=120) as session:
        yield session


def test_modifier_zones_et_lien(tmp_path, pp):
    cible = _fabriquer(tmp_path / "cas_analyse_38068_1.pptx", CAS)
    assert zone_analyse(cible)["nom"] == "Rectangle 8"
    r = pp.modifier(cible, cible, "Ligne 1\nLigne 2", "https://sv.example/nouveau")
    assert r["analyse"] and r["lien"] and r["non_verifie"] == []
    diapos, liens = _lire(cible)
    assert diapos[0] == ["38068 : adresse au 1 rue X", "Analyses:", "Ligne 1", "Ligne 2",
                         "Lien Street View : https://sv.example/nouveau"]
    assert liens == ["https://sv.example/nouveau"]            # l'ancien lien a disparu
    r = pp.modifier(cible, cible, None, "")                    # lien vidé, Analyse non touchée
    diapos, liens = _lire(cible)
    assert diapos[0][-1] == "Lien Street View :" and liens == [] and diapos[0][2] == "Ligne 1"


def test_remplacer_les_diapositives(tmp_path, pp):
    cible = _fabriquer(tmp_path / "cas_analyse_38068_1.pptx", CAS)
    nouveau = _fabriquer(tmp_path / "nouveau.pptx", [[("Titre 2", "NOUVEAU", None)], [("Z", "diapo 2", None)]])
    r = pp.remplacer(cible, nouveau)
    assert r["diapositives"] == 2 and r["non_verifie"] == []
    assert _lire(cible)[0] == [["NOUVEAU"], ["diapo 2"]]


def test_presentation_deja_ouverte_chez_l_utilisateur(tmp_path, pp):
    """Ouverte dans le PowerPoint de l'utilisateur : écrite là, et laissée ouverte."""
    import pythoncom
    import win32com.client as wc
    cible = _fabriquer(tmp_path / "cas_analyse_38068_1.pptx", CAS)
    pythoncom.CoInitialize()
    app = wc.Dispatch("PowerPoint.Application")
    ouverte = app.Presentations.Open(cible, 0, 0, 0)
    try:
        r = pp.modifier(cible, cible, "Écrite dans la présentation ouverte", None)
        assert r["ouverture"] == "présentation déjà ouverte dans votre PowerPoint" and r["analyse"]
        assert ouverte.Slides(1).Shapes("Rectangle 8").TextFrame.TextRange.Text == "Écrite dans la présentation ouverte"
        assert app.Presentations.Count >= 1                     # toujours ouverte
    finally:
        ouverte.Close()
        ouverte = app = None
        pythoncom.CoUninitialize()
    assert "Écrite dans la présentation ouverte" in _lire(cible)[0][0]
