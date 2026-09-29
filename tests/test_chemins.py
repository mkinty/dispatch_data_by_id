import os

from dispatch_sna.services.chemins import (detecter_colonnes, extraire_commune_nom, get_dep,
                                           lister_communes_prod, supprimer_si_vide)
from tests.fichiers import ENTETES, arbo, log


def test_extraire_commune_et_dep():
    assert extraire_commune_nom("audit_74143.xlsx") == "74143"
    assert extraire_commune_nom("2a004") == "2A004"
    assert extraire_commune_nom("Analyse") is None
    assert get_dep("04121") == "Dep04"


def test_lister_communes_prod(tmp_path):
    arbo(tmp_path)
    (tmp_path / "prod" / "divers").mkdir()
    assert lister_communes_prod(str(tmp_path / "prod")) == {"74143": str(tmp_path / "prod" / "74143")}
    assert lister_communes_prod(str(tmp_path / "absent")) == {}


def test_detecter_colonnes(tmp_path):
    root, _, _ = arbo(tmp_path)
    cols, f = detecter_colonnes(str(tmp_path / "prod"), str(root))
    assert cols == ENTETES and f.endswith("audit_74143.xlsx") and "prod" in f
    # Prod IA vide → repli sur l'Excel SharePoint
    cols, f = detecter_colonnes(str(tmp_path / "vide"), str(root))
    assert cols == ENTETES and os.path.join("Dep74", "74143") in f
    assert detecter_colonnes(str(tmp_path / "x"), str(tmp_path / "y")) == ([], None)


def test_supprimer_si_vide(tmp_path):
    vide = tmp_path / "a_deposer"; vide.mkdir()
    supprimer_si_vide(str(vide), log)
    assert not vide.exists()
    plein = tmp_path / "plein"; (plein / "74143").mkdir(parents=True)
    supprimer_si_vide(str(plein), log)
    assert plein.exists()
