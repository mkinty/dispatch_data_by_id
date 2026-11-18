import os

import pytest
from openpyxl import Workbook

from dispatch_sna.services.excel_rows import ExcelErreur, lire_entetes

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
    cols, f, _ = detecter_colonnes(str(tmp_path / "prod"), str(root))
    assert cols == ENTETES and f.endswith("audit_74143.xlsx") and "prod" in f
    # Prod IA vide → repli sur l'Excel SharePoint
    cols, f, rapport = detecter_colonnes(str(tmp_path / "vide"), str(root))
    assert cols == ENTETES and os.path.join("Dep74", "74143") in f
    assert any("Prod IA : dossier introuvable" in l for l in rapport)
    cols, f, rapport = detecter_colonnes(str(tmp_path / "x"), str(tmp_path / "y"))
    assert (cols, f) == ([], None) and len(rapport) == 2


def _excel(path, onglets, entetes=ENTETES):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook(); wb.remove(wb.active)
    for o in onglets:
        wb.create_sheet(o).append(entetes)
    wb.save(path)
    return path


def test_detection_explique_pourquoi_chaque_fichier_est_ecarte(tmp_path):
    prod = tmp_path / "prod"
    _excel(prod / "74143" / "a_sans_audit.xlsx", ["DGFIP"], ["Adresse", "Remarque"])
    _excel(prod / "74143" / "b_sans_id.xlsx", ["Audit"], ["Adresse", "Remarque"])
    (prod / "74143" / "c_abime.xlsx").write_bytes(b"pas un zip")
    cols, f, rapport = detecter_colonnes(str(prod), str(tmp_path / "rien"))
    assert (cols, f) == ([], None)
    texte = "\n".join(rapport)
    assert "3 fichier(s) Excel trouvé(s)" in texte
    assert ("a_sans_audit.xlsx : onglet « Audit » introuvable (onglets présents : DGFIP) ; "
            "la 1re feuille « DGFIP » n'a pas non plus de colonne « ID erreur »") in texte
    assert "b_sans_id.xlsx : pas de colonne « ID erreur »" in texte
    assert "c_abime.xlsx : n'est pas un .xlsx lisible" in texte
    assert "SharePoint : dossier introuvable" in texte


def test_detection_ne_s_arrete_pas_aux_3_premiers_fichiers(tmp_path):
    root = tmp_path / "sp"
    for n in range(1, 6):                                   # 5 fichiers sans onglet Audit ni « ID erreur »
        _excel(root / "Dep74" / f"7410{n}" / f"audit_7410{n}.xlsx", ["Feuil1"], ["Adresse"])
    _excel(root / "Dep74" / "74143" / "audit_74143.xlsx", ["Audit"])
    cols, f, _ = detecter_colonnes(str(tmp_path / "vide"), str(root))
    assert cols == ENTETES and f.endswith("audit_74143.xlsx")


def test_detection_dans_un_dossier_lot(tmp_path):
    """À déposer/LOT82/38068/audit_38068.xlsx : un niveau de plus qu'avant."""
    _excel(tmp_path / "prod" / "LOT82" / "38068" / "audit_38068.xlsx", ["Audit"])
    cols, f, _ = detecter_colonnes(str(tmp_path / "prod"), str(tmp_path / "rien"))
    assert cols == ENTETES and "LOT82" in f


def test_onglet_audit_au_nom_proche(tmp_path):
    for nom in ("AUDIT", " Audit ", "Audit SNA", "audit_38068"):
        p = _excel(tmp_path / f"{nom.strip()}.xlsx", ["DGFIP", nom, "IPE"])
        assert lire_entetes(str(p)) == ENTETES


def test_onglets_ambigus_refuses(tmp_path):
    p = _excel(tmp_path / "a.xlsx", ["Audit SNA", "Audit DGFIP"])
    with pytest.raises(ExcelErreur, match="plusieurs onglets"):
        lire_entetes(str(p))


def test_supprimer_si_vide(tmp_path):
    vide = tmp_path / "a_deposer"; vide.mkdir()
    supprimer_si_vide(str(vide), log)
    assert not vide.exists()
    plein = tmp_path / "plein"; (plein / "74143").mkdir(parents=True)
    supprimer_si_vide(str(plein), log)
    assert plein.exists()
