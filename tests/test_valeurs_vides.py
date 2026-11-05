"""Colonne cochée dont la valeur Prod IA est vide : l'ancienne valeur est remplacée par le vide."""
import zipfile

import pytest
from openpyxl import Workbook, load_workbook

from dispatch_sna.services.dispatcher import ranger_ids
from dispatch_sna.services.excel_com import SessionExcel, excel_disponible
from dispatch_sna.services.excel_rows import maj_lignes_excel
from dispatch_sna.services.ppt_editor import modifier_ppt
from tests.fichiers import log, ppt_babayaga, ppt_bp, textes

ENT = ["ID erreur", "Remarque", "Lien Street View", "Date Analyse", "Calcul", "Garde"]
COCHEES = ["Remarque", "Lien Street View", "Date Analyse", "Calcul"]


def fichiers(tmp_path):
    src, dst = tmp_path / "s.xlsx", tmp_path / "d.xlsx"
    wb = Workbook(); ws = wb.active; ws.title = "Audit"; ws.append(ENT)
    ws.append(["74143_1", None, None, None, None, None])
    wb.save(src)
    wb = Workbook(); ws = wb.active; ws.title = "Audit"; ws.append(ENT)
    ws.append(["74143_1", "ancienne", "https://old/1", "01/01/2020", "=LEN(A2)", "non cochée"])
    ws["C2"].hyperlink = "https://old/1"
    wb.save(dst)
    return str(src), str(dst)


def verifier(dst):
    ws = load_workbook(dst)["Audit"]
    assert [ws.cell(2, c).value for c in range(2, 6)] == [None, None, None, None]
    assert ws["C2"].hyperlink is None
    assert ws["F2"].value == "non cochée"


def test_reecriture_fichier_efface_les_valeurs(tmp_path):
    src, dst = fichiers(tmp_path)
    r = maj_lignes_excel(src, dst, ["74143_1"], COCHEES)
    assert r["maj"] == ["74143_1"] and r["videes"] == {"74143_1": COCHEES}
    verifier(dst)


@pytest.mark.skipif(not excel_disponible(), reason="Excel / pywin32 indisponible")
class TestViaExcel:
    def test_efface_valeur_lien_et_formule(self, tmp_path):
        src, dst = fichiers(tmp_path)
        with SessionExcel() as s:
            r = s.maj_lignes(src, dst, ["74143_1"], COCHEES)
        assert r["maj"] == ["74143_1"] and r["non_verifies"] == []
        assert r["videes"] == {"74143_1": COCHEES}
        verifier(dst)

    def test_cellule_deja_vide_n_est_pas_signalee(self, tmp_path):
        src, dst = fichiers(tmp_path)
        with SessionExcel() as s:
            s.maj_lignes(src, dst, ["74143_1"], COCHEES)
            r = s.maj_lignes(src, dst, ["74143_1"], COCHEES)
        assert r["maj"] == ["74143_1"] and r["videes"] == {}

    def test_cellule_fusionnee(self, tmp_path):
        src, dst = fichiers(tmp_path)
        with SessionExcel() as s:
            wb = s.xl.Workbooks.Open(dst)
            wb.Worksheets("Audit").Range("B2:C2").Merge()      # Remarque + Lien : valeur portée par B2
            wb.Worksheets("Audit").Range("E2:F2").Merge()      # Calcul + Garde
            wb.Save(); wb.Close(False)
            r = s.maj_lignes(src, dst, ["74143_1"], ["Remarque"])
            assert r["maj"] == ["74143_1"]
            assert load_workbook(dst)["Audit"]["B2"].value is None
            # 2e cellule de la fusion cochée : ne doit pas effacer la valeur de la 1re (Calcul)
            r = s.maj_lignes(src, dst, ["74143_1"], ["Garde"])
        assert r["maj"] == ["74143_1"]
        assert load_workbook(dst)["Audit"]["E2"].value == "=LEN(A2)"


def test_ppt_zones_videes_et_lien_retire(tmp_path):
    p = ppt_babayaga(tmp_path / "b.pptx")
    assert modifier_ppt(str(p), "", "") == {"analyse": True, "lien": True}
    t = textes(p)
    assert "Justification IA" not in t and "Lien Street View :" in t
    assert "hlinkClick" not in zipfile.ZipFile(p).read("ppt/slides/slide1.xml").decode()
    assert modifier_ppt(str(p), "Nouvelle", "https://sv/n") == {"analyse": True, "lien": True}
    t = textes(p)
    assert "Nouvelle" in t and "Lien Street View : https://sv/n" in t


def test_ppt_bp_zone_videe_reste_retrouvable(tmp_path):
    p = ppt_bp(tmp_path / "a.pptx")
    modifier_ppt(str(p), "", None)
    assert "Ancienne analyse" not in textes(p)
    assert modifier_ppt(str(p), "Rejeu", None)["analyse"] is True
    t = textes(p)
    assert "Rejeu" in t and "SNA nombre de logement: 2" in t       # jamais un libellé à la place


def _arbo(tmp_path):
    root = tmp_path / "sp"; d = root / "Dep74" / "74143"; (d / "analyse").mkdir(parents=True)
    src_dir = tmp_path / "prod" / "74143"; src_dir.mkdir(parents=True)
    wb = Workbook(); ws = wb.active; ws.title = "Audit"; ws.append(ENT)
    ws.append(["74143_1", None, None, None, None, None]); wb.save(src_dir / "audit_74143.xlsx")
    wb = Workbook(); ws = wb.active; ws.title = "Audit"; ws.append(ENT)
    ws.append(["74143_1", "ancienne", "https://old/1", None, None, None]); wb.save(d / "audit_74143.xlsx")
    ppt_bp(d / "analyse" / "cas_analyse_74143_1.pptx")
    return root, src_dir, d / "analyse" / "cas_analyse_74143_1.pptx"


def test_ranger_ids_colonne_cochee_vide_aussi_le_ppt(tmp_path):
    root, src_dir, ppt = _arbo(tmp_path)
    st = ranger_ids(str(root), "74143", str(src_dir), ["74143_1"], "modifier", False, log, colonnes=COCHEES)
    assert st["echecs"] == {} and st["ppt"] == 1
    assert "Ancienne analyse" not in textes(ppt)


def test_ranger_ids_colonne_non_cochee_laisse_le_ppt(tmp_path):
    root, src_dir, ppt = _arbo(tmp_path)
    avant = open(ppt, "rb").read()
    st = ranger_ids(str(root), "74143", str(src_dir), ["74143_1"], "modifier", False, log, colonnes=["Calcul"])
    assert st["echecs"] == {} and st["ppt"] == 0
    assert open(ppt, "rb").read() == avant
