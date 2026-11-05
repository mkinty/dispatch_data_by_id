"""Fichiers avec filtres, lignes / colonnes masquées, onglet actif différent d'« Audit »."""
import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.filters import FilterColumn, Filters
from openpyxl.worksheet.table import Table

from dispatch_sna.services.dispatcher import ranger_ids
from dispatch_sna.services.excel_com import SessionExcel, excel_disponible
from dispatch_sna.services.excel_rows import (ExcelErreur, lire_entetes, maj_lignes_excel,
                                              retirer_filtres_fichier)
from tests.fichiers import log

ENT = ["ID erreur", "Etat", "Adresse", "Remarque", "Lien Street View", "Date Analyse"]
IDS = ["74143_2", "74143_15", "74143_20"]      # 74143_2 visible ; 15 et 20 masqués par le filtre
COLS = ["Remarque", "Lien Street View"]


def classeur(path, remarque):
    wb = Workbook(); ws = wb.active; ws.title = "Audit"
    ws.append(ENT)
    for n in range(1, 21):
        ws.append([f"74143_{n}", "OK" if n <= 5 else "KO", f"adr {n}", remarque, f"https://sv/{remarque}/{n}", None])
    wb.create_sheet("DGFIP")["A1"] = "ID erreur"
    wb.active = wb["DGFIP"]                     # onglet actif ≠ Audit
    wb.save(path)
    return str(path)


def filtrer_openpyxl(path):
    wb = load_workbook(path); ws = wb["Audit"]
    ws.auto_filter.ref = "A1:F21"
    ws.auto_filter.filterColumn.append(FilterColumn(colId=1, filters=Filters(filter=["OK"])))
    for r in range(7, 22):
        ws.row_dimensions[r].hidden = True
    ws.column_dimensions["E"].hidden = True
    wb.save(path)


def verifier_nettoye(path, remarque):
    ws = load_workbook(path)["Audit"]
    assert not ws.auto_filter.filterColumn
    assert not any(d.hidden for d in ws.row_dimensions.values())
    assert not any(d.hidden for d in ws.column_dimensions.values())
    for n in (3, 16, 21):
        assert ws.cell(n, 4).value == remarque


# ── Onglet ────────────────────────────────────────────────────────
def test_onglet_audit_introuvable_est_une_erreur(tmp_path):
    wb = Workbook(); wb.active.title = "DGFIP"; wb.active["A1"] = "ID erreur"
    wb.save(tmp_path / "a.xlsx")
    with pytest.raises(ExcelErreur, match="Audit.*DGFIP"):
        lire_entetes(str(tmp_path / "a.xlsx"))


def test_le_bon_onglet_est_lu_meme_si_un_autre_est_actif(tmp_path):
    assert lire_entetes(classeur(tmp_path / "a.xlsx", "x")) == ENT


# ── Réécriture de fichier (openpyxl) ──────────────────────────────
def test_retirer_filtres_fichier(tmp_path):
    p = classeur(tmp_path / "a.xlsx", "x"); filtrer_openpyxl(p)
    r = retirer_filtres_fichier(p)
    assert r == {"filtres": True, "lignes": True, "colonnes": ["Lien Street View"]}
    verifier_nettoye(p, "x")
    assert load_workbook(p)["Audit"].auto_filter.ref == "A1:F21"      # flèches conservées


def test_retirer_filtres_ne_reecrit_pas_un_fichier_sans_filtre(tmp_path):
    p = classeur(tmp_path / "a.xlsx", "x")
    avant = open(p, "rb").read()
    assert retirer_filtres_fichier(p) == {"filtres": False, "lignes": False, "colonnes": []}
    assert open(p, "rb").read() == avant


def test_retirer_filtre_d_un_tableau(tmp_path):
    p = classeur(tmp_path / "a.xlsx", "x")
    wb = load_workbook(p); ws = wb["Audit"]
    t = Table(displayName="T1", ref="A1:F21"); ws.add_table(t)
    wb.save(p)
    wb = load_workbook(p); t = wb["Audit"].tables["T1"]
    t.autoFilter.filterColumn.append(FilterColumn(colId=1, filters=Filters(filter=["OK"])))
    wb.save(p)
    assert retirer_filtres_fichier(p)["filtres"] is True
    assert not load_workbook(p)["Audit"].tables["T1"].autoFilter.filterColumn


def test_maj_lignes_excel_sur_fichiers_filtres(tmp_path):
    src, dst = classeur(tmp_path / "s.xlsx", "NOUV"), classeur(tmp_path / "d.xlsx", "vierge")
    filtrer_openpyxl(src); filtrer_openpyxl(dst)
    r = maj_lignes_excel(src, dst, IDS, COLS)
    assert r["maj"] == IDS and r["absents_dst"] == [] and r["col_absentes"] == []
    ws = load_workbook(dst)["Audit"]
    assert [ws.cell(n, 4).value for n in (3, 16, 21, 2)] == ["NOUV", "NOUV", "NOUV", "vierge"]
    assert not ws.auto_filter.filterColumn and not any(d.hidden for d in ws.row_dimensions.values())


def test_aucune_colonne_cochee_laisse_le_fichier_sharepoint_intact(tmp_path):
    src, dst = classeur(tmp_path / "s.xlsx", "NOUV"), classeur(tmp_path / "d.xlsx", "vierge")
    filtrer_openpyxl(dst)
    avant = open(dst, "rb").read()
    maj_lignes_excel(src, dst, IDS, [])
    assert open(dst, "rb").read() == avant


def test_ranger_ids_retire_les_filtres_des_deux_sources(tmp_path):
    root = tmp_path / "sp"; d = root / "Dep74" / "74143"; d.mkdir(parents=True)
    src_dir = tmp_path / "prod" / "74143"; src_dir.mkdir(parents=True)
    src, dst = classeur(src_dir / "audit_74143.xlsx", "NOUV"), classeur(d / "audit_74143.xlsx", "vierge")
    filtrer_openpyxl(src); filtrer_openpyxl(dst)
    lignes = []
    st = ranger_ids(str(root), "74143", str(src_dir), IDS, None, False, lambda m, c=None: lignes.append(m),
                    colonnes=COLS)
    assert st["lignes"] == 3 and st["echecs"] == {}
    assert not load_workbook(src)["Audit"].auto_filter.filterColumn
    assert sum("🧹" in l for l in lignes) == 2


# ── Via Excel : filtres posés par Excel lui-même ──────────────────
@pytest.mark.skipif(not excel_disponible(), reason="Excel / pywin32 indisponible")
class TestViaExcel:
    @staticmethod
    def _filtrer(s, path):
        wb = s.xl.Workbooks.Open(path); ws = wb.Worksheets("Audit")
        ws.Range("A1:F21").AutoFilter(2, "OK")
        ws.Columns("E:F").Hidden = True
        wb.Worksheets("DGFIP").Activate()
        wb.Save(); wb.Close(False)

    def test_id_filtres_et_colonnes_masquees_sont_mis_a_jour(self, tmp_path):
        src, dst = classeur(tmp_path / "s.xlsx", "NOUV"), classeur(tmp_path / "d.xlsx", "vierge")
        with SessionExcel() as s:
            self._filtrer(s, dst)
            r = s.maj_lignes(src, dst, IDS, COLS)
        assert r["maj"] == IDS and r["absents_dst"] == [] and r["col_absentes"] == []
        assert r["affichage"]["filtres"]
        assert r["affichage"]["colonnes"] == ["Lien Street View", "Date Analyse"]
        ws = load_workbook(dst)["Audit"]
        assert [ws.cell(n, 4).value for n in (3, 16, 21, 2)] == ["NOUV", "NOUV", "NOUV", "vierge"]
        assert ws.cell(21, 5).value == "https://sv/NOUV/20"
        assert not ws.auto_filter.filterColumn
        assert ws.auto_filter.ref                                       # flèches conservées
        assert not any(d.hidden for d in ws.row_dimensions.values())
        assert not any(d.hidden for d in ws.column_dimensions.values())

    def test_filtre_retire_et_enregistre_meme_sans_ligne_a_ecrire(self, tmp_path):
        src, dst = classeur(tmp_path / "s.xlsx", "NOUV"), classeur(tmp_path / "d.xlsx", "vierge")
        wb = load_workbook(src); wb["Audit"]["A3"] = "AUTRE"; wb.save(src)
        wb = load_workbook(dst); wb["Audit"]["A3"] = "AUTRE2"; wb.save(dst)   # 74143_2 absent de dst
        with SessionExcel() as s:
            self._filtrer(s, dst)
            r = s.maj_lignes(src, dst, ["74143_15"], ["Remarque"])
        assert r["maj"] == ["74143_15"]
        assert not load_workbook(dst)["Audit"].auto_filter.filterColumn

    def test_filtre_de_tableau_excel(self, tmp_path):
        src, dst = classeur(tmp_path / "s.xlsx", "NOUV"), classeur(tmp_path / "d.xlsx", "vierge")
        with SessionExcel() as s:
            wb = s.xl.Workbooks.Open(dst); ws = wb.Worksheets("Audit")
            ws.ListObjects.Add(1, ws.Range("A1:F21"), None, 1).Range.AutoFilter(2, "OK")
            wb.Save(); wb.Close(False)
            r = s.maj_lignes(src, dst, IDS, COLS)
        assert r["maj"] == IDS and r["affichage"]["filtres"] and r["affichage"]["erreurs"] == []

    def test_feuille_protegee_filtree_trouve_les_id_et_explique(self, tmp_path):
        src, dst = classeur(tmp_path / "s.xlsx", "NOUV"), classeur(tmp_path / "d.xlsx", "vierge")
        with SessionExcel() as s:
            self._filtrer(s, dst)
            wb = s.xl.Workbooks.Open(dst); wb.Worksheets("Audit").Protect(); wb.Save(); wb.Close(False)
            with pytest.raises(ExcelErreur, match="protégé"):
                s.maj_lignes(src, dst, IDS, COLS)

    def test_ligne_masquee_a_la_main_sans_filtre(self, tmp_path):
        src, dst = classeur(tmp_path / "s.xlsx", "NOUV"), classeur(tmp_path / "d.xlsx", "vierge")
        with SessionExcel() as s:
            wb = s.xl.Workbooks.Open(dst)
            wb.Worksheets("Audit").Rows("16:21").Hidden = True
            wb.Save(); wb.Close(False)
            r = s.maj_lignes(src, dst, IDS, COLS)
        assert r["maj"] == IDS and r["affichage"]["lignes"] and not r["affichage"]["filtres"]
        assert not any(d.hidden for d in load_workbook(dst)["Audit"].row_dimensions.values())

    def test_onglet_audit_absent(self, tmp_path):
        src = classeur(tmp_path / "s.xlsx", "NOUV")
        wb = Workbook(); wb.active.title = "DGFIP"; wb.active.append(ENT); wb.save(tmp_path / "d.xlsx")
        with SessionExcel() as s:
            with pytest.raises(ExcelErreur, match="Audit"):
                s.maj_lignes(src, str(tmp_path / "d.xlsx"), IDS, COLS)
