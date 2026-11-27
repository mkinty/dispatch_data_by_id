"""Mise à jour par ID imm (unique par ligne), comme dans audit_validation_finale.

Données à l'image des vrais fichiers : un ID erreur peut couvrir plusieurs logements, donc
plusieurs lignes, chacune avec son ID imm ; un ID imm peut être en double (erreur de saisie).
"""
import pytest
from openpyxl import Workbook, load_workbook

from dispatch_sna.services.chemins import calculer_apercu
from dispatch_sna.services.dispatcher import ranger_ids
from dispatch_sna.services.error_ids import extraire_ids_imm, insee_de_id_imm
from dispatch_sna.services.excel_com import SessionExcel, excel_disponible
from dispatch_sna.services.excel_rows import lire_ids_erreur, maj_lignes_excel
from tests.fichiers import log, ppt_babayaga, ppt_bp, textes

ENT = ["ID imm", "ID erreur", "Adresse", "Remarque", "Lien Street View"]
# 74143_5 couvre 3 logements ; ID-…0003 est saisi deux fois (doublon)
LIGNES = [
    ["ID-7414300000A1", "74143_5", "1 rue A", None, None],
    ["ID-7414300000A2", "74143_5", "1 rue A", None, None],
    ["ID-7414300000A3", "74143_5", "1 rue A", None, None],
    ["ID-7414300000B1", "74143_6", "2 rue B", None, None],
    ["ID-7414300000B1", "74143_6", "2 rue B bis", None, None],
]


def classeur(path, remarque, lignes=LIGNES):
    wb = Workbook(); ws = wb.active; ws.title = "Audit"; ws.append(ENT)
    for l in lignes:
        ws.append(l[:3] + [None if remarque is None else f"{remarque} {l[0]}", l[4]])
    wb.save(path)
    return str(path)


def remarques(path):
    ws = load_workbook(path)["Audit"]
    return [ws.cell(r, 4).value for r in range(2, ws.max_row + 1)]


# ── Saisie ────────────────────────────────────────────────────────
def test_extraire_ids_imm():
    texte = "ID imm\nID-3806800000IW\nid-3806800000iv ; ID-2B030000006D\nID-3806800000IW 38068_207"
    assert extraire_ids_imm(texte) == {"38068": ["ID-3806800000IW", "ID-3806800000IV"],
                                       "2B030": ["ID-2B030000006D"]}
    assert insee_de_id_imm("ID-2B030000006D") == "2B030" and insee_de_id_imm("38068_207") == ""


def test_apercu_signale_l_ancien_format(tmp_path):
    a = calculer_apercu(True, "38068_207\n38068_32", str(tmp_path), str(tmp_path))
    assert a["nb"] == 0 and a["anciens"] == 2 and a["lignes"] == []


# ── Excel : seule la ligne de l'ID imm est écrite ─────────────────
def test_seule_la_ligne_de_l_id_imm_est_mise_a_jour(tmp_path):
    src, dst = classeur(tmp_path / "s.xlsx", "NOUV"), classeur(tmp_path / "d.xlsx", None)
    r = maj_lignes_excel(src, dst, ["ID-7414300000A2"], ["Remarque"])
    assert r["maj"] == ["ID-7414300000A2"] and r["id_erreur"] == {"ID-7414300000A2": "74143_5"}
    # les deux autres logements du même ID erreur 74143_5 ne sont PAS touchés
    assert remarques(dst) == [None, "NOUV ID-7414300000A2", None, None, None]


def test_id_imm_en_double_toutes_ses_lignes(tmp_path):
    src, dst = classeur(tmp_path / "s.xlsx", "NOUV"), classeur(tmp_path / "d.xlsx", None)
    maj_lignes_excel(src, dst, ["ID-7414300000B1"], ["Remarque"])
    assert remarques(dst) == [None, None, None, "NOUV ID-7414300000B1", "NOUV ID-7414300000B1"]


def test_id_imm_absent_et_colonne_absente(tmp_path):
    src, dst = classeur(tmp_path / "s.xlsx", "NOUV"), classeur(tmp_path / "d.xlsx", None)
    r = maj_lignes_excel(src, dst, ["ID-7414300000ZZ"], ["Remarque"])
    assert r["absents_src"] == ["ID-7414300000ZZ"] and r["maj"] == []
    wb = Workbook(); wb.active.title = "Audit"; wb.active.append(["ID erreur", "Remarque"]); wb.save(tmp_path / "x.xlsx")
    from dispatch_sna.services.excel_rows import ExcelErreur
    with pytest.raises(ExcelErreur, match="colonne « ID imm » absente"):
        maj_lignes_excel(src, str(tmp_path / "x.xlsx"), ["ID-7414300000A1"], ["Remarque"])


def test_lire_ids_erreur(tmp_path):
    src = classeur(tmp_path / "s.xlsx", "x")
    assert lire_ids_erreur(src, ["ID-7414300000A1", "ID-7414300000B1", "ID-7414300000ZZ"]) == {
        "ID-7414300000A1": "74143_5", "ID-7414300000B1": "74143_6"}


@pytest.mark.skipif(not excel_disponible(), reason="Excel / pywin32 indisponible")
def test_via_excel_seule_la_ligne_de_l_id_imm(tmp_path):
    src, dst = classeur(tmp_path / "s.xlsx", "NOUV"), classeur(tmp_path / "d.xlsx", None)
    with SessionExcel() as xl:
        r = xl.maj_lignes(src, dst, ["ID-7414300000A2", "ID-7414300000B1"], ["Remarque"])
    assert sorted(r["maj"]) == ["ID-7414300000A2", "ID-7414300000B1"] and r["non_verifies"] == []
    assert remarques(dst) == [None, "NOUV ID-7414300000A2", None, "NOUV ID-7414300000B1", "NOUV ID-7414300000B1"]


# ── ranger_ids : PPT nommés par ID erreur ─────────────────────────
def _arbo(tmp_path, excel_prod=True):
    root = tmp_path / "sp"; d = root / "Dep74" / "74143"; (d / "analyse").mkdir(parents=True)
    src = tmp_path / "prod" / "74143"; src.mkdir(parents=True)
    classeur(d / "audit_74143.xlsx", None)
    if excel_prod:
        classeur(src / "audit_74143.xlsx", "NOUV")
    for e in ("74143_5", "74143_6"):
        ppt_bp(d / "analyse" / f"cas_analyse_{e}.pptx"); ppt_babayaga(src / f"cas_analyse_{e}.pptx")
    return root, src, d


def test_plusieurs_id_imm_d_un_meme_id_erreur_un_seul_ppt(tmp_path, monkeypatch):
    from dispatch_sna.services import dispatcher
    root, src, d = _arbo(tmp_path)
    copies = []
    vraie = dispatcher.copier_ppt
    monkeypatch.setattr(dispatcher, "copier_ppt", lambda s, t: copies.append(t) or vraie(s, t))
    lignes = []
    st = ranger_ids(str(root), "74143", str(src), ["ID-7414300000A1", "ID-7414300000A3"], "remplacer", False,
                    lambda m, c=None: lignes.append(m), colonnes=["Remarque"])
    assert st["echecs"] == {} and st["lignes"] == 2 and st["ppt"] == 1
    assert len(copies) == 1 and copies[0].endswith("cas_analyse_74143_5.pptx")
    assert remarques(d / "audit_74143.xlsx")[:3] == ["NOUV ID-7414300000A1", None, "NOUV ID-7414300000A3"]
    assert any("partagent 1 PPT" in l for l in lignes)


def test_ppt_seuls_id_erreur_lu_dans_l_excel_sharepoint(tmp_path):
    """Pas d'Excel dans Prod IA : l'ID erreur (nom du PPT) est lu dans l'Excel SharePoint."""
    root, src, d = _arbo(tmp_path, excel_prod=False)
    st = ranger_ids(str(root), "74143", str(src), ["ID-7414300000B1"], "remplacer", True, log, colonnes=[])
    assert st["echecs"] == {} and st["ppt"] == 1
    assert "Justification IA" in textes(d / "analyse" / "cas_analyse_74143_6.pptx")


def test_ppt_seuls_sans_aucun_excel_bloquant(tmp_path):
    root, src, d = _arbo(tmp_path, excel_prod=False)
    (d / "audit_74143.xlsx").unlink()
    st = ranger_ids(str(root), "74143", str(src), ["ID-7414300000B1"], "remplacer", True, log, colonnes=[])
    assert st["bloquant"] == "aucun Excel dans Prod IA/74143 ni dans SharePoint"


def test_id_imm_introuvable_en_attente(tmp_path):
    root, src, _ = _arbo(tmp_path)
    st = ranger_ids(str(root), "74143", str(src), ["ID-7414300000ZZ"], "remplacer", True, log, colonnes=[])
    assert "ID erreur inconnu, PPT introuvable" in st["echecs"]["ID-7414300000ZZ"] and src.exists()


def test_ppt_seuls_nettoyage_compare_les_id_erreur(tmp_path):
    """PPT seuls : les PPT sont nommés par ID erreur ; ceux des ID saisis ne retiennent pas le dossier."""
    root, src, _ = _arbo(tmp_path)
    lignes = []
    st = ranger_ids(str(root), "74143", str(src), ["ID-7414300000A1", "ID-7414300000B1"], "remplacer", True,
                    lambda m, c=None: lignes.append(m), colonnes=[])
    assert st["echecs"] == {} and not src.exists()


def test_ppt_seuls_autre_id_erreur_garde_le_dossier(tmp_path):
    root, src, _ = _arbo(tmp_path)
    lignes = []
    st = ranger_ids(str(root), "74143", str(src), ["ID-7414300000A1"], "remplacer", True,
                    lambda m, c=None: lignes.append(m), colonnes=[])
    assert st["echecs"] == {} and src.exists()
    assert any("PPT d'autres ID" in l and "74143_6" in l for l in lignes)


def test_modifier_ppt_par_id_imm(tmp_path):
    root, src, d = _arbo(tmp_path)
    st = ranger_ids(str(root), "74143", str(src), ["ID-7414300000B1"], "modifier", False, log, colonnes=["Remarque"])
    assert st["echecs"] == {}
    assert "NOUV ID-7414300000B1" in textes(d / "analyse" / "cas_analyse_74143_6.pptx")
