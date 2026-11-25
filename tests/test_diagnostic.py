"""Messages d'échec compréhensibles : raison précise par ID, PPT introuvable expliqué,
bilan regroupé par cause et rapport CSV (cas réel : 475 communes dont 436 déjà rangées)."""
import csv
import os

from dispatch_sna.services import attente
from dispatch_sna.services.dispatcher import ranger_ids, trouver_ppt
from tests.fichiers import log, ppt_babayaga
from tests.test_surveillance import _arbo


# ── Commune non traitable : la raison accompagne chaque ID ────────
def test_dossier_prod_ia_absent_raison_et_conseil(tmp_path):
    root, _, _ = _arbo(tmp_path, ["74143_1"], [])
    lignes = []
    st = ranger_ids(str(root), "74143", None, ["74143_1", "74143_2"], "remplacer", True,
                    lambda m, c=None: lignes.append(m), colonnes=[])
    assert st["bloquant"] == "aucun dossier 74143 dans Prod IA" and st["err"] == 1
    assert set(st["echecs"]) == {"74143_1", "74143_2"}
    assert "déjà rangée lors d'un précédent dispatch" in st["echecs"]["74143_2"]
    assert any(l.startswith("  ❌ Aucun dossier 74143 dans Prod IA") for l in lignes)


def test_destination_sharepoint_absente(tmp_path):
    root, src, _ = _arbo(tmp_path, ["74143_1"], ["74143_1"])
    st = ranger_ids(str(tmp_path / "ailleurs"), "74143", str(src), ["74143_1"], "remplacer", True, log, colonnes=[])
    assert st["bloquant"] == "dossier SharePoint Dep74/74143 introuvable"
    assert src.exists()


# ── PPT absent : ce que contient le dossier ───────────────────────
def test_ppt_dans_un_sous_dossier_est_utilise(tmp_path):
    root, src, d = _arbo(tmp_path, ["74143_1"], [])
    (src / "PPT").mkdir(); ppt_babayaga(src / "PPT" / "cas_analyse_74143_1.pptx")
    lignes = []
    st = ranger_ids(str(root), "74143", str(src), ["74143_1"], "remplacer", True,
                    lambda m, c=None: lignes.append(m), colonnes=[])
    assert st["echecs"] == {} and st["ppt"] == 1
    assert any("trouvé sous le nom" in l and "PPT" in l for l in lignes)


def test_ppt_casse_differente_est_utilise(tmp_path):
    src = tmp_path / "74143"; src.mkdir()
    ppt_babayaga(src / "CAS_ANALYSE_74143_1.PPTX")
    chemin, _ = trouver_ppt(str(src), "cas_analyse_74143_1.pptx")
    assert chemin and os.path.samefile(chemin, src / "CAS_ANALYSE_74143_1.PPTX")


def test_ppt_casse_differente_dans_un_sous_dossier(tmp_path):
    src = tmp_path / "74143"; (src / "Exports").mkdir(parents=True)
    ppt_babayaga(src / "Exports" / "CAS_ANALYSE_74143_1.PPTX")
    chemin, info = trouver_ppt(str(src), "cas_analyse_74143_1.pptx")
    assert chemin and chemin.endswith("CAS_ANALYSE_74143_1.PPTX") and "Exports" in info


def test_ppt_nom_proche_signale_mais_pas_utilise(tmp_path):
    src = tmp_path / "73150"; src.mkdir()
    for nom in ("cas_analyse_73150_088.pptx", "cas_analyse_73150_88 (1).pptx", "cas_analyse_73150_9.pptx"):
        ppt_babayaga(src / nom)
    chemin, info = trouver_ppt(str(src), "cas_analyse_73150_88.pptx")
    assert chemin is None
    assert "nom proche" in info and "cas_analyse_73150_088.pptx" in info and "(1)" in info
    assert "cas_analyse_73150_9" not in info


def test_ppt_absent_liste_le_contenu_et_sharepoint(tmp_path):
    src = tmp_path / "73150"; src.mkdir()
    for n in (1, 2, 3, 4):
        ppt_babayaga(src / f"cas_analyse_73150_{n}.pptx")
    deja = tmp_path / "cas_analyse_73150_88.pptx"; deja.write_bytes(b"x")
    chemin, info = trouver_ppt(str(src), "cas_analyse_73150_88.pptx", str(deja))
    assert chemin is None
    assert "le dossier contient 4 autre(s) PPT (cas_analyse_73150_1.pptx" in info and "…" in info
    assert "existe déjà dans SharePoint" in info


def test_dossier_sans_ppt(tmp_path):
    src = tmp_path / "73150"; src.mkdir(); (src / "audit_73150.xlsx").write_bytes(b"x")
    assert trouver_ppt(str(src), "cas_analyse_73150_88.pptx") == \
        (None, "le dossier ne contient aucun PPT (1 autre(s) fichier(s))")


# ── Bilan regroupé et rapport CSV ─────────────────────────────────
def _echecs_comme_le_journal():
    """436 communes sans dossier Prod IA (≈10 ID chacune) + 39 PPT absents, comme dans le journal."""
    echecs = {}
    for c in range(436):
        commune = f"{4000 + c:05d}"
        for n in range(1, 11):
            echecs[f"{commune}_{n}"] = (f"aucun dossier {commune} dans Prod IA — commune déjà rangée lors "
                                        "d'un précédent dispatch (son dossier est supprimé après un rangement "
                                        "réussi), ou dossier pas encore déposé")
    for c in range(39):
        commune = f"{73000 + c:05d}"
        echecs[f"{commune}_{c + 7}"] = (f"PPT absent de Prod IA — le dossier contient {c % 4 + 1} autre(s) PPT "
                                        f"(cas_analyse_{commune}_1.pptx) ; un PPT existe déjà dans SharePoint "
                                        "pour cet ID (peut-être déjà livré)")
    return echecs


def test_regrouper_par_cause():
    groupes = attente.regrouper(_echecs_comme_le_journal())
    assert len(groupes) == 2                                       # 4399 lignes → 2 causes
    assert len(groupes[0]["ids"]) == 4360 and len(groupes[0]["communes"]) == 436
    assert groupes[0]["cause"].startswith("aucun dossier «commune» dans Prod IA")
    assert len(groupes[1]["ids"]) == 39
    assert groupes[1]["cause"] == ("PPT absent de Prod IA — le dossier contient N autre(s) PPT ; un PPT existe "
                                   "déjà dans SharePoint pour cet ID (peut-être déjà livré)")


def test_cause_neutralise_id_commune_et_departement():
    assert attente.cause("dossier SharePoint Dep2A/2A004 introuvable") == "dossier SharePoint DepXX/«commune» introuvable"
    assert attente.cause("cas_analyse_73150_88.pptx absent") == "cas_analyse_«ID».pptx absent"


def test_rapport_csv(tmp_path):
    chemin = attente.ecrire_rapport(str(tmp_path / "rapports"), _echecs_comme_le_journal())
    with open(chemin, encoding="utf-8-sig") as f:
        lignes = list(csv.reader(f, delimiter=";"))
    assert lignes[0] == ["ID erreur", "Commune", "Cause", "Raison détaillée"]
    assert len(lignes) == 1 + 4399
    assert lignes[1][1] == lignes[1][0].split("_")[0]
