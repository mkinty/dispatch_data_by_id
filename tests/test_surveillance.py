"""Garde-fou anti-blocage d'Excel (SessionExcelSurveillee) et mode « PPT seuls »."""
import pytest
from openpyxl import Workbook, load_workbook

from dispatch_sna.services.dispatcher import ranger_ids
from dispatch_sna.services.excel_com import (ExcelBloque, SessionExcel, SessionExcelSurveillee,
                                             _tuer, excel_disponible)
from tests.fichiers import log, ppt_babayaga, ppt_bp, textes
from tests.test_filtres import COLS, IDS, classeur

ENT = ["ID erreur", "Remarque", "Lien Street View"]


# ── Mode PPT seuls (aucune colonne cochée) ────────────────────────
def _arbo(tmp_path, ids_excel, ppt_prod, avec_excel_sp=True, avec_excel_prod=True):
    root = tmp_path / "sp"; d = root / "Dep74" / "74143"; (d / "analyse").mkdir(parents=True)
    src = tmp_path / "prod" / "74143"; src.mkdir(parents=True)
    if avec_excel_prod:
        wb = Workbook(); ws = wb.active; ws.title = "Audit"; ws.append(ENT)
        for i in ids_excel:
            ws.append([i, f"Rem {i}", f"https://sv/{i}"])
        wb.save(src / "audit_74143.xlsx")
    if avec_excel_sp:
        wb = Workbook(); ws = wb.active; ws.title = "Audit"; ws.append(ENT); wb.save(d / "audit_74143.xlsx")
    for i in ppt_prod:
        ppt_babayaga(src / f"cas_analyse_{i}.pptx")
    for i in ids_excel:
        ppt_bp(d / "analyse" / f"cas_analyse_{i}.pptx")
    return root, src, d


def test_ppt_seuls_remplacer_supprime_le_dossier_malgre_l_excel(tmp_path):
    """L'Excel Prod IA contient toute la commune : il n'empêche plus la suppression."""
    root, src, d = _arbo(tmp_path, [f"74143_{n}" for n in range(1, 30)], ["74143_1", "74143_2"])
    lignes = []
    st = ranger_ids(str(root), "74143", str(src), ["74143_1", "74143_2"], "remplacer", True,
                    lambda m, c=None: lignes.append(m), colonnes=[])
    assert st["err"] == 0 and st["echecs"] == {} and st["ppt"] == 2
    assert not src.exists()
    assert "Justification IA" in textes(d / "analyse" / "cas_analyse_74143_1.pptx")
    assert any("seuls les PPT" in l for l in lignes)


def test_ppt_seuls_id_absent_de_l_excel_n_est_plus_une_erreur(tmp_path):
    root, src, _ = _arbo(tmp_path, ["74143_9"], ["74143_1"])          # 74143_1 absent de l'Excel
    st = ranger_ids(str(root), "74143", str(src), ["74143_1"], "remplacer", True, log, colonnes=[])
    assert st["echecs"] == {} and not src.exists()


def test_ppt_seuls_sans_aucun_excel(tmp_path):
    root, src, _ = _arbo(tmp_path, [], ["74143_1"], avec_excel_sp=False, avec_excel_prod=False)
    st = ranger_ids(str(root), "74143", str(src), ["74143_1"], "remplacer", True, log, colonnes=[])
    assert st is not None and st["echecs"] == {} and st["ppt"] == 1 and not src.exists()


def test_ppt_seuls_modifier_supprime_le_dossier(tmp_path):
    root, src, d = _arbo(tmp_path, ["74143_1", "74143_2", "74143_3"], [])
    st = ranger_ids(str(root), "74143", str(src), ["74143_1"], "modifier", True, log, colonnes=[])
    assert st["echecs"] == {} and st["ppt"] == 1 and not src.exists()
    assert "Rem 74143_1" in textes(d / "analyse" / "cas_analyse_74143_1.pptx")


def test_ppt_seuls_modifier_sans_excel_prod_est_bloquant(tmp_path):
    root, src, _ = _arbo(tmp_path, ["74143_1"], [], avec_excel_prod=False)
    st = ranger_ids(str(root), "74143", str(src), ["74143_1"], "modifier", True, log, colonnes=[])
    assert st["bloquant"] == "aucun Excel dans Prod IA/74143" and "Modifier les PPT" in st["echecs"]["74143_1"]
    assert src.exists()


def test_ppt_seuls_ppt_d_autres_id_conserve_le_dossier(tmp_path):
    """Des PPT d'ID non saisis = travail pas encore livré : on ne les perd pas."""
    root, src, _ = _arbo(tmp_path, ["74143_1"], ["74143_1", "74143_7"])
    lignes = []
    st = ranger_ids(str(root), "74143", str(src), ["74143_1"], "remplacer", True,
                    lambda m, c=None: lignes.append(m), colonnes=[])
    assert st["echecs"] == {} and src.exists()
    assert any("PPT d'autres ID" in l and "74143_7" in l for l in lignes)


def test_ppt_seuls_erreur_ppt_conserve_le_dossier(tmp_path):
    root, src, _ = _arbo(tmp_path, ["74143_1"], [])                     # PPT absent de Prod IA
    st = ranger_ids(str(root), "74143", str(src), ["74143_1"], "remplacer", True, log, colonnes=[])
    assert st["echecs"] == {"74143_1": "PPT absent de Prod IA — le dossier ne contient aucun PPT (1 autre(s) fichier(s)) ; un PPT existe déjà dans SharePoint pour cet ID (peut-être déjà livré)"}
    assert src.exists()


def test_colonnes_cochees_garde_la_regle_des_lignes_excel(tmp_path):
    root, src, _ = _arbo(tmp_path, ["74143_1", "74143_2"], ["74143_1"])
    wb = load_workbook(src.parent.parent / "sp" / "Dep74" / "74143" / "audit_74143.xlsx")
    wb["Audit"].append(["74143_1", None, None]); wb.save(src.parent.parent / "sp" / "Dep74" / "74143" / "audit_74143.xlsx")
    st = ranger_ids(str(root), "74143", str(src), ["74143_1"], "remplacer", True, log, colonnes=["Remarque"])
    assert st["echecs"] == {} and src.exists()                         # 74143_2 reste dans l'Excel Prod IA


# ── Garde-fou anti-blocage (vrai Excel) ───────────────────────────
pytest_excel = pytest.mark.skipif(not excel_disponible(), reason="Excel / pywin32 indisponible")


def _vivant(pid):
    import win32api, win32con, win32event
    try:
        h = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except Exception:
        return False
    try:
        return win32event.WaitForSingleObject(h, 0) != win32event.WAIT_OBJECT_0
    finally:
        win32api.CloseHandle(h)


@pytest.fixture
def blocage(monkeypatch):
    """Les `n` premiers appels de maj_lignes restent bloqués dans Excel sans signe de vie
    (Application.Wait : l'appel ne revient pas, comme sur une fenêtre modale invisible)."""
    etat = {"restants": 0, "pids": []}
    vraie = SessionExcel.maj_lignes

    def maj_lignes(self, *a):
        etat["pids"].append(self._pid)
        if etat["restants"] > 0:
            etat["restants"] -= 1
            self.xl.Wait(self.xl.Evaluate("NOW()+300/86400"))
        return vraie(self, *a)

    monkeypatch.setattr(SessionExcel, "maj_lignes", maj_lignes)
    return etat


@pytest_excel
def test_excel_bloque_est_tue_et_relance(tmp_path, blocage):
    src, dst = classeur(tmp_path / "s.xlsx", "NOUV"), classeur(tmp_path / "d.xlsx", "vierge")
    blocage["restants"] = 1
    with SessionExcelSurveillee(delai_blocage_s=6) as s:
        r = s.maj_lignes(src, dst, IDS, COLS)
        assert s.relances == 1
    assert r["maj"] == IDS and r["relance_excel"] == 1
    pid_bloque, pid_neuf = blocage["pids"]
    assert pid_bloque != pid_neuf and not _vivant(pid_bloque) and not _vivant(pid_neuf)
    assert load_workbook(dst)["Audit"]["D16"].value == "NOUV"


@pytest_excel
def test_excel_toujours_bloque_echoue_sans_geler(tmp_path, blocage):
    import time
    src, dst = classeur(tmp_path / "s.xlsx", "NOUV"), classeur(tmp_path / "d.xlsx", "vierge")
    blocage["restants"] = 99
    t = time.time()
    with SessionExcelSurveillee(delai_blocage_s=5, essais=2) as s:
        with pytest.raises(ExcelBloque, match="ne répond plus"):
            s.maj_lignes(src, dst, IDS, COLS)
    assert time.time() - t < 90                               # jamais de gel
    assert len(blocage["pids"]) == 2 and not any(_vivant(p) for p in blocage["pids"])


@pytest_excel
def test_excel_disparu_entre_deux_communes_est_relance(tmp_path):
    src, dst = classeur(tmp_path / "s.xlsx", "NOUV"), classeur(tmp_path / "d.xlsx", "vierge")
    with SessionExcelSurveillee(delai_blocage_s=30) as s:
        _tuer(s._etat["session"]._pid)                        # Excel plante ou est fermé de force
        r = s.maj_lignes(src, dst, IDS, COLS)
        assert s.relances == 1
    assert r["maj"] == IDS


@pytest_excel
def test_ranger_ids_excel_bloque_met_la_commune_en_attente(tmp_path, blocage):
    root = tmp_path / "sp"; d = root / "Dep74" / "74143"; d.mkdir(parents=True)
    src_dir = tmp_path / "prod" / "74143"; src_dir.mkdir(parents=True)
    classeur(src_dir / "audit_74143.xlsx", "NOUV"); classeur(d / "audit_74143.xlsx", "vierge")
    blocage["restants"] = 99
    with SessionExcelSurveillee(delai_blocage_s=5, essais=1) as s:
        st = ranger_ids(str(root), "74143", str(src_dir), IDS, None, True, log, colonnes=COLS,
                        maj_excel=s.maj_lignes)
    assert set(st["echecs"]) == set(IDS) and "ne répond plus" in st["echecs"][IDS[0]]
    assert src_dir.exists()


# ── Connexion Microsoft 365 demandée par Excel (comme audit_validation_finale) ──
@pytest.mark.skipif(not excel_disponible(), reason="Excel absent")
def test_fenetre_de_connexion_detectee_puis_connexion_et_reprise(tmp_path, monkeypatch):
    """Excel invisible bloqué sur la connexion Microsoft 365 : détecté tout de suite (pas au bout du
    délai anti-blocage), Excel visible proposé pour se connecter, puis le fichier est écrit."""
    import time
    from dispatch_sna.services import excel_com
    src, dst = tmp_path / "src.xlsx", tmp_path / "dst.xlsx"
    for chemin, rem in ((src, "NOUV"), (dst, "vieux")):
        wb = Workbook(); ws = wb.active; ws.title = "Audit"; ws.append(ENT); ws.append(["74143_1", rem, None])
        wb.save(chemin)
    vraie_maj, etat = SessionExcel.maj_lignes, {"n": 0, "connexions": []}

    def maj_lente(self, *a, **k):
        etat["n"] += 1
        if etat["n"] == 1:
            time.sleep(3)                    # le temps que la fenêtre de connexion soit « vue »
        return vraie_maj(self, *a, **k)
    monkeypatch.setattr(SessionExcel, "maj_lignes", maj_lente)
    monkeypatch.setattr(excel_com, "fenetre_connexion", lambda pid: etat["n"] == 1)
    monkeypatch.setattr(excel_com, "url_sharepoint", lambda chemin: str(chemin))
    monkeypatch.setattr(excel_com, "connexion_microsoft365", lambda adresse, delai: etat["connexions"].append(adresse))
    debut = time.time()
    with SessionExcelSurveillee(delai_blocage_s=60) as xl:
        r = xl.maj_lignes(str(src), str(dst), ["74143_1"], ["Remarque"])
    assert etat["connexions"] == [str(dst)] and r["maj"] == ["74143_1"]
    assert time.time() - debut < 50                       # bien avant le délai anti-blocage
    assert load_workbook(dst).active["B2"].value == "NOUV"
