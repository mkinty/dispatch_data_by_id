"""Écritures robustes face aux refus OneDrive ([WinError 389] L'opération de cloud a échoué).

OneDrive n'étant pas disponible en test, les refus sont simulés en remplaçant os.replace,
os.remove ou open par des versions qui lèvent la vraie erreur Windows (winerror 389 / 362).
"""
import builtins
import os
import stat

import pytest

from dispatch_sna.services import dispatcher, fichiers_cloud as fc
from dispatch_sna.services.dispatcher import ranger_ids
from dispatch_sna.services.ppt_editor import copier_ppt, modifier_ppt
from tests.fichiers import log, ppt_babayaga, ppt_bp, textes
from tests.test_surveillance import _arbo


def refus(code=389):
    return OSError(None, "L'opération de cloud a échoué", None, code)


@pytest.fixture(autouse=True)
def sans_attente(monkeypatch):
    pauses = []
    monkeypatch.setattr(fc, "_dormir", pauses.append)
    return pauses


@pytest.fixture
def fichiers(tmp_path):
    src, dst = tmp_path / "src.pptx", tmp_path / "dst.pptx"
    src.write_bytes(b"NOUVEAU" * 1000); dst.write_bytes(b"ANCIEN" * 500)
    return src, dst


def _replace_qui_refuse(monkeypatch, combien, condition=lambda tmp, dst: True):
    vrai, compteur = os.replace, {"n": 0}

    def faux(tmp, dst):
        if condition(tmp, dst) and compteur["n"] < combien:
            compteur["n"] += 1
            raise refus()
        return vrai(tmp, dst)
    monkeypatch.setattr(fc.os, "replace", faux)
    return compteur


# ── classification des erreurs ────────────────────────────────────
def test_erreurs_cloud_reconnues():
    assert fc.est_erreur_cloud(refus(389)) and fc.est_erreur_cloud(refus(362))
    assert not fc.est_erreur_cloud(PermissionError(13, "refusé"))
    assert not fc.est_erreur_cloud(FileNotFoundError(2, "absent"))
    assert not fc.est_erreur_cloud(OSError(None, "autre", None, 5))
    assert not fc.est_erreur_cloud(ValueError("x"))


# ── refus passagers ───────────────────────────────────────────────
def test_refus_passager_reussi_au_reessai(monkeypatch, fichiers, sans_attente):
    src, dst = fichiers
    _replace_qui_refuse(monkeypatch, 1)
    fc.copier(str(src), str(dst))
    assert dst.read_bytes() == src.read_bytes()
    assert sans_attente == [1] and not os.path.exists(str(dst) + ".tmp")


def test_destination_en_ligne_uniquement_rapatriee_puis_remplacee(monkeypatch, fichiers):
    """Cas des PPT déjà livrés (désépinglés) : le remplacement échoue tant que l'ancien
    fichier n'a pas été rapatrié ; on le lit, puis le remplacement passe."""
    src, dst = fichiers
    etat = {"rapatrie": False}
    vrai_open = builtins.open

    def suivi_open(f, mode="r", *a, **k):
        if str(f) == str(dst) and "r" in mode:
            etat["rapatrie"] = True
        return vrai_open(f, mode, *a, **k)
    monkeypatch.setattr(builtins, "open", suivi_open)
    _replace_qui_refuse(monkeypatch, 99, lambda tmp, d: not etat["rapatrie"])
    fc.copier(str(src), str(dst))
    assert etat["rapatrie"] and dst.read_bytes() == src.read_bytes()


def test_dernier_recours_supprimer_puis_deposer(monkeypatch, fichiers):
    """os.replace refusé tant que la destination existe (remplacement de coquille impossible)."""
    src, dst = fichiers
    _replace_qui_refuse(monkeypatch, 99, lambda tmp, d: os.path.exists(d))
    fc.copier(str(src), str(dst))
    assert dst.read_bytes() == src.read_bytes() and not os.path.exists(str(dst) + ".tmp")


# ── refus persistants : jamais de perte ni de faux succès ─────────
def test_refus_persistant_erreur_claire_et_ancien_fichier_intact(monkeypatch, fichiers):
    src, dst = fichiers
    _replace_qui_refuse(monkeypatch, 999)
    with pytest.raises(fc.ErreurOneDrive, match="WinError 389.*OneDrive est lancé"):
        fc.copier(str(src), str(dst))
    assert dst.read_bytes() == b"ANCIEN" * 500                 # remis en place après la suppression
    assert not os.path.exists(str(dst) + ".tmp")


def test_ancien_illisible_jamais_supprime(monkeypatch, fichiers):
    """Si l'ancien fichier ne peut pas être rapatrié, on ne le supprime pas (pas de copie de secours)."""
    src, dst = fichiers
    _replace_qui_refuse(monkeypatch, 999)
    vrai_open, supprimes = builtins.open, []

    def open_dst_refuse(f, mode="r", *a, **k):
        if str(f) == str(dst) and "r" in mode:
            raise refus()
        return vrai_open(f, mode, *a, **k)
    monkeypatch.setattr(builtins, "open", open_dst_refuse)
    monkeypatch.setattr(fc.os, "remove", lambda p: supprimes.append(p))
    with pytest.raises(fc.ErreurOneDrive, match="lecture de dst.pptx"):
        fc.copier(str(src), str(dst))
    assert str(dst) not in supprimes


def test_source_illisible(monkeypatch, fichiers):
    src, dst = fichiers
    vrai_open = builtins.open
    monkeypatch.setattr(builtins, "open", lambda f, mode="r", *a, **k:
                        (_ for _ in ()).throw(refus(362)) if str(f) == str(src) else vrai_open(f, mode, *a, **k))
    with pytest.raises(fc.ErreurOneDrive, match="lecture de src.pptx.*WinError 362"):
        fc.copier(str(src), str(dst))
    assert dst.read_bytes() == b"ANCIEN" * 500


def test_verrou_n_est_pas_reessaye(monkeypatch, fichiers, sans_attente):
    src, dst = fichiers
    monkeypatch.setattr(fc.os, "replace", lambda a, b: (_ for _ in ()).throw(PermissionError(13, "verrouillé")))
    with pytest.raises(PermissionError):
        fc.copier(str(src), str(dst))
    assert sans_attente == []


def test_attributs_de_la_source_non_recopies(fichiers):
    """shutil.copy2 recopiait la lecture seule (et les attributs OneDrive) de la source."""
    src, dst = fichiers
    os.chmod(src, stat.S_IREAD)
    try:
        copier_ppt(str(src), str(dst))
        assert os.access(dst, os.W_OK)
    finally:
        os.chmod(src, stat.S_IWRITE | stat.S_IREAD)


# ── PPT modifiés ──────────────────────────────────────────────────
def test_modifier_ppt_resiste_a_un_refus(monkeypatch, tmp_path):
    p = ppt_bp(tmp_path / "a.pptx")
    _replace_qui_refuse(monkeypatch, 1)
    assert modifier_ppt(str(p), "Nouvelle", "https://sv/n") == {"analyse": True, "lien": True}
    assert "Nouvelle" in textes(p)


# ── disjoncteur dans ranger_ids ───────────────────────────────────
def test_disjoncteur_reporte_la_fin_de_la_commune(monkeypatch, tmp_path):
    ids = [f"74143_{n}" for n in range(1, 11)]
    root, src, _ = _arbo(tmp_path, ids, ids)
    appels = []

    def copie_refusee(s, d):
        appels.append(d)
        raise fc.ErreurOneDrive("OneDrive a refusé le dépôt (WinError 389)")
    monkeypatch.setattr(dispatcher, "copier_ppt", copie_refusee)
    lignes = []
    st = ranger_ids(str(root), "74143", str(src), ids, "remplacer", True,
                    lambda m, c=None: lignes.append(m), colonnes=[])
    assert len(appels) == dispatcher.SEUIL_REFUS_ONEDRIVE           # pas 10 tentatives inutiles
    assert set(st["echecs"]) == set(ids) and st["onedrive_ko"] and src.exists()
    assert all("reporté" in st["echecs"][i] for i in ids[3:])
    assert any("⏸️" in l for l in lignes)


def test_disjoncteur_remis_a_zero_par_un_succes(monkeypatch, tmp_path):
    ids = [f"74143_{n}" for n in range(1, 8)]
    root, src, d = _arbo(tmp_path, ids, ids)
    vraie, n = dispatcher.copier_ppt, {"k": 0}

    def une_sur_deux(s, dst):
        n["k"] += 1
        if n["k"] % 2:
            raise fc.ErreurOneDrive("refus")
        return vraie(s, dst)
    monkeypatch.setattr(dispatcher, "copier_ppt", une_sur_deux)
    st = ranger_ids(str(root), "74143", str(src), ids, "remplacer", True, log, colonnes=[])
    assert n["k"] == 7 and "onedrive_ko" not in st                   # jamais 3 refus d'affilée
    assert st["ppt"] == 3 and len(st["echecs"]) == 4
