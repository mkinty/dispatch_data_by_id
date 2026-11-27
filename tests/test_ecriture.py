"""Fichier ouvert : jamais écrasé, écrit en co-édition si possible (comme audit_validation_finale),
et causes de plantage corrigées."""
import gc
import os
import threading
import time

from openpyxl import load_workbook

from dispatch_sna.services import ecriture
from dispatch_sna.services.chemins import calculer_apercu
from dispatch_sna.services.dispatcher import ranger_commune, ranger_ids
from dispatch_sna.services.ecriture import SessionEcriture, proprietaire_verrou
from dispatch_sna.services.excel_com import _liberer_objets_com
from tests.fichiers import arbo, log, textes


def verrou(chemin, nom="Alice Martin"):
    """Fichier verrou d'Office : longueur du nom puis le nom (ANSI), comme Excel / PowerPoint."""
    dossier, f = os.path.split(str(chemin))
    with open(os.path.join(dossier, "~$" + f), "wb") as v:
        v.write(bytes([len(nom)]) + nom.encode("cp1252") + b"\x00" * 40)


# ── Fichier verrou ~$ ─────────────────────────────────────────────
def test_proprietaire_verrou(tmp_path):
    f = tmp_path / "audit_74143.xlsx"; f.write_bytes(b"x")
    assert proprietaire_verrou(str(f)) is None
    verrou(f)
    assert proprietaire_verrou(str(f)) == "Alice Martin"
    assert proprietaire_verrou(str(tmp_path / "absent.xlsx")) is None


# ── Choix de la voie d'écriture ───────────────────────────────────
class FausseSessionExcel:
    def __init__(self, journal): self.journal = journal
    def __enter__(self): self.journal.append("demarre"); return self
    def __exit__(self, *e): self.journal.append("ferme")
    def maj_lignes(self, src, dst, ids, colonnes):
        self.journal.append(("excel", os.path.basename(dst)))
        return {"maj": list(ids), "ouverture": "co-édition SharePoint", "bloque": None}


def test_fichier_local_libre_reecriture_directe(tmp_path):
    root, src, dst = arbo(tmp_path)
    journal = []
    with SessionEcriture(lambda: FausseSessionExcel(journal)) as e:
        r = e.maj_lignes(str(src / "audit_74143.xlsx"), str(dst / "audit_74143.xlsx"), ["74143_1"], ["Remarque"])
    assert r["maj"] == ["74143_1"] and journal == []                     # Excel jamais démarré
    assert load_workbook(dst / "audit_74143.xlsx").active["C2"].value == "Rem 1"


def test_fichier_ouvert_ecrit_via_excel_a_la_demande(tmp_path):
    root, src, dst = arbo(tmp_path)
    verrou(dst / "audit_74143.xlsx")
    journal = []
    with SessionEcriture(lambda: FausseSessionExcel(journal)) as e:
        r = e.maj_lignes(str(src / "audit_74143.xlsx"), str(dst / "audit_74143.xlsx"), ["74143_1"], ["Remarque"])
    assert journal == ["demarre", ("excel", "audit_74143.xlsx"), "ferme"]
    assert "ouvert par Alice Martin, écrit en co-édition" in r["ouverture"]


def test_fichier_ouvert_sans_excel_rien_ecrit(tmp_path):
    root, src, dst = arbo(tmp_path)
    cible = dst / "audit_74143.xlsx"
    verrou(cible)
    avant = cible.read_bytes()
    with SessionEcriture(None) as e:
        r = e.maj_lignes(str(src / "audit_74143.xlsx"), str(cible), ["74143_1"], ["Remarque"])
    assert r["maj"] == [] and "Alice Martin (fichier ouvert)" in r["bloque"]
    assert cible.read_bytes() == avant                                    # jamais écrasé


def test_fichier_synchronise_toujours_via_excel(tmp_path, monkeypatch):
    root, src, dst = arbo(tmp_path)
    monkeypatch.setattr(ecriture, "url_sharepoint", lambda chemin: "https://sfr.sharepoint.com/x.xlsx")
    journal = []
    with SessionEcriture(lambda: FausseSessionExcel(journal)) as e:
        e.maj_lignes(str(src / "audit_74143.xlsx"), str(dst / "audit_74143.xlsx"), ["74143_1"], ["Remarque"])
    assert ("excel", "audit_74143.xlsx") in journal


def test_excel_qui_ne_demarre_pas_repli_explique(tmp_path):
    root, src, dst = arbo(tmp_path)
    verrou(dst / "audit_74143.xlsx")

    def en_panne():
        raise RuntimeError("Excel absent")
    with SessionEcriture(en_panne) as e:
        r = e.maj_lignes(str(src / "audit_74143.xlsx"), str(dst / "audit_74143.xlsx"), ["74143_1"], ["Remarque"])
    assert r["maj"] == [] and "Excel n'a pas pu démarrer : Excel absent" in r["bloque"]


def test_ranger_ids_fichier_ouvert_sans_excel_en_attente(tmp_path):
    root, src, dst = arbo(tmp_path)
    verrou(dst / "audit_74143.xlsx")
    with SessionEcriture(None) as e:
        st = ranger_ids(str(root), "74143", str(src), ["74143_1"], None, True, log, colonnes=["Remarque"],
                        maj_excel=e.maj_lignes)
    assert "verrouillé par Alice Martin (fichier ouvert)" in st["echecs"]["74143_1"] and src.exists()


# ── PPT ouvert ────────────────────────────────────────────────────
def test_ppt_ouvert_non_ecrase(tmp_path):
    root, src, dst = arbo(tmp_path)
    ppt = dst / "analyse" / "cas_analyse_74143_1.pptx"
    verrou(ppt, "Bob Durand")
    avant = ppt.read_bytes()
    st = ranger_ids(str(root), "74143", str(src), ["74143_1"], "remplacer", True, log, colonnes=[])
    assert "ouvert par Bob Durand" in st["echecs"]["74143_1"] and ppt.read_bytes() == avant
    assert src.exists()


def test_mode_commune_fichier_ouvert_non_remplace(tmp_path):
    root, src, dst = arbo(tmp_path)
    verrou(dst / "audit_74143.xlsx")
    avant = (dst / "audit_74143.xlsx").read_bytes()
    st = ranger_commune(str(tmp_path / "prod"), str(root), "74143", str(src), True, log)
    assert st["err"] == 1 and (dst / "audit_74143.xlsx").read_bytes() == avant and src.exists()


# ── Plantages ─────────────────────────────────────────────────────
def test_trace_vide_avant_passage_a_un_autre_fil():
    """Les objets retenus par la trace d'une exception sont libérés sur le fil qui la lève."""
    libere = threading.Event()

    class ObjetCom:
        def __del__(self): libere.set()

    def travail():
        objet = ObjetCom()   # comme un classeur / une cellule COM
        raise ValueError("échec Excel")

    try:
        travail()
    except ValueError as e:
        erreur = e
        _liberer_objets_com(e)
    gc.collect()
    assert libere.is_set() and str(erreur) == "échec Excel"


def test_apercu_de_centaines_de_communes_sans_interface(tmp_path):
    """Calculé hors du fil de l'interface : centaines de communes, sans Tkinter."""
    prod, sp = tmp_path / "prod", tmp_path / "sp"
    texte = []
    for n in range(400):
        c = f"{38000 + n:05d}"
        (prod / c).mkdir(parents=True)
        if n % 2:
            (sp / f"Dep{c[:2]}" / c).mkdir(parents=True)
        texte.append(f"ID-{c}00000A{n % 10}")
    t = time.time()
    a = calculer_apercu(True, "\n".join(texte), str(prod), str(sp))
    assert len(a["lignes"]) == 400 and a["n_ok"] == 200 and a["n_no"] == 200 and a["nb"] == 400
    assert time.time() - t < 10
