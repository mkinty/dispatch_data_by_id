"""Choix de la voie d'écriture de l'Excel SharePoint, pour ne jamais écraser un fichier ouvert.

Le problème : réécrire tout le classeur (openpyxl) pendant que quelqu'un l'a ouvert fait perdre
son travail, et OneDrive ne sait pas fusionner une version produite hors d'Office. La solution
est celle d'Office lui-même : la co-édition, via Excel (services.excel_com), qui n'écrit que les
cellules des lignes visées. Même logique que audit_validation_finale (avf/services/ecriture.py) :

| Fichier                                              | Voie                                       |
|------------------------------------------------------|--------------------------------------------|
| « Écrire via Excel » coché, ou fichier synchronisé    | Excel, co-édition, pour chaque fichier     |
|   avec SharePoint (dossier OneDrive)                  |                                            |
| Fichier local ouvert (fichier verrou ~$ d'Excel)      | Excel démarré à la demande, co-édition     |
| Fichier local libre                                   | réécriture openpyxl                        |
| Fichier ouvert et Excel indisponible                  | rien n'est écrit : ID en attente, expliqué |
"""
from __future__ import annotations

import os

from .excel_rows import lire_cellules, maj_lignes_excel
from .onedrive import url_sharepoint


def proprietaire_verrou(chemin: str, ouvrir: bool = True) -> str | None:
    """Nom de la personne qui a `chemin` ouvert (Excel, PowerPoint…), ou None s'il est libre.

    Office crée « ~$<nom> » à côté du fichier ouvert ; il commence par le nom de l'utilisateur.
    ouvrir=True : un fichier qu'on ne peut pas ouvrir en écriture est aussi considéré comme ouvert
    (False pour un fichier peut-être « en ligne uniquement » : l'ouvrir le ferait télécharger)."""
    if not chemin or not os.path.isfile(chemin):
        return None
    dossier, nom = os.path.split(chemin)
    for candidat in (f"~${nom}", f"~${nom[2:]}"):
        verrou = os.path.join(dossier, candidat)
        if os.path.isfile(verrou):
            return _nom_dans_verrou(verrou) or "un autre utilisateur"
    if not ouvrir:
        return None
    try:
        with open(chemin, "r+b"):
            return None
    except PermissionError:
        return "un autre utilisateur (fichier ouvert ou verrouillé)"
    except OSError:
        return None


def _nom_dans_verrou(chemin: str) -> str:
    """Le fichier verrou commence par la longueur du nom, puis le nom (ANSI)."""
    try:
        with open(chemin, "rb") as f:
            donnees = f.read(165)
    except OSError:
        return ""
    if not donnees:
        return ""
    nom = donnees[1:1 + donnees[0]].decode("cp1252", errors="ignore").strip("\x00 ").strip()
    return nom if nom.isprintable() else ""


class SessionEcriture:
    """Même interface que SessionExcel(Surveillee).maj_lignes : utilisable par ranger_ids.

    fabrique_excel : crée une session Excel (co-édition) ; None = Excel absent du poste.
    toujours_excel : Excel pour chaque fichier (case « Écrire via Excel »).
    """

    def __init__(self, fabrique_excel=None, toujours_excel=False):
        self._fabrique = fabrique_excel
        self.toujours_excel = toujours_excel and fabrique_excel is not None
        self.excel = None
        self._essaye = False
        self.raison_sans_excel = "" if fabrique_excel else "Excel / pywin32 introuvables sur ce poste"

    def __enter__(self):
        if self.toujours_excel:
            self._excel_pret()
        return self

    def __exit__(self, *exc):
        if self.excel is not None:
            self.excel.__exit__(*exc)
            self.excel = None
        return False

    def _excel_pret(self) -> bool:
        """Démarre Excel au premier besoin ; un seul essai par lot."""
        if self.excel is not None:
            return True
        if self._fabrique is None or self._essaye:
            return False
        self._essaye = True
        try:
            session = self._fabrique()
            session.__enter__()
            self.excel = session
            return True
        except Exception as e:
            self.raison_sans_excel = f"Excel n'a pas pu démarrer : {str(e).splitlines()[0][:200]}"
            return False

    def maj_lignes(self, src, dst, ids, colonnes=None) -> dict:
        synchro = url_sharepoint(dst) is not None
        qui = proprietaire_verrou(dst)
        if (self.toujours_excel or synchro or qui) and self._excel_pret():
            r = self.excel.maj_lignes(src, dst, ids, colonnes)
            if qui and r.get("ouverture") and not r.get("bloque"):
                r["ouverture"] = f"{r['ouverture']} — fichier ouvert par {qui}, écrit en co-édition"
            return r
        if qui:
            # Ouvert, et pas d'Excel pour co-éditer : ne surtout pas réécrire le fichier entier
            lecture = lire_cellules(src, ids, colonnes)
            return {"maj": [], "absents_src": lecture["absents_src"], "absents_dst": [],
                    "donnees": lecture["donnees"], "ids_src": lecture["ids_src"], "col_absentes": [],
                    "id_erreur": lecture["id_erreur"], "onglets": {"src": lecture["onglet"], "dst": None},
                    "bloque": f"{qui} (fichier ouvert) — l'écriture dans un fichier ouvert passe par la "
                              f"co-édition Excel, indisponible : {self.raison_sans_excel}"}
        return maj_lignes_excel(src, dst, ids, colonnes)
