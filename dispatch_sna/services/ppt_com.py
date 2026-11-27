"""PPT ouvert par quelqu'un : écriture par PowerPoint, en co-édition (COM / pywin32).

Le problème est le même que pour l'Excel audit : un PPT ouvert par un collègue (PowerPoint,
PowerPoint Online, Teams) ne peut pas être remplacé « en bloc ». SharePoint refuse (verrou
partagé, 423) et écraser le fichier ferait perdre son travail. La solution est encore celle
d'Office : PowerPoint ouvre la présentation par son adresse SharePoint, rejoint la co-édition,
ne change que ce qu'il faut et enregistre ; SharePoint fusionne avec le travail des autres.

- **Modifier** : seules la zone « Analyse » (la même que celle choisie par ppt_editor, repérée
  par sa diapositive et sa forme) et la ligne « Lien Street View : » (texte + hyperlien) sont
  réécrites ;
- **Remplacer** : les diapositives du PPT de Prod IA sont insérées, puis les anciennes
  supprimées, dans la présentation ouverte (le fichier lui-même n'est pas écrasé).

Les valeurs sont relues après l'enregistrement. PowerPoint n'a qu'une instance par session
Windows : si vous l'avez déjà lancé, c'est le vôtre qui est utilisé, sans fenêtre, et il n'est
jamais fermé ; une présentation que vous avez déjà ouverte est écrite telle quelle et reste
ouverte. Si PowerPoint ne répond plus (fenêtre bloquante), l'outil abandonne le fichier (ID en
attente) au lieu de geler ; il ne tue que l'instance qu'il a lui-même démarrée.
"""
from __future__ import annotations

import os
import queue
import re
import threading
import time
from urllib.parse import unquote

from .excel_com import _liberer_objets_com, _pythoncom, _tuer, fenetre_connexion
from .ppt_editor import _LABEL_LIEN, zone_analyse

_MSO_VRAI = -1
_PP_CLIC = 1                     # ppMouseClick
_PP_ALERTES_AUCUNE = 1           # ppAlertsNone


class PowerPointIndisponible(Exception):
    """pywin32 absent, ou PowerPoint non installé / impossible à démarrer."""


class PowerPointBloque(Exception):
    """PowerPoint n'a pas terminé dans le délai (fenêtre bloquante, connexion demandée…)."""


def powerpoint_disponible() -> bool:
    try:
        import winreg
        import win32com.client  # noqa: F401
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CLSID"))
        return True
    except Exception:  # noqa: BLE001
        return False


def _processus_powerpoint() -> set[int]:
    try:
        import subprocess
        sortie = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
                                capture_output=True, text=True, timeout=15,
                                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)).stdout
        return {int(m) for m in re.findall(r'"POWERPNT\.EXE","(\d+)"', sortie, re.I)}
    except Exception:  # noqa: BLE001
        return set()


def _cle(chemin: str) -> str:
    return unquote(str(chemin)).replace("\\", "/").rstrip("/").casefold()


def _normaliser(texte: str) -> str:
    return re.sub(r"[\r\n\x0b]+", "\n", texte or "").strip()


def _formes(formes):
    """Toutes les formes d'une diapositive, y compris celles des groupes."""
    for i in range(1, formes.Count + 1):
        f = formes(i)
        yield f
        try:
            if f.Type == 6:          # msoGroup
                yield from (f.GroupItems(j) for j in range(1, f.GroupItems.Count + 1))
        except Exception:  # noqa: BLE001
            pass


def _texte_forme(f):
    try:
        if f.HasTextFrame == _MSO_VRAI and f.TextFrame.HasText == _MSO_VRAI:
            return f.TextFrame.TextRange
    except Exception:  # noqa: BLE001
        pass
    return None


class SessionPowerPoint:
    """``with SessionPowerPoint() as pp: pp.modifier(adresse, modele, analyse, lien)``.

    Toutes les opérations COM tournent sur un fil dédié (propriétaire de son appartement COM).
    """

    def __init__(self, delai_s: float = 300):
        self.delai_s = delai_s
        self._jobs: queue.Queue = queue.Queue()
        self._fil = None
        self._pid = None
        self._demarre_par_nous = False
        self.hors_service = ""

    # cycle de vie -----------------------------------------------------
    def __enter__(self):
        pret, etat = threading.Event(), {}

        def boucle():
            pythoncom = _pythoncom()
            pythoncom.CoInitialize()
            app = None
            try:
                import win32com.client as wc
                avant = _processus_powerpoint()
                app = wc.DispatchEx("PowerPoint.Application")
                try:
                    import win32process
                    self._pid = win32process.GetWindowThreadProcessId(app.HWND)[1]
                except Exception:  # noqa: BLE001
                    self._pid = None
                self._demarre_par_nous = bool(self._pid) and self._pid not in avant
                alertes = app.DisplayAlerts
                app.DisplayAlerts = _PP_ALERTES_AUCUNE
            except BaseException as e:  # noqa: BLE001
                etat["erreur"] = e
                pret.set()
                pythoncom.CoUninitialize()
                return
            pret.set()
            try:
                while True:
                    job = self._jobs.get()
                    if job is None:
                        break
                    fn, args, boite = job
                    try:
                        boite["res"] = fn(app, *args)
                    except BaseException as e:  # noqa: BLE001
                        _liberer_objets_com(e)
                        boite["err"] = e
                    boite["fait"].set()
            finally:
                try:
                    app.DisplayAlerts = alertes
                    if self._demarre_par_nous and app.Presentations.Count == 0:
                        app.Quit()                 # jamais le PowerPoint de l'utilisateur
                except Exception:  # noqa: BLE001
                    pass
                app = None
                pythoncom.CoUninitialize()

        self._fil = threading.Thread(target=boucle, name="powerpoint-com", daemon=True)
        self._fil.start()
        if not pret.wait(90):
            raise PowerPointIndisponible("PowerPoint n'a pas démarré dans le délai")
        if "erreur" in etat:
            raise PowerPointIndisponible(f"impossible de démarrer PowerPoint : {etat['erreur']}")
        return self

    def __exit__(self, *exc):
        if self._fil is not None:
            self._jobs.put(None)
            self._fil.join(60)
            self._fil = None
        return False

    # travail ------------------------------------------------------------
    def _executer(self, fn, *args):
        if self.hors_service:
            raise PowerPointBloque(self.hors_service)
        boite = {"fait": threading.Event()}
        self._jobs.put((fn, args, boite))
        fin = time.monotonic() + self.delai_s
        while not boite["fait"].wait(1.0):
            if fenetre_connexion(self._pid):
                self._abandonner("PowerPoint demande une connexion Microsoft 365 : ouvrez une fois un PPT du "
                                 "site SharePoint dans PowerPoint, connectez-vous avec votre compte "
                                 "professionnel, puis rejouez les ID en attente")
            if time.monotonic() > fin:
                self._abandonner(f"PowerPoint ne répond plus depuis {self.delai_s:.0f} s (fenêtre bloquante ?)")
        if "err" in boite:
            raise boite["err"]
        return boite["res"]

    def _abandonner(self, raison):
        """Le fil reste coincé dans l'appel : la session est hors service pour la suite du lot."""
        self.hors_service = raison
        if self._demarre_par_nous:
            _tuer(self._pid)
        self._fil = None
        raise PowerPointBloque(raison)

    def modifier(self, cible: str, modele: str, analyse: str | None, lien: str | None) -> dict:
        """Réécrit Analyse / Lien Street View dans la présentation ``cible`` (adresse SharePoint
        ou chemin). ``modele`` : copie locale non modifiée, pour repérer la zone « Analyse »
        exactement comme ppt_editor. Retourne {"analyse", "lien", "ouverture", "non_verifie"}."""
        zone = zone_analyse(modele) if analyse is not None else None
        return self._executer(_modifier, cible, zone, analyse, lien)

    def remplacer(self, cible: str, src: str) -> dict:
        """Remplace les diapositives de ``cible`` par celles de ``src`` (fichier local)."""
        return self._executer(_remplacer, cible, os.path.abspath(src))


# ── Sur le fil PowerPoint ─────────────────────────────────────────
def _ouvrir(app, cible):
    for i in range(1, app.Presentations.Count + 1):
        p = app.Presentations(i)
        if _cle(p.FullName) == _cle(cible):
            return p, True, "présentation déjà ouverte dans votre PowerPoint"
    # Open(FileName, ReadOnly, Untitled, WithWindow) : sans fenêtre
    p = app.Presentations.Open(cible, 0, 0, 0)
    mode = "co-édition SharePoint" if str(cible).lower().startswith("http") else "fichier local"
    if p.ReadOnly == _MSO_VRAI:
        try:
            p.Close()
        except Exception:  # noqa: BLE001
            pass
        raise PermissionError("PowerPoint ne l'obtient qu'en lecture seule : ouvert en mode exclusif "
                              "(pas en co-édition) — il doit être fermé, ou ouvert dans PowerPoint Online / "
                              "Microsoft 365 avec l'enregistrement automatique")
    return p, False, mode


def _diapo(p, zone):
    if zone.get("slide_id"):
        try:
            return p.Slides.FindBySlideID(zone["slide_id"])
        except Exception:  # noqa: BLE001
            pass
    n = int(re.search(r"(\d+)\.xml$", zone["fichier"]).group(1))
    return p.Slides(n) if n <= p.Slides.Count else None


def _forme_zone(diapo, zone):
    formes = list(_formes(diapo.Shapes))
    par_id = [f for f in formes if f.Id == zone["shape_id"]]
    if len(par_id) == 1 and par_id[0].Name == zone["nom"]:
        return par_id[0]
    par_nom = [f for f in formes if f.Name == zone["nom"]]
    return par_nom[zone["rang"]] if zone["rang"] < len(par_nom) else None


def _ecrire_lien(para, lien: str):
    corps = para.Text.rstrip("\r\n\x0b")
    r = para.Characters(1, len(corps))
    try:
        r.ActionSettings(_PP_CLIC).Hyperlink.Delete()        # plus de lien vers l'ancienne adresse
    except Exception:  # noqa: BLE001
        pass
    nouveau = f"{_LABEL_LIEN}{lien}".rstrip()
    r.Text = nouveau
    if lien:
        para.Characters(len(_LABEL_LIEN) + 1, len(lien)).ActionSettings(_PP_CLIC).Hyperlink.Address = lien


def _paragraphes_lien(p):
    for i in range(1, p.Slides.Count + 1):
        for f in _formes(p.Slides(i).Shapes):
            tr = _texte_forme(f)
            if tr is None:
                continue
            for k in range(tr.Paragraphs().Count, 0, -1):
                para = tr.Paragraphs(k, 1)
                if para.Text.strip().lower().startswith("lien street view"):
                    yield tr, k


def _modifier(app, cible, zone, analyse, lien):
    p, externe, mode = _ouvrir(app, cible)
    res = {"analyse": False, "lien": False, "ouverture": mode, "non_verifie": []}
    try:
        forme = None
        if analyse is not None and zone:
            diapo = _diapo(p, zone)
            forme = _forme_zone(diapo, zone) if diapo is not None else None
            if forme is not None and forme.HasTextFrame == _MSO_VRAI:
                forme.TextFrame.TextRange.Text = analyse.replace("\r\n", "\n").replace("\n", "\r")
                res["analyse"] = True
        if lien is not None:
            for tr, k in list(_paragraphes_lien(p)):
                _ecrire_lien(tr.Paragraphs(k, 1), lien)
                res["lien"] = True
        if not (res["analyse"] or res["lien"]):
            return res
        p.Save()
        # Relecture : le texte attendu est-il bien en place ?
        if res["analyse"] and _normaliser(forme.TextFrame.TextRange.Text) != _normaliser(analyse):
            res["non_verifie"].append("Analyse")
        if res["lien"]:
            attendu = _normaliser(f"{_LABEL_LIEN}{lien}")
            if any(_normaliser(tr.Paragraphs(k, 1).Text) != attendu for tr, k in _paragraphes_lien(p)):
                res["non_verifie"].append("Lien Street View")
        return res
    finally:
        if not externe:
            try:
                p.Close()
            except Exception:  # noqa: BLE001
                pass


def _remplacer(app, cible, src):
    p, externe, mode = _ouvrir(app, cible)
    s = None
    try:
        s = app.Presentations.Open(src, _MSO_VRAI, 0, 0)
        n_nouv, n_anc = s.Slides.Count, p.Slides.Count
        s.Close(); s = None
        if n_nouv == 0:
            raise ValueError("le PPT de Prod IA ne contient aucune diapositive")
        p.Slides.InsertFromFile(src, n_anc, 1, n_nouv)           # après les anciennes…
        for _ in range(n_anc):                                    # … puis on retire les anciennes
            p.Slides(1).Delete()
        p.Save()
        res = {"ouverture": mode, "diapositives": n_nouv, "non_verifie": []}
        if p.Slides.Count != n_nouv:
            res["non_verifie"].append("diapositives")
        return res
    finally:
        if s is not None:
            try:
                s.Close()
            except Exception:  # noqa: BLE001
                pass
        if not externe:
            try:
                p.Close()
            except Exception:  # noqa: BLE001
                pass


class PowerPointALaDemande:
    """PowerPoint démarré au premier PPT ouvert par quelqu'un (un seul essai par lot)."""

    def __init__(self, fabrique=SessionPowerPoint):
        self._fabrique = fabrique if powerpoint_disponible() or fabrique is not SessionPowerPoint else None
        self.session = None
        self._essaye = False
        self.raison = "" if self._fabrique else "PowerPoint / pywin32 introuvables sur ce poste"

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        if self.session is not None:
            try:
                self.session.__exit__(*exc)
            finally:
                self.session = None
        return False

    def pret(self) -> bool:
        if self.session is not None:
            if self.session.hors_service:
                self.raison = self.session.hors_service
                return False
            return True
        if self._fabrique is None or self._essaye:
            return False
        self._essaye = True
        try:
            s = self._fabrique()
            s.__enter__()
            self.session = s
            return True
        except Exception as e:  # noqa: BLE001
            self.raison = f"PowerPoint n'a pas pu démarrer : {str(e).splitlines()[0][:200]}"
            return False

    def modifier(self, cible, modele, analyse, lien):
        if not self.pret():
            raise PermissionError(f"co-édition PowerPoint indisponible : {self.raison}")
        return self.session.modifier(cible, modele, analyse, lien)

    def remplacer(self, cible, src):
        if not self.pret():
            raise PermissionError(f"co-édition PowerPoint indisponible : {self.raison}")
        return self.session.remplacer(cible, src)
