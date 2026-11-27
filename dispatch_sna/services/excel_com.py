"""Écriture cellule par cellule via Excel (COM / pywin32), dans la session de co-édition.

Pourquoi : réécrire le fichier entier (openpyxl) puis laisser OneDrive l'envoyer
échoue dès qu'un autre utilisateur a le classeur ouvert — OneDrive ne sait pas
fusionner une version produite hors d'Office (conflit, copie « -NOMDUPC », ou
envoi refusé). Ici, Excel ouvre le classeur par son adresse SharePoint, rejoint
la co-édition, ne modifie que les cellules des ID / colonnes demandés et enregistre :
SharePoint fusionne avec le travail des autres. Les valeurs sont relues après
l'enregistrement pour confirmer qu'elles sont bien en place.

Si le classeur est ouvert en mode exclusif ailleurs, Excel l'ouvre en lecture
seule : rien n'est écrit et le résultat l'indique (« bloque »), au lieu d'un faux succès.
"""
from __future__ import annotations

import datetime as _dt
import math
from urllib.parse import unquote

from openpyxl.utils import get_column_letter

from .constantes import ONGLET_AUDIT, COL_ID_ERR, COL_ID_IMM
from .excel_rows import (CLES, ExcelErreur, OngletIntrouvable, affichage_modifie, choisir_onglet, cle_de,
                         repli_premiere_feuille, lire_cellules, formule_pour, _norm)
from .onedrive import url_sharepoint

_EN = 1033                 # propriétés NumberFormat / Formula toujours en anglais (sinon « General » est refusé en français)
_XL_CALC_MANUEL = -4135
_EPOQUE = _dt.datetime(1899, 12, 30)


class ExcelIndisponible(Exception):
    """pywin32 absent ou Excel non installé / impossible à démarrer."""


class ExcelOccupe(Exception):
    """Excel a refusé les appels trop longtemps (RPC_E_CALL_REJECTED persistant)."""


# HRESULT renvoyés quand Excel est occupé (chargement / synchro co-édition / recalcul / enregistrement)
RPC_E_CALL_REJECTED = -2147418111          # 0x80010001 « L'appel a été rejeté par l'appelé »
RPC_E_SERVERCALL_RETRYLATER = -2147417846  # 0x8001010A
_REJETS = (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER)


def _est_rejet(e) -> bool:
    return bool(getattr(e, "args", None)) and e.args[0] in _REJETS


# L'application Excel elle-même a disparu ou ne répond plus : seule une instance neuve peut continuer.
_DECONNEXIONS = (-2147417848,   # 0x80010108 RPC_E_DISCONNECTED « déconnecté de ses clients »
                 -2147023174,   # 0x800706BA serveur RPC indisponible
                 -2147023170,   # 0x800706BE l'appel de procédure distante a échoué
                 -2147417851)   # 0x80010105 RPC_E_SERVERFAULT
_DISP_E_EXCEPTION = -2147352567
_OBJET_REQUIS = -2146827864     # 0x800A01A8 : objet invalidé (classeur rechargé par la fusion co-édition)


def _est_deconnexion(e) -> bool:
    return bool(getattr(e, "args", None)) and e.args[0] in _DECONNEXIONS


def _est_objet_invalide(e) -> bool:
    """Objet COM rendu caduc par un rechargement du classeur (fusion SharePoint) : il suffit de rouvrir."""
    args = getattr(e, "args", None) or ()
    info = args[2] if len(args) > 2 else None
    return bool(args) and args[0] == _DISP_E_EXCEPTION and bool(info) and len(info) > 5 and info[5] == _OBJET_REQUIS


class _FiltreMessages:
    """Filtre de messages COM (IMessageFilter) pour le thread courant.

    Sans lui, tout appel envoyé pendant qu'Excel est occupé échoue aussitôt avec
    RPC_E_CALL_REJECTED. Avec lui, Windows réessaie l'appel toutes les 250 ms
    jusqu'à `delai_max_ms`. pywin32 n'exposant pas CoRegisterMessageFilter,
    l'interface est construite avec ctypes (vtable de 6 méthodes).
    """
    _ATTENTE_MS = 250

    def __init__(self, delai_max_ms=120_000, signe_de_vie=None):
        import ctypes
        import uuid
        from ctypes import wintypes
        self.delai_max_ms = delai_max_ms
        self.signe_de_vie = signe_de_vie
        self._ancien = ctypes.c_void_p()
        self._installe = False
        HRESULT, DWORD, ULONG, LPVOID = ctypes.c_long, wintypes.DWORD, ctypes.c_ulong, ctypes.c_void_p

        class GUID(ctypes.Structure):
            _fields_ = [("d", ctypes.c_ubyte * 16)]

        guid = lambda s: GUID.from_buffer_copy(uuid.UUID(s).bytes_le)
        connus = {bytes(guid("00000000-0000-0000-C000-000000000046").d),    # IUnknown
                  bytes(guid("00000016-0000-0000-C000-000000000046").d)}    # IMessageFilter
        F = ctypes.WINFUNCTYPE

        def query_interface(this, riid, ppv):
            if bytes(riid.contents.d) in connus:
                ppv[0] = this; return 0
            ppv[0] = None; return -2147467262                 # E_NOINTERFACE

        def handle_incoming(this, call_type, task, tick, info):
            return 0                                          # SERVERCALL_ISHANDLED

        def retry_rejected(this, task, tick, reject_type):
            # tick = ms écoulées depuis le 1er refus de CET appel ; 0xFFFFFFFF = abandon
            if self.signe_de_vie:
                try: self.signe_de_vie()
                except Exception: pass
            return self._ATTENTE_MS if tick < self.delai_max_ms else 0xFFFFFFFF

        def message_pending(this, task, tick, pending_type):
            return 2                                          # PENDINGMSG_WAITDEFPROCESS

        self._fns = [
            F(HRESULT, LPVOID, ctypes.POINTER(GUID), ctypes.POINTER(LPVOID))(query_interface),
            F(ULONG, LPVOID)(lambda this: 1),                 # AddRef  (objet statique)
            F(ULONG, LPVOID)(lambda this: 1),                 # Release
            F(DWORD, LPVOID, DWORD, LPVOID, DWORD, LPVOID)(handle_incoming),
            F(DWORD, LPVOID, LPVOID, DWORD, DWORD)(retry_rejected),
            F(DWORD, LPVOID, LPVOID, DWORD, DWORD)(message_pending),
        ]

        class VTable(ctypes.Structure):
            _fields_ = [(f"f{i}", type(f)) for i, f in enumerate(self._fns)]

        class Objet(ctypes.Structure):
            _fields_ = [("vtbl", ctypes.POINTER(VTable))]

        self._vtbl = VTable(*self._fns)
        self._objet = Objet(ctypes.pointer(self._vtbl))
        self._ole32 = ctypes.windll.ole32
        self._ole32.CoRegisterMessageFilter.argtypes = [LPVOID, ctypes.POINTER(LPVOID)]

    def installer(self):
        import ctypes
        hr = self._ole32.CoRegisterMessageFilter(ctypes.addressof(self._objet), ctypes.byref(self._ancien))
        self._installe = hr == 0
        return self._installe

    def retirer(self):
        if self._installe:
            self._ole32.CoRegisterMessageFilter(self._ancien, None)
            self._installe = False


def excel_disponible() -> bool:
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CurVer"))
        return True
    except Exception:
        return False


# ── conversions ───────────────────────────────────────────────────
def _serie(v):
    """datetime / date / time → numéro de série Excel (évite le décalage UTC de pywin32)."""
    if isinstance(v, _dt.datetime):
        return (v.replace(tzinfo=None) - _EPOQUE).total_seconds() / 86400
    if isinstance(v, _dt.date):
        return float((v - _EPOQUE.date()).days)
    if isinstance(v, _dt.time):
        return (v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6) / 86400
    return None


def _put_en(obj, prop, valeur):
    dispid = obj._oleobj_.GetIDsOfNames(prop)
    obj._oleobj_.Invoke(dispid, _EN, _pythoncom().DISPATCH_PROPERTYPUT, 0, valeur)


def _get_en(obj, prop):
    dispid = obj._oleobj_.GetIDsOfNames(prop)
    return obj._oleobj_.Invoke(dispid, _EN, _pythoncom().DISPATCH_PROPERTYGET, 1)


def _pythoncom():
    import pythoncom
    return pythoncom


def _plat(valeurs):
    """Range.Value d'une ligne ou d'une colonne → liste plate (scalaire si 1 seule cellule)."""
    if not isinstance(valeurs, tuple):
        return [valeurs]
    return [x for rangee in valeurs for x in rangee]


def _formule_norm(f) -> str:
    return "".join(str(f or "").split()).upper()


def _lien_norm(u) -> str:
    return unquote(str(u or "")).strip().rstrip("/")


# ── écriture / vérification d'une cellule ─────────────────────────
def _ecrire(ws, cell, attendu, lien, fmt) -> bool:
    """Écrit valeur / formule, format et lien. Retourne False si le lien cliquable n'a pas pu
    être créé parce que l'onglet est protégé (Excel y interdit l'insertion de liens, même dans
    une cellule déverrouillée) : la valeur est alors écrite, le lien manque."""
    if isinstance(attendu, str) and attendu.startswith("="):
        _put_en(cell, "Formula", attendu)
    elif attendu is None or attendu == "":
        # ClearContents est refusé sur une partie de cellule fusionnée. Seule la 1re cellule
        # porte la valeur : vider les suivantes effacerait celle d'une autre colonne.
        if not cell.MergeCells:
            cell.ClearContents()
        elif cell.MergeArea.Cells(1, 1).Address == cell.Address:
            cell.MergeArea.ClearContents()
    elif _serie(attendu) is not None:
        cell.Value2 = _serie(attendu)
    else:
        cell.Value2 = attendu
        if isinstance(attendu, str) and not isinstance(cell.Value2, str):
            cell.Value2 = "'" + attendu        # Excel a converti le texte (ex. « 04121 » → 4121) : on force le texte
    if fmt and fmt != "General":
        try:
            if _get_en(cell, "NumberFormat") != fmt:
                _put_en(cell, "NumberFormat", fmt)
        except Exception:
            pass                               # format exotique : la valeur prime
    if lien and not (isinstance(attendu, str) and attendu.startswith("=")):
        if cell.Hyperlinks.Count:
            cell.Hyperlinks.Delete()
        try:
            if isinstance(attendu, str) and attendu:
                ws.Hyperlinks.Add(cell, lien, "", "", attendu)
            else:
                ws.Hyperlinks.Add(cell, lien)  # garde la valeur de la cellule (affiche le lien si vide)
        except Exception as e:
            if _est_rejet(e) or not ws.ProtectContents:
                raise
            if attendu in (None, ""):
                cell.Value2 = lien             # au moins l'adresse, en texte
            return False
    elif not lien and cell.Hyperlinks.Count:
        cell.Hyperlinks.Delete()
    return True


def _conforme(cell, attendu, lien) -> bool:
    if isinstance(attendu, str) and attendu.startswith("="):
        ok = _formule_norm(_get_en(cell, "Formula")) == _formule_norm(attendu)
    else:
        lu = cell.Value2
        if attendu is None or attendu == "":
            ok = lu in (None, "")
        elif _serie(attendu) is not None:
            ok = isinstance(lu, float) and abs(lu - _serie(attendu)) < 1e-6
        elif isinstance(attendu, bool):
            ok = lu == attendu
        elif isinstance(attendu, (int, float)):
            ok = isinstance(lu, (int, float)) and math.isclose(float(lu), float(attendu), rel_tol=1e-9, abs_tol=1e-9)
        else:
            ok = str(lu).replace("\r\n", "\n") == str(attendu).replace("\r\n", "\n")
    if ok and lien and not (isinstance(attendu, str) and attendu.startswith("=")):
        ok = cell.Hyperlinks.Count > 0 and _lien_norm(cell.Hyperlinks(1).Address) == _lien_norm(lien)
    elif ok and not lien:
        ok = cell.Hyperlinks.Count == 0                # un ancien lien ne doit pas survivre
    return ok


def _liberer_objets_com(exc):
    """Vide, sur le fil courant, les variables locales des cadres de la trace de `exc`.

    Une exception garde dans sa trace les cadres d'exécution, donc leurs variables locales :
    classeur, feuille, cellules… Ces objets COM seraient sinon libérés plus tard par le
    ramasse-miettes, depuis n'importe quel fil — par exemple le fil de traitement à qui
    l'exception est transmise. Libérer un objet COM hors de son appartement peut faire
    planter Python sans message. Le type et le message de l'exception sont conservés.
    (Même correctif que dans audit_validation_finale.)
    """
    import traceback
    vues = set()
    while exc is not None and id(exc) not in vues:
        vues.add(id(exc))
        if exc.__traceback__ is not None:
            traceback.clear_frames(exc.__traceback__)
        exc = exc.__cause__ or exc.__context__


# ── processus Excel ───────────────────────────────────────────────
def _pid_de(xl):
    """PID de l'instance Excel pilotée (via la fenêtre principale)."""
    try:
        import win32process
        return win32process.GetWindowThreadProcessId(xl.Hwnd)[1] or None
    except Exception:
        return None


def _terminer_si_toujours_ouvert(pid, attente_s=15):
    if not pid:
        return
    try:
        import win32api, win32con, win32event
        h = win32api.OpenProcess(win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE, False, pid)
    except Exception:
        return                                         # déjà fermé
    try:
        if win32event.WaitForSingleObject(h, int(attente_s * 1000)) == win32event.WAIT_TIMEOUT:
            win32api.TerminateProcess(h, 1)
            win32event.WaitForSingleObject(h, 15_000)
    finally:
        win32api.CloseHandle(h)


def _processus_vivant(pid) -> bool:
    if not pid:
        return True                                    # inconnu : on ne conclut pas à sa disparition
    try:
        import win32api, win32con, win32event
        h = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    except Exception:
        return False
    try:
        return win32event.WaitForSingleObject(h, 0) == win32event.WAIT_TIMEOUT
    finally:
        win32api.CloseHandle(h)


def _tuer(pid, attente_s=15) -> bool:
    """Termine le processus `pid` et attend qu'il ait VRAIMENT disparu : rouvrir le même fichier
    alors que Windows n'a pas encore relâché ses verrous donne un objet Excel à moitié initialisé."""
    if not pid:
        return False
    try:
        import win32api, win32con, win32event
        h = win32api.OpenProcess(win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE, False, pid)
    except Exception:
        return True                                    # déjà terminé
    try:
        win32api.TerminateProcess(h, 1)
        return win32event.WaitForSingleObject(h, int(attente_s * 1000)) == win32event.WAIT_OBJECT_0
    except Exception:
        return False
    finally:
        win32api.CloseHandle(h)


# ── Connexion Microsoft 365 demandée par Office (comme audit_validation_finale) ──
# Classe de fenêtre de la connexion « moderne » d'Office (WebView) : son apparition dans le
# processus Excel / PowerPoint pendant un appel = connexion au compte professionnel demandée.
_CLASSES_CONNEXION = {"BasicEmbeddedBrowser"}


def fenetre_connexion(pid) -> bool:
    """Vrai si le processus `pid` affiche la fenêtre de connexion Microsoft 365."""
    if not pid:
        return False
    try:
        import win32gui, win32process
    except ImportError:
        return False
    trouvee = []

    def examiner(fenetre, _):
        try:
            if (win32process.GetWindowThreadProcessId(fenetre)[1] == pid
                    and win32gui.GetClassName(fenetre) in _CLASSES_CONNEXION):
                trouvee.append(fenetre)
                return False
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(examiner, None)
    except Exception:          # EnumWindows « échoue » quand le rappel l'arrête
        pass
    return bool(trouvee)


def connexion_microsoft365(adresse: str, delai_s: float = 300) -> None:
    """Ouvre `adresse` (fichier SharePoint) dans un Excel VISIBLE pour que l'utilisateur se
    connecte une fois à son compte professionnel. Office mémorise la connexion : les instances
    invisibles suivantes ouvrent ensuite les fichiers sans rien demander.
    Lève ConnexionRequise si la connexion n'est pas faite dans le délai."""
    import threading
    etat = {"pid": None}
    fini = threading.Event()

    def ouvrir():
        pythoncom = _pythoncom()
        pythoncom.CoInitialize()
        xl = None
        try:
            import win32com.client as wc
            xl = wc.DispatchEx("Excel.Application")
            etat["pid"] = _pid_de(xl)
            xl.Visible = True
            try: xl.WindowState = -4137                # xlMaximized, au premier plan
            except Exception: pass
            wb = xl.Workbooks.Open(adresse, 0, True)   # lecture seule : seulement pour se connecter
            wb.Close(False)
            etat["ok"] = True
        except BaseException as e:
            etat["erreur"] = e
        finally:
            try: xl.Quit()
            except Exception: pass
            xl = None
            pythoncom.CoUninitialize()
            fini.set()

    threading.Thread(target=ouvrir, name="excel-connexion", daemon=True).start()
    fini.wait(delai_s)
    if not etat.get("ok"):
        _tuer(etat["pid"])
        raise ConnexionRequise(f"connexion Microsoft 365 non effectuée en {delai_s / 60:.0f} min"
                               + (f" ({etat['erreur']})" if etat.get("erreur") else "")
                               + " — ouvrez une fois un fichier du site SharePoint dans Excel, connectez-vous "
                                 "avec votre compte professionnel, puis rejouez les ID en attente")
    _terminer_si_toujours_ouvert(etat["pid"], attente_s=10)


# ── session ───────────────────────────────────────────────────────
class SessionExcel:
    """Une instance Excel dédiée (jamais celle de l'utilisateur), fermée à la sortie.

        with SessionExcel() as xl:
            r = xl.maj_lignes(src, dst, ids, colonnes)
    """

    TENTATIVES = 3              # essais complets d'un classeur si Excel reste occupé au-delà du délai

    def __init__(self, visible=False, delai_occupe_s=120):
        import time
        self.visible = visible
        self.delai_occupe_s = delai_occupe_s
        self.xl = None
        self._filtre = None
        self._pid = None
        self.activite = time.time()     # dernier signe de vie d'Excel (lu par SessionExcelSurveillee)

    def _vivant(self):
        import time
        self.activite = time.time()

    def __enter__(self):
        try:
            pythoncom = _pythoncom()
            import win32com.client as wc
        except ImportError as e:
            raise ExcelIndisponible("pywin32 n'est pas installé") from e
        pythoncom.CoInitialize()
        # Avant tout appel à Excel : les refus « occupé » seront réessayés au lieu d'échouer
        # Un refus « occupé » réessayé prouve qu'Excel est vivant : c'est un signe de vie.
        self._filtre = _FiltreMessages(delai_max_ms=int(self.delai_occupe_s * 1000), signe_de_vie=self._vivant)
        self._filtre.installer()
        try:
            self.xl = wc.DispatchEx("Excel.Application")   # nouvelle instance, n'interfère pas avec l'Excel ouvert
            self._pid = _pid_de(self.xl)
            self.xl.Visible = bool(self.visible)
            self.xl.DisplayAlerts = False
            self.xl.AskToUpdateLinks = False
            self.xl.ScreenUpdating = bool(self.visible)
            if not self.visible:
                try: self.xl.Interactive = False   # aucune saisie possible dans cette instance invisible
                except Exception: pass
            self._vivant()
        except Exception as e:
            self.__exit__(None, None, None)
            raise ExcelIndisponible(f"impossible de démarrer Excel : {e}") from e
        return self

    def __exit__(self, *exc):
        if self.xl is not None:
            try:
                for i in range(self.xl.Workbooks.Count, 0, -1):
                    self.xl.Workbooks(i).Close(False)
            except Exception:
                pass
            try: self.xl.Quit()
            except Exception: pass
            self.xl = None
        if self._filtre is not None:
            self._filtre.retirer(); self._filtre = None
        _pythoncom().CoUninitialize()
        # Si une référence COM traîne encore (exception en cours, objet gardé ailleurs), Excel
        # ne se ferme pas et reste invisible en verrouillant les fichiers : on termine CETTE
        # instance-là, identifiée par son PID au démarrage — jamais l'Excel de l'utilisateur.
        _terminer_si_toujours_ouvert(self._pid, attente_s=15)
        self._pid = None
        return False

    def _attendre_pret(self):
        """Attend qu'Excel ait fini (chargement, synchro co-édition, recalcul, envoi)."""
        import time
        limite = time.time() + self.delai_occupe_s
        while time.time() < limite:
            try:
                if self.xl.Ready:
                    self._vivant()
                    return
                self._vivant()                       # Excel répond (« pas prêt ») : il est vivant
            except Exception as e:
                if not _est_rejet(e):
                    raise
            _pythoncom().PumpWaitingMessages()
            time.sleep(0.25)
        raise ExcelOccupe(f"Excel n'est pas redevenu disponible après {self.delai_occupe_s} s")

    def _ouvrir(self, dst):
        """Ouvre par l'adresse SharePoint si le fichier est synchronisé, sinon par le chemin local."""
        url = url_sharepoint(dst)
        erreur = None
        for cible, mode in ((url, "co-édition SharePoint"), (dst, "fichier local OneDrive")):
            if not cible or (cible == url and mode != "co-édition SharePoint"):
                continue                     # source par lien : pas de copie locale à essayer
            try:
                # Open(Filename, UpdateLinks, ReadOnly, Format, Password, WriteResPassword,
                #      IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru)
                wb = self.xl.Workbooks.Open(cible, 0, False, None, None, None, True,
                                            None, None, False, False, None, False)
                return wb, mode
            except Exception as e:
                erreur = e
        raise erreur

    @staticmethod
    def _onglet(wb):
        """Onglet « Audit », à défaut la 1re feuille si elle a « ID erreur » (même règle que la lecture)."""
        noms = [str(wb.Worksheets(i).Name) for i in range(1, wb.Worksheets.Count + 1)]
        try:
            return wb.Worksheets(choisir_onglet(noms, ONGLET_AUDIT))
        except OngletIntrouvable as e:
            ws = wb.Worksheets(1)
            utile = ws.UsedRange
            der_col = utile.Column + utile.Columns.Count - 1
            entetes = [str(v).strip() for v in _plat(ws.Range(ws.Cells(1, 1), ws.Cells(1, der_col)).Value)
                       if v is not None]
            repli_premiere_feuille(e, ws.Name, entetes)
            return ws

    @staticmethod
    def _retirer_filtres(ws) -> dict:
        """Retire les filtres et réaffiche lignes et colonnes masquées de l'onglet.
        Les flèches de filtre restent en place.

        Ordre des commandes, de la plus tolérante à la plus stricte :
        Worksheet.ShowAllData échoue sur une feuille protégée (« La méthode ShowAllData de
        la classe Worksheet a échoué »), même si le filtrage y est autorisé, alors que
        AutoFilter.ShowAllData réussit. Elle ne sert donc qu'en dernier recours (filtre avancé).
        Sur une feuille protégée, une ligne ou colonne masquée À LA MAIN ne peut pas être
        réaffichée : c'est signalé (« protege »), sans effet sur la mise à jour, qui lit
        toute la zone utilisée. La protection n'est jamais retirée.
        Retourne {"filtres", "lignes", "colonnes", "protege": [ce qui reste masqué], "erreurs"}."""
        res = {"filtres": False, "lignes": False, "colonnes": [], "protege": [], "erreurs": []}
        protegee = bool(ws.ProtectContents)

        def message(e):
            return str(getattr(e, "excepinfo", None) and e.excepinfo[2] or e)

        def essai(action):
            try:
                action()
            except Exception as e:
                if _est_rejet(e):
                    raise
                res["erreurs"].append(message(e))

        def filtres():
            af = ws.AutoFilter if ws.AutoFilterMode else None
            if af is not None and af.FilterMode:
                af.ShowAllData(); res["filtres"] = True
            for i in range(1, ws.ListObjects.Count + 1):      # tableaux Excel : filtre propre à chacun
                af = ws.ListObjects(i).AutoFilter
                if af is not None and af.FilterMode:
                    af.ShowAllData(); res["filtres"] = True
            if ws.FilterMode:                                 # reste : filtre avancé (Données > Avancé)
                try:
                    ws.ShowAllData(); res["filtres"] = True
                except Exception as e:
                    if _est_rejet(e):
                        raise
                    if not protegee:
                        raise
                    res["protege"].append("filtre avancé")

        # Range.Hidden vaut False sur une plage partiellement masquée : il ne dit rien ici.
        def colonnes():
            ur = ws.UsedRange
            for c in range(ur.Column, ur.Column + ur.Columns.Count):
                if not ws.Columns(c).Hidden:
                    continue
                nom = str(ws.Cells(1, c).Value or get_column_letter(c))
                if protegee:
                    res["protege"].append(f"colonne {nom}"); continue
                ws.Columns(c).Hidden = False
                res["colonnes"].append(nom)

        def lignes():
            # SUBTOTAL(103) = cellules remplies VISIBLES ; COUNTA = toutes. Contrairement à
            # SpecialCells, ce calcul est autorisé sur une feuille protégée.
            ur, wf = ws.UsedRange, ws.Application.WorksheetFunction
            if wf.Subtotal(103, ur) >= wf.CountA(ur):
                return
            if protegee:
                res["protege"].append("lignes masquées"); return
            ur.EntireRow.Hidden = False; res["lignes"] = True

        for action in (filtres, colonnes, lignes):
            essai(action)
        return res

    def maj_lignes(self, src: str, dst: str, ids, colonnes=None) -> dict:
        """Même contrat que excel_rows.maj_lignes_excel, plus :
        "bloque" (None ou qui bloque), "non_verifies" ([ID]), "ouverture" (mode d'ouverture),
        "tentatives" (nombre d'essais). Le classeur est retenté entièrement si Excel
        reste occupé (l'écriture est idempotente : mêmes valeurs aux mêmes cellules)."""
        import time
        lecture = lire_cellules(src, ids, colonnes)
        for essai in range(1, self.TENTATIVES + 1):
            try:
                res = self._maj_une_fois(lecture, dst, colonnes)
                res["tentatives"] = essai
                return res
            except Exception as e:
                _liberer_objets_com(e)           # classeur, feuille, cellules : libérés ici, sur ce fil
                # Objet invalidé par une fusion SharePoint (0x800A01A8) : rouvrir le classeur suffit
                if not (_est_rejet(e) or isinstance(e, ExcelOccupe) or _est_objet_invalide(e)):
                    raise
                if essai == self.TENTATIVES:
                    if _est_objet_invalide(e):
                        raise ExcelErreur(f"classeur rechargé par la co-édition pendant l'écriture, "
                                          f"{self.TENTATIVES} fois de suite : rejouez plus tard") from e
                    raise ExcelOccupe(
                        f"Excel est resté occupé ({self.TENTATIVES} essais de {self.delai_occupe_s} s) — "
                        "synchronisation ou co-édition en cours ? Les ID sont mis en attente, "
                        "rejouez-les plus tard") from e
                time.sleep(5 * essai)

    def _maj_une_fois(self, lecture, dst, colonnes) -> dict:
        voulues = None if colonnes is None else {str(c).strip() for c in colonnes} - set(CLES)
        res = {"maj": [], "absents_src": lecture["absents_src"], "absents_dst": [],
               "donnees": lecture["donnees"], "ids_src": lecture["ids_src"], "col_absentes": [],
               "bloque": None, "non_verifies": [], "ouverture": None, "affichage": None, "videes": {},
               "onglets": {"src": lecture["onglet"], "dst": None}, "liens_refuses": {},
               "id_erreur": lecture.get("id_erreur", {})}
        a_ecrire = {i: c for i, c in lecture["cellules"].items() if c}
        if not a_ecrire:
            if voulues:
                res["col_absentes"] = sorted(voulues - set(lecture["entetes"]))
            return res                          # aucune colonne → on n'ouvre même pas le classeur

        self._vivant()
        wb, res["ouverture"] = self._ouvrir(dst)
        self._vivant()
        calcul = None
        externe = False      # classeur ouvert dans l'Excel de l'UTILISATEUR : jamais fermé ni réglé
        xl = self.xl
        try:
            # Enregistrement automatique coupé pour NOTRE instance : chaque sauvegarde auto
            # déclenche une fusion SharePoint (et sa fenêtre « Nous rechargeons votre classeur… »).
            # On enregistre une seule fois, à la fin. Sans effet pour les collègues.
            try: wb.AutoSaveOn = False
            except Exception as e:
                if _est_rejet(e): raise
            self._attendre_pret()
            if wb.ReadOnly:
                try: qui = wb.WriteReservedBy
                except Exception: qui = ""
                # Ouvert en écriture dans VOTRE propre Excel : notre instance ne l'a qu'en lecture
                # seule. On écrit alors dans votre classeur ouvert (cellules seulement), qui est
                # enregistré et reste ouvert. (Même solution que dans audit_validation_finale.)
                ouvert = _classeur_ouvert_ailleurs([url_sharepoint(dst), dst], self.xl)
                if ouvert is None:
                    # Ouvert en mode EXCLUSIF (Excel ancien, enregistrement automatique coupé,
                    # fichier .xls…) : la co-édition est impossible tant qu'il reste ouvert.
                    res["bloque"] = (f"{qui or 'un autre utilisateur'} (classeur ouvert en mode exclusif, pas en "
                                     "co-édition : Excel ne l'obtient qu'en lecture seule — il doit être fermé, ou "
                                     "ouvert dans Excel Online / Microsoft 365 avec l'enregistrement automatique)")
                    return res
                try: wb.Close(False)
                except Exception: pass
                wb, externe, xl = ouvert, True, ouvert.Application
                res["ouverture"] = "classeur déjà ouvert dans votre Excel (il reste ouvert)"
            ws = self._onglet(wb)
            res["onglets"]["dst"] = str(ws.Name)
            res["affichage"] = self._retirer_filtres(ws)
            # UsedRange et non End(xlUp/xlToLeft) : End s'arrête à la dernière ligne / colonne
            # VISIBLE, les ID filtrés passeraient pour absents et les colonnes masquées aussi.
            utile = ws.UsedRange
            der_col = utile.Column + utile.Columns.Count - 1
            der_lig = utile.Row + utile.Rows.Count - 1
            entetes = {str(v).strip(): c for c, v in enumerate(_plat(ws.Range(ws.Cells(1, 1), ws.Cells(1, der_col)).Value), 1)
                       if v is not None and str(v).strip()}
            if voulues:
                res["col_absentes"] = sorted(voulues - (set(lecture["entetes"]) & set(entetes)))
            lignes = _index_cles_com(ws, entetes, list(a_ecrire), der_lig)

            # Calcul suspendu pendant l'écriture : sinon chaque cellule relance un recalcul
            # complet du classeur et Excel refuse les appels suivants. Rétabli avant Save.
            # Jamais dans l'Excel de l'utilisateur : ses réglages ne sont pas touchés.
            if not externe:
                calcul = xl.Calculation
                xl.Calculation = _XL_CALC_MANUEL
            ecrits = {}                          # ID → [(cellule, attendu, lien)]
            for i, cells in a_ecrire.items():
                rangs = lignes.get(i)
                if not rangs:
                    res["absents_dst"].append(i); continue
                faits = []
                for r in rangs:                  # toutes les lignes d'un ID imm en doublon
                    for nom, cel in cells.items():
                        self._vivant()
                        c = entetes.get(nom)
                        if c is None:
                            continue
                        attendu = formule_pour(cel.valeur, cel.coord, f"{get_column_letter(c)}{r}")
                        cell = ws.Cells(r, c)
                        if attendu in (None, "") and cell.Formula != "" and nom not in res["videes"].get(i, []):
                            res["videes"].setdefault(i, []).append(nom)
                        try:
                            lien_ok = _ecrire(ws, cell, attendu, cel.lien, cel.format)
                        except Exception as e:
                            if not _est_rejet(e) and ws.ProtectContents:
                                raise ExcelErreur(f"onglet « {ONGLET_AUDIT} » protégé et cellule {nom} verrouillée : "
                                                  "ôtez la protection de la feuille puis rejouez") from e
                            raise
                        if not lien_ok and nom not in res["liens_refuses"].get(i, []):
                            res["liens_refuses"].setdefault(i, []).append(nom)
                        faits.append((cell, attendu if lien_ok or attendu not in (None, "") else cel.lien,
                                      cel.lien if lien_ok else None))
                if faits:
                    ecrits[i] = faits
            if calcul is not None:
                xl.Calculation = calcul; calcul = None   # enregistré avec son mode de calcul d'origine
            if not ecrits and not affichage_modifie(res["affichage"]):
                return res
            self._attendre_pret()

            wb.Save()
            self._vivant()
            self._attendre_pret()
            # Relecture : un ID n'est compté que si TOUTES ses cellules sont bien en place
            # Pas de « wb.Saved » ici : c'est un drapeau GLOBAL du classeur, remis à False par la
            # moindre activité de co-édition des collègues ; il rejetait à tort tous les ID d'une
            # commune. Save() a abouti sans erreur ; seule la relecture de NOS cellules fait foi.
            for i, faits in ecrits.items():
                self._vivant()
                if all(_conforme(cell, a, l) for cell, a, l in faits):
                    res["maj"].append(i)
                else:
                    res["non_verifies"].append(i)
            return res
        finally:
            if calcul is not None:
                try: xl.Calculation = calcul
                except Exception: pass
            if not externe:
                try: wb.Close(False)
                except Exception: pass


def _index_cles_com(ws, entetes, ids, der_lig) -> dict[str, list[int]]:
    """{ID → lignes} dans la feuille Excel : toutes les lignes d'un ID imm, la 1re d'un ID erreur."""
    besoin = {cle_de(i) for i in ids}
    for col in CLES:
        if col in besoin and col not in entetes:
            raise ValueError(f"colonne « {col} » absente de l'Excel SharePoint")
    voulus = {_norm(i) for i in ids}
    res: dict[str, list[int]] = {}
    if der_lig < 2:
        return res
    for col in besoin:
        c = entetes[col]
        for r, v in enumerate(_plat(ws.Range(ws.Cells(2, c), ws.Cells(der_lig, c)).Value), 2):
            k = _norm(v)
            if k not in voulus or cle_de(k) != col:
                continue
            if col == COL_ID_IMM:
                res.setdefault(k, []).append(r)
            elif k not in res:
                res[k] = [r]
    return res


def _cle_classeur(nom) -> str:
    """Forme comparable d'un chemin ou d'une adresse de classeur."""
    return unquote(str(nom or "")).replace("\\", "/").strip().rstrip("/").lower()


def _classeur_ouvert_ailleurs(cibles, notre_excel=None):
    """Classeur `cibles` (adresse SharePoint ou chemin) ouvert EN ÉCRITURE dans une autre instance
    d'Excel de cette session Windows — celle de l'utilisateur —, trouvé dans la table des objets COM
    en cours (ROT). None sinon. Ne lève jamais. (Repris d'audit_validation_finale.)"""
    cles = {_cle_classeur(c) for c in cibles if c}
    if not cles:
        return None
    try:
        pythoncom = _pythoncom()
        import win32com.client as wc
        rot = pythoncom.GetRunningObjectTable()
        contexte = pythoncom.CreateBindCtx(0)
        try:
            notre = notre_excel.Hwnd if notre_excel is not None else None
        except Exception:
            notre = None
        for moniker in rot:
            try:
                if _cle_classeur(moniker.GetDisplayName(contexte, None)) not in cles:
                    continue
                classeur = wc.Dispatch(rot.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch))
                if _cle_classeur(classeur.FullName) not in cles or classeur.ReadOnly:
                    continue
                if notre is not None and classeur.Application.Hwnd == notre:
                    continue
                return classeur
            except Exception:
                continue
    except Exception:
        return None
    return None


# ── session surveillée (anti-blocage) ─────────────────────────────
class ExcelBloque(Exception):
    """Excel n'a plus donné signe de vie : très probablement une fenêtre modale invisible."""


class ConnexionRequise(ExcelBloque):
    """Excel affiche la fenêtre de connexion Microsoft 365 : dans une instance invisible,
    personne ne peut y répondre. Détectée tout de suite (au lieu d'attendre le délai anti-blocage)."""


class SessionExcelSurveillee:
    """Même usage que SessionExcel, avec un garde-fou contre le blocage définitif.

    Pourquoi : Excel tourne invisible. Si SharePoint affiche une boîte modale (« Nous
    rechargeons votre classeur pour terminer la fusion »…), personne ne peut cliquer dessus et
    l'appel COM en cours (souvent Save) ne revient JAMAIS : sans garde-fou, l'outil gèle.

    Comment : la session Excel vit sur un fil dédié, propriétaire de son appartement COM, qui
    exécute les classeurs un par un. Le fil appelant attend le résultat en surveillant
    SessionExcel.activite (signes de vie : réponses d'Excel, refus « occupé » réessayés,
    avancement cellule par cellule). Sans signe de vie pendant `delai_blocage_s`, le processus
    de CETTE instance est tué (jamais l'Excel de l'utilisateur), on attend sa disparition, une
    instance neuve est lancée et le classeur est retenté. Au pire une commune échoue (ExcelBloque)
    et ses ID rejoignent « ID en attente » : l'outil ne gèle plus.
    """

    def __init__(self, visible=False, delai_occupe_s=120, delai_blocage_s=120, essais=2,
                 connexion_interactive=True, delai_connexion_s=300):
        self.visible, self.delai_occupe_s = visible, delai_occupe_s
        self.delai_blocage_s, self.essais = delai_blocage_s, essais
        # Fenêtre de connexion Microsoft 365 : l'afficher (Excel visible) pour que l'utilisateur
        # se connecte une fois, puis reprendre — plutôt qu'échouer à chaque fichier.
        self.connexion_interactive, self.delai_connexion_s = connexion_interactive, delai_connexion_s
        self._connexion_faite = False
        self._fil = None
        self.relances = 0          # nombre d'instances Excel tuées puis relancées

    # cycle de vie ---------------------------------------------------
    def __enter__(self):
        self._demarrer()
        return self

    def __exit__(self, *exc):
        self._arreter(proprement=True)
        return False

    def _demarrer(self):
        import queue, threading
        etat = {"pret": threading.Event(), "jobs": queue.Queue(), "session": None, "erreur": None}

        def boucle():
            try:
                s = SessionExcel(visible=self.visible, delai_occupe_s=self.delai_occupe_s)
                s.__enter__()
            except BaseException as e:
                etat["erreur"] = e; etat["pret"].set(); return
            etat["session"] = s; etat["pret"].set()
            try:
                while True:
                    job = etat["jobs"].get()
                    if job is None:
                        break
                    fn, args, boite = job
                    try:
                        boite["res"] = getattr(s, fn)(*args)
                    except BaseException as e:
                        # Ici, sur le fil propriétaire : sinon les objets COM retenus par la trace
                        # seraient libérés plus tard depuis un autre fil (plantage de Python).
                        _liberer_objets_com(e)
                        boite["err"] = e
                    boite["fait"].set()
            finally:
                try: s.__exit__(None, None, None)
                except BaseException: pass

        fil = threading.Thread(target=boucle, name="excel-com", daemon=True)
        fil.start()
        if not etat["pret"].wait(max(60, self.delai_blocage_s)):
            raise ExcelIndisponible("Excel n'a pas démarré dans le délai")
        if etat["erreur"] is not None:
            e = etat["erreur"]
            raise e if isinstance(e, ExcelIndisponible) else ExcelIndisponible(str(e))
        self._fil, self._etat = fil, etat

    def _arreter(self, proprement):
        if self._fil is None:
            return
        etat, fil, self._fil = self._etat, self._fil, None
        if proprement:
            etat["jobs"].put(None)
            fil.join(60)
        if fil.is_alive() and etat["session"] is not None:   # bloqué : on coupe court
            _tuer(etat["session"]._pid)
            fil.join(20)

    def _relancer(self):
        """Tue l'instance bloquée, attend sa disparition, en démarre une neuve."""
        etat, self._fil = self._etat, None
        _tuer(etat["session"]._pid)
        etat["jobs"].put(None)              # le fil débloqué (appel COM en erreur) s'arrêtera
        self.relances += 1
        self._demarrer()

    # travail ----------------------------------------------------------
    def _executer(self, fn, *args):
        import threading, time
        boite = {"fait": threading.Event()}
        session = self._etat["session"]
        session._vivant()
        self._etat["jobs"].put((fn, args, boite))
        while not boite["fait"].wait(1.0):
            if fenetre_connexion(session._pid):
                raise ConnexionRequise("Excel demande une connexion à Microsoft 365 (compte professionnel)")
            if time.time() - session.activite > self.delai_blocage_s:
                raise ExcelBloque(f"Excel ne répond plus depuis {self.delai_blocage_s} s "
                                  "(fenêtre bloquante invisible, ex. « Nous rechargeons votre classeur… »)")
        if "err" in boite:
            raise boite["err"]
        return boite["res"]

    def maj_lignes(self, src, dst, ids, colonnes=None) -> dict:
        """SessionExcel.maj_lignes, protégé contre le blocage et la perte de l'instance Excel."""
        derniere = None
        essai = 0
        while essai < self.essais:
            essai += 1
            if self._fil is None:
                self._demarrer()
            try:
                res = self._executer("maj_lignes", src, dst, ids, colonnes)
                if essai > 1:
                    res["relance_excel"] = essai - 1
                return res
            except ConnexionRequise as e:
                self._relancer()                 # l'instance invisible reste bloquée sur la fenêtre
                adresse = url_sharepoint(dst)
                if not self.connexion_interactive or self._connexion_faite or not adresse:
                    raise ConnexionRequise(f"{e} — ouvrez une fois un fichier du site SharePoint dans Excel, "
                                           "connectez-vous avec votre compte professionnel, puis rejouez") from None
                self._connexion_faite = True
                connexion_microsoft365(adresse, self.delai_connexion_s)
                essai -= 1                       # la connexion ne compte pas comme un essai
                continue
            except ExcelBloque as e:
                derniere = e
            except Exception as e:
                # Excel a disparu (plantage, fermeture forcée) : pywin32 masque parfois la
                # déconnexion en AttributeError ; on vérifie donc le processus lui-même.
                if not (_est_deconnexion(e) or not _processus_vivant(self._etat["session"]._pid)):
                    raise
                derniere = e
            self._relancer()
        raise ExcelBloque(f"{derniere} — {self.essais} essai(s) avec une instance Excel neuve, "
                          "commune mise en attente") from derniere
