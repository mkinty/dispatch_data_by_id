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

from .constantes import ONGLET_AUDIT, COL_ID_ERR
from .excel_rows import lire_cellules, formule_pour, _norm
from .onedrive import url_sharepoint

_EN = 1033                 # propriétés NumberFormat / Formula toujours en anglais (sinon « General » est refusé en français)
_XL_UP, _XL_TO_LEFT = -4162, -4159
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


class _FiltreMessages:
    """Filtre de messages COM (IMessageFilter) pour le thread courant.

    Sans lui, tout appel envoyé pendant qu'Excel est occupé échoue aussitôt avec
    RPC_E_CALL_REJECTED. Avec lui, Windows réessaie l'appel toutes les 250 ms
    jusqu'à `delai_max_ms`. pywin32 n'exposant pas CoRegisterMessageFilter,
    l'interface est construite avec ctypes (vtable de 6 méthodes).
    """
    _ATTENTE_MS = 250

    def __init__(self, delai_max_ms=120_000):
        import ctypes
        import uuid
        from ctypes import wintypes
        self.delai_max_ms = delai_max_ms
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
def _ecrire(ws, cell, attendu, lien, fmt):
    if isinstance(attendu, str) and attendu.startswith("="):
        _put_en(cell, "Formula", attendu)
    elif attendu is None or attendu == "":
        cell.ClearContents()
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
        if isinstance(attendu, str) and attendu:
            ws.Hyperlinks.Add(cell, lien, "", "", attendu)
        else:
            ws.Hyperlinks.Add(cell, lien)      # garde la valeur de la cellule (affiche le lien si vide)
    elif not lien and cell.Hyperlinks.Count:
        cell.Hyperlinks.Delete()


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
    return ok


# ── session ───────────────────────────────────────────────────────
class SessionExcel:
    """Une instance Excel dédiée (jamais celle de l'utilisateur), fermée à la sortie.

        with SessionExcel() as xl:
            r = xl.maj_lignes(src, dst, ids, colonnes)
    """

    TENTATIVES = 3              # essais complets d'un classeur si Excel reste occupé au-delà du délai

    def __init__(self, visible=False, delai_occupe_s=120):
        self.visible = visible
        self.delai_occupe_s = delai_occupe_s
        self.xl = None
        self._filtre = None

    def __enter__(self):
        try:
            pythoncom = _pythoncom()
            import win32com.client as wc
        except ImportError as e:
            raise ExcelIndisponible("pywin32 n'est pas installé") from e
        pythoncom.CoInitialize()
        # Avant tout appel à Excel : les refus « occupé » seront réessayés au lieu d'échouer
        self._filtre = _FiltreMessages(delai_max_ms=int(self.delai_occupe_s * 1000))
        self._filtre.installer()
        try:
            self.xl = wc.DispatchEx("Excel.Application")   # nouvelle instance, n'interfère pas avec l'Excel ouvert
            self.xl.Visible = bool(self.visible)
            self.xl.DisplayAlerts = False
            self.xl.AskToUpdateLinks = False
            self.xl.ScreenUpdating = bool(self.visible)
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
        return False

    def _attendre_pret(self):
        """Attend qu'Excel ait fini (chargement, synchro co-édition, recalcul, envoi)."""
        import time
        limite = time.time() + self.delai_occupe_s
        while time.time() < limite:
            try:
                if self.xl.Ready:
                    return
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
            if not cible:
                continue
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
        cible = "".join(ONGLET_AUDIT.split()).casefold()
        for i in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(i)
            if "".join(str(ws.Name).split()).casefold() == cible:
                return ws
        return wb.ActiveSheet

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
                if not (_est_rejet(e) or isinstance(e, ExcelOccupe)):
                    raise
                if essai == self.TENTATIVES:
                    raise ExcelOccupe(
                        f"Excel est resté occupé ({self.TENTATIVES} essais de {self.delai_occupe_s} s) — "
                        "synchronisation ou co-édition en cours ? Les ID sont mis en attente, "
                        "rejouez-les plus tard") from e
                time.sleep(5 * essai)

    def _maj_une_fois(self, lecture, dst, colonnes) -> dict:
        voulues = None if colonnes is None else {str(c).strip() for c in colonnes} - {COL_ID_ERR}
        res = {"maj": [], "absents_src": lecture["absents_src"], "absents_dst": [],
               "donnees": lecture["donnees"], "ids_src": lecture["ids_src"], "col_absentes": [],
               "bloque": None, "non_verifies": [], "ouverture": None}
        a_ecrire = {i: c for i, c in lecture["cellules"].items() if c}
        if not a_ecrire:
            if voulues:
                res["col_absentes"] = sorted(voulues - set(lecture["entetes"]))
            return res                          # aucune colonne → on n'ouvre même pas le classeur

        wb, res["ouverture"] = self._ouvrir(dst)
        calcul = None
        try:
            self._attendre_pret()
            if wb.ReadOnly:
                try: qui = wb.WriteReservedBy
                except Exception: qui = ""
                res["bloque"] = qui or "un autre utilisateur"
                return res
            ws = self._onglet(wb)
            der_col = ws.Cells(1, ws.Columns.Count).End(_XL_TO_LEFT).Column
            entetes = {str(v).strip(): c for c, v in enumerate(_plat(ws.Range(ws.Cells(1, 1), ws.Cells(1, der_col)).Value), 1)
                       if v is not None and str(v).strip()}
            if COL_ID_ERR not in entetes:
                raise ValueError(f"colonne « {COL_ID_ERR} » absente de l'Excel SharePoint")
            if voulues:
                res["col_absentes"] = sorted(voulues - (set(lecture["entetes"]) & set(entetes)))
            c_id = entetes[COL_ID_ERR]
            der_lig = ws.Cells(ws.Rows.Count, c_id).End(_XL_UP).Row
            lignes = {}
            if der_lig >= 2:
                for r, v in enumerate(_plat(ws.Range(ws.Cells(2, c_id), ws.Cells(der_lig, c_id)).Value), 2):
                    k = _norm(v)
                    if k and k not in lignes:
                        lignes[k] = r

            # Calcul suspendu pendant l'écriture : sinon chaque cellule relance un recalcul
            # complet du classeur et Excel refuse les appels suivants. Rétabli avant Save.
            calcul = self.xl.Calculation
            self.xl.Calculation = _XL_CALC_MANUEL
            ecrits = {}                          # ID → [(cellule, attendu, lien)]
            for i, cells in a_ecrire.items():
                r = lignes.get(i)
                if r is None:
                    res["absents_dst"].append(i); continue
                faits = []
                for nom, cel in cells.items():
                    c = entetes.get(nom)
                    if c is None:
                        continue
                    attendu = formule_pour(cel.valeur, cel.coord, f"{get_column_letter(c)}{r}")
                    cell = ws.Cells(r, c)
                    _ecrire(ws, cell, attendu, cel.lien, cel.format)
                    faits.append((cell, attendu, cel.lien))
                if faits:
                    ecrits[i] = faits
            self.xl.Calculation = calcul; calcul = None   # le fichier est enregistré avec son mode de calcul d'origine
            if not ecrits:
                return res
            self._attendre_pret()

            wb.Save()
            self._attendre_pret()
            # Relecture : un ID n'est compté que si TOUTES ses cellules sont bien en place
            for i, faits in ecrits.items():
                if wb.Saved and all(_conforme(cell, a, l) for cell, a, l in faits):
                    res["maj"].append(i)
                else:
                    res["non_verifies"].append(i)
            return res
        finally:
            if calcul is not None:
                try: self.xl.Calculation = calcul
                except Exception: pass
            try: wb.Close(False)
            except Exception: pass
