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
_EPOQUE = _dt.datetime(1899, 12, 30)


class ExcelIndisponible(Exception):
    """pywin32 absent ou Excel non installé / impossible à démarrer."""


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

    def __init__(self, visible=False):
        self.visible = visible
        self.xl = None

    def __enter__(self):
        try:
            pythoncom = _pythoncom()
            import win32com.client as wc
        except ImportError as e:
            raise ExcelIndisponible("pywin32 n'est pas installé") from e
        pythoncom.CoInitialize()
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
        _pythoncom().CoUninitialize()
        return False

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
        "bloque" (None ou qui bloque), "non_verifies" ([ID]), "ouverture" (mode d'ouverture)."""
        lecture = lire_cellules(src, ids, colonnes)
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
        try:
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
            if not ecrits:
                return res

            wb.Save()
            # Relecture : un ID n'est compté que si TOUTES ses cellules sont bien en place
            for i, faits in ecrits.items():
                if wb.Saved and all(_conforme(cell, a, l) for cell, a, l in faits):
                    res["maj"].append(i)
                else:
                    res["non_verifies"].append(i)
            return res
        finally:
            try: wb.Close(False)
            except Exception: pass
