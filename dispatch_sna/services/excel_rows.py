"""Excel audit : lecture des en-têtes / lignes, remplacement des lignes par ID erreur."""
from __future__ import annotations

import os
from copy import copy
from dataclasses import dataclass

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string, get_column_letter

from .constantes import ONGLET_AUDIT, COL_ID_ERR, COL_REMARQUE, COL_LIEN_SV


def _norm(v) -> str:
    return str(v).strip().upper() if v is not None else ""


class ExcelErreur(Exception):
    """Structure inattendue (colonne ID erreur absente...)."""


class OngletIntrouvable(ExcelErreur):
    """Aucun onglet au nom d'« Audit » (≠ plusieurs candidats, qui reste une ExcelErreur)."""


def _cle_onglet(nom) -> str:
    return "".join(str(nom).split()).replace("_", "").replace("-", "").casefold()


def choisir_onglet(noms, nom=ONGLET_AUDIT) -> str:
    """Nom de l'onglet Audit parmi `noms` : « Audit » (espaces / casse ignorés), sinon
    l'UNIQUE onglet dont le nom contient « audit » (« Audit SNA », « Audit_38068 »…).
    Jamais l'onglet actif : il dépend du dernier enregistrement (DGFIP, IPE…).
    Lève ExcelErreur si aucun ou plusieurs onglets possibles."""
    cible = _cle_onglet(nom)
    noms = [str(n) for n in noms]
    exacts = [n for n in noms if _cle_onglet(n) == cible]
    if exacts:
        return exacts[0]
    proches = [n for n in noms if cible in _cle_onglet(n)]
    if len(proches) == 1:
        return proches[0]
    if proches:
        raise ExcelErreur(f"plusieurs onglets possibles pour « {nom} » : {', '.join(proches)}")
    raise OngletIntrouvable(f"onglet « {nom} » introuvable (onglets présents : {', '.join(noms)})")


def est_onglet_audit(nom, cible=ONGLET_AUDIT) -> bool:
    return _cle_onglet(cible) in _cle_onglet(nom)


def repli_premiere_feuille(erreur, titre, entetes):
    """Sans onglet Audit, la 1re feuille sert d'onglet Audit si sa ligne 1 contient
    « ID erreur » (sinon ce n'est pas un fichier audit : on refuse plutôt que d'écrire ailleurs)."""
    if COL_ID_ERR in entetes:
        return
    raise OngletIntrouvable(f"{erreur} ; la 1re feuille « {titre} » n'a pas non plus "
                            f"de colonne « {COL_ID_ERR} » en ligne 1") from None


def _ligne1(ws) -> list[str]:
    ligne = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    return [str(v).strip() for v in ligne if v is not None and str(v).strip()]


def trouver_onglet(wb, nom=ONGLET_AUDIT):
    """Onglet « Audit » (cf. choisir_onglet) ; à défaut la 1re feuille, si elle a « ID erreur »."""
    try:
        return wb[choisir_onglet(wb.sheetnames, nom)]
    except OngletIntrouvable as e:
        premiere = wb.worksheets[0]
        repli_premiere_feuille(e, premiere.title, _ligne1(premiere))
        return premiere


def retirer_filtres(ws) -> dict:
    """Retire les critères de filtre et réaffiche lignes et colonnes masquées d'un onglet.
    Les flèches de filtre restent en place. Retourne {"filtres", "lignes", "colonnes"}."""
    res = {"filtres": bool(ws.auto_filter.filterColumn), "lignes": False, "colonnes": []}
    ws.auto_filter.filterColumn = []
    for table in ws.tables.values():
        if table.autoFilter is not None and table.autoFilter.filterColumn:
            table.autoFilter.filterColumn = []
            res["filtres"] = True
    for d in ws.row_dimensions.values():
        if d.hidden:
            d.hidden = False; res["lignes"] = True
    for d in ws.column_dimensions.values():
        if d.hidden:
            d.hidden = False
            if not d.width:
                d.width = 13
            debut = d.min or column_index_from_string(d.index)
            res["colonnes"] += [str(ws.cell(1, c).value or get_column_letter(c))
                                for c in range(debut, (d.max or debut) + 1) if c <= ws.max_column]
    return res


def affichage_modifie(a: dict | None) -> bool:
    return bool(a and (a["filtres"] or a["lignes"] or a["colonnes"]))


def retirer_filtres_fichier(chemin: str) -> dict:
    """Retire les filtres de l'onglet Audit d'un fichier, réécrit seulement s'il y en avait."""
    wb = load_workbook(chemin)
    res = retirer_filtres(trouver_onglet(wb))
    if affichage_modifie(res):
        _enregistrer(wb, chemin)
    return res


def _enregistrer(wb, chemin: str):
    tmp = chemin + ".tmp.xlsx"
    try:
        wb.save(tmp)
        os.replace(tmp, chemin)
    finally:
        if os.path.isfile(tmp):
            os.remove(tmp)


def _entetes(ws) -> dict[str, int]:
    return {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}


def _index_ids(ws, col_id: int) -> dict[str, int]:
    """{ID erreur → n° de ligne} (première occurrence)."""
    idx = {}
    for r in range(2, ws.max_row + 1):
        v = _norm(ws.cell(r, col_id).value)
        if v and v not in idx:
            idx[v] = r
    return idx


def _valeur_cellule(cell) -> str:
    """Texte d'une cellule — l'hyperlien est préféré au texte affiché."""
    hl = getattr(cell, "hyperlink", None)
    cible = getattr(hl, "target", None) if hl is not None else None
    if cible:
        return str(cible).strip()
    return "" if cell.value is None else str(cell.value).strip()


def formule_pour(valeur, origine: str, cible: str):
    """Recale les références relatives d'une formule copiée de `origine` vers `cible` (ex. J57 → J3).
    Toute autre valeur est rendue telle quelle."""
    if isinstance(valeur, str) and valeur.startswith("=") and origine != cible:
        try: return Translator(valeur, origin=origine).translate_formula(cible)
        except Exception: pass
    return valeur


def _copier_cellule(src, dst):
    dst.value = formule_pour(src.value, src.coordinate, dst.coordinate)
    dst.hyperlink = copy(src.hyperlink) if src.hyperlink else None
    if src.has_style:
        dst.font = copy(src.font); dst.fill = copy(src.fill); dst.border = copy(src.border)
        dst.alignment = copy(src.alignment); dst.number_format = src.number_format
        dst.protection = copy(src.protection)


def lire_entetes(chemin: str) -> list[str]:
    """En-têtes (ligne 1) de l'onglet Audit, dans l'ordre du fichier."""
    wb = load_workbook(chemin, read_only=True)
    try:
        ws = trouver_onglet(wb)
        ligne = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        return list(dict.fromkeys(str(v).strip() for v in ligne if v is not None and str(v).strip()))
    finally:
        wb.close()


def lire_lignes(chemin: str, ids) -> dict[str, dict[str, str]]:
    """{ID → {"Remarque": ..., "Lien Street View": ...}} pour les ID trouvés."""
    wb = load_workbook(chemin)
    ws = trouver_onglet(wb)
    h = _entetes(ws)
    if COL_ID_ERR not in h:
        return {}
    idx = _index_ids(ws, h[COL_ID_ERR])
    res = {}
    for i in ids:
        r = idx.get(_norm(i))
        if r is None:
            continue
        res[_norm(i)] = {col: (_valeur_cellule(ws.cell(r, h[col])) if col in h else "")
                         for col in (COL_REMARQUE, COL_LIEN_SV)}
    return res




@dataclass
class Cellule:
    """Contenu d'une cellule source à reporter : valeur (ou formule), lien, format, position."""
    valeur: object
    lien: str | None
    format: str
    coord: str


def lire_cellules(src: str, ids, colonnes=None) -> dict:
    """Lit dans l'Excel Prod IA les cellules à reporter, pour une écriture cellule par cellule.

    `colonnes` : None = toutes, [] = aucune (même convention que maj_lignes_excel).
    Retourne {"entetes": [...], "cellules": {ID → {colonne → Cellule}},
              "donnees": {ID → {Remarque, Lien Street View}}, "ids_src": [...], "absents_src": [...]}.
    """
    wanted = list(dict.fromkeys(_norm(i) for i in ids))
    wb = load_workbook(src); ws = trouver_onglet(wb)
    h = _entetes(ws)
    if COL_ID_ERR not in h:
        raise ExcelErreur(f"colonne « {COL_ID_ERR} » absente de l'Excel Prod IA")
    idx = _index_ids(ws, h[COL_ID_ERR])
    voulues = None if colonnes is None else {str(c).strip() for c in colonnes} - {COL_ID_ERR}
    noms = [n for n in h if n != COL_ID_ERR and (voulues is None or n in voulues)]
    res = {"entetes": list(h), "cellules": {}, "donnees": {}, "ids_src": list(idx), "absents_src": [],
           "onglet": ws.title}
    for i in wanted:
        r = idx.get(i)
        if r is None:
            res["absents_src"].append(i); continue
        res["donnees"][i] = {col: (_valeur_cellule(ws.cell(r, h[col])) if col in h else "")
                             for col in (COL_REMARQUE, COL_LIEN_SV)}
        cells = {}
        for n in noms:
            c = ws.cell(r, h[n])
            lien = getattr(c.hyperlink, "target", None) if c.hyperlink else None
            cells[n] = Cellule(c.value, lien, c.number_format or "General", c.coordinate)
        res["cellules"][i] = cells
    return res


def maj_lignes_excel(src: str, dst: str, ids, colonnes=None) -> dict:
    """Remplace dans dst les lignes des ID donnés par celles de src.

    Colonnes associées par nom d'en-tête ; valeurs, hyperliens et mise en forme
    copiés. `colonnes` : noms des colonnes à mettre à jour (None = toutes les
    colonnes communes, [] = aucune → dst non modifié). Les autres lignes et
    colonnes de dst ne sont pas touchées. Retourne
    {"maj": [...], "absents_src": [...], "absents_dst": [...],
     "donnees": {ID → {Remarque, Lien Street View}}, "ids_src": [...],
     "col_absentes": [...], "affichage": résultat de retirer_filtres sur dst,
     "videes": {ID → [colonnes dont l'ancienne valeur a été effacée]}}.
    Une colonne cochée vide dans src efface donc la valeur de dst.
    Lève PermissionError si dst est verrouillé.
    """
    wanted = list(dict.fromkeys(_norm(i) for i in ids))
    wb_s = load_workbook(src); ws_s = trouver_onglet(wb_s)
    wb_d = load_workbook(dst); ws_d = trouver_onglet(wb_d)
    h_s, h_d = _entetes(ws_s), _entetes(ws_d)
    for nom, h in (("Prod IA", h_s), ("SharePoint", h_d)):
        if COL_ID_ERR not in h:
            raise ExcelErreur(f"colonne « {COL_ID_ERR} » absente de l'Excel {nom}")
    idx_s = _index_ids(ws_s, h_s[COL_ID_ERR])
    idx_d = _index_ids(ws_d, h_d[COL_ID_ERR])
    voulues = None if colonnes is None else {str(c).strip() for c in colonnes} - {COL_ID_ERR}
    paires = [(h_s[n], c) for n, c in h_d.items()
              if n in h_s and n != COL_ID_ERR and (voulues is None or n in voulues)]

    res = {"maj": [], "absents_src": [], "absents_dst": [], "donnees": {},
           "ids_src": list(idx_s), "videes": {},
           "onglets": {"src": ws_s.title, "dst": ws_d.title},
           "col_absentes": sorted(voulues - (h_s.keys() & h_d.keys())) if voulues else [],
           "affichage": retirer_filtres(ws_d) if paires else None}
    for i in wanted:
        rs = idx_s.get(i)
        if rs is None:
            res["absents_src"].append(i); continue
        res["donnees"][i] = {col: (_valeur_cellule(ws_s.cell(rs, h_s[col])) if col in h_s else "")
                             for col in (COL_REMARQUE, COL_LIEN_SV)}
        if not paires:
            continue
        rd = idx_d.get(i)
        if rd is None:
            res["absents_dst"].append(i); continue
        for cs, cd in paires:
            c_s, c_d = ws_s.cell(rs, cs), ws_d.cell(rd, cd)
            if c_s.value in (None, "") and c_d.value not in (None, ""):
                res["videes"].setdefault(i, []).append(str(ws_d.cell(1, cd).value))
            _copier_cellule(c_s, c_d)
        res["maj"].append(i)

    if res["maj"] or affichage_modifie(res["affichage"]):
        _enregistrer(wb_d, dst)
    return res
