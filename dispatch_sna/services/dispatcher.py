"""Orchestration du rangement : commune complète, ou mise à jour ciblée par ID erreur.
Aucune dépendance à l'interface : la progression passe par log_fn / phase_fn."""
import os
import shutil

from .constantes import (AUDIT_PREFIX, PPT_FOLDER, COL_REMARQUE, COL_LIEN_SV,
                         LOG_OK, LOG_INFO, LOG_WARN, LOG_ERR)
from .chemins import get_dep, onedrive_pin, copier_remplacer
from .error_ids import nom_ppt
from .excel_rows import maj_lignes_excel, lire_lignes
from .ppt_editor import modifier_ppt, copier_ppt


# ── Traitement d'une commune ──────────────────────────────────────
def ranger_commune(prod_ia, root_sp, commune, dossier_src, supprimer, log_fn, phase_fn=None):
    """Range les fichiers d'une commune depuis Prod IA vers Dep/commune.
    Retourne un dict stats ou None si erreur bloquante."""
    def phase(n, e="run"):
        if phase_fn: phase_fn(n, e)

    dep = get_dep(commune)
    dossier_dst = os.path.join(root_sp, dep, commune)
    ppt_dst = os.path.join(dossier_dst, PPT_FOLDER)

    phase("verif", "run")
    # 1. Vérifier que la destination existe
    if not os.path.isdir(dossier_dst):
        log_fn(f"  ❌ Destination introuvable : {dep}/{commune}", LOG_ERR)
        log_fn(f"     → dossier laissé dans Prod IA (non supprimé)", LOG_WARN)
        phase("verif", "err")
        return None
    phase("verif", "ok")

    # 2. Lister les fichiers source
    phase("analyse", "run")
    fichiers = os.listdir(dossier_src)
    excels = [f for f in fichiers if f.lower().endswith(".xlsx") and not f.startswith("~")]
    ppts   = [f for f in fichiers if f.lower().endswith(".pptx")]
    log_fn(f"  📁 {len(excels)} Excel · {len(ppts)} PPT à ranger", LOG_INFO)
    phase("analyse", "ok")

    stats = {"excel": 0, "ppt": 0, "err": 0}

    # 3. Ranger l'Excel (remplace le vierge du client)
    phase("excel", "run")
    for f in excels:
        src = os.path.join(dossier_src, f)
        dst = os.path.join(dossier_dst, f"{AUDIT_PREFIX}{commune}.xlsx")
        if copier_remplacer(src, dst, log_fn):
            stats["excel"] += 1
            log_fn(f"  📊 Excel → {dep}/{commune}/ (remplacé)", LOG_OK)
        else:
            stats["err"] += 1
    phase("excel", "ok" if stats["err"] == 0 else "err")

    # 4. Ranger les PPT dans analyse/
    phase("ppt", "run")
    os.makedirs(ppt_dst, exist_ok=True)
    for f in ppts:
        src = os.path.join(dossier_src, f)
        dst = os.path.join(ppt_dst, f)
        if copier_remplacer(src, dst, log_fn):
            stats["ppt"] += 1
        else:
            stats["err"] += 1
    log_fn(f"  🖼️  {stats['ppt']} PPT → {dep}/{commune}/{PPT_FOLDER}/", LOG_OK)
    phase("ppt", "ok" if stats["err"] == 0 else "err")

    # 5. Supprimer le dossier commune de Prod IA (si tout a réussi)
    phase("nettoyage", "run")
    if stats["err"] == 0:
        if supprimer:
            try:
                shutil.rmtree(dossier_src)
                log_fn(f"  🗑️  Dossier Prod IA/{commune} supprimé", LOG_WARN)
            except Exception as e:
                log_fn(f"  ⚠️ Impossible de supprimer Prod IA/{commune} : {e}", LOG_WARN)
        else:
            log_fn(f"  ⏭️ Suppression désactivée (dossier gardé)", LOG_INFO)
        phase("nettoyage", "ok")
    else:
        log_fn(f"  ⚠️ {stats['err']} erreur(s) → dossier Prod IA/{commune} conservé", LOG_WARN)
        phase("nettoyage", "err")

    return stats


# ── Traitement ciblé par ID erreur ────────────────────────────────
def ranger_ids(root_sp, commune, dossier_src, ids, ppt_mode, supprimer, log_fn, phase_fn=None,
               colonnes=None):
    """Met à jour Dep/commune uniquement pour les ID erreur donnés.
    ppt_mode : "remplacer", "modifier" ou None (Excel seul).
    colonnes : colonnes Excel à mettre à jour (None = toutes, [] = aucune).
    Retourne un dict stats ou None si erreur bloquante."""
    def phase(n, e="run"):
        if phase_fn: phase_fn(n, e)

    dep = get_dep(commune)
    dossier_dst = os.path.join(root_sp, dep, commune)
    ppt_dst = os.path.join(dossier_dst, PPT_FOLDER)
    excel_dst = os.path.join(dossier_dst, f"{AUDIT_PREFIX}{commune}.xlsx")
    stats = {"lignes": 0, "excel": 0, "ppt": 0, "err": 0}

    # 1. Vérifications
    phase("verif", "run")
    if not dossier_src or not os.path.isdir(dossier_src):
        log_fn(f"  ❌ Aucun dossier {commune} dans Prod IA", LOG_ERR)
        phase("verif", "err"); return None
    if not os.path.isdir(dossier_dst):
        log_fn(f"  ❌ Destination introuvable : {dep}/{commune}", LOG_ERR)
        phase("verif", "err"); return None
    if not os.path.isfile(excel_dst):
        log_fn(f"  ❌ Excel SharePoint introuvable : {dep}/{commune}/{os.path.basename(excel_dst)}", LOG_ERR)
        phase("verif", "err"); return None
    phase("verif", "ok")

    # 2. Excel source
    phase("analyse", "run")
    excels = [f for f in os.listdir(dossier_src) if f.lower().endswith(".xlsx") and not f.startswith("~")]
    nom_src = f"{AUDIT_PREFIX}{commune}.xlsx"
    nom_src = nom_src if nom_src in excels else (excels[0] if excels else None)
    if not nom_src:
        log_fn(f"  ❌ Aucun Excel dans Prod IA/{commune}", LOG_ERR)
        phase("analyse", "err"); return None
    excel_src = os.path.join(dossier_src, nom_src)
    log_fn(f"  🔎 {len(ids)} ID erreur · source {nom_src}", LOG_INFO)
    phase("analyse", "ok")

    # 3. Excel : remplacement des seules lignes des ID
    phase("excel", "run")
    donnees, ids_src = None, []
    try:
        r = maj_lignes_excel(excel_src, excel_dst, ids, colonnes)
        donnees, ids_src = r["donnees"], r["ids_src"]
        stats["lignes"] = len(r["maj"]); stats["excel"] = 1 if r["maj"] else 0
        if colonnes is not None and not colonnes:
            log_fn("  ⏭️ Excel non modifié (aucune colonne cochée)", LOG_INFO)
        if r["maj"]:
            quoi = "toutes colonnes" if colonnes is None else f"{len(colonnes)} colonne(s)"
            log_fn(f"  📊 {len(r['maj'])} ligne(s) mise(s) à jour ({quoi}) dans {dep}/{commune}/{os.path.basename(excel_dst)}", LOG_OK)
        if r["col_absentes"]:
            log_fn(f"    ⚠️ Colonne(s) cochée(s) absente(s) d'un des deux Excel, ignorée(s) : {', '.join(r['col_absentes'])}", LOG_WARN)
        if r["absents_src"]:
            stats["err"] += len(r["absents_src"])
            log_fn(f"    ⚠️ Absent(s) de l'Excel Prod IA : {', '.join(r['absents_src'])}", LOG_WARN)
        if r["absents_dst"]:
            stats["err"] += len(r["absents_dst"])
            log_fn(f"    ⚠️ Absent(s) de l'Excel SharePoint (non ajoutés) : {', '.join(r['absents_dst'])}", LOG_WARN)
    except PermissionError:
        stats["err"] += 1
        log_fn(f"    ⚠️ {os.path.basename(excel_dst)} verrouillé (ouvert dans Excel ?)", LOG_WARN)
    except Exception as e:
        stats["err"] += 1
        log_fn(f"    ❌ Erreur Excel : {e}", LOG_ERR)
    phase("excel", "ok" if stats["err"] == 0 else "err")

    # 4. PPT
    phase("ppt", "run")
    err_avant = stats["err"]
    if ppt_mode is None:
        log_fn("  ⏭️ PPT non traités (aucune option PPT cochée)", LOG_INFO)
    else:
        os.makedirs(ppt_dst, exist_ok=True)
        if ppt_mode == "modifier" and donnees is None:
            try: donnees = lire_lignes(excel_src, ids)
            except Exception as e:
                donnees = {}; log_fn(f"    ❌ Lecture Excel Prod IA impossible : {e}", LOG_ERR)
        for id_err in ids:
            nom = nom_ppt(id_err)
            src, dst = os.path.join(dossier_src, nom), os.path.join(ppt_dst, nom)
            try:
                if ppt_mode == "remplacer" or not os.path.isfile(dst):
                    if not os.path.isfile(src):
                        stats["err"] += 1
                        log_fn(f"    ⚠️ {nom} absent de Prod IA/{commune}", LOG_WARN); continue
                    if ppt_mode == "modifier":
                        log_fn(f"    ℹ️ {nom} absent de SharePoint → copie du nouveau PPT", LOG_INFO)
                    copier_ppt(src, dst); onedrive_pin(dst); stats["ppt"] += 1
                    continue
                d = donnees.get(id_err.upper())
                if d is None:
                    stats["err"] += 1
                    log_fn(f"    ⚠️ {id_err} absent de l'Excel Prod IA → {nom} non modifié", LOG_WARN); continue
                analyse, lien = d[COL_REMARQUE] or None, d[COL_LIEN_SV] or None
                if analyse is None and lien is None:
                    log_fn(f"    ⚠️ {id_err} : « {COL_REMARQUE} » et « {COL_LIEN_SV} » vides → {nom} non modifié", LOG_WARN)
                    continue
                res = modifier_ppt(dst, analyse, lien)
                manque = [z for z, v, ok in (("Analyse", analyse, res["analyse"]),
                                             ("Lien Street View", lien, res["lien"])) if v and not ok]
                vides = [c for c, v in ((COL_REMARQUE, analyse), (COL_LIEN_SV, lien)) if not v]
                if manque:
                    stats["err"] += 1
                    log_fn(f"    ⚠️ {nom} : zone(s) introuvable(s) : {', '.join(manque)}", LOG_WARN)
                if vides:
                    log_fn(f"    ℹ️ {id_err} : « {', '.join(vides)} » vide → zone laissée telle quelle", LOG_INFO)
                if res["analyse"] or res["lien"]:
                    onedrive_pin(dst); stats["ppt"] += 1
            except PermissionError:
                stats["err"] += 1
                log_fn(f"    ⚠️ {nom} verrouillé (ouvert dans PowerPoint ?)", LOG_WARN)
            except Exception as e:
                stats["err"] += 1
                log_fn(f"    ❌ {nom} : {e}", LOG_ERR)
        verbe = "remplacé(s)" if ppt_mode == "remplacer" else "modifié(s)"
        log_fn(f"  🖼️  {stats['ppt']} PPT {verbe} → {dep}/{commune}/{PPT_FOLDER}/", LOG_OK)
    phase("ppt", "ok" if stats["err"] == err_avant else "err")

    # 5. Nettoyage : seulement si tout a réussi ET que le dossier ne contient
    #    pas d'autres ID que ceux saisis (sinon on les perdrait)
    phase("nettoyage", "run")
    autres = [i for i in ids_src if i not in set(ids)]
    if stats["err"]:
        log_fn(f"  ⚠️ {stats['err']} erreur(s) → dossier Prod IA/{commune} conservé", LOG_WARN)
        phase("nettoyage", "err")
    elif not supprimer:
        log_fn(f"  ⏭️ Suppression désactivée (dossier gardé)", LOG_INFO); phase("nettoyage", "ok")
    elif autres:
        log_fn(f"  ⏭️ Prod IA/{commune} conservé : {len(autres)} autre(s) ID non traité(s)", LOG_INFO)
        phase("nettoyage", "ok")
    else:
        try:
            shutil.rmtree(dossier_src)
            log_fn(f"  🗑️  Dossier Prod IA/{commune} supprimé", LOG_WARN)
        except Exception as e:
            log_fn(f"  ⚠️ Impossible de supprimer Prod IA/{commune} : {e}", LOG_WARN)
        phase("nettoyage", "ok")
    return stats
