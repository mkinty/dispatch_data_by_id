"""Orchestration du rangement : commune complète, ou mise à jour ciblée par ID erreur.
Aucune dépendance à l'interface : la progression passe par log_fn / phase_fn."""
import os
import re
import shutil

from .constantes import (AUDIT_PREFIX, PPT_FOLDER, PPT_PREFIX, COL_ID_IMM, COL_REMARQUE, COL_LIEN_SV,
                         LOG_OK, LOG_INFO, LOG_WARN, LOG_ERR)
from .chemins import get_dep, onedrive_pin, copier_remplacer
from .error_ids import nom_ppt
from .ecriture import proprietaire_verrou
from .excel_rows import (cle_de, est_onglet_audit, lire_ids_erreur, lire_lignes, maj_lignes_excel,
                         retirer_filtres_fichier)
from .fichiers_cloud import ErreurOneDrive
from .onedrive import url_sharepoint
from .ppt_editor import modifier_ppt, copier_ppt

SEUIL_REFUS_ONEDRIVE = 3     # refus OneDrive consécutifs avant de reporter le reste de la commune


# ── Traitement d'une commune ──────────────────────────────────────
def ranger_commune(prod_ia, root_sp, commune, dossier_src, supprimer, log_fn, phase_fn=None, publier=None,
                   ppt_coedition=None):
    """Range les fichiers d'une commune depuis Prod IA vers Dep/commune.
    publier : appelé avec chaque fichier écrit (source SharePoint par lien : envoi par l'API).
    ppt_coedition : PowerPoint (ppt_com.PowerPointALaDemande) pour un PPT ouvert par quelqu'un.
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
        qui = proprietaire_verrou(dst, ouvrir=False)
        if qui:
            stats["err"] += 1
            log_fn(f"  🔒 {os.path.basename(dst)} ouvert par {qui} : non remplacé (le remplacer ferait "
                   "perdre son travail) — rejouez quand il sera fermé", LOG_WARN)
            continue
        if copier_remplacer(src, dst, log_fn) and _publier(publier, dst, log_fn, {"mode": "remplacer", "src": src}):
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
        qui = proprietaire_verrou(dst, ouvrir=False)
        if qui:
            if ppt_coedition is not None and _ppt_coedition(ppt_coedition, "remplacer", dst, qui, log_fn, src=src):
                stats["ppt"] += 1
            else:
                stats["err"] += 1
            continue
        if copier_remplacer(src, dst, log_fn) and _publier(publier, dst, log_fn, {"mode": "remplacer", "src": src}):
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


def _publier(publier, chemin, log_fn, op=None) -> bool:
    """Envoie un fichier écrit vers SharePoint (source par lien) ; False + journal si refusé."""
    if publier is None:
        return True
    try:
        message = publier(chemin, op)
        if message:
            log_fn(f"    ✍️ {os.path.basename(chemin)} : {message}", LOG_INFO)
        return True
    except Exception as e:
        log_fn(f"    ❌ {os.path.basename(chemin)} non envoyé sur SharePoint : {e}", LOG_ERR)
        return False


def _ppt_coedition(ppt, mode, dst, qui, log_fn, src=None, analyse=None, lien=None):
    """PPT local ouvert (fichier verrou ~$) : écriture par PowerPoint, en co-édition.
    Retourne le résultat (dict) ou None si impossible (raison journalisée)."""
    nom = os.path.basename(dst)
    cible = url_sharepoint(dst) or dst
    try:
        r = ppt.remplacer(cible, src) if mode == "remplacer" else ppt.modifier(cible, dst, analyse, lien)
    except Exception as e:
        log_fn(f"    🔒 {nom} ouvert par {qui} : non modifié — {e}", LOG_WARN)
        return None
    if r.get("non_verifie"):
        log_fn(f"    ⚠️ {nom} : écrit en co-édition PowerPoint, mais non confirmé à la relecture "
               f"({', '.join(r['non_verifie'])})", LOG_WARN)
        return None
    log_fn(f"    ✍️ {nom} ouvert par {qui} : écrit en co-édition PowerPoint ({r.get('ouverture')})", LOG_INFO)
    return r


def _envoyer(publier, dst, op, nom, log_fn):
    """Source par lien : envoi du PPT écrit (les erreurs remontent à l'appelant)."""
    if publier:
        message = publier(dst, op)
        if message:
            log_fn(f"    ✍️ {nom} : {message}", LOG_INFO)


# ── Traitement ciblé par ID erreur ────────────────────────────────
def _journal_affichage(a, fichier, log_fn):
    """Trace ce que retirer_filtres a changé dans l'onglet Audit d'un fichier."""
    if not a:
        return
    faits = (["filtre retiré"] if a["filtres"] else []) + (["lignes réaffichées"] if a["lignes"] else [])
    if a["colonnes"]:
        faits.append(f"colonnes réaffichées : {', '.join(a['colonnes'])}")
    if faits:
        log_fn(f"  🧹 {fichier} : {' · '.join(faits)}", LOG_INFO)
    if a.get("protege"):
        log_fn(f"    🔒 {fichier} : onglet protégé, laissé masqué : {', '.join(a['protege'])} "
               "(sans effet sur la mise à jour ; ôtez la protection pour les réafficher)", LOG_INFO)
    for e in a.get("erreurs", []):
        log_fn(f"    ⚠️ {fichier} : affichage non rétabli ({e}) — sans effet sur la recherche des ID", LOG_WARN)


def _groupes_ppt(ids, id_erreur, excel_src, excel_dst, commune, dep, nom_dst, stats, echec, log_fn):
    """{ID erreur → [ID saisis]} : le PPT à traiter et les ID imm qu'il concerne.

    ID erreur d'un ID imm : celui déjà lu à l'étape Excel, sinon lu dans l'Excel Prod IA, sinon
    dans l'Excel SharePoint (PPT seuls). Un ID erreur (ancien mode) est son propre PPT."""
    manquants = [i for i in ids if cle_de(i) == COL_ID_IMM and not id_erreur.get(i.upper())]
    for chemin, ou in ((excel_src, f"Prod IA/{commune}"), (excel_dst, f"{dep}/{commune}/{nom_dst}")):
        if not manquants or not chemin or not os.path.isfile(chemin):
            continue
        try:
            id_erreur = {**id_erreur, **lire_ids_erreur(chemin, manquants)}
        except Exception as e:
            log_fn(f"    ⚠️ ID erreur des ID imm non lus dans {ou} : {e}", LOG_WARN)
        manquants = [i for i in manquants if not id_erreur.get(i.upper())]
    groupes = {}
    for i in ids:
        e = i.upper() if cle_de(i) != COL_ID_IMM else id_erreur.get(i.upper(), "")
        if e:
            groupes.setdefault(e, []).append(i)
    if manquants:
        stats["err"] += len(manquants)
        echec(manquants, "ID imm absent de l'Excel (Prod IA et SharePoint) : ID erreur inconnu, PPT introuvable")
        log_fn(f"    ⚠️ {len(manquants)} ID imm introuvable(s) dans l'Excel, PPT non traité(s) : "
               f"{', '.join(manquants[:6])}{' …' if len(manquants) > 6 else ''}", LOG_WARN)
    partages = {e: g for e, g in groupes.items() if len(g) > 1}
    if partages:
        log_fn(f"    ℹ️ {sum(len(g) for g in partages.values())} ID imm partagent {len(partages)} PPT "
               f"(même ID erreur) : chaque PPT est traité une fois", LOG_INFO)
    return groupes


def _cle_ppt(nom: str) -> str:
    """Nom de PPT normalisé pour repérer les quasi-doublons : casse, espaces, « (1) »,
    zéros en tête du numéro (cas_analyse_73150_088 ≈ cas_analyse_73150_88)."""
    base = re.sub(r"\s*\(\d+\)", "", os.path.splitext(nom)[0]).casefold().replace(" ", "")
    return re.sub(r"_0+(\d)", r"_\1", base)


def trouver_ppt(dossier_src, nom, dst=None):
    """Cherche le PPT `nom` dans le dossier commune de Prod IA.

    Retourne (chemin, info) : chemin = fichier à utiliser (nom exact, ou même nom à la casse
    près, éventuellement dans un sous-dossier) ou None ; info = explication pour le journal
    (où il a été trouvé, ou ce que contient le dossier quand il est absent)."""
    exact = os.path.join(dossier_src, nom)
    if os.path.isfile(exact):
        return exact, ""
    ppts, autres, proches = [], 0, []
    for racine, _dossiers, fichiers in os.walk(dossier_src):
        for f in fichiers:
            if f.casefold() == nom.casefold():
                rel = os.path.relpath(os.path.join(racine, f), dossier_src)
                return os.path.join(racine, f), f"trouvé sous le nom « {rel} »"
            if f.lower().endswith(".pptx"):
                ppts.append(f)
                if _cle_ppt(f) == _cle_ppt(nom):
                    proches.append(os.path.relpath(os.path.join(racine, f), dossier_src))
            else:
                autres += 1
    if proches:
        info = f"nom proche présent, non utilisé par prudence : {', '.join(proches[:3])} (renommez-le en {nom})"
    elif not ppts:
        info = f"le dossier ne contient aucun PPT ({autres} autre(s) fichier(s))"
    else:
        apercu = ", ".join(sorted(ppts)[:3]) + (" …" if len(ppts) > 3 else "")
        info = f"le dossier contient {len(ppts)} autre(s) PPT ({apercu})"
    if dst and os.path.isfile(dst):
        info += " ; un PPT existe déjà dans SharePoint pour cet ID (peut-être déjà livré)"
    return None, info


def _autres_ppt(dossier, saisis) -> list[str]:
    """ID des cas_analyse_<ID>.pptx du dossier qui ne font pas partie des ID saisis."""
    res = []
    for f in sorted(os.listdir(dossier)):
        if f.lower().startswith(PPT_PREFIX) and f.lower().endswith(".pptx"):
            i = f[len(PPT_PREFIX):-len(".pptx")].upper()
            if i not in saisis:
                res.append(i)
    return res


def _etape_excel(maj_excel, excel_src, excel_dst, ids, colonnes, commune, dep, nom_src, nom_dst,
                 stats, echec, log_fn):
    """Étape 3 de ranger_ids (au moins une colonne cochée) : écrit les lignes des ID.
    Retourne (donnees, ids_src, id_erreur) ; erreurs et échecs par ID vont dans stats / echec."""
    donnees, ids_src, id_erreur = None, [], {}
    # Source : la lecture ne dépend pas des filtres, ce retrait n'est donc jamais bloquant
    try:
        _journal_affichage(retirer_filtres_fichier(excel_src), f"Prod IA/{commune}/{nom_src}", log_fn)
    except PermissionError:
        log_fn(f"    ℹ️ {nom_src} (Prod IA) ouvert ailleurs : filtres non retirés, lecture faite quand même", LOG_INFO)
    except Exception as e:
        log_fn(f"    ⚠️ Filtres de {nom_src} (Prod IA) non retirés : {e}", LOG_WARN)
    try:
        r = maj_excel(excel_src, excel_dst, ids, colonnes)
        for cote, fichier in (("src", f"Prod IA/{commune}/{nom_src}"), ("dst", f"{dep}/{commune}/{nom_dst}")):
            onglet = (r.get("onglets") or {}).get(cote)
            if onglet and not est_onglet_audit(onglet):
                log_fn(f"  📄 {fichier} : pas d'onglet Audit → 1re feuille « {onglet} » utilisée "
                       f"(elle contient la colonne « ID erreur »)", LOG_INFO)
        _journal_affichage(r.get("affichage"), f"{dep}/{commune}/{nom_dst}", log_fn)
        donnees, ids_src, id_erreur = r["donnees"], r["ids_src"], r.get("id_erreur") or {}
        stats["lignes"] = len(r["maj"]); stats["excel"] = 1 if r["maj"] else 0
        if r.get("ouverture"):
            if r["ouverture"].startswith("écrit"):      # source SharePoint par lien, sans Excel
                log_fn(f"  🔗 {nom_dst} {r['ouverture']}", LOG_INFO)
            else:
                log_fn(f"  🔗 {nom_dst} ouvert dans Excel ({r['ouverture']})", LOG_INFO)
        if r.get("tentatives", 1) > 1:
            log_fn(f"    ℹ️ Excel était occupé : réussi au {r['tentatives']}e essai", LOG_INFO)
        if r.get("relance_excel"):
            log_fn(f"    🔄 Excel ne répondait plus (fenêtre bloquante invisible ?) : instance tuée et "
                   f"relancée, réussi ensuite", LOG_WARN)
        if r.get("bloque"):
            restants = [i for i in ids if i not in r["absents_src"]]
            stats["err"] += 1
            echec(restants, f"{nom_dst} verrouillé par {r['bloque']}")
            log_fn(f"    🔒 {nom_dst} verrouillé par {r['bloque']} → rien n'a été écrit, "
                   f"{len(restants)} ID mis en attente (rejouez une fois le fichier libéré)", LOG_WARN)
        if r["maj"]:
            quoi = "toutes colonnes" if colonnes is None else f"{len(colonnes)} colonne(s)"
            verif = ", valeurs vérifiées" if "non_verifies" in r else ""
            log_fn(f"  📊 {len(r['maj'])} ligne(s) mise(s) à jour ({quoi}{verif}) dans {dep}/{commune}/{nom_dst}", LOG_OK)
        for i, cols in r.get("liens_refuses", {}).items():
            log_fn(f"    🔗 {i} : onglet protégé, lien cliquable non créé (adresse écrite en texte) : {', '.join(cols)}",
                   LOG_INFO)
        for i, cols in r.get("videes", {}).items():
            log_fn(f"    🧽 {i} : ancienne valeur effacée, vide dans Prod IA : {', '.join(cols)}", LOG_INFO)
        if r.get("non_verifies"):
            stats["err"] += len(r["non_verifies"])
            echec(r["non_verifies"], "valeurs non confirmées après enregistrement")
            log_fn(f"    ⚠️ Valeurs non confirmées après enregistrement : {', '.join(r['non_verifies'])}", LOG_WARN)
        if r["col_absentes"]:
            log_fn(f"    ⚠️ Colonne(s) cochée(s) absente(s) d'un des deux Excel, ignorée(s) : {', '.join(r['col_absentes'])}", LOG_WARN)
        if r["absents_src"]:
            stats["err"] += len(r["absents_src"])
            echec(r["absents_src"], "absent de l'Excel Prod IA")
            log_fn(f"    ⚠️ Absent(s) de l'Excel Prod IA : {', '.join(r['absents_src'])}", LOG_WARN)
        if r["absents_dst"]:
            stats["err"] += len(r["absents_dst"])
            echec(r["absents_dst"], "absent de l'Excel SharePoint")
            log_fn(f"    ⚠️ Absent(s) de l'Excel SharePoint (non ajoutés) : {', '.join(r['absents_dst'])}", LOG_WARN)
    except PermissionError:
        stats["err"] += 1
        echec(ids, f"{nom_dst} verrouillé")
        log_fn(f"    ⚠️ {nom_dst} verrouillé (ouvert dans Excel ?)", LOG_WARN)
    except Exception as e:
        stats["err"] += 1
        echec(ids, f"erreur Excel : {e}")
        log_fn(f"    ❌ Erreur Excel : {e}", LOG_ERR)
    return donnees, ids_src, id_erreur


def ranger_ids(root_sp, commune, dossier_src, ids, ppt_mode, supprimer, log_fn, phase_fn=None,
               colonnes=None, maj_excel=None, publier=None, ppt_coedition=None):
    """Met à jour Dep/commune uniquement pour les ID donnés : des ID imm (ID-…), qui désignent
    chacun UNE ligne (ou un ID erreur, ancien mode). Les PPT étant nommés par ID erreur
    (cas_analyse_<ID erreur>.pptx), l'ID erreur de chaque ID imm est lu dans l'Excel ; plusieurs
    ID imm d'un même ID erreur partagent son PPT, traité une seule fois.
    ppt_mode : "remplacer", "modifier" ou None (Excel seul).
    colonnes : colonnes Excel à mettre à jour (None = toutes, [] = aucune).
    maj_excel : fonction (src, dst, ids, colonnes) → résultat ; par défaut réécriture du
                fichier (maj_lignes_excel), ou SessionExcel.maj_lignes pour la co-édition.
    publier : appelé avec chaque PPT écrit (source SharePoint par lien : envoi par l'API) ;
              ses erreurs comptent comme celles de l'écriture.
    ppt_coedition : PowerPoint (ppt_com.PowerPointALaDemande) : un PPT ouvert par quelqu'un est
              écrit en co-édition au lieu d'être laissé de côté.
    Retourne un dict stats (dont "echecs" : {ID → raison}). Si la commune ne peut pas être
    traitée du tout, stats["bloquant"] donne la raison, reprise pour chacun de ses ID."""
    def phase(n, e="run"):
        if phase_fn: phase_fn(n, e)

    def echec(liste, raison):
        for i in liste:
            stats["echecs"][i] = f"{stats['echecs'][i]} ; {raison}" if i in stats["echecs"] else raison

    def bloquant(etape, raison, conseil=""):
        log_fn(f"  ❌ {raison[0].upper()}{raison[1:]}" + (f"\n     → {conseil}" if conseil else ""), LOG_ERR)
        stats["bloquant"] = raison; stats["err"] += 1
        echec(ids, raison + (f" — {conseil}" if conseil else ""))
        phase(etape, "err")
        return stats

    maj_excel = maj_excel or maj_lignes_excel
    dep = get_dep(commune)
    dossier_dst = os.path.join(root_sp, dep, commune)
    ppt_dst = os.path.join(dossier_dst, PPT_FOLDER)
    excel_dst = os.path.join(dossier_dst, f"{AUDIT_PREFIX}{commune}.xlsx")
    nom_dst = os.path.basename(excel_dst)
    stats = {"lignes": 0, "excel": 0, "ppt": 0, "err": 0, "echecs": {}}
    # Aucune colonne cochée = PPT seuls : aucun des deux Excel n'est alors utilisé pour écrire.
    # L'Excel Prod IA ne sert qu'en mode « modifier » (valeurs Remarque / Lien Street View).
    excel_voulu = colonnes is None or bool(colonnes)
    besoin_src = excel_voulu or ppt_mode == "modifier"
    par_imm = any(cle_de(i) == COL_ID_IMM for i in ids)

    # 1. Vérifications
    phase("verif", "run")
    if not dossier_src or not os.path.isdir(dossier_src):
        return bloquant("verif", f"aucun dossier {commune} dans Prod IA",
                        "commune déjà rangée lors d'un précédent dispatch (son dossier est supprimé "
                        "après un rangement réussi), ou dossier pas encore déposé")
    if not os.path.isdir(dossier_dst):
        return bloquant("verif", f"dossier SharePoint {dep}/{commune} introuvable",
                        "vérifiez la racine SharePoint et la synchronisation OneDrive de ce département")
    if excel_voulu and not os.path.isfile(excel_dst):
        return bloquant("verif", f"Excel SharePoint {dep}/{commune}/{nom_dst} introuvable",
                        "pour ne traiter que les PPT, décochez toutes les colonnes Excel")
    phase("verif", "ok")

    # 2. Excel source
    phase("analyse", "run")
    excels = [f for f in os.listdir(dossier_src) if f.lower().endswith(".xlsx") and not f.startswith("~")]
    nom_src = f"{AUDIT_PREFIX}{commune}.xlsx"
    nom_src = nom_src if nom_src in excels else (excels[0] if excels else None)
    excel_src = os.path.join(dossier_src, nom_src) if nom_src else None
    if besoin_src and not nom_src:
        return bloquant("analyse", f"aucun Excel dans Prod IA/{commune}",
                        "nécessaire pour lire Remarque / Lien Street View en mode « Modifier les PPT »"
                        if not excel_voulu else "il contient les valeurs à reporter dans SharePoint")
    if ppt_mode and par_imm and not nom_src and not os.path.isfile(excel_dst):
        return bloquant("analyse", f"aucun Excel dans Prod IA/{commune} ni dans SharePoint",
                        "il faut un Excel pour retrouver l'ID erreur de chaque ID imm (nom du PPT)")
    genre = "ID imm" if par_imm else "ID erreur"
    log_fn(f"  🔎 {len(ids)} {genre}" + (f" · source {nom_src}" if besoin_src else " · PPT seuls"), LOG_INFO)
    phase("analyse", "ok")

    # 3. Excel : remplacement des seules lignes des ID
    phase("excel", "run")
    donnees, ids_src, id_erreur = None, [], {}
    if not excel_voulu:
        log_fn("  ⏭️ Excel non modifié (aucune colonne cochée) : seuls les PPT sont traités", LOG_INFO)
        phase("excel", "ok")
    else:
        donnees, ids_src, id_erreur = _etape_excel(maj_excel, excel_src, excel_dst, ids, colonnes, commune,
                                                   dep, nom_src, nom_dst, stats, echec, log_fn)
        phase("excel", "ok" if stats["err"] == 0 else "err")

    # 4. PPT
    phase("ppt", "run")
    err_avant = stats["err"]
    groupes = {}                    # {ID erreur → ID saisis} : PPT à traiter
    if ppt_mode is None:
        log_fn("  ⏭️ PPT non traités (aucune option PPT cochée)", LOG_INFO)
    else:
        os.makedirs(ppt_dst, exist_ok=True)
        if ppt_mode == "modifier" and donnees is None:
            try: donnees = lire_lignes(excel_src, ids)
            except Exception as e:
                donnees = {}; log_fn(f"    ❌ Lecture Excel Prod IA impossible : {e}", LOG_ERR)
        groupes = _groupes_ppt(ids, id_erreur, excel_src, excel_dst, commune, dep, nom_dst, stats, echec, log_fn)
        refus_onedrive = 0          # refus OneDrive consécutifs (malgré les réessais) : disjoncteur
        a_liberer = []              # fichiers à libérer de l'espace local APRÈS la boucle
        liste_groupes = list(groupes.items())
        for n, (id_err, ids_ppt) in enumerate(liste_groupes):
            if refus_onedrive >= SEUIL_REFUS_ONEDRIVE:
                reste = [i for _e, g in liste_groupes[n:] for i in g]
                stats["err"] += len(reste); stats["onedrive_ko"] = True
                echec(reste, "reporté : OneDrive refuse les opérations sur ce dossier (voir journal)")
                log_fn(f"    ⏸️ OneDrive a refusé {refus_onedrive} PPT de suite : les {len(reste)} suivants "
                       "sont mis en attente sans être tentés (vérifiez que OneDrive est lancé, connecté, "
                       "non suspendu, avec de l'espace disque)", LOG_WARN)
                break
            nom = nom_ppt(id_err)
            src, dst = os.path.join(dossier_src, nom), os.path.join(ppt_dst, nom)
            qui = proprietaire_verrou(dst, ouvrir=False)
            if qui and ppt_coedition is None:
                stats["err"] += 1; echec(ids_ppt, f"PPT {nom} ouvert par {qui} : non modifié, rejouez quand il sera fermé")
                log_fn(f"    🔒 {nom} ouvert par {qui} : non modifié (l'écraser ferait perdre son travail)", LOG_WARN)
                continue
            try:
                if ppt_mode == "remplacer" or not os.path.isfile(dst):
                    src, info = trouver_ppt(dossier_src, nom, dst)
                    if src is None:
                        stats["err"] += 1; echec(ids_ppt, f"PPT absent de Prod IA — {info}")
                        log_fn(f"    ⚠️ {nom} absent de Prod IA/{commune} : {info}", LOG_WARN); continue
                    if info:
                        log_fn(f"    ℹ️ {nom} : {info}", LOG_INFO)
                    if ppt_mode == "modifier" and not os.path.isfile(dst):
                        log_fn(f"    ℹ️ {nom} absent de SharePoint → copie du nouveau PPT", LOG_INFO)
                    if qui:                 # ouvert : remplacement des diapositives en co-édition
                        if not _ppt_coedition(ppt_coedition, "remplacer", dst, qui, log_fn, src=src):
                            stats["err"] += 1; echec(ids_ppt, f"PPT {nom} ouvert par {qui} : non remplacé")
                            continue
                        stats["ppt"] += 1; refus_onedrive = 0
                        continue
                    copier_ppt(src, dst)
                    _envoyer(publier, dst, {"mode": "remplacer", "src": src}, nom, log_fn)
                    a_liberer.append(dst); stats["ppt"] += 1
                    refus_onedrive = 0
                    continue
                d = next((donnees[i] for i in ids_ppt if i in donnees), None)
                if d is None:
                    stats["err"] += 1; echec(ids_ppt, "PPT non modifié (ID absent de l'Excel Prod IA)")
                    log_fn(f"    ⚠️ {', '.join(ids_ppt)} absent(s) de l'Excel Prod IA → {nom} non modifié", LOG_WARN); continue
                # Vide + colonne cochée : la zone est vidée, comme dans l'Excel. Vide + non cochée : on n'y touche pas.
                analyse, lien = (d[col] or ("" if colonnes is None or col in colonnes else None)
                                 for col in (COL_REMARQUE, COL_LIEN_SV))
                if analyse is None and lien is None:
                    log_fn(f"    ⚠️ {id_err} : « {COL_REMARQUE} » et « {COL_LIEN_SV} » vides et non cochées → {nom} non modifié", LOG_WARN)
                    continue
                if qui:                     # ouvert : seules les zones, en co-édition PowerPoint
                    res = _ppt_coedition(ppt_coedition, "modifier", dst, qui, log_fn, analyse=analyse, lien=lien)
                    if res is None:
                        stats["err"] += 1; echec(ids_ppt, f"PPT {nom} ouvert par {qui} : non modifié")
                        continue
                else:
                    res = modifier_ppt(dst, analyse, lien)
                manque = [z for z, v, ok in (("Analyse", analyse, res["analyse"]),
                                             ("Lien Street View", lien, res["lien"])) if v is not None and not ok]
                vides = [c for c, v in ((COL_REMARQUE, analyse), (COL_LIEN_SV, lien)) if v is None]
                videes = [c for c, v in ((COL_REMARQUE, analyse), (COL_LIEN_SV, lien)) if v == ""]
                if manque:
                    stats["err"] += 1; echec(ids_ppt, f"PPT : zone(s) introuvable(s) {', '.join(manque)}")
                    log_fn(f"    ⚠️ {nom} : zone(s) introuvable(s) : {', '.join(manque)}", LOG_WARN)
                if vides:
                    log_fn(f"    ℹ️ {id_err} : « {', '.join(vides)} » vide et non cochée → zone du PPT laissée telle quelle", LOG_INFO)
                if videes:
                    log_fn(f"    🧽 {nom} : zone vidée, « {', '.join(videes)} » vide dans Prod IA", LOG_INFO)
                refus_onedrive = 0
                if res["analyse"] or res["lien"]:
                    if not qui:
                        _envoyer(publier, dst, {"mode": "modifier", "analyse": analyse, "lien": lien}, nom, log_fn)
                        a_liberer.append(dst)
                    stats["ppt"] += 1
            except PermissionError as e:
                detail = str(e) if publier and str(e) else "ouvert dans PowerPoint ?"
                stats["err"] += 1; echec(ids_ppt, f"PPT verrouillé ({detail})")
                log_fn(f"    ⚠️ {nom} verrouillé ({detail})", LOG_WARN)
            except ErreurOneDrive as e:
                refus_onedrive += 1
                stats["err"] += 1; echec(ids_ppt, f"PPT : {e}")
                log_fn(f"    ❌ {nom} : {e}", LOG_ERR)
            except Exception as e:
                stats["err"] += 1; echec(ids_ppt, f"PPT : {e}")
                log_fn(f"    ❌ {nom} : {e}", LOG_ERR)
        # Libération de l'espace local seulement une fois tous les PPT écrits : désépingler un
        # fichier encore en cours d'envoi le rend « en ligne uniquement » au pire moment.
        for f in a_liberer:
            onedrive_pin(f)
        verbe = "remplacé(s)" if ppt_mode == "remplacer" else "modifié(s)"
        log_fn(f"  🖼️  {stats['ppt']} PPT {verbe} → {dep}/{commune}/{PPT_FOLDER}/", LOG_OK)
    phase("ppt", "ok" if stats["err"] == err_avant else "err")

    # 5. Nettoyage : seulement si tout a réussi ET que le dossier ne contient pas de travail
    #    d'autres ID que ceux saisis (sinon on le perdrait).
    #    - Excel mis à jour : les autres ID sont ceux des lignes de l'Excel Prod IA.
    #    - PPT seuls : l'Excel n'a servi à rien (ou seulement de lecture) et n'est pas un livrable
    #      en attente ; seuls comptent les PPT d'autres ID présents dans le dossier.
    phase("nettoyage", "run")
    saisis = {i.upper() for i in ids}
    if excel_voulu:
        autres, quoi = [i for i in ids_src if i not in saisis], "ligne(s) Excel d'autres ID"
    else:   # PPT seuls : les PPT sont nommés par ID erreur — ceux des ID saisis (groupes)
        autres, quoi = _autres_ppt(dossier_src, {e.upper() for e in groupes}), "PPT d'autres ID"
    if stats["err"]:
        log_fn(f"  ⚠️ {stats['err']} erreur(s) → dossier Prod IA/{commune} conservé", LOG_WARN)
        phase("nettoyage", "err")
    elif not supprimer:
        log_fn(f"  ⏭️ Suppression désactivée (dossier gardé)", LOG_INFO); phase("nettoyage", "ok")
    elif autres:
        apercu = ", ".join(autres[:5]) + (" …" if len(autres) > 5 else "")
        log_fn(f"  ⏭️ Prod IA/{commune} conservé : {len(autres)} {quoi} non traité(s) ({apercu})", LOG_INFO)
        phase("nettoyage", "ok")
    else:
        try:
            shutil.rmtree(dossier_src)
            log_fn(f"  🗑️  Dossier Prod IA/{commune} supprimé", LOG_WARN)
        except Exception as e:
            log_fn(f"  ⚠️ Impossible de supprimer Prod IA/{commune} : {e}", LOG_WARN)
        phase("nettoyage", "ok")
    return stats
