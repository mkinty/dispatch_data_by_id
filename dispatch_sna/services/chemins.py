"""Arborescence Prod IA / SharePoint : codes commune, copies, détection des colonnes."""
import os
import re
import shutil

from .constantes import DEP_PREFIX, AUDIT_PREFIX, COL_ID_ERR, LOG_WARN, LOG_ERR
from .excel_rows import lire_entetes
from .fichiers_cloud import copier


def extraire_commune_nom(nom):
    """Code commune (5 chiffres ou 2A/2B+3) depuis un nom de dossier/fichier."""
    m = re.search(r'(2[ABab]\d{3}|\d{5})', nom)
    return m.group(1).upper() if m else None

def get_dep(commune):
    return f"{DEP_PREFIX}{str(commune).strip().upper()[:2]}"

def onedrive_pin(path):
    """Force OneDrive à garder le fichier en local (libère l'espace après)."""
    try: os.system(f'attrib +U -P "{path}"')
    except: pass

# ── Détection des colonnes ────────────────────────────────────────
# Attributs Windows d'un fichier OneDrive « disponible en ligne uniquement »
_EN_LIGNE_SEULEMENT = 0x00400000 | 0x00040000 | 0x00001000   # RECALL_ON_DATA_ACCESS | RECALL_ON_OPEN | OFFLINE


def _est_excel(nom):
    return nom.lower().endswith(".xlsx") and not nom.startswith(("~$", "."))


def _en_ligne_seulement(f) -> bool:
    try:
        return bool(getattr(os.stat(f), "st_file_attributes", 0) & _EN_LIGNE_SEULEMENT)
    except OSError:
        return False


def _excels_prod(prod_ia, profondeur=3):
    """Excel de Prod IA jusqu'à `profondeur` niveaux (À déposer/<commune>/ ou À déposer/LOTxx/<commune>/)."""
    res = []
    for racine, dossiers, fichiers in os.walk(prod_ia):
        niveau = os.path.relpath(racine, prod_ia).count(os.sep) + (racine != prod_ia)
        dossiers[:] = sorted(dossiers) if niveau < profondeur else []
        if racine != prod_ia:
            res += [os.path.join(racine, f) for f in sorted(fichiers) if _est_excel(f)]
    return res


def _excels_sharepoint(root_sp, limite=40):
    """audit_<commune>.xlsx de Dep*/<commune>/ (à défaut, tout audit*.xlsx du dossier commune)."""
    res = []
    for dep in sorted(os.listdir(root_sp)):
        dep_dir = os.path.join(root_sp, dep)
        if not (dep.casefold().startswith(DEP_PREFIX.casefold()) and os.path.isdir(dep_dir)):
            continue
        for com in sorted(os.listdir(dep_dir)):
            com_dir = os.path.join(dep_dir, com)
            if not os.path.isdir(com_dir):
                continue
            noms = [f for f in sorted(os.listdir(com_dir)) if _est_excel(f)]
            attendu = f"{AUDIT_PREFIX}{com}.xlsx".casefold()
            exact = [f for f in noms if f.casefold() == attendu]
            res += [os.path.join(com_dir, f) for f in (exact or [f for f in noms
                                                                if f.casefold().startswith(AUDIT_PREFIX)])]
            if len(res) >= limite:
                return res
    return res


def _raison(e) -> str:
    """Explication lisible de l'échec de lecture d'un Excel."""
    from zipfile import BadZipFile
    if isinstance(e, PermissionError):
        return "accès refusé (fichier verrouillé ou droits insuffisants)"
    if isinstance(e, BadZipFile) or type(e).__name__ == "InvalidFileException":
        return "n'est pas un .xlsx lisible (fichier abîmé, protégé par mot de passe, ou téléchargement OneDrive incomplet)"
    if isinstance(e, OSError):
        return f"OneDrive / disque n'a pas fourni le fichier : {getattr(e, 'strerror', None) or e}"
    return str(e)


def detecter_colonnes(prod_ia, root_sp, essais=12, delai_s=45):
    """En-têtes de l'onglet Audit d'un Excel de Prod IA, sinon de SharePoint.

    Les fichiers déjà présents sur le disque sont essayés avant ceux « en ligne
    uniquement » (dont la lecture force un téléchargement OneDrive). Un fichier
    n'est retenu que si son onglet Audit a une colonne « ID erreur ».
    Retourne (colonnes, fichier, rapport) ; rapport = lignes « où on a cherché et
    pourquoi chaque fichier a été écarté », pour un message d'erreur utile.
    """
    import time
    rapport = []
    for nom, dossier, lister in (("Prod IA", prod_ia, _excels_prod), ("SharePoint", root_sp, _excels_sharepoint)):
        if not dossier or not os.path.isdir(dossier):
            rapport.append(f"{nom} : dossier introuvable ({dossier or 'non renseigné'})")
            continue
        try:
            candidats = lister(dossier)
        except OSError as e:
            rapport.append(f"{nom} : impossible de parcourir {dossier} ({_raison(e)})"); continue
        if not candidats:
            ou = "dans ses sous-dossiers" if nom == "Prod IA" else f"dans {DEP_PREFIX}XX/<commune>/"
            rapport.append(f"{nom} : aucun fichier .xlsx {ou} ({dossier})")
            continue
        candidats.sort(key=_en_ligne_seulement)          # stable : présents sur le disque d'abord
        rapport.append(f"{nom} : {len(candidats)} fichier(s) Excel trouvé(s)")
        limite = time.time() + delai_s
        for f in candidats[:essais]:
            if time.time() > limite:
                rapport.append(f"  … arrêt après {delai_s} s (téléchargements OneDrive trop longs ?)"); break
            try:
                cols = lire_entetes(f)
            except Exception as e:
                rapport.append(f"  ✗ {os.path.basename(f)} : {_raison(e)}"); continue
            if COL_ID_ERR not in cols:
                rapport.append(f"  ✗ {os.path.basename(f)} : pas de colonne « {COL_ID_ERR} » en ligne 1 de l'onglet Audit")
                continue
            return cols, f, rapport
    return [], None, rapport

def copier_remplacer(src, dst, log_fn):
    """Copie src → dst en remplaçant proprement. Retourne True si OK."""
    try:
        copier(src, dst)                  # fichier temporaire + remplacement robuste OneDrive
        onedrive_pin(dst)
        return True
    except PermissionError:
        log_fn(f"    ⚠️ {os.path.basename(dst)} verrouillé (ouvert dans Excel ?)", LOG_WARN)
        return False
    except Exception as e:
        log_fn(f"    ❌ Erreur copie {os.path.basename(dst)} : {e}", LOG_ERR)
        return False



def lister_communes_prod(prod_ia):
    """{code commune → dossier} des sous-dossiers commune de Prod IA.

    Si plusieurs dossiers portent le même code (« 73283 », « 73283 - Copie »…), le dossier
    nommé exactement comme la commune l'emporte, sinon le premier par ordre alphabétique —
    avant, c'était le dernier rencontré, donc parfois la copie au lieu du vrai dossier."""
    res = {}
    if not os.path.isdir(prod_ia): return res
    for d in sorted(os.listdir(prod_ia)):
        full = os.path.join(prod_ia, d)
        if os.path.isdir(full):
            com = extraire_commune_nom(d)
            if com and (com not in res or d.upper() == com):
                res[com] = full
    return res

def supprimer_si_vide(prod_ia, log_fn):
    """Supprime le dossier "À déposer" entier s'il ne contient plus de dossier commune."""
    try:
        restant = [d for d in os.listdir(prod_ia)
                   if os.path.isdir(os.path.join(prod_ia, d))]
        if not restant:
            # dossier vide → on le supprime entièrement
            base = os.path.basename(prod_ia.rstrip("\\/"))
            shutil.rmtree(prod_ia)
            log_fn(f"\n🗑️  Dossier '{base}' supprimé (vide) — prêt pour un nouveau dépôt", LOG_WARN)
        else:
            log_fn(f"\n⚠️ '{os.path.basename(prod_ia)}' conservé ({len(restant)} commune(s) non rangée(s) : {', '.join(restant)})", LOG_WARN)
    except Exception as e:
        log_fn(f"\n⚠️ Impossible de supprimer le dossier : {e}", LOG_WARN)
