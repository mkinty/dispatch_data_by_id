"""Arborescence Prod IA / SharePoint : codes commune, copies, détection des colonnes."""
import os
import re
import shutil

from .constantes import DEP_PREFIX, AUDIT_PREFIX, LOG_WARN, LOG_ERR
from .excel_rows import lire_entetes


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

def detecter_colonnes(prod_ia, root_sp, essais=3):
    """En-têtes de l'onglet Audit : 1er Excel lisible dans Prod IA, sinon un
    audit_<commune>.xlsx de SharePoint. Retourne (colonnes, fichier) ou ([], None)."""
    candidats = []
    if os.path.isdir(prod_ia):
        for d in sorted(os.listdir(prod_ia)):
            full = os.path.join(prod_ia, d)
            if os.path.isdir(full) and extraire_commune_nom(d):
                candidats += [os.path.join(full, f) for f in sorted(os.listdir(full))
                              if f.lower().endswith(".xlsx") and not f.startswith("~")]
    if not candidats and os.path.isdir(root_sp):
        for dep in sorted(os.listdir(root_sp)):
            dep_dir = os.path.join(root_sp, dep)
            if not (dep.startswith(DEP_PREFIX) and os.path.isdir(dep_dir)): continue
            for com in sorted(os.listdir(dep_dir)):
                f = os.path.join(dep_dir, com, f"{AUDIT_PREFIX}{com}.xlsx")
                if os.path.isfile(f): candidats.append(f)
                if len(candidats) >= essais: break
            if len(candidats) >= essais: break
    for f in candidats[:essais]:
        try:
            cols = lire_entetes(f)
            if cols: return cols, f
        except Exception:
            continue
    return [], None

def copier_remplacer(src, dst, log_fn):
    """Copie src → dst en remplaçant proprement. Retourne True si OK."""
    try:
        # Écriture via un fichier temporaire pour éviter la corruption
        tmp = dst + ".tmp"
        shutil.copy2(src, tmp)
        if os.path.isfile(dst):
            os.remove(dst)
        shutil.move(tmp, dst)
        onedrive_pin(dst)
        return True
    except PermissionError:
        log_fn(f"    ⚠️ {os.path.basename(dst)} verrouillé (ouvert dans Excel ?)", LOG_WARN)
        return False
    except Exception as e:
        log_fn(f"    ❌ Erreur copie {os.path.basename(dst)} : {e}", LOG_ERR)
        return False



def lister_communes_prod(prod_ia):
    """{code commune → dossier} des sous-dossiers commune de Prod IA."""
    res = {}
    if not os.path.isdir(prod_ia): return res
    for d in os.listdir(prod_ia):
        full = os.path.join(prod_ia, d)
        if os.path.isdir(full):
            com = extraire_commune_nom(d)
            if com: res[com] = full
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
