"""Chargement / sauvegarde de config_dispatcher.json (à la racine du projet)."""
import json
import os

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_FILE = os.path.join(RACINE, "config_dispatcher.json")
ATTENTE_FILE = os.path.join(RACINE, "ids_en_attente.json")   # ID non appliqués, à rejouer

DEFAULT_CFG = {
    "root_sp":  r"C:\Users\u269775\Altice Campus SFR\Swap Adresse - Etude Cible\audit_SNA",
    "prod_ia":  r"C:\Users\u269775\Altice Campus SFR\Swap Adresse - Etude Cible\audit_SNA\Prod IA\À déposer",
    "supprimer": True,   # supprimer le dossier commune de Prod IA après rangement
    "supp_adeposer": True,  # supprimer le dossier "À déposer" entier à la fin
    "mode": "commune",      # "commune" (tout ranger) ou "ids" (lignes des ID erreur)
    "ppt_remplacer": False, # mode ids : remplacer les PPT par ceux de Prod IA
    "ppt_modifier": True,   # mode ids : modifier Analyse + Lien Street View
    "colonnes_maj": None,   # mode ids : colonnes Excel à mettre à jour (None = toutes)
    "excel_com": True,      # mode ids : écrire via Excel (co-édition SharePoint) plutôt que réécrire le fichier
    "excel_visible": False, # mode ids : afficher Excel pendant l'écriture (diagnostic)
    "excel_delai_occupe_s": 120,  # mode ids : attente max quand Excel répond « occupé » (par appel)
}


def load_config():
    if os.path.isfile(CONFIG_FILE):
        try:
            d = json.load(open(CONFIG_FILE, encoding="utf-8"))
            for k, v in DEFAULT_CFG.items(): d.setdefault(k, v)
            return d
        except: pass
    return dict(DEFAULT_CFG)

def save_config(d):
    try: json.dump(d, open(CONFIG_FILE, "w", encoding="utf-8"), indent=2, ensure_ascii=False)
    except: pass
