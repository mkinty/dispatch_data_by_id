"""Chargement / sauvegarde de config_dispatcher.json (à la racine du projet)."""
import json
import os

CONFIG_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "config_dispatcher.json")

DEFAULT_CFG = {
    "root_sp":  r"C:\Users\u269775\Altice Campus SFR\Swap Adresse - Etude Cible\audit_SNA",
    "prod_ia":  r"C:\Users\u269775\Altice Campus SFR\Swap Adresse - Etude Cible\audit_SNA\Prod IA\À déposer",
    "supprimer": True,   # supprimer le dossier commune de Prod IA après rangement
    "supp_adeposer": True,  # supprimer le dossier "À déposer" entier à la fin
    "mode": "commune",      # "commune" (tout ranger) ou "ids" (lignes des ID erreur)
    "ppt_remplacer": False, # mode ids : remplacer les PPT par ceux de Prod IA
    "ppt_modifier": True,   # mode ids : modifier Analyse + Lien Street View
    "colonnes_maj": None,   # mode ids : colonnes Excel à mettre à jour (None = toutes)
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
