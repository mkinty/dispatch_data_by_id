"""ID erreur non appliqués, conservés d'un lancement à l'autre pour être rejoués.

Fichier JSON : {"74143_1": {"raison": "...", "date": "2026-10-16 18:20"}, ...}
"""
from __future__ import annotations

import json
import os
from datetime import datetime


def charger(fichier: str) -> dict[str, dict]:
    try:
        with open(fichier, encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


def mettre_a_jour(fichier: str, echecs: dict[str, str], reussis) -> dict[str, dict]:
    """Retire les ID réussis, ajoute / remplace les échecs (avec leur raison). Retourne la liste à jour."""
    d = charger(fichier)
    for i in reussis:
        d.pop(i, None)
    date = datetime.now().strftime("%Y-%m-%d %H:%M")
    for i, raison in echecs.items():
        d[i] = {"raison": raison, "date": date}
    _ecrire(fichier, d)
    return d


def vider(fichier: str):
    _ecrire(fichier, {})


def _ecrire(fichier, d):
    tmp = fichier + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, indent=2, ensure_ascii=False)
    os.replace(tmp, fichier)
