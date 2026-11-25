"""ID erreur non appliqués, conservés d'un lancement à l'autre pour être rejoués.

Fichier JSON : {"74143_1": {"raison": "...", "date": "2026-10-16 18:20"}, ...}
"""
from __future__ import annotations

import json
import re
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


# ── Bilan lisible des échecs ──────────────────────────────────────
# Bornes explicites et non \b : « _ » compte comme une lettre pour \b, qui ne verrait donc pas
# l'ID dans « cas_analyse_73150_88.pptx ».
_RE_ID = re.compile(r"(?<![0-9A-Za-z])(?:2[AB]|\d{2})\d{3}_\d+(?![0-9])")
_RE_COMMUNE = re.compile(r"(?<![0-9A-Za-z])(?:2[AB]|\d{2})\d{3}(?![0-9])")
_RE_DEP = re.compile(r"\bDep(?:2[AB]|\d{2})(?![0-9])")
_RE_LISTE_ID = re.compile(r"\s*\([^)]*«ID»[^)]*\)")    # « (cas_analyse_…, …) » : détail propre à chaque ID


def cause(raison: str) -> str:
    """Raison débarrassée de ce qui est propre à un ID ou une commune, pour regrouper."""
    t = _RE_ID.sub("«ID»", raison)
    t = _RE_DEP.sub("DepXX", _RE_COMMUNE.sub("«commune»", t))
    t = _RE_LISTE_ID.sub("", t)
    return re.sub(r"\b\d+\b", "N", t)


def regrouper(echecs: dict[str, str]) -> list[dict]:
    """[{"cause", "ids", "communes"}] du plus fréquent au plus rare."""
    groupes: dict[str, dict] = {}
    for i, raison in echecs.items():
        g = groupes.setdefault(cause(raison), {"cause": cause(raison), "ids": [], "communes": []})
        g["ids"].append(i)
        commune = i.split("_")[0]
        if commune not in g["communes"]:
            g["communes"].append(commune)
    return sorted(groupes.values(), key=lambda g: -len(g["ids"]))


def ecrire_rapport(dossier: str, echecs: dict[str, str]) -> str:
    """Rapport CSV complet (ID ; commune ; cause ; raison détaillée), lisible dans Excel."""
    import csv
    os.makedirs(dossier, exist_ok=True)
    chemin = os.path.join(dossier, f"echecs_{datetime.now():%Y-%m-%d_%H%M%S}.csv")
    with open(chemin, "w", encoding="utf-8-sig", newline="") as f:      # BOM : accents corrects dans Excel
        w = csv.writer(f, delimiter=";")
        w.writerow(["ID erreur", "Commune", "Cause", "Raison détaillée"])
        for i, raison in sorted(echecs.items()):
            w.writerow([i, i.split("_")[0], cause(raison), raison])
    return chemin


def _ecrire(fichier, d):
    tmp = fichier + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(d, f, indent=2, ensure_ascii=False)
    os.replace(tmp, fichier)
