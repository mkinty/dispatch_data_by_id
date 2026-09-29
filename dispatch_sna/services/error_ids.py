"""ID erreur saisis par l'utilisateur : extraction et regroupement par commune."""
from __future__ import annotations

import re

from .constantes import PPT_PREFIX

# <code INSEE>_<numéro>, ex : 74143_1, 2A004_12 (même format que export_lot_by_id)
ERROR_ID_PATTERN = re.compile(r'\b((?:2[AB]|\d{2})\d{3})_(\d+)\b')


def extraire_ids(texte: str) -> dict[str, list[str]]:
    """{"74143": ["74143_1", "74143_2"], ...} — ordre conservé, doublons supprimés."""
    groupes: dict[str, list[str]] = {}
    for insee, num in ERROR_ID_PATTERN.findall(texte.upper()):
        ids = groupes.setdefault(insee, [])
        id_err = f"{insee}_{num}"
        if id_err not in ids:
            ids.append(id_err)
    return groupes


def nom_ppt(id_err: str) -> str:
    return f"{PPT_PREFIX}{id_err}.pptx"
