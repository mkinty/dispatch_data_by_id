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


# ── ID imm ────────────────────────────────────────────────────────
# « ID- » + code INSEE (5 caractères, Corse 2A/2B comprise) + numéro : ID-3806800000IW.
# Même motif que audit_validation_finale (avf/services/detection.py).
ID_IMM_PATTERN = re.compile(r"(?<![0-9A-Z])ID-((?:2[AB]\d{3})|\d{5})[0-9A-Z]+(?![0-9A-Z])")


def extraire_ids_imm(texte: str) -> dict[str, list[str]]:
    """{"38068": ["ID-3806800000IW", …], …} — majuscules, ordre conservé, doublons supprimés.
    Un en-tête « ID imm » collé avec les valeurs est ignoré (il ne correspond pas au motif)."""
    groupes: dict[str, list[str]] = {}
    for m in ID_IMM_PATTERN.finditer((texte or "").upper()):
        ids = groupes.setdefault(m.group(1), [])
        if m.group(0) not in ids:
            ids.append(m.group(0))
    return groupes


def insee_de_id_imm(id_imm: str) -> str:
    """Code commune d'un ID imm : les 5 caractères après « ID- » (ID-2B030000006D → 2B030)."""
    m = ID_IMM_PATTERN.fullmatch("".join(str(id_imm or "").split()).upper())
    return m.group(1) if m else ""
