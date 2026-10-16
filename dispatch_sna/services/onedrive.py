"""Correspondance dossier local synchronisé par OneDrive ↔ adresse SharePoint.

OneDrive enregistre chaque bibliothèque synchronisée dans
HKCU\\Software\\SyncEngines\\Providers\\OneDrive\\<id> (valeurs MountPoint et
UrlNamespace). Ouvrir un classeur par son adresse SharePoint garantit
qu'Excel rejoint la session de co-édition au lieu de travailler sur la copie locale.
"""
from __future__ import annotations

import os
from urllib.parse import quote

_CLE = r"Software\SyncEngines\Providers\OneDrive"


def montages_onedrive() -> list[tuple[str, str]]:
    """[(dossier local, URL SharePoint)] des bibliothèques synchronisées ; [] hors Windows."""
    try:
        import winreg
    except ImportError:
        return []
    res = []
    try:
        racine = winreg.OpenKey(winreg.HKEY_CURRENT_USER, _CLE)
    except OSError:
        return res
    with racine:
        i = 0
        while True:
            try: sous = winreg.EnumKey(racine, i)
            except OSError: break
            i += 1
            try:
                with winreg.OpenKey(racine, sous) as k:
                    mp = winreg.QueryValueEx(k, "MountPoint")[0]
                    url = winreg.QueryValueEx(k, "UrlNamespace")[0]
            except OSError:
                continue
            if mp and url:
                res.append((mp, url))
    return res


def url_sharepoint(chemin: str, montages=None) -> str | None:
    """Adresse SharePoint d'un fichier situé dans un dossier synchronisé, sinon None."""
    montages = montages_onedrive() if montages is None else montages
    absolu = os.path.abspath(chemin)
    p = os.path.normcase(absolu)
    meilleur = None
    for mp, url in montages:
        m = os.path.normcase(os.path.abspath(mp)).rstrip("\\/")
        if (p == m or p.startswith(m + os.sep)) and (meilleur is None or len(m) > len(meilleur[0])):
            meilleur = (m, url)
    if meilleur is None:
        return None
    rel = absolu[len(meilleur[0]):].lstrip("\\/").replace("\\", "/")
    return meilleur[1].rstrip("/") + "/" + quote(rel, safe="/")
