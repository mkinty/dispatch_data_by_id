"""Liens SharePoint et connexion par le navigateur, avec l'identité de l'utilisateur.

Repris d'audit_validation_finale (avf/services/sharepoint.py : analyser_lien,
connecter_navigateur, registre des sessions), même processus :

1. l'utilisateur colle le **lien du dossier** (Racine SharePoint et/ou Dossier à déposer) ;
2. « 🔗 Se connecter » : d'abord **sans fenêtre** (le profil dédié a peut-être encore une
   session valide), sinon une fenêtre **Microsoft Edge** (Chrome à défaut) s'ouvre sur le
   dossier ; l'utilisateur s'identifie normalement (compte professionnel, MFA) ; dès que la
   session ouvre le dossier par l'API, la fenêtre se ferme ;
3. l'outil interroge ensuite l'**API SharePoint** avec cette session (services.sharepoint.rest).

L'outil n'enregistre ni mot de passe ni cookie : la session reste dans le profil dédié du
navigateur (%LOCALAPPDATA%\\DispatcherSNA\\navigateur), chiffré par Windows.
"""
from __future__ import annotations

import re
import shutil
import threading
import time
from dataclasses import dataclass
from urllib.parse import parse_qs, quote, unquote, urlsplit

from . import rest
from .navigateur import NavigateurFerme, NavigateurIndisponible, NavigateurPilote, cookies_pour, trouver_navigateur

_RACINE_WEBDAV = re.compile(r"^\\\\(?P<hote>[^\\@]+)@SSL(?:@\d+)?\\DavWWWRoot(?P<chemin>\\.*)?$", re.I)
_PARTAGE_RESOLUBLE = re.compile(r"^/:[a-z]:/r(/.*)$", re.I)       # /:f:/r/sites/… : chemin lisible
_PARTAGE_OPAQUE = re.compile(r"^/:[a-z]:/[a-z]/", re.I)           # /:f:/s/…/EaBc… : sans chemin


class LienInvalide(ValueError):
    """Le texte saisi n'est pas un lien SharePoint de dossier exploitable."""


class AccesImpossible(RuntimeError):
    """Connexion impossible : délai dépassé, dossier introuvable, droits insuffisants."""


@dataclass(frozen=True)
class LienSharePoint:
    """Dossier SharePoint : hôte et chemin serveur décodé (``/sites/X/Shared Documents/Audit``)."""
    hote: str
    chemin: str

    @property
    def url(self) -> str:
        return f"https://{self.hote}{quote(self.chemin, safe='/')}"


def est_lien(texte: str) -> bool:
    """Vrai si le texte saisi est un lien SharePoint (https:// ou chemin WebDAV), pas un dossier local."""
    t = (texte or "").strip().strip('"')
    return bool(re.match(r"^https?://", t, re.I) or _RACINE_WEBDAV.match(t))


def analyser_lien(texte: str) -> LienSharePoint:
    """Lien collé → dossier SharePoint. Acceptés, comme dans audit_validation_finale :

    - l'adresse du dossier : ``https://x.sharepoint.com/sites/S/Shared%20Documents/Audit`` ;
    - la barre d'adresse sur la bibliothèque : ``…/Forms/AllItems.aspx?id=%2Fsites%2FS%2F…`` ;
    - un lien de partage au chemin lisible : ``https://x.sharepoint.com/:f:/r/sites/S/…?csf=1`` ;
    - un chemin WebDAV ``\\\\x.sharepoint.com@SSL\\DavWWWRoot\\sites\\S\\…``.

    Un lien de partage opaque (``/:f:/s/S/EaBc…``) ne contient pas le chemin : refusé, avec la
    marche à suivre.
    """
    brut = (texte or "").strip().strip('"').strip()
    if not brut:
        raise LienInvalide("aucun lien saisi")
    webdav = _RACINE_WEBDAV.match(brut)
    if webdav:
        chemin = (webdav.group("chemin") or "").replace("\\", "/")
        return LienSharePoint(webdav.group("hote").lower(), _nettoyer(unquote(chemin)))
    if not re.match(r"^https?://", brut, re.I):
        raise LienInvalide("le lien doit commencer par https:// (adresse SharePoint du dossier)")
    morceaux = urlsplit(brut)
    hote = (morceaux.hostname or "").lower()
    if not hote:
        raise LienInvalide("adresse sans nom de serveur")
    chemin = morceaux.path
    identifiant = parse_qs(morceaux.query).get("id") or parse_qs(morceaux.query).get("RootFolder")
    if identifiant:
        return LienSharePoint(hote, _nettoyer(unquote(identifiant[0])))
    if _PARTAGE_OPAQUE.match(chemin) and not _PARTAGE_RESOLUBLE.match(chemin):
        raise LienInvalide("lien de partage sans chemin de dossier — ouvrez le dossier dans SharePoint et "
                           "copiez l'adresse de la barre du navigateur")
    partage = _PARTAGE_RESOLUBLE.match(chemin)
    if partage:
        chemin = partage.group(1)
    chemin = unquote(chemin)
    if "/Forms/" in chemin:
        chemin = chemin.split("/Forms/")[0]
    elif chemin.lower().endswith(".aspx"):
        raise LienInvalide("ce lien désigne une page, pas un dossier de documents")
    chemin = _nettoyer(chemin)
    if chemin.count("/") < 2:
        raise LienInvalide("le lien doit désigner un dossier d'une bibliothèque (…/sites/<site>/<bibliothèque>/…)")
    return LienSharePoint(hote, chemin)


def _nettoyer(chemin: str) -> str:
    return re.sub(r"/{2,}", "/", "/" + chemin.strip().strip("/"))


# ── Sessions établies (une par hôte SharePoint) ───────────────────
_CLIENTS: dict[str, rest.ClientSharePoint] = {}
_VERROU = threading.Lock()


def client_pour(lien: str) -> rest.ClientSharePoint:
    hote = analyser_lien(lien).hote
    with _VERROU:
        client = _CLIENTS.get(hote)
    if client is None:
        raise rest.AuthentificationRequise("pas de session SharePoint : cliquez sur « 🔗 Se connecter »")
    return client


def est_connecte(lien: str) -> bool:
    try:
        client_pour(lien)
        return True
    except (rest.AuthentificationRequise, LienInvalide):
        return False


def enregistrer_client(client: rest.ClientSharePoint) -> None:
    with _VERROU:
        _CLIENTS[client.hote] = client


def oublier_connexion() -> None:
    """Oublie les sessions de l'outil (mémoire et profil du navigateur dédié)."""
    from .navigateur import dossier_profil
    with _VERROU:
        _CLIENTS.clear()
    shutil.rmtree(dossier_profil(), ignore_errors=True)


def navigateur_disponible() -> bool:
    return trouver_navigateur() is not None


def connecter_navigateur(texte, on_etape=None, interactif=True, delai_s=300, delai_silencieux_s=15,
                         fabrique=NavigateurPilote, client_factory=rest.ClientSharePoint) -> rest.ClientSharePoint:
    """Connexion au dossier SharePoint avec l'identité de l'utilisateur, via le navigateur.

    Raises:
        LienInvalide, NavigateurIndisponible, NavigateurFerme, AccesImpossible,
        rest.AuthentificationRequise (non interactif et aucune session valide).
    """
    lien = analyser_lien(texte)
    with _VERROU:
        deja = _CLIENTS.get(lien.hote)
    if deja is not None:                     # même hôte que l'autre source : même session
        try:
            deja.verifier(lien.chemin)
            return deja
        except rest.AuthentificationRequise:
            pass
        except FileNotFoundError as exc:
            raise AccesImpossible(f"connecté à SharePoint, mais le dossier est introuvable : {lien.chemin}") from exc

    def dire(message):
        if on_etape is not None:
            on_etape(message)

    def essayer(nav):
        client = client_factory(lien.hote, cookies_pour(nav.cookies(), lien.hote))
        try:
            client.verifier(lien.chemin)
        except rest.AuthentificationRequise:
            return None
        except FileNotFoundError as exc:
            raise AccesImpossible(f"connecté à SharePoint, mais le dossier est introuvable : {lien.chemin} "
                                  "— vérifiez le lien (il doit ouvrir le dossier dans le navigateur)") from exc
        return client

    def attendre(nav, limite_s, pas_s):
        fin = time.monotonic() + limite_s
        while time.monotonic() < fin:
            if not nav.ouvert():
                raise NavigateurFerme("fenêtre du navigateur fermée avant la fin de la connexion")
            try:
                client = essayer(nav)
            except NavigateurFerme:
                raise
            except (OSError, RuntimeError) as exc:     # navigateur pas encore prêt, réseau…
                if isinstance(exc, (AccesImpossible, NavigateurIndisponible)):
                    raise
                client = None
            if client is not None:
                return client
            time.sleep(pas_s)
        return None

    dire("vérification d'une session SharePoint existante…")
    with fabrique(lien.url, visible=False) as nav:
        client = attendre(nav, delai_silencieux_s, 1.0)
    if client is None:
        if not interactif:
            raise rest.AuthentificationRequise("connexion SharePoint requise : cliquez sur « 🔗 Se connecter »")
        dire("🔑 identifiez-vous dans la fenêtre du navigateur (compte professionnel) — cochez "
             "« Rester connecté » pour ne plus avoir à le faire")
        with fabrique(lien.url, visible=True) as nav:
            client = attendre(nav, delai_s, 2.0)
            if client is None:
                raise AccesImpossible(f"connexion SharePoint non terminée en {delai_s / 60:.0f} min")
            dire("✅ identification réussie — fermeture du navigateur")
    enregistrer_client(client)
    return client
