"""(Repris d'audit_validation_finale, avf/services/sharepoint_rest.py, complété pour le
Dispatcher : dépôt de fichiers neufs, création de dossier, mise à la corbeille, téléchargement
d'un dossier entier.)

Accès au dossier SharePoint par son API (REST), avec la session ouverte dans le navigateur.

Après l'identification dans le navigateur (``services.navigateur``), les cookies
de session SharePoint suffisent pour appeler l'API du site, exactement comme le
fait la page SharePoint elle-même :

    GET {site}/_api/web/GetFolderByServerRelativePath(decodedurl='…')/Folders
    GET {site}/_api/web/GetFolderByServerRelativePath(decodedurl='…')/Files
    GET {site}/_api/web/GetFileByServerRelativePath(decodedurl='…')/$value

C'est la version **serveur** des fichiers qui est lue : ni OneDrive, ni WebDAV,
ni Excel.

Écriture (mise à jour, cf. ``services.sharepoint.sources``) :

    POST {site}/_api/contextinfo                              → jeton de formulaire
    POST {site}/_api/web/GetFileByServerRelativePath(…)/$value
         X-HTTP-Method: PUT, If-Match: <ETag lu au téléchargement>

``If-Match`` garantit qu'on ne remplace jamais une version que l'on n'a pas
lue : si quelqu'un a enregistré le fichier entre-temps, SharePoint refuse
(:class:`VersionModifiee`) et la mise à jour repart de la nouvelle version. Un
fichier extrait ou verrouillé (ouvert en co-édition) est signalé
(:class:`FichierVerrouille`) plutôt qu'écrasé. SharePoint garde l'historique
des versions : chaque mise à jour en crée une nouvelle.
"""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import json
import re
import threading
import time
import urllib.error
import urllib.request
from urllib.parse import quote, unquote, urlsplit

from ..constantes import AUDIT_PREFIX as PREFIXE_AUDIT, DEP_PREFIX as PREFIXE_DEP

EXTENSION = ".xlsx"

_MOTIF_DATE_NOM = re.compile(r"\d{4}-\d{2}-\d{2}|\d{2}[/-]\d{2}[/-]\d{4}")
_SITE = re.compile(r"^/(sites|teams|personal|portals)/[^/]+", re.I)


class AuthentificationRequise(PermissionError):
    """La session SharePoint est absente, expirée, ou n'ouvre pas ce dossier."""


class TelechargementInterdit(PermissionError):
    """Accès conditionnel de l'organisation : téléchargement interdit depuis cet appareil.

    SharePoint laisse lister les dossiers mais refuse le contenu des fichiers
    (« vous ne pouvez pas télécharger ces ressources à partir de cet appareil »,
    code ``conditionalAccessPolicyEnforced``) — typiquement sur un poste non géré
    (ni joint ni enregistré dans l'annuaire de l'organisation). C'est une règle de
    sécurité de l'organisation : l'outil la signale, il ne la contourne pas.
    """

    EXPLICATION = (
        "Votre organisation interdit le téléchargement des fichiers SharePoint depuis cet "
        "appareil (accès conditionnel : poste non géré). La liste des dossiers reste lisible, "
        "mais pas le contenu des fichiers audit.\n\n"
        "• Utilisez l'outil sur un poste de l'entreprise (géré), ou\n"
        "• demandez à votre service informatique d'enregistrer ce poste ou d'autoriser cet usage, ou\n"
        "• auditez une copie locale des fichiers (source « Dossier local »)."
    )


class VersionModifiee(RuntimeError):
    """Le fichier a été enregistré par quelqu'un d'autre depuis sa lecture (``If-Match``)."""


class FichierVerrouille(RuntimeError):
    """Fichier extrait ou verrouillé (ouvert en co-édition) : SharePoint refuse l'écriture."""


def site_de(chemin_serveur: str) -> str:
    """Chemin du site qui contient un chemin serveur (``/sites/X/Docs/…`` → ``/sites/X``)."""
    correspondance = _SITE.match(chemin_serveur)
    return correspondance.group(0) if correspondance else ""


def chemin_serveur(url: str) -> str:
    """Adresse https d'un dossier ou d'un fichier → chemin serveur décodé."""
    return unquote(urlsplit(url).path).rstrip("/")


class ClientSharePoint:
    """Appels en lecture à l'API SharePoint, avec les cookies de la session navigateur."""

    def __init__(self, hote: str, cookies: dict[str, str], delai_s: float = 60, base: str | None = None):
        self.hote = hote.lower()
        self.cookies = dict(cookies)
        self.delai_s = delai_s
        self.base = (base or f"https://{self.hote}").rstrip("/")   # remplaçable en test
        self._arborescences: dict[str, tuple[float, dict[str, list[dict]]]] = {}
        self._verrou_arbo = threading.Lock()

    # ── Requêtes ──────────────────────────────────────────────────
    def _requete(self, chemin: str, binaire: bool = False) -> bytes:
        return self._echange(chemin, binaire=binaire)[0]

    def _echange(self, chemin: str, binaire: bool = False, methode: str = "GET",
                 donnees: bytes | None = None, entetes: dict | None = None) -> tuple[bytes, dict]:
        """Requête à l'API ; renvoie ``(corps, en-têtes de réponse)``."""
        requete = urllib.request.Request(f"{self.base}{chemin}", data=donnees, method=methode, headers={
            "Cookie": "; ".join(f"{k}={v}" for k, v in self.cookies.items()),
            "Accept": "application/octet-stream" if binaire else "application/json;odata=nometadata",
            "User-Agent": "DispatcherSNA",
            **(entetes or {}),
        })
        try:
            with urllib.request.urlopen(requete, timeout=self.delai_s) as reponse:
                if urlsplit(reponse.geturl()).hostname != urlsplit(self.base).hostname:
                    raise AuthentificationRequise("session SharePoint absente ou expirée")
                return reponse.read(), dict(reponse.headers.items())
        except urllib.error.HTTPError as exc:
            corps = _corps(exc)
            message = _message_erreur(corps, exc.reason)
            if exc.code == 403 and "conditionalaccesspolicyenforced" in corps.lower():
                raise TelechargementInterdit(
                    f"téléchargement interdit depuis cet appareil par les stratégies de l'organisation "
                    f"({message})") from None
            if exc.code == 412:
                raise VersionModifiee("fichier enregistré par quelqu'un d'autre depuis sa lecture") from None
            if exc.code == 423 or _est_verrou(message):
                raise FichierVerrouille(f"fichier verrouillé sur SharePoint ({message or 'ouvert ou extrait'})") from None
            if exc.code in (401, 403):
                raise AuthentificationRequise(
                    "session SharePoint absente ou expirée (ou droits insuffisants sur ce dossier)"
                ) from None
            if exc.code == 404:
                raise FileNotFoundError(f"introuvable sur SharePoint : {chemin}") from None
            raise OSError(f"SharePoint a répondu {exc.code} : {message}") from None

    def _api(self, chemin_serveur: str, ressource: str, suite: str = "") -> str:
        valeur = quote(chemin_serveur.replace("'", "''"), safe="/")
        return f"{site_de(chemin_serveur)}/_api/web/{ressource}(decodedurl='{valeur}'){suite}"

    def _json(self, chemin: str) -> dict:
        return json.loads(self._requete(chemin).decode("utf-8"))

    # ── Dossiers et fichiers ──────────────────────────────────────
    def dossier_existe(self, chemin_serveur: str) -> bool:
        try:
            return bool(self._json(self._api(chemin_serveur, "GetFolderByServerRelativePath",
                                              "?$select=Exists")).get("Exists"))
        except FileNotFoundError:
            return False

    def verifier(self, chemin_serveur: str) -> None:
        """Lève si la session n'ouvre pas ce dossier (ou s'il n'existe pas)."""
        if not self.dossier_existe(chemin_serveur):
            raise FileNotFoundError(f"dossier introuvable sur SharePoint : {chemin_serveur}")

    def dossiers(self, chemin_serveur: str) -> list[str]:
        donnees = self._json(self._api(chemin_serveur, "GetFolderByServerRelativePath",
                                       "/Folders?$select=Name&$top=5000"))
        return sorted(d["Name"] for d in donnees.get("value", []))

    def fichiers(self, chemin_serveur: str) -> list[dict]:
        """``[{"Name", "TimeLastModified"}]`` des fichiers d'un dossier, triés par nom."""
        donnees = self._json(self._api(chemin_serveur, "GetFolderByServerRelativePath",
                                       "/Files?$select=Name,TimeLastModified&$top=5000"))
        return sorted(donnees.get("value", []), key=lambda f: f["Name"])

    def dossiers_et_fichiers(self, chemin_serveur: str) -> dict[str, list[dict]]:
        """Sous-dossiers d'un dossier et leurs fichiers, en **une** requête."""
        donnees = self._json(self._api(
            chemin_serveur, "GetFolderByServerRelativePath",
            "/Folders?$select=Name,Files/Name,Files/TimeLastModified&$expand=Files&$top=5000",
        ))
        return {d["Name"]: sorted(d.get("Files", []), key=lambda f: f["Name"]) for d in donnees.get("value", [])}

    def telecharger(self, chemin_fichier: str, destination: str) -> str:
        contenu = self._requete(self._api(chemin_fichier, "GetFileByServerRelativePath", "/$value"),
                                binaire=True)
        with open(destination, "wb") as f:
            f.write(contenu)
        return destination

    # ── Écriture ──────────────────────────────────────────────────
    def telecharger_version(self, chemin_fichier: str, destination: str) -> str:
        """Télécharge le fichier et renvoie l'ETag **du contenu téléchargé**.

        L'ETag vient de la réponse elle-même quand SharePoint le fournit ; sinon il
        est lu avant et après le téléchargement, qui est refait s'il a changé
        entre les deux (enregistrement concurrent).
        """
        for _ in range(3):
            avant = self.etat_fichier(chemin_fichier)["ETag"]
            contenu, entetes = self._echange(
                self._api(chemin_fichier, "GetFileByServerRelativePath", "/$value"), binaire=True)
            etag = next((v for k, v in entetes.items() if k.lower() == "etag"), "") or ""
            if not etag:
                etag = self.etat_fichier(chemin_fichier)["ETag"]
                if etag != avant:
                    continue
            with open(destination, "wb") as f:
                f.write(contenu)
            return etag
        raise VersionModifiee("fichier enregistré en continu par quelqu'un d'autre : rejouez plus tard")

    def etat_fichier(self, chemin_fichier: str) -> dict:
        """``{"ETag", "CheckOutType", "CheckedOutBy"}`` : version courante et extraction."""
        donnees = self._json(self._api(chemin_fichier, "GetFileByServerRelativePath",
                                       "?$select=ETag,CheckOutType&$expand=CheckedOutByUser"))
        qui = (donnees.get("CheckedOutByUser") or {}).get("Title", "")
        return {"ETag": donnees.get("ETag", ""), "CheckOutType": donnees.get("CheckOutType", 2),
                "CheckedOutBy": qui}

    def verrou_par(self, chemin_fichier: str) -> str:
        """Nom de la personne qui a le fichier ouvert (verrou de co-édition / d'édition),
        ``""`` si SharePoint ne le dit pas. Ne lève jamais : sert à expliquer un refus."""
        try:
            donnees = self._json(self._api(chemin_fichier, "GetFileByServerRelativePath",
                                           "/LockedByUser?$select=Title,Email"))
        except Exception:  # noqa: BLE001
            return ""
        return str(donnees.get("Title") or donnees.get("Email") or "").strip()

    def jeton(self, chemin_serveur: str) -> str:
        """Jeton de formulaire (« request digest ») exigé par SharePoint pour écrire."""
        corps, _ = self._echange(f"{site_de(chemin_serveur)}/_api/contextinfo", methode="POST", donnees=b"")
        return json.loads(corps.decode("utf-8"))["FormDigestValue"]

    def remplacer(self, chemin_fichier: str, contenu: bytes, etag: str) -> None:
        """Remplace le contenu du fichier, **seulement** s'il est encore à la version ``etag``.

        Raises:
            VersionModifiee: quelqu'un l'a enregistré depuis.
            FichierVerrouille: fichier extrait / verrouillé.
        """
        self._echange(
            self._api(chemin_fichier, "GetFileByServerRelativePath", "/$value"),
            methode="POST", donnees=contenu,
            entetes={"X-HTTP-Method": "PUT", "If-Match": etag or "*",
                     "X-RequestDigest": self.jeton(chemin_fichier),
                     "Content-Type": "application/octet-stream"},
        )

    # ── Ajouts du Dispatcher ──────────────────────────────────────
    def deposer(self, chemin_dossier: str, nom: str, contenu: bytes) -> None:
        """Dépose un fichier (le crée, ou le remplace s'il existe). Pour remplacer un fichier
        qu'on a LU, préférer :meth:`remplacer` (If-Match : ne jamais écraser la version d'un autre).

        Raises:
            FichierVerrouille: fichier existant ouvert / extrait.
        """
        nom_api = quote(nom.replace("'", "''"), safe="")
        self._echange(
            self._api(chemin_dossier, "GetFolderByServerRelativePath",
                      f"/Files/AddUsingPath(decodedurl='{nom_api}',overwrite=true)"),
            methode="POST", donnees=contenu,
            entetes={"X-RequestDigest": self.jeton(chemin_dossier), "Content-Type": "application/octet-stream"},
        )

    def creer_dossier(self, chemin_dossier: str) -> None:
        """Crée le dossier s'il n'existe pas (ex. ``…/Dep74/74143/analyse``)."""
        if self.dossier_existe(chemin_dossier):
            return
        valeur = quote(chemin_dossier.replace("'", "''"), safe="/")
        self._echange(f"{site_de(chemin_dossier)}/_api/web/Folders/AddUsingPath(decodedurl='{valeur}')",
                      methode="POST", donnees=b"", entetes={"X-RequestDigest": self.jeton(chemin_dossier)})

    def recycler_dossier(self, chemin_dossier: str) -> None:
        """Met le dossier à la **corbeille** du site (restaurable), plutôt que de le supprimer."""
        self._echange(self._api(chemin_dossier, "GetFolderByServerRelativePath", "/recycle()"),
                      methode="POST", donnees=b"", entetes={"X-RequestDigest": self.jeton(chemin_dossier)})

    def telecharger_dossier(self, chemin_dossier: str, destination: str, profondeur: int = 3) -> list[str]:
        """Télécharge un dossier et ses sous-dossiers (jusqu'à ``profondeur``) dans ``destination``,
        pour LECTURE seule (pas d'ETag). Retourne les chemins locaux des fichiers."""
        import os
        os.makedirs(destination, exist_ok=True)
        locaux: list[str] = []
        for f in self.fichiers(chemin_dossier):
            if f["Name"].startswith("~$"):
                continue
            locaux.append(self.telecharger(f"{chemin_dossier}/{f['Name']}", os.path.join(destination, f["Name"])))
        if profondeur > 1:
            for d in self.dossiers(chemin_dossier):
                locaux += self.telecharger_dossier(f"{chemin_dossier}/{d}", os.path.join(destination, d),
                                                   profondeur - 1)
        return locaux

    def arborescence(self, racine: str, forcer: bool = False, age_max_s: float = 600,
                     paralleles: int = 8) -> dict[str, list[dict]]:
        """``{CODE COMMUNE: [fichiers]}`` de ``<racine>/Dep*/<commune>``, gardé en mémoire.

        Scan rapide : une requête pour la liste des ``Dep<XX>``, puis une par
        ``Dep<XX>`` (ses communes **et** leurs fichiers ensemble), en parallèle —
        quelques requêtes pour des centaines de communes. Le résultat sert ensuite à
        vérifier n'importe quel nombre de communes ou d'ID collés **sans requête**.
        Refait après ``age_max_s`` secondes, ou si ``forcer``.
        """
        with self._verrou_arbo:
            memo = self._arborescences.get(racine)
            if memo and not forcer and time.monotonic() - memo[0] < age_max_s:
                return memo[1]
        departements = [d for d in self.dossiers(racine) if d.lower().startswith(PREFIXE_DEP.lower())]
        with ThreadPoolExecutor(max_workers=max(1, min(paralleles, len(departements) or 1))) as pool:
            contenus = list(pool.map(lambda d: self.dossiers_et_fichiers(f"{racine}/{d}"), departements))
        arbo: dict[str, list[dict]] = {}
        for contenu in contenus:
            for code, fichiers in contenu.items():
                arbo[code.upper()] = fichiers
        with self._verrou_arbo:
            self._arborescences[racine] = (time.monotonic(), arbo)
        return arbo

    def url(self, chemin_serveur: str) -> str:
        return f"https://{self.hote}{quote(chemin_serveur, safe='/')}"


# ── Règles de complete_excel_files, appliquées aux listes de l'API ─
def fichier_audit(fichiers: list[dict]) -> str | None:
    """Premier ``audit_*.xlsx`` par ordre de nom (même règle que ``trouver_fichier_audit``)."""
    return next((f["Name"] for f in fichiers if _est_audit(f["Name"])), None)


def copies_de_conflit(fichiers: list[dict]) -> list[tuple[str, str]]:
    """``[(nom, date)]`` des ``audit_*.xlsx`` surnuméraires (cf. ``detecter_fichiers_conflit``)."""
    principal = fichier_audit(fichiers)
    resultat = []
    for f in fichiers:
        if _est_audit(f["Name"]) and f["Name"] != principal:
            date = _MOTIF_DATE_NOM.search(f["Name"])
            resultat.append((f["Name"], date.group(0) if date else _date(f.get("TimeLastModified"))))
    return resultat


def lister_communes(client: ClientSharePoint, racine: str, forcer: bool = True):
    """``[(code, département, a_fichier)]`` de toutes les communes ``<racine>/Dep*/<commune>``."""
    return sorted(
        (code, code[:2], fichier_audit(fichiers) is not None)
        for code, fichiers in client.arborescence(racine, forcer=forcer).items()
    )


def _est_verrou(message: str) -> bool:
    texte = (message or "").lower()
    return any(m in texte for m in ("locked", "verrouill", "checked out", "extrait"))


def _est_audit(nom: str) -> bool:
    return nom.startswith(PREFIXE_AUDIT) and nom.endswith(EXTENSION) and not nom.startswith("~$")


def _date(horodatage: str | None) -> str:
    try:
        return datetime.fromisoformat(str(horodatage).replace("Z", "+00:00")).astimezone().strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return "date inconnue"


def _corps(exc: urllib.error.HTTPError) -> str:
    try:
        return exc.read().decode("utf-8", errors="replace")
    except Exception:  # noqa: BLE001
        return ""


def _message_erreur(corps: str, raison: str = "") -> str:
    """Message d'une réponse d'erreur SharePoint (formats « odata.error » et « error »)."""
    try:
        donnees = json.loads(corps)
    except ValueError:
        return raison or corps[:200]
    message = donnees.get("odata.error", {}).get("message", {})
    if isinstance(message, dict) and message.get("value"):
        return message["value"]
    erreur = donnees.get("error", {})
    return erreur.get("message") or str(donnees)[:200]
