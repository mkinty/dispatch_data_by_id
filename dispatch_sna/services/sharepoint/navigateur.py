"""(Repris d'audit_validation_finale, avf/services/navigateur.py ; profil propre au Dispatcher.)

Navigateur piloté pour la connexion SharePoint (Microsoft Edge, sinon Google Chrome).

L'utilisateur s'identifie dans une **vraie fenêtre de navigateur** (mot de passe,
MFA, accès conditionnel…). L'outil récupère ensuite les cookies de session de
SharePoint (``FedAuth``, ``rtFa``) par le protocole de débogage du navigateur
(Chrome DevTools Protocol), ouvert **uniquement sur 127.0.0.1**, et s'en sert
pour interroger l'API SharePoint avec l'identité de l'utilisateur
(cf. ``services.sharepoint.rest``).

Le navigateur tourne sur un **profil dédié** (``%LOCALAPPDATA%\\DispatcherSNA\\
navigateur``), jamais sur celui de l'utilisateur : rien n'est lu dans ses autres
sessions. Le profil est conservé : avec « Rester connecté », la connexion
suivante se fait sans fenêtre (navigateur sans affichage). L'outil n'enregistre
lui-même aucun cookie ni mot de passe : ils restent dans le profil du
navigateur, chiffrés par Windows.

Aucune dépendance : le client WebSocket nécessaire au protocole est inclus.
"""

from __future__ import annotations

import base64
import json
import os
import socket
import struct
import subprocess
import time
import urllib.request
from urllib.parse import urlsplit

_EXECUTABLES = [
    r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe",
    r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe",
    r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe",
    r"%ProgramFiles%\Google\Chrome\Application\chrome.exe",
    r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe",
    r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe",
]


class NavigateurIndisponible(RuntimeError):
    """Ni Edge ni Chrome trouvés, ou le navigateur n'a pas démarré."""


class NavigateurFerme(RuntimeError):
    """L'utilisateur a fermé la fenêtre du navigateur."""


def trouver_navigateur() -> str | None:
    """Chemin de Microsoft Edge (sinon Google Chrome), ou ``None``."""
    for modele in _EXECUTABLES:
        chemin = os.path.expandvars(modele)
        if os.path.isfile(chemin):
            return chemin
    return None


def dossier_profil() -> str:
    """Profil de navigateur réservé à l'outil (conservé entre deux lancements)."""
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "DispatcherSNA", "navigateur")


class NavigateurPilote:
    """Un navigateur lancé avec le protocole de débogage, sur un profil dédié.

        with NavigateurPilote(url, visible=True) as nav:
            cookies = nav.cookies()
    """

    def __init__(self, url: str, visible: bool = True, profil: str | None = None,
                 executable: str | None = None, delai_demarrage_s: float = 30):
        self.url = url
        self.visible = visible
        self.profil = profil or dossier_profil()
        self.executable = executable or trouver_navigateur()
        self.delai_demarrage_s = delai_demarrage_s
        self.processus: subprocess.Popen | None = None
        self.port: int | None = None

    def __enter__(self) -> "NavigateurPilote":
        self.demarrer()
        return self

    def __exit__(self, *exc) -> bool:
        self.fermer()
        return False

    def demarrer(self) -> None:
        if not self.executable:
            raise NavigateurIndisponible("ni Microsoft Edge ni Google Chrome n'ont été trouvés sur ce poste")
        os.makedirs(self.profil, exist_ok=True)
        fichier_port = os.path.join(self.profil, "DevToolsActivePort")
        if os.path.exists(fichier_port):
            os.remove(fichier_port)
        arguments = [
            self.executable, "--remote-debugging-port=0", "--remote-allow-origins=http://127.0.0.1",
            f"--user-data-dir={self.profil}", "--no-first-run", "--no-default-browser-check",
            "--disable-features=msEdgeSidebarV2,msUndersideButton",
        ]
        arguments += ["--new-window"] if self.visible else ["--headless=new", "--disable-gpu"]
        self.processus = subprocess.Popen(
            arguments + [self.url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0) if not self.visible else 0,
        )
        limite = time.monotonic() + self.delai_demarrage_s
        while time.monotonic() < limite:
            if self.processus.poll() is not None:
                raise NavigateurIndisponible(
                    "le navigateur s'est arrêté au démarrage (profil déjà ouvert dans une autre fenêtre ?)")
            try:
                with open(fichier_port, encoding="utf-8") as f:
                    self.port = int(f.readline().strip())
                self._version()
                return
            except (OSError, ValueError):
                time.sleep(0.2)
        self.fermer()
        raise NavigateurIndisponible(f"le navigateur n'a pas répondu en {self.delai_demarrage_s:.0f} s")

    def ouvert(self) -> bool:
        return self.processus is not None and self.processus.poll() is None

    def _version(self) -> dict:
        with urllib.request.urlopen(f"http://127.0.0.1:{self.port}/json/version", timeout=5) as reponse:
            return json.loads(reponse.read().decode("utf-8"))

    def _appel(self, methode: str, params: dict | None = None) -> dict:
        if not self.ouvert():
            raise NavigateurFerme("fenêtre du navigateur fermée")
        try:
            adresse = self._version()["webSocketDebuggerUrl"]
        except OSError as exc:
            raise NavigateurFerme("fenêtre du navigateur fermée") from exc
        with _WebSocket(adresse) as ws:
            return ws.appel(methode, params or {})

    def cookies(self) -> list[dict]:
        """Tous les cookies du profil (``HttpOnly`` compris), au format CDP."""
        return self._appel("Storage.getCookies").get("cookies", [])

    def url_courante(self) -> str:
        """Adresse de l'onglet affiché (pour suivre la connexion)."""
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{self.port}/json/list", timeout=5) as reponse:
                pages = [p for p in json.loads(reponse.read().decode("utf-8")) if p.get("type") == "page"]
        except OSError:
            return ""
        return pages[0].get("url", "") if pages else ""

    def fermer(self) -> None:
        if self.processus is None:
            return
        if self.ouvert() and self.port:
            try:
                self._appel("Browser.close")
            except Exception:  # noqa: BLE001 — fermeture « polie » impossible : on termine
                pass
            try:
                self.processus.wait(10)
            except subprocess.TimeoutExpired:
                pass
        if self.ouvert():
            self.processus.terminate()
            try:
                self.processus.wait(5)
            except subprocess.TimeoutExpired:
                self.processus.kill()
        self.processus = None


def cookies_pour(cookies: list[dict], hote: str) -> dict[str, str]:
    """Cookies (format CDP) envoyés par un navigateur à ``https://hote`` : ``{nom: valeur}``."""
    hote = hote.lower()
    retenus: dict[str, str] = {}
    for c in sorted(cookies, key=lambda c: len(c.get("domain", ""))):   # le plus précis l'emporte
        domaine = c.get("domain", "").lower().lstrip(".")
        if hote == domaine or hote.endswith("." + domaine):
            if c.get("expires", -1) not in (-1, 0) and c["expires"] < time.time():
                continue
            retenus[c["name"]] = c["value"]
    return retenus


# ── Client WebSocket minimal (RFC 6455), pour le protocole de débogage local ──
class _WebSocket:
    def __init__(self, adresse: str, delai_s: float = 15):
        morceaux = urlsplit(adresse)
        if morceaux.hostname not in ("127.0.0.1", "localhost"):
            raise ValueError("le protocole de débogage n'est utilisé qu'en local")
        self._socket = socket.create_connection((morceaux.hostname, morceaux.port), timeout=delai_s)
        cle = base64.b64encode(os.urandom(16)).decode()
        chemin = morceaux.path + (f"?{morceaux.query}" if morceaux.query else "")
        self._socket.sendall((
            f"GET {chemin} HTTP/1.1\r\nHost: {morceaux.hostname}:{morceaux.port}\r\n"
            "Upgrade: websocket\r\nConnection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {cle}\r\nSec-WebSocket-Version: 13\r\n\r\n"
        ).encode())
        entete = b""
        while b"\r\n\r\n" not in entete:
            morceau = self._socket.recv(4096)
            if not morceau:
                raise ConnectionError("poignée de main WebSocket interrompue")
            entete += morceau
        entete, self._tampon = entete.split(b"\r\n\r\n", 1)
        if b" 101 " not in entete.split(b"\r\n", 1)[0]:
            raise ConnectionError(f"poignée de main WebSocket refusée : {entete[:80]!r}")
        self._id = 0

    def __enter__(self) -> "_WebSocket":
        return self

    def __exit__(self, *exc) -> bool:
        try:
            self._trame(0x8, b"")
        except OSError:
            pass
        self._socket.close()
        return False

    def appel(self, methode: str, params: dict) -> dict:
        self._id += 1
        self._trame(0x1, json.dumps({"id": self._id, "method": methode, "params": params}).encode())
        while True:
            message = json.loads(self._recevoir())
            if message.get("id") == self._id:
                if "error" in message:
                    raise RuntimeError(f"{methode} : {message['error'].get('message')}")
                return message.get("result", {})

    def _trame(self, code: int, donnees: bytes) -> None:
        masque = os.urandom(4)
        n = len(donnees)
        entete = bytes([0x80 | code])
        if n < 126:
            entete += bytes([0x80 | n])
        elif n < 65536:
            entete += bytes([0x80 | 126]) + struct.pack(">H", n)
        else:
            entete += bytes([0x80 | 127]) + struct.pack(">Q", n)
        self._socket.sendall(entete + masque + bytes(b ^ masque[i % 4] for i, b in enumerate(donnees)))

    def _lire(self, n: int) -> bytes:
        while len(self._tampon) < n:
            morceau = self._socket.recv(65536)
            if not morceau:
                raise ConnectionError("connexion WebSocket fermée")
            self._tampon += morceau
        donnees, self._tampon = self._tampon[:n], self._tampon[n:]
        return donnees

    def _recevoir(self) -> str:
        message = b""
        while True:
            octet1, octet2 = self._lire(2)
            fin, code = octet1 & 0x80, octet1 & 0x0F
            n = octet2 & 0x7F
            if n == 126:
                n = struct.unpack(">H", self._lire(2))[0]
            elif n == 127:
                n = struct.unpack(">Q", self._lire(8))[0]
            masque = self._lire(4) if octet2 & 0x80 else None
            donnees = self._lire(n)
            if masque:
                donnees = bytes(b ^ masque[i % 4] for i, b in enumerate(donnees))
            if code == 0x9:                       # ping → pong
                self._trame(0xA, donnees)
                continue
            if code == 0x8:
                raise ConnectionError("connexion WebSocket fermée par le navigateur")
            if code in (0x0, 0x1, 0x2):
                message += donnees
                if fin:
                    return message.decode("utf-8")
