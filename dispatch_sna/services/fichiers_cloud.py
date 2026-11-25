"""Lecture / écriture de fichiers robustes dans un dossier synchronisé par OneDrive.

Pourquoi : dans un dossier OneDrive, un fichier peut être « disponible en ligne uniquement »
(simple coquille sur le disque). Plusieurs opérations courantes y échouent alors avec une
erreur du fournisseur cloud (ERROR_CLOUD_FILE_*), la plus fréquente étant
« [WinError 389] L'opération de cloud a échoué » (0x80070185) :
  - renommer un fichier PAR-DESSUS une coquille en ligne uniquement (os.replace) ;
  - lire un fichier quand OneDrive est saturé, suspendu, arrêté (WinError 362) ou sans disque ;
  - shutil.copy2, qui recopie aussi les attributs du fichier source (lecture seule,
    attributs OneDrive), source de refus supplémentaires.

Ce module :
  - réessaie les erreurs cloud, qui sont souvent passagères (OneDrive occupé) ;
  - écrit uniquement les octets (jamais les attributs du fichier source) ;
  - remplace un fichier en ligne uniquement en trois temps : remplacement direct, sinon
    rapatriement de l'ancien fichier puis remplacement, sinon suppression puis renommage
    (ce que fait déjà le mode « Commune complète », qui n'échoue pas) ;
  - vérifie la taille écrite ;
  - lève ErreurOneDrive avec l'étape en cause et des pistes, au lieu d'un WinError brut.
"""
from __future__ import annotations

import errno
import os
import time

# ERROR_CLOUD_FILE_* (362 fournisseur arrêté, 389 opération de cloud échouée, 375 en cours
# d'utilisation, 377/382 requête annulée…) : 0x166 → 0x18F
ERREURS_CLOUD = range(358, 400)
_EN_LIGNE_SEULEMENT = 0x00400000 | 0x00040000 | 0x00001000   # RECALL_ON_DATA_ACCESS | RECALL_ON_OPEN | OFFLINE
PAUSES_S = (1, 3, 8, 15)          # entre deux essais d'une même opération (~27 s au total)
_dormir = time.sleep              # remplaçable dans les tests


class ErreurOneDrive(OSError):
    """OneDrive a refusé l'opération, même après plusieurs essais."""


def en_ligne_seulement(chemin) -> bool:
    try:
        return bool(getattr(os.stat(chemin), "st_file_attributes", 0) & _EN_LIGNE_SEULEMENT)
    except OSError:
        return False


def est_erreur_cloud(e, chemin=None) -> bool:
    if not isinstance(e, OSError) or isinstance(e, (FileNotFoundError, PermissionError)):
        return False
    if getattr(e, "winerror", None) in ERREURS_CLOUD:
        return True
    # Lecture d'une coquille quand OneDrive ne répond pas : Windows renvoie parfois EINVAL
    return getattr(e, "errno", None) == errno.EINVAL and chemin is not None and en_ligne_seulement(chemin)


def _essayer(action, etape, chemin, pauses=PAUSES_S):
    """Exécute `action`, réessaie sur erreur cloud ; ErreurOneDrive si elle persiste."""
    for k in range(len(pauses) + 1):
        try:
            return action()
        except OSError as e:
            if not est_erreur_cloud(e, chemin):
                raise
            if k == len(pauses):
                raise ErreurOneDrive(
                    f"OneDrive a refusé {etape} après {k + 1} essais ({_court(e)}). "
                    "Vérifiez que OneDrive est lancé, connecté, non suspendu, à jour dans sa "
                    "synchronisation et qu'il reste de l'espace disque") from e
            _dormir(pauses[k])


def _court(e) -> str:
    code = getattr(e, "winerror", None)
    return f"WinError {code} : {e.strerror}" if code else str(e)


def lire_octets(chemin) -> bytes:
    """Contenu complet d'un fichier (force le rapatriement d'un fichier en ligne uniquement)."""
    def lire():
        with open(chemin, "rb") as f:
            return f.read()
    return _essayer(lire, f"la lecture de {os.path.basename(chemin)}", chemin)


def remplacer(tmp, dst, pauses=PAUSES_S):
    """os.replace(tmp, dst) robuste quand dst est un fichier OneDrive (en ligne uniquement)."""
    nom = os.path.basename(dst)
    courtes = pauses[:2]
    try:
        return _essayer(lambda: os.replace(tmp, dst), f"le remplacement de {nom}", dst, courtes)
    except ErreurOneDrive:
        if not os.path.exists(dst):
            raise
    # 2. Rapatrier l'ancien fichier (il redevient un fichier local ordinaire), puis remplacer.
    #    Son contenu est gardé en mémoire : c'est ce qui autorise l'étape 3 sans risque de perte.
    ancien = lire_octets(dst)                      # ErreurOneDrive si impossible : on s'arrête là
    try:
        return _essayer(lambda: os.replace(tmp, dst), f"le remplacement de {nom}", dst, courtes)
    except ErreurOneDrive:
        pass
    # 3. Dernier recours : supprimer l'ancien puis renommer le nouveau à sa place
    _essayer(lambda: os.remove(dst), f"la suppression de l'ancien {nom}", dst, pauses)
    try:
        return _essayer(lambda: os.replace(tmp, dst), f"le dépôt de {nom}", dst, pauses)
    except ErreurOneDrive:
        # Ne jamais laisser la destination sans fichier : on remet l'ancien contenu
        try:
            with open(dst, "wb") as f:
                f.write(ancien)
        except OSError:
            pass
        raise


def ecrire_octets(dst, data: bytes):
    """Écrit `data` dans dst via un fichier temporaire ; ne copie aucun attribut ; vérifie la taille."""
    tmp = dst + ".tmp"

    def ecrire_tmp():
        with open(tmp, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
    try:
        _essayer(ecrire_tmp, f"l'écriture de {os.path.basename(dst)}", tmp)
        remplacer(tmp, dst)
    finally:
        if os.path.exists(tmp):
            try: os.remove(tmp)
            except OSError: pass
    taille = _essayer(lambda: os.path.getsize(dst), f"la vérification de {os.path.basename(dst)}", dst)
    if taille != len(data):
        raise ErreurOneDrive(f"{os.path.basename(dst)} : {taille} octets écrits au lieu de {len(data)}")


def copier(src, dst):
    """Copie le contenu de src vers dst (remplace dst), robuste OneDrive, sans les attributs de src."""
    ecrire_octets(dst, lire_octets(src))
