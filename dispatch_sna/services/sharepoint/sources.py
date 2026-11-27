"""Les deux sources du Dispatcher (Racine SharePoint, Dossier à déposer) : dossier local OU lien.

Un **dossier local** (synchronisé OneDrive) est utilisé tel quel, comme avant.

Un **lien SharePoint** est lu et écrit par l'API SharePoint, avec la session ouverte dans le
navigateur (services.sharepoint.connexion), comme dans audit_validation_finale. La logique
métier (services.dispatcher) ne change pas : pour chaque commune, un **dossier de travail**
temporaire reproduit la partie utile de SharePoint :

- Dossier à déposer par lien : le dossier de la commune est téléchargé ; s'il est supprimé
  par le rangement (tout a réussi), le dossier SharePoint est mis à la **corbeille** du site
  (restaurable) ;
- Racine par lien : ``DepXX/<commune>/`` n'est créé que si ce dossier existe sur SharePoint ;
  l'Excel audit et les PPT à modifier sont téléchargés avec leur **ETag** (version lue).
  Chaque fichier écrit est renvoyé aussitôt (``publier``) :

  * fichier lu → remplacé **seulement s'il n'a pas changé depuis** (``If-Match``) ;
  * nouveau fichier → déposé (le dossier ``analyse`` est créé au besoin) ;
  * Excel audit (mise à jour par ID) : téléchargé, lignes écrites, renvoyé avec If-Match puis
    RELU sur le serveur ; version modifiée entre-temps → on recommence sur la nouvelle version ;
    fichier **ouvert sur SharePoint** (verrou 423) → écriture **cellule par cellule en
    co-édition Excel**, par son adresse, comme audit_validation_finale ; Excel indisponible →
    rien n'est écrit et l'ID passe en attente, avec le nom de la personne ;
  * PPT **ouvert sur SharePoint** → écriture **en co-édition PowerPoint** par son adresse :
    seules les zones Analyse / Lien Street View (modifier), ou les diapositives (remplacer).
"""
from __future__ import annotations

import os
import shutil
import tempfile
from concurrent.futures import ThreadPoolExecutor

from .. import chemins
from ..constantes import AUDIT_PREFIX, COL_ID_ERR, COL_ID_IMM, DEP_PREFIX, PPT_FOLDER
from ..error_ids import nom_ppt
from ..excel_rows import OngletIntrouvable, cle_de, lire_cellules, lire_entetes, lire_ids_erreur, maj_lignes_excel
from . import rest
from .connexion import analyser_lien, client_pour, est_lien

ESSAIS_VERSION = 3          # nouvelles tentatives si l'Excel est enregistré par un autre entre-temps
VOIE_API = "écrit par l'API SharePoint (session du navigateur)"


def _message_session(e) -> str:
    return f"session SharePoint expirée ou refusée — cliquez sur « 🔗 Se connecter » puis rejouez ({e})"


class Sources:
    """Résout les deux sources d'un lot. ``with Sources(prod, root) as s:`` ; le dossier de
    travail est supprimé à la sortie. Lève rest.AuthentificationRequise si un lien est saisi
    sans connexion établie, connexion.LienInvalide si le lien est inexploitable."""

    def __init__(self, prod: str, root: str):
        self.prod_texte, self.root_texte = (prod or "").strip(), (root or "").strip()
        self.prod_lien = analyser_lien(self.prod_texte) if est_lien(self.prod_texte) else None
        self.root_lien = analyser_lien(self.root_texte) if est_lien(self.root_texte) else None
        self.client_prod = client_pour(self.prod_texte) if self.prod_lien else None
        self.client_root = client_pour(self.root_texte) if self.root_lien else None
        self._tmp = None
        self._prod_serveur: dict[str, tuple[str, str]] | None = None   # commune → (chemin serveur, nom)
        self._dest: dict[str, tuple[str, list[dict]]] | None = None     # commune → (chemin serveur, fichiers)
        self.session_expiree = ""

    # ── Cycle de vie ──────────────────────────────────────────────
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        if self._tmp:
            shutil.rmtree(self._tmp, ignore_errors=True)
            self._tmp = None
        return False

    @property
    def distant(self) -> bool:
        return bool(self.prod_lien or self.root_lien)

    def _travail(self) -> str:
        if self._tmp is None:
            self._tmp = tempfile.mkdtemp(prefix="dispatch_sna_")
        return self._tmp

    @property
    def prod_local(self) -> str:
        return os.path.join(self._travail(), "prod") if self.prod_lien else self.prod_texte

    @property
    def root_local(self) -> str:
        return os.path.join(self._travail(), "sp") if self.root_lien else self.root_texte

    def libelle(self, quoi: str) -> str:
        lien = self.prod_lien if quoi == "prod" else self.root_lien
        return f"{lien.url} (SharePoint, par l'API)" if lien else (self.prod_texte if quoi == "prod" else self.root_texte)

    # ── Inventaire ────────────────────────────────────────────────
    def communes_prod(self) -> dict[str, str]:
        """{commune → dossier local} : pour un lien, le dossier de travail (rempli par preparer)."""
        if not self.prod_lien:
            return chemins.lister_communes_prod(self.prod_texte)
        if self._prod_serveur is None:
            racine = self.prod_lien.chemin
            res = {}
            for d in self.client_prod.dossiers(racine):     # même règle que lister_communes_prod
                com = chemins.extraire_commune_nom(d)
                if com and (com not in res or d.upper() == com):
                    res[com] = (f"{racine}/{d}", d)
            self._prod_serveur = res
        base = os.path.join(self._travail(), "prod")
        return {c: os.path.join(base, nom) for c, (_s, nom) in self._prod_serveur.items()}

    def _prod_serveur_charge(self) -> dict[str, tuple[str, str]]:
        self.communes_prod()
        return self._prod_serveur

    def _destinations(self) -> dict[str, tuple[str, list[dict]]]:
        if self._dest is None:
            racine, client = self.root_lien.chemin, self.client_root
            deps = [d for d in client.dossiers(racine) if d.lower().startswith(DEP_PREFIX.lower())]
            with ThreadPoolExecutor(max_workers=max(1, min(8, len(deps) or 1))) as pool:
                contenus = list(pool.map(lambda d: (d, client.dossiers_et_fichiers(f"{racine}/{d}")), deps))
            self._dest = {nom.upper(): (f"{racine}/{d}/{nom}", fichiers)
                          for d, contenu in contenus for nom, fichiers in contenu.items()}
        return self._dest

    def dest_existe(self, commune: str) -> bool:
        if not self.root_lien:
            return os.path.isdir(os.path.join(self.root_texte, chemins.get_dep(commune), commune))
        return commune.upper() in self._destinations()

    def supprimer_si_vide(self, log_fn) -> None:
        if self.prod_lien:
            log_fn("\nℹ️ Dossier « À déposer » sur SharePoint : jamais supprimé par l'outil (seuls les "
                   "dossiers commune rangés sont mis à la corbeille du site)", "#7d8aa5")
            return
        chemins.supprimer_si_vide(self.prod_texte, log_fn)

    # ── Préparation d'une commune ─────────────────────────────────
    def preparer(self, commune: str, ids=None, ppt_mode=None, excel=True, ppt=None) -> "Chantier":
        """Dossier de travail de la commune. ids / ppt_mode / excel : ce qu'il faut télécharger
        de SharePoint (mode par ID) ; ids=None = commune complète (rien à télécharger).
        ppt : PowerPoint (ppt_com.PowerPointALaDemande) pour un PPT ouvert sur SharePoint."""
        return Chantier(self, commune, ids, ppt_mode, excel, ppt)


class Chantier:
    """Une commune prête pour ranger_ids / ranger_commune : chemins locaux, ``publier``,
    ``maj_excel`` et ``finaliser`` (mise à la corbeille du dossier Prod IA distant)."""

    def __init__(self, sources: Sources, commune: str, ids, ppt_mode, excel, ppt=None):
        self.s, self.commune, self.ppt = sources, commune, ppt
        self._originaux: dict[str, str] = {}     # PPT téléchargé → copie intacte (zones à repérer)
        self.root_sp, self.prod_ia = sources.root_local, sources.prod_local
        self.dossier_src = sources.communes_prod().get(commune)
        self.etags: dict[str, str] = {}          # fichier local → ETag de la version téléchargée
        self._crees: set[str] = set()
        self._src_serveur = None
        self.telecharges = 0
        if sources.prod_lien and self.dossier_src:
            self._src_serveur = sources._prod_serveur[commune][0]
            shutil.rmtree(self.dossier_src, ignore_errors=True)
            self.telecharges += len(sources.client_prod.telecharger_dossier(self._src_serveur, self.dossier_src))
        if sources.root_lien:
            self._preparer_destination(ids, ppt_mode, excel)

    # Destination SharePoint par lien
    def _local_dst(self) -> str:
        return os.path.join(self.root_sp, chemins.get_dep(self.commune), self.commune)

    def _preparer_destination(self, ids, ppt_mode, excel):
        dest = self.s._destinations().get(self.commune.upper())
        if dest is None:
            return                                  # pas de dossier : ranger_* le signalera
        self._dst_serveur, fichiers = dest
        client = self.s.client_root
        local = self._local_dst()
        os.makedirs(os.path.join(local, PPT_FOLDER), exist_ok=True)
        if ids is None:
            return                                  # commune complète : on remplace, rien à lire
        nom_excel = f"{AUDIT_PREFIX}{self.commune}.xlsx"
        excel_local = os.path.join(local, nom_excel)
        if any(f["Name"].lower() == nom_excel.lower() for f in fichiers) and (excel or ppt_mode):
            self.etags[excel_local] = client.telecharger_version(f"{self._dst_serveur}/{nom_excel}", excel_local)
            self.telecharges += 1
        if ppt_mode != "modifier":
            return
        # PPT à modifier : nommés par ID erreur, lu dans l'Excel Prod IA puis SharePoint
        voulus = {nom_ppt(i).lower() for i in ids if cle_de(i) != COL_ID_IMM}
        imm = [i for i in ids if cle_de(i) == COL_ID_IMM]
        for chemin in (self._excel_src(), excel_local if os.path.isfile(excel_local) else None):
            if imm and chemin:
                try:
                    trouves = lire_ids_erreur(chemin, imm)
                except Exception:   # noqa: BLE001 — ranger_ids le signalera
                    trouves = {}
                voulus |= {nom_ppt(e).lower() for e in trouves.values() if e}
                imm = [i for i in imm if i.upper() not in trouves]
        if not voulus:
            return
        try:
            presents = client.fichiers(f"{self._dst_serveur}/{PPT_FOLDER}")
        except FileNotFoundError:
            presents = []
        for f in presents:
            if f["Name"].lower() in voulus:
                loc = os.path.join(local, PPT_FOLDER, f["Name"])
                self.etags[loc] = client.telecharger_version(f"{self._dst_serveur}/{PPT_FOLDER}/{f['Name']}", loc)
                orig = os.path.join(self.s._travail(), "originaux", self.commune, f["Name"])
                os.makedirs(os.path.dirname(orig), exist_ok=True)
                shutil.copyfile(loc, orig)
                self._originaux[loc] = orig
                self.telecharges += 1

    def _excel_src(self):
        if not self.dossier_src or not os.path.isdir(self.dossier_src):
            return None
        noms = [f for f in os.listdir(self.dossier_src) if f.lower().endswith(".xlsx") and not f.startswith("~")]
        voulu = f"{AUDIT_PREFIX}{self.commune}.xlsx"
        nom = voulu if voulu in noms else (noms[0] if noms else None)
        return os.path.join(self.dossier_src, nom) if nom else None

    def _serveur(self, local: str) -> str:
        rel = os.path.relpath(local, self._local_dst()).replace("\\", "/")
        if rel.startswith(".."):
            raise ValueError(f"{local} hors du dossier de la commune")
        return f"{self._dst_serveur}/{rel}"

    @property
    def publier(self):
        """Fonction d'envoi pour ranger_* (None si la racine est un dossier local)."""
        return self._publier if self.s.root_lien else None

    def _publier(self, local: str, op=None):
        """Envoie le fichier écrit. Retourne un message à journaliser (ou None).
        op : ce qui a été fait ({"mode": "remplacer", "src"} / {"mode": "modifier", "analyse", "lien"}),
        pour le refaire en co-édition PowerPoint si le PPT est ouvert sur SharePoint."""
        client, serveur = self.s.client_root, self._serveur(local)
        dossier, nom = serveur.rsplit("/", 1)
        with open(local, "rb") as f:
            contenu = f.read()
        try:
            etag = self.etags.pop(local, None)
            if etag:
                client.remplacer(serveur, contenu, etag)
            else:
                if dossier not in self._crees:
                    client.creer_dossier(dossier); self._crees.add(dossier)
                client.deposer(dossier, nom, contenu)
        except rest.FichierVerrouille as e:
            qui = client.verrou_par(serveur) or "un autre utilisateur"
            if local.lower().endswith(".pptx") and op and self.ppt is not None:
                return self._ppt_coedition(local, serveur, op, qui)
            raise PermissionError(f"ouvert par {qui} sur SharePoint ({e})") from None
        except rest.VersionModifiee:
            raise RuntimeError("modifié sur SharePoint par quelqu'un d'autre depuis son téléchargement : "
                               "non écrasé, rejouez cet ID") from None
        except rest.AuthentificationRequise as e:
            self.s.session_expiree = str(e)
            raise RuntimeError(_message_session(e)) from None

    def _ppt_coedition(self, local, serveur, op, qui):
        """PPT ouvert sur SharePoint : PowerPoint l'ouvre par son adresse (co-édition)."""
        adresse = self.s.client_root.url(serveur)
        if op["mode"] == "remplacer":
            r = self.ppt.remplacer(adresse, op["src"])
        else:
            modele = self._originaux.get(local)
            if modele is None:
                raise PermissionError(f"ouvert par {qui} sur SharePoint (version d'origine non téléchargée)")
            r = self.ppt.modifier(adresse, modele, op.get("analyse"), op.get("lien"))
            manque = [z for z, cle in (("Analyse", "analyse"), ("Lien Street View", "lien"))
                      if op.get(cle) is not None and not r.get(cle)]
            if manque:
                raise RuntimeError(f"co-édition PowerPoint : zone(s) introuvable(s) {', '.join(manque)}")
        if r.get("non_verifie"):
            raise RuntimeError(f"écrit en co-édition PowerPoint mais non confirmé à la relecture "
                               f"({', '.join(r['non_verifie'])})")
        return f"ouvert par {qui} sur SharePoint → écrit en co-édition PowerPoint ({r.get('ouverture')})"

    def maj_excel(self, session=None):
        """Fonction maj_excel(src, dst, ids, colonnes) pour ranger_ids.
        session : SessionEcriture (ou None : PPT seuls)."""
        if not self.s.root_lien:
            return session.maj_lignes if session else None
        return lambda src, dst, ids, colonnes=None: self._maj_excel_distant(session, src, dst, ids, colonnes)

    def _co_edition(self, session, src, serveur, ids, colonnes, qui=""):
        r = session.excel.maj_lignes(src, self.s.client_root.url(serveur), ids, colonnes)
        if qui and r.get("ouverture") and not r.get("bloque"):
            r["ouverture"] = f"{r['ouverture']} — fichier ouvert par {qui}, écrit en co-édition"
        return r

    def _maj_excel_distant(self, session, src, dst, ids, colonnes):
        client, serveur = self.s.client_root, self._serveur(dst)
        try:
            etat = client.etat_fichier(serveur)
            if etat.get("CheckOutType", 2) != 2:     # 2 = non extrait
                return self._bloque(src, ids, colonnes, f"{etat.get('CheckedOutBy') or 'un autre utilisateur'} "
                                                        "(fichier extrait sur SharePoint : archivez-le d'abord)")
            excel_ok = session is not None and getattr(session, "toujours_excel", False) and session._excel_pret()
            note = ""
            if excel_ok:                              # « Écrire via Excel » : co-édition directe
                try:
                    return self._co_edition(session, src, serveur, ids, colonnes)
                except (OngletIntrouvable, rest.AuthentificationRequise):
                    raise
                except Exception as e:   # noqa: BLE001 — Excel n'a pas pu CE fichier : l'API reste possible
                    note = f" — co-édition Excel impossible ({_court(e)}), envoi par l'API à la place"
            for essai in range(1, ESSAIS_VERSION + 1):
                etag = self.etags.pop(dst, None) if essai == 1 else None
                if not etag or not os.path.isfile(dst):
                    etag = client.telecharger_version(serveur, dst)
                r = maj_lignes_excel(src, dst, ids, colonnes)
                if not r["maj"]:
                    return r
                with open(dst, "rb") as f:
                    contenu = f.read()
                try:
                    client.remplacer(serveur, contenu, etag)
                except rest.VersionModifiee:
                    continue                          # enregistré entre-temps : on repart de la nouvelle version
                except rest.FichierVerrouille:
                    # Ouvert sur SharePoint (Excel, Excel Online, Teams) : SharePoint refuse de remplacer
                    # le fichier. Écriture cellule par cellule en co-édition, comme audit_validation_finale.
                    qui = client.verrou_par(serveur) or "un autre utilisateur"
                    raison = getattr(session, "raison_sans_excel", "") or "Excel indisponible"
                    if session is not None and not note and session._excel_pret():
                        try:
                            return self._co_edition(session, src, serveur, ids, colonnes, qui)
                        except (OngletIntrouvable, rest.AuthentificationRequise):
                            raise
                        except Exception as e:   # noqa: BLE001
                            raison = f"Excel n'a pas pu ouvrir le fichier : {_court(e)}"
                    elif note:
                        raison = note.split("(", 1)[-1].rsplit(")", 1)[0]
                    return self._bloque(src, ids, colonnes, f"{qui} (ouvert sur SharePoint) — l'écriture d'un fichier "
                                                            f"ouvert passe par la co-édition Excel, indisponible : {raison}")
                r["ouverture"] = (VOIE_API + (f", {essai}e essai (enregistré entre-temps par un autre)" if essai > 1 else "")
                                  + note)
                r["non_verifies"] = self._verifier_api(serveur, src, dst, r["maj"], colonnes)
                r["maj"] = [i for i in r["maj"] if i not in r["non_verifies"]]
                return r
            raise rest.VersionModifiee(f"{os.path.basename(dst)} enregistré en continu par quelqu'un d'autre : rejouez plus tard")
        except rest.AuthentificationRequise as e:
            self.s.session_expiree = str(e)
            raise RuntimeError(_message_session(e)) from None

    def _verifier_api(self, serveur, src, dst, ids, colonnes) -> list[str]:
        """Relit la version serveur après l'envoi : un ID n'est compté que si ses valeurs y sont."""
        if not ids:
            return []
        relu = dst + ".relu.xlsx"
        try:
            self.s.client_root.telecharger(serveur, relu)
            attendu = lire_cellules(src, ids, colonnes)["cellules"]
            obtenu = lire_cellules(relu, ids, colonnes)["cellules"]
        except Exception:   # noqa: BLE001 — envoi fait, relecture impossible : à rejouer par prudence
            return list(ids)
        finally:
            if os.path.isfile(relu):
                os.remove(relu)
        non = []
        for i in ids:
            for col, c in (attendu.get(i) or {}).items():
                r = (obtenu.get(i) or {}).get(col)
                if r is None or (isinstance(c.valeur, str) and c.valeur.startswith("=")):
                    continue                          # colonne absente côté SharePoint, ou formule recalculée
                if not _egal(r.valeur, c.valeur):
                    non.append(i)
                    break
        return non

    @staticmethod
    def _bloque(src, ids, colonnes, qui):
        lecture = lire_cellules(src, ids, colonnes)
        return {"maj": [], "absents_src": lecture["absents_src"], "absents_dst": [], "donnees": lecture["donnees"],
                "ids_src": lecture["ids_src"], "col_absentes": [], "id_erreur": lecture["id_erreur"],
                "onglets": {"src": lecture["onglet"], "dst": None}, "bloque": qui}

    def finaliser(self, log_fn) -> None:
        """Dossier Prod IA supprimé par le rangement (tout a réussi) → corbeille SharePoint."""
        if not self._src_serveur or os.path.isdir(self.dossier_src):
            return
        try:
            self.s.client_prod.recycler_dossier(self._src_serveur)
            log_fn(f"  ♻️  Dossier SharePoint « {self._src_serveur.rsplit('/', 1)[-1]} » mis à la corbeille du site "
                   "(restaurable)", "#ffb547")
        except rest.AuthentificationRequise as e:
            self.s.session_expiree = str(e)
            log_fn(f"  ⚠️ Dossier Prod IA/{self.commune} non mis à la corbeille : {_message_session(e)}", "#ffb547")
        except Exception as e:   # noqa: BLE001
            log_fn(f"  ⚠️ Dossier Prod IA/{self.commune} non mis à la corbeille sur SharePoint : {e}", "#ffb547")


def _court(e) -> str:
    return (str(e) or type(e).__name__).splitlines()[0][:160]


def _egal(a, b) -> bool:
    vide = (None, "")
    if a in vide and b in vide:
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool):
        return abs(a - b) < 1e-9
    return a == b


# ── Détection des colonnes ────────────────────────────────────────
def detecter_colonnes(prod: str, root: str, essais: int = 6):
    """Comme chemins.detecter_colonnes, en acceptant des liens SharePoint (les Excel sont
    téléchargés dans un dossier temporaire pour en lire les en-têtes)."""
    if not est_lien(prod) and not est_lien(root):
        return chemins.detecter_colonnes(prod, root)
    rapport = []
    try:
        sources = Sources(prod, root)
    except Exception as e:   # noqa: BLE001 — pas connecté, lien invalide
        return [], None, [f"SharePoint : {e}"]
    with sources, tempfile.TemporaryDirectory(prefix="dispatch_cols_") as tmp:
        for nom, lien, client in (("Prod IA", sources.prod_lien, sources.client_prod),
                                  ("SharePoint", sources.root_lien, sources.client_root)):
            texte = prod if nom == "Prod IA" else root
            if lien is None:
                cols, f, r = chemins.detecter_colonnes(texte, None) if nom == "Prod IA" else chemins.detecter_colonnes(None, texte)
                rapport += [l for l in r if "non renseigné" not in l and "(None)" not in l]
                if cols:
                    return cols, f, rapport
                continue
            try:
                candidats = _excels_distants(sources, nom, essais)
            except Exception as e:   # noqa: BLE001
                rapport.append(f"{nom} : lecture impossible ({e})"); continue
            if not candidats:
                rapport.append(f"{nom} : aucun fichier .xlsx trouvé ({lien.url})"); continue
            rapport.append(f"{nom} : {len(candidats)} fichier(s) Excel trouvé(s) sur SharePoint")
            for i, serveur in enumerate(candidats):
                local = os.path.join(tmp, f"{i}_{serveur.rsplit('/', 1)[-1]}")
                try:
                    client.telecharger(serveur, local)
                    cols = lire_entetes(local)
                except Exception as e:   # noqa: BLE001
                    rapport.append(f"  ✗ {serveur.rsplit('/', 1)[-1]} : {chemins._raison(e)}"); continue
                if COL_ID_IMM not in cols and COL_ID_ERR not in cols:
                    rapport.append(f"  ✗ {serveur.rsplit('/', 1)[-1]} : pas de colonne « {COL_ID_IMM} » ni "
                                   f"« {COL_ID_ERR} » en ligne 1 de l'onglet Audit")
                    continue
                return cols, client.url(serveur), rapport
    return [], None, rapport


def _excels_distants(sources: Sources, nom: str, essais: int) -> list[str]:
    def excel(n):
        return n.lower().endswith(".xlsx") and not n.startswith("~")
    if nom == "Prod IA":
        res = []
        for _com, (serveur, _d) in sorted(sources._prod_serveur_charge().items())[:essais]:
            res += [f"{serveur}/{f['Name']}" for f in sources.client_prod.fichiers(serveur) if excel(f["Name"])]
            if len(res) >= essais:
                break
        return res[:essais]
    res = [f"{serveur}/{f['Name']}" for _c, (serveur, fichiers) in sorted(sources._destinations().items())
           for f in fichiers if excel(f["Name"]) and f["Name"].lower().startswith(AUDIT_PREFIX.lower())]
    return res[:essais]

