"""Sources SharePoint par lien (connexion navigateur + API), contre un faux SharePoint local.

Le faux serveur expose, sur une arborescence de test, les appels de SharePoint Online utilisés
par l'outil (lecture, ETag, If-Match, dépôt, création de dossier, corbeille, verrou) et refuse
toute requête sans le cookie de session ``FedAuth`` (comme dans audit_validation_finale).
"""
import http.server
import json
import os
import re
import shutil
import threading
from datetime import datetime, timezone
from urllib.parse import unquote, urlsplit

import pytest
from openpyxl import Workbook, load_workbook

from dispatch_sna.services import chemins
from dispatch_sna.services.dispatcher import ranger_commune, ranger_ids
from dispatch_sna.services.sharepoint import connexion as cx
from dispatch_sna.services.sharepoint import rest
from dispatch_sna.services.sharepoint.sources import Sources, detecter_colonnes

from .fichiers import ppt_babayaga, ppt_bp, textes

HOTE = "contoso.sharepoint.com"
AUDIT = "/sites/SNA/Shared Documents/Audit"
DEPOT = "/sites/SNA/Shared Documents/Prod IA/A deposer"
LIEN_AUDIT = "https://contoso.sharepoint.com/sites/SNA/Shared%20Documents/Audit"
LIEN_DEPOT = "https://contoso.sharepoint.com/sites/SNA/Shared%20Documents/Prod%20IA/A%20deposer"
BIBLIO = "/sites/SNA/Shared Documents"
_APPEL = re.compile(r"^/sites/SNA/_api/web/(?P<res>GetFolderByServerRelativePath|GetFileByServerRelativePath)"
                    r"\(decodedurl='(?P<chemin>.*?)'\)(?P<suite>/[^?]*)?(?:\?(?P<q>.*))?$")
IMM, IMM2 = "ID-3806800000IW", "ID-3806800000AB"


class _FauxSharePoint(http.server.BaseHTTPRequestHandler):
    racine = ""
    versions: dict = {}
    verrouilles: set = set()
    extraits: dict = {}
    avant_ecriture = None
    ignorer_ecritures = False      # simule un serveur qui « perd » l'écriture (relecture)
    appels: list = []

    def log_message(self, *a):
        pass

    def _etag(self, local):
        return f'"{{AAAA}},{type(self).versions.get(local, 1)}"'

    def _local(self, chemin):
        if not chemin.startswith(BIBLIO):
            return None
        return os.path.join(self.racine, *[p for p in chemin[len(BIBLIO):].split("/") if p])

    def _envoyer(self, code, corps, type_="application/json", entetes=None):
        donnees = corps if isinstance(corps, bytes) else json.dumps(corps).encode()
        self.send_response(code)
        for k, v in (entetes or {}).items():
            self.send_header(k, v)
        self.send_header("Content-Type", type_)
        self.send_header("Content-Length", str(len(donnees)))
        self.end_headers()
        self.wfile.write(donnees)

    def _autorise(self):
        return "FedAuth=valide" in (self.headers.get("Cookie") or "")

    def _analyser(self):
        brut = urlsplit(self.path)
        m = _APPEL.match(unquote(brut.path) + ("?" + unquote(brut.query) if brut.query else ""))
        if not m:
            return None
        return m, self._local(m.group("chemin").replace("''", "'")), m.group("suite") or "", m.group("q") or ""

    def do_GET(self):
        type(self).appels.append("GET " + unquote(self.path))
        if not self._autorise():
            return self._envoyer(403, {"odata.error": {"message": {"value": "Accès refusé"}}})
        a = self._analyser()
        if not a:
            return self._envoyer(400, {})
        m, local, suite, q = a
        if m.group("res") == "GetFileByServerRelativePath":
            if not (local and os.path.isfile(local)):
                return self._envoyer(404, {"odata.error": {"message": {"value": "fichier introuvable"}}})
            if suite == "/$value":
                return self._envoyer(200, open(local, "rb").read(), "application/octet-stream",
                                     {"ETag": self._etag(local)})
            if suite == "/LockedByUser":
                return self._envoyer(200, {"Title": "Collègue Martin"} if local in type(self).verrouilles
                                     else {"odata.null": True})
            qui = type(self).extraits.get(local)
            return self._envoyer(200, {"ETag": self._etag(local), "CheckOutType": 0 if qui else 2,
                                       "CheckedOutByUser": {"Title": qui} if qui else None})
        if not local or not os.path.isdir(local):
            return self._envoyer(404, {"odata.error": {"message": {"value": "introuvable"}}})
        if suite == "":
            return self._envoyer(200, {"Exists": True})
        if suite == "/Files":
            return self._envoyer(200, {"value": _fichiers(local)})
        if suite == "/Folders":
            dossiers = [n for n in sorted(os.listdir(local)) if os.path.isdir(os.path.join(local, n))]
            if "$expand=Files" in q:
                return self._envoyer(200, {"value": [{"Name": n, "Files": _fichiers(os.path.join(local, n))}
                                                     for n in dossiers]})
            return self._envoyer(200, {"value": [{"Name": n} for n in dossiers]})
        return self._envoyer(400, {})

    def do_POST(self):
        cls = type(self)
        cls.appels.append("POST " + unquote(self.path))
        corps = self.rfile.read(int(self.headers.get("Content-Length") or 0))
        if not self._autorise():
            return self._envoyer(403, {"odata.error": {"message": {"value": "Accès refusé"}}})
        if self.path == "/sites/SNA/_api/contextinfo":
            return self._envoyer(200, {"FormDigestValue": "jeton-valide"})
        if self.headers.get("X-RequestDigest") != "jeton-valide":
            return self._envoyer(403, {"odata.error": {"message": {"value": "jeton invalide"}}})
        dossier = re.match(r"^/sites/SNA/_api/web/Folders/AddUsingPath\(decodedurl='(.*)'\)$", unquote(self.path))
        if dossier:
            os.makedirs(self._local(dossier.group(1)), exist_ok=True)
            return self._envoyer(200, {})
        a = self._analyser()
        if not a:
            return self._envoyer(400, {})
        m, local, suite, _q = a
        if suite == "/recycle()":
            corbeille = os.path.join(self.racine, "_corbeille")
            os.makedirs(corbeille, exist_ok=True)
            shutil.move(local, os.path.join(corbeille, os.path.basename(local)))
            return self._envoyer(200, {})
        ajout = re.match(r"^/Files/AddUsingPath\(decodedurl='(.*)',overwrite=true\)$", suite)
        if ajout:
            cible = os.path.join(local, ajout.group(1).replace("''", "'"))
            if cible in cls.verrouilles:
                return self._envoyer(423, {"odata.error": {"message": {"value": "The file is locked"}}})
            with open(cible, "wb") as f:
                f.write(corps)
            cls.versions[cible] = cls.versions.get(cible, 1) + 1
            return self._envoyer(200, {})
        if suite != "/$value" or self.headers.get("X-HTTP-Method") != "PUT":
            return self._envoyer(400, {})
        if cls.avant_ecriture:
            cls.avant_ecriture(local)
        if local in cls.verrouilles:
            return self._envoyer(423, {"odata.error": {"message": {"value": "The file is locked for shared use"}}})
        if self.headers.get("If-Match") not in ("*", self._etag(local)):
            return self._envoyer(412, {"odata.error": {"message": {"value": "version modifiée"}}})
        if cls.ignorer_ecritures:
            return self._envoyer(204, b"")
        with open(local, "wb") as f:
            f.write(corps)
        cls.versions[local] = cls.versions.get(local, 1) + 1
        return self._envoyer(204, b"")


def _fichiers(dossier):
    return [{"Name": n, "TimeLastModified": datetime.fromtimestamp(os.path.getmtime(os.path.join(dossier, n)),
                                                                  timezone.utc).isoformat().replace("+00:00", "Z")}
            for n in sorted(os.listdir(dossier)) if os.path.isfile(os.path.join(dossier, n))]


def _excel(chemin, lignes):
    wb = Workbook(); ws = wb.active; ws.title = "Audit"
    ws.append(["ID imm", "ID erreur", "Remarque", "Lien Street View", "Autre"])
    for l in lignes:
        ws.append(l)
    wb.save(chemin)


@pytest.fixture
def racine(tmp_path):
    """Bibliothèque « Shared Documents » : Audit/Dep38/38068 et Prod IA/A deposer/38068."""
    bib = tmp_path / "bib"
    sp = bib / "Audit" / "Dep38" / "38068"; (sp / "analyse").mkdir(parents=True)
    (bib / "Audit" / "Dep38" / "38070").mkdir(parents=True)         # commune sans Excel
    _excel(sp / "audit_38068.xlsx", [[IMM, "38068_1", "ancienne", None, "garde"],
                                    [IMM2, "38068_2", "autre", None, "garde"]])
    ppt_bp(sp / "analyse" / "cas_analyse_38068_1.pptx")
    pr = bib / "Prod IA" / "A deposer" / "38068"; pr.mkdir(parents=True)
    _excel(pr / "audit_38068.xlsx", [[IMM, "38068_1", "NOUVELLE", "https://sv/1", "x"]])
    ppt_babayaga(pr / "cas_analyse_38068_1.pptx")
    return str(bib)


@pytest.fixture
def serveur(racine):
    g = type("Faux", (_FauxSharePoint,), {"racine": racine, "versions": {}, "verrouilles": set(),
                                          "extraits": {}, "avant_ecriture": None, "appels": [],
                                          "ignorer_ecritures": False})
    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), g)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    yield srv, g
    srv.shutdown()


@pytest.fixture
def connecte(serveur):
    srv, g = serveur
    cx.enregistrer_client(rest.ClientSharePoint(HOTE, {"FedAuth": "valide"},
                                                base=f"http://127.0.0.1:{srv.server_port}"))
    yield g
    cx._CLIENTS.clear()


def _journal():
    lignes = []
    return lignes, (lambda m, c=None: lignes.append(m))


def _sp(racine, *p):
    return os.path.join(racine, "Audit", "Dep38", "38068", *p)


# ── Liens ─────────────────────────────────────────────────────────
@pytest.mark.parametrize("texte, chemin", [
    (LIEN_AUDIT, AUDIT),
    (LIEN_AUDIT + "/", AUDIT),
    ("https://contoso.sharepoint.com/sites/SNA/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FSNA%2FShared"
     "%20Documents%2FAudit&viewid=1", AUDIT),
    ("https://contoso.sharepoint.com/:f:/r/sites/SNA/Shared%20Documents/Audit?csf=1&web=1", AUDIT),
    (r"\\contoso.sharepoint.com@SSL\DavWWWRoot\sites\SNA\Shared Documents\Audit", AUDIT),
])
def test_analyser_lien(texte, chemin):
    lien = cx.analyser_lien(texte)
    assert (lien.hote, lien.chemin) == (HOTE, chemin)
    assert cx.est_lien(texte)


def test_lien_refuse_et_dossier_local():
    with pytest.raises(cx.LienInvalide):
        cx.analyser_lien("https://contoso.sharepoint.com/:f:/s/SNA/EaBcDeF")     # partage opaque
    assert not cx.est_lien(r"C:\Users\x\OneDrive\Audit")


def test_lien_sans_connexion_demande_de_se_connecter():
    cx._CLIENTS.clear()
    with pytest.raises(rest.AuthentificationRequise):
        Sources(LIEN_DEPOT, LIEN_AUDIT)
    assert not cx.est_connecte(LIEN_AUDIT)


# ── Connexion par le navigateur ───────────────────────────────────
class _FauxNavigateur:
    lances = []

    def __init__(self, url, visible=True, session_apres=None):
        self.url, self.visible, self.session_apres, self.releves = url, visible, session_apres, 0
        type(self).lances.append(self)

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def ouvert(self):
        return True

    def cookies(self):
        self.releves += 1
        ok = self.session_apres is not None and self.releves >= self.session_apres
        return [{"name": "FedAuth", "value": "valide" if ok else "anonyme", "domain": HOTE, "expires": -1}]


def _connecter(serveur, silencieux, visible, **k):
    srv, _ = serveur
    _FauxNavigateur.lances = []
    fab = lambda url, visible=True: _FauxNavigateur(url, visible, silencieux if not visible else visible_apres)
    visible_apres = visible
    return cx.connecter_navigateur(LIEN_AUDIT, fabrique=fab, delai_silencieux_s=1, delai_s=5,
                                   client_factory=lambda h, c: rest.ClientSharePoint(
                                       h, c, base=f"http://127.0.0.1:{srv.server_port}"), **k)


def test_connexion_silencieuse_puis_fenetre(serveur):
    cx._CLIENTS.clear()
    try:
        client = _connecter(serveur, silencieux=1, visible=None)
        assert [n.visible for n in _FauxNavigateur.lances] == [False]         # session du profil : sans fenêtre
        assert cx.client_pour(LIEN_DEPOT) is client                          # même hôte : même session
        cx._CLIENTS.clear()
        _connecter(serveur, silencieux=None, visible=2)
        assert [n.visible for n in _FauxNavigateur.lances] == [False, True]  # fenêtre pour s'identifier
        cx._CLIENTS.clear()
        with pytest.raises(rest.AuthentificationRequise):                    # au démarrage : jamais de fenêtre
            _connecter(serveur, silencieux=None, visible=1, interactif=False)
        assert [n.visible for n in _FauxNavigateur.lances] == [False]
    finally:
        cx._CLIENTS.clear()


# ── Aperçu, inventaire, colonnes ──────────────────────────────────
def test_apercu_et_colonnes_par_lien(connecte):
    with Sources(LIEN_DEPOT, LIEN_AUDIT) as s:
        a = chemins.calculer_apercu(True, f"{IMM}\nID-3807000000ZZ\nID-3899900000QQ", None, None,
                                    dispo=s.communes_prod(), existe=s.dest_existe)
    statuts = {l[0]: l[3] for l in a["lignes"]}
    assert statuts == {"38068": "⏳ À mettre à jour", "38070": "❌ Absent de Prod IA", "38999": "❌ Absent de Prod IA"}
    cols, fichier, _rapport = detecter_colonnes(LIEN_DEPOT, LIEN_AUDIT)
    assert cols[:3] == ["ID imm", "ID erreur", "Remarque"]
    assert fichier.startswith("https://contoso.sharepoint.com/sites/SNA/Shared%20Documents/Prod%20IA/")


# ── Mise à jour par ID imm ────────────────────────────────────────
def _ranger(s, ids, ppt_mode, colonnes, session=None, supprimer=True):
    lignes, log = _journal()
    ch = s.preparer("38068", ids, ppt_mode, excel=bool(colonnes))
    stats = ranger_ids(ch.root_sp, "38068", ch.dossier_src, ids, ppt_mode, supprimer, log, colonnes=colonnes,
                       maj_excel=ch.maj_excel(session), publier=ch.publier)
    ch.finaliser(log)
    return stats, lignes


def test_maj_par_id_deux_sources_par_lien(connecte, racine):
    with Sources(LIEN_DEPOT, LIEN_AUDIT) as s:
        stats, lignes = _ranger(s, [IMM], "modifier", ["Remarque", "Lien Street View"])
    assert stats["err"] == 0 and stats["lignes"] == 1 and stats["ppt"] == 1, lignes
    ws = load_workbook(_sp(racine, "audit_38068.xlsx")).active
    assert [c.value for c in ws[2]] == [IMM, "38068_1", "NOUVELLE", "https://sv/1", "garde"]
    assert ws["C3"].value == "autre"                                           # autre ID : intact
    assert any("NOUVELLE" in t for t in textes(_sp(racine, "analyse", "cas_analyse_38068_1.pptx")))
    # Tout a réussi : le dossier Prod IA SharePoint est à la corbeille (pas supprimé définitivement)
    assert not os.path.exists(os.path.join(racine, "Prod IA", "A deposer", "38068"))
    assert os.path.isdir(os.path.join(racine, "_corbeille", "38068"))
    assert any("écrit par l'API SharePoint" in l for l in lignes)
    # Remplacement conditionnel : toujours avec l'ETag de la version lue, jamais « * »
    assert any(a.startswith("POST") and "audit_38068.xlsx" in a for a in connecte.appels)


def test_racine_par_lien_prod_ia_local(connecte, racine, tmp_path):
    local = tmp_path / "depot"
    shutil.copytree(os.path.join(racine, "Prod IA", "A deposer"), local)
    with Sources(str(local), LIEN_AUDIT) as s:
        stats, _ = _ranger(s, [IMM], "remplacer", ["Remarque"])
    assert stats["err"] == 0 and stats["ppt"] == 1
    assert load_workbook(_sp(racine, "audit_38068.xlsx")).active["C2"].value == "NOUVELLE"
    assert not (local / "38068").exists()                     # dossier local supprimé comme avant


def test_version_modifiee_entre_temps_on_repart_de_la_nouvelle(connecte, racine):
    fait = []

    def collegue(local):                                      # enregistre juste avant notre 1re écriture
        if local.endswith("audit_38068.xlsx") and not fait:
            fait.append(1)
            wb = load_workbook(local); wb.active["C3"] = "par Martin"; wb.save(local)
            connecte.versions[local] = connecte.versions.get(local, 1) + 1
    connecte.avant_ecriture = staticmethod(collegue)
    with Sources(LIEN_DEPOT, LIEN_AUDIT) as s:
        stats, lignes = _ranger(s, [IMM], None, ["Remarque"])
    ws = load_workbook(_sp(racine, "audit_38068.xlsx")).active
    assert stats["err"] == 0
    assert (ws["C2"].value, ws["C3"].value) == ("NOUVELLE", "par Martin")   # rien de perdu
    assert any("2e essai" in l for l in lignes)


def test_excel_ouvert_sans_excel_rien_ecrit_et_en_attente(connecte, racine):
    connecte.verrouilles.add(_sp(racine, "audit_38068.xlsx"))
    with Sources(LIEN_DEPOT, LIEN_AUDIT) as s:
        stats, lignes = _ranger(s, [IMM], None, ["Remarque"])
    assert "Collègue Martin" in stats["echecs"][IMM]
    assert load_workbook(_sp(racine, "audit_38068.xlsx")).active["C2"].value == "ancienne"
    assert os.path.isdir(os.path.join(racine, "Prod IA", "A deposer", "38068"))    # conservé


def test_fichier_extrait_bloque(connecte, racine):
    connecte.extraits[_sp(racine, "audit_38068.xlsx")] = "Julie"
    with Sources(LIEN_DEPOT, LIEN_AUDIT) as s:
        stats, _ = _ranger(s, [IMM], None, ["Remarque"])
    assert "Julie" in stats["echecs"][IMM] and "extrait" in stats["echecs"][IMM]


def test_ppt_ouvert_sur_sharepoint(connecte, racine):
    connecte.verrouilles.add(_sp(racine, "analyse", "cas_analyse_38068_1.pptx"))
    with Sources(LIEN_DEPOT, LIEN_AUDIT) as s:
        stats, _ = _ranger(s, [IMM], "remplacer", [])
    assert "Collègue Martin" in stats["echecs"][IMM]
    assert os.path.isdir(os.path.join(racine, "Prod IA", "A deposer", "38068"))


def test_session_expiree_signalee(connecte, racine):
    with Sources(LIEN_DEPOT, LIEN_AUDIT) as s:
        s.client_root.cookies = {"FedAuth": "expiree"}
        with pytest.raises(rest.AuthentificationRequise):
            s.preparer("38068", [IMM], None, excel=True)


# ── Commune complète ──────────────────────────────────────────────
def test_commune_complete_par_lien(connecte, racine):
    shutil.rmtree(_sp(racine, "analyse"))                     # dossier analyse créé au besoin
    lignes, log = _journal()
    with Sources(LIEN_DEPOT, LIEN_AUDIT) as s:
        ch = s.preparer("38068")
        stats = ranger_commune(ch.prod_ia, ch.root_sp, "38068", ch.dossier_src, True, log, publier=ch.publier)
        ch.finaliser(log)
    assert stats == {"excel": 1, "ppt": 1, "err": 0}, lignes
    assert load_workbook(_sp(racine, "audit_38068.xlsx")).active["C2"].value == "NOUVELLE"
    assert os.path.isfile(_sp(racine, "analyse", "cas_analyse_38068_1.pptx"))
    assert os.path.isdir(os.path.join(racine, "_corbeille", "38068"))


# ── Fichier ouvert sur SharePoint : co-édition (comme audit_validation_finale) ──
class _FauxExcel:
    """Session Excel simulée : la co-édition écrit cellule par cellule DANS le fichier du serveur."""

    def __init__(self, racine, echoue=None):
        self.racine, self.echoue, self.adresses = racine, echoue, []

    def maj_lignes(self, src, adresse, ids, colonnes=None):
        from dispatch_sna.services.excel_rows import maj_lignes_excel
        self.adresses.append(adresse)
        if self.echoue:
            raise RuntimeError(self.echoue)
        chemin = unquote(urlsplit(adresse).path)[len(BIBLIO):]
        r = maj_lignes_excel(src, os.path.join(self.racine, *chemin.strip("/").split("/")), ids, colonnes)
        r["ouverture"] = "co-édition SharePoint"
        return r


class _FausseEcriture:
    """Même interface que ecriture.SessionEcriture pour Chantier.maj_excel."""

    def __init__(self, excel, toujours=False):
        self.excel, self.toujours_excel, self.raison_sans_excel = excel, toujours, ""

    def _excel_pret(self):
        return self.excel is not None


def test_excel_ouvert_sur_sharepoint_ecrit_en_co_edition(connecte, racine):
    connecte.verrouilles.add(_sp(racine, "audit_38068.xlsx"))           # ouvert par Collègue Martin
    xl = _FauxExcel(racine)
    with Sources(LIEN_DEPOT, LIEN_AUDIT) as s:
        stats, lignes = _ranger(s, [IMM], None, ["Remarque"], session=_FausseEcriture(xl))
    assert stats["err"] == 0 and stats["lignes"] == 1, lignes
    assert xl.adresses == [LIEN_AUDIT + "/Dep38/38068/audit_38068.xlsx"]   # ouvert par son adresse
    assert load_workbook(_sp(racine, "audit_38068.xlsx")).active["C2"].value == "NOUVELLE"
    assert any("Collègue Martin" in l and "co-édition" in l for l in lignes)


def test_excel_coche_mais_echec_de_co_edition_repli_api(connecte, racine):
    xl = _FauxExcel(racine, echoue="adresse refusée par Office")
    with Sources(LIEN_DEPOT, LIEN_AUDIT) as s:
        stats, lignes = _ranger(s, [IMM], None, ["Remarque"], session=_FausseEcriture(xl, toujours=True))
    assert stats["err"] == 0 and stats["lignes"] == 1
    assert load_workbook(_sp(racine, "audit_38068.xlsx")).active["C2"].value == "NOUVELLE"
    assert any("co-édition Excel impossible (adresse refusée par Office)" in l for l in lignes)


def test_envoi_api_relu_sur_le_serveur(connecte, racine):
    connecte.ignorer_ecritures = True                       # le serveur « perd » l'écriture
    with Sources(LIEN_DEPOT, LIEN_AUDIT) as s:
        stats, _ = _ranger(s, [IMM], None, ["Remarque"])
    assert stats["lignes"] == 0 and "non confirmées" in stats["echecs"][IMM]


class _FauxPowerPoint:
    def __init__(self):
        self.appels = []

    def remplacer(self, adresse, src):
        self.appels.append(("remplacer", adresse, os.path.basename(src)))
        return {"ouverture": "co-édition SharePoint", "non_verifie": []}

    def modifier(self, adresse, modele, analyse, lien):
        self.appels.append(("modifier", adresse, textes(modele), analyse, lien))
        return {"analyse": True, "lien": True, "ouverture": "co-édition SharePoint", "non_verifie": []}


@pytest.mark.parametrize("mode", ["remplacer", "modifier"])
def test_ppt_ouvert_sur_sharepoint_ecrit_en_co_edition_powerpoint(connecte, racine, mode):
    connecte.verrouilles.add(_sp(racine, "analyse", "cas_analyse_38068_1.pptx"))
    pp = _FauxPowerPoint()
    with Sources(LIEN_DEPOT, LIEN_AUDIT) as s:
        lignes, log = _journal()
        ch = s.preparer("38068", [IMM], mode, excel=False, ppt=pp)
        stats = ranger_ids(ch.root_sp, "38068", ch.dossier_src, [IMM], mode, True, log, colonnes=["Remarque"]
                           if mode == "modifier" else [], publier=ch.publier)
    assert stats["err"] == 0 and stats["ppt"] == 1, lignes
    adresse = LIEN_AUDIT + "/Dep38/38068/analyse/cas_analyse_38068_1.pptx"
    assert pp.appels[0][:2] == (mode, adresse)
    if mode == "modifier":           # zones repérées sur la version d'ORIGINE du serveur
        assert "Ancienne analyse" in pp.appels[0][2] and pp.appels[0][3] == "NOUVELLE"
    assert any("co-édition PowerPoint" in l and "Collègue Martin" in l for l in lignes)


def test_ppt_ouvert_sans_powerpoint_reste_en_attente(connecte, racine):
    connecte.verrouilles.add(_sp(racine, "analyse", "cas_analyse_38068_1.pptx"))
    with Sources(LIEN_DEPOT, LIEN_AUDIT) as s:
        stats, _ = _ranger(s, [IMM], "remplacer", [])
    assert "Collègue Martin" in stats["echecs"][IMM]


def test_ppt_local_ouvert_ecrit_en_co_edition_powerpoint(tmp_path):
    """Dossier local (OneDrive) : PPT ouvert (verrou ~$) → PowerPoint, plus « non modifié »."""
    from tests.fichiers import arbo
    root, src, dst = arbo(tmp_path)
    analyse = dst / "analyse"
    (analyse / "~$cas_analyse_74143_1.pptx").write_bytes(b"\x05Julie" + b"\x00" * 50)
    pp = _FauxPowerPoint()
    lignes, log = _journal()
    st = ranger_ids(str(root), "74143", str(src), ["74143_1"], "modifier", False, log,
                    colonnes=[], ppt_coedition=pp)
    assert st["err"] == 0 and st["ppt"] == 1, lignes
    assert pp.appels and pp.appels[0][0] == "modifier"
    assert any("ouvert par Julie" in l and "co-édition PowerPoint" in l for l in lignes)
