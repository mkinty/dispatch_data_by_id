"""PPT cas_analyse : réécriture des zones « Analyse » et « Lien Street View »
par édition XML directe (sans python-pptx), et copie avec remplacement."""
from __future__ import annotations

import os
import re
import shutil
import zipfile
from html import unescape

_SHAPE = re.compile(r"<p:sp>.*?</p:sp>", re.S)
_TXBODY = re.compile(r"(<p:txBody>)(.*?)(</p:txBody>)", re.S)
_PARA = re.compile(r"<a:p(?:\s[^>]*)?>.*?</a:p>", re.S)   # n'attrape pas <a:p/> ni <a:pPr>
_RUN = re.compile(r"<a:r>.*?</a:r>", re.S)
_T = re.compile(r"(<a:t(?:\s[^>]*)?>)(.*?)(</a:t>)", re.S)
_END_RPR = re.compile(r"<a:endParaRPr\b[^>]*?(?:/>|>.*?</a:endParaRPr>)", re.S)
_HLINK = re.compile(r'<a:hlinkClick\b[^>]*\br:id="([^"]+)"')

# Libellés fixes du template : jamais la zone « Analyse »
_LIBELLES = re.compile(r"^(analyses\s*:|lien street view|sna nombre|ipe nombre|type de probl"
                       r"|\{\{photo|photo$)", re.I)
_LABEL_LIEN = "Lien Street View : "


def _esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _texte(xml: str) -> str:
    return unescape("".join(m.group(2) for m in _T.finditer(xml)))


def _est_libelle(texte: str) -> bool:
    t = texte.strip()
    return bool(_LIBELLES.match(t)) or "adresse au" in t.lower()


def _set_para(para: str, texte: str) -> str:
    """Met `texte` dans le 1er fragment <a:t> du paragraphe, vide les autres."""
    if _T.search(para):
        n = {"i": 0}
        def rep(m):
            n["i"] += 1
            return f"{m.group(1)}{_esc(texte) if n['i'] == 1 else ''}{m.group(3)}"
        return _T.sub(rep, para)
    # Paragraphe vide : on crée un run avec la mise en forme de fin de paragraphe
    m = _END_RPR.search(para)
    rpr = m.group(0).replace("a:endParaRPr", "a:rPr") if m else ""
    run = f"<a:r>{rpr}<a:t>{_esc(texte)}</a:t></a:r>"
    pos = m.start() if m else para.rindex("</a:p>")
    return para[:pos] + run + para[pos:]


def _remplir_analyse(sp: str, texte: str) -> str:
    """Remplace le texte d'une zone : 1ère ligne dans le 1er paragraphe écrit,
    lignes suivantes en paragraphes clonés ; anciens paragraphes texte retirés,
    paragraphes vides (espacement) conservés."""
    mb = _TXBODY.search(sp)
    if not mb:
        return sp
    body = mb.group(2)
    paras = list(_PARA.finditer(body))
    if not paras:
        return sp
    avec_texte = [p for p in paras if _texte(p.group(0)).strip()]
    modele = (avec_texte or paras)[0]
    lignes = texte.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    nouveau = "".join(_set_para(modele.group(0), l) for l in lignes)
    out, pos = [], 0
    for p in paras:
        out.append(body[pos:p.start()])
        if p is modele:
            out.append(nouveau)
        elif p not in avec_texte:
            out.append(p.group(0))
        pos = p.end()
    out.append(body[pos:])
    body2 = "".join(out)
    return sp[:mb.start(2)] + body2 + sp[mb.end(2):]


def _remplir_lien(para: str, lien: str, rels: str) -> tuple[str, str]:
    """Réécrit « Lien Street View : <lien> ». Si un run porte déjà un
    hyperlien, le libellé va dans le 1er run et l'URL dans ce run (cible mise à jour)."""
    runs = list(_RUN.finditer(para))
    lien_run = next((r for r in runs[1:] if _HLINK.search(r.group(0))), None)
    if lien_run is None:
        m = _HLINK.search(runs[0].group(0)) if runs else None
        if m:
            rels = _maj_rel(rels, m.group(1), lien)
        return _set_para(para, f"{_LABEL_LIEN}{lien}".rstrip()), rels
    rels = _maj_rel(rels, _HLINK.search(lien_run.group(0)).group(1), lien)
    out, pos = [], 0
    for r in runs:
        txt = _LABEL_LIEN if r is runs[0] else (lien if r is lien_run else "")
        out.append(para[pos:r.start()])
        out.append(_T.sub(lambda m: f"{m.group(1)}{_esc(txt)}{m.group(3)}", r.group(0), count=1))
        pos = r.end()
    out.append(para[pos:])
    return "".join(out), rels


def _maj_rel(rels: str, rid: str, cible: str) -> str:
    if not cible:
        return rels
    return re.sub(rf'(<Relationship\b[^>]*\bId="{re.escape(rid)}"[^>]*\bTarget=")[^"]*(")',
                  lambda m: f"{m.group(1)}{_esc(cible)}{m.group(2)}", rels, count=1)


def _modifier_slide(xml: str, rels: str, analyse: str | None, lien: str | None):
    """Retourne (xml, rels, analyse_ok, lien_ok)."""
    shapes = list(_SHAPE.finditer(xml))
    textes = [_texte(s.group(0)) for s in shapes]
    a_ok = l_ok = False
    remplacements = {}   # index shape → nouveau XML

    # Lien Street View : paragraphe commençant par « Lien Street View »
    i_lien = None
    if lien is not None:
        for i, s in enumerate(shapes):
            sp = s.group(0); changed = False
            for p in list(_PARA.finditer(sp))[::-1]:
                if _texte(p.group(0)).strip().lower().startswith("lien street view"):
                    np_, rels = _remplir_lien(p.group(0), lien, rels)
                    sp = sp[:p.start()] + np_ + sp[p.end():]; changed = True
            if changed:
                remplacements[i] = sp; l_ok = True
                if i_lien is None: i_lien = i
    else:
        i_lien = next((i for i, t in enumerate(textes)
                       if t.strip().lower().startswith("lien street view")), None)

    # Analyse : uniquement sur la slide qui porte le libellé « Analyses: »
    if analyse is not None and any(t.strip().lower().startswith("analyses") for t in textes):
        cible = next((i for i, t in enumerate(textes) if t.strip() and not _est_libelle(t)), None)
        if cible is None and i_lien is not None:   # zone vide : celle juste avant le lien
            j = i_lien - 1
            if j >= 0 and not _est_libelle(textes[j]) and _TXBODY.search(shapes[j].group(0)):
                cible = j
        if cible is not None:
            remplacements[cible] = _remplir_analyse(remplacements.get(cible, shapes[cible].group(0)), analyse)
            a_ok = True

    if remplacements:
        out, pos = [], 0
        for i, s in enumerate(shapes):
            out.append(xml[pos:s.start()]); out.append(remplacements.get(i, s.group(0))); pos = s.end()
        out.append(xml[pos:]); xml = "".join(out)
    return xml, rels, a_ok, l_ok


def modifier_ppt(chemin: str, analyse: str | None, lien: str | None) -> dict:
    """Réécrit Analyse et/ou Lien Street View dans le PPT (None = ne pas toucher).

    Retourne {"analyse": bool, "lien": bool} — ce qui a effectivement été trouvé
    et modifié. Écriture via fichier temporaire. Lève PermissionError si verrouillé.
    """
    slide_re = re.compile(r"^ppt/slides/slide\d+\.xml$")
    nouveaux, res = {}, {"analyse": False, "lien": False}
    with zipfile.ZipFile(chemin) as z:
        noms = z.namelist()
        for nom in sorted(n for n in noms if slide_re.match(n)):
            rel_nom = nom.replace("slides/", "slides/_rels/") + ".rels"
            xml = z.read(nom).decode("utf-8")
            rels = z.read(rel_nom).decode("utf-8") if rel_nom in noms else ""
            xml2, rels2, a_ok, l_ok = _modifier_slide(
                xml, rels, None if res["analyse"] else analyse, lien)
            res["analyse"] |= a_ok; res["lien"] |= l_ok
            if xml2 != xml: nouveaux[nom] = xml2
            if rels2 != rels: nouveaux[rel_nom] = rels2
    if not nouveaux:
        return res
    tmp = chemin + ".tmp.pptx"
    try:
        with zipfile.ZipFile(chemin) as zi, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for item in zi.infolist():
                data = nouveaux[item.filename].encode("utf-8") if item.filename in nouveaux \
                    else zi.read(item.filename)
                zo.writestr(item, data)
        os.replace(tmp, chemin)
    finally:
        if os.path.isfile(tmp):
            os.remove(tmp)
    return res


def copier_ppt(src: str, dst: str):
    tmp = dst + ".tmp"
    shutil.copy2(src, tmp)
    os.replace(tmp, dst)
