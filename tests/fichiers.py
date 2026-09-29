"""Fabriques de fichiers de test : Excel audit, PPT (structures bp_app / Babayaga), arborescence."""
import re
import zipfile

from openpyxl import Workbook

ENTETES = ["ID erreur", "Adresse", "Remarque", "Lien Street View", "Calcul"]


def excel(path, lignes, entetes=ENTETES, onglets=("Audit",)):
    wb = Workbook(); ws = wb.active; ws.title = onglets[0]
    for autre in onglets[1:]:
        wb.create_sheet(autre)["A1"] = "garde-moi"
    ws.append(entetes)
    for l in lignes:
        ws.append(l)
    wb.save(path)
    return path


def sp(nom, paras):
    ps = "".join(f"<a:p>{p}</a:p>" for p in paras)
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="1" name="{nom}"/></p:nvSpPr>'
            f"<p:txBody><a:bodyPr/><a:lstStyle/>{ps}</p:txBody></p:sp>")


def run(t, extra=""):
    return f'<a:r><a:rPr lang="fr-FR" sz="1000">{extra}</a:rPr><a:t>{t}</a:t></a:r>'


def pptx(path, shapes, rels=""):
    xml = ('<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="a" xmlns:p="p" xmlns:r="r">'
           f"<p:cSld><p:spTree>{''.join(shapes)}</p:spTree></p:cSld></p:sld>")
    rels = rels or '<?xml version="1.0"?><Relationships></Relationships>'
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("ppt/slides/slide1.xml", xml)
        z.writestr("ppt/slides/_rels/slide1.xml.rels", rels)
    return path


def textes(path):
    xml = zipfile.ZipFile(path).read("ppt/slides/slide1.xml").decode()
    return [''.join(re.findall(r"<a:t[^>]*>(.*?)</a:t>", p, re.S))
            for p in re.findall(r"<a:p>.*?</a:p>", xml, re.S)]


def ppt_bp(path):
    """Structure du template bp_app (zone « Analyse » = Rectangle 8)."""
    return pptx(path, [
        sp("ZoneTexte 5", [run("Analyses:")]),
        sp("Rectangle 8", [run("Ancienne analyse"), "<a:endParaRPr/>", ""]),
        sp("Titre 2", [run("74143 : adresse au 1 rue X (Etat IPE: OK)")]),
        sp("ZoneTexte 3", [run("SNA nombre de logement: 2")]),
        sp("ZoneTexte 7", [run("Lien Street View : ")]),
        sp("TextBox 6", [run("")]),
    ])


def ppt_babayaga(path):
    """Structure Babayaga : lien éclaté en plusieurs runs + run hyperlien."""
    hl = '<a:hlinkClick r:id="rId5"/>'
    rels = ('<?xml version="1.0"?><Relationships><Relationship Id="rId5" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            'Target="https://old.example" TargetMode="External"/></Relationships>')
    return pptx(path, [
        sp("Titre 2", [run("(2 - Maison): adresse au 3 rue Y (Etat IPE: OK)")]),
        sp("ZoneTexte 4", [run("IPE nombre de logement sur l'adresse: 1")]),
        sp("ZoneTexte 5", [run("Analyses:")]),
        sp("Rectangle 8", [""]),
        sp("ZoneTexte 6", [run("   \t\t"), "", ""]),
        sp("ZoneTexte 7", [run("Justification IA")]),
        sp("ZoneTexte 7", [run("Lien Street ") + run("View") + run(" : ")
                           + run("https://old.example", hl)]),
    ], rels)


def arbo(tmp_path):
    """SharePoint (Dep74/74143 avec 2 lignes + 1 PPT) et Prod IA (74143 avec l'ID 74143_1)."""
    root = tmp_path / "sp"; prod = tmp_path / "prod"
    dst_dir = root / "Dep74" / "74143"; (dst_dir / "analyse").mkdir(parents=True)
    src_dir = prod / "74143"; src_dir.mkdir(parents=True)
    excel(dst_dir / "audit_74143.xlsx", [["74143_1", "A1", "vierge", None, None],
                                         ["74143_2", "A2", "vierge", None, None]])
    excel(src_dir / "audit_74143.xlsx", [["74143_1", "A1", "Rem 1", "https://sv/1", None]])
    ppt_bp(dst_dir / "analyse" / "cas_analyse_74143_1.pptx")
    ppt_babayaga(src_dir / "cas_analyse_74143_1.pptx")
    return root, src_dir, dst_dir


def log(*a):
    pass
