from dispatch_sna.services import attente
from dispatch_sna.services.dispatcher import ranger_ids
from dispatch_sna.services.onedrive import url_sharepoint
from tests.fichiers import arbo, log


def test_attente_ajoute_remplace_et_retire(tmp_path):
    f = str(tmp_path / "attente.json")
    assert attente.charger(f) == {}
    attente.mettre_a_jour(f, {"74143_1": "verrouillé", "74143_2": "absent"}, [])
    d = attente.mettre_a_jour(f, {"74143_2": "toujours absent"}, ["74143_1"])
    assert list(d) == ["74143_2"] and d["74143_2"]["raison"] == "toujours absent"
    assert attente.charger(f) == d
    attente.vider(f)
    assert attente.charger(f) == {}


def test_attente_fichier_corrompu(tmp_path):
    f = tmp_path / "attente.json"; f.write_text("pas du json", encoding="utf-8")
    assert attente.charger(str(f)) == {}


def test_url_sharepoint():
    montages = [(r"C:\Users\u1\Altice\Swap Adresse - Etude Cible",
                 "https://sfr.sharepoint.com/sites/Swap/Documents partages"),
                (r"C:\Users\u1\Altice", "https://sfr.sharepoint.com/sites/Racine/Docs")]
    # le montage le plus précis l'emporte, espaces encodés dans la partie relative
    assert url_sharepoint(r"C:\Users\u1\Altice\Swap Adresse - Etude Cible\audit_SNA\Dep04\04121\audit_04121.xlsx",
                          montages) == ("https://sfr.sharepoint.com/sites/Swap/Documents partages/"
                                        "audit_SNA/Dep04/04121/audit_04121.xlsx")
    assert url_sharepoint(r"c:\users\U1\altice\Autre dossier\a.xlsx", montages) == \
        "https://sfr.sharepoint.com/sites/Racine/Docs/Autre%20dossier/a.xlsx"
    assert url_sharepoint(r"D:\local\a.xlsx", montages) is None
    assert url_sharepoint(r"C:\Users\u1\AlticeBis\a.xlsx", montages) is None   # préfixe ≠ dossier parent


def test_ranger_ids_remonte_les_echecs_par_id(tmp_path):
    root, src_dir, _ = arbo(tmp_path)
    st = ranger_ids(str(root), "74143", str(src_dir), ["74143_1", "74143_9"], "modifier", True, log)
    assert set(st["echecs"]) == {"74143_9"}
    assert "absent de l'Excel Prod IA" in st["echecs"]["74143_9"]
    assert src_dir.exists()                          # échec → dossier Prod IA conservé pour rejouer


def test_ranger_ids_excel_bloque_met_tout_en_attente(tmp_path):
    root, src_dir, _ = arbo(tmp_path)

    def maj_bloquee(src, dst, ids, colonnes):
        return {"maj": [], "absents_src": [], "absents_dst": [], "donnees": {}, "ids_src": ["74143_1"],
                "col_absentes": [], "bloque": "Jean Dupont", "non_verifies": [], "ouverture": "co-édition"}

    st = ranger_ids(str(root), "74143", str(src_dir), ["74143_1"], None, True, log, maj_excel=maj_bloquee)
    assert st["lignes"] == 0 and st["err"] == 1
    assert st["echecs"] == {"74143_1": "Excel ouvert en mode exclusif par Jean Dupont"}
    assert src_dir.exists()
