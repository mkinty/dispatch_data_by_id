from dispatch_sna.services.error_ids import extraire_ids, nom_ppt


def test_extraire_ids_groupe_et_dedoublonne():
    txt = "74143_1, 74143_2\n38068_207 74143_1 2a004_12 foo 123_4"
    assert extraire_ids(txt) == {"74143": ["74143_1", "74143_2"], "38068": ["38068_207"],
                                 "2A004": ["2A004_12"]}


def test_extraire_ids_un_par_ligne():
    assert extraire_ids("04121_93\r\n04121_97\n\n04121_108\n") == {
        "04121": ["04121_93", "04121_97", "04121_108"]}


def test_nom_ppt():
    assert nom_ppt("74143_1") == "cas_analyse_74143_1.pptx"
