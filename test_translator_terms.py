from translator import translate


def test_standardized_elevator_terms():
    expected = {
        "轿厢上梁": "Car Top Beam",
        "轿架下梁": "Car Bottom Beam",
        "主机支架梁": "Machine Support Beam",
        "轿厢支撑梁": "Car Support Beam",
        "对重支撑梁": "Counterweight Support Beam",
        "钢带轮": "Belt Sheave",
        "夹30绳，4槽": "Clamp for 30 mm Rope, 4 Grooves",
        "夹30绳,4槽": "Clamp for 30 mm Rope, 4 Grooves",
        "含弹簧及紧固件": "Including Springs and Fasteners",
        "对称件": "Symmetrical Part",
        "焊接件": "Welded Part",
        "色": "Color",
        "残疾人操纵箱": "Disabled Control Panel",
        "Side Wall Panel（后侧）": "Rear Side Wall Panel",
        "Side Wall Panel(后侧)": "Rear Side Wall Panel",
    }

    for source, target in expected.items():
        assert translate(source) == target


def test_color_phrase_is_not_partially_translated():
    assert translate("颜色：RAL 7035") == "Color: RAL 7035"
