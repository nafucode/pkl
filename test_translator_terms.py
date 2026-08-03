from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook

from translator import scan_source_for_quality_findings, translate


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
        "轿厢侧支撑梁": "Car Side Support Beam",
        "对重侧支撑梁": "Counterweight Side Support Beam",
        "绳头支撑件": "Rope End Support Bracket",
        "侧中壁": "Side Center Wall Panel",
        "后侧壁": "Rear Side Wall Panel",
        "后Side Wall Panel": "Rear Side Wall Panel",
        "后中壁": "Rear Center Wall Panel",
        "扶手": "Handrail",
        "上下围板": "Upper and Lower Skirting Panels",
        "观光Glass橡胶垫": "Glass Rubber Pad for Observation Car",
        "千里马Glass胶及枪": "Glass Adhesive and Gun (Qianlima Brand)",
        "控制柜箱": "Control Cabinet Wooden Case",
        "电气控制": "Electrical Control System",
        "轿壁": "Car Wall Panel",
        "门套": "Door Jamb",
        "层站召唤盒": "Landing Call Panel",
        "轿厢架上梁": "Car Sling Upper Beam",
        "轿厢架下梁": "Car Sling Lower Beam",
        "轿厢反绳轮": "Car Deflector Sheave",
        "轿厢Deflector Sheave": "Car Deflector Sheave",
        "爬梯": "Ladder",
        "侧梁": "Side Beam",
        "支架导靴": "Bracket Guide Shoe",
        "含提拉机构、导靴板": "with lifting mechanism and guide shoe plate",
        "绳头组合": "Wedge Rope Socket Assembly",
        "钢丝绳夹": "Wire Rope Clip",
        "导轨支架底座": "Guide Rail Bracket Base",
        "张紧轮": "Tension Pulley",
        "终端限位开关": "Final Limit Switch",
        "技术文件": "Technical Documents",
    }

    for source, target in expected.items():
        assert translate(source) == target


def test_color_phrase_is_not_partially_translated():
    assert translate("颜色：RAL 7035") == "Color: RAL 7035"


def test_standardized_english_replacements():
    assert translate("Bare Packing") == "Bare Package"
    assert translate("Soft Packing") == "Soft Package"
    assert translate("Contral Cabinet") == "Control Cabinet"
    assert translate("轿壁箱") == "Car Wall Panel Wooden Case"


def test_quality_check_flags_missing_number_and_unit():
    with TemporaryDirectory() as temp_dir:
        path = Path(temp_dir) / "packing-list.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "装箱单"
        sheet["A25"] = "装箱清单"
        sheet["A30"] = 10
        sheet["C30"] = "Overspeed Governor"
        sheet["F30"] = 1
        sheet["G30"] = "unit"
        sheet["A31"] = 12
        sheet["C31"] = "Rope Rod"
        sheet["F31"] = 8
        workbook.save(path)

        findings = scan_source_for_quality_findings(path)

    assert len(findings) == 2
    assert "jumps from 10 to 12" in findings[0].text
    assert "Unit is missing for 'Rope Rod' with Qty 8" in findings[1].text
