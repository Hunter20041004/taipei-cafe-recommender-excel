from pathlib import Path
import os
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "workbook" / "taipei-cafe-recommender.xlsm"
VBA_SOURCE = ROOT / "src" / "vba" / "CafeRecommender.bas"

CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

ET.register_namespace("cp", CORE_NS)
ET.register_namespace("dc", DC_NS)
ET.register_namespace("dcterms", "http://purl.org/dc/terms/")
ET.register_namespace("dcmitype", "http://purl.org/dc/dcmitype/")
ET.register_namespace("xsi", "http://www.w3.org/2001/XMLSchema-instance")
ET.register_namespace("", SHEET_NS)


def remove_creator_metadata(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)
    for tag in (f"{{{DC_NS}}}creator", f"{{{CORE_NS}}}lastModifiedBy"):
        element = root.find(tag)
        if element is not None:
            root.remove(element)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def rewrite_zip_part(path: Path, part_name: str, transform) -> None:
    with tempfile.NamedTemporaryFile(dir=path.parent, suffix=path.suffix, delete=False) as handle:
        temporary_path = Path(handle.name)

    try:
        with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(temporary_path, "w") as target:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == part_name:
                    data = transform(data)
                target.writestr(info, data)
        os.replace(temporary_path, path)
    finally:
        temporary_path.unlink(missing_ok=True)


def replace_cell_formula(xml_bytes: bytes, address: str, formula_text: str) -> bytes:
    root = ET.fromstring(xml_bytes)
    cell = root.find(f".//{{{SHEET_NS}}}c[@r='{address}']")
    if cell is None:
        raise ValueError(f"Cell {address} not found")
    formula = cell.find(f"{{{SHEET_NS}}}f")
    if formula is None:
        raise ValueError(f"Formula not found in {address}")
    formula.text = formula_text
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def bound_all_data_references(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)
    pattern = re.compile(r"全部資料!\$([A-Z]+):\$([A-Z]+)")
    for formula in root.findall(f".//{{{SHEET_NS}}}f"):
        if formula.text:
            formula.text = pattern.sub(
                lambda match: (
                    f"全部資料!${match.group(1)}$2:${match.group(2)}$41"
                ),
                formula.text,
            )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def replace_shared_strings(xml_bytes: bytes, replacements: dict[str, str]) -> bytes:
    root = ET.fromstring(xml_bytes)
    for item in root.findall(f"{{{SHEET_NS}}}si"):
        texts = list(item.iter(f"{{{SHEET_NS}}}t"))
        if len(texts) == 1 and (texts[0].text or "") in replacements:
            texts[0].text = replacements[texts[0].text or ""]
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def inject_vba_source() -> None:
    from pyopenvba import ExcelFile, VBAModuleKind

    with ExcelFile(WORKBOOK) as workbook:
        source = VBA_SOURCE.read_text(encoding="utf-8")
        project = workbook.vba_project()
        if "CafeRecommender" in workbook.module_names():
            workbook.set_module("CafeRecommender", source)
        else:
            project.delete_module("Module1")
            project.add_module(
                "CafeRecommender",
                source,
                kind=VBAModuleKind.standard,
            )
        workbook.save()


def main() -> None:
    rewrite_zip_part(WORKBOOK, "docProps/core.xml", remove_creator_metadata)
    unlimited_highest_formula = (
        'INDEX(全部資料!$A$2:$A$41,MATCH(1,'
        '(全部資料!$H$2:$H$41="不限時")*'
        '(全部資料!$I$2:$I$41=MAX(IF(全部資料!$H$2:$H$41="不限時",'
        '全部資料!$I$2:$I$41))),0))'
    )
    rewrite_zip_part(
        WORKBOOK,
        "xl/worksheets/sheet3.xml",
        lambda data: replace_cell_formula(data, "C32", unlimited_highest_formula),
    )
    wenshan_cheapest_formula = (
        'INDEX(全部資料!$A$2:$A$41,MATCH(1,'
        '(全部資料!$B$2:$B$41="文山區")*'
        '(全部資料!$E$2:$E$41=MIN(IF(全部資料!$B$2:$B$41="文山區",'
        '全部資料!$E$2:$E$41))),0))'
    )
    rewrite_zip_part(
        WORKBOOK,
        "xl/worksheets/sheet3.xml",
        lambda data: replace_cell_formula(data, "C33", wenshan_cheapest_formula),
    )
    for sheet_part in (
        "xl/worksheets/sheet2.xml",
        "xl/worksheets/sheet3.xml",
    ):
        rewrite_zip_part(WORKBOOK, sheet_part, bound_all_data_references)
    rewrite_zip_part(
        WORKBOOK,
        "xl/sharedStrings.xml",
        lambda data: replace_shared_strings(
            data,
            {
                "☕ 大台北咖啡廳推薦系統": "☕ 新店・文山・大安咖啡廳推薦系統",
                "?? 搜尋統計摘要": "搜尋統計摘要",
                "INDEX + MATCH + MAXIFS": "INDEX + MATCH + MAX + IF",
                "INDEX + MATCH + MINIFS": "INDEX + MATCH + MIN + IF",
            },
        ),
    )
    inject_vba_source()


if __name__ == "__main__":
    main()
