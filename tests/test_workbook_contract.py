from pathlib import Path
import re
import unittest
import xml.etree.ElementTree as ET
import zipfile


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "workbook" / "taipei-cafe-recommender.xlsm"
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def read_cell_formula(sheet_part: str, address: str) -> str:
    with zipfile.ZipFile(WORKBOOK) as archive:
        sheet = ET.fromstring(archive.read(sheet_part))
    cell = sheet.find(f".//{{{SHEET_NS}}}c[@r='{address}']")
    if cell is None:
        raise AssertionError(f"Cell {address} not found in {sheet_part}")
    formula = cell.find(f"{{{SHEET_NS}}}f")
    if formula is None or formula.text is None:
        raise AssertionError(f"Formula not found in {sheet_part}!{address}")
    return formula.text


def read_all_formulas(sheet_part: str) -> list[str]:
    with zipfile.ZipFile(WORKBOOK) as archive:
        sheet = ET.fromstring(archive.read(sheet_part))
    return [
        formula.text or ""
        for formula in sheet.findall(f".//{{{SHEET_NS}}}f")
    ]


def read_shared_strings() -> list[str]:
    with zipfile.ZipFile(WORKBOOK) as archive:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return [
        "".join(text.text or "" for text in item.iter(f"{{{SHEET_NS}}}t"))
        for item in root.findall(f"{{{SHEET_NS}}}si")
    ]


def read_embedded_module(name: str = "CafeRecommender") -> str:
    from pyopenvba import ExcelFile

    with ExcelFile(WORKBOOK) as workbook:
        return workbook.get_module(name)


def read_vba_subroutine(source: str, name: str) -> str:
    start = source.index(f"Sub {name}")
    end = source.index("End Sub", start)
    return source[start:end]


class WorkbookContractTests(unittest.TestCase):
    def test_creator_metadata_removed(self):
        with zipfile.ZipFile(WORKBOOK) as archive:
            core = ET.fromstring(archive.read("docProps/core.xml"))

        namespaces = {
            "dc": "http://purl.org/dc/elements/1.1/",
            "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        }
        creator = core.find("dc:creator", namespaces)
        modified_by = core.find("cp:lastModifiedBy", namespaces)

        self.assertIsNone(creator)
        self.assertIsNone(modified_by)

    def test_unlimited_highest_rated_lookup_keeps_both_constraints(self):
        formula = read_cell_formula("xl/worksheets/sheet3.xml", "C32")

        self.assertIn('全部資料!$H$2:$H$41="不限時"', formula)
        self.assertIn(
            '全部資料!$I$2:$I$41=MAX(IF(全部資料!$H$2:$H$41="不限時",全部資料!$I$2:$I$41))',
            formula,
        )
        self.assertIn("MATCH(1,", formula)

    def test_wenshan_cheapest_lookup_keeps_both_constraints(self):
        formula = read_cell_formula("xl/worksheets/sheet3.xml", "C33")

        self.assertIn('全部資料!$B$2:$B$41="文山區"', formula)
        self.assertIn(
            '全部資料!$E$2:$E$41=MIN(IF(全部資料!$B$2:$B$41="文山區",全部資料!$E$2:$E$41))',
            formula,
        )
        self.assertIn("MATCH(1,", formula)

    def test_portfolio_formulas_avoid_whole_column_references(self):
        formulas = read_all_formulas("xl/worksheets/sheet2.xml")
        formulas += read_all_formulas("xl/worksheets/sheet3.xml")
        whole_column = re.compile(r"\$?[A-Z]+:\$?[A-Z]+")

        offenders = [formula for formula in formulas if whole_column.search(formula)]

        self.assertEqual([], offenders)

    def test_product_copy_is_precise_and_encoding_safe(self):
        strings = read_shared_strings()

        self.assertIn("☕ 新店・文山・大安咖啡廳推薦系統", strings)
        self.assertIn("搜尋統計摘要", strings)
        self.assertIn("INDEX + MATCH + MAX + IF", strings)
        self.assertIn("INDEX + MATCH + MIN + IF", strings)
        self.assertNotIn("☕ 大台北咖啡廳推薦系統", strings)
        self.assertNotIn("?? 搜尋統計摘要", strings)
        self.assertNotIn("INDEX + MATCH + MAXIFS", strings)
        self.assertNotIn("INDEX + MATCH + MINIFS", strings)

    def test_no_results_branch_clears_stale_chart(self):
        source = read_embedded_module()
        branch_start = source.index("If matchCount = 0 Then")
        branch_end = source.index("End If", branch_start)
        no_results_branch = source[branch_start:branch_end]

        self.assertIn("更新圖表 0", no_results_branch)

    def test_clear_results_clears_stale_chart(self):
        source = read_embedded_module()
        clear_results = read_vba_subroutine(source, "清除結果()")

        self.assertIn("更新圖表 0", clear_results)

    def test_chart_update_handles_zero_results_without_stale_series(self):
        source = read_embedded_module()
        update_chart = read_vba_subroutine(source, "更新圖表(dataCount As Long)")

        self.assertIn("If dataCount = 0 Then", update_chart)
        self.assertIn(".SeriesCollection(1).Delete", update_chart)
        self.assertIn("Exit Sub", update_chart)

    def test_search_skips_rows_with_nonnumeric_filter_values(self):
        source = read_embedded_module()
        search = read_vba_subroutine(source, "搜尋咖啡廳()")

        self.assertIn("If Not IsNumeric(wsData.Cells(i, 4).Value)", search)
        self.assertIn("Not IsNumeric(wsData.Cells(i, 5).Value)", search)
        self.assertIn("Not IsNumeric(wsData.Cells(i, 9).Value) Then GoTo NextRow", search)

    def test_vba_copy_is_precise_and_encoding_safe(self):
        source = read_embedded_module()

        self.assertIn("新店、文山、大安咖啡廳推薦系統", source)
        self.assertIn('wsDash.Range("B27").Value = "搜尋統計摘要"', source)
        self.assertNotIn("?? 搜尋統計摘要", source)
        self.assertNotIn("大台北咖啡廳推薦系統", source)

    def test_embedded_vba_matches_reviewable_source(self):
        embedded = read_embedded_module().replace("\r\n", "\n")
        reviewable = (ROOT / "src" / "vba" / "CafeRecommender.bas").read_text(
            encoding="utf-8"
        ).replace("\r\n", "\n")

        self.assertEqual(reviewable, embedded)

    def test_vba_uses_fresh_module_without_stale_compiled_cache(self):
        from pyopenvba import ExcelFile

        with ExcelFile(WORKBOOK) as workbook:
            names = workbook.module_names()

        self.assertIn("CafeRecommender", names)
        self.assertNotIn("Module1", names)

    def test_vba_avoids_ambiguous_call_keyword(self):
        source = read_embedded_module()

        self.assertNotIn("Call ", source)

    def test_vba_has_no_auto_run_or_system_side_effects(self):
        source = read_embedded_module().casefold()
        forbidden = (
            "workbook_open",
            "auto_open",
            "shell(",
            "createobject(",
            "filesystemobject",
            "powershell",
            "wscript",
            "xmlhttp",
            "urlmon",
            "kill ",
            "open " + '"' + " for output",
        )

        for keyword in forbidden:
            with self.subTest(keyword=keyword):
                self.assertNotIn(keyword, source)


if __name__ == "__main__":
    unittest.main()
