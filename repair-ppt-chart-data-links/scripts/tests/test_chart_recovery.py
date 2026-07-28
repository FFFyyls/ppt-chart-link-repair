# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

import sys
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET


SCRIPTS = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SCRIPTS))

import chart_recovery as cr


C = "http://schemas.openxmlformats.org/drawingml/2006/chart"


class FormulaTests(unittest.TestCase):
    def test_parse_formula_with_quoted_sheet(self):
        parsed = cr.parse_formula("'[old.xlsx]Sheet Name'!$F$22:$K$22")
        self.assertEqual(parsed.workbook, "old.xlsx")
        self.assertEqual(parsed.sheet, "Sheet Name")
        self.assertEqual(parsed.address, "F22:K22")


class CacheTests(unittest.TestCase):
    def test_sparse_cache_preserves_missing_indexes_and_flags_count(self):
        node = ET.fromstring(
            f'<c:numRef xmlns:c="{C}"><c:numCache><c:ptCount val="3"/>'
            '<c:pt idx="0"><c:v>10</c:v></c:pt>'
            '<c:pt idx="2"><c:v>30</c:v></c:pt>'
            '</c:numCache></c:numRef>'
        )
        result = cr.cache_points(node, numeric=True)
        self.assertEqual(result.values, [10.0, None, 30.0])
        self.assertEqual(result.declared_count, 3)
        self.assertEqual(result.actual_count, 2)
        self.assertTrue(result.has_gaps)

    def test_multilevel_cache_maps_levels_to_two_dimensional_range(self):
        node = ET.fromstring(
            f'<c:multiLvlStrRef xmlns:c="{C}">'
            '<c:f>\'[data.xlsx]Sheet1\'!$A$1:$B$2</c:f>'
            '<c:multiLvlStrCache><c:ptCount val="2"/>'
            '<c:lvl><c:pt idx="0"><c:v>North</c:v></c:pt>'
            '<c:pt idx="1"><c:v>South</c:v></c:pt></c:lvl>'
            '<c:lvl><c:pt idx="0"><c:v>2024</c:v></c:pt>'
            '<c:pt idx="1"><c:v>2025</c:v></c:pt></c:lvl>'
            '</c:multiLvlStrCache></c:multiLvlStrRef>'
        )
        result = cr.multilevel_cache_points(node, "A1:B2")
        self.assertEqual(result.values, ["North", "2024", "South", "2025"])
        self.assertEqual(result.declared_count, 4)
        self.assertEqual(result.actual_count, 4)
        self.assertFalse(result.has_gaps)

    def test_leading_sparse_points_are_an_intentional_gap_not_cache_corruption(self):
        xml = f'''<c:chartSpace xmlns:c="{C}"><c:chart><c:plotArea><c:lineChart>
        <c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef>
        <c:f>'[data.xlsx]Sheet1'!$A$1:$A$5</c:f>
        <c:numCache><c:ptCount val="3"/>
        <c:pt idx="2"><c:v>6</c:v></c:pt>
        <c:pt idx="3"><c:v>7</c:v></c:pt>
        <c:pt idx="4"><c:v>8</c:v></c:pt>
        </c:numCache></c:numRef></c:val></c:ser>
        </c:lineChart></c:plotArea></c:chart></c:chartSpace>'''.encode("utf-8")

        chart = cr._analyze_chart(xml, "slide-1-chart-1")

        self.assertEqual("A", chart["recoverability"])
        self.assertEqual([2, 3, 4], chart["references"][0]["cache_indexes"])
        self.assertEqual("leading", chart["references"][0]["intentional_gap"])

    def test_defined_name_series_reference_is_recoverable(self):
        xml = f'''<c:chartSpace xmlns:c="{C}"><c:chart><c:plotArea><c:lineChart>
        <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef>
        <c:f>SeriesName</c:f><c:strCache><c:ptCount val="1"/>
        <c:pt idx="0"><c:v>收益率</c:v></c:pt>
        </c:strCache></c:strRef></c:tx><c:val><c:numRef>
        <c:f>'[data.xlsx]Sheet1'!$A$1</c:f><c:numCache><c:ptCount val="1"/>
        <c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache>
        </c:numRef></c:val></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>'''.encode("utf-8")

        chart = cr._analyze_chart(xml, "slide-1-chart-1")

        self.assertEqual("A", chart["recoverability"])
        self.assertEqual("SeriesName", chart["references"][0]["defined_name"])
        self.assertEqual(["收益率"], chart["references"][0]["values"])


class SafetyTests(unittest.TestCase):
    def test_recovery_requires_explicit_confirmation(self):
        with self.assertRaises(cr.ConfirmationRequired):
            cr.require_confirmation({"selection_confirmed": False})

    def test_group_key_uses_complete_original_target(self):
        self.assertNotEqual(
            cr.source_group_key(r"E:\\ProjectA\\data.xlsx", "chart1"),
            cr.source_group_key(r"D:\\ProjectB\\data.xlsx", "chart2"),
        )
        self.assertEqual(
            cr.source_group_key(r"E:\\ProjectA\\data.xlsx", "chart1"),
            cr.source_group_key(r"E:\\ProjectA\\data.xlsx", "chart9"),
        )

    def test_subprocess_output_decodes_utf8_before_windows_legacy_codepage(self):
        self.assertEqual(cr.decode_process_bytes("中文输出".encode("utf-8")), "中文输出")

    def test_powerpoint_external_uri_matches_windows_office_link_spelling(self):
        self.assertEqual(
            cr._file_uri(Path(r"C:\Data\中文 数据#%.xlsx")),
            r"file:///C:\Data\中文%20数据%23%25.xlsx",
        )


if __name__ == "__main__":
    unittest.main()
