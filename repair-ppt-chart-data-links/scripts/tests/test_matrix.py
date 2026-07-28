# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

import json
import sys
import tempfile
import time
import unittest
import zipfile
from pathlib import Path

from lxml import etree


TESTS = Path(__file__).resolve().parent
SCRIPTS = TESTS.parent
sys.path.insert(0, str(TESTS))
sys.path.insert(0, str(SCRIPTS))

import chart_recovery as cr
from test_workflow import make_fixture
from openpyxl import load_workbook


def rewrite_package(path: Path, additions=None, transforms=None) -> None:
    additions = additions or {}
    transforms = transforms or {}
    temp = path.with_suffix(".matrix.pptx")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(temp, "w") as target:
        names = set()
        for info in source.infolist():
            payload = source.read(info.filename)
            if info.filename in transforms:
                payload = transforms[info.filename](payload)
            target.writestr(info, payload)
            names.add(info.filename)
        for name, payload in additions.items():
            if name not in names:
                target.writestr(name, payload)
    deadline = time.monotonic() + 2.0
    while True:
        try:
            temp.replace(path)
            break
        except PermissionError:
            if time.monotonic() >= deadline:
                raise
            time.sleep(0.02)


def duplicate_chart_on_second_slide(path: Path, conflict=False) -> None:
    with zipfile.ZipFile(path) as archive:
        slide = archive.read("ppt/slides/slide1.xml")
        slide_rels = archive.read("ppt/slides/_rels/slide1.xml.rels").replace(
            b"chart1.xml", b"chart2.xml"
        )
        chart = archive.read("ppt/charts/chart1.xml")
        chart_rels = archive.read("ppt/charts/_rels/chart1.xml.rels")
    if conflict:
        chart = chart.replace(b"<c:v>20</c:v>", b"<c:v>21</c:v>")
    rewrite_package(
        path,
        additions={
            "ppt/slides/slide2.xml": slide,
            "ppt/slides/_rels/slide2.xml.rels": slide_rels,
            "ppt/charts/chart2.xml": chart,
            "ppt/charts/_rels/chart2.xml.rels": chart_rels,
        },
    )


class RequiredMatrixTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def test_multiple_charts_with_same_old_path_share_one_workbook(self):
        source = self.root / "shared.pptx"
        make_fixture(source)
        duplicate_chart_on_second_slide(source)
        inventory = cr.inspect_pptx(source, self.root / "shared-work")
        inventory["selection_confirmed"] = True
        for chart in inventory["charts"]:
            chart["selected"] = True
        path = Path(inventory["inventory_path"])
        path.write_text(json.dumps(inventory, ensure_ascii=False), encoding="utf-8")
        repair = cr.recover_manifest(path)
        self.assertEqual(len(repair["groups"]), 1)
        self.assertEqual(len(repair["groups"][0]["chart_ids"]), 2)

    def test_conflicting_shared_cell_blocks_relink(self):
        source = self.root / "conflict.pptx"
        make_fixture(source)
        duplicate_chart_on_second_slide(source, conflict=True)
        inventory = cr.inspect_pptx(source, self.root / "conflict-work")
        inventory["selection_confirmed"] = True
        for chart in inventory["charts"]:
            chart["selected"] = True
        path = Path(inventory["inventory_path"])
        path.write_text(json.dumps(inventory, ensure_ascii=False), encoding="utf-8")
        repair = cr.recover_manifest(path)
        self.assertFalse(repair["relink_allowed"])
        self.assertTrue(repair["groups"][0]["conflicts"])

    def test_conventional_combination_chart_is_analyzed_series_by_series(self):
        source = self.root / "combo.pptx"
        make_fixture(source)

        def add_line_chart(payload: bytes) -> bytes:
            text = payload.decode("utf-8")
            series = text.split("<c:ser>", 1)[1].split("</c:ser>", 1)[0]
            return text.replace(
                "</c:barChart>",
                f"</c:barChart><c:lineChart><c:ser>{series}</c:ser></c:lineChart>",
            ).encode("utf-8")

        rewrite_package(
            source, transforms={"ppt/charts/chart1.xml": add_line_chart}
        )
        inventory = cr.inspect_pptx(source, self.root / "combo-work")
        chart = inventory["charts"][0]
        self.assertEqual(chart["recoverability"], "A")
        self.assertEqual(set(chart["chart_types"]), {"barChart", "lineChart"})
        self.assertEqual(len({ref["series_index"] for ref in chart["references"]}), 2)

    def test_chartex_is_listed_as_unsupported(self):
        source = self.root / "chartex.pptx"
        make_fixture(source)

        def make_chartex(payload: bytes) -> bytes:
            text = payload.decode("utf-8")
            text = text.replace(
                'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"',
                'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
                'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"',
            )
            return text.replace("<c:chart r:id=", "<cx:chart r:id=").replace(
                "/></a:graphicData>", "/></a:graphicData>"
            ).encode("utf-8")

        rewrite_package(
            source, transforms={"ppt/slides/slide1.xml": make_chartex}
        )
        inventory = cr.inspect_pptx(source, self.root / "chartex-work")
        self.assertEqual(inventory["charts"][0]["recoverability"], "U")
        self.assertIn("ChartEx", inventory["charts"][0]["issues"][0])

    def test_structural_verify_detects_cache_regenerated_to_wrong_value(self):
        source = self.root / "wrong-cache.pptx"
        make_fixture(source)
        inventory = cr.inspect_pptx(source, self.root / "wrong-cache-work")
        inventory["selection_confirmed"] = True
        inventory["charts"][0]["selected"] = True
        path = Path(inventory["inventory_path"])
        path.write_text(json.dumps(inventory, ensure_ascii=False), encoding="utf-8")
        repair = cr.recover_manifest(path)
        manifest_path = Path(repair["manifest_path"])
        output = Path(repair["output_pptx"])
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_bytes(source.read_bytes())
        workbook = Path(repair["groups"][0]["workbook_path"])

        def relink_relationship(payload: bytes) -> bytes:
            root = etree.fromstring(payload)
            for relationship in root:
                if relationship.get("Id") == "rId1":
                    relationship.set("Target", cr._file_uri(workbook))
            return etree.tostring(root, encoding="utf-8", xml_declaration=True)

        rewrite_package(
            output,
            transforms={
                "ppt/charts/_rels/chart1.xml.rels": relink_relationship
            },
        )

        def corrupt(payload: bytes) -> bytes:
            return payload.replace(b"<c:v>20</c:v>", b"<c:v>999</c:v>")

        rewrite_package(output, transforms={"ppt/charts/chart1.xml": corrupt})
        result = cr.verify_manifest(manifest_path, office_required=False)
        self.assertFalse(result["passed"])

    def test_date_axis_round_trips_as_numeric_excel_serial(self):
        source = self.root / "date-axis.pptx"
        make_fixture(source)

        def make_date_axis(payload: bytes) -> bytes:
            text = payload.decode("utf-8")
            old = (
                "<c:cat><c:strRef><c:f>'[data.xlsx]Sheet A'!$B$1:$C$1</c:f>"
                '<c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>2024</c:v></c:pt>'
                '<c:pt idx="1"><c:v>2025</c:v></c:pt></c:strCache></c:strRef></c:cat>'
            )
            new = (
                "<c:cat><c:numRef><c:f>'[data.xlsx]Sheet A'!$B$1:$C$1</c:f>"
                '<c:numCache><c:formatCode>yyyy-mm-dd</c:formatCode><c:ptCount val="2"/>'
                '<c:pt idx="0"><c:v>44469</c:v></c:pt>'
                '<c:pt idx="1"><c:v>44470</c:v></c:pt></c:numCache></c:numRef></c:cat>'
            )
            self.assertIn(old, text)
            return text.replace(old, new).encode("utf-8")

        rewrite_package(
            source, transforms={"ppt/charts/chart1.xml": make_date_axis}
        )
        inventory = cr.inspect_pptx(source, self.root / "date-axis-work")
        inventory["selection_confirmed"] = True
        inventory["charts"][0]["selected"] = True
        path = Path(inventory["inventory_path"])
        path.write_text(json.dumps(inventory, ensure_ascii=False), encoding="utf-8")
        repair = cr.recover_manifest(path)
        workbook = Path(repair["groups"][0]["workbook_path"])
        book = load_workbook(workbook, read_only=True, data_only=False)
        try:
            self.assertEqual(
                cr._chart_cache_value(book["Sheet A"]["B1"].value, True, book.epoch),
                44469.0,
            )
            self.assertEqual(
                cr._chart_cache_value(book["Sheet A"]["C1"].value, True, book.epoch),
                44470.0,
            )
        finally:
            book.close()
        category = next(
            ref for ref in repair["charts"][0]["references"] if ref["role"] == "cat"
        )
        self.assertEqual(category["format_code"], "yyyy-mm-dd")


if __name__ == "__main__":
    unittest.main()
