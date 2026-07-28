# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from lxml import etree


SCRIPTS = Path(__file__).resolve().parents[1]
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

import native_link_pipeline as nlp  # noqa: E402
from native_link_pipeline import RecipeConflict, _plot_metadata, build_chart_recipes  # noqa: E402
from transplant_native_links import replace_formulas  # noqa: E402


def reference(role: str, sheet: str, cells: list[str], values: list[object]) -> dict:
    return {
        "series_index": 0,
        "role": role,
        "formula": f"'[book.xlsx]{sheet}'!$A$1:$A${len(cells)}",
        "sheet": sheet,
        "address": f"A1:A{len(cells)}",
        "cells": cells,
        "values": values,
        "cache_indexes": [index for index, value in enumerate(values) if value is not None],
        "cache_type": "numeric" if role == "val" else "string",
    }


class NativeLinkRecipeTests(unittest.TestCase):
    def test_office_workers_clean_all_owned_office_descendants(self) -> None:
        for worker_name in (
            "office_native_paste_batch.ps1",
            "office_capture_series_names.ps1",
        ):
            with self.subTest(worker=worker_name):
                worker = (SCRIPTS / worker_name).read_text(encoding="utf-8")
                self.assertIn("function Stop-OwnedOfficeDescendants", worker)
                self.assertIn("Stop-OwnedOfficeDescendants $PID", worker)

    def test_native_paste_pipeline_invokes_one_batch_worker(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source = root / "source.pptx"
            source.write_bytes(b"ppt")
            candidate = root / "candidate.pptx"
            recipes_path = root / "recipes.json"
            support_map = root / "support.json"
            paste_dir = root / "paste-results"
            paste_dir.mkdir()
            transplant_result = root / "transplant-result.json"
            stage_status = root / "stage-status.json"
            recipes = [{"chart_id": "chart-1"}, {"chart_id": "chart-2"}]
            calls: list[list[str]] = []

            def fake_run(command: list[str], timeout: int):
                calls.append(command)
                if any(part.endswith("transplant_native_links.py") for part in command):
                    transplant_result.write_text(
                        json.dumps({"passed": True}), encoding="utf-8"
                    )
                return subprocess.CompletedProcess(command, 0, b"", b"")

            with patch.object(nlp, "_run", side_effect=fake_run):
                result = nlp._paste_and_transplant(
                    scripts=SCRIPTS,
                    source=source,
                    output=root / "output.pptx",
                    candidate=candidate,
                    recipes=recipes,
                    recipes_path=recipes_path,
                    support_map_path=support_map,
                    paste_dir=paste_dir,
                    transplant_result=transplant_result,
                    stage_status=stage_status,
                )

            batch_calls = [
                command
                for command in calls
                if any("office_native_paste_batch.ps1" in part for part in command)
            ]
            legacy_calls = [
                command
                for command in calls
                if any(part.endswith("office_native_paste.ps1") for part in command)
            ]
            self.assertTrue(result["passed"])
            self.assertEqual(len(batch_calls), 1)
            self.assertEqual(legacy_calls, [])
            self.assertEqual(len(calls), 2)

    def test_runtime_series_name_enrichment_requires_unambiguous_capture(self) -> None:
        manifest = {
            "charts": [
                {
                    "chart_id": "chart-1",
                    "data_mappings": [
                        {"series_index": 0, "series_name": None},
                    ],
                }
            ]
        }
        results = {
            "chart-1": {"runtime_series_names": ["系列1"]},
        }

        nlp.enrich_runtime_series_names(manifest, results)

        self.assertEqual(
            manifest["charts"][0]["data_mappings"][0]["runtime_series_name"],
            "系列1",
        )
        manifest["charts"][0]["data_mappings"][0]["runtime_series_name"] = "Series 1"
        nlp.enrich_runtime_series_names(
            manifest,
            {"chart-1": {"runtime_series_names": ["Series 2"]}},
        )
        self.assertEqual(
            manifest["charts"][0]["data_mappings"][0]["runtime_series_name"],
            "Series 2",
        )
        ambiguous = {
            "charts": [
                {
                    "chart_id": "chart-1",
                    "data_mappings": [
                        {"series_index": 0, "series_name": None},
                    ],
                }
            ]
        }
        with self.assertRaisesRegex(nlp.NativeLinkError, "unique runtime series"):
            nlp.enrich_runtime_series_names(
                ambiguous,
                {"chart-1": {"runtime_series_names": ["系列1", "系列2"]}},
            )

    def test_runtime_series_name_enrichment_rejects_duplicates_and_mismatch(self) -> None:
        missing = {
            "charts": [{"chart_id": "chart-1", "data_mappings": [
                {"series_index": 0, "series_name": None}
            ]}]
        }
        with self.assertRaisesRegex(nlp.NativeLinkError, "duplicate runtime series"):
            nlp.enrich_runtime_series_names(
                missing,
                {"chart-1": {"runtime_series_names": ["dup", "dup"]}},
            )

        known = {
            "charts": [{"chart_id": "chart-1", "data_mappings": [
                {"series_index": 0, "series_name": "Expected"}
            ]}]
        }
        with self.assertRaisesRegex(nlp.NativeLinkError, "does not match"):
            nlp.enrich_runtime_series_names(
                known,
                {"chart-1": {"runtime_series_names": ["Unexpected"]}},
            )

    def test_plot_metadata_uses_document_position_when_plot_orders_overlap(self) -> None:
        chart = b'''<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea>
        <c:barChart><c:barDir val="col"/><c:ser><c:idx val="3"/><c:order val="3"/></c:ser><c:ser><c:idx val="4"/><c:order val="4"/></c:ser></c:barChart>
        <c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/></c:ser><c:ser><c:idx val="1"/><c:order val="1"/></c:ser></c:lineChart>
        </c:plotArea></c:chart></c:chartSpace>'''
        _, by_position = _plot_metadata(chart)
        self.assertEqual((by_position[0]["plot_type"], by_position[0]["order"]), ("barChart", 3))
        self.assertEqual((by_position[2]["plot_type"], by_position[2]["order"]), ("lineChart", 0))

    def test_transplant_matches_series_by_value_reference_not_candidate_order(self) -> None:
        source = etree.fromstring(
            b'''<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart>
            <c:ser><c:order val="3"/><c:val><c:numRef><c:f>Sheet1!$E$3:$E$5</c:f></c:numRef></c:val></c:ser>
            <c:ser><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!$B$3:$B$5</c:f></c:numRef></c:val></c:ser>
            </c:lineChart></c:plotArea></c:chart></c:chartSpace>'''
        )
        candidate = etree.fromstring(
            b'''<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart>
            <c:ser><c:order val="0"/><c:val><c:numRef><c:f>'[new.xlsx]Sheet1'!$B$3:$B$5</c:f></c:numRef></c:val></c:ser>
            <c:ser><c:order val="1"/><c:val><c:numRef><c:f>'[new.xlsx]Sheet1'!$E$3:$E$5</c:f></c:numRef></c:val></c:ser>
            </c:lineChart></c:plotArea></c:chart></c:chartSpace>'''
        )
        recipe = {
            "series": [
                {"order": 3, "refs": {"val": {"parsed": {"kind": "range", "sheet": "Sheet1", "address": "E3:E5"}}}},
                {"order": 0, "refs": {"val": {"parsed": {"kind": "range", "sheet": "Sheet1", "address": "B3:B5"}}}},
            ]
        }
        changes = replace_formulas(source, candidate, recipe)
        by_order = {change["order"]: change["new"] for change in changes}
        self.assertIn("$E$3:$E$5", by_order[3])
        self.assertIn("$B$3:$B$5", by_order[0])

    def test_preserves_sparse_cache_indexes_and_cross_sheet_references(self) -> None:
        chart = {
            "chart_id": "slide-12-chart-1",
            "actual_slide": 12,
            "shape_name": "Chart 4",
            "chart_part": "ppt/charts/chart11.xml",
            "chart_rel_part": "ppt/charts/_rels/chart11.xml.rels",
            "chart_types": ["lineChart", "barChart"],
            "references": [
                reference("cat", "分类", [f"A{i}" for i in range(1, 6)], ["a", "b", "c", "d", "e"]),
                reference("val", "数值", [f"B{i}" for i in range(1, 6)], [None, None, 6.0, 7.0, 8.0]),
            ],
        }
        manifest = {
            "charts": [chart],
            "groups": [{"chart_ids": [chart["chart_id"]], "workbook_path": r"C:\包\数据.xlsx"}],
        }

        recipe = build_chart_recipes(manifest)[0]
        value_ref = recipe["series"][0]["refs"]["val"]

        self.assertEqual([2, 3, 4], value_ref["cache_indexes"])
        self.assertEqual(["B3", "B4", "B5"], [point["cell"] for point in value_ref["materialized_points"]])
        self.assertEqual([6.0, 7.0, 8.0], [point["value"] for point in value_ref["materialized_points"]])
        self.assertEqual("分类", recipe["series"][0]["refs"]["cat"]["sheet"])
        self.assertEqual("数值", value_ref["sheet"])

    def test_keeps_defined_name_reference(self) -> None:
        chart = {
            "chart_id": "slide-23-chart-1",
            "actual_slide": 23,
            "shape_name": "Chart 5",
            "chart_part": "ppt/charts/chart23.xml",
            "chart_rel_part": "ppt/charts/_rels/chart23.xml.rels",
            "chart_types": ["barChart", "scatterChart"],
            "references": [
                {
                    "series_index": 0,
                    "role": "tx",
                    "formula": "系列名称一",
                    "defined_name": "系列名称一",
                    "cells": [],
                    "values": ["收益率"],
                    "cache_indexes": [0],
                    "cache_type": "string",
                },
                reference("val", "Sheet4", ["D14"], [0.056102]),
            ],
        }
        manifest = {
            "charts": [chart],
            "groups": [{"chart_ids": [chart["chart_id"]], "workbook_path": r"C:\包\数据.xlsx"}],
        }

        recipe = build_chart_recipes(manifest)[0]

        self.assertEqual("系列名称一", recipe["series"][0]["refs"]["tx"]["defined_name"])

    def test_conflicting_cells_block_recipe(self) -> None:
        chart = {
            "chart_id": "slide-1-chart-1",
            "actual_slide": 1,
            "shape_name": "Chart 1",
            "chart_part": "ppt/charts/chart1.xml",
            "chart_rel_part": "ppt/charts/_rels/chart1.xml.rels",
            "chart_types": ["lineChart"],
            "references": [
                reference("val", "Sheet1", ["A1"], [1.0]),
                {**reference("val", "Sheet1", ["A1"], [2.0]), "series_index": 1},
            ],
        }
        manifest = {
            "charts": [chart],
            "groups": [{"chart_ids": [chart["chart_id"]], "workbook_path": r"C:\包\数据.xlsx"}],
        }

        with self.assertRaises(RecipeConflict):
            build_chart_recipes(manifest)


if __name__ == "__main__":
    unittest.main()
