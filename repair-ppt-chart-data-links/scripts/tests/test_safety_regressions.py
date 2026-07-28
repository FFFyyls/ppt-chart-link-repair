# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
from xml.etree import ElementTree as ET


TESTS = Path(__file__).resolve().parent
SCRIPTS = TESTS.parent
sys.path.insert(0, str(TESTS))
sys.path.insert(0, str(SCRIPTS))

import chart_recovery as cr
import ooxml_security as security
import transplant_native_links as transplant
from test_workflow import make_fixture


PR = "http://schemas.openxmlformats.org/package/2006/relationships"


def rewrite_zip_member(path: Path, member: str, transform) -> None:
    replacement = path.with_suffix(".rewritten.pptx")
    with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(replacement, "w") as target:
        for info in source.infolist():
            payload = source.read(info.filename)
            if info.filename == member:
                payload = transform(payload)
            target.writestr(info, payload)
    replacement.replace(path)


def make_declared_count_mismatch(path: Path) -> None:
    make_fixture(path)

    def alter_chart(payload: bytes) -> bytes:
        text = payload.decode("utf-8")
        # Keep both cached cells populated while making the cache metadata invalid.
        return text.replace(
            '<c:numCache><c:formatCode>0.00</c:formatCode><c:ptCount val="2"/>',
            '<c:numCache><c:formatCode>0.00</c:formatCode><c:ptCount val="3"/>',
        ).encode("utf-8")

    rewrite_zip_member(path, "ppt/charts/chart1.xml", alter_chart)


class SafetyRegressionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def _recover(self, fixture: Path, name: str = "case") -> tuple[dict, Path]:
        inventory = cr.inspect_pptx(fixture, self.root / name)
        inventory["selection_confirmed"] = True
        inventory["charts"][0]["selected"] = True
        inventory_path = Path(inventory["inventory_path"])
        inventory_path.write_text(
            json.dumps(inventory, ensure_ascii=False), encoding="utf-8"
        )
        repair = cr.recover_manifest(inventory_path)
        return repair, Path(repair["manifest_path"])

    def _write_updates(self, name: str, payload: dict) -> Path:
        path = self.root / name
        path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        return path

    def test_archive_total_expanded_size_is_limited_before_reading_members(self):
        archive_path = self.root / "oversized.pptx"
        with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml", b"x" * 64)

        with zipfile.ZipFile(archive_path) as archive:
            with patch.object(security, "MAX_TOTAL_UNCOMPRESSED_BYTES", 32):
                with self.assertRaisesRegex(
                    security.OoxmlSecurityError, "expanded size"
                ):
                    security.validate_ooxml_archive(archive)

    def test_inspect_rejects_dtd_and_entity_declarations(self):
        source = self.root / "xml-entity.pptx"
        make_fixture(source)

        def add_doctype(payload: bytes) -> bytes:
            declaration = (
                b'<!DOCTYPE p:sld [<!ENTITY local_entity "unexpected">]>\n'
            )
            marker = payload.find(b"?>")
            if marker >= 0:
                return payload[: marker + 2] + b"\n" + declaration + payload[marker + 2 :]
            return declaration + payload

        rewrite_zip_member(source, "ppt/slides/slide1.xml", add_doctype)

        with self.assertRaisesRegex(cr.RecoveryBlocked, "DTD|entity"):
            cr.inspect_pptx(source, self.root / "xml-entity-work")

    def test_empty_updates_cannot_promote_original_b_cache_to_a(self):
        source = self.root / "metadata-mismatch.pptx"
        make_declared_count_mismatch(source)
        repair, manifest_path = self._recover(source, "metadata-mismatch-work")
        self.assertEqual(repair["charts"][0]["recoverability"], "B")

        updates = self._write_updates(
            "empty-updates.json", {"updates_confirmed": True, "updates": []}
        )
        result = cr.apply_updates(manifest_path, updates)

        self.assertEqual(result["charts"][0]["recoverability"], "B")
        self.assertFalse(result["relink_allowed"])

    def test_apply_updates_requires_explicit_update_confirmation(self):
        source = self.root / "confirmation.pptx"
        make_fixture(source)
        _, manifest_path = self._recover(source, "confirmation-work")
        updates = self._write_updates(
            "unconfirmed.json",
            {
                "updates": [
                    {
                        "chart_id": "slide-1-chart-1",
                        "sheet": "Sheet A",
                        "target_cell": "C2",
                        "old_value": 20,
                        "new_value": 25,
                    }
                ]
            },
        )
        with self.assertRaises(cr.ConfirmationRequired):
            cr.apply_updates(manifest_path, updates)

    def test_apply_updates_requires_old_value(self):
        source = self.root / "old-value.pptx"
        make_fixture(source)
        _, manifest_path = self._recover(source, "old-value-work")
        updates = self._write_updates(
            "missing-old-value.json",
            {
                "updates_confirmed": True,
                "updates": [
                    {
                        "chart_id": "slide-1-chart-1",
                        "sheet": "Sheet A",
                        "target_cell": "C2",
                        "new_value": 25,
                    }
                ],
            },
        )
        with self.assertRaises(cr.RecoveryBlocked):
            cr.apply_updates(manifest_path, updates)

    def test_apply_updates_rejects_cell_outside_selected_chart_references(self):
        source = self.root / "out-of-range.pptx"
        make_fixture(source)
        _, manifest_path = self._recover(source, "out-of-range-work")
        updates = self._write_updates(
            "out-of-range.json",
            {
                "updates_confirmed": True,
                "updates": [
                    {
                        "chart_id": "slide-1-chart-1",
                        "sheet": "Sheet A",
                        "target_cell": "D99",
                        "old_value": None,
                        "new_value": 25,
                    }
                ],
            },
        )
        with self.assertRaises(cr.RecoveryBlocked):
            cr.apply_updates(manifest_path, updates)

    def test_apply_updates_rejects_new_value_with_wrong_data_type(self):
        source = self.root / "wrong-type.pptx"
        make_fixture(source)
        _, manifest_path = self._recover(source, "wrong-type-work")
        updates = self._write_updates(
            "wrong-type.json",
            {
                "updates_confirmed": True,
                "updates": [
                    {
                        "chart_id": "slide-1-chart-1",
                        "sheet": "Sheet A",
                        "target_cell": "C2",
                        "old_value": 20,
                        "new_value": "25",
                    }
                ],
            },
        )
        with self.assertRaises(cr.RecoveryBlocked):
            cr.apply_updates(manifest_path, updates)

    def test_relink_only_rewrites_relationship_used_by_external_data(self):
        source = self.root / "multiple-external.pptx"
        make_fixture(source)

        untouched_target = "https://example.invalid/source"

        def add_external_relationship(payload: bytes) -> bytes:
            root = ET.fromstring(payload)
            ET.SubElement(
                root,
                f"{{{PR}}}Relationship",
                {
                    "Id": "rId2",
                    "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                    "Target": untouched_target,
                    "TargetMode": "External",
                },
            )
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)

        rewrite_zip_member(
            source,
            "ppt/charts/_rels/chart1.xml.rels",
            add_external_relationship,
        )
        candidate = self.root / "candidate.pptx"
        candidate.write_bytes(source.read_bytes())
        new_target = "file:///C:\\Recovered\\data.xlsx"

        def update_candidate_target(payload: bytes) -> bytes:
            root = ET.fromstring(payload)
            for relationship in root.findall(f"{{{PR}}}Relationship"):
                if relationship.get("Id") == "rId1":
                    relationship.set("Target", new_target)
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)

        rewrite_zip_member(
            candidate,
            "ppt/charts/_rels/chart1.xml.rels",
            update_candidate_target,
        )
        manifest = self.root / "native-recipes.json"
        manifest.write_text(
            json.dumps(
                {
                    "charts": [
                        {
                            "chart_id": "slide-1-chart-1",
                            "actual_slide": 1,
                            "chart_part": "ppt/charts/chart1.xml",
                            "chart_rel_part": "ppt/charts/_rels/chart1.xml.rels",
                            "series": [
                                {
                                    "order": 0,
                                    "refs": {
                                        "val": {
                                            "parsed": {
                                                "kind": "range",
                                                "sheet": "Sheet A",
                                                "address": "B2:C2",
                                            }
                                        }
                                    },
                                }
                            ],
                        }
                    ]
                }
            ),
            encoding="utf-8",
        )
        paste_dir = self.root / "paste"
        paste_dir.mkdir()
        (paste_dir / "slide-1-chart-1.json").write_text(
            json.dumps({"linked_shape_name": "Chart 1"}), encoding="utf-8"
        )
        output = self.root / "transplanted.pptx"
        transplant.build(source, candidate, manifest, paste_dir, output)

        with zipfile.ZipFile(output) as archive:
            rels = ET.fromstring(
                archive.read("ppt/charts/_rels/chart1.xml.rels")
            )
        targets = {
            rel.get("Id"): rel.get("Target")
            for rel in rels.findall(f"{{{PR}}}Relationship")
        }
        self.assertEqual(targets["rId1"], new_target)
        self.assertEqual(targets["rId2"], untouched_target)

    def test_recover_rejects_selected_c_u_and_n_charts(self):
        for grade in ("C", "U", "N"):
            with self.subTest(recoverability=grade):
                source = self.root / f"grade-{grade}.pptx"
                make_fixture(source)
                inventory = cr.inspect_pptx(source, self.root / f"grade-{grade}-work")
                inventory["selection_confirmed"] = True
                inventory["charts"][0]["selected"] = True
                inventory["charts"][0]["recoverability"] = grade
                inventory_path = Path(inventory["inventory_path"])
                inventory_path.write_text(
                    json.dumps(inventory, ensure_ascii=False), encoding="utf-8"
                )
                with self.assertRaises(cr.RecoveryBlocked):
                    cr.recover_manifest(inventory_path)


if __name__ == "__main__":
    unittest.main()
