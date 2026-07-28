# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

from __future__ import annotations

import argparse
import json
import posixpath
import re
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from ooxml_security import safe_xml_fromstring, validate_ooxml_archive


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
R_ID = f"{{{NS['r']}}}id"
ROLES = ("tx", "cat", "val", "xVal", "yVal", "bubbleSize")


def resolve(base: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(base), target))


def chart_rel_part(chart_part: str) -> str:
    return posixpath.join(posixpath.dirname(chart_part), "_rels", posixpath.basename(chart_part) + ".rels")


def chart_part_by_shape(zf: zipfile.ZipFile, slide_number: int, shape_name: str) -> str:
    slide_part = f"ppt/slides/slide{slide_number}.xml"
    rel_part = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
    slide = safe_xml_fromstring(zf.read(slide_part))
    rels_root = safe_xml_fromstring(zf.read(rel_part))
    rels = {rel.get("Id"): resolve(slide_part, rel.get("Target")) for rel in rels_root}
    for frame in slide.findall(".//p:graphicFrame", NS):
        name_node = frame.find("p:nvGraphicFramePr/p:cNvPr", NS)
        chart_node = frame.find(".//c:chart", NS)
        if name_node is not None and chart_node is not None and name_node.get("name") == shape_name:
            return rels[chart_node.get(R_ID)]
    raise ValueError(f"shape not found in slide XML: slide={slide_number} name={shape_name}")


def series_formula_maps(root: etree._Element) -> list[dict]:
    result = []
    for series in root.findall(".//c:ser", NS):
        order_node = series.find("c:order", NS)
        if order_node is None:
            continue
        order = int(order_node.get("val"))
        formulas = {}
        for role in ROLES:
            role_node = series.find(f"c:{role}", NS)
            if role_node is None:
                continue
            formula = role_node.findtext(".//c:f", namespaces=NS)
            if formula:
                formulas[role] = formula
        result.append({"candidate_order": order, "formulas": formulas})
    return result


def formula_signature(formula: str) -> tuple[str, str] | None:
    if not formula or "!" not in formula:
        return None
    left, address = formula.rsplit("!", 1)
    left = left.strip().strip("'").replace("''", "'")
    workbook = re.search(r"\[[^\]]+\]", left)
    if workbook:
        left = left[workbook.end() :]
    return left.casefold(), address.replace("$", "").casefold()


def _recipe_value_signature(recipe_series: dict) -> tuple[str, str]:
    refs = recipe_series.get("refs", {})
    reference = refs.get("val") or refs.get("yVal")
    if not reference or reference.get("parsed", {}).get("kind") != "range":
        raise ValueError(
            f"Recipe series {recipe_series.get('order')} has no range value reference"
        )
    parsed = reference["parsed"]
    return str(parsed["sheet"]).casefold(), str(parsed["address"]).replace("$", "").casefold()


def candidate_map_by_recipe(
    candidate_root: etree._Element, recipe_chart: dict
) -> dict[tuple[int, str], str]:
    candidates = series_formula_maps(candidate_root)
    unused = set(range(len(candidates)))
    result: dict[tuple[int, str], str] = {}
    for recipe_series in sorted(recipe_chart.get("series", []), key=lambda item: int(item["order"])):
        expected = _recipe_value_signature(recipe_series)
        matches = []
        for index in sorted(unused):
            formulas = candidates[index]["formulas"]
            value_formula = formulas.get("val") or formulas.get("yVal")
            if value_formula and formula_signature(value_formula) == expected:
                matches.append(index)
        if not matches:
            raise ValueError(
                f"No candidate series matches original order={recipe_series['order']} value={expected}"
            )
        selected = matches[0]
        unused.remove(selected)
        for role, formula in candidates[selected]["formulas"].items():
            result[(int(recipe_series["order"]), role)] = formula
    return result


def replace_formulas(
    source_root: etree._Element,
    candidate_root: etree._Element,
    recipe_chart: dict,
) -> list[dict]:
    candidate = candidate_map_by_recipe(candidate_root, recipe_chart)
    changes = []
    for series in source_root.findall(".//c:ser", NS):
        order_node = series.find("c:order", NS)
        if order_node is None:
            continue
        order = int(order_node.get("val"))
        for role in ROLES:
            role_node = series.find(f"c:{role}", NS)
            if role_node is None:
                continue
            formula_node = role_node.find(".//c:f", NS)
            if formula_node is None:
                continue
            key = (order, role)
            fallback = None
            if role == "yVal":
                fallback = candidate.get((order, "val"))
            elif role == "val":
                fallback = candidate.get((order, "yVal"))
            new_formula = candidate.get(key, fallback)
            if not new_formula:
                raise ValueError(f"candidate formula missing for series order={order} role={role}; available={sorted(candidate)}")
            changes.append({"order": order, "role": role, "old": formula_node.text, "new": new_formula})
            formula_node.text = new_formula
    return changes


def relationship_by_id(root: etree._Element, relationship_id: str) -> etree._Element:
    for rel in root:
        if rel.get("Id") == relationship_id:
            return rel
    raise ValueError(f"relationship not found: {relationship_id}")


def external_relationship(root: etree._Element) -> etree._Element:
    for rel in root:
        if (rel.get("Type") or "").endswith("/oleObject") and rel.get("TargetMode") == "External":
            return rel
    raise ValueError("candidate external oleObject relationship not found")


def build(source: Path, candidate: Path, manifest_path: Path, paste_dir: Path, output: Path) -> dict:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    patches: dict[str, bytes] = {}
    details = []
    with zipfile.ZipFile(source) as source_zip, zipfile.ZipFile(candidate) as candidate_zip:
        validate_ooxml_archive(source_zip)
        validate_ooxml_archive(candidate_zip)
        for chart in manifest["charts"]:
            chart_id = chart["chart_id"]
            paste_result = json.loads((paste_dir / f"{chart_id}.json").read_text(encoding="utf-8"))
            candidate_part = chart_part_by_shape(
                candidate_zip,
                int(chart["actual_slide"]),
                paste_result["linked_shape_name"],
            )
            source_part = chart["chart_part"]
            source_rel_part = chart["chart_rel_part"]
            candidate_rel_part = chart_rel_part(candidate_part)
            source_root = safe_xml_fromstring(source_zip.read(source_part))
            candidate_root = safe_xml_fromstring(candidate_zip.read(candidate_part))
            changes = replace_formulas(source_root, candidate_root, chart)

            source_external = source_root.find("c:externalData", NS)
            candidate_external = candidate_root.find("c:externalData", NS)
            if source_external is None or candidate_external is None:
                raise ValueError(f"externalData missing: {chart_id}")
            source_rel_root = safe_xml_fromstring(source_zip.read(source_rel_part))
            candidate_rel_root = safe_xml_fromstring(candidate_zip.read(candidate_rel_part))
            source_external_rel = relationship_by_id(source_rel_root, source_external.get(R_ID))
            candidate_external_rel = external_relationship(candidate_rel_root)
            source_external_rel.set("Target", candidate_external_rel.get("Target"))
            source_external_rel.set("TargetMode", "External")

            patches[source_part] = etree.tostring(source_root, xml_declaration=True, encoding="UTF-8", standalone=True)
            patches[source_rel_part] = etree.tostring(source_rel_root, xml_declaration=True, encoding="UTF-8", standalone=True)
            details.append(
                {
                    "chart_id": chart_id,
                    "source_part": source_part,
                    "candidate_part": candidate_part,
                    "formula_changes": changes,
                    "external_target": candidate_external_rel.get("Target"),
                }
            )

        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, "w") as out_zip:
            for item in source_zip.infolist():
                out_zip.writestr(item, patches.get(item.filename, source_zip.read(item.filename)))
    return {"passed": True, "output": str(output.resolve()), "charts": details}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--candidate", type=Path, required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--paste-dir", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    args = parser.parse_args()
    result = build(args.source, args.candidate, args.manifest, args.paste_dir, args.output)
    args.result.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"passed": result["passed"], "charts": len(result["charts"]), "output": result["output"]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
