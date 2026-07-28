#!/usr/bin/env python3
# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

"""Build index-aware recipes for the native Office chart-link pipeline."""

from __future__ import annotations

import copy
import argparse
import json
import re
import shutil
import stat
import subprocess
import sys
import time
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Any

from lxml import etree

from ooxml_security import safe_xml_fromstring, validate_ooxml_archive


C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS = {"c": C}
DATA_ROLES = ("tx", "cat", "val", "xVal", "yVal", "bubbleSize")


class RecipeConflict(RuntimeError):
    """Raised when recovered references cannot map to one workbook safely."""


class NativeLinkError(RuntimeError):
    """Raised when the native Office carrier pipeline fails."""


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _group_for_chart(manifest: dict[str, Any], chart_id: str) -> dict[str, Any]:
    matches = [
        group
        for group in manifest.get("groups", [])
        if chart_id in group.get("chart_ids", [])
    ]
    if len(matches) != 1:
        raise RecipeConflict(
            f"Chart {chart_id} must belong to exactly one workbook group; found {len(matches)}"
        )
    return matches[0]


def _cache_indexes(ref: dict[str, Any]) -> list[int]:
    values = list(ref.get("values", []))
    explicit = ref.get("cache_indexes")
    if explicit is not None:
        indexes = [int(value) for value in explicit]
    else:
        indexes = [index for index, value in enumerate(values) if value is not None]
    if len(indexes) != len(set(indexes)) or any(index < 0 for index in indexes):
        raise RecipeConflict(f"Invalid cache indexes for {ref.get('formula')!r}: {indexes}")
    return indexes


def _materialized_points(ref: dict[str, Any]) -> list[dict[str, Any]]:
    values = list(ref.get("values", []))
    cells = list(ref.get("cells", []))
    indexes = _cache_indexes(ref)
    points: list[dict[str, Any]] = []
    for index in indexes:
        if index >= len(values):
            raise RecipeConflict(
                f"Cache index {index} exceeds values for {ref.get('formula')!r}"
            )
        cell = None
        if cells:
            if index >= len(cells):
                raise RecipeConflict(
                    f"Cache index {index} exceeds formula cells for {ref.get('formula')!r}"
                )
            cell = cells[index]
        points.append({"index": index, "cell": cell, "value": values[index]})
    return points


def _plot_metadata(chart_xml: bytes) -> tuple[list[dict[str, Any]], dict[int, dict[str, Any]]]:
    root = safe_xml_fromstring(chart_xml)
    axis_positions: dict[int, str] = {}
    for axis_name in ("catAx", "valAx", "dateAx", "serAx"):
        for axis in root.findall(f".//c:{axis_name}", NS):
            axis_id = axis.find("c:axId", NS)
            axis_pos = axis.find("c:axPos", NS)
            if axis_id is not None and axis_id.get("val") is not None:
                axis_positions[int(axis_id.get("val"))] = (
                    axis_pos.get("val") if axis_pos is not None else ""
                )
    plots: list[dict[str, Any]] = []
    by_position: dict[int, dict[str, Any]] = {}
    fallback_order = 0
    series_position = 0
    for plot in root.findall(".//c:plotArea/*", NS):
        plot_type = _local_name(plot.tag)
        if not plot_type.endswith("Chart"):
            continue
        bar_dir = plot.find("c:barDir", NS)
        grouping = plot.find("c:grouping", NS)
        axis_ids = [
            int(node.get("val"))
            for node in plot.findall("c:axId", NS)
            if node.get("val") is not None
        ]
        axis_group = 2 if any(axis_positions.get(axis_id) in {"r", "t"} for axis_id in axis_ids) else 1
        normalized_bar_dir = bar_dir.get("val") if bar_dir is not None else None
        normalized_grouping = grouping.get("val") if grouping is not None else None
        if plot_type == "barChart":
            normalized_bar_dir = normalized_bar_dir or "col"
            normalized_grouping = normalized_grouping or "clustered"
        elif plot_type in {"lineChart", "areaChart"}:
            normalized_grouping = normalized_grouping or "standard"
        plot_info = {
            "type": plot_type,
            "bar_dir": normalized_bar_dir,
            "grouping": normalized_grouping,
            "axis_ids": axis_ids,
            "axis_group": axis_group,
            "series_orders": [],
        }
        for series in plot.findall("c:ser", NS):
            order_node = series.find("c:order", NS)
            idx_node = series.find("c:idx", NS)
            order = int(order_node.get("val")) if order_node is not None else fallback_order
            series_idx = int(idx_node.get("val")) if idx_node is not None else order
            fallback_order = max(fallback_order + 1, order + 1)
            metadata = {
                "order": order,
                "idx": series_idx,
                "plot_type": plot_type,
                "bar_dir": plot_info["bar_dir"],
                "grouping": plot_info["grouping"],
                "axis_ids": axis_ids,
                "axis_group": axis_group,
            }
            by_position[series_position] = metadata
            series_position += 1
            plot_info["series_orders"].append(order)
        plots.append(plot_info)
    return plots, by_position


def _reference_recipe(ref: dict[str, Any]) -> dict[str, Any]:
    result = copy.deepcopy(ref)
    formula = str(result.get("formula", ""))
    if "!" not in formula:
        result.setdefault("defined_name", formula.lstrip("="))
        result["parsed"] = {"kind": "defined-name", "name": result["defined_name"]}
    else:
        left, address = formula.rsplit("!", 1)
        left = left.strip().strip("'").replace("''", "'")
        workbook_match = re.search(r"\[([^\]]+)\]", left)
        sheet = left[workbook_match.end() :] if workbook_match else left
        result["parsed"] = {
            "kind": "range",
            "workbook": workbook_match.group(1) if workbook_match else None,
            "sheet": result.get("sheet", sheet),
            "address": result.get("address", address.replace("$", "")),
        }
    result["cache_indexes"] = _cache_indexes(result)
    result["materialized_points"] = _materialized_points(result)
    return result


def build_chart_recipes(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    """Return index-aware, series-by-series Office carrier recipes."""

    source_pptx = manifest.get("source_pptx")
    archive: zipfile.ZipFile | None = None
    if source_pptx and Path(source_pptx).is_file():
        archive = zipfile.ZipFile(source_pptx)
        validate_ooxml_archive(archive)
    recipes: list[dict[str, Any]] = []
    ledger: dict[tuple[str, str, str], tuple[Any, str]] = {}
    try:
        for chart in manifest.get("charts", []):
            chart_id = str(chart["chart_id"])
            group = _group_for_chart(manifest, chart_id)
            plots: list[dict[str, Any]] = []
            plot_by_position: dict[int, dict[str, Any]] = {}
            if archive is not None and chart.get("chart_part") in archive.namelist():
                plots, plot_by_position = _plot_metadata(archive.read(chart["chart_part"]))

            refs_by_series: dict[int, dict[str, dict[str, Any]]] = defaultdict(dict)
            for raw_ref in chart.get("references", []):
                role = str(raw_ref.get("role", ""))
                if role not in DATA_ROLES:
                    continue
                series_index = int(raw_ref.get("series_index", 0))
                ref = _reference_recipe(raw_ref)
                if role in refs_by_series[series_index]:
                    raise RecipeConflict(
                        f"Duplicate {role} reference for {chart_id} series {series_index}"
                    )
                refs_by_series[series_index][role] = ref
                sheet = ref.get("sheet")
                if sheet:
                    for point in ref["materialized_points"]:
                        if point["cell"] is None:
                            continue
                        key = (str(Path(group["workbook_path"]).resolve()), str(sheet), point["cell"])
                        typed = (point["value"], str(ref.get("cache_type", "")))
                        if key in ledger and ledger[key] != typed:
                            raise RecipeConflict(
                                f"Conflicting recovered values at {sheet}!{point['cell']}"
                            )
                        ledger[key] = typed

            series: list[dict[str, Any]] = []
            for series_index, refs in sorted(refs_by_series.items()):
                metadata = plot_by_position.get(series_index, {})
                series.append(
                    {
                        "series_index": series_index,
                        "order": int(metadata.get("order", series_index)),
                        "idx": int(metadata.get("idx", series_index)),
                        "plot_type": metadata.get("plot_type"),
                        "bar_dir": metadata.get("bar_dir"),
                        "grouping": metadata.get("grouping"),
                        "axis_ids": list(metadata.get("axis_ids", [])),
                        "axis_group": int(metadata.get("axis_group", 1)),
                        "refs": refs,
                    }
                )
            recipes.append(
                {
                    "chart_id": chart_id,
                    "actual_slide": int(chart["actual_slide"]),
                    "shape_name": chart.get("shape_name", ""),
                    "chart_part": chart.get("chart_part"),
                    "chart_rel_part": chart.get("chart_rel_part"),
                    "chart_types": list(chart.get("chart_types", [])),
                    "plots": plots,
                    "series": series,
                    "workbook_path": str(Path(group["workbook_path"]).resolve()),
                }
            )
    finally:
        if archive is not None:
            archive.close()
    return recipes


def write_chart_recipes(manifest_path: str | Path, output_path: str | Path) -> list[dict[str, Any]]:
    manifest = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
    recipes = build_chart_recipes(manifest)
    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(
        json.dumps({"charts": recipes}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return recipes


def _run(command: list[str], timeout: int) -> subprocess.CompletedProcess[bytes]:
    completed = subprocess.run(command, capture_output=True, text=False, timeout=timeout)
    if completed.returncode not in {0, 20}:
        stdout = completed.stdout.decode("utf-8", errors="replace")
        stderr = completed.stderr.decode("utf-8", errors="replace")
        raise NativeLinkError(
            f"Command failed ({completed.returncode}): {' '.join(command)}\n{stdout}\n{stderr}"
        )
    return completed


def _clean_carrier_sheets(
    scripts: Path,
    recipes_path: Path,
    support_map_path: Path,
    result_path: Path,
) -> None:
    if not support_map_path.exists():
        return
    _run(
        [
            sys.executable,
            str(scripts / "build_carrier_workbooks.py"),
            "clean",
            "--recipes",
            str(recipes_path),
            "--support-map",
            str(support_map_path),
            "--result",
            str(result_path),
        ],
        timeout=120,
    )


def _paste_and_transplant(
    *,
    scripts: Path,
    source: Path,
    output: Path,
    candidate: Path,
    recipes: list[dict[str, Any]],
    recipes_path: Path,
    support_map_path: Path,
    paste_dir: Path,
    transplant_result: Path,
    stage_status: Path,
) -> dict[str, Any]:
    if candidate.exists():
        candidate.chmod(stat.S_IREAD | stat.S_IWRITE)
        candidate.unlink()
    shutil.copyfile(source, candidate)
    batch_result = paste_dir.parent / "native-paste-batch.json"
    completed = _run(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(scripts / "office_native_paste_batch.ps1"),
            "-ManifestPath",
            str(recipes_path),
            "-SupportMapPath",
            str(support_map_path),
            "-CandidatePptx",
            str(candidate),
            "-PasteDir",
            str(paste_dir),
            "-BatchResultPath",
            str(batch_result),
        ],
        timeout=max(300, 60 * max(1, len(recipes))),
    )
    if completed.returncode == 20:
        stage_status.write_text(
            json.dumps(
                {
                    "stage": "awaiting-powerpoint-close",
                    "remaining": len(recipes),
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        return {
            "passed": False,
            "requires_user_action": True,
            "message": "PowerPoint正在运行，请保存并关闭所有PowerPoint窗口后继续",
            "stage": "native-paste",
        }
    _run(
        [
            sys.executable,
            str(scripts / "transplant_native_links.py"),
            "--source",
            str(source),
            "--candidate",
            str(candidate),
            "--manifest",
            str(recipes_path),
            "--paste-dir",
            str(paste_dir),
            "--output",
            str(output),
            "--result",
            str(transplant_result),
        ],
        timeout=180,
    )
    result = json.loads(transplant_result.read_text(encoding="utf-8"))
    if batch_result.exists():
        result["native_paste_batch"] = json.loads(
            batch_result.read_text(encoding="utf-8")
        )
    return result


def _valid_series_name(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.casefold() in {"none", "null"}:
        return None
    return text


def enrich_runtime_series_names(
    manifest: dict[str, Any], paste_results: dict[str, dict[str, Any]]
) -> None:
    """Fill absent cache titles only when COM exposes one unique unmatched name."""

    for chart in manifest.get("charts", []):
        chart_id = str(chart["chart_id"])
        captured = paste_results.get(chart_id, {})
        captured_names = list(captured.get("runtime_series_names", []))
        runtime_names = [
            name
            for raw in captured_names
            if (name := _valid_series_name(raw)) is not None
        ]
        if len(runtime_names) != len(captured_names):
            raise NativeLinkError(
                f"Chart {chart_id} has a missing runtime series name"
            )
        if len(set(runtime_names)) != len(runtime_names):
            raise NativeLinkError(
                f"Chart {chart_id} has duplicate runtime series names: {runtime_names}"
            )
        mappings = list(chart.get("data_mappings", []))
        known_names = [
            name
            for mapping in mappings
            for raw in (mapping.get("series_name"),)
            if (name := _valid_series_name(raw)) is not None
        ]
        if len(set(known_names)) != len(known_names):
            raise NativeLinkError(
                f"Chart {chart_id} has duplicate cached series names: {known_names}"
            )
        known = set(known_names)
        if len(runtime_names) != len(mappings) or not known.issubset(runtime_names):
            raise NativeLinkError(
                f"Chart {chart_id} requires a unique runtime series mapping and "
                f"does not match the manifest; "
                f"runtime={runtime_names}, cached={known_names}, mappings={len(mappings)}"
            )
        missing = [
            mapping
            for mapping in mappings
            if _valid_series_name(mapping.get("series_name")) is None
        ]
        candidates = [name for name in runtime_names if name not in known]
        if not missing:
            chart["runtime_series_names"] = runtime_names
            continue
        if len(missing) != 1 or len(candidates) != 1:
            raise NativeLinkError(
                f"Chart {chart_id} requires a unique runtime series name; "
                f"missing={len(missing)}, candidates={candidates}"
            )
        missing[0]["runtime_series_name"] = candidates[0]
        chart["runtime_series_names"] = runtime_names


def capture_final_runtime_series_names(
    scripts: Path, recipes_path: Path, output: Path, result_path: Path
) -> dict[str, Any]:
    """Read COM series names from the transplanted, saved output presentation."""

    completed = _run(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(scripts / "office_capture_series_names.ps1"),
            "-ManifestPath",
            str(recipes_path),
            "-Pptx",
            str(output),
            "-ResultPath",
            str(result_path),
        ],
        timeout=180,
    )
    if not result_path.exists():
        raise NativeLinkError("Final runtime-series capture did not produce a result")
    result = json.loads(result_path.read_text(encoding="utf-8"))
    if completed.returncode == 20 or result.get("requires_user_action"):
        return result
    if not result.get("passed"):
        raise NativeLinkError("Final runtime-series capture returned passed=false")
    return result


def run_native_relink(manifest_path: str | Path) -> dict[str, Any]:
    """Execute the only supported native Office link-rebuild pipeline."""

    started_at = time.perf_counter()
    path = Path(manifest_path).resolve()
    manifest = json.loads(path.read_text(encoding="utf-8"))
    if manifest.get("selection_confirmed") is not True:
        raise NativeLinkError("selection_confirmed must be true before relinking")
    if not manifest.get("relink_allowed"):
        raise NativeLinkError("Relink is blocked by incomplete or conflicting cache data")
    source = Path(manifest["source_pptx"]).resolve()
    output = Path(manifest["output_pptx"]).resolve()
    if source == output:
        raise NativeLinkError("Output must not overwrite the source presentation")
    work = Path(manifest["work_dir"]).resolve() / "native-link"
    work.mkdir(parents=True, exist_ok=True)
    recipes_path = work / "native-link-recipes.json"
    support_map_path = work / "support-map.json"
    candidate = work / "office-link-candidate.pptx"
    paste_dir = work / "paste-results"
    paste_dir.mkdir(parents=True, exist_ok=True)
    transplant_result = work / "transplant-result.json"
    stage_status = work / "stage-status.json"

    recipes = build_chart_recipes(manifest)
    recipe_manifest = {
        "schema_version": "2.0",
        "source_pptx": str(source),
        "output_pptx": str(output),
        "work_dir": str(work),
        "charts": recipes,
        "groups": manifest.get("groups", []),
    }
    recipes_path.write_text(
        json.dumps(recipe_manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    stage_status.write_text(
        json.dumps({"stage": "recipes-ready", "charts": len(recipes)}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    scripts = Path(__file__).resolve().parent
    if support_map_path.exists():
        _clean_carrier_sheets(
            scripts,
            recipes_path,
            support_map_path,
            work / "stale-carrier-cleanup.json",
        )
        support_map_path.unlink(missing_ok=True)
    carrier = _run(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(scripts / "office_native_carrier.ps1"),
            "-ManifestPath",
            str(recipes_path),
            "-ResultPath",
            str(support_map_path),
        ],
        timeout=300,
    )
    if carrier.returncode == 20:
        return {
            "passed": False,
            "requires_user_action": True,
            "message": "PowerPoint正在运行，请保存并关闭所有PowerPoint窗口后继续",
            "stage": "carrier",
        }

    clean_result_path = work / "carrier-cleanup.json"
    final_series_result_path = work / "final-series-names.json"
    try:
        result = _paste_and_transplant(
            scripts=scripts,
            source=source,
            output=output,
            candidate=candidate,
            recipes=recipes,
            recipes_path=recipes_path,
            support_map_path=support_map_path,
            paste_dir=paste_dir,
            transplant_result=transplant_result,
            stage_status=stage_status,
        )
    finally:
        _clean_carrier_sheets(
            scripts, recipes_path, support_map_path, clean_result_path
        )
    if result.get("requires_user_action"):
        result["elapsed_seconds"] = round(time.perf_counter() - started_at, 3)
        return result
    final_series_capture = capture_final_runtime_series_names(
        scripts, recipes_path, output, final_series_result_path
    )
    if final_series_capture.get("requires_user_action"):
        result.update(final_series_capture)
        result["elapsed_seconds"] = round(time.perf_counter() - started_at, 3)
        return result
    final_series_by_chart = {
        str(item["chart_id"]): item
        for item in final_series_capture.get("charts", [])
    }
    enrich_runtime_series_names(manifest, final_series_by_chart)
    result["final_series_capture"] = final_series_capture
    result["elapsed_seconds"] = round(time.perf_counter() - started_at, 3)
    for chart in manifest.get("charts", []):
        chart["relink_method"] = (
            "native Excel chart carrier + Office PasteSourceFormatting + OOXML relationship transplant"
        )
        chart["relink_object_kind"] = "native-powerpoint-chart"
    manifest["native_relink"] = result
    manifest["native_relink"]["recipe_manifest"] = str(recipes_path)
    manifest["native_relink"]["support_map"] = str(support_map_path)
    path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    stage_status.write_text(
        json.dumps({"stage": "native-relink-complete", "passed": True}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", required=True)
    args = parser.parse_args()
    try:
        result = run_native_relink(args.manifest)
    except (NativeLinkError, RecipeConflict, FileNotFoundError, ValueError) as exc:
        print(json.dumps({"passed": False, "error": str(exc)}, ensure_ascii=False))
        return 2
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("passed") else (20 if result.get("requires_user_action") else 1)


if __name__ == "__main__":
    raise SystemExit(main())
