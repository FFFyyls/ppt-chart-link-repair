# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

"""Build the minimal manifest consumed by isolated Office verification stages."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any


class VerificationManifestError(RuntimeError):
    pass


def _chart_ordinal(chart_id: str) -> int:
    match = re.search(r"-chart-(\d+)$", chart_id)
    if not match or int(match.group(1)) < 1:
        raise VerificationManifestError(
            f"Chart {chart_id} does not contain a valid one-based chart ordinal"
        )
    return int(match.group(1))


def _runtime_series_name(mapping: dict[str, Any], chart_id: str) -> str:
    for key in ("series_name", "runtime_series_name"):
        raw = mapping.get(key)
        if raw is None:
            continue
        value = str(raw).strip()
        if value and value.casefold() not in {"none", "null"}:
            return value
    raise VerificationManifestError(
        f"Chart {chart_id} requires a unique runtime series name before Office verification"
    )


def _group_for_chart(manifest: dict[str, Any], chart_id: str) -> dict[str, Any]:
    groups = [
        group
        for group in manifest.get("groups", [])
        if chart_id in group.get("chart_ids", [])
    ]
    if len(groups) != 1:
        raise VerificationManifestError(
            f"Chart {chart_id} must belong to exactly one workbook group"
        )
    return groups[0]


def _first_numeric_probe(chart: dict[str, Any]) -> dict[str, Any]:
    for mapping in sorted(
        chart.get("data_mappings", []), key=lambda item: int(item["series_index"])
    ):
        cells = list(mapping.get("value_cells", []))
        values = list(mapping.get("values", []))
        indexes = list(mapping.get("cache_indexes", range(len(values))))
        for index in indexes:
            index = int(index)
            if index < 0 or index >= len(values) or index >= len(cells):
                continue
            value = values[index]
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                continue
            return {
                "series_index": int(mapping["series_index"]),
                "series_name": _runtime_series_name(mapping, str(chart["chart_id"])),
                "point_index": index,
                "sheet": str(mapping["sheet"]),
                "cell": str(cells[index]),
                "expected_value": float(value),
            }
    raise VerificationManifestError(
        f"No numeric mutation probe is available for {chart['chart_id']}"
    )


def build_manifest(source: dict[str, Any]) -> dict[str, Any]:
    charts: list[dict[str, Any]] = []
    groups: list[dict[str, Any]] = []
    for group in source.get("groups", []):
        workbook = Path(group["workbook_path"]).resolve()
        if not workbook.exists():
            raise FileNotFoundError(workbook)
        groups.append(
            {
                "workbook_path": str(workbook),
                "chart_ids": list(group.get("chart_ids", [])),
            }
        )
    for chart in source.get("charts", []):
        chart_id = str(chart["chart_id"])
        _group_for_chart(source, chart_id)
        charts.append(
            {
                "chart_id": chart_id,
                "actual_slide": int(chart["actual_slide"]),
                "shape_name": str(chart.get("shape_name", "")),
                "chart_ordinal": _chart_ordinal(chart_id),
                "probe": _first_numeric_probe(chart),
            }
        )
    if not charts:
        raise VerificationManifestError("No target charts were found")
    return {
        "schema_version": "2.0",
        "work_dir": str(Path(source["work_dir"]).resolve()),
        "source_pptx": str(Path(source["source_pptx"]).resolve()),
        "output_pptx": str(Path(source["output_pptx"]).resolve()),
        "groups": groups,
        "charts": charts,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-manifest", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    source = json.loads(args.source_manifest.read_text(encoding="utf-8"))
    result = build_manifest(source)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps({"charts": len(result["charts"])}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
