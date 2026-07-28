# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

"""Create an optional self-contained cross-computer chart-link data bundle."""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import shutil
import tempfile
import uuid
from pathlib import Path
from typing import Any


BUNDLE_MARKER = ".ppt-chart-link-repair-bundle"


class PortableBundleError(RuntimeError):
    pass


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _safe_name(value: str) -> str:
    forbidden = '<>:"/\\|?*'
    result = "".join("_" if char in forbidden else char for char in value).strip(" .")
    return result or "PPT图表数据包"


def _first_probe(chart: dict[str, Any]) -> dict[str, Any]:
    for mapping in sorted(
        chart.get("data_mappings", []), key=lambda item: int(item["series_index"])
    ):
        cells = list(mapping.get("value_cells", []))
        values = list(mapping.get("values", []))
        indexes = list(mapping.get("cache_indexes", range(len(values))))
        for raw_index in indexes:
            index = int(raw_index)
            if index < 0 or index >= len(cells) or index >= len(values):
                continue
            value = values[index]
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                continue
            return {
                "series_index": int(mapping["series_index"]),
                "series_name": str(mapping.get("series_name", "")),
                "point_index": index,
                "sheet": str(mapping["sheet"]),
                "cell": str(cells[index]),
                "expected_value": float(value),
            }
    raise PortableBundleError(f"No numeric probe available for {chart['chart_id']}")


def _build_bundle_in_root(
    manifest_path: Path, output_root: Path | None = None
) -> dict[str, Any]:
    manifest_path = manifest_path.resolve()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    source_pptx = Path(manifest["output_pptx"]).resolve()
    if not source_pptx.exists():
        raise FileNotFoundError(source_pptx)
    native = manifest.get("native_relink", {})
    if not native.get("passed"):
        raise PortableBundleError("Native relink must pass before portable packaging")
    if not manifest.get("charts"):
        raise PortableBundleError("No repaired charts are available for packaging")
    office_report_path = Path(manifest["work_dir"]).resolve() / "office-verification.json"
    if not office_report_path.exists():
        raise PortableBundleError("Office verification must pass before portable packaging")
    office_report = json.loads(office_report_path.read_text(encoding="utf-8"))
    verified_ids = {
        str(item.get("chart_id"))
        for item in office_report.get("charts", [])
        if item.get("passed") is True
        and item.get("workbook_restored_exactly") is True
    }
    selected_ids = {str(chart["chart_id"]) for chart in manifest["charts"]}
    if office_report.get("passed") is not True or verified_ids != selected_ids:
        raise PortableBundleError(
            "Office verification must pass for every repaired chart before portable packaging"
        )

    root = (
        output_root.resolve()
        if output_root
        else Path(manifest["work_dir"]).resolve() / "portable-bundle"
    )
    bundle_name = _safe_name(f"{source_pptx.stem}-图表数据包")
    bundle_dir = root / bundle_name
    if bundle_dir.exists():
        shutil.rmtree(bundle_dir)
    data_dir = bundle_dir / "data"
    data_dir.mkdir(parents=True)

    master_name = _safe_name(f"{source_pptx.stem}-可迁移母版.pptx")
    master = bundle_dir / master_name
    shutil.copy2(source_pptx, master)

    workbook_map: dict[str, dict[str, Any]] = {}
    for group in manifest.get("groups", []):
        source_workbook = Path(group["workbook_path"]).resolve()
        if not source_workbook.exists():
            raise FileNotFoundError(source_workbook)
        destination = data_dir / source_workbook.name
        if destination.exists() and sha256_file(destination) != sha256_file(source_workbook):
            raise PortableBundleError(f"Workbook name collision: {destination.name}")
        shutil.copy2(source_workbook, destination)
        workbook_map[str(source_workbook)] = {
            "relative_path": destination.relative_to(bundle_dir).as_posix(),
            "initial_sha256": sha256_file(destination),
            "chart_ids": list(group.get("chart_ids", [])),
        }

    charts: list[dict[str, Any]] = []
    for chart in manifest["charts"]:
        group = next(
            group
            for group in manifest["groups"]
            if chart["chart_id"] in group.get("chart_ids", [])
        )
        workbook = workbook_map[str(Path(group["workbook_path"]).resolve())]
        charts.append(
            {
                "chart_id": chart["chart_id"],
                "actual_slide": int(chart["actual_slide"]),
                "shape_name": chart.get("shape_name", ""),
                "chart_part": chart.get("chart_part"),
                "chart_rel_part": chart.get("chart_rel_part"),
                "workbook_relative_path": workbook["relative_path"],
                "probe": _first_probe(chart),
            }
        )

    created_at = dt.datetime.now(dt.timezone.utc).isoformat()
    bundle_fingerprint = hashlib.sha256(
        (
            created_at
            + sha256_file(master)
            + "".join(
                sorted(item["initial_sha256"] for item in workbook_map.values())
            )
        ).encode("utf-8")
    ).hexdigest()[:24]
    bundle_manifest = {
        "schema_version": "1.0",
        "bundle_id": bundle_fingerprint,
        "created_at": created_at,
        "mode": "portable-rebind",
        "unsupported_locations": [
            "OneDrive",
            "SharePoint",
            "UNC/network paths",
            "mapped network drives",
        ],
        "master_pptx": master.name,
        "master_sha256": sha256_file(master),
        "source_office_verification_sha256": sha256_file(office_report_path),
        "local_output_pptx": _safe_name(f"{source_pptx.stem}-本机链接版.pptx"),
        "workbooks": list(workbook_map.values()),
        "charts": charts,
    }
    (bundle_dir / "bundle-manifest.json").write_text(
        json.dumps(bundle_manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    scripts = Path(__file__).resolve().parent
    shutil.copy2(scripts / "portable_rebind.ps1", bundle_dir / "本机重绑.ps1")
    launcher = bundle_dir / "启动并重绑PPT图表.cmd"
    launcher.write_text(
        "@echo off\r\n"
        "chcp 65001 >nul\r\n"
        'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0本机重绑.ps1"\r\n'
        "set EXITCODE=%ERRORLEVEL%\r\n"
        "if not %EXITCODE%==0 pause\r\n"
        "exit /b %EXITCODE%\r\n",
        encoding="utf-8-sig",
    )
    readme = bundle_dir / "使用说明.txt"
    readme.write_text(
        "1. 将整个文件夹复制到目标电脑的本地磁盘，不要拆散文件。\n"
        "2. 不支持 OneDrive、SharePoint、UNC/网络路径和映射网络驱动器。\n"
        "3. 保存并关闭所有 PowerPoint 窗口；Excel 可继续使用，但重绑测试期间可能短暂显示数据工作簿。\n"
        "4. 双击“启动并重绑PPT图表.cmd”。程序会从母版生成本机链接版，不覆盖母版。\n"
        "5. 文件夹移动或改名后，再次运行启动器即可按新路径重绑。\n",
        encoding="utf-8-sig",
    )
    result = {
        "passed": True,
        "bundle_dir": str(bundle_dir),
        "master_pptx": str(master),
        "workbooks": len(workbook_map),
        "charts": len(charts),
    }
    (bundle_dir / "package-result.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return result


def _remove_managed_directory(path: Path, parent: Path, prefix: str) -> None:
    resolved = path.resolve()
    expected_parent = parent.resolve()
    if resolved.parent != expected_parent or not resolved.name.startswith(prefix):
        raise PortableBundleError(f"Refusing to remove unmanaged directory: {resolved}")
    if resolved.exists():
        shutil.rmtree(resolved)


def _publish_staged_bundle(staged: Path, final: Path) -> None:
    root = final.parent.resolve()
    backup: Path | None = None
    if final.exists():
        if not (final / BUNDLE_MARKER).is_file():
            raise PortableBundleError(
                f"Existing bundle directory was not created by this tool: {final}"
            )
        backup = root / f".{final.name}.backup-{uuid.uuid4().hex}"
        final.replace(backup)
    try:
        staged.replace(final)
    except Exception:
        if backup is not None and backup.exists() and not final.exists():
            backup.replace(final)
        raise
    if backup is not None:
        _remove_managed_directory(backup, root, f".{final.name}.backup-")


def build_bundle(manifest_path: Path, output_root: Path | None = None) -> dict[str, Any]:
    manifest_path = manifest_path.resolve()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    source_pptx = Path(manifest["output_pptx"]).resolve()
    root = (
        output_root.resolve()
        if output_root
        else Path(manifest["work_dir"]).resolve() / "portable-bundle"
    )
    root.mkdir(parents=True, exist_ok=True)
    bundle_name = _safe_name(f"{source_pptx.stem}-图表数据包")
    final = root / bundle_name
    if final.exists() and not (final / BUNDLE_MARKER).is_file():
        raise PortableBundleError(
            f"Existing bundle directory was not created by this tool: {final}"
        )

    staging_root = Path(
        tempfile.mkdtemp(prefix=f".{bundle_name}.staging-", dir=root)
    )
    try:
        result = _build_bundle_in_root(manifest_path, staging_root)
        staged = Path(result["bundle_dir"])
        final_master = final / Path(result["master_pptx"]).name
        final_result = {
            **result,
            "bundle_dir": str(final),
            "master_pptx": str(final_master),
        }
        (staged / BUNDLE_MARKER).write_text(
            "ppt-chart-link-repair portable bundle\n", encoding="utf-8"
        )
        (staged / "package-result.json").write_text(
            json.dumps(final_result, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        _publish_staged_bundle(staged, final)
        return final_result
    finally:
        if staging_root.exists():
            _remove_managed_directory(
                staging_root, root, f".{bundle_name}.staging-"
            )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path)
    args = parser.parse_args()
    result = build_bundle(args.manifest, args.output_dir)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
