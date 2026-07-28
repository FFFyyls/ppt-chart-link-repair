# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

from __future__ import annotations

import subprocess
import sys
import unittest
from pathlib import Path


SKILL_ROOT = Path(__file__).resolve().parents[2]
SCRIPTS = SKILL_ROOT / "scripts"


class V09ProductionContractTests(unittest.TestCase):
    def test_old_relink_engine_is_absent(self) -> None:
        production_sources = "\n".join(
            path.read_text(encoding="utf-8")
            for path in SCRIPTS.iterdir()
            if path.suffix.lower() in {".py", ".ps1"}
        )
        self.assertNotIn("Chart.SetSourceData", production_sources)
        self.assertNotIn("use_linked_excel_fallback", production_sources)

    def test_new_engine_files_exist(self) -> None:
        required = {
            "native_link_pipeline.py",
            "build_carrier_workbooks.py",
            "office_native_carrier.ps1",
            "transplant_native_links.py",
            "office_stage_worker.ps1",
            "portable_bundle.py",
            "portable_rebind.ps1",
        }
        self.assertEqual([], sorted(name for name in required if not (SCRIPTS / name).is_file()))

    def test_office_gate_has_explicit_user_message(self) -> None:
        source = (SCRIPTS / "office_verify.ps1").read_text(encoding="utf-8")
        self.assertIn("PowerPoint正在运行", source)
        self.assertIn("请保存并关闭所有PowerPoint窗口后继续", source)
        self.assertIn("exit 20", source)

    def test_cli_exposes_portable_package(self) -> None:
        completed = subprocess.run(
            [sys.executable, str(SCRIPTS / "chart_recovery.py"), "--help"],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        self.assertEqual(0, completed.returncode, completed.stderr)
        self.assertIn("portable-package", completed.stdout)

    def test_skill_documents_default_and_portable_modes(self) -> None:
        content = (SKILL_ROOT / "SKILL.md").read_text(encoding="utf-8")
        self.assertIn("默认本机重连", content)
        self.assertIn("可迁移数据包", content)
        self.assertNotIn("新 Excel 使用当前机器的绝对路径；", content)

    def test_skill_documents_combo_series_runtime_matching(self) -> None:
        content = (SKILL_ROOT / "SKILL.md").read_text(encoding="utf-8")
        standard = (SKILL_ROOT / "references" / "verification-standard.md").read_text(
            encoding="utf-8"
        )
        self.assertIn("系列名称", content)
        self.assertIn("SeriesCollection", standard)
        self.assertIn("重复系列名称", standard)

    def test_support_matrix_distinguishes_tested_from_unverified(self) -> None:
        matrix = (SKILL_ROOT / "references" / "supported-chart-types.md").read_text(
            encoding="utf-8"
        )
        self.assertIn("绿色（实机通过）", matrix)
        self.assertIn("黄色（尚未充分实机验证）", matrix)
        self.assertIn("红色（不支持）", matrix)

    def test_portable_mode_requires_local_office_success_first(self) -> None:
        content = (SKILL_ROOT / "SKILL.md").read_text(encoding="utf-8")
        self.assertIn("本机 Office 硬验收已经全部通过", content)


if __name__ == "__main__":
    unittest.main()
