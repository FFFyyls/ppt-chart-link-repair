# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

from __future__ import annotations

import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[3]
SKILL_ROOT = REPO_ROOT / "repair-ppt-chart-data-links"


class InstallationContractTests(unittest.TestCase):
    def test_standalone_skill_contains_its_python_requirements(self) -> None:
        root_requirements = (REPO_ROOT / "requirements.txt").read_text(encoding="utf-8")
        skill_requirements = (SKILL_ROOT / "requirements.txt").read_text(encoding="utf-8")
        self.assertEqual(root_requirements, skill_requirements)

    def test_readme_documents_three_installation_paths(self) -> None:
        readme = (REPO_ROOT / "README.md").read_text(encoding="utf-8")
        required_fragments = {
            "Codex Skill Installer",
            "install-skill-from-github.py",
            "--repo FFFyyls/ppt-chart-link-repair",
            "--path repair-ppt-chart-data-links",
            "--ref v1.0.0",
            "Git 克隆安装",
            "Release ZIP 安装",
            "repair-ppt-chart-data-links-v1.0.0.zip",
            "一句话交给 Codex 安装",
            "安装完成后告诉我安装目录和验证结果",
        }
        missing = sorted(fragment for fragment in required_fragments if fragment not in readme)
        self.assertEqual([], missing)


if __name__ == "__main__":
    unittest.main()
