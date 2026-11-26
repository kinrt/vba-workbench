# -*- coding: utf-8 -*-
"""配置覆盖层 merge_local 的纯函数自测（不碰磁盘、不需要 Excel）。"""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from toolkit.config import merge_local  # noqa: E402


BASE = {
    "defaults": {"common_dir": "src/common", "require_version": True},
    "projects": [
        {"name": "示例表A", "prefix": "示例表A", "src": "src/projects/示例表A"},
        {"name": "示例表B", "prefix": "示例表B", "src": "src/projects/示例表B"},
    ],
}


class MergeLocalTest(unittest.TestCase):
    def test_defaults_override_by_key(self):
        out = merge_local(BASE, {"defaults": {"require_version": False}})
        self.assertFalse(out["defaults"]["require_version"])
        self.assertEqual(out["defaults"]["common_dir"], "src/common")

    def test_project_patch_keeps_other_fields(self):
        out = merge_local(BASE, {"projects": [{"name": "示例表A", "prefix": "新基名"}]})
        a = next(p for p in out["projects"] if p["name"] == "示例表A")
        self.assertEqual(a["prefix"], "新基名")
        self.assertEqual(a["src"], "src/projects/示例表A")

    def test_new_project_appended(self):
        out = merge_local(BASE, {"projects": [{"name": "示例表C"}]})
        self.assertEqual(len(out["projects"]), 3)

    def test_drop_removes_project(self):
        out = merge_local(BASE, {"_drop": ["示例表B"]})
        self.assertEqual([p["name"] for p in out["projects"]], ["示例表A"])

    def test_base_not_mutated(self):
        merge_local(BASE, {"projects": [{"name": "示例表A", "prefix": "x"}]})
        self.assertEqual(BASE["projects"][0]["prefix"], "示例表A")


if __name__ == "__main__":
    unittest.main()
