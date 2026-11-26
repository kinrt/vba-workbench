# -*- coding: utf-8 -*-
"""随包发布的 mapping.example.json 必须能被真正解析（不是只长得像 JSON）。"""

import json
import os
import subprocess
import sys
import unittest
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
SYNC = PKG / "scripts" / "sync.py"
EXAMPLE = PKG / "assets" / "mapping.example.json"


class ExampleConfigTest(unittest.TestCase):
    def test_example_is_loadable(self):
        env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(EXAMPLE), PYTHONIOENCODING="utf-8")
        r = subprocess.run([sys.executable, str(SYNC), "list", "--json"],
                           cwd=str(PKG), env=env, capture_output=True, text=True, encoding="utf-8")
        payload = json.loads(r.stdout.strip())
        self.assertGreaterEqual(payload["stats"]["total"], 1, "示例配置里至少要有一个项目")
        self.assertEqual(payload["details"][0]["name"], "示例表A")


if __name__ == "__main__":
    unittest.main()
