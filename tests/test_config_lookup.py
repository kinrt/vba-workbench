# -*- coding: utf-8 -*-
"""配置查找顺序自测：环境变量优先、向上逐级找到 mapping.json。"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
SYNC = PKG / "scripts" / "sync.py"


def run_cfg(config_path: Path | None, cwd: Path, script: Path | None = None) -> str:
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    if config_path is None:
        env.pop("VBA_WORKBENCH_CONFIG", None)
    else:
        env["VBA_WORKBENCH_CONFIG"] = str(config_path)
    out = subprocess.run([sys.executable, str(script or SYNC), "list"],
                         cwd=str(cwd), env=env, capture_output=True, text=True, encoding="utf-8")
    for line in out.stdout.splitlines():
        if "配置：" in line:
            return line.split("配置：", 1)[1].strip()
    raise AssertionError(f"没打印配置行：{out.stdout}\n{out.stderr}")


class ConfigLookupTest(unittest.TestCase):
    def test_env_var_wins(self):
        with tempfile.TemporaryDirectory() as d:
            cfg = Path(d) / "mapping.json"
            cfg.write_text(json.dumps({
                "defaults": {"workspace": d, "common_dir": "src/common"},
                "projects": [{"name": "示例表A", "search_dir": d, "prefix": "示例表A",
                              "require_version": False, "src": "src/projects/示例表A"}],
            }, ensure_ascii=False), encoding="utf-8")
            got = run_cfg(cfg, Path(d))
            self.assertEqual(Path(got), cfg)

    def test_walk_up_finds_config(self):
        # 查找是「从包所在位置向上」，所以要把 scripts/ 真复制到临时仓库里再跑
        with tempfile.TemporaryDirectory() as d:
            root = Path(d) / "repo"
            pkg = root / "vba-workbench"
            pkg.mkdir(parents=True)
            shutil.copytree(PKG / "scripts", pkg / "scripts")
            cfg = root / "mapping.json"
            cfg.write_text(json.dumps({
                "defaults": {"common_dir": "src/common"},
                "projects": [{"name": "示例表A", "search_dir": d, "prefix": "示例表A",
                              "require_version": False, "src": "src/projects/示例表A"}],
            }, ensure_ascii=False), encoding="utf-8")
            got = run_cfg(None, pkg / "scripts", pkg / "scripts" / "sync.py")
            self.assertEqual(Path(got), cfg)


if __name__ == "__main__":
    unittest.main()
