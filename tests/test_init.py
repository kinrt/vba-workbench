# -*- coding: utf-8 -*-
"""init 的契约：默认只预览、--apply 才写配置与 src 骨架。"""

import hashlib
import json
import os
import shlex
import subprocess
import sys
import tempfile
import time
import types
import unittest
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
SYNC = PKG / "scripts" / "sync.py"
sys.path.insert(0, str(PKG / "scripts"))

from toolkit.base import ROOT  # noqa: E402
from toolkit.context import AppContext, Progress  # noqa: E402
import features.sync_push as sp  # noqa: E402
import features.sync_init as sync_init  # noqa: E402


def make_ctx(args, config: dict) -> AppContext:
    return AppContext(
        tool="sync", args=args, config=config, root=ROOT,
        json_mode=bool(getattr(args, "json", False)),
        started=time.perf_counter(),
        progress=Progress(tool="sync", enabled=False),
    )


def _sha256(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def _write_prod_config(p: Path) -> None:
    p.write_text(json.dumps({
        "defaults": {"common_dir": "src/common"},
        "projects": [{"name": "既有项目", "search_dir": "D:/VBA/既有",
                      "prefix": "既有", "require_version": False,
                      "src": "src/projects/既有"}],
    }, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


class InitTest(unittest.TestCase):
    def run_init(self, *args, cfg: Path):
        env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(cfg), PYTHONIOENCODING="utf-8",
                   PYTHONUTF8="1")
        return subprocess.run([sys.executable, str(SYNC), "init", *args, "--json"],
                              cwd=str(cfg.parent), env=env, capture_output=True,
                              text=True, encoding="utf-8")

    def test_preview_does_not_write(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            cfg = d / "mapping.json"
            table_dir = d / "表"
            table_dir.mkdir()
            r = self.run_init("--project", "甲", "--dir", str(table_dir), "--config", str(cfg), cfg=cfg)
            self.assertEqual(r.returncode, 0, r.stderr)
            payload = json.loads(r.stdout.strip())
            self.assertTrue(payload["ok"])
            self.assertTrue(payload["stats"]["dry"])
            self.assertFalse(cfg.exists())
            self.assertFalse((d / "src" / "projects" / "甲").exists())

    def test_apply_writes_config_and_skeleton(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            cfg = d / "mapping.json"
            table_dir = d / "表"
            table_dir.mkdir()
            r = self.run_init("--project", "甲", "--dir", str(table_dir), "--config", str(cfg),
                             "--apply", cfg=cfg)
            self.assertEqual(r.returncode, 0, r.stderr)
            payload = json.loads(r.stdout.strip())
            self.assertFalse(payload["stats"]["dry"])
            data = json.loads(cfg.read_text(encoding="utf-8"))
            self.assertEqual(data["projects"][0]["name"], "甲")
            self.assertEqual(data["projects"][0]["prefix"], "甲")
            self.assertEqual(data["projects"][0]["search_dir"], table_dir.as_posix())
            self.assertTrue((d / "src" / "projects" / "甲").is_dir())
            self.assertTrue((d / "src" / "common").is_dir())

    def test_reinit_updates_existing_project(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            cfg = d / "mapping.json"
            table_dir = d / "表"
            table_dir.mkdir()
            self.run_init("--project", "甲", "--dir", str(table_dir), "--config", str(cfg),
                          "--apply", cfg=cfg)
            r = self.run_init("--project", "甲", "--dir", str(table_dir), "--config", str(cfg),
                             "--prefix", "甲表", "--apply", cfg=cfg)
            self.assertEqual(r.returncode, 0, r.stderr)
            data = json.loads(cfg.read_text(encoding="utf-8"))
            self.assertEqual(len(data["projects"]), 1)
            self.assertEqual(data["projects"][0]["prefix"], "甲表")

    def test_missing_dir_fails(self):
        with tempfile.TemporaryDirectory() as d:
            cfg = Path(d) / "mapping.json"
            r = self.run_init("--project", "甲", "--dir", str(Path(d) / "没有"), "--config", str(cfg), cfg=cfg)
            self.assertEqual(r.returncode, 2, r.stdout)

    def test_file_without_version_registers_exact_path(self):
        """文件名没有 V<数字>：登记 file 精确路径，不再要求改名或按版本号猜。"""
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            cfg = d / "mapping.json"
            table_dir = d / "表"
            table_dir.mkdir()
            book = table_dir / "示例表B.xlsm"
            book.write_bytes(b"")
            r = self.run_init("--project", "示例表B", "--file", str(book),
                              "--config", str(cfg), "--apply", cfg=cfg)
            self.assertEqual(r.returncode, 0, r.stderr)
            data = json.loads(cfg.read_text(encoding="utf-8"))
            proj = data["projects"][0]
            self.assertEqual(proj["name"], "示例表B")
            self.assertEqual(proj["prefix"], "示例表B")
            self.assertEqual(proj["search_dir"], table_dir.as_posix())
            self.assertEqual(proj["file"], book.as_posix())
            self.assertFalse(proj["require_version"])

    def test_file_with_version_keeps_directory_version_rule(self):
        """文件名带 V<数字>：只登记目录和前缀，留给 resolve_target 自动取最高版本。"""
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            cfg = d / "mapping.json"
            table_dir = d / "表"
            table_dir.mkdir()
            book = table_dir / "示例表A V3.xlsm"
            book.write_bytes(b"")
            r = self.run_init("--project", "示例表A", "--file", str(book),
                              "--config", str(cfg), "--apply", cfg=cfg)
            self.assertEqual(r.returncode, 0, r.stderr)
            proj = json.loads(cfg.read_text(encoding="utf-8"))["projects"][0]
            self.assertEqual(proj["prefix"], "示例表A")
            self.assertEqual(proj["search_dir"], table_dir.as_posix())
            self.assertNotIn("file", proj)
            self.assertTrue(proj["require_version"])

    def test_allow_outside_registers_foreign_workbook(self):
        """菜单拖入外部工作簿：显式 --allow-outside 才允许写进当前 mapping。"""
        with tempfile.TemporaryDirectory() as t1, tempfile.TemporaryDirectory() as t2:
            prod_cfg = Path(t1) / "mapping.json"
            _write_prod_config(prod_cfg)
            book = Path(t2) / "外部表 V1.xlsm"
            book.write_bytes(b"")
            r = self.run_init("--project", "外部表", "--file", str(book),
                              "--allow-outside", "--apply", cfg=prod_cfg)
            self.assertEqual(r.returncode, 0, r.stderr)
            data = json.loads(prod_cfg.read_text(encoding="utf-8"))
            names = [p["name"] for p in data["projects"]]
            self.assertIn("外部表", names)


    def test_lab_template_missing_has_semantic_code(self):
        """安装包缺 make_lab_workbook.py 时，要报 MISSING_LAB_TEMPLATE/init。"""
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            cfg = d / "mapping.json"
            args = types.SimpleNamespace(project="甲", prefix="", lab=True, dir=str(d / ".lab"),
                                         src="", config=str(cfg), apply=True, json=False)
            old_root = sync_init.PKG_ROOT
            sync_init.PKG_ROOT = d   # 模拟安装副本：没有 scripts/make_lab_workbook.py
            try:
                outcome = sp.init(make_ctx(args, {"defaults": {}, "projects": []}))
            finally:
                sync_init.PKG_ROOT = old_root
            err = outcome.errors[0]
            self.assertEqual(err.code, "MISSING_LAB_TEMPLATE")
            self.assertEqual(err.stage, "init")
            self.assertEqual(outcome.exit_code, 1)


    def test_lab_foreign_dir_refused_and_production_untouched(self):
        """env 指向生产配置时，游离目录的 lab init 必须被拒绝，生产 mapping 逐字节不变。"""
        with tempfile.TemporaryDirectory() as t1, tempfile.TemporaryDirectory() as t2:
            prod = Path(t1) / "prod"
            prod.mkdir()
            prod_cfg = prod / "mapping.json"
            _write_prod_config(prod_cfg)
            foreign = Path(t2)
            before = _sha256(prod_cfg)

            r = self.run_init("--project", "临时", "--lab", "--dir", str(foreign),
                              cfg=prod_cfg)
            after = _sha256(prod_cfg)
            self.assertEqual(r.returncode, 2, r.stderr)
            payload = json.loads(r.stdout.strip())
            self.assertFalse(payload["ok"])
            self.assertEqual(payload["errors"][0]["code"], "USAGE")
            self.assertEqual(before, after)
            self.assertFalse((foreign / "mapping.json").exists())

            suggested = payload["errors"][0]["suggested_command"]
            self.assertIn("--config", suggested)
            self.assertIn("--project 临时", suggested)
            self.assertIn((foreign / "mapping.json").as_posix(), suggested)

    def test_explicit_config_writes_lab_config_production_untouched(self):
        """显式 --config 才写 lab 局部配置；生产配置文件逐字节不变（接缝替掉 Excel 生成）。"""
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            prod_cfg = tmp / "prod" / "mapping.json"
            prod_cfg.parent.mkdir()
            _write_prod_config(prod_cfg)
            before = _sha256(prod_cfg)
            lab_dir = tmp / "lab"
            lab_dir.mkdir()
            lab_cfg = lab_dir / "mapping.json"

            called: list[Path] = []

            def fake_generator(script: Path, target: Path) -> None:
                called.append(target)
                target.write_text("lab workbook placeholder", encoding="utf-8")

            old = sync_init._run_lab_generator
            sync_init._run_lab_generator = fake_generator
            self.addCleanup(setattr, sync_init, "_run_lab_generator", old)

            args = types.SimpleNamespace(
                project="临时", prefix="", lab=True, dir=str(lab_dir), src="",
                config=str(lab_cfg), apply=True, json=False)
            outcome = sp.init(make_ctx(args, {"defaults": {}, "projects": []}))

            self.assertEqual(outcome.exit_code, 0, outcome.errors)
            self.assertEqual(Path(outcome.stats["config"]), lab_cfg)
            self.assertTrue(lab_cfg.exists())
            data = json.loads(lab_cfg.read_text(encoding="utf-8"))
            self.assertEqual(data["projects"][0]["name"], "临时")
            self.assertEqual(before, _sha256(prod_cfg))
            self.assertEqual(len(called), 1)

    def test_inside_workspace_apply_still_writes_active_config(self):
        """目标目录在 workspace 内、不给 --config：正常预览/写回，行为不回归。"""
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            cfg = tmp / "mapping.json"
            table_dir = tmp / "表"
            table_dir.mkdir()
            cfg.write_text(json.dumps({
                "defaults": {"workspace": str(tmp), "common_dir": "src/common"},
                "projects": [{"name": "既有项目", "search_dir": str(table_dir),
                              "prefix": "既有", "require_version": False,
                              "src": "src/projects/既有"}],
            }, ensure_ascii=False), encoding="utf-8")

            r = self.run_init("--project", "新表", "--dir", str(table_dir),
                              "--apply", cfg=cfg)
            self.assertEqual(r.returncode, 0, r.stderr)
            payload = json.loads(r.stdout.strip())
            self.assertTrue(payload["ok"])
            data = json.loads(cfg.read_text(encoding="utf-8"))
            names = [p["name"] for p in data["projects"]]
            self.assertIn("新表", names)
            self.assertTrue((tmp / "src" / "projects" / "新表").is_dir())

    def test_refusal_message_carries_executable_config_command(self):
        """拒绝提示必须给出可直接照做的 --config 绝对入口命令，而不是空话或相对路径。"""
        with tempfile.TemporaryDirectory() as t1, tempfile.TemporaryDirectory() as t2:
            prod = Path(t1) / "prod"
            prod.mkdir()
            prod_cfg = prod / "mapping.json"
            _write_prod_config(prod_cfg)
            foreign = Path(t2)

            r = self.run_init("--project", "临时", "--dir", str(foreign),
                              cfg=prod_cfg)
            self.assertEqual(r.returncode, 2, r.stderr)
            payload = json.loads(r.stdout.strip())
            err = payload["errors"][0]
            suggested = err["suggested_command"]
            self.assertIn("--config", suggested)
            self.assertIn("--project 临时", suggested)
            self.assertIn((foreign / "mapping.json").as_posix(), suggested)
            # 绝对入口：含当前解释器，入口脚本真实存在，不是包内相对路径
            self.assertNotIn("vba-workbench/scripts/", suggested)
            self.assertIn(sys.executable, suggested)
            tokens = shlex.split(suggested, posix=False)
            script = Path(tokens[1].strip('"'))
            self.assertTrue(script.is_file(), str(script))
            self.assertEqual(script.name, "sync.py")
            self.assertIn(str(foreign), err["message"] or err["hint"] or suggested)


if __name__ == "__main__":
    unittest.main()
