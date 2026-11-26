# -*- coding: utf-8 -*-
"""入口契约：--version、裸跑退出 2、菜单只认 -i。"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
SYNC = PKG / "scripts" / "sync.py"
AUDIT = PKG / "scripts" / "audit.py"
AGENT = PKG / "scripts" / "agent.py"
VERSION = (PKG / "VERSION").read_text(encoding="utf-8").strip()


class CliContractTest(unittest.TestCase):
    def run_cli(self, script, *args):
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
        return subprocess.run([sys.executable, str(script), *args],
                              cwd=str(PKG.parent), env=env,
                              capture_output=True, text=True, encoding="utf-8")

    def test_version_prints_package_version(self):
        for script in (SYNC, AUDIT, AGENT):
            with self.subTest(script=script.name):
                r = self.run_cli(script, "--version")
                self.assertEqual(r.returncode, 0, r.stderr)
                self.assertIn(f"vba-workbench {VERSION}", r.stdout)

    def test_sync_bare_prints_help_and_exits_2(self):
        r = self.run_cli(SYNC)
        self.assertEqual(r.returncode, 2)
        self.assertIn("--apply", r.stdout)
        self.assertNotIn('"action"', r.stdout)

    def test_project_and_all_project_are_mutually_exclusive(self):
        """--project 与 --all-project 同用必须在解析层退出 2，不能静默选一边。"""
        for verb in ("pull", "push", "gbk-export"):
            with self.subTest(verb=verb):
                r = self.run_cli(SYNC, verb, "--project", "不存在",
                                 "--all-project", "--json")
                self.assertEqual(r.returncode, 2, r.stderr)
                self.assertIn("--all-project", r.stderr)
                self.assertNotIn('"ok": true', r.stdout)

    def test_audit_bare_prints_help_and_exits_2(self):
        r = self.run_cli(AUDIT)
        self.assertEqual(r.returncode, 2)
        self.assertIn("-i", r.stdout)

    def test_audit_business_failure_is_not_silent(self):
        """audit 子命令业务失败必须像 sync/agent 一样在人话模式报错误。"""
        with tempfile.TemporaryDirectory() as d:
            cfg = Path(d) / "mapping.json"
            cfg.write_text(json.dumps({
                "defaults": {"common_dir": "src/common"},
                "projects": [{"name": "示例表A", "search_dir": str(d),
                              "prefix": "示例表A", "require_version": False,
                              "src": "src/projects/示例表A"}],
            }, ensure_ascii=False), encoding="utf-8")
            env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(cfg),
                       PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
            r = subprocess.run(
                [sys.executable, str(AUDIT), "project", "--project", "不存在",
                 "--no-input"],
                cwd=str(PKG.parent), env=env, capture_output=True, text=True,
                encoding="utf-8")
            self.assertEqual(r.returncode, 1)
            self.assertIn("[错误]", r.stderr)
            self.assertTrue(r.stderr.strip(), "业务失败不能 stdout/stderr 全空")
            self.assertNotIn("Traceback", r.stderr)

    def test_menu_requires_explicit_i(self):
        # 没有 -i 但 stdin 是管道：也必须退出 2，不能进菜单
        r = self.run_cli(SYNC, "--no-input")
        self.assertEqual(r.returncode, 2)

    def test_no_input_flag_is_defined(self):
        """--no-input 必须能真正置位；不能因为漏 import 在入口就 NameError。"""
        with tempfile.TemporaryDirectory() as d:
            cfg = Path(d) / "mapping.json"
            cfg.write_text(json.dumps({
                "defaults": {"common_dir": "src/common"},
                "projects": [{"name": "示例表A", "search_dir": str(d),
                              "prefix": "示例表A", "require_version": False,
                              "src": "src/projects/示例表A"}],
            }, ensure_ascii=False), encoding="utf-8")
            (Path(d) / "示例表A V1.xlsm").write_bytes(b"")
            env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(cfg),
                       PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
            r = subprocess.run([sys.executable, str(SYNC), "list", "--no-input"],
                               cwd=str(PKG.parent), env=env, capture_output=True,
                               text=True, encoding="utf-8")
            self.assertEqual(r.returncode, 0, r.stderr)
            self.assertNotIn("Traceback", r.stderr)

            r = subprocess.run([sys.executable, str(AUDIT), "members", "--no-input",
                                "--module", "不存在的模块"],
                               cwd=str(PKG.parent), env=env, capture_output=True,
                               text=True, encoding="utf-8")
            self.assertEqual(r.returncode, 1, r.stderr)
            self.assertNotIn("Traceback", r.stderr)

    def test_watch_help_lists_no_backup(self):
        r = self.run_cli(SYNC, "watch", "--help")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn("--no-backup", r.stdout)

    def test_agent_logs_help_lists_tail(self):
        r = self.run_cli(AGENT, "logs", "--help")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn("--tail", r.stdout)
        self.assertIn("40", r.stdout)

    def test_interactive_non_tty_is_clean(self):
        """audit -i 在管道里只出人话，不能留 Failure 的 traceback。"""
        with tempfile.TemporaryDirectory() as d:
            cfg = Path(d) / "mapping.json"
            cfg.write_text(json.dumps({
                "defaults": {"common_dir": "src/common"},
                "projects": [{"name": "示例表A", "search_dir": str(d),
                              "prefix": "示例表A", "require_version": False,
                              "src": "src/projects/示例表A"}],
            }, ensure_ascii=False), encoding="utf-8")
            env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(cfg),
                       PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
            r = subprocess.run(
                [sys.executable, str(AUDIT), "-i"],
                cwd=str(PKG.parent), env=env, input="",
                capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(r.returncode, 1, f"{AUDIT.name}: {r.stderr}")
            self.assertNotIn("Traceback", r.stderr, f"{AUDIT.name} 还有 traceback")
            self.assertIn("没有可交互终端", r.stderr, f"{AUDIT.name} 没给原因")

    def test_sync_interactive_non_tty_is_clean(self):
        """sync -i 在管道里只出人话，不能留 Failure 的 traceback。"""
        with tempfile.TemporaryDirectory() as d:
            cfg = Path(d) / "mapping.json"
            cfg.write_text(json.dumps({
                "defaults": {"common_dir": "src/common"},
                "projects": [{"name": "示例表A", "search_dir": str(d),
                              "prefix": "示例表A", "require_version": False,
                              "src": "src/projects/示例表A"}],
            }, ensure_ascii=False), encoding="utf-8")
            env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(cfg),
                       PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
            r = subprocess.run(
                [sys.executable, str(SYNC), "-i"],
                cwd=str(PKG.parent), env=env, input="",
                capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(r.returncode, 1, f"{SYNC.name}: {r.stderr}")
            self.assertNotIn("Traceback", r.stderr, f"{SYNC.name} 还有 traceback")
            self.assertIn("没有可交互终端", r.stderr, f"{SYNC.name} 没给原因")


if __name__ == "__main__":
    unittest.main()
