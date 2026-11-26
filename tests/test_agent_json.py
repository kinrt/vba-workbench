# -*- coding: utf-8 -*-
"""agent.py 的 --json 是全局开关：任何子命令都只吐一行 JSON。"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
AGENT = PKG / "scripts" / "agent.py"


class AgentJsonTest(unittest.TestCase):
    def test_probe_single_line(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            cfg = d / "mapping.json"
            cfg.write_text(json.dumps({
                "projects": [{"name": "示例表A", "search_dir": str(d), "prefix": "示例表A",
                              "require_version": True, "src": "src/projects/示例表A"}],
            }, ensure_ascii=False), encoding="utf-8")
            env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(cfg), PYTHONIOENCODING="utf-8")
            r = subprocess.run([sys.executable, str(AGENT), "probe", "--json"],
                               cwd=str(d), env=env, capture_output=True, text=True, encoding="utf-8")
            out = r.stdout.strip()
            self.assertEqual(len(out.splitlines()), 1, f"stdout 不止一行：{r.stdout!r}")
            payload = json.loads(out)
            self.assertIn("ok", payload)
            self.assertEqual(payload["action"], "agent.probe")
            self.assertEqual(payload["stats"]["projects"], 0 if not payload["ok"] else 1)
            if payload["ok"]:
                self.assertEqual(payload["details"][0]["project"], "示例表A")
            else:
                self.assertTrue(payload["errors"][0]["message"])

    def test_logs_missing_reports_logs_not_found(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            (d / "示例表A V1.xlsm").write_bytes(b"")
            logs = d / "空日志"
            logs.mkdir()
            cfg = d / "mapping.json"
            cfg.write_text(json.dumps({
                "projects": [{"name": "示例表A", "search_dir": str(d), "prefix": "示例表A",
                              "require_version": True, "src": "src/projects/示例表A"}],
            }, ensure_ascii=False), encoding="utf-8")
            env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(cfg), PYTHONIOENCODING="utf-8")
            r = subprocess.run([sys.executable, str(AGENT), "logs", "--project", "示例表A",
                                "--log", str(logs), "--json"],
                               cwd=str(d), env=env, capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(r.returncode, 1, r.stderr)
            self.assertEqual(len(r.stdout.strip().splitlines()), 1, r.stdout)
            payload = json.loads(r.stdout.strip())
            self.assertEqual(payload["errors"][0]["code"], "LOGS_NOT_FOUND")
            self.assertEqual(payload["errors"][0]["stage"], "logs")
            self.assertIn("--log", payload["errors"][0]["hint"])

    def test_style_failure_json_has_errors(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            cfg = d / "mapping.json"
            cfg.write_text(json.dumps({"projects": []}, ensure_ascii=False), encoding="utf-8")
            src = d / "src" / "模块"
            src.mkdir(parents=True)
            (src / "坏样例.bas").write_text(
                'Attribute VB_Name = "坏样例"\n'
                "Option Explicit\n\n"
                "Public Sub 坏样例()\n"
                '    MsgBox "x"\n'
                "End Sub\n",
                encoding="utf-8", newline="\n")
            env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(cfg),
                       PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
            r = subprocess.run([sys.executable, str(AGENT), "style",
                                "--action", "check", "--json"],
                               cwd=str(d), env=env, capture_output=True,
                               text=True, encoding="utf-8")
            self.assertEqual(r.returncode, 1, r.stderr)
            payload = json.loads(r.stdout.strip())
            self.assertFalse(payload["ok"])
            self.assertTrue(payload["errors"], f"失败时 errors 为空：{payload}")

    def test_metrics_patch_assert_json_acceptance(self):
        """四个遗留动词的 CLI --json 接受度：metrics 正常、patch/assert 失败信封。"""
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            (d / "示例表 V1.xlsm").write_bytes(b"")
            cfg = d / "mapping.json"
            cfg.write_text(json.dumps({
                "projects": [{"name": "示例表", "file": str(d / "示例表 V1.xlsm"),
                              "search_dir": str(d), "prefix": "示例表",
                              "src": "src/projects/示例表"}],
            }, ensure_ascii=False), encoding="utf-8")
            (d / "src" / "模块").mkdir(parents=True)
            (d / "src" / "模块" / "示例.bas").write_text(
                'Attribute VB_Name = "示例"\nOption Explicit\n\n'
                "Public Sub 示例()\nEnd Sub\n", encoding="utf-8", newline="\n")
            env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(cfg),
                       PYTHONIOENCODING="utf-8", PYTHONUTF8="1")

            def run(*args):
                return subprocess.run([sys.executable, str(AGENT), *args, "--json"],
                                      cwd=str(d), env=env, capture_output=True,
                                      text=True, encoding="utf-8")

            cases = [
                ("metrics", (), 0),
                ("patch", ("--file", str(d / "不存在.patch")), 1),
                ("assert", ("--project", "不存在", "--apply"), 1),
            ]
            for verb, extra, want_rc in cases:
                with self.subTest(verb=verb):
                    r = run(verb, *extra)
                    self.assertEqual(r.returncode, want_rc, r.stderr)
                    out = r.stdout.strip()
                    self.assertEqual(len(out.splitlines()), 1, r.stdout)
                    payload = json.loads(out)
                    self.assertEqual(payload["action"], f"agent.{verb}")


if __name__ == "__main__":
    unittest.main()
