# -*- coding: utf-8 -*-
"""sync_transform 切片契约测试：align / promote / gbk-export / clean-backup。

四类覆盖：成功 / 业务失败（exit 1，errors[].code）/ 用法错（exit 2）/ --json 单行纯度。
测试只在 tempfile 造假的 src / .backup 目录，不启动 Excel、不回壳别的 CLI。
"""
from __future__ import annotations

import io
import json
import sys
import tempfile
import time
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))

from toolkit import commands, registry  # noqa: E402
from toolkit.contract import Failure, Outcome  # noqa: E402
from toolkit.context import Progress  # noqa: E402
from toolkit.render import render  # noqa: E402

PROJECT = "示例表A"


class FakeCtx:
    """轻量运行上下文：只实现 handler 用到的 config / args / root / 进度计时。"""

    def __init__(self, args, cfg: dict, root: Path) -> None:
        self.args = args
        self.config = cfg
        self.root = Path(root)
        self.json_mode = bool(getattr(args, "json", False))
        self.started = time.perf_counter()
        self.progress = Progress(tool="sync", enabled=False)

    def elapsed_ms(self) -> int:
        return int((time.perf_counter() - self.started) * 1000)

    def stage_timings(self) -> dict[str, int]:
        return dict(self.progress.timings)


def build_parser():
    return registry.build_parser(commands.specs_for("sync"), prog="sync.py")


def make_sandbox() -> tuple[tempfile.TemporaryDirectory, Path, dict]:
    d = tempfile.TemporaryDirectory()
    root = Path(d.name)
    common = root / "src" / "common" / "模块"
    proj = root / "src" / "projects" / PROJECT / "模块"
    common.mkdir(parents=True)
    proj.mkdir(parents=True)
    cfg = {
        "defaults": {
            "common_dir": str(root / "src" / "common"),
            "backup_dir": str(root / ".backup"),
            "projects_dir": str(root / "src" / "projects"),
        },
        "projects": [{
            "name": PROJECT,
            "src": str(root / "src" / "projects" / PROJECT),
            "search_dir": str(root),
            "prefix": PROJECT,
            "require_version": False,
        }],
    }
    return d, root, cfg


def write_file(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def run_handler(argv: list[str], cfg: dict, root: Path):
    """解析参数 -> 跑 handler；Failure 包成 exit 1 的 Outcome（与入口壳同口径）。"""
    parser = build_parser()
    args = parser.parse_args(argv)
    ctx = FakeCtx(args, cfg, root)
    try:
        outcome = registry.load_handler(args._spec)(ctx)
    except Failure as e:
        outcome = Outcome(action=f"sync.{args._spec.verb}", ok=False,
                          errors=[e], exit_code=1, elapsed_ms=0)
    return args, outcome


def render_json(outcome: Outcome) -> tuple[int, dict, str]:
    buf, err = io.StringIO(), io.StringIO()
    with redirect_stdout(buf), redirect_stderr(err):
        code = render(outcome, json_mode=True)
    lines = buf.getvalue().splitlines()
    payload = json.loads(lines[0]) if lines else {}
    return code, payload, buf.getvalue()


class AlignTest(unittest.TestCase):
    def setUp(self) -> None:
        self._d, self.root, self.cfg = make_sandbox()
        common = self.root / "src" / "common" / "模块"
        proj = self.root / "src" / "projects" / PROJECT / "模块"
        write_file(common / "公共.bas",
                   'Attribute VB_Name = "公共"\nOption Explicit\n'
                   'Public Function AddOne(ByVal x As Long) As Long\nEnd Function\n')
        write_file(proj / "公共.bas",
                   'Attribute VB_Name = "公共"\nOption Explicit\n'
                   'Public Sub OldCopy()\nEnd Sub\n')

    def tearDown(self) -> None:
        self._d.cleanup()

    def test_success_preview_lists_same_name(self):
        _args, outcome = run_handler(["align", "--project", PROJECT, "--json"],
                                     self.cfg, self.root)
        self.assertTrue(outcome.ok, outcome.errors)
        self.assertEqual(outcome.exit_code, 0)
        self.assertEqual(outcome.details[0]["name"], PROJECT)
        self.assertEqual(outcome.details[0]["swapped"], 1)
        self.assertEqual(outcome.details[0]["kept"], 0)
        self.assertEqual(outcome.details[0]["already"], 0)
        self.assertTrue(outcome.stats["dry"])

    def test_success_apply_swaps_and_backs_up(self):
        from toolkit.base import set_assume_yes
        set_assume_yes(True)
        self.addCleanup(set_assume_yes, False)
        proj_file = self.root / "src" / "projects" / PROJECT / "模块" / "公共.bas"
        _args, outcome = run_handler(
            ["align", "--project", PROJECT, "--apply"], self.cfg, self.root)
        self.assertTrue(outcome.ok, outcome.errors)
        self.assertFalse(outcome.stats["dry"])
        self.assertEqual(outcome.details[0]["swapped"], 1)
        text = proj_file.read_text(encoding="utf-8")
        self.assertIn("真实文件在", text)
        self.assertNotIn("OldCopy", text)
        backups = list(self.root.rglob("公共.bas"))
        self.assertTrue(any(".backup" in str(p) for p in backups))

    def test_business_failure_unknown_project(self):
        _args, outcome = run_handler(["align", "--project", "不存在项目", "--json"],
                                     self.cfg, self.root)
        self.assertEqual(outcome.exit_code, 1)
        self.assertFalse(outcome.ok)
        self.assertEqual(outcome.errors[0].code, "PROJECT_NOT_FOUND")
        self.assertEqual(outcome.errors[0].stage, "resolve")

    def test_usage_error_bad_flag_exits_2(self):
        parser = build_parser()
        with self.assertRaises(SystemExit) as cm:
            parser.parse_args(["align", "--bogus"])
        self.assertEqual(cm.exception.code, 2)

    def test_json_purity_single_envelope(self):
        _args, outcome = run_handler(["align", "--project", PROJECT, "--json"],
                                     self.cfg, self.root)
        code, payload, raw = render_json(outcome)
        self.assertEqual(code, 0)
        self.assertEqual(len(raw.strip().splitlines()), 1)
        for key in ("ok", "action", "details", "errors", "elapsed_ms", "stats", "exit_code"):
            self.assertIn(key, payload)
        self.assertEqual(payload["action"], "sync.align")


class PromoteTest(unittest.TestCase):
    def setUp(self) -> None:
        self._d, self.root, self.cfg = make_sandbox()
        proj = self.root / "src" / "projects" / PROJECT / "模块"
        write_file(proj / "工具.bas",
                   'Attribute VB_Name = "工具"\nOption Explicit\n'
                   'Public Sub DoIt()\nEnd Sub\n')

    def tearDown(self) -> None:
        self._d.cleanup()

    def test_success_preview(self):
        _args, outcome = run_handler(
            ["promote", "--project", PROJECT, "--module", "工具", "--json"],
            self.cfg, self.root)
        self.assertTrue(outcome.ok, outcome.errors)
        self.assertEqual(outcome.exit_code, 0)
        self.assertTrue(outcome.stats["dry"])
        self.assertEqual(outcome.stats["module"], "工具")
        self.assertEqual(outcome.stats["project"], PROJECT)
        self.assertIn("from_path", outcome.stats)
        self.assertIn("to_path", outcome.stats)
        self.assertTrue(outcome.details[0]["dry"])
        # 预览不动文件
        self.assertTrue((self.root / "src" / "projects" / PROJECT / "模块" / "工具.bas").exists())

    def test_success_apply_moves_and_writes_shell(self):
        _args, outcome = run_handler(
            ["promote", "--project", PROJECT, "--module", "工具", "--apply"],
            self.cfg, self.root)
        self.assertTrue(outcome.ok, outcome.errors)
        self.assertEqual(outcome.exit_code, 0)
        common_file = self.root / "src" / "common" / "模块" / "工具.bas"
        self.assertTrue(common_file.exists())
        # 原位置留下同名占位壳
        shell = self.root / "src" / "projects" / PROJECT / "模块" / "工具.bas"
        self.assertTrue(shell.exists())
        self.assertTrue(outcome.stats["shell_written"])
        self.assertFalse(outcome.stats["dry"])

    def test_business_failure_module_not_found(self):
        _args, outcome = run_handler(
            ["promote", "--project", PROJECT, "--module", "不存在的模块", "--json"],
            self.cfg, self.root)
        self.assertEqual(outcome.exit_code, 1)
        self.assertEqual(outcome.errors[0].code, "TARGET_NOT_FOUND")
        self.assertEqual(outcome.errors[0].stage, "resolve")

    def test_business_failure_target_exists(self):
        common = self.root / "src" / "common" / "模块"
        write_file(common / "工具.bas",
                   'Attribute VB_Name = "工具"\nOption Explicit\n')
        _args, outcome = run_handler(
            ["promote", "--project", PROJECT, "--module", "工具", "--apply"],
            self.cfg, self.root)
        self.assertEqual(outcome.exit_code, 1)
        self.assertEqual(outcome.errors[0].code, "SOURCE_PREFLIGHT")

    def test_usage_error_missing_args_exits_2(self):
        parser = build_parser()
        with self.assertRaises(SystemExit) as cm:
            parser.parse_args(["promote"])
        self.assertEqual(cm.exception.code, 2)
        with self.assertRaises(SystemExit) as cm2:
            parser.parse_args(["promote", "--project", PROJECT])
        self.assertEqual(cm2.exception.code, 2)

    def test_json_purity_single_envelope(self):
        _args, outcome = run_handler(
            ["promote", "--project", PROJECT, "--module", "工具", "--json"],
            self.cfg, self.root)
        code, payload, raw = render_json(outcome)
        self.assertEqual(code, 0)
        self.assertEqual(len(raw.strip().splitlines()), 1)
        for key in ("ok", "action", "details", "errors", "elapsed_ms", "stats", "exit_code"):
            self.assertIn(key, payload)
        self.assertEqual(payload["action"], "sync.promote")


class GbkExportTest(unittest.TestCase):
    def setUp(self) -> None:
        self._d, self.root, self.cfg = make_sandbox()
        proj = self.root / "src" / "projects" / PROJECT / "模块"
        write_file(proj / "中文模块.bas",
                   'Attribute VB_Name = "中文模块"\r\nOption Explicit\n'
                   'Public Sub 演示()\nEnd Sub\n')

    def tearDown(self) -> None:
        self._d.cleanup()

    def test_success_preview(self):
        _args, outcome = run_handler(
            ["gbk-export", "--project", PROJECT, "--json"], self.cfg, self.root)
        self.assertTrue(outcome.ok, outcome.errors)
        self.assertEqual(outcome.exit_code, 0)
        self.assertTrue(outcome.stats["dry"])
        self.assertEqual(outcome.stats["exported"], 1)
        self.assertEqual(outcome.stats["shells_skipped"], 0)
        row = outcome.details[0]
        self.assertEqual(row["name"], PROJECT)
        self.assertIn("out_dir", row)
        self.assertIn("cleaned", row)

    def test_success_apply_writes_gbk_crlf(self):
        _args, outcome = run_handler(
            ["gbk-export", "--project", PROJECT, "--apply"], self.cfg, self.root)
        self.assertTrue(outcome.ok, outcome.errors)
        self.assertFalse(outcome.stats["dry"])
        dest = list((self.root / ".backup" / "GBK" / PROJECT).rglob("*.bas"))
        self.assertEqual(len(dest), 1)
        raw = dest[0].read_bytes()
        # GBK 可解码、换行是 CRLF
        text = raw.decode("gbk")
        self.assertIn("\r\n", text)
        self.assertIn("演示", text)

    def test_business_failure_unknown_project(self):
        _args, outcome = run_handler(
            ["gbk-export", "--project", "不存在项目", "--json"], self.cfg, self.root)
        self.assertEqual(outcome.exit_code, 1)
        self.assertEqual(outcome.errors[0].code, "PROJECT_NOT_FOUND")

    def test_usage_error_no_project_exits_2(self):
        parser = build_parser()
        with self.assertRaises(SystemExit) as cm:
            parser.parse_args(["gbk-export"])
        self.assertEqual(cm.exception.code, 2)

    def test_json_purity_single_envelope(self):
        _args, outcome = run_handler(
            ["gbk-export", "--project", PROJECT, "--json"], self.cfg, self.root)
        code, payload, raw = render_json(outcome)
        self.assertEqual(code, 0)
        self.assertEqual(len(raw.strip().splitlines()), 1)
        for key in ("ok", "action", "details", "errors", "elapsed_ms", "stats", "exit_code"):
            self.assertIn(key, payload)
        self.assertEqual(payload["action"], "sync.gbk-export")


class CleanBackupTest(unittest.TestCase):
    def setUp(self) -> None:
        self._d, self.root, self.cfg = make_sandbox()
        backup = self.root / ".backup"
        for stamp, tag in (("20261120-100000", "-抓取前"), ("20261121-100000", ""),
                           ("20261122-100000", "")):
            b = backup / f"{stamp}{tag}"
            b.mkdir(parents=True)
            write_file(b / "sample.bas", "Attribute VB_Name = \"Sample\"\n")

    def tearDown(self) -> None:
        self._d.cleanup()

    def test_success_preview_defaults(self):
        _args, outcome = run_handler(["clean-backup", "--json"], self.cfg, self.root)
        self.assertTrue(outcome.ok, outcome.errors)
        self.assertEqual(outcome.exit_code, 0)
        self.assertFalse(outcome.stats["applied"])
        for key in ("applied", "freed_bytes", "deleted_batches", "kept", "kept_bytes",
                    "keep_recent", "keep_requested", "keep_note", "keep_daily",
                    "min_keep", "auto_clean", "keep_scope", "by_tag", "kept_by_tag",
                    "kept_names", "leftover", "protected", "timings"):
            self.assertIn(key, outcome.stats)
        self.assertIn("clean-backup", outcome.stats["timings"])
        self.assertEqual(outcome.stats["keep_scope"], "per-tag")
        self.assertIsNone(outcome.stats["keep_requested"])
        # 预览不删盘
        self.assertEqual(len(list((self.root / ".backup").iterdir())), 3)

    def test_keep_below_one_clamped_to_one(self):
        _args, outcome = run_handler(["clean-backup", "--keep", "0", "--json"],
                                     self.cfg, self.root)
        self.assertEqual(outcome.exit_code, 0)
        self.assertEqual(outcome.stats["keep_recent"], 1)
        self.assertIn("无效", outcome.stats["keep_note"])

    def test_apply_deletes_only_over_keep(self):
        cfg = json.loads(json.dumps(self.cfg))
        cfg["defaults"]["backup_min_keep"] = 0
        cfg["defaults"]["backup_keep_daily"] = 0
        _args, outcome = run_handler(
            ["clean-backup", "--keep", "1", "--apply"], cfg, self.root)
        self.assertTrue(outcome.ok, outcome.errors)
        self.assertTrue(outcome.stats["applied"])
        self.assertEqual(outcome.stats["deleted_batches"], 1)
        # 抓取前 1 批与默认来源最新 1 批留下，旧的默认批次被删。
        self.assertEqual(len(list((self.root / ".backup").iterdir())), 2)

    def test_keep_tag_protects_batch(self):
        _args, outcome = run_handler(
            ["clean-backup", "--keep-tag", "抓取前", "--json"], self.cfg, self.root)
        self.assertEqual(outcome.exit_code, 0)
        self.assertTrue(any("抓取前" in name for name in outcome.stats["protected"]))

    def test_business_failure_bad_config(self):
        bad = json.loads(json.dumps(self.cfg))
        bad["defaults"]["backup_keep"] = "坏"
        _args, outcome = run_handler(["clean-backup", "--json"], bad, self.root)
        self.assertEqual(outcome.exit_code, 1)
        self.assertEqual(outcome.errors[0].code, "CONFIG_INVALID")
        self.assertEqual(outcome.errors[0].stage, "config")

    def test_usage_error_non_integer_keep_exits_2(self):
        parser = build_parser()
        with self.assertRaises(SystemExit) as cm:
            parser.parse_args(["clean-backup", "--keep", "abc"])
        self.assertEqual(cm.exception.code, 2)

    def test_json_purity_single_envelope(self):
        _args, outcome = run_handler(["clean-backup", "--json"], self.cfg, self.root)
        code, payload, raw = render_json(outcome)
        self.assertEqual(code, 0)
        self.assertEqual(len(raw.strip().splitlines()), 1)
        for key in ("ok", "action", "details", "errors", "elapsed_ms", "stats", "exit_code"):
            self.assertIn(key, payload)
        self.assertEqual(payload["action"], "sync.clean-backup")


if __name__ == "__main__":
    unittest.main()
