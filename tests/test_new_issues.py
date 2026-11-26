# -*- coding: utf-8 -*-
"""新问题回归：dry 预检、watch 推荐命令、宏不可用提示统一。"""

import io
import contextlib
import json
import sys
import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))

from features import (audit_a, audit_b, agent_io, agent_patch,  # noqa: E402
                      sync_push, sync_read)
from features.audit_engine import AuditEngine  # noqa: E402
from toolkit.base import PKG_ROOT, entry_command  # noqa: E402
from toolkit.contract import Failure  # noqa: E402
from toolkit.context import Progress  # noqa: E402
from toolkit.render import render  # noqa: E402
from toolkit.state import Purpose, TargetState, rule_for  # noqa: E402


def _ctx(args, config: dict, root: Path) -> SimpleNamespace:
    """最小运行上下文：形状对齐 ``toolkit.context.AppContext``，测试可直接塞桩。"""
    return SimpleNamespace(
        args=args, config=config, root=Path(root), json_mode=False,
        started=time.perf_counter(), progress=Progress(tool="test", enabled=False),
        elapsed_ms=lambda: 1, stage_timings=lambda: {})


class DryPreflightTest(unittest.TestCase):
    def test_push_dry_runs_source_preflight(self):
        with tempfile.TemporaryDirectory() as d:
            src = Path(d) / "src"
            src.mkdir()
            proj = {"name": "P", "src": str(src)}
            cfg = {"defaults": {}, "projects": [proj]}
            args = SimpleNamespace(project="P", apply=False, doc_modules=False,
                                    no_backup=False, no_save=False, open_only=False,
                                    rollback_on_fail=False, all_project=False)
            ctx = _ctx(args, cfg, Path(d))
            with patch.object(sync_push, "_find_project", return_value=proj), \
                    patch.object(sync_push, "resolve_target",
                                 return_value=(Path(d) / "P.xlsm", "")), \
                    patch.object(sync_push, "_preflight_or_fail",
                                 side_effect=Failure("预检失败", code="SOURCE_PREFLIGHT",
                                                      stage="preflight")):
                outcome = sync_push.push(ctx)
        # 预检失败被收进 Outcome.errors（架构统一 Outcome 契约）；dry 没直接退 0 跳过预检
        self.assertEqual(outcome.exit_code, 1)
        self.assertEqual(outcome.errors[0].code, "SOURCE_PREFLIGHT")
        self.assertFalse(outcome.stats["results"][0]["ok"])


class OverviewSeverityTest(unittest.TestCase):
    def test_manual_accel_is_warning_not_fatal(self):
        issues = {"manual_accel": ["a", "b", "c"], "sheet_mismatch": [], "dup": []}
        self.assertEqual(AuditEngine.issues_fatal(issues), 0)

    def test_sheet_mismatch_is_fatal(self):
        self.assertEqual(AuditEngine.issues_fatal({"sheet_mismatch": ["x"]}), 1)


class OverviewCommandTest(unittest.TestCase):
    def test_manual_accel_only_overview_is_ok_and_exit_zero(self):
        keys = ["stale", "orphan", "kind", "missing", "dup", "shadow", "shadow_known",
                "doc_orphan", "doc_ungrabbed", "sheet_mismatch", "suffix", "naked_new",
                "manual_accel", "no_comment", "binding_target"]
        issues = {k: [] for k in keys}
        issues["manual_accel"] = [f"m{i}" for i in range(6)]
        cfg = {"defaults": {}, "projects": []}
        with tempfile.TemporaryDirectory() as d:
            args = SimpleNamespace(project=[], exclude=[], ignore=[], strict=False)
            ctx = _ctx(args, cfg, Path(d))
            with patch.object(AuditEngine, "common_analyses", return_value=[]), \
                    patch.object(AuditEngine, "structure_issues", return_value=issues), \
                    contextlib.redirect_stdout(io.StringIO()):
                outcome = audit_b.overview(ctx)
        self.assertEqual(outcome.exit_code, 0)
        self.assertEqual(outcome.errors, [])
        self.assertNotIn("errors", outcome.stats)
        self.assertEqual(outcome.stats["issues"], 6)
        self.assertEqual(outcome.stats["fatal"], 0)


class H10SuggestionTest(unittest.TestCase):
    def test_warning_row_carries_clean_backup_command(self):
        with tempfile.TemporaryDirectory() as d:
            bk = Path(d) / ".backup"
            bk.mkdir()
            for i in range(61):
                (bk / f"b{i:02d}").mkdir()
            cfg = {"projects": [], "defaults": {"backup_dir": str(bk), "backup_keep": 10}}
            rows = AuditEngine(Path(d), cfg).indicator_rows(ignore="H05a,H06")
        h10 = [r for r in rows if r["id"] == "H10"][0]
        self.assertEqual(h10["state"], "警告")
        self.assertIn("clean-backup --apply", h10["suggested_command"])

    def test_min_keep_does_not_trigger_per_tag_warning(self):
        with tempfile.TemporaryDirectory() as d:
            bk = Path(d) / ".backup"
            bk.mkdir()
            for i in range(12):
                (bk / f"b{i:02d}").mkdir()
            cfg = {"projects": [], "defaults": {
                "backup_dir": str(bk), "backup_keep": 10,
                "backup_keep_daily": 0, "backup_min_keep": 3}}
            rows = AuditEngine(Path(d), cfg).indicator_rows(ignore="H05a,H06")
        h10 = [r for r in rows if r["id"] == "H10"][0]
        self.assertEqual(h10["state"], "通过")
        self.assertIn("每类 ≤13", h10["detail"])


class WatchPlanCommandTest(unittest.TestCase):
    def test_watch_without_open_gives_open_command(self):
        with tempfile.TemporaryDirectory() as d:
            proj = {"name": "P", "src": str(Path(d))}
            cfg = {"defaults": {}, "projects": [proj]}
            args = SimpleNamespace(project="P", intent="watch")
            ctx = _ctx(args, cfg, Path(d))
            with patch.object(sync_read, "_find_project", return_value=proj), \
                    patch.object(sync_read, "resolve_target",
                                 return_value=(Path(d) / "P.xlsm", "")), \
                    patch.object(sync_read.excel_gateway, "attach_running_excel",
                                 side_effect=Failure("没开", code="NO_EXCEL", stage="attach")):
                outcome = sync_read.plan(ctx)
        self.assertEqual(outcome.exit_code, 0)
        self.assertEqual(outcome.stats["command"],
                         "agent.py book --project P --action open --apply")
        self.assertTrue(outcome.stats["needs_open"])


class TargetClosedHintTest(unittest.TestCase):
    def test_hint_is_per_verb(self):
        # 1.6 起关闭态提示按用途登记在 toolkit.state 矩阵；pull 的磁盘替代提示留在
        # sync_read._closed_hint（旧单体 _target_closed_hint 已拆到这两处）。
        run = rule_for(TargetState.CLOSED, Purpose.RUN)
        pull = sync_read._closed_hint("P", "抓取")
        watch = rule_for(TargetState.CLOSED, Purpose.WATCH)
        push = rule_for(TargetState.CLOSED, Purpose.PUSH)
        self.assertIn(sys.executable, run.suggested)
        self.assertIn(str(PKG / "scripts" / "agent.py"), run.suggested)
        self.assertNotIn("--open-only", run.hint + run.suggested)
        self.assertIn("sync.py pull", pull[0])
        self.assertIn("sync.py pull", pull[1])
        self.assertIn("agent.py book", watch.suggested)
        # PUSH 关闭态在矩阵里未登记统一规则（旧 --open-only 提示只留在 CLI 帮助与适配器兜底）
        self.assertIsNone(push)

    def test_book_close_quits_empty_instance(self):
        class FakeBooks:
            Count = 0
        class FakeXl:
            Workbooks = FakeBooks()
            quit = False
            def Quit(self):
                self.quit = True
        xl = FakeXl()
        self.assertTrue(agent_io._quit_if_empty(xl))
        self.assertTrue(xl.quit)


class AuditCommonHintTest(unittest.TestCase):
    def test_common_not_configured_is_warning_exit3(self):
        out = io.StringIO()
        with tempfile.TemporaryDirectory() as d:
            cfg = {"defaults": {"common_dir": d}, "projects": []}
            args = SimpleNamespace(project=None, exclude=[], all=False)
            ctx = _ctx(args, cfg, Path(d))
            with patch.object(audit_a, "_common_analyses", return_value=([], [])), \
                    contextlib.redirect_stdout(out):
                outcome = audit_a.handle_common(ctx)
                code = render(outcome, json_mode=True)
        payload = json.loads(out.getvalue().strip().splitlines()[-1])
        self.assertEqual(code, 3)
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["errors"], [])
        self.assertFalse(payload["stats"]["configured"])
        self.assertEqual(payload["stats"]["config_key"], "defaults.common_dir")
        self.assertEqual(payload["stats"]["common_dir"], d)


class AssertSettingsTest(unittest.TestCase):
    def test_cli_overrides_mapping_and_sheet_default(self):
        cfg = {"defaults": {"assert_macro": "默认.全部", "assert_sheet": "默认表"},
               "projects": []}
        args = SimpleNamespace(macro="指定.这个", sheet="")
        ctx = _ctx(args, cfg, Path(tempfile.gettempdir()))
        self.assertEqual(agent_patch._assert_settings(ctx, {"name": "P"}, args),
                         ("指定.这个", "默认表"))

    def test_template_default_and_invalid_macro(self):
        cfg = {"defaults": {}, "projects": []}
        ctx = _ctx(SimpleNamespace(macro="", sheet=""), cfg,
                   Path(tempfile.gettempdir()))
        self.assertEqual(
            agent_patch._assert_settings(ctx, {"name": "P"},
                                         SimpleNamespace(macro="", sheet="")),
            (agent_patch.DEFAULT_ASSERT_MACRO, agent_patch.TESTSHEET))
        with self.assertRaises(Failure) as cm:
            agent_patch._assert_settings(
                ctx, {"name": "P"}, SimpleNamespace(macro="没有点", sheet=""))
        self.assertEqual(cm.exception.code, "ASSERT_MACRO_INVALID")


class EntryCommandContractTest(unittest.TestCase):
    """入口命令契约：entry_command 拼解释器绝对路径 + 本包入口绝对路径。"""

    def test_list_command_carries_absolute_interpreter_and_script(self):
        cmd = entry_command("sync.py", "list", "--json")
        self.assertIn(str(PKG_ROOT / "scripts" / "sync.py"), cmd)
        self.assertIn(Path(sys.executable).name, cmd)

    def test_slice_sources_hardcode_no_package_prefixed_command(self):
        # 静态断言：本切片 9 个源文件不含 `python vba-workbench` 前缀命令（斜杠/反斜杠写法）
        sources = [
            "scripts/features/audit_b.py",
            "scripts/features/audit_engine.py",
            "scripts/toolkit/excel.py",
            "scripts/toolkit/state.py",
            "scripts/toolkit/style_rules.py",
            "scripts/toolkit/vbabin.py",
            "scripts/toolkit/cli.py",
            "scripts/toolkit/config.py",
            "scripts/make_lab_workbook.py",
        ]
        for rel in sources:
            text = (PKG / rel).read_text(encoding="utf-8")
            self.assertNotIn("python vba-workbench/", text, rel)
            self.assertNotIn("python vba-workbench\\", text, rel)


if __name__ == "__main__":
    unittest.main()
