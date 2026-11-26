# -*- coding: utf-8 -*-
"""结构体检门槛：登记类提醒不算失败，真问题才算。"""

import sys
import tempfile
import unittest
from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))

from features import audit_b  # noqa: E402
from features.audit_engine import AuditEngine  # noqa: E402
from toolkit.vba_parse import scan_rules  # noqa: E402


class _StubProgress:
    @contextmanager
    def stage(self, name):
        yield


def _ctx(root: Path, cfg: dict, **args) -> SimpleNamespace:
    """最小运行上下文：形状对齐 ``toolkit.context.AppContext``，测试可直接塞桩。"""
    ns = SimpleNamespace(exclude=[], ignore=[], strict=False, all=False, md=None,
                         out=None, project=[], json=False, no_input=False)
    for key, value in args.items():
        setattr(ns, key, value)
    return SimpleNamespace(args=ns, config=cfg, root=Path(root), json_mode=False,
                            elapsed_ms=lambda: 1, stage_timings=lambda: {},
                            progress=_StubProgress())


class NewGuardTest(unittest.TestCase):
    def test_warning_only_exits_three_strict_exits_one(self):
        # audit structure：仅警告退 3；strict 模式把同一批警告升格为失败退 1。
        cfg = {"projects": [], "defaults": {}}
        issues = {k: [] for k in (
            "stale", "orphan", "kind", "missing", "dup", "shadow", "shadow_known",
            "doc_orphan", "doc_ungrabbed", "sheet_mismatch", "naked_new",
            "manual_accel", "no_comment", "binding_target")}
        issues["stale"] = [(Path("x"), Path("y"))]
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            with patch.object(AuditEngine, "structure_issues", return_value=issues):
                self.assertEqual(audit_b.structure(_ctx(root, cfg)).exit_code, 3)
                self.assertEqual(
                    audit_b.structure(_ctx(root, cfg, strict=True)).exit_code, 1)

    def test_leading_comment_is_per_procedure(self):
        """紧邻过程头的注释块才算；上一过程的注释不能算到下一过程。"""
        code = ("Option Explicit\n"
                "' 功能：A\n"
                "Public Sub A()\nEnd Sub\n"
                "Public Sub B()\nEnd Sub\n"
                "\n"
                "' 功能：C\n"
                "Public Sub C()\nEnd Sub\n")
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "注释.bas"
            p.write_text(code, encoding="utf-8", newline="\n")
            rules = scan_rules([p])
        names = [x[2] for x in rules["no_comment"]]
        self.assertEqual(names, ["Public Sub B()"])

    def test_paired_accel_restores_are_not_violations(self):
        """关闭与错误分支里的重复还原成对出现时，还原行本身不该报 H05a。"""
        code = ("Option Explicit\n"
                "Public Function 安全执行() As Boolean\n"
                "    On Error GoTo 清理\n"
                "    Application.ScreenUpdating = False\n"
                "    Application.Calculation = xlCalculationManual\n"
                "    Application.EnableEvents = False\n"
                "    Application.ScreenUpdating = True\n"
                "    Application.Calculation = xlCalculationAutomatic\n"
                "    Application.EnableEvents = True\n"
                "    Exit Function\n"
                "清理:\n"
                "    Application.ScreenUpdating = True\n"
                "    Application.Calculation = xlCalculationAutomatic\n"
                "    Application.EnableEvents = True\n"
                "End Function\n")
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "通用工具.bas"
            p.write_text(code, encoding="utf-8", newline="\n")
            rules = scan_rules([p])
        self.assertEqual(rules["manual_accel"], [])

    def test_unpaired_accel_off_is_still_flagged(self):
        code = ("Option Explicit\n"
                "Public Sub 快()\n"
                "    Application.ScreenUpdating = False\n"
                "End Sub\n")
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "快.bas"
            p.write_text(code, encoding="utf-8", newline="\n")
            rules = scan_rules([p])
        self.assertEqual(len(rules["manual_accel"]), 1)

    def test_binding_shortcut_target_is_checked(self):
        """快捷键目标必须是 Public 零必填参数过程；动作过程参数不够也要报。"""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            proj, common = root / "proj", root / "common"
            proj.mkdir(); common.mkdir()
            (proj / "界面清单.bas").write_text(
                "Option Explicit\n"
                "Public Sub 好()\nEnd Sub\n"
                "Public Sub 有参(x As Long)\nEnd Sub\n"
                "Private Sub 私()\nEnd Sub\n"
                "Public Sub 注册()\n"
                "    Call 注册快捷键(\"^a\", \"好\")\n"
                "    Call 注册快捷键(\"^b\", \"有参\")\n"
                "    Call 注册快捷键(\"^c\", \"私\")\n"
                "    Call 注册快捷键(\"^d\", \"不存在\")\n"
                "    Call 注册动作(\"动作1\", \"有参\")\n"
                "    ' Call 注册快捷键(\"^e\", \"注释里的不存在\")\n"
                "End Sub\n",
                encoding="utf-8", newline="\n")
            cfg = {"defaults": {"common_dir": str(common)},
                   "projects": [{"name": "P", "src": str(proj)}]}
            issues = AuditEngine(root, cfg)._binding_issues(exclude=None)
        text = "\n".join(x[3] for x in issues)
        self.assertIn("有参", text)
        self.assertIn("私", text)
        self.assertIn("不存在", text)
        self.assertNotIn("注释里的不存在", text)
        self.assertGreaterEqual(len(issues), 4)

    def test_protected_backup_is_not_automatically_deleted(self):
        from toolkit.files import prune_backups
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            for name in ("20260101-000000", "20260101-000001-覆盖前", "20260102-000000"):
                (root / name).mkdir()
            r = prune_backups(root, keep_recent=1, keep_daily=1, min_keep=1,
                              dry_run=True, protect=["-覆盖前"])
        self.assertIn("20260101-000001-覆盖前", r["protected"])
        self.assertNotIn("20260101-000001-覆盖前", [n for n, _s in r["deleted"]])


    def test_registered_notes_do_not_fail(self):
        issues = {"shadow_known": [1, 2], "stale": [], "orphan": []}
        self.assertEqual(AuditEngine.issues_total(issues), 0)
        self.assertEqual(AuditEngine.issues_total({"shadow_known": [1], "stale": [1]}), 1)

    def test_overview_without_common_still_reports(self):
        """公共库为空时 overview 只报项目/结构，不能整条命令失败。"""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            cfg = {"defaults": {"common_dir": str(root)}, "projects": []}
            with patch.object(AuditEngine, "common_analyses", return_value=[]):
                outcome = audit_b.overview(_ctx(root, cfg))
        self.assertEqual(outcome.exit_code, 0)
        self.assertEqual(outcome.stats["modules"], 0)
        self.assertFalse(outcome.stats["common_configured"])

    def test_paired_accel_with_error_handler_is_exempt(self):
        """On Error 后成对还原的三项加速，不该再报 manual_accel。"""
        code = ("Option Explicit\n"
                "Public Sub 安全执行()\n"
                "    On Error GoTo 收尾\n"
                "    Application.ScreenUpdating = False\n"
                "    Application.Calculation = xlCalculationManual\n"
                "    Application.EnableEvents = False\n"
                "    Debug.Print 1\n"
                "    Exit Sub\n"
                "收尾:\n"
                "    Application.ScreenUpdating = True\n"
                "    Application.Calculation = xlCalculationAutomatic\n"
                "    Application.EnableEvents = True\n"
                "End Sub\n")
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "安全执行.bas"
            p.write_text(code, encoding="utf-8", newline="\n")
            rules = scan_rules([p])
        self.assertEqual(rules["manual_accel"], [])

    def test_unpaired_accel_still_reported(self):
        code = ("Option Explicit\n"
                "Public Sub 裸关加速()\n"
                "    Application.ScreenUpdating = False\n"
                "End Sub\n")
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "裸关加速.bas"
            p.write_text(code, encoding="utf-8", newline="\n")
            rules = scan_rules([p])
        self.assertEqual(len(rules["manual_accel"]), 1)

    def test_statusbar_clear_and_progress_class_are_exempt(self):
        """清空式写入、自带 结束 的进度器类、自测模块不报 H06；普通消息仍报。"""
        files = {}
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            (d / "ThisWorkbook.bas").write_text(
                "Option Explicit\nPrivate Sub Workbook_Open()\n"
                "    Application.StatusBar = False\nEnd Sub\n", encoding="utf-8")
            (d / "进度器.cls").write_text(
                "VERSION 1.0 CLASS\nBEGIN\nEND\nOption Explicit\n"
                "Public Sub 写(a As String)\n    Application.StatusBar = a\nEnd Sub\n"
                "Public Sub 结束()\n    Application.StatusBar = \"\"\nEnd Sub\n",
                encoding="utf-8")
            (d / "运行自测.bas").write_text(
                "Option Explicit\nPublic Sub 跑()\n    Application.StatusBar = \"跑\"\nEnd Sub\n",
                encoding="utf-8")
            (d / "表格工具.bas").write_text(
                "Option Explicit\nPublic Sub proPrint(ByVal s As String)\n"
                "    Application.StatusBar = s\nEnd Sub\n", encoding="utf-8")
            (d / "业务.bas").write_text(
                "Option Explicit\nPublic Sub 跑()\n    Application.StatusBar = \"业务\"\nEnd Sub\n",
                encoding="utf-8")
            for p in d.glob("*.*"):
                files[p] = (b"", p.read_text(encoding="utf-8").splitlines())
            bad = AuditEngine(d, {})._statusbar_violations(files)
        self.assertEqual(len(bad), 1)
        self.assertIn("业务.bas", bad[0])


if __name__ == "__main__":
    unittest.main()
