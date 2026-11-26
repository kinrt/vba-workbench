# -*- coding: utf-8 -*-
"""
test_sync_read —— sync_read 切片（vbamod / plan / pull / pull-open）契约测试

覆盖四类：成功 / 业务失败（exit 1 且 errors[].code）/ 用法错（exit 2）/ --json 单行纯度。
COM 一律打桩，不落盘、不启动 Excel；临时文件走 tempfile。
"""
from __future__ import annotations

import io
import json
import sys
import time
import unittest
from argparse import Namespace
from pathlib import Path
from unittest import mock

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))

from toolkit import commands, registry  # noqa: E402
from toolkit.base import Failure  # noqa: E402
from toolkit.contract import ErrorCode  # noqa: E402
from toolkit.context import AppContext, Progress  # noqa: E402
from toolkit.render import render  # noqa: E402
from features import sync_read  # noqa: E402

SPECS = commands.specs_for("sync")
ENVELOPE_KEYS = {"ok", "action", "details", "errors", "elapsed_ms", "stats", "exit_code"}


def make_parser():
    return registry.build_parser(SPECS, prog="sync.py")


def parse(argv: list[str]) -> Namespace:
    return make_parser().parse_args(argv)


def make_ctx(args: Namespace, config: dict, root: Path) -> AppContext:
    return AppContext(
        tool="sync", args=args, config=config, root=root,
        json_mode=bool(getattr(args, "json", False)),
        started=time.perf_counter(), progress=Progress(tool="sync", enabled=False),
    )


def render_json(outcome) -> tuple[int, str, str]:
    """跑 JSON 渲染，返回 (退出码, stdout 文本, stderr 文本)。"""
    out_buf, err_buf = io.StringIO(), io.StringIO()
    old_out, old_err = sys.stdout, sys.stderr
    sys.stdout, sys.stderr = out_buf, err_buf
    try:
        code = render(outcome, json_mode=True)
    finally:
        sys.stdout, sys.stderr = old_out, old_err
    return code, out_buf.getvalue(), err_buf.getvalue()


def write_xlsm(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"PK\x03\x04")
    return path


class _Fixture:
    def __init__(self, tmp: Path, *, file_exists: bool = True):
        self.tmp = tmp
        xlsm = tmp / "示例表A V1.xlsm"
        if file_exists:
            write_xlsm(xlsm)
        self.config = {
            "defaults": {
                "common_dir": str(tmp / "src" / "common"),
                "backup_dir": str(tmp / ".backup"),
            },
            "projects": [{
                "name": "示例表A",
                "file": str(xlsm),
                "src": str(tmp / "src" / "projA"),
                "common_modules": None,
            }],
        }

    def ctx(self, args: Namespace) -> AppContext:
        return make_ctx(args, self.config, self.tmp)


class TestVbamod(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.tmp = Path(tempfile.mkdtemp(prefix="vba-wb-"))
        self.fx = _Fixture(self.tmp)

    def test_vbamod_success(self):
        args = parse(["vbamod", "--project", "示例表A", "--json"])
        backup = self.tmp / ".backup" / "20261117-231633" / "示例表A V1.xlsm"
        backup.parent.mkdir(parents=True, exist_ok=True)
        backup.write_bytes(b"PK\x03\x04")
        with mock.patch.object(sync_read, "vba_module_names", return_value=["模块1"]), \
                mock.patch.object(sync_read, "latest_backup", return_value=backup), \
                mock.patch.object(sync_read, "compare_vba_modules",
                                   return_value={"same": True, "count_before": 1,
                                                 "count_after": 1, "lost": [], "added": []}):
            outcome = sync_read.vbamod(self.fx.ctx(args))
        self.assertEqual(outcome.exit_code, 0)
        self.assertTrue(outcome.ok)
        self.assertEqual(outcome.details, [])
        self.assertEqual(len(outcome.stats["results"]), 1)
        row = outcome.stats["results"][0]
        self.assertTrue(row["compared"])
        self.assertTrue(row["same"])
        self.assertEqual(row["module_names"], ["模块1"])
        self.assertEqual(outcome.stats["examined"], 1)
        self.assertEqual(outcome.stats["changed"], 0)
        self.assertEqual(outcome.stats["unverifiable"], 0)

    def test_vbamod_project_not_found(self):
        args = parse(["vbamod", "--project", "不存在"])
        outcome = sync_read.vbamod(self.fx.ctx(args))
        self.assertEqual(outcome.exit_code, 1)
        self.assertFalse(outcome.ok)
        self.assertEqual(outcome.errors[0].code, "PROJECT_NOT_FOUND")
        self.assertEqual(outcome.errors[0].stage, "resolve")

    def test_vbamod_usage_error(self):
        with self.assertRaises(SystemExit) as cm:
            parse(["vbamod", "--no-such-flag"])
        self.assertEqual(cm.exception.code, 2)

    def test_vbamod_json_purity(self):
        args = parse(["vbamod", "--project", "示例表A", "--json"])
        with mock.patch.object(sync_read, "vba_module_names", return_value=["模块1"]), \
                mock.patch.object(sync_read, "latest_backup", return_value=None):
            outcome = sync_read.vbamod(self.fx.ctx(args))
        code, stdout, _err = render_json(outcome)
        self.assertEqual(code, 0)
        self.assertEqual(len(stdout.strip().splitlines()), 1)
        payload = json.loads(stdout)
        self.assertTrue(ENVELOPE_KEYS <= set(payload))
        self.assertEqual(payload["exit_code"], 0)
        self.assertEqual(payload["action"], "sync.vbamod")
        self.assertEqual(payload["details"], [])
        self.assertIn("results", payload["stats"])


class TestPlan(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.tmp = Path(tempfile.mkdtemp(prefix="vba-wb-"))
        self.fx = _Fixture(self.tmp)

    def test_plan_success_closed(self):
        args = parse(["plan", "--project", "示例表A", "--intent", "pull", "--json"])
        xl = mock.Mock()
        with mock.patch.object(sync_read.excel_gateway, "attach_running_excel", return_value=xl), \
                mock.patch.object(sync_read.excel_gateway, "find_open_workbook",
                                  return_value=(None, False)):
            outcome = sync_read.plan(self.fx.ctx(args))
        self.assertEqual(outcome.exit_code, 0)
        self.assertTrue(outcome.ok)
        self.assertEqual(outcome.details, [])
        row = outcome.stats["results"][0]
        self.assertEqual(row["mode"], "disk")
        self.assertFalse(row["open"])
        self.assertEqual(row["command"], "sync.py pull --project 示例表A --apply")

    def test_plan_pull_open_command_uses_pull(self):
        # 表已打开：pull 自己会转会话，plan 推荐命令统一给 pull，不再要求手敲 pull-open
        args = parse(["plan", "--project", "示例表A", "--intent", "pull", "--json"])
        wb = mock.Mock(Saved=True)
        xl = mock.Mock()
        with mock.patch.object(sync_read.excel_gateway, "attach_running_excel", return_value=xl), \
                mock.patch.object(sync_read.excel_gateway, "find_open_workbook",
                                  return_value=(wb, [])):
            outcome = sync_read.plan(self.fx.ctx(args))
        self.assertEqual(outcome.exit_code, 0)
        self.assertTrue(outcome.ok)
        row = outcome.stats["results"][0]
        self.assertTrue(row["open"])
        self.assertEqual(row["mode"], "open")
        self.assertTrue(row["preserve_unsaved"])
        self.assertEqual(row["command"], "sync.py pull --project 示例表A --apply")

    def test_plan_matrix_closed_open_for_pull_push_run_watch(self):
        """表关/表开 × pull/push/run/watch：mode / needs_open / will_save 固定。"""
        cases = [
            # intent, open, mode, needs_open, will_save, command
            ("pull", False, "disk", False, False,
             "sync.py pull --project 示例表A --apply"),
            ("pull", True, "open", False, False,
             "sync.py pull --project 示例表A --apply"),
            ("push", False, "disk", False, True,
             "sync.py push --project 示例表A --apply"),
            ("push", True, "open", False, True,
             "sync.py push --project 示例表A --apply"),
            ("run", False, "isolate", False, False,
             "sync.py run --project 示例表A --macro <宏名> --isolate --apply"),
            ("run", True, "open", False, False,
             "sync.py run --project 示例表A --macro <宏名> --apply"),
            ("watch", False, "open_required", True, False,
             "agent.py book --project 示例表A --action open --apply"),
            ("watch", True, "open", False, False,
             "sync.py watch --project 示例表A --apply"),
        ]
        for intent, is_open, mode, needs_open, will_save, command in cases:
            with self.subTest(intent=intent, open=is_open):
                args = parse(["plan", "--project", "示例表A",
                              "--intent", intent, "--json"])
                xl = mock.Mock()
                wb = mock.Mock(Saved=True)
                found = (wb, []) if is_open else (None, False)
                with mock.patch.object(sync_read.excel_gateway,
                                       "attach_running_excel", return_value=xl), \
                        mock.patch.object(sync_read.excel_gateway,
                                          "find_open_workbook", return_value=found), \
                        mock.patch.object(sync_read, "vbe_state",
                                          return_value=("ok", "")):
                    outcome = sync_read.plan(self.fx.ctx(args))
                row = outcome.stats["results"][0]
                self.assertEqual(outcome.exit_code, 0)
                self.assertEqual(outcome.stats["mode"], mode)
                self.assertEqual(row["mode"], mode)
                self.assertEqual(row["needs_open"], needs_open)
                self.assertEqual(row["will_save"], will_save)
                self.assertEqual(row["command"], command)
                self.assertEqual(row["open"], is_open)

    def test_plan_target_missing(self):
        import tempfile
        missing_tmp = Path(tempfile.mkdtemp(prefix="vba-wb-missing-"))
        fx = _Fixture(missing_tmp, file_exists=False)
        args = parse(["plan", "--project", "示例表A"])
        outcome = sync_read.plan(fx.ctx(args))
        self.assertEqual(outcome.exit_code, 1)
        self.assertFalse(outcome.ok)
        self.assertEqual(outcome.errors[0].code, "TARGET_NOT_FOUND")
        self.assertEqual(outcome.errors[0].stage, "resolve")
        self.assertFalse(outcome.stats["results"][0]["ok"])

    def test_plan_usage_error(self):
        with self.assertRaises(SystemExit) as cm:
            parse(["plan"])
        self.assertEqual(cm.exception.code, 2)

    def test_plan_probe_failure_fallback_is_registered_code(self):
        # 探针异常没带 code 时，回退码必须是已登记的 PROBE_FAILED
        self.assertEqual(ErrorCode.PROBE_FAILED.value, "PROBE_FAILED")
        args = parse(["plan", "--project", "示例表A", "--intent", "pull", "--json"])
        boom = Failure("附着失败，没带码")
        with mock.patch.object(sync_read.excel_gateway, "attach_running_excel",
                               side_effect=boom):
            outcome = sync_read.plan(self.fx.ctx(args))
        row = outcome.stats["results"][0]
        self.assertEqual(row["probe"], ErrorCode.PROBE_FAILED.value)
        self.assertEqual(row["probe_detail"], "附着失败，没带码")

    def test_plan_json_purity(self):
        args = parse(["plan", "--project", "示例表A", "--intent", "pull", "--json"])
        with mock.patch.object(sync_read.excel_gateway, "attach_running_excel", return_value=mock.Mock()), \
                mock.patch.object(sync_read.excel_gateway, "find_open_workbook",
                                  return_value=(None, False)):
            outcome = sync_read.plan(self.fx.ctx(args))
        code, stdout, _err = render_json(outcome)
        self.assertEqual(code, 0)
        self.assertEqual(len(stdout.strip().splitlines()), 1)
        payload = json.loads(stdout)
        self.assertTrue(ENVELOPE_KEYS <= set(payload))
        self.assertEqual(payload["action"], "sync.plan")
        self.assertEqual(payload["stats"]["intent"], "pull")


class TestPull(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.tmp = Path(tempfile.mkdtemp(prefix="vba-wb-"))
        self.fx = _Fixture(self.tmp)

    def test_pull_success_dry(self):
        args = parse(["pull", "--all-project", "--json"])
        with mock.patch.object(sync_read, "find_excel_with_workbook", return_value=(None, None)):
            outcome = sync_read.pull(self.fx.ctx(args))
        self.assertEqual(outcome.exit_code, 0)
        self.assertTrue(outcome.ok)
        self.assertTrue(outcome.stats["dry"])
        self.assertEqual(outcome.stats["projects"], 1)
        self.assertEqual(outcome.stats["failed"], 0)
        self.assertEqual(outcome.details, [])
        self.assertTrue(outcome.stats["results"][0]["dry"])
        self.assertEqual(outcome.stats["results"][0]["mode"], "disk")
        self.assertNotIn("modules", outcome.stats["results"][0])

    def test_pull_target_missing(self):
        import tempfile
        missing_tmp = Path(tempfile.mkdtemp(prefix="vba-wb-missing-"))
        fx = _Fixture(missing_tmp, file_exists=False)
        args = parse(["pull", "--project", "示例表A"])
        with mock.patch.object(sync_read, "find_excel_with_workbook", return_value=(None, None)):
            outcome = sync_read.pull(fx.ctx(args))
        self.assertEqual(outcome.exit_code, 1)
        self.assertFalse(outcome.ok)
        self.assertEqual(outcome.errors[0].code, "TARGET_NOT_FOUND")
        self.assertEqual(outcome.stats["results"][0]["ok"], False)
        self.assertIn("reason", outcome.stats["results"][0])

    def test_pull_apply_closed_goes_through_disk_gateway(self):
        # 表没开（NO_EXCEL）→ 走隐藏实例磁盘路径；成功路径行为与 1.5.6 一致
        args = parse(["pull", "--project", "示例表A", "--apply", "--json"])
        fake_wb = mock.Mock()
        fake_session = mock.Mock()
        fake_session.app.return_value = mock.Mock()
        no_excel = Failure("没找到正在运行的 Excel", code=ErrorCode.NO_EXCEL, stage="attach")
        with mock.patch.object(sync_read, "ExcelSession", return_value=fake_session), \
                mock.patch.object(sync_read.excel_gateway, "attach_running_excel",
                                  side_effect=no_excel), \
                mock.patch.object(sync_read.excel_gateway, "open_disk_workbook",
                                  return_value=fake_wb), \
                mock.patch.object(sync_read, "has_vbproject_access", return_value=True), \
                mock.patch.object(sync_read, "_export_workbook", return_value=2), \
                mock.patch.object(sync_read, "_write_placeholder_shells"), \
                mock.patch.object(sync_read, "_verify_export_set"):
            outcome = sync_read.pull(self.fx.ctx(args))
        self.assertEqual(outcome.exit_code, 0)
        self.assertTrue(outcome.ok)
        self.assertFalse(outcome.stats["dry"])
        self.assertEqual(outcome.stats["failed"], 0)
        row = outcome.stats["results"][0]
        self.assertTrue(row["ok"])
        self.assertEqual(row["mode"], "disk")
        self.assertEqual(row["modules"], 2)
        fake_session.close.assert_called_once()

    def test_pull_apply_open_uses_session_not_disk(self):
        # 表已打开 → 借已打开会话导出；绝不打开磁盘副本、不读陈旧磁盘
        args = parse(["pull", "--project", "示例表A", "--apply", "--json"])
        fake_wb = mock.Mock(FullName=str(self.tmp / "示例表A V1.xlsm"), Saved=True)
        with mock.patch.object(sync_read, "_attach_pull_session",
                                  return_value=(fake_wb, "open")) as attach, \
                mock.patch.object(sync_read.excel_gateway, "open_disk_workbook") as open_disk, \
                mock.patch.object(sync_read, "_export_workbook", return_value=3), \
                mock.patch.object(sync_read, "_write_placeholder_shells"), \
                mock.patch.object(sync_read, "_verify_export_set"):
            outcome = sync_read.pull(self.fx.ctx(args))
        self.assertEqual(outcome.exit_code, 0)
        self.assertTrue(outcome.ok)
        attach.assert_called_once()
        open_disk.assert_not_called()  # 关键：没走磁盘读取
        row = outcome.stats["results"][0]
        self.assertEqual(row["mode"], "open")
        self.assertEqual(row["modules"], 3)
        fake_wb.Close.assert_not_called()  # 借用户会话，不关窗

    def test_pull_apply_mixed_open_and_disk(self):
        # --all-project 逐项目判定：已打开借会话、没打开走磁盘，不整批一刀切
        import tempfile
        tmp = Path(tempfile.mkdtemp(prefix="vba-wb-mix-"))
        xlsm_a = write_xlsm(tmp / "示例表A V1.xlsm")
        xlsm_b = write_xlsm(tmp / "示例表C V1.xlsm")
        cfg = {
            "defaults": {"common_dir": str(tmp / "src" / "common"),
                         "backup_dir": str(tmp / ".backup")},
            "projects": [
                {"name": "示例表A", "file": str(xlsm_a), "src": str(tmp / "src" / "projA"),
                 "common_modules": None},
                {"name": "示例表C", "file": str(xlsm_b), "src": str(tmp / "src" / "projC"),
                 "common_modules": None},
            ],
        }
        args = parse(["pull", "--all-project", "--apply", "--json"])
        ctx = make_ctx(args, cfg, tmp)
        wb_a = mock.Mock(FullName=str(xlsm_a), Saved=True)
        wb_b = mock.Mock()

        def fake_attach(proj, path, purpose="抓取"):
            return (wb_a, "open") if proj["name"] == "示例表A" else (None, "disk")

        with mock.patch.object(sync_read, "_attach_pull_session", side_effect=fake_attach), \
                mock.patch.object(sync_read.excel_gateway, "open_disk_workbook",
                                  return_value=wb_b) as open_disk, \
                mock.patch.object(sync_read, "has_vbproject_access", return_value=True), \
                mock.patch.object(sync_read, "_export_workbook", return_value=1), \
                mock.patch.object(sync_read, "_write_placeholder_shells"), \
                mock.patch.object(sync_read, "_verify_export_set"):
            outcome = sync_read.pull(ctx)
        self.assertEqual(outcome.exit_code, 0)
        self.assertTrue(outcome.ok)
        self.assertEqual(outcome.stats["failed"], 0)
        rows = {r["name"]: r for r in outcome.stats["results"]}
        self.assertEqual(rows["示例表A"]["mode"], "open")
        self.assertEqual(rows["示例表C"]["mode"], "disk")
        open_disk.assert_called_once()  # 只有未打开的 B 走磁盘

    def test_pull_apply_open_session_failure_not_fallback_disk(self):
        # 已打开但会话连不上（COM_CONTEXT）：明确错误码，不静默退回磁盘读陈旧副本
        args = parse(["pull", "--project", "示例表A", "--apply", "--json"])
        com_err = Failure("COM 不接话", code=ErrorCode.COM_CONTEXT, stage="attach")
        with mock.patch.object(sync_read.excel_gateway, "attach_running_excel",
                               side_effect=com_err), \
                mock.patch.object(sync_read.excel_gateway, "open_disk_workbook") as open_disk:
            outcome = sync_read.pull(self.fx.ctx(args))
        self.assertEqual(outcome.exit_code, 1)
        self.assertFalse(outcome.ok)
        self.assertEqual(outcome.stats["failed"], 1)
        open_disk.assert_not_called()
        row = outcome.stats["results"][0]
        self.assertFalse(row["ok"])
        self.assertEqual(row["reason"], "COM 不接话")
        self.assertEqual(outcome.errors[0].code, "COM_CONTEXT")
        self.assertEqual(outcome.errors[0].stage, "attach")
        self.assertIn("会话", outcome.errors[0].hint)
        self.assertIn("未退回磁盘", outcome.errors[0].hint)

    def test_pull_apply_disk_open_failure_keeps_contract_code(self):
        # 磁盘路径网关打开失败走单表容错：记失败行；业务码经 to_contract_failure 保留，
        # 不吞成 UNEXPECTED；单表失败不中断整批
        args = parse(["pull", "--project", "示例表A", "--apply", "--json"])
        boom = Failure("磁盘打不开", code="WORKBOOK_OPEN_FAILED", stage="open")
        fake_session = mock.Mock()
        fake_session.app.return_value = mock.Mock()
        no_excel = Failure("没找到正在运行的 Excel", code=ErrorCode.NO_EXCEL, stage="attach")
        with mock.patch.object(sync_read, "ExcelSession", return_value=fake_session), \
                mock.patch.object(sync_read.excel_gateway, "attach_running_excel",
                                  side_effect=no_excel), \
                mock.patch.object(sync_read.excel_gateway, "open_disk_workbook",
                                  side_effect=boom):
            outcome = sync_read.pull(self.fx.ctx(args))
        self.assertEqual(outcome.exit_code, 1)
        self.assertFalse(outcome.ok)
        self.assertEqual(outcome.stats["failed"], 1)
        row = outcome.stats["results"][0]
        self.assertFalse(row["ok"])
        self.assertEqual(row["reason"], "磁盘打不开")
        self.assertEqual(outcome.errors[0].code, "WORKBOOK_OPEN_FAILED")
        self.assertEqual(outcome.errors[0].stage, "open")

    def test_pull_usage_error_neither_project(self):
        with self.assertRaises(SystemExit) as cm:
            parse(["pull"])
        self.assertEqual(cm.exception.code, 2)

    def test_pull_usage_error_mutually_exclusive(self):
        with self.assertRaises(SystemExit) as cm:
            parse(["pull", "--project", "示例表A", "--all-project"])
        self.assertEqual(cm.exception.code, 2)

    def test_pull_json_purity(self):
        args = parse(["pull", "--all-project", "--json"])
        with mock.patch.object(sync_read, "find_excel_with_workbook", return_value=(None, None)):
            outcome = sync_read.pull(self.fx.ctx(args))
        code, stdout, _err = render_json(outcome)
        self.assertEqual(code, 0)
        self.assertEqual(len(stdout.strip().splitlines()), 1)
        payload = json.loads(stdout)
        self.assertTrue(ENVELOPE_KEYS <= set(payload))
        self.assertEqual(payload["action"], "sync.pull")
        self.assertEqual(payload["stats"]["dry"], True)

    def test_pull_results_keys_match_baseline(self):
        """stats.results 形状：dry 行键为 name/ok/dry/mode；1.6.x 起附加大致公共模块范围键。"""
        args = parse(["pull", "--all-project", "--json"])
        with mock.patch.object(sync_read, "find_excel_with_workbook", return_value=(None, None)):
            outcome = sync_read.pull(self.fx.ctx(args))
        code, stdout, _err = render_json(outcome)
        self.assertEqual(code, 0)
        payload = json.loads(stdout)
        self.assertIn("results", payload["stats"])
        self.assertEqual(payload["details"], [])
        self.assertEqual(set(payload["stats"]["results"][0].keys()),
                         {"name", "ok", "dry", "mode", "common_scope", "common_modules"})


class TestPullOpen(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.tmp = Path(tempfile.mkdtemp(prefix="vba-wb-"))
        self.fx = _Fixture(self.tmp)

    def test_pull_open_success_dry(self):
        args = parse(["pull-open", "--project", "示例表A", "--json"])
        fake_wb = mock.Mock(Saved=True, FullName=str(self.tmp / "示例表A V1.xlsm"))
        with mock.patch.object(sync_read, "_connect_open_workbook", return_value=fake_wb):
            outcome = sync_read.pull_open(self.fx.ctx(args))
        self.assertEqual(outcome.exit_code, 0)
        self.assertTrue(outcome.ok)
        self.assertTrue(outcome.stats["dry"])
        self.assertFalse(outcome.stats["wrote_src"])
        self.assertFalse(outcome.stats["saved"])
        self.assertEqual(outcome.stats["workbook_saved"], True)
        self.assertEqual(outcome.details, [])
        self.assertTrue(outcome.stats["results"][0]["dry"])

    def test_pull_open_project_not_found(self):
        args = parse(["pull-open", "--project", "不存在"])
        outcome = sync_read.pull_open(self.fx.ctx(args))
        self.assertEqual(outcome.exit_code, 1)
        self.assertFalse(outcome.ok)
        self.assertEqual(outcome.errors[0].code, "PROJECT_NOT_FOUND")

    def test_pull_open_usage_error(self):
        with self.assertRaises(SystemExit) as cm:
            parse(["pull-open"])
        self.assertEqual(cm.exception.code, 2)

    def test_pull_open_json_purity(self):
        args = parse(["pull-open", "--project", "示例表A", "--json"])
        fake_wb = mock.Mock(Saved=False, FullName=str(self.tmp / "示例表A V1.xlsm"))
        with mock.patch.object(sync_read, "_connect_open_workbook", return_value=fake_wb):
            outcome = sync_read.pull_open(self.fx.ctx(args))
        code, stdout, _err = render_json(outcome)
        self.assertEqual(code, 0)
        self.assertEqual(len(stdout.strip().splitlines()), 1)
        payload = json.loads(stdout)
        self.assertTrue(ENVELOPE_KEYS <= set(payload))
        self.assertEqual(payload["action"], "sync.pull-open")
        self.assertEqual(payload["stats"]["workbook_saved"], False)

    def test_pull_open_results_keys_match_baseline(self):
        """stats.results 保留 1.5.6 旧形状：dry 行键为 name/ok/dry；1.6.x 起附加大致公共模块范围键。"""
        args = parse(["pull-open", "--project", "示例表A", "--json"])
        fake_wb = mock.Mock(Saved=True, FullName=str(self.tmp / "示例表A V1.xlsm"))
        with mock.patch.object(sync_read, "_connect_open_workbook", return_value=fake_wb):
            outcome = sync_read.pull_open(self.fx.ctx(args))
        code, stdout, _err = render_json(outcome)
        self.assertEqual(code, 0)
        payload = json.loads(stdout)
        self.assertIn("results", payload["stats"])
        self.assertEqual(payload["details"], [])
        self.assertEqual(set(payload["stats"]["results"][0].keys()),
                         {"name", "ok", "dry", "common_scope", "common_modules"})


if __name__ == "__main__":
    unittest.main(verbosity=2)
