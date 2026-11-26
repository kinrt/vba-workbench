# -*- coding: utf-8 -*-
"""agent_io 切片契约测试：probe / book / logs / backup / push。

覆盖四类：成功（exit 0）/ 业务失败（exit 1 且 errors[].code）/ 用法错（exit 2）/
``--json`` 单行纯度（stdout 恰好一行、可 json.loads、信封键齐）。

不启动真实 Excel：成功路径走 dry 预览或桩化 handler；业务表只放临时目录。
"""
from __future__ import annotations

import io
import json
import os
import shlex
import subprocess
import sys
import tempfile
import time
import unittest
from argparse import Namespace
from contextlib import redirect_stdout, redirect_stderr
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from toolkit.context import AppContext, Progress  # noqa: E402
from toolkit.base import (PKG_ROOT, entry_command, tool_version,  # noqa: E402
                          Failure as BaseFailure, make_envelope)
from toolkit.contract import ErrorCode  # noqa: E402
import features.agent_io as agent_io  # noqa: E402

ENVELOPE_KEYS = {"ok", "action", "details", "errors", "elapsed_ms", "ts",
                 "stats", "exit_code"}


def _env(config_path: Path) -> dict[str, str]:
    env = dict(os.environ)
    env["VBA_WORKBENCH_CONFIG"] = str(config_path)
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def _run(args: list[str], config_path: Path) -> tuple[int, str, str]:
    proc = subprocess.run(
        [sys.executable, str(SCRIPTS / "agent.py"), *args],
        capture_output=True, text=True, env=_env(config_path),
        encoding="utf-8", errors="replace",
    )
    return proc.returncode, proc.stdout, proc.stderr


MODULE_A = (
    'Attribute VB_Name = "模块A"\n'
    "Option Explicit\n"
    "\n"
    "Public Sub Demo()\n"
    "End Sub\n"
)


def _config_dir(tmp: Path, name: str = "示例表A", prefix: str = "示例表A",
                src: str | None = None) -> Path:
    """写一份临时 mapping.json 和空 xlsm；测试件全部落在临时目录。"""
    common = tmp / "common"
    common.mkdir(exist_ok=True)
    cfg = {
        "projects": [{
            "name": name,
            "search_dir": str(tmp).replace("\\", "/"),
            "prefix": prefix,
            "require_version": True,
            "src": src or ("src/projects/" + name),
        }],
        "defaults": {"common_dir": str(common).replace("\\", "/")},
    }
    (tmp / f"{prefix} V1.xlsm").write_bytes(b"")
    (tmp / "mapping.json").write_text(
        json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
    return tmp / "mapping.json"


def _write_src_module(tmp: Path, name: str = "示例表A") -> Path:
    """在临时目录落一个可通过预检的标准模块源码目录。"""
    folder = tmp / "src" / "projects" / name
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "模块A.bas").write_text(MODULE_A, encoding="utf-8")
    return folder


def _parse_one(stdout: str) -> dict:
    """单行纯度：恰好一行且可解析；返回信封字典。"""
    lines = [ln for ln in stdout.splitlines() if ln.strip()]
    if len(lines) != 1:
        raise AssertionError(f"stdout 不是单行：{len(lines)} 行")
    payload = json.loads(lines[0])
    missing = ENVELOPE_KEYS - set(payload)
    if missing:
        raise AssertionError(f"信封缺键：{missing}")
    return payload


class ProbeContract(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.cfg = _config_dir(self.tmp)

    def tearDown(self):
        self._tmp.cleanup()

    def test_probe_success_json_pure(self):
        rc, out, err = _run(["probe", "--json"], self.cfg)
        self.assertEqual(rc, 0, err)
        payload = _parse_one(out)
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["action"], "agent.probe")
        self.assertEqual(payload["exit_code"], 0)
        self.assertEqual(payload["errors"], [])
        self.assertEqual(payload["stats"]["projects"], 1)
        self.assertEqual(payload["details"][0]["project"], "示例表A")

    def test_probe_unknown_project_business_failure(self):
        rc, out, err = _run(["probe", "--project", "不存在", "--json"], self.cfg)
        self.assertEqual(rc, 1)
        payload = _parse_one(out)
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["exit_code"], 1)
        self.assertEqual(payload["errors"][0]["code"], ErrorCode.PROJECT_NOT_FOUND)
        self.assertEqual(payload["errors"][0]["stage"], "config")

    def test_probe_usage_error_exit_2(self):
        rc, _, _ = _run(["probe", "--bogus-flag", "--json"], self.cfg)
        self.assertEqual(rc, 2)

    def test_probe_stats_lines_restored_timings_is_new_key(self):
        rc, out, err = _run(["probe", "--json"], self.cfg)
        self.assertEqual(rc, 0, err)
        payload = _parse_one(out)
        # stats.lines 是旧版就有的人类行列表（Excel 状态行 + 表头 + 数据行），必须保留
        self.assertIsInstance(payload["stats"]["lines"], list)
        self.assertTrue(payload["stats"]["lines"])
        self.assertTrue(any("项目" in ln for ln in payload["stats"]["lines"]))
        self.assertTrue(any("Excel" in ln for ln in payload["stats"]["lines"]))
        # timings 是 1.6 新增键，单列在差异表，不顶替旧键
        self.assertIn("timings", payload["stats"])


class ShellArgTests(unittest.TestCase):
    """入口壳参数集：与 1.5.6 完全一致，不得新增参数。"""

    def test_version_flag_prints_version_and_exits_0(self):
        # 1.6 三入口统一注册 --version：打印版本行并退出 0（不再是未知参数退 2）
        tmp = tempfile.TemporaryDirectory()
        cfg = _config_dir(Path(tmp.name))
        rc, out, _ = _run(["--version"], cfg)
        self.assertEqual(rc, 0)
        self.assertIn(f"vba-workbench {tool_version()}", out)
        tmp.cleanup()


class BookContract(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.cfg = _config_dir(self.tmp)

    def tearDown(self):
        self._tmp.cleanup()

    def test_book_open_preview_success_json_pure(self):
        rc, out, err = _run(
            ["book", "--action", "open", "--project", "示例表A", "--json"], self.cfg)
        self.assertEqual(rc, 0, err)
        payload = _parse_one(out)
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["action"], "agent.book")
        self.assertEqual(payload["stats"]["dry"], True)
        self.assertTrue(payload["details"][0]["dry"])

    def test_book_unknown_project_business_failure(self):
        rc, out, err = _run(
            ["book", "--action", "close", "--project", "不存在", "--apply", "--json"],
            self.cfg)
        self.assertEqual(rc, 1)
        payload = _parse_one(out)
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["errors"][0]["code"], ErrorCode.PROJECT_NOT_FOUND)

    def test_book_missing_action_usage_error_exit_2(self):
        rc, _, _ = _run(["book", "--project", "示例表A", "--json"], self.cfg)
        self.assertEqual(rc, 2)


class BackupContract(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.cfg = _config_dir(self.tmp)

    def tearDown(self):
        self._tmp.cleanup()

    def test_backup_preview_success_json_pure(self):
        rc, out, err = _run(["backup", "--project", "示例表A", "--json"], self.cfg)
        self.assertEqual(rc, 0, err)
        payload = _parse_one(out)
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["action"], "agent.backup")
        self.assertEqual(payload["stats"]["dry"], True)
        self.assertEqual(payload["stats"]["tag"], "-agent")

    def test_backup_unknown_project_business_failure(self):
        rc, out, err = _run(
            ["backup", "--project", "不存在", "--apply", "--json"], self.cfg)
        self.assertEqual(rc, 1)
        payload = _parse_one(out)
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["errors"][0]["code"], ErrorCode.PROJECT_NOT_FOUND)

    def test_backup_usage_error_exit_2(self):
        rc, _, _ = _run(["backup", "--bogus-flag", "--json"], self.cfg)
        self.assertEqual(rc, 2)


class LogsContract(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.cfg = _config_dir(self.tmp)

    def tearDown(self):
        self._tmp.cleanup()

    def test_logs_success_json_pure(self):
        log = self.tmp / "run_20261122.log"
        log.write_text("通过 3 / 失败 0\n第一行断言\n第二行断言\n", encoding="gbk")
        rc, out, err = _run(
            ["logs", "--project", "示例表A", "--log", str(self.tmp), "--json"], self.cfg)
        self.assertEqual(rc, 0, err)
        payload = _parse_one(out)
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["action"], "agent.logs")
        self.assertEqual(payload["stats"]["logs"], 1)
        self.assertEqual(payload["details"][0]["passed"], "3")
        self.assertEqual(payload["details"][0]["failed"], "0")

    def test_logs_not_found_business_failure(self):
        rc, out, err = _run(
            ["logs", "--project", "示例表A", "--log", str(self.tmp), "--json"], self.cfg)
        self.assertEqual(rc, 1)
        payload = _parse_one(out)
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["errors"][0]["code"], ErrorCode.LOGS_NOT_FOUND)
        self.assertEqual(payload["errors"][0]["stage"], "logs")
        self.assertIn("--log", payload["errors"][0]["hint"])

    def test_logs_usage_error_exit_2(self):
        rc, _, _ = _run(["logs", "--bogus-flag", "--json"], self.cfg)
        self.assertEqual(rc, 2)

    def test_logs_tail_window_and_zero(self):
        log = self.tmp / "run_20261122.log"
        log.write_text("表头\n第一行\n第二行\n第三行\n", encoding="gbk")
        rc, out, err = _run(
            ["logs", "--project", "示例表A", "--log", str(self.tmp),
             "--tail", "2", "--json"], self.cfg)
        self.assertEqual(rc, 0, err)
        payload = _parse_one(out)
        self.assertEqual(payload["stats"]["tail"], 2)
        self.assertEqual(payload["details"][0]["lines"], ["第二行", "第三行"])

        rc, out, err = _run(
            ["logs", "--project", "示例表A", "--log", str(self.tmp),
             "--tail", "0", "--json"], self.cfg)
        self.assertEqual(rc, 0, err)
        payload = _parse_one(out)
        self.assertEqual(payload["stats"]["tail"], 0)
        self.assertEqual(payload["details"][0]["lines"], [])


class PushContract(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        src = _write_src_module(self.tmp)
        self.cfg = _config_dir(
            self.tmp, src=str(src).replace("\\", "/"))

    def tearDown(self):
        self._tmp.cleanup()

    def test_push_preview_success_json_pure(self):
        rc, out, err = _run(["push", "--project", "示例表A", "--json"], self.cfg)
        self.assertEqual(rc, 0, err)
        payload = _parse_one(out)
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["action"], "agent.push")
        self.assertEqual(payload["stats"]["dry"], True)
        self.assertTrue(payload["details"][0]["dry"])
        self.assertIn("modules", payload["details"][0])
        self.assertEqual(payload["stats"]["modules"], 1)

    def test_push_unknown_project_business_failure(self):
        rc, out, err = _run(
            ["push", "--project", "不存在", "--apply", "--json"], self.cfg)
        self.assertEqual(rc, 1)
        payload = _parse_one(out)
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["errors"][0]["code"], ErrorCode.PROJECT_NOT_FOUND)

    def test_push_usage_error_exit_2(self):
        rc, _, _ = _run(["push", "--bogus-flag", "--json"], self.cfg)
        self.assertEqual(rc, 2)

    def test_push_preflight_failure_reports_source_preflight(self):
        # 源码目录不存在：预检先行，attach/备份都不该被碰到
        bad_cfg = _config_dir(
            self.tmp, src=str(self.tmp / "no-such-src").replace("\\", "/"))
        rc, out, err = _run(
            ["push", "--project", "示例表A", "--apply", "--json"], bad_cfg)
        self.assertEqual(rc, 1)
        payload = _parse_one(out)
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["errors"][0]["code"], ErrorCode.SOURCE_PREFLIGHT)
        self.assertEqual(payload["errors"][0]["stage"], "push")


def _ctx(tmp: Path, projects: list[dict], **kwargs) -> AppContext:
    base = dict(command="probe", project=None, json=True, apply=False)
    base.update(kwargs)
    ns = Namespace(**base)
    config = {"projects": projects, "defaults": {}}
    return AppContext(tool="agent", args=ns, config=config, root=Path(tmp),
                      json_mode=True, started=time.perf_counter(),
                      progress=Progress("agent"))


class _FakeWB:
    Path = "fake"
    FullName = "fake.xlsm"

    def Activate(self):
        pass


class _FakeXL:
    Hwnd = 1
    Visible = True
    Workbooks = type("W", (), {"Count": 1})()

    def Quit(self):
        pass


class _FakeCodeModule:
    def __init__(self, text: str):
        self._text = text

    @property
    def CountOfLines(self) -> int:
        return len(self._text.splitlines())

    def Lines(self, start: int, count: int) -> str:
        return self._text


class _FakeComp:
    def __init__(self, name: str, text: str):
        self.Name = name
        self.Type = 1  # vbext_ct_StdModule
        self.CodeModule = _FakeCodeModule(text)


class _FakeVBWB:
    """桩工作簿：记录 Save/Close 调用，VBProject 里挂桩组件。"""

    def __init__(self, comps: list[_FakeComp], full_name: str, calls: list[str]):
        self.VBProject = type("P", (), {"VBComponents": comps})()
        self.Application = object()
        self.FullName = full_name
        self.Path = str(Path(full_name).parent)
        self._calls = calls
        self.saved = False

    def Save(self):
        self.saved = True
        self._calls.append("wb.Save")

    def Close(self, SaveChanges: bool = False):
        self._calls.append("wb.Close")


class HandlerStubbedTests(unittest.TestCase):
    """桩化 COM 的进程内用例：覆盖预览之外的 handler 分支。"""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        src = _write_src_module(self.tmp)
        (self.tmp / "示例表A V1.xlsm").write_bytes(b"")
        self.project = {
            "name": "示例表A",
            "search_dir": str(self.tmp).replace("\\", "/"),
            "prefix": "示例表A",
            "src": str(src).replace("\\", "/"),
        }
        self._orig = {
            "new_excel": agent_io.excel_gateway.new_excel,
            "open_disk": agent_io.excel_gateway.open_disk_workbook,
            "attach": agent_io.excel_gateway.attach_running_excel,
            "find_open": agent_io.excel_gateway.find_open_workbook,
            "backup_copy": agent_io.backup_copy,
            "snapshot": agent_io.component_snapshot,
            "find_comp": agent_io.find_component,
            "remove_mod": agent_io.remove_module,
            "import_comp": agent_io.import_component,
            "replace_code": agent_io.replace_module_code,
            "gbk_copy": agent_io.gbk_import_copy,
            "guard": agent_io.app_guard.guard,
        }

    def tearDown(self):
        agent_io.excel_gateway.new_excel = self._orig["new_excel"]
        agent_io.excel_gateway.open_disk_workbook = self._orig["open_disk"]
        agent_io.excel_gateway.attach_running_excel = self._orig["attach"]
        agent_io.excel_gateway.find_open_workbook = self._orig["find_open"]
        agent_io.backup_copy = self._orig["backup_copy"]
        agent_io.component_snapshot = self._orig["snapshot"]
        agent_io.find_component = self._orig["find_comp"]
        agent_io.remove_module = self._orig["remove_mod"]
        agent_io.import_component = self._orig["import_comp"]
        agent_io.replace_module_code = self._orig["replace_code"]
        agent_io.gbk_import_copy = self._orig["gbk_copy"]
        agent_io.app_guard.guard = self._orig["guard"]
        self._tmp.cleanup()

    def test_book_open_with_stub_gateway_succeeds(self):
        agent_io.excel_gateway.new_excel = lambda **kw: _FakeXL()
        agent_io.excel_gateway.open_disk_workbook = lambda xl, path, read_only=False: _FakeWB()
        ctx = _ctx(self.tmp, [self.project], command="book", project="示例表A",
                   action="open", apply=True)
        out = agent_io.book(ctx)
        self.assertTrue(out.ok, out.errors)
        self.assertEqual(out.exit_code, 0)

    def test_book_close_target_workbook_and_quit_empty_instance(self):
        class _CloseWB:
            FullName = str(self.tmp / "示例表A V1.xlsm")

            def __init__(self):
                self.closed = False

            def Close(self, SaveChanges: bool = False):
                self.closed = True

        class _EmptyBooks:
            Count = 0

            def __iter__(self):
                return iter(())

        class _CloseXL:
            Hwnd = 9101
            Workbooks = _EmptyBooks()

            def __init__(self):
                self.quit = False

            def Quit(self):
                self.quit = True

        wb = _CloseWB()
        xl = _CloseXL()
        agent_io.excel_gateway.attach_running_excel = (
            lambda path=None, **kw: xl)
        agent_io.excel_gateway.find_open_workbook = (
            lambda xl, path: (wb, False))
        ctx = _ctx(self.tmp, [self.project], command="book", project="示例表A",
                   action="close", apply=True)
        out = agent_io.book(ctx)
        self.assertTrue(out.ok, out.errors)
        self.assertEqual(out.exit_code, 0)
        self.assertTrue(wb.closed)
        self.assertTrue(xl.quit)
        self.assertEqual(out.stats["isolate_cleanup"], "ok")

    def test_backup_apply_copies_closed_workbook(self):
        target = self.tmp / "示例表A V1.xlsm"
        target.write_bytes(b"backup-me")
        agent_io.excel_gateway.attach_running_excel = (
            lambda path=None, **kw: (_ for _ in ()).throw(
                BaseFailure("没有 Excel", code="NO_EXCEL", stage="attach")))
        ctx = _ctx(self.tmp, [self.project], command="backup",
                   project="示例表A", apply=True)
        with patch.object(agent_io, "ROOT", self.tmp):
            out = agent_io.backup(ctx)
        self.assertTrue(out.ok, out.errors)
        self.assertTrue(out.stats["backup"])
        copied = list((self.tmp / ".backup").rglob("*.xlsm"))
        self.assertEqual(len(copied), 1)
        self.assertEqual(copied[0].read_bytes(), b"backup-me")

    def test_push_apply_real_writeback_chain(self):
        # 桩 gateway / session / toolkit 原语，断言真实写回链的调用序列
        calls: list[str] = []
        xl = _FakeXL()
        comp = _FakeComp("模块A", MODULE_A)
        wb = _FakeVBWB([comp], str(self.tmp / "示例表A V1.xlsm"), calls)
        imported = {"v": False}
        snap_n = {"n": 0}

        agent_io.excel_gateway.attach_running_excel = (
            lambda path=None, **kw: calls.append("attach") or xl)
        agent_io.excel_gateway.find_open_workbook = (
            lambda xl, path: calls.append("find_open") or (wb, False))
        agent_io.backup_copy = (
            lambda src, dst, tag="": calls.append("backup_copy") or Path(dst) / "bak.xlsm")

        def snapshot(vb):
            calls.append("snapshot")
            snap_n["n"] += 1
            return {} if snap_n["n"] == 1 else {"模块A": 1}
        agent_io.component_snapshot = snapshot

        def find_comp(vb, name):
            calls.append("find_component")
            return comp if imported["v"] else None
        agent_io.find_component = find_comp
        agent_io.remove_module = (
            lambda vb, name, retries=3: calls.append("remove_module") or True)
        agent_io.gbk_import_copy = lambda f, tmp: calls.append("gbk_import_copy") or f

        def imp(vb, path, name, retries=2):
            calls.append("import_component")
            imported["v"] = True
            return comp
        agent_io.import_component = imp
        agent_io.replace_module_code = (
            lambda comp, code: calls.append("replace_module_code"))

        class _Guard:
            changes = {}
            errors: list[str] = []

            def restore(self):
                calls.append("guard.restore")
                return {}

        agent_io.app_guard.guard = lambda app: _Guard()

        ctx = _ctx(self.tmp, [self.project], command="push",
                   project="示例表A", apply=True)
        out = agent_io.push(ctx)

        self.assertTrue(out.ok, out.errors)
        self.assertEqual(out.exit_code, 0)
        self.assertTrue(out.stats["saved"])
        self.assertIn("模块A", out.stats["written"])
        self.assertNotIn(
            ErrorCode.UNSUPPORTED_OPERATION, [e.code for e in out.errors])

        # 序列：附着 → 找表 → 备份 → 写回前快照 → 查找/导入 → 写回后快照 →
        #       回读查找 → 还原应用状态（单次） → 落盘保存
        def idx(label: str) -> int:
            return calls.index(label)

        self.assertLess(idx("attach"), idx("find_open"))
        self.assertLess(idx("find_open"), idx("backup_copy"))
        snap_pos = [i for i, c in enumerate(calls) if c == "snapshot"]
        self.assertEqual(len(snap_pos), 2)
        self.assertLess(snap_pos[0], idx("import_component"))
        self.assertLess(idx("import_component"), snap_pos[1])
        find_pos = [i for i, c in enumerate(calls) if c == "find_component"]
        self.assertGreaterEqual(len(find_pos), 2)
        self.assertLess(idx("import_component"), find_pos[-1])
        self.assertLess(find_pos[-1], idx("guard.restore"))
        self.assertLess(idx("guard.restore"), idx("wb.Save"))
        self.assertTrue(wb.saved)

    def test_push_path_has_no_unsupported_operation_seam(self):
        # 不注入任何写回引擎：handler 源码里不得再留 UNSUPPORTED_OPERATION 接缝
        text = Path(agent_io.__file__).read_text(encoding="utf-8")
        self.assertNotIn("UNSUPPORTED_OPERATION", text)


class InterruptSignalTest(unittest.TestCase):
    """派发期间 Ctrl+C：agent 入口必须返回 130、提示「已中断」，不出 JSON 信封。"""

    def _run_interrupt(self, tmp: Path, *argv: str):
        import agent as agent_entry  # noqa: E402
        from toolkit.base import set_json_mode  # noqa: E402

        cfg = tmp / "mapping.json"
        cfg.write_text(json.dumps({"projects": []}, ensure_ascii=False), encoding="utf-8")
        old_env = os.environ.get("VBA_WORKBENCH_CONFIG")
        os.environ["VBA_WORKBENCH_CONFIG"] = str(cfg)
        orig = agent_entry.dispatch

        def _boom(specs, args, ctx):
            raise KeyboardInterrupt

        agent_entry.dispatch = _boom
        buf, err = io.StringIO(), io.StringIO()
        try:
            with redirect_stdout(buf), redirect_stderr(err):
                code = agent_entry.main(list(argv))
        finally:
            agent_entry.dispatch = orig
            set_json_mode(False)
            if old_env is None:
                os.environ.pop("VBA_WORKBENCH_CONFIG", None)
            else:
                os.environ["VBA_WORKBENCH_CONFIG"] = old_env
        return code, buf.getvalue(), err.getvalue()

    def test_agent_keyboard_interrupt_exits_130_without_envelope(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            # 非 JSON：人话含「已中断」，stdout 不出信封
            code, out, err = self._run_interrupt(d, "probe")
            self.assertEqual(code, 130, err)
            self.assertIn("已中断", out)
            self.assertFalse(out.strip().startswith("{"))
            # JSON 模式：同样不得吐半截信封，「已中断」走 stderr
            code2, out2, err2 = self._run_interrupt(d, "probe", "--json")
            self.assertEqual(code2, 130, err2)
            self.assertFalse(out2.strip(), f"JSON 中断不得出 stdout：{out2!r}")
            self.assertIn("已中断", err2)


class EntryCommandTest(unittest.TestCase):
    """绝对入口命令契约：suggested_command 必须由 entry_command() 生成，安装后可直接复制。"""

    SOURCES = (
        SCRIPTS / "features" / "agent_io.py",
        SCRIPTS / "features" / "agent_patch.py",
        SCRIPTS / "features" / "run.py",
    )

    def test_entry_command_uses_absolute_interpreter_and_script(self):
        cmd = entry_command("agent.py", "probe", "--json")
        script = str(PKG_ROOT / "scripts" / "agent.py")
        # 含本包入口脚本绝对路径与当前解释器名
        self.assertIn(script, cmd)
        self.assertIn(Path(sys.executable).name, cmd)
        tokens = shlex.split(cmd, posix=False)
        self.assertTrue(Path(tokens[0].strip('"')).is_absolute())
        self.assertTrue(Path(tokens[1].strip('"')).is_absolute())

    def test_envelope_keeps_absolute_suggested_command(self):
        cmd = entry_command("agent.py", "probe", "--json")
        env = make_envelope("agent.probe", False, errors=["探测失败"],
                            suggested_command=cmd)
        row = env["errors"][0]
        # error 行 suggested_command 非空且仍是绝对入口命令（不被降级成 hint）
        self.assertTrue(row["suggested_command"])
        self.assertEqual(row["suggested_command"], cmd)
        tokens = shlex.split(row["suggested_command"], posix=False)
        self.assertTrue(Path(tokens[1].strip('"')).is_absolute())

    def test_sources_have_no_relative_package_command(self):
        for src in self.SOURCES:
            text = src.read_text(encoding="utf-8")
            self.assertNotIn("python vba-workbench", text)
            self.assertNotIn("python vba-workbench\\scripts", text)


if __name__ == "__main__":
    unittest.main(verbosity=2)
