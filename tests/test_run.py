# -*- coding: utf-8 -*-
"""05 切片：run / watch 编排与 MacroRunner 契约测试。

约束：
  · 不真启动 Excel、不碰业务表；COM 边界（attach/new_excel/Run/Components）用桩对象，
    会话获取、源码预检、写回、回读、探活、比对全部走 toolkit/gateway 真实函数；
  · 测试件只落在临时目录；示例名均为虚构（示例表A / 示例求和 / 目标.xlsm）。

覆盖四类：成功 / 业务失败（exit 1 且 errors[].code）/ 用法错（exit 2）/ JSON 纯度
（run 真 JSON 一行；watch 拒绝 --json，exit 2 且 stdout 无 JSON）。
"""
from __future__ import annotations

import io
import json
import os
import subprocess
import sys
import tempfile
import time
import types
import unittest
from contextlib import contextmanager
from pathlib import Path
from unittest import mock

PKG = Path(__file__).resolve().parent.parent / "scripts"
if str(PKG) not in sys.path:
    sys.path.insert(0, str(PKG))

from adapters import excel_gateway  # noqa: E402
from features import run as run_feature  # noqa: E402
from toolkit import repo  # noqa: E402
from toolkit.base import Failure as BaseFailure  # noqa: E402
from toolkit.contract import ErrorCode, Failure  # noqa: E402
from toolkit.macro import MacroRunner  # noqa: E402
from toolkit.registry import build_parser  # noqa: E402

SRC_CODE = "Public Sub 示例求和()\nEnd Sub"


# --------------------------------------------------------------------------- 桩对象
class _FakeCM:
    """CodeModule 桩：保持行列表，支持 CountOfLines / Lines / DeleteLines / AddFromString。"""

    def __init__(self, code: str = "") -> None:
        self._lines = code.splitlines()

    @property
    def CountOfLines(self) -> int:
        return len(self._lines)

    def Lines(self, start: int, count: int) -> str:
        return "\n".join(self._lines[:max(count, 0)])

    def DeleteLines(self, start: int, count: int) -> None:
        del self._lines[start - 1:start - 1 + count]

    def AddFromString(self, code: str) -> None:
        self._lines = code.splitlines()


class _FakeComp:
    def __init__(self, name: str, ctype: int = repo.VBXT_STANDARD, code: str = "") -> None:
        self.Name = name
        self.Type = ctype
        self.CodeModule = _FakeCM(code)
        self.exported: list[str] = []

    def Export(self, path: str) -> None:
        # 桩：按 VBE 口径落 GBK 产物，供 transcode_exported_file 回读转码。
        self.exported.append(path)
        text = self.CodeModule.Lines(1, self.CodeModule.CountOfLines)
        Path(path).write_bytes(text.encode("gbk"))


class _FakeComponents:
    def __init__(self, comps: list[_FakeComp]) -> None:
        self._comps = list(comps)

    def __iter__(self):
        return iter(self._comps)

    def Item(self, name: str):
        return next((c for c in self._comps if c.Name == name), None)

    def Import(self, path: str):
        comp = _FakeComp(Path(path).stem)
        self._comps.append(comp)
        return comp

    def Remove(self, comp: _FakeComp) -> None:
        self._comps.remove(comp)


class _FakeVBProject:
    def __init__(self, comps: list[_FakeComp]) -> None:
        self.VBComponents = _FakeComponents(comps)

    @property
    def Name(self) -> str:
        return "示例工程"


class _FakeWBList:
    def __init__(self, wbs: list) -> None:
        self._wbs = list(wbs)

    def __iter__(self):
        return iter(self._wbs)


class _FakeApp:
    def __init__(self, ret: object = 5, wbs: list | None = None,
                 active_name: str = "目标.xlsm") -> None:
        self.ActiveWorkbook = types.SimpleNamespace(Name=active_name)
        self.AutomationSecurity = 2
        self.Workbooks = wbs or _FakeWBList([])
        self._ret = ret
        self.calls: list[tuple[str, tuple]] = []

    def Run(self, macro: str, *args):
        self.calls.append((macro, tuple(args)))
        if isinstance(self._ret, Exception):
            raise self._ret
        return self._ret


class _FakeWB:
    def __init__(self, app: _FakeApp, name: str, full_name: str,
                 comps: list[_FakeComp] | None = None, closed: bool = False) -> None:
        self.Application = app
        app.ActiveWorkbook = types.SimpleNamespace(Name=name)
        self._name = name
        self.FullName = full_name
        self.Path = str(Path(full_name).parent)
        self.activated = False
        self.saved = False
        self.VBProject = _FakeVBProject(comps or [])
        self.closed = closed

    @property
    def Name(self) -> str:
        if self.closed:
            raise RuntimeError("closed")
        return self._name

    def Activate(self) -> None:
        self.activated = True

    def Save(self) -> None:
        self.saved = True


class _Prog:
    @contextmanager
    def stage(self, name: str):
        yield


class _Ctx:
    """最小上下文：字段对齐 AppContext，journal_dir 指向临时目录。"""

    def __init__(self, args: object, cfg: dict, root: Path, journal_dir: Path | None = None):
        self.args = args
        self.config = cfg
        self.root = root
        self.json_mode = False
        self.started = time.perf_counter()
        self.progress = _Prog()
        self.journal_dir = journal_dir or (root / ".run")

    def elapsed_ms(self) -> int:
        return 0

    def stage_timings(self) -> dict:
        return {}


# --------------------------------------------------------------------------- 环境搭建
def _make_env(tmp: Path) -> tuple[dict, Path, Path]:
    src = tmp / "srcproj"
    common = tmp / "common"
    search = tmp / "search"
    backup = tmp / "backup"
    for d in (src, common, search, backup):
        d.mkdir(parents=True, exist_ok=True)
    (src / "示例模块.bas").write_text(SRC_CODE, encoding="utf-8")
    xlsm = search / "示例表A V1.xlsm"
    xlsm.write_bytes(b"x")
    cfg = {
        "defaults": {"common_dir": str(common), "backup_dir": str(backup)},
        "projects": [{
            "name": "示例表A",
            "src": str(src),
            "search_dir": str(search),
            "prefix": "示例表A",
        }],
    }
    return cfg, tmp, xlsm


def _parse(argv: list[str]):
    parser = build_parser(run_feature.SPECS, prog="sync")
    return parser.parse_args(argv)


def _patch_attach(test: unittest.TestCase, app: _FakeApp) -> None:
    test.addCleanup(setattr, excel_gateway, "attach_running_excel",
                    excel_gateway.attach_running_excel)
    excel_gateway.attach_running_excel = lambda path=None, **kw: app


# --------------------------------------------------------------------------- MacroRunner 单元
class MacroRunnerTest(unittest.TestCase):
    def _session(self, wb: _FakeWB) -> types.SimpleNamespace:
        return types.SimpleNamespace(workbook=wb, macro_declared=lambda m: True)

    def test_success_attach_sequence(self):
        app = _FakeApp(ret=7)
        wb = _FakeWB(app, "目标.xlsm", r"C:\tmp\目标.xlsm")
        runner = MacroRunner(self._session(wb))
        result = runner.run("'目标.xlsm'!示例求和", [2])
        self.assertTrue(result.ok)
        self.assertEqual(result.return_value, 7)
        self.assertTrue(wb.activated)
        self.assertEqual(app.calls, [("'目标.xlsm'!示例求和", (2,))])

    def test_business_failure_maps_run_failed(self):
        app = _FakeApp(ret=RuntimeError("运行时错误 13：类型不匹配"))
        wb = _FakeWB(app, "目标.xlsm", r"C:\tmp\目标.xlsm")
        runner = MacroRunner(self._session(wb))
        with self.assertRaises(Failure) as cm:
            runner.run("'目标.xlsm'!示例求和")
        self.assertEqual(cm.exception.code, ErrorCode.RUN_FAILED.value)
        self.assertEqual(cm.exception.stage, "run")

    def test_macro_unavailable_maps_run_failed(self):
        # 旧版该兜底分支即 RUN_FAILED；不另造 MACRO_NOT_AVAILABLE 之类的新码。
        app = _FakeApp(ret=RuntimeError("宏“示例求和”在此工作簿中不可用，或已被禁用。"))
        wb = _FakeWB(app, "目标.xlsm", r"C:\tmp\目标.xlsm")
        runner = MacroRunner(self._session(wb))
        with self.assertRaises(Failure) as cm:
            runner.run("'目标.xlsm'!示例求和")
        self.assertEqual(cm.exception.code, ErrorCode.RUN_FAILED.value)

    def test_macro_disabled_maps_run_failed(self):
        app = _FakeApp(ret=RuntimeError("宏都被禁用，因为宏在此文件中被阻止运行。"))
        wb = _FakeWB(app, "目标.xlsm", r"C:\tmp\目标.xlsm")
        runner = MacroRunner(self._session(wb))
        with self.assertRaises(Failure) as cm:
            runner.run("'目标.xlsm'!示例求和")
        self.assertEqual(cm.exception.code, ErrorCode.RUN_FAILED.value)

    def test_arg_mismatch_classified(self):
        app = _FakeApp(ret=RuntimeError("-2147352562 参数个数不符"))
        wb = _FakeWB(app, "目标.xlsm", r"C:\tmp\目标.xlsm")
        runner = MacroRunner(self._session(wb))
        with self.assertRaises(Failure) as cm:
            runner.run("'目标.xlsm'!示例求和", [1, 2])
        self.assertEqual(cm.exception.code, ErrorCode.ARG_MISMATCH.value)

    def test_active_workbook_mismatch_rejected(self):
        app = _FakeApp(ret=5)
        wb = _FakeWB(app, "目标.xlsm", r"C:\tmp\目标.xlsm")
        app.ActiveWorkbook = types.SimpleNamespace(Name="别的表.xlsm")
        runner = MacroRunner(self._session(wb))
        with self.assertRaises(Failure) as cm:
            runner.run("'目标.xlsm'!示例求和")
        self.assertEqual(cm.exception.code, ErrorCode.TARGET_NOT_ACTIVE.value)

    def test_isolate_save_persists(self):
        app = _FakeApp(ret=None)
        wb = _FakeWB(app, "目标.xlsm", r"C:\tmp\目标.xlsm")
        runner = MacroRunner(self._session(wb))
        result = runner.run("'目标.xlsm'!示例求和", isolate=True, save=True)
        self.assertTrue(result.ok)
        self.assertTrue(wb.saved)
        self.assertTrue(result.saved)


# --------------------------------------------------------------------------- run 动词
class RunFeatureTest(unittest.TestCase):
    def setUp(self):
        self._tmp = Path(tempfile.mkdtemp(prefix="vba-run-"))
        self.cfg, self.root, self.xlsm = _make_env(self._tmp)

    def _ctx(self, argv: list[str]) -> _Ctx:
        return _Ctx(_parse(argv), self.cfg, self.root, self.root / ".run")

    def test_dry_run_preview_touches_no_excel(self):
        # 不 patch attach：若编排层真去开表会直接抛 COM 相关错误。
        ctx = self._ctx(["run", "--project", "示例表A", "--macro", "示例求和"])
        outcome = run_feature.run(ctx)
        self.assertEqual(outcome.exit_code, 0)
        self.assertTrue(outcome.stats["dry"])
        self.assertIn("results", outcome.stats)
        self.assertNotIn("cost_sec", outcome.stats)

    def test_isolate_falls_back_on_base_no_excel(self):
        """基类 Failure（NO_EXCEL）也必须命中 isolate 回退，不能穿透成 exit 1。"""
        proj = self.cfg["projects"][0]
        err = BaseFailure("没装 Excel", code="NO_EXCEL", stage="attach")
        excel_gateway.attach_running_excel = lambda path, **kw: (_ for _ in ()).throw(err)
        self.addCleanup(setattr, excel_gateway, "attach_running_excel",
                        excel_gateway.attach_running_excel)
        fake_app = _FakeApp(ret=None)
        fake_wb = _FakeWB(fake_app, "示例表A V1.xlsm", str(self.xlsm))
        excel_gateway.new_excel = lambda hidden=True: fake_app
        self.addCleanup(setattr, excel_gateway, "new_excel", excel_gateway.new_excel)
        excel_gateway.open_disk_workbook = lambda xl, path, read_only=False: fake_wb
        self.addCleanup(setattr, excel_gateway, "open_disk_workbook",
                        excel_gateway.open_disk_workbook)
        orig_vbe = run_feature.excel.ensure_vbe_access
        run_feature.excel.ensure_vbe_access = lambda wb, name, interactive=False: True
        self.addCleanup(setattr, run_feature.excel, "ensure_vbe_access", orig_vbe)

        wb, _xl, mode, owned = run_feature._acquire(
            self.cfg, proj, isolate=True, push=False, macro="示例求和")
        self.assertEqual(mode, "isolate")
        self.assertIs(wb, fake_wb)
        self.assertIsNotNone(owned)

    def test_isolate_non_fallback_contract_failure_propagates(self):
        """反向保护：契约层 COM_CONTEXT 不属可回退口径，原样上抛、不归一吞掉。"""
        proj = self.cfg["projects"][0]
        err = Failure("COM 不接话", code="COM_CONTEXT", stage="attach")
        excel_gateway.attach_running_excel = lambda path, **kw: (_ for _ in ()).throw(err)
        self.addCleanup(setattr, excel_gateway, "attach_running_excel",
                        excel_gateway.attach_running_excel)
        with self.assertRaises(Failure) as cm:
            run_feature._acquire(self.cfg, proj, isolate=True, push=False,
                                 macro="示例求和")
        self.assertEqual(cm.exception.code, "COM_CONTEXT")

    def test_attach_success_real_chain(self):
        app = _FakeApp(ret=5)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm))
        app.Workbooks = _FakeWBList([wb])
        _patch_attach(self, app)

        ctx = self._ctx(["run", "--project", "示例表A", "--macro", "示例求和",
                         "--arg", "2", "--apply"])
        outcome = run_feature.run(ctx)

        self.assertEqual(outcome.exit_code, 0, outcome.errors)
        self.assertTrue(outcome.ok)
        self.assertTrue(wb.activated)
        self.assertEqual(app.calls, [("'示例表A V1.xlsm'!示例求和", (2,))])
        # 旧形状保留：cost_sec / run_sec / result / log_status 都在。
        self.assertIn("cost_sec", outcome.stats)
        self.assertIn("run_sec", outcome.stats)
        self.assertEqual(outcome.stats["result"], 5)
        self.assertIn("log_status", outcome.stats)
        # journal 落盘。
        journals = list((self.root / ".run").glob("*.json"))
        self.assertEqual(len(journals), 1)
        journal = json.loads(journals[0].read_text(encoding="utf-8"))
        self.assertEqual(journal["project"], "示例表A")
        self.assertEqual(journal["exit_code"], 0)

    def test_push_writes_back_real_chain(self):
        comp = _FakeComp("示例模块", code="Public Sub 旧过程()\nEnd Sub")
        app = _FakeApp(ret=None)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm), comps=[comp])
        app.Workbooks = _FakeWBList([wb])
        _patch_attach(self, app)

        ctx = self._ctx(["run", "--project", "示例表A", "--macro", "示例求和",
                         "--push", "--apply"])
        outcome = run_feature.run(ctx)

        self.assertEqual(outcome.exit_code, 0, outcome.errors)
        self.assertTrue(outcome.stats["pushed"])
        # 工程内模块正文已被源码覆盖。
        cm = comp.CodeModule
        self.assertIn("示例求和", cm.Lines(1, cm.CountOfLines))
        self.assertNotIn("旧过程", cm.Lines(1, cm.CountOfLines))
        # 运行前备份已落盘（backup 目录下有复制出来的 xlsm）。
        backups = list((self.root / "backup").rglob("*.xlsm"))
        self.assertTrue(backups, "应产生运行前备份")

    def test_push_postcheck_failure_rolls_back(self):
        # 桩模块拒绝写入（模拟回读对不上）→ postcheck 失败，POSTCHECK_FAILED。
        class _StubbornCM(_FakeCM):
            def AddFromString(self, code: str) -> None:
                pass

        comp = _FakeComp("示例模块", code="Public Sub 别的过程()\nEnd Sub")
        comp.CodeModule = _StubbornCM("Public Sub 别的过程()\nEnd Sub")
        app = _FakeApp(ret=None)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm), comps=[comp])
        app.Workbooks = _FakeWBList([wb])
        _patch_attach(self, app)

        ctx = self._ctx(["run", "--project", "示例表A", "--macro", "示例求和",
                         "--push", "--rollback-on-fail", "--apply"])
        outcome = run_feature.run(ctx)
        self.assertEqual(outcome.exit_code, 1)
        self.assertEqual(outcome.errors[0].code, ErrorCode.POSTCHECK_FAILED.value)

    def test_business_failure_exit1_with_code(self):
        app = _FakeApp(ret=RuntimeError("运行时错误 13：类型不匹配"))
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm))
        app.Workbooks = _FakeWBList([wb])
        _patch_attach(self, app)

        ctx = self._ctx(["run", "--project", "示例表A", "--macro", "示例求和", "--apply"])
        outcome = run_feature.run(ctx)

        self.assertEqual(outcome.exit_code, 1)
        self.assertEqual(outcome.errors[0].code, ErrorCode.RUN_FAILED.value)
        self.assertEqual(outcome.stats["error_code"], ErrorCode.RUN_FAILED.value)
        self.assertEqual(len(outcome.stats["results"]), 1)
        self.assertFalse(outcome.stats["results"][0]["ok"])
        self.assertIn("cost_sec", outcome.stats)

    def test_project_not_found_raises(self):
        # 入口派发把 Failure 归一成 exit 1；这里在 handler 边界核对错误码。
        ctx = self._ctx(["run", "--project", "不存在项目", "--macro", "示例求和", "--apply"])
        with self.assertRaises(Failure) as cm:
            run_feature.run(ctx)
        self.assertEqual(cm.exception.code, ErrorCode.PROJECT_NOT_FOUND.value)

    def test_usage_missing_macro_raises_usage(self):
        args = types.SimpleNamespace(
            project="示例表A", macro=None, arg=[], apply=True, isolate=False,
            save=False, push=False, tail=40, log=None, log_optional=False,
            rollback_on_fail=False)
        ctx = _Ctx(args, self.cfg, self.root)
        with self.assertRaises(Failure) as cm:
            run_feature.run(ctx)
        self.assertEqual(cm.exception.code, ErrorCode.USAGE.value)

    def test_parser_usage_exit2(self):
        parser = build_parser(run_feature.SPECS, prog="sync")
        with self.assertRaises(SystemExit) as cm:
            parser.parse_args(["run"])
        self.assertEqual(cm.exception.code, 2)

    def test_run_json_is_single_line(self):
        from toolkit.render import render
        app = _FakeApp(ret=3)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm))
        app.Workbooks = _FakeWBList([wb])
        _patch_attach(self, app)
        ctx = self._ctx(["run", "--project", "示例表A", "--macro", "示例求和",
                         "--apply", "--json"])
        outcome = run_feature.run(ctx)

        buf = io.StringIO()
        with mock.patch("sys.stdout", buf):
            code = render(outcome, json_mode=True)
        self.assertEqual(code, 0)
        lines = [ln for ln in buf.getvalue().splitlines() if ln.strip()]
        self.assertEqual(len(lines), 1)
        payload = json.loads(lines[0])
        self.assertEqual(payload["exit_code"], 0)

    def _run_with_log(self, *extra: str):
        log = self.root / "测试输出" / "run.log"
        log.parent.mkdir(parents=True, exist_ok=True)
        log.write_text("第一行\n第二行\n第三行\n", encoding="gbk")
        app = _FakeApp(ret=None)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm))
        app.Workbooks = _FakeWBList([wb])
        _patch_attach(self, app)
        ctx = self._ctx(["run", "--project", "示例表A", "--macro", "示例求和",
                         "--apply", "--log", str(log), *extra])
        return run_feature.run(ctx), log

    def test_run_log_target_and_tail_window(self):
        outcome, log = self._run_with_log("--tail", "2")
        self.assertEqual(outcome.exit_code, 0, outcome.errors)
        self.assertEqual(outcome.stats["log_status"], "found")
        self.assertEqual(outcome.stats["log"], str(log))
        text = "\n".join(outcome.human)
        self.assertIn("第二行", text)
        self.assertIn("第三行", text)
        self.assertNotIn("第一行", text)

    def test_run_tail_zero_skips_log(self):
        outcome, log = self._run_with_log("--tail", "0")
        self.assertEqual(outcome.exit_code, 0, outcome.errors)
        self.assertEqual(outcome.stats["log_status"], "")
        self.assertEqual(outcome.stats["log"], "")
        self.assertNotIn("日志尾部", "\n".join(outcome.human))

    def test_run_log_optional_missing_is_normal(self):
        app = _FakeApp(ret=None)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm))
        app.Workbooks = _FakeWBList([wb])
        _patch_attach(self, app)
        ctx = self._ctx(["run", "--project", "示例表A", "--macro", "示例求和",
                         "--apply", "--log-optional", "--tail", "2"])
        outcome = run_feature.run(ctx)
        self.assertEqual(outcome.exit_code, 0, outcome.errors)
        self.assertEqual(outcome.stats["log_status"], "not_applicable")
        self.assertIn("按 --log-optional", "\n".join(outcome.human))

    def test_run_default_log_directory(self):
        log = self.xlsm.parent / "测试输出" / "run.log"
        log.parent.mkdir(parents=True, exist_ok=True)
        log.write_text("第一行\n第二行\n", encoding="gbk")
        app = _FakeApp(ret=None)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm))
        app.Workbooks = _FakeWBList([wb])
        _patch_attach(self, app)
        ctx = self._ctx(["run", "--project", "示例表A", "--macro", "示例求和",
                         "--apply", "--tail", "2"])
        outcome = run_feature.run(ctx)
        self.assertEqual(outcome.exit_code, 0, outcome.errors)
        self.assertEqual(outcome.stats["log_status"], "found")
        self.assertEqual(Path(outcome.stats["log"]), log)


# --------------------------------------------------------------------------- watch 动词
class WatchFeatureTest(unittest.TestCase):
    def setUp(self):
        self._tmp = Path(tempfile.mkdtemp(prefix="vba-watch-"))
        self.cfg, self.root, self.xlsm = _make_env(self._tmp)
        self._orig_sleep = run_feature._sleep
        run_feature._sleep = lambda _gap: None

    def tearDown(self):
        run_feature._sleep = self._orig_sleep

    def _ctx(self, argv: list[str]) -> _Ctx:
        return _Ctx(_parse(argv), self.cfg, self.root, self.root / ".run")

    def test_target_resolution_and_probe_real_chain(self):
        app = _FakeApp(ret=None)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm))
        app.Workbooks = _FakeWBList([wb])
        _patch_attach(self, app)

        ctx = self._ctx(["watch", "--project", "示例表A", "--apply"])
        # 首轮探活成功；第二轮表关闭 → 监听退出。
        def _second_closed():
            wb.closed = True
            app.Workbooks = _FakeWBList([])
        run_feature._sleep = lambda _gap: _second_closed()

        outcome = run_feature.watch(ctx)
        self.assertEqual(outcome.exit_code, 0)
        self.assertEqual(outcome.stats["targets"], 0)
        self.assertIn("written", outcome.stats)
        self.assertIn("log", outcome.stats)

    def test_source_change_writes_back_once(self):
        comp = _FakeComp("示例模块", code=SRC_CODE)
        app = _FakeApp(ret=None)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm), comps=[comp])
        app.Workbooks = _FakeWBList([wb])
        _patch_attach(self, app)

        ctx = self._ctx(["watch", "--project", "示例表A", "--startup", "skip", "--apply"])
        src = self.root / "srcproj" / "示例模块.bas"
        state = {"n": 0}

        def _sleep_once(_gap):
            state["n"] += 1
            if state["n"] == 1:
                src.write_text("Public Sub 新过程()\nEnd Sub\n", encoding="utf-8")
                return
            if state["n"] == 2:
                return
            wb.closed = True
            app.Workbooks = _FakeWBList([])

        run_feature._sleep = _sleep_once
        outcome = run_feature.watch(ctx)

        self.assertEqual(outcome.exit_code, 0)
        self.assertEqual(outcome.stats["written"], 1)
        code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
        self.assertIn("新过程", code)
        self.assertNotIn("示例求和", code)

    def test_startup_src_write_counts_in_summary(self):
        comp = _FakeComp("示例模块", code="Public Sub 旧过程()\nEnd Sub")
        app = _FakeApp(ret=None)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm), comps=[comp])
        app.Workbooks = _FakeWBList([wb])
        _patch_attach(self, app)

        ctx = self._ctx(["watch", "--project", "示例表A", "--startup", "src", "--apply"])

        def _second_closed():
            wb.closed = True
            app.Workbooks = _FakeWBList([])
        run_feature._sleep = lambda _gap: _second_closed()

        outcome = run_feature.watch(ctx)
        self.assertEqual(outcome.exit_code, 0)
        self.assertEqual(outcome.stats["written"], 1)
        self.assertEqual(outcome.stats["skipped"], 0)
        self.assertIn("示例求和", comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines))

    def test_probe_alive_closed_busy_classified(self):
        app = _FakeApp(ret=None)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm))
        app.Workbooks = _FakeWBList([wb])
        target = {"name": "示例表A", "book": "示例表A V1.xlsm", "wb": wb, "xl": app}
        self.assertEqual(run_feature._watch_probe(target, tries=1, gap=0), "alive")

        wb.closed = True
        app.Workbooks = _FakeWBList([])
        self.assertEqual(run_feature._watch_probe(target, tries=1, gap=0), "closed")

        class _RaisingBooks:
            def __iter__(self):
                raise RuntimeError("busy")
        app.Workbooks = _RaisingBooks()
        self.assertEqual(run_feature._watch_probe(target, tries=2, gap=0), "busy")

    def test_no_target_exit1(self):
        cfg = {"defaults": {"common_dir": str(self.root / "common"),
                            "backup_dir": str(self.root / "backup")},
               "projects": [{"name": "示例表A", "src": str(self.root / "srcproj"),
                             "search_dir": str(self.root / "missing-dir"),
                             "prefix": "示例表A"}]}
        ctx = _Ctx(_parse(["watch", "--apply"]), cfg, self.root)
        with self.assertRaises(Failure) as cm:
            run_feature.watch(ctx)
        self.assertEqual(cm.exception.code, ErrorCode.TARGET_NOT_OPEN.value)

    def test_vbe_pull_exports_to_source_real_chain(self):
        # 以表内为准：比对 → Export 到 src 分类目录 → 转码 → 刷新基线。
        comp = _FakeComp("示例模块", code="Public Sub 表内过程()\nEnd Sub")
        app = _FakeApp(ret=None)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm), comps=[comp])
        app.Workbooks = _FakeWBList([wb])
        _patch_attach(self, app)

        ctx = self._ctx(["watch", "--project", "示例表A", "--startup", "vbe", "--apply"])

        def _second_closed():
            wb.closed = True
            app.Workbooks = _FakeWBList([])
        run_feature._sleep = lambda _gap: _second_closed()

        outcome = run_feature.watch(ctx)
        self.assertEqual(outcome.exit_code, 0)
        self.assertEqual(len(comp.exported), 1)
        pulled = Path(comp.exported[0])
        self.assertTrue(pulled.is_file())
        self.assertEqual(pulled.parent.name, "模块")
        self.assertIn("表内过程", pulled.read_text(encoding="utf-8"))

    def test_each_non_interactive_records_and_skips(self):
        from toolkit.base import set_non_interactive
        set_non_interactive(True)
        self.addCleanup(set_non_interactive, False)

        comp = _FakeComp("示例模块", code="Public Sub 表内过程()\nEnd Sub")
        app = _FakeApp(ret=None)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm), comps=[comp])
        app.Workbooks = _FakeWBList([wb])
        _patch_attach(self, app)

        ctx = self._ctx(["watch", "--project", "示例表A", "--startup", "each",
                         "--no-input", "--apply"])

        def _second_closed():
            wb.closed = True
            app.Workbooks = _FakeWBList([])
        run_feature._sleep = lambda _gap: _second_closed()

        outcome = run_feature.watch(ctx)
        self.assertEqual(outcome.exit_code, 0)
        # 非交互：既不抓回表内代码，也不把源码写回工程。
        self.assertEqual(comp.exported, [])
        self.assertIn("表内过程", comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines))
        logs = list((self.root / "backup" / "watch").rglob("*.log"))
        self.assertTrue(logs)
        text = "\n".join(p.read_text(encoding="utf-8") for p in logs)
        self.assertIn("非交互", text)

    def test_each_interactive_picks_vbe(self):
        from toolkit import ui
        self.addCleanup(mock.patch.stopall)
        mock.patch.object(ui, "choose", return_value=2).start()
        self.addCleanup(setattr, run_feature, "interactive_ok", run_feature.interactive_ok)
        run_feature.interactive_ok = lambda: True

        comp = _FakeComp("示例模块", code="Public Sub 表内过程()\nEnd Sub")
        app = _FakeApp(ret=None)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm), comps=[comp])
        app.Workbooks = _FakeWBList([wb])
        _patch_attach(self, app)

        ctx = self._ctx(["watch", "--project", "示例表A", "--startup", "each", "--apply"])

        def _second_closed():
            wb.closed = True
            app.Workbooks = _FakeWBList([])
        run_feature._sleep = lambda _gap: _second_closed()

        outcome = run_feature.watch(ctx)
        self.assertEqual(outcome.exit_code, 0)
        self.assertEqual(len(comp.exported), 1)

    def test_reject_json_exit2_no_json_on_stdout(self):
        args = _parse(["watch", "--project", "示例表A", "--json"])
        self.assertTrue(run_feature.json_rejected(args))
        self.assertEqual(run_feature.precheck(args), 2)

        ctx = _Ctx(args, self.cfg, self.root)
        buf = io.StringIO()
        with mock.patch("sys.stdout", buf):
            outcome = run_feature.watch(ctx)
        self.assertEqual(outcome.exit_code, 2)
        self.assertEqual(outcome.errors[0].code, ErrorCode.USAGE.value)
        self.assertNotIn("{", buf.getvalue())
        self.assertNotIn("}", buf.getvalue())

    def test_watch_reject_json_main_entry_exit2(self):
        """端到端走 sync.main(argv)：watch --json 用法错必须 exit 2，stdout 不出 JSON 信封。"""
        cfg_path = self.root / "mapping.json"
        cfg_path.write_text(json.dumps(self.cfg, ensure_ascii=False), encoding="utf-8")
        env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(cfg_path),
                   PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
        sync_py = PKG / "sync.py"
        for argv in (["watch", "--json"],
                     ["watch", "--project", "示例表A", "--json"]):
            with self.subTest(argv=argv):
                r = subprocess.run([sys.executable, str(sync_py), *argv],
                                   cwd=str(self.root), env=env, capture_output=True,
                                   text=True, encoding="utf-8")
                self.assertEqual(r.returncode, 2,
                                 f"stdout={r.stdout!r} stderr={r.stderr!r}")
                self.assertFalse(r.stdout.strip().startswith("{"),
                                 "用法错 stdout 不得出 JSON 信封")

    def test_parser_bad_choice_exit2(self):
        parser = build_parser(run_feature.SPECS, prog="sync")
        with self.assertRaises(SystemExit) as cm:
            parser.parse_args(["watch", "--startup", "bogus"])
        self.assertEqual(cm.exception.code, 2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
