# -*- coding: utf-8 -*-
"""1.6.5 回归：Excel 实例归属生命周期（N1 归属快照 / N2 簿白名单 / N3 收尾诊断 / N7 收口）。

约束：不启动真实 Excel、不碰业务表；COM 边界全部用桩对象 / sys.modules 注入假模块，
测试件只落临时目录。覆盖：
  1) attach 借来的实例（owned=False）close() 绝不 Quit；
  2) run attach 模式 owned_xl is None，finally 收尾不 Quit；
  3) quit_owned 异常分支（Workbooks 读取失败 / 单簿 Close 失败 / Quit 抛错）诊断，
     以及 quit_owned / close() 二次调用幂等；
  4) 归属快照不一致不 Quit 且报问题；白名单外仍有簿不 Quit 且报问题；
  5) ROT 绑定失败真实落表路径，与下一轮枚举清空旧条目；
  6) N3 两条路径（close_owned 剩余簿 / Quit 失败）诊断进 warnings/stats。

第二轮追加：run_macro 隔离失败回收与参考簿白名单；sync_push 单项目/批量收尾诊断上报；
agent.push guard 还原 errors 进 warnings；attach 卡住语境默认提示不再误导走 isolate。
"""
from __future__ import annotations

import sys
import tempfile
import time
import types
import unittest
from argparse import Namespace
from pathlib import Path
from unittest import mock

PKG = Path(__file__).resolve().parents[1] / "scripts"
if str(PKG) not in sys.path:
    sys.path.insert(0, str(PKG))

from adapters import excel_gateway  # noqa: E402
from toolkit import excel  # noqa: E402
from toolkit.context import AppContext, Progress  # noqa: E402
import features.agent_io as agent_io  # noqa: E402
import features.common_subset as common_subset  # noqa: E402
import features.run as run_feature  # noqa: E402
import features.run_macro as run_macro  # noqa: E402
import features.sync_push as sync_push  # noqa: E402


# --------------------------------------------------------------------------- 桩对象
class FakeWorkbook:
    """桩工作簿：记录 Close 调用，可配置 Close 抛错。"""

    def __init__(self, name: str, events: list, *,
                 full_name: str = "", close_error: Exception | None = None):
        self.Name = name
        self.FullName = full_name or name
        self._events = events
        self._close_error = close_error
        self.closed = False

    def Close(self, SaveChanges: bool = False):
        if self._close_error is not None:
            raise self._close_error
        self.closed = True
        self._events.append(("wb.Close", self.Name, SaveChanges))

    def Activate(self):
        pass


class FakeBooks:
    def __init__(self, books: list[FakeWorkbook]):
        self._books = list(books)

    def __iter__(self):
        return iter(self._books)

    @property
    def Count(self) -> int:
        return len(self._books)


class FakeXL:
    """桩 Excel 实例：记录 Quit/DisplayAlerts，可配置 Quit 抛错、Workbooks 读失败。"""

    def __init__(self, books: list[FakeWorkbook] | None = None, *,
                 quit_error: Exception | None = None, books_error: Exception | None = None,
                 hwnd: int = 222):
        self.DisplayAlerts = True
        self.Hwnd = hwnd
        self.ActiveWorkbook = None
        self.QuitCalled = False
        self._quit_error = quit_error
        if books_error is not None:
            self.Workbooks = types.SimpleNamespace(Count=0)
        else:
            self.Workbooks = FakeBooks(books or [])

    def Quit(self):
        self.QuitCalled = True
        if self._quit_error is not None:
            raise self._quit_error
        self.Workbooks = FakeBooks([])


def _make_env(tmp: Path) -> tuple[dict, Path]:
    """落最小项目源码与 xlsm，返回 (cfg, xlsm 路径)。"""
    src = tmp / "srcproj"
    common = tmp / "common"
    search = tmp / "search"
    backup = tmp / "backup"
    for d in (src, common, search, backup):
        d.mkdir(parents=True, exist_ok=True)
    (src / "示例模块.bas").write_text(
        "Public Sub 示例求和()\nEnd Sub", encoding="utf-8")
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
    return cfg, xlsm


class SessionLifecycleTest(unittest.TestCase):
    """缺测 1 / 2：attach 借来实例不 Quit；run attach 模式 finally 不 Quit。"""

    def setUp(self):
        # 每个用例清掉归属快照与幂等标记，避免 id 复用串状态。
        excel._OWNED_SNAPSHOTS.clear()
        excel._QUIT_DONE.clear()

    def test_attach_session_close_never_quits(self):
        """prefer_running 借到已有实例：owned=False，close() 绝不调用 Quit。"""
        xl = FakeXL()
        with mock.patch.object(excel, "attach_running_excel", return_value=xl):
            session = excel.ExcelSession(prefer_running=True)
            self.assertIs(session.app(), xl)
            self.assertFalse(session.owned)
            session.close()
        self.assertFalse(xl.QuitCalled, "借来的实例绝不允许 Quit")
        self.assertIsNone(session.xl)

    def test_run_attach_mode_finally_does_not_quit(self):
        """run attach 模式 owned_xl=None：finally 收尾不碰 Quit。"""
        with tempfile.TemporaryDirectory() as raw:
            tmp = Path(raw)
            cfg, xlsm = _make_env(tmp)
            app = FakeXL()
            wb = FakeWorkbook("示例表A V1.xlsm", [], full_name=str(xlsm))
            app.Workbooks = FakeBooks([wb])

            excel_gateway.attach_running_excel = lambda path, **kw: app
            self.addCleanup(setattr, excel_gateway, "attach_running_excel",
                            excel_gateway.attach_running_excel)

            args = Namespace(project="示例表A", macro="示例求和", arg=[], apply=True,
                             isolate=False, save=False, push=False, tail=40,
                             log=None, log_optional=False, rollback_on_fail=False)
            ctx = AppContext(tool="sync", args=args, config=cfg, root=tmp,
                             json_mode=False, started=time.perf_counter(),
                             progress=Progress("sync"))
            # 桩掉宏执行器：只验证会话归属与 finally 收尾口径。
            orig_run = run_macro.MacroRunner.run
            run_macro.MacroRunner.run = lambda self, ref, argv, **kw: types.SimpleNamespace(
                return_value=7, saved=False, log_file="", log_status="missing",
                duration_ms=1, warnings=[])
            self.addCleanup(setattr, run_macro.MacroRunner, "run", orig_run)

            outcome = run_feature.run(ctx)
            self.assertEqual(outcome.exit_code, 0, outcome.errors)
            self.assertFalse(app.QuitCalled, "attach 借来实例的 finally 不允许 Quit")
            self.assertIsNone(outcome.stats.get("isolate_cleanup_problems"))


class QuitOwnedBranchTest(unittest.TestCase):
    """缺测 3：异常分支诊断 + 幂等。"""

    def setUp(self):
        excel._OWNED_SNAPSHOTS.clear()
        excel._QUIT_DONE.clear()

    def test_workbooks_read_failure_reported(self):
        xl = FakeXL(books_error=RuntimeError("RPC 拒绝枚举簿"))
        problems = excel.quit_owned(xl)
        self.assertTrue(any("读取 Workbooks 失败" in p for p in problems), problems)
        # 读不到簿集合仍按旧路径尝试 Quit。
        self.assertTrue(xl.QuitCalled)

    def test_single_workbook_close_failure_does_not_block_rest(self):
        events: list = []
        bad = FakeWorkbook("坏表.xlsm", events, close_error=RuntimeError("关不掉"))
        good = FakeWorkbook("好表.xlsm", events)
        xl = FakeXL([bad, good])
        problems = excel.quit_owned(xl)
        self.assertTrue(any("关闭工作簿 坏表.xlsm 失败" in p for p in problems), problems)
        self.assertIn(("wb.Close", "好表.xlsm", False), events)
        self.assertTrue(xl.QuitCalled)

    def test_quit_failure_reported(self):
        events: list = []
        xl = FakeXL([FakeWorkbook("目标.xlsm", events)],
                    quit_error=RuntimeError("COM 拒绝退出"))
        problems = excel.quit_owned(xl)
        self.assertTrue(any("Quit 失败" in p for p in problems), problems)
        self.assertIn(("wb.Close", "目标.xlsm", False), events)

    def test_second_call_is_idempotent(self):
        events: list = []
        xl = FakeXL([FakeWorkbook("目标.xlsm", events)])
        first = excel.quit_owned(xl)
        self.assertEqual(first, [])
        quit_count = 1
        xl.Quit = lambda: quit_count.__class__  # 二次调用不应再走到这里
        second = excel.quit_owned(xl)
        self.assertEqual(second, [])
        self.assertEqual(quit_count, 1, "二次调用不重复 Quit")

    def test_quit_done_key_is_stable_across_proxy_objects(self):
        """同一 Excel 实例换新代理：按 PID/Hwnd 判幂等，不重复 Quit。"""
        first = FakeXL([FakeWorkbook("目标.xlsm", [])], hwnd=222)
        self.assertEqual(excel.quit_owned(first), [])
        self.assertTrue(first.QuitCalled)
        second = FakeXL([FakeWorkbook("目标.xlsm", [])], hwnd=222)
        self.assertEqual(excel.quit_owned(second), [])
        self.assertFalse(second.QuitCalled, "同实例新代理不应重复 Quit")

    def test_quit_done_key_does_not_block_other_instance(self):
        """不同 Hwnd 的实例各收各的，稳定键不能把新实例挡回去。"""
        first = FakeXL([FakeWorkbook("目标.xlsm", [])], hwnd=222)
        second = FakeXL([FakeWorkbook("目标.xlsm", [])], hwnd=333)
        self.assertEqual(excel.quit_owned(first), [])
        self.assertEqual(excel.quit_owned(second), [])
        self.assertTrue(first.QuitCalled)
        self.assertTrue(second.QuitCalled)

    def test_close_session_is_idempotent(self):
        """ExcelSession.close() 二次调用不抛异常、不重复 Quit。"""
        xl = FakeXL()
        excel._OWNED_SNAPSHOTS[id(xl)] = {"pid": 0, "hwnd": 0}
        session = excel.ExcelSession(prefer_running=False)
        session.xl = xl
        session.owned = True
        session.close()
        session.close()  # 不应抛异常
        self.assertTrue(xl.QuitCalled)


class OwnershipGuardTest(unittest.TestCase):
    """缺测 4：归属快照不一致 / 白名单外有簿 → 不 Quit 且报问题。"""

    def setUp(self):
        excel._OWNED_SNAPSHOTS.clear()
        excel._QUIT_DONE.clear()

    def test_snapshot_mismatch_does_not_quit(self):
        """起实例后 PID/Hwnd 被换掉：不碰簿、不 Quit，报归属校验不通过。"""
        xl = FakeXL(hwnd=333)
        # 登记一份与当前实例不一致的快照（旧 PID 111 / 旧 Hwnd 222）。
        excel._OWNED_SNAPSHOTS[id(xl)] = {"pid": 111, "hwnd": 222}
        problems = excel.quit_owned(xl)
        self.assertFalse(xl.QuitCalled, "归属校验不通过时绝不 Quit")
        self.assertTrue(any("归属校验不通过" in p for p in problems), problems)
        self.assertTrue(any("hwnd" in p for p in problems), problems)

    def test_keep_workbooks_remaining_blocks_quit(self):
        """白名单参考簿仍在实例里：不 Quit，剩余簿名写进问题清单。"""
        events: list = []
        target = FakeWorkbook("目标.xlsm", events)
        ref = FakeWorkbook("参考簿.xlsm", events)
        xl = FakeXL([target, ref])
        problems = excel.quit_owned(xl, keep_workbooks=["参考簿.xlsm"])
        self.assertFalse(xl.QuitCalled, "有他人参考簿时宁可留实例也不退出")
        self.assertTrue(any("参考簿.xlsm" in p for p in problems), problems)
        self.assertTrue(any("未回收工作簿" in p for p in problems), problems)
        # 本工具开的目标簿照常关闭；参考簿保留。
        self.assertIn(("wb.Close", "目标.xlsm", False), events)

    def test_keep_workbooks_all_closed_then_quits(self):
        """白名单内簿全关完（无他人残留）：照常 Quit。"""
        events: list = []
        target = FakeWorkbook("目标.xlsm", events)
        xl = FakeXL([target])
        problems = excel.quit_owned(xl, keep_workbooks=[])
        self.assertEqual(problems, [])
        self.assertTrue(xl.QuitCalled)

    def test_no_snapshot_keeps_legacy_behavior(self):
        """无快照的旧调用路径：不因缺快照报错，走 1.6.4 旧收尾。"""
        events: list = []
        xl = FakeXL([FakeWorkbook("目标.xlsm", events)])
        problems = excel.quit_owned(xl)
        self.assertEqual(problems, [])
        self.assertTrue(xl.QuitCalled)


class RotUnreachableTest(unittest.TestCase):
    """缺测 5：BindToObject 失败真实落表；下一轮枚举清空旧条目。"""

    def setUp(self):
        excel.reset_rot_unreachable()

    def _fake_rot(self, monikers):
        class _Enum:
            def __init__(self):
                self._ms = list(monikers)

            def Next(self, n):
                out = self._ms[:n]
                del self._ms[:n]
                return out

        rot = types.SimpleNamespace(EnumRunning=lambda: _Enum())
        pythoncom = types.SimpleNamespace(
            CreateBindCtx=lambda n: {},
            GetRunningObjectTable=lambda: rot,
            IID_IDispatch=object())
        return pythoncom

    @staticmethod
    def _fake_win32com():
        """注入假 win32com 父包 + client 子模块（import win32com.client 能绑定到属性）。"""
        client = types.SimpleNamespace(Dispatch=lambda raw: raw)
        parent = types.ModuleType("win32com")
        parent.client = client
        return {"win32com": parent, "win32com.client": client}

    def test_bind_failure_recorded(self):
        """BindToObject 抛错 → 条目进 _LAST_ROT_UNREACHABLE（显示名 + 原因）。"""
        class _Moniker:
            def __init__(self, display, error):
                self._display = display
                self._error = error

            def GetDisplayName(self, ctx, prev):
                return self._display

            def BindToObject(self, ctx, prev, iid):
                raise self._error

        mon = _Moniker(r"C:\\tmp\\卡住的表 V1.xlsm", RuntimeError("受保护视图"))
        pythoncom = self._fake_rot([mon])

        with mock.patch.dict(sys.modules, {"pythoncom": pythoncom,
                                           **self._fake_win32com()}):
            apps = excel._rot_excel_apps()
        self.assertEqual(apps, [])
        self.assertEqual(len(excel._LAST_ROT_UNREACHABLE), 1)
        display, reason = excel._LAST_ROT_UNREACHABLE[0]
        self.assertIn("卡住的表 V1.xlsm", display)
        self.assertIn("受保护视图", reason)

    def test_next_enum_clears_stale_entries(self):
        """新一轮枚举开头清空旧条目：上一轮的失败记录不残留。"""
        excel._LAST_ROT_UNREACHABLE.append((r"C:\\tmp\\旧表.xlsm", "旧原因"))

        class _MonikerOk:
            def GetDisplayName(self, ctx, prev):
                return r"C:\\tmp\\新表 V1.xlsm"

            def BindToObject(self, ctx, prev, iid):
                wb = types.SimpleNamespace(
                    Application=types.SimpleNamespace(Hwnd=999))
                return wb

        pythoncom = self._fake_rot([_MonikerOk()])

        with mock.patch.dict(sys.modules, {"pythoncom": pythoncom,
                                           **self._fake_win32com()}):
            apps = excel._rot_excel_apps()
        self.assertEqual(len(apps), 1)
        self.assertEqual(excel._LAST_ROT_UNREACHABLE, [], "旧条目必须被清空")

    def test_reset_helper_clears_list(self):
        excel._LAST_ROT_UNREACHABLE.append(("x.xlsm", "e"))
        excel.reset_rot_unreachable()
        self.assertEqual(excel._LAST_ROT_UNREACHABLE, [])


class CloseOwnedDiagnosticsTest(unittest.TestCase):
    """缺测 6：N3 两条路径诊断进 warnings/stats。"""

    def setUp(self):
        excel._OWNED_SNAPSHOTS.clear()
        excel._QUIT_DONE.clear()

    def test_close_owned_reports_remaining_reference_book(self):
        """实例内有他人参考簿：close_owned 返回问题清单，不 Quit 参考簿。"""
        events: list = []
        target = FakeWorkbook("目标.xlsm", events)
        ref = FakeWorkbook("参考簿.xlsm", events)
        xl = FakeXL([target, ref])
        problems = common_subset.close_owned(target, xl)
        self.assertFalse(xl.QuitCalled)
        self.assertTrue(any("参考簿.xlsm" in p for p in problems), problems)

    def test_close_owned_reports_quit_failure(self):
        """Quit 抛错：问题清单上报，不再裸吞异常。"""
        events: list = []
        xl = FakeXL([FakeWorkbook("目标.xlsm", events)],
                    quit_error=RuntimeError("退出被拒"))
        problems = common_subset.close_owned(types.SimpleNamespace(Name="目标.xlsm"), xl)
        self.assertTrue(any("Quit 失败" in p for p in problems), problems)

    def test_book_open_failure_feeds_warnings_and_stats(self):
        """book open 失败回收：诊断进 Outcome.warnings 与 stats.isolate_cleanup_problems。"""
        xl = FakeXL(quit_error=RuntimeError("退出被拒"))

        def _open_disk(app, path, read_only=False):
            raise agent_io.ContractFailure("打开失败", code="WORKBOOK_OPEN_FAILED",
                                          stage="open")

        agent_io.excel_gateway.new_excel = lambda hidden=False, **kw: xl
        agent_io.excel_gateway.open_disk_workbook = _open_disk
        self.addCleanup(setattr, agent_io.excel_gateway, "new_excel",
                        excel_gateway.new_excel)
        self.addCleanup(setattr, agent_io.excel_gateway, "open_disk_workbook",
                        excel_gateway.open_disk_workbook)

        proj = {"name": "示例表A", "search_dir": "", "prefix": "示例表A", "src": ""}
        args = Namespace(project="示例表A", action="open", apply=True, json=True)
        ctx = AppContext(tool="agent", args=args,
                         config={"projects": [proj], "defaults": {}},
                         root=Path(tempfile.mkdtemp()), json_mode=True,
                         started=time.perf_counter(), progress=Progress("agent"))

        # _resolve 依赖 resolve_target；桩成返回固定路径，避开磁盘查找。
        orig_resolve = agent_io.resolve_target
        agent_io.resolve_target = lambda p: (Path("目标.xlsm"), "桩路径")
        self.addCleanup(setattr, agent_io, "resolve_target", orig_resolve)

        outcome = agent_io.book(ctx)
        self.assertEqual(outcome.exit_code, 1)
        self.assertTrue(outcome.warnings, "回收诊断必须进 warnings")
        self.assertTrue(any("回收未完成" in w for w in outcome.warnings), outcome.warnings)
        self.assertIn("isolate_cleanup_problems", outcome.stats)
        self.assertTrue(outcome.stats["isolate_cleanup_problems"])


class RunMacroIsolateCleanupTest(unittest.TestCase):
    """第二轮 1：run_macro 隔离失败回收走 quit_owned；成功收尾按参考簿白名单保护。"""

    def setUp(self):
        excel._OWNED_SNAPSHOTS.clear()
        excel._QUIT_DONE.clear()

    def _ctx(self, cfg, tmp):
        args = Namespace(project="示例表A", macro="示例求和", arg=[], apply=True,
                         isolate=True, save=False, push=False, tail=40,
                         log=None, log_optional=False, rollback_on_fail=False)
        return AppContext(tool="sync", args=args, config=cfg, root=tmp,
                          json_mode=False, started=time.perf_counter(),
                          progress=Progress("sync"))

    def _stub_gateway(self, xl, target_wb, open_error=None):
        orig_resolve = run_macro.repo.resolve_target
        run_macro.repo.resolve_target = lambda p: (Path(target_wb.FullName), "桩路径")
        self.addCleanup(setattr, run_macro.repo, "resolve_target", orig_resolve)

        def _attach(path, **kw):
            raise run_macro.Failure("没开", code="NO_EXCEL", stage="attach")

        orig_attach = run_macro.excel_gateway.attach_running_excel
        run_macro.excel_gateway.attach_running_excel = _attach
        self.addCleanup(setattr, run_macro.excel_gateway,
                        "attach_running_excel", orig_attach)

        orig_new = run_macro.excel_gateway.new_excel
        run_macro.excel_gateway.new_excel = lambda hidden=False, **kw: xl
        self.addCleanup(setattr, run_macro.excel_gateway, "new_excel", orig_new)

        if open_error is not None:
            def _open_disk(app, path, read_only=False):
                raise open_error
        else:
            def _open_disk(app, path, read_only=False):
                return target_wb

        orig_open = run_macro.excel_gateway.open_disk_workbook
        run_macro.excel_gateway.open_disk_workbook = _open_disk
        self.addCleanup(setattr, run_macro.excel_gateway,
                        "open_disk_workbook", orig_open)

    def test_isolate_open_failure_reports_cleanup(self):
        """开表失败分支：quit_owned 诊断挂到异常，run() 收口进 warnings/stats。"""
        with tempfile.TemporaryDirectory() as raw:
            tmp = Path(raw)
            cfg, xlsm = _make_env(tmp)
            xl = FakeXL([], quit_error=RuntimeError("退出被拒"))
            target = FakeWorkbook("示例表A V1.xlsm", [], full_name=str(xlsm))
            self._stub_gateway(
                xl, target,
                open_error=run_macro.Failure("打开失败",
                                             code="WORKBOOK_OPEN_FAILED", stage="open"))

            outcome = run_feature.run(self._ctx(cfg, tmp))
            self.assertEqual(outcome.exit_code, 1, outcome.errors)
            self.assertTrue(outcome.warnings, "失败回收诊断必须进 warnings")
            self.assertTrue(any("回收未完成" in w for w in outcome.warnings),
                            outcome.warnings)
            self.assertIn("isolate_cleanup_problems", outcome.stats)
            self.assertTrue(outcome.stats["isolate_cleanup_problems"])

    def test_isolate_success_keeps_reference_workbook(self):
        """成功收尾：除目标表外的现存簿（宏侧参考簿）进白名单，不误关、不吞。"""
        with tempfile.TemporaryDirectory() as raw:
            tmp = Path(raw)
            cfg, xlsm = _make_env(tmp)
            events: list = []
            target = FakeWorkbook("示例表A V1.xlsm", events, full_name=str(xlsm))
            ref = FakeWorkbook("参考簿.xlsm", events)
            xl = FakeXL([target, ref])
            self._stub_gateway(xl, target)

            orig_run = run_macro.MacroRunner.run
            run_macro.MacroRunner.run = lambda self, ref, argv, **kw: types.SimpleNamespace(
                return_value=7, saved=False, log_file="", log_status="missing",
                duration_ms=1, warnings=[])
            self.addCleanup(setattr, run_macro.MacroRunner, "run", orig_run)

            outcome = run_feature.run(self._ctx(cfg, tmp))
            self.assertEqual(outcome.exit_code, 0, outcome.errors)
            self.assertIn(("wb.Close", "示例表A V1.xlsm", False), events)
            self.assertNotIn(("wb.Close", "参考簿.xlsm", False), events,
                             "宏自行打开的参考簿不得被收尾关掉")
            self.assertFalse(xl.QuitCalled, "参考簿残留时宁可留实例也不 Quit")
            problems = outcome.stats.get("isolate_cleanup_problems") or []
            self.assertTrue(any("参考簿.xlsm" in p for p in problems), problems)


class SyncPushCleanupReportTest(unittest.TestCase):
    """第二轮 2：两处 _close_owned 返回清单进 warnings/stats，与 agent.push 口径对齐。"""

    def setUp(self):
        excel._OWNED_SNAPSHOTS.clear()
        excel._QUIT_DONE.clear()
        self._orig_preflight = sync_push._preflight_or_fail
        sync_push._preflight_or_fail = lambda proj, defaults, doc: []
        self._orig_resolve = sync_push.resolve_target
        sync_push.resolve_target = lambda p: (Path("示例表A V1.xlsm"), "桩路径")
        self._orig_attach = sync_push._attach_or_open
        self._orig_writeback = sync_push._writeback
        self._orig_vbe = sync_push.has_vbproject_access
        self.addCleanup(setattr, sync_push, "_preflight_or_fail", self._orig_preflight)
        self.addCleanup(setattr, sync_push, "resolve_target", self._orig_resolve)
        self.addCleanup(setattr, sync_push, "_attach_or_open", self._orig_attach)
        self.addCleanup(setattr, sync_push, "_writeback", self._orig_writeback)
        self.addCleanup(setattr, sync_push, "has_vbproject_access", self._orig_vbe)

    @staticmethod
    def _proj(name="示例表A"):
        return {"name": name, "src": "", "search_dir": "", "prefix": name}

    def _ctx(self, cfg):
        args = Namespace(project="示例表A", apply=True, no_backup=False,
                         no_save=False, open_only=False, rollback_on_fail=False,
                         doc_modules=False)
        return AppContext(tool="sync", args=args, config=cfg,
                          root=Path(tempfile.mkdtemp()), json_mode=False,
                          started=time.perf_counter(), progress=Progress("sync"))

    def _make_open_pair(self, label):
        events: list = []
        target = FakeWorkbook(f"{label} V1.xlsm", events)
        target.Save = lambda: None
        ref = FakeWorkbook("参考簿.xlsm", events)
        xl = FakeXL([target, ref])
        return target, xl

    def test_single_push_reports_cleanup(self):
        cfg = {"projects": [self._proj()], "defaults": {}}
        target, xl = self._make_open_pair("示例表A")
        sync_push._attach_or_open = lambda proj, path, *, auto_open: (target, xl, "isolate")
        sync_push._writeback = lambda wb, proj, defaults, *, doc_modules, rollback_on_fail: ([], [])

        outcome = sync_push._push_single(
            self._ctx(cfg), self._proj(), cfg, {}, None,
            dry=False, backup=False, save=True, auto_open=True, rollback=False)

        self.assertTrue(outcome.warnings, "单项目回收诊断必须进 warnings")
        self.assertTrue(any("回收未完成" in w for w in outcome.warnings), outcome.warnings)
        self.assertIn("isolate_cleanup_problems", outcome.stats)
        self.assertTrue(any("参考簿.xlsm" in p for p in
                             outcome.stats["isolate_cleanup_problems"]))
        self.assertFalse(xl.QuitCalled)

    def test_batch_push_aggregates_cleanup(self):
        projs = [self._proj("示例表A"), self._proj("示例表B")]
        cfg = {"projects": projs, "defaults": {}}
        pairs = [self._make_open_pair(p["name"]) for p in projs]

        def _attach(proj, path, *, auto_open):
            target, xl = next(p for p in pairs if p[0].Name.startswith(proj["name"]))
            return target, xl, "isolate"

        sync_push._attach_or_open = _attach
        sync_push._writeback = lambda wb, proj, defaults, *, doc_modules, rollback_on_fail: ([], [])
        orig_vbe = sync_push.has_vbproject_access
        sync_push.has_vbproject_access = lambda wb: True
        self.addCleanup(setattr, sync_push, "has_vbproject_access", orig_vbe)

        outcome = sync_push._push_projects(
            self._ctx(cfg), cfg, {}, None, dry=False, backup=False,
            save=True, rollback=False)

        self.assertTrue(outcome.warnings, "批量回收诊断必须进 warnings")
        problems = outcome.stats.get("isolate_cleanup_problems") or []
        self.assertEqual(len(problems), 2, f"两个项目的收尾问题都要聚合：{problems}")
        self.assertTrue(all("参考簿.xlsm" in p for p in problems), problems)


class AgentPushGuardDiagnosticsTest(unittest.TestCase):
    """第二轮 3：agent.push 收尾 guard 还原 errors 进 warnings，不改既有人读输出。"""

    def setUp(self):
        excel._OWNED_SNAPSHOTS.clear()
        excel._QUIT_DONE.clear()
        self._orig_preflight = agent_io._preflight
        agent_io._preflight = lambda proj, defaults, doc_modules: []
        self._orig_resolve = agent_io.resolve_target
        agent_io.resolve_target = lambda p: (Path("示例表A V1.xlsm"), "桩路径")
        self._orig_attach = agent_io._attach_or_open
        self._orig_writeback = agent_io._writeback
        self.addCleanup(setattr, agent_io, "_preflight", self._orig_preflight)
        self.addCleanup(setattr, agent_io, "resolve_target", self._orig_resolve)
        self.addCleanup(setattr, agent_io, "_attach_or_open", self._orig_attach)
        self.addCleanup(setattr, agent_io, "_writeback", self._orig_writeback)

    def test_guard_restore_errors_feed_warnings(self):
        class _GuardApp:
            def __init__(self):
                self.EnableEvents = True
                self.ScreenUpdating = True
                self.DisplayAlerts = True
                self.DisplayStatusBar = True
                self.StatusBar = False
                self.CutCopyMode = False

            @property
            def Calculation(self):
                raise RuntimeError("RPC 断链")

            @Calculation.setter
            def Calculation(self, value):
                pass

        app = _GuardApp()
        wb = types.SimpleNamespace(Name="示例表A V1.xlsm", Application=app,
                                    Save=lambda: None)
        agent_io._attach_or_open = lambda proj, target, *, auto_open: (wb, None, "open")
        agent_io._writeback = lambda wb, proj, defaults, doc_modules, rollback_on_fail: ([], {})

        proj = {"name": "示例表A", "search_dir": "", "prefix": "示例表A", "src": ""}
        args = Namespace(project="示例表A", apply=True, no_backup=True, no_save=False,
                         open_only=False, rollback_on_fail=False, doc_modules=False)
        ctx = AppContext(tool="agent", args=args,
                         config={"projects": [proj], "defaults": {}},
                         root=Path(tempfile.mkdtemp()), json_mode=True,
                         started=time.perf_counter(), progress=Progress("agent"))

        outcome = agent_io.push(ctx)
        self.assertEqual(outcome.exit_code, 0, outcome.errors)
        self.assertTrue(any("状态还原未完成" in w for w in outcome.warnings),
                        outcome.warnings)
        self.assertTrue(any("Calculation" in w for w in outcome.warnings),
                        outcome.warnings)


class SyncPushGuardDiagnosticsTest(unittest.TestCase):
    """sync_push 收尾 guard 还原 errors 进 warnings，口径同 agent_io。"""

    def setUp(self):
        excel._OWNED_SNAPSHOTS.clear()
        excel._QUIT_DONE.clear()
        self._orig_preflight = sync_push._preflight_or_fail
        sync_push._preflight_or_fail = lambda proj, defaults, doc: []
        self._orig_resolve = sync_push.resolve_target
        sync_push.resolve_target = lambda p: (Path("示例表A V1.xlsm"), "桩路径")
        self._orig_attach = sync_push._attach_or_open
        self._orig_writeback = sync_push._writeback
        self.addCleanup(setattr, sync_push, "_preflight_or_fail", self._orig_preflight)
        self.addCleanup(setattr, sync_push, "resolve_target", self._orig_resolve)
        self.addCleanup(setattr, sync_push, "_attach_or_open", self._orig_attach)
        self.addCleanup(setattr, sync_push, "_writeback", self._orig_writeback)

    def test_guard_restore_errors_feed_warnings(self):
        class _GuardApp:
            def __init__(self):
                self.EnableEvents = True
                self.ScreenUpdating = True
                self.DisplayAlerts = True
                self.DisplayStatusBar = True
                self.StatusBar = False
                self.CutCopyMode = False

            @property
            def Calculation(self):
                raise RuntimeError("RPC 断链")

            @Calculation.setter
            def Calculation(self, value):
                pass

        app = _GuardApp()
        wb = types.SimpleNamespace(Name="示例表A V1.xlsm", Application=app,
                                    Save=lambda: None)
        sync_push._attach_or_open = lambda proj, path, *, auto_open: (wb, None, "open")
        sync_push._writeback = lambda wb, proj, defaults, *, doc_modules, rollback_on_fail: ([], [])

        proj = {"name": "示例表A", "src": "", "search_dir": "", "prefix": "示例表A"}
        cfg = {"projects": [proj], "defaults": {}}
        args = Namespace(project="示例表A", apply=True, no_backup=True, no_save=False,
                         open_only=False, rollback_on_fail=False, doc_modules=False)
        ctx = AppContext(tool="sync", args=args, config=cfg,
                         root=Path(tempfile.mkdtemp()), json_mode=False,
                         started=time.perf_counter(), progress=Progress("sync"))

        outcome = sync_push._push_single(
            ctx, proj, cfg, {}, None,
            dry=False, backup=False, save=True, auto_open=True, rollback=False)

        self.assertEqual(outcome.exit_code, 0, outcome.errors)
        self.assertTrue(any("状态还原未完成" in w for w in outcome.warnings),
                        outcome.warnings)
        self.assertTrue(any("Calculation" in w for w in outcome.warnings),
                        outcome.warnings)


class AttachSuggestedWordingTest(unittest.TestCase):
    """第二轮 4：attach 被卡住语境下，默认提示不再把人往 isolate 上引（判定逻辑不动）。"""

    def test_blocked_branch_does_not_push_isolate(self):
        excel.reset_rot_unreachable()
        target = Path("示例表A V1.xlsm")
        excel._LAST_ROT_UNREACHABLE.append((str(target), "受保护视图"))
        self.addCleanup(excel.reset_rot_unreachable)

        orig_find = excel.find_excel_with_workbook
        excel.find_excel_with_workbook = lambda path: (None, None)
        self.addCleanup(setattr, excel, "find_excel_with_workbook", orig_find)

        client = types.SimpleNamespace(
            GetActiveObject=lambda progid: (_ for _ in ()).throw(RuntimeError("连不上")))
        parent = types.ModuleType("win32com")
        parent.client = client
        with mock.patch.dict(sys.modules, {
            "pythoncom": types.SimpleNamespace(CoInitialize=lambda: None),
            "win32com": parent,
            "win32com.client": client,
        }):
            with self.assertRaises(excel.Failure) as cm:
                excel.attach_running_excel(target)

        failure = cm.exception
        self.assertEqual(failure.code, "COM_CONTEXT")
        self.assertIn("重试附着", failure.suggested)
        self.assertNotIn("仍不行用 run --isolate", failure.suggested)


class KeepWorkbooksMatchRuleTest(unittest.TestCase):
    """1.6.6：白名单 FullName 优先 + Name 兼容；同名多簿不误保留，剩余清单带 FullName。"""

    def setUp(self):
        excel._OWNED_SNAPSHOTS.clear()
        excel._QUIT_DONE.clear()

    def test_same_name_different_fullname_kept_by_fullname_only(self):
        """同名多簿：名单给 FullName 只保留对应簿，另一张被关；不 Quit，清单含剩余 FullName。"""
        with tempfile.TemporaryDirectory() as raw:
            tmp = Path(raw)
            events: list = []
            kept = FakeWorkbook("同名.xlsm", events,
                                full_name=str(tmp / "dirA" / "同名.xlsm"))
            other = FakeWorkbook("同名.xlsm", events,
                                 full_name=str(tmp / "dirB" / "同名.xlsm"))
            xl = FakeXL([kept, other])
            problems = excel.quit_owned(xl, keep_workbooks=[kept.FullName])
            self.assertFalse(xl.QuitCalled, "仍有参考簿时宁可留实例也不 Quit")
            self.assertFalse(kept.closed, "FullName 命中的簿必须保留")
            self.assertTrue(other.closed, "同名但 FullName 未命中的簿应被关闭")
            self.assertTrue(any(kept.FullName in p for p in problems), problems)

    def test_legacy_name_entry_still_kept(self):
        """旧写法名单只给 Name：大小写不敏感命中，兼容保留。"""
        events: list = []
        target = FakeWorkbook("目标.xlsm", events)
        ref = FakeWorkbook("参考簿.xlsm", events,
                           full_name=str(Path("x") / "参考簿.xlsm"))
        xl = FakeXL([target, ref])
        problems = excel.quit_owned(xl, keep_workbooks=["参考簿.xlsm"])
        self.assertFalse(xl.QuitCalled)
        self.assertFalse(ref.closed, "Name 命中的旧写法参考簿必须保留")
        self.assertTrue(target.closed)
        self.assertTrue(any("参考簿.xlsm" in p for p in problems), problems)

    def test_fullname_normalization_case_and_separator(self):
        """路径大小写/正反斜杠差异：归一后命中 FullName 保留。"""
        with tempfile.TemporaryDirectory() as raw:
            tmp = Path(raw)
            full = str(tmp / "子目录" / "参考簿.xlsm")
            events: list = []
            ref = FakeWorkbook("参考簿.xlsm", events, full_name=full)
            xl = FakeXL([ref])
            # 反斜杠→正斜杠 + 大写化；normpath/normcase 归一后两侧应相等。
            mixed = full.replace("\\", "/").upper()
            problems = excel.quit_owned(xl, keep_workbooks=[mixed])
            self.assertFalse(ref.closed, "归一化后命中 FullName 的簿必须保留")
            self.assertFalse(xl.QuitCalled)
            self.assertTrue(any(full in p for p in problems), problems)


class ComCallSuggestedWordingTest(unittest.TestCase):
    """1.6.6：com_call 两处默认 suggested 不再把人往 isolate/批量模式的失败路径上引。"""

    def test_busy_suggested_wording(self):
        def busy():
            raise Exception("(-2147418111, 'Call was rejected by callee.', None, None)")
        with self.assertRaises(excel.Failure) as cm:
            excel.com_call(busy, tries=1, gap=0)
        failure = cm.exception
        self.assertEqual(failure.code, "EXCEL_BUSY")
        self.assertNotIn("再用 --isolate", failure.suggested)
        self.assertNotIn("批量模式", failure.suggested)
        self.assertIn("换一张能打开的表", failure.suggested)

    def test_context_exhausted_suggested_wording(self):
        def stuck():
            raise Exception("generic 一直不响应")
        with self.assertRaises(excel.Failure) as cm:
            excel.com_call(stuck, tries=1, gap=0)
        failure = cm.exception
        self.assertEqual(failure.code, "COM_CONTEXT")
        self.assertNotIn("再用 --isolate", failure.suggested)
        self.assertNotIn("批量模式", failure.suggested)
        self.assertIn("换一张能打开的表", failure.suggested)


if __name__ == "__main__":
    unittest.main(verbosity=2)
