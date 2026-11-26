# -*- coding: utf-8 -*-
"""1.6.4 回归：隔离收尾、ROT 绑定失败、未开表兜底文案、状态栏假值。

约束：不启动 Excel；COM 边界全部用桩对象或 mock，测试件只落临时目录。
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

PKG = Path(__file__).resolve().parents[1] / "scripts"
if str(PKG) not in sys.path:
    sys.path.insert(0, str(PKG))

from toolkit import app_guard, excel  # noqa: E402
from toolkit.base import Failure  # noqa: E402


class QuitOwnedTest(unittest.TestCase):
    class FakeWorkbook:
        def __init__(self, events):
            self.Name = "示例表.xlsm"
            self.events = events

        def Close(self, SaveChanges=False):
            self.events.append(("close", SaveChanges))

    class FakeBooks:
        def __init__(self, workbooks):
            self._workbooks = list(workbooks)
            self.Count = len(self._workbooks)

        def __iter__(self):
            return iter(self._workbooks)

    class FakeXl:
        def __init__(self, events):
            self.DisplayAlerts = True
            self.QuitCalled = False
            self.Workbooks = QuitOwnedTest.FakeBooks(
                [QuitOwnedTest.FakeWorkbook(events)])

        def Quit(self):
            self.QuitCalled = True
            self.Workbooks.Count = 0
            if self.DisplayAlerts:
                raise RuntimeError("保存提示挡住了 Quit")

    def test_quit_owned_closes_dirty_workbook_before_quit(self):
        events = []
        xl = self.FakeXl(events)
        problems = excel.quit_owned(xl)
        self.assertEqual(problems, [])
        self.assertEqual(events, [("close", False)])
        self.assertTrue(xl.QuitCalled)
        self.assertFalse(xl.DisplayAlerts)

    def test_quit_owned_reports_quit_failure(self):
        events = []
        xl = self.FakeXl(events)
        xl.Quit = lambda: (_ for _ in ()).throw(RuntimeError("COM 拒绝退出"))
        problems = excel.quit_owned(xl)
        self.assertTrue(any("Quit 失败" in p for p in problems), problems)
        self.assertEqual(events, [("close", False)])


class RotUnreachableTest(unittest.TestCase):
    def test_rot_target_bind_failure_is_com_context(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "示例表 V1.xlsm"
            old = excel._LAST_ROT_UNREACHABLE
            excel._LAST_ROT_UNREACHABLE = [(str(target), "unknown application")]
            self.addCleanup(setattr, excel, "_LAST_ROT_UNREACHABLE", old)

            with mock.patch.object(excel, "find_excel_with_workbook",
                                   return_value=(None, None)), \
                    mock.patch("win32com.client.GetActiveObject",
                               side_effect=RuntimeError("no active")):
                with self.assertRaises(Failure) as cm:
                    excel.attach_running_excel(target)

        self.assertEqual(cm.exception.code, "COM_CONTEXT")
        self.assertIn("ROT", str(cm.exception))
        self.assertIn("示例表 V1.xlsm", str(cm.exception))

    def test_target_not_open_fallback_has_no_push_copy(self):
        class FakeApp:
            class Workbooks:
                Count = 0

            Workbooks = Workbooks()

        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "示例表 V1.xlsm"
            with mock.patch.object(excel, "find_excel_with_workbook",
                                   return_value=(None, None)), \
                    mock.patch.object(excel, "find_open_workbook",
                                      return_value=(None, [])), \
                    mock.patch("win32com.client.GetActiveObject",
                               return_value=FakeApp()):
                with self.assertRaises(Failure) as cm:
                    excel.attach_running_excel(target)

        self.assertEqual(cm.exception.code, "TARGET_NOT_OPEN")
        self.assertNotIn("push 默认", cm.exception.suggested)
        self.assertIn("book --action open", cm.exception.suggested)


class AppGuardStatusBarTest(unittest.TestCase):
    class FakeApp:
        EnableEvents = True
        ScreenUpdating = True
        DisplayAlerts = True
        DisplayStatusBar = True
        Calculation = app_guard.XL_AUTOMATIC
        StatusBar = "FALSE"
        CutCopyMode = False

    def test_string_false_is_not_a_change(self):
        changes = app_guard.AppGuard(self.FakeApp(), calculation=None).restore()
        self.assertEqual(changes, {})

    def test_real_statusbar_is_still_restored(self):
        app = self.FakeApp()
        app.StatusBar = "正在汇总"
        changes = app_guard.AppGuard(app, calculation=None).restore()
        self.assertEqual(changes["StatusBar"], ("正在汇总", False))
        self.assertEqual(app.StatusBar, False)


if __name__ == "__main__":
    unittest.main()
