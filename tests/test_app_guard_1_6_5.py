# -*- coding: utf-8 -*-
"""1.6.5 定向回归：N4 还原未生效不再弱报 + 风险点 4 缺测补齐。

覆盖要点（对应 1.6.4 验收报告 §7.2 风险点 4）：
    1. CutCopyMode="FALSE" 字符串归类为 off，不记 changed；
    2. 设置后回读仍为 on（还原未生效）→ 进 errors，describe() 人读文本能体现；
    3. _is_off 边界参数化：None / 0 / False / "" / "false" / "0" / "none" /
       大写 "FALSE" / 非空文本 / 数字 1；
    4. StatusBar 自定义文本被还原后 changes 正确、describe() 文案正确。

约束：不启动真实 Excel；COM 边界全部用桩对象。
"""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

PKG = Path(__file__).resolve().parents[1] / "scripts"
if str(PKG) not in sys.path:
    sys.path.insert(0, str(PKG))

from toolkit import app_guard  # noqa: E402


class _StubApp:
    """桩 Application：ignore 里的属性赋值被静默忽略，模拟 Excel 吞赋值。"""

    def __init__(self, *, statusbar="FALSE", cutcopymode=False, ignore=()):
        self.EnableEvents = True
        self.ScreenUpdating = True
        self.DisplayAlerts = True
        self.DisplayStatusBar = True
        self.Calculation = app_guard.XL_AUTOMATIC
        self._statusbar = statusbar
        self._cutcopy = cutcopymode
        self._ignore = set(ignore)

    @property
    def StatusBar(self):
        return self._statusbar

    @StatusBar.setter
    def StatusBar(self, value):
        if "StatusBar" not in self._ignore:
            self._statusbar = value

    @property
    def CutCopyMode(self):
        return self._cutcopy

    @CutCopyMode.setter
    def CutCopyMode(self, value):
        if "CutCopyMode" not in self._ignore:
            self._cutcopy = value


class OffStringClassificationTest(unittest.TestCase):
    """要点 1：COM 假值字符串归类为 off，不记 changed、不进 errors。"""

    def test_cutcopymode_false_string_is_off(self):
        app = _StubApp(statusbar="FALSE", cutcopymode="FALSE")
        guard = app_guard.AppGuard(app, calculation=None)
        changes = guard.restore()
        self.assertEqual(changes, {})
        self.assertEqual(guard.errors, [])

    def test_mixed_off_values_do_not_interfere(self):
        # StatusBar 自定义文本能正常还原；CutCopyMode 假值不记改动，两者独立迭代。
        app = _StubApp(statusbar="正在汇总", cutcopymode="0")
        guard = app_guard.AppGuard(app, calculation=None)
        changes = guard.restore()
        self.assertEqual(list(changes), ["StatusBar"])
        self.assertEqual(changes["StatusBar"], ("正在汇总", False))
        self.assertEqual(guard.errors, [])


class RestoreNotEffectiveTest(unittest.TestCase):
    """要点 2：赋值被静默忽略、回读仍为 on → 进 errors，describe() 体现。"""

    def test_readback_still_on_reports_error(self):
        app = _StubApp(statusbar="正在汇总", cutcopymode=True,
                       ignore=("StatusBar", "CutCopyMode"))
        guard = app_guard.AppGuard(app, calculation=None)
        changes = guard.restore()

        # 未落成关闭：不记 change，也不能被弱化成“无需还原”。
        self.assertEqual(changes, {})
        self.assertEqual(len(guard.errors), 2)
        self.assertTrue(any("StatusBar" in e and "未生效" in e
                             for e in guard.errors), guard.errors)
        self.assertTrue(any("CutCopyMode" in e and "未生效" in e
                             for e in guard.errors), guard.errors)

        text = guard.describe()
        self.assertIn("未生效", text)
        self.assertIn("StatusBar", text)
        self.assertIn("CutCopyMode", text)

    def test_repeated_restore_does_not_spam_errors(self):
        app = _StubApp(statusbar="正在汇总", cutcopymode=True,
                       ignore=("StatusBar", "CutCopyMode"))
        guard = app_guard.AppGuard(app, calculation=None)
        guard.restore()
        guard.restore()
        self.assertEqual(len(guard.errors), 2, guard.errors)


class IsOffBoundaryTest(unittest.TestCase):
    """要点 3：_is_off 归一边界参数化。"""

    OFF_CASES = [None, 0, False, "", "false", "0", "none", "FALSE"]
    ON_CASES = ["正在汇总", 1]

    def test_off_cases(self):
        for value in self.OFF_CASES:
            with self.subTest(value=value):
                self.assertTrue(app_guard._is_off(value), repr(value))

    def test_on_cases(self):
        for value in self.ON_CASES:
            with self.subTest(value=value):
                self.assertFalse(app_guard._is_off(value), repr(value))


class StatusBarRestoreDescribeTest(unittest.TestCase):
    """要点 4：StatusBar 自定义文本还原后 changes 与 describe() 文案正确。"""

    def test_custom_statusbar_restored_and_described(self):
        app = _StubApp(statusbar="正在汇总", cutcopymode=False)
        guard = app_guard.AppGuard(app, calculation=None)
        changes = guard.restore()

        self.assertEqual(changes["StatusBar"], ("正在汇总", False))
        self.assertEqual(app.StatusBar, False)
        self.assertEqual(guard.errors, [])

        text = guard.describe()
        self.assertIn("状态栏文本", text)
        self.assertIn("正在汇总→False", text)


if __name__ == "__main__":
    unittest.main()
