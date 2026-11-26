# -*- coding: utf-8 -*-
"""看门狗 / PID 存活 / 弹窗识别纯函数单测：不启动 Excel、不真等待。

覆盖 1.7 B1（线程看门狗公共原语）、B2（弹窗分类纯函数族）、B4（PID 存活三态探测）：
  · ``run_with_watchdog``：超时不被慢调用阻塞、正常返回值 / 异常原样透传；
  · ``pid_alive``：三态返回与 probe 注入；
  · ``dialogs``：编译错误框 / VBE 通用模态 / Excel 模态 / 未知 四类判定与 PID 过滤、优先级。
"""
import sys
import threading
import time
import unittest
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))

from toolkit.dialogs import (KIND_COMPILE_ERROR, KIND_MODAL_DIALOG, KIND_UNKNOWN,
                             KIND_VBE_DIALOG, WindowInfo, classify_window, detect_dialog)
from toolkit.watchdog import WatchdogTimeout, pid_alive, run_with_watchdog


class _FakeClock:
    """假时钟：按预设序列依次取值，末值重复；单测不靠真实时间流逝。"""

    def __init__(self, values: list[float]):
        self._values = list(values)
        self._index = 0

    def __call__(self) -> float:
        value = self._values[min(self._index, len(self._values) - 1)]
        self._index += 1
        return value


class _FakeSleep:
    """假休眠：只记录休眠时长，不真睡。"""

    def __init__(self) -> None:
        self.calls: list[float] = []

    def __call__(self, seconds: float) -> None:
        self.calls.append(seconds)


class WatchdogTest(unittest.TestCase):
    def test_timeout_does_not_wait_for_slow_call(self):
        # 慢调用永远阻塞在事件上；假时钟第二次跳到截止时间之后，看门狗应立即放弃等待。
        clock = _FakeClock([0.0, 0.0, 100.0])
        sleep = _FakeSleep()
        blocker = threading.Event()
        started = time.perf_counter()
        with self.assertRaises(WatchdogTimeout):
            run_with_watchdog(blocker.wait, 0.2, clock=clock, sleep=sleep)
        self.assertLess(time.perf_counter() - started, 1.0, "超时后不应继续等待慢调用")
        self.assertTrue(sleep.calls, "超时前应有轮询休眠")

    def test_value_and_exception_pass_through(self):
        self.assertEqual(run_with_watchdog(lambda: 42, 1), 42)
        with self.assertRaises(ValueError):
            run_with_watchdog(lambda: (_ for _ in ()).throw(ValueError("boom")), 1)

    def test_timeout_disabled_runs_synchronously(self):
        # timeout <= 0 不启用看门狗：直接同步执行。
        self.assertEqual(run_with_watchdog(lambda: 7, 0), 7)


class PidAliveTest(unittest.TestCase):
    def test_empty_pid_is_unknown(self):
        self.assertIsNone(pid_alive(None))
        self.assertIsNone(pid_alive(0))
        self.assertIsNone(pid_alive(-3))

    def test_probe_injection_tri_state(self):
        self.assertTrue(pid_alive(123, probe=lambda pid: True))
        self.assertFalse(pid_alive(123, probe=lambda pid: False))
        self.assertIsNone(pid_alive(123, probe=lambda pid: None))

    def test_probe_receives_pid(self):
        seen: list[int] = []

        def probe(pid: int) -> bool:
            seen.append(pid)
            return True

        pid_alive(456, probe=probe)
        self.assertEqual(seen, [456])


def _dialog(hwnd: int, pid: int, title: str, text: str = "",
             class_name: str = "#32770") -> WindowInfo:
    return WindowInfo(hwnd=hwnd, pid=pid, title=title,
                      class_name=class_name, text=text)


class DialogClassifyTest(unittest.TestCase):
    def test_compile_error_dialog_recognized(self):
        win = _dialog(2, 123,
                      "编译错误：# ElseIf、# Else 或 # EndIf 的前面必须为 # If 子句")
        self.assertEqual(classify_window(win), KIND_COMPILE_ERROR)

    def test_compile_keyword_in_child_text_only(self):
        # 标题读不到时，子窗口正文里的编译关键字也要能认出来。
        win = _dialog(2, 123, "Microsoft Excel", text="# Else 没有对应的 # If")
        self.assertEqual(classify_window(win), KIND_COMPILE_ERROR)

    def test_generic_vbe_dialog_is_not_compile_error(self):
        win = _dialog(1, 123, "Microsoft Visual Basic for Applications",
                      text="工程被锁定，是否解锁？")
        self.assertEqual(classify_window(win), KIND_VBE_DIALOG)

    def test_excel_modal_dialog_recognized(self):
        win = _dialog(1, 123, "Microsoft Excel", text="测试弹窗")
        self.assertEqual(classify_window(win), KIND_MODAL_DIALOG)

    def test_non_dialog_class_is_unknown(self):
        win = _dialog(1, 123, "计算器", class_name="Calc", text="")
        self.assertEqual(classify_window(win), KIND_UNKNOWN)

    def test_detect_filters_by_pid(self):
        other = _dialog(2, 999, "编译错误：缺少 # End If")
        mine = _dialog(1, 123, "Microsoft Visual Basic for Applications",
                       text="工程被锁定，是否解锁？")
        # 别的进程的编译错误框不能误报；只认 PID 123 的 VBE 通用框。
        match = detect_dialog([other, mine], target_pid=123)
        self.assertIsNotNone(match)
        self.assertEqual(match.kind, KIND_VBE_DIALOG)
        self.assertEqual(match.pid, 123)
        self.assertIsNone(detect_dialog([other, mine], target_pid=456))

    def test_compile_error_outranks_vbe_dialog(self):
        vbe = _dialog(1, 123, "Microsoft Visual Basic for Applications", text="普通提示")
        compile_err = _dialog(2, 123, "编译错误：语法错误")
        match = detect_dialog([vbe, compile_err], target_pid=123)
        self.assertIsNotNone(match)
        self.assertEqual(match.kind, KIND_COMPILE_ERROR)
        self.assertEqual(match.title, "编译错误：语法错误")

    def test_no_recognized_dialog_returns_none(self):
        wins = [_dialog(1, 123, "计算器", class_name="Calc"),
                _dialog(2, 123, "设置", class_name="#32770", text="")]
        self.assertIsNone(detect_dialog(wins, target_pid=123))


if __name__ == "__main__":
    unittest.main()
