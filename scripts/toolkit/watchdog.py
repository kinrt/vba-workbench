# -*- coding: utf-8 -*-
"""线程看门狗与进程存活探测公共原语（1.7 B1 / B4 回补）。

本模块只放**与 Excel/COM 无关**的并发与进程探测骨架：

  · ``WatchdogTimeout``：看门狗硬超时异常类型；
  · ``run_with_watchdog(fn, timeout, ...)``：独立守护线程执行 ``fn``，超时不再等待，
    正常返回值 / 线程内异常原样回传；时钟、休眠均可注入，单测不需要真线程真等待；
  · ``pid_alive(pid, ...)``：按 PID 判进程是否还在，三态返回（存活 / 已退出 / 未知），
    win32 不可用或权限不足时安全降级为“未知”，不崩、不把权限问题误报成崩溃；
  · ``pid_from_hwnd(hwnd)``：窗口句柄反查所属 PID，读不到返回 0。

与断言超时口径的映射关系
------------------------
``MacroRunner.run`` 跑完后把执行结论记在 ``MacroResult.timed_out`` 上；
``features.agent_patch.handle_assert`` 读到 ``timed_out=True`` 后走本模块的探测原语
做二次分类，再落到登记错误码：

  1. 目标 PID 已退出 → ``EXCEL_CLOSED``（``stage=assert``，优先于超时码）；
  2. 仍存活（或存活探测未知）→ 交 ``toolkit.dialogs`` 认弹窗：编译错误框
     ``VBE_COMPILE_ERROR``（``stage=vbe``），其余模态框仍归 ``WATCHDOG_TIMEOUT``；
  3. 判不出实例状态也认不出弹窗 → 回退通用 ``WATCHDOG_TIMEOUT``（``stage=assert``）。

``run_with_watchdog`` 本身不抛业务失败码：它只负责“不再等待”，分类由上层编排层完成。
COM 对象不能跨线程直接复用，若用本原语包 COM 调用，``fn`` 内部要重新 attach 会话。
"""
from __future__ import annotations

import threading
import time
from typing import Any, Callable

# PROCESS_QUERY_LIMITED_INFORMATION：只读查询权限，足够判存活，不需要进程全部访问权。
_PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
# STILL_ACTIVE：GetExitCodeProcess 返回 259 表示进程尚未退出。
_STILL_ACTIVE = 259
# 打开不存在的进程句柄时 winerror 87（ERROR_INVALID_PARAMETER）/ 6（ERROR_INVALID_HANDLE）。
_DEAD_WINERRORS = frozenset({6, 87})


class WatchdogTimeout(Exception):
    """运行期硬超时：调用线程不再等待慢调用，交给上层编排层分类成超时 / 关闭 / 弹窗。"""

    def __init__(self, message: str, *, timeout: float | None = None) -> None:
        super().__init__(message)
        self.timeout = timeout


def run_with_watchdog(fn: Callable[[], Any], timeout: float, *,
                      clock: Callable[[], float] = time.monotonic,
                      sleep: Callable[[float], None] = time.sleep,
                      poll: float = 0.05) -> Any:
    """在守护线程里执行 ``fn``；超时立即抛 ``WatchdogTimeout``，不阻塞主流程。

    参数：
      · ``fn``：无参可调用；返回值或线程内异常原样回传调用线程；
      · ``timeout``：硬超时秒数；``<= 0`` 表示不启用看门狗，直接同步执行；
      · ``clock``：可注入时钟（默认 ``time.monotonic``），单测用假时钟避免真等待；
      · ``sleep``：可注入休眠（默认 ``time.sleep``），单测可记录 / 跳过；
      · ``poll``：轮询间隔秒数；假时钟下仅作为休眠上界使用。

    守护线程只负责“不再等待”：超时后不会 join 慢调用，也不会强行结束用户已打开的
    Excel；慢调用线程是 daemon，随主进程退出回收。
    """
    if timeout is None or timeout <= 0:
        return fn()

    box: dict[str, Any] = {}
    done = threading.Event()

    def target() -> None:
        try:
            box["value"] = fn()
        except BaseException as exc:  # noqa: BLE001 - 原样交回调用线程，不在线程内收口
            box["error"] = exc
        finally:
            done.set()

    thread = threading.Thread(target=target, name="vba-watchdog", daemon=True)
    thread.start()

    deadline = clock() + timeout
    while not done.is_set():
        remaining = deadline - clock()
        if remaining <= 0:
            break
        sleep(min(poll, remaining))

    if not done.is_set():
        raise WatchdogTimeout(f"{timeout:g}s 内宏没有返回", timeout=timeout)
    if "error" in box:
        raise box["error"]  # type: ignore[misc]
    return box.get("value")


def pid_from_hwnd(hwnd: int) -> int:
    """窗口句柄 → 所属进程 PID；读不到（无 pywin32 / 句柄失效）返回 0，不抛。"""
    if not hwnd:
        return 0
    try:
        import win32process
        return int(win32process.GetWindowThreadProcessId(int(hwnd))[1])
    except Exception:  # noqa: BLE001 - 权限 / 句柄问题不能盖住原超时分类
        return 0


def pid_alive(pid: int | None, *,
              probe: Callable[[int], bool | None] | None = None) -> bool | None:
    """按 PID 判进程存活，三态返回：``True`` 存活 / ``False`` 已退出 / ``None`` 未知。

    · ``pid`` 为空或 <= 0：返回 ``None``（没有可探测的目标）；
    · ``probe`` 可注入：单测传桩函数直接回三态，不碰真实进程；
    · 默认走 win32 ``OpenProcess`` + ``GetExitCodeProcess``；pywin32 缺失、句柄被拒、
      查询异常一律返回 ``None``（未知）——权限不足不能断言进程已退，避免误报崩溃。
    """
    if pid is None:
        return None
    try:
        pid_int = int(pid)
    except (TypeError, ValueError):
        return None
    if pid_int <= 0:
        return None

    if probe is not None:
        result = probe(pid_int)
        return None if result is None else bool(result)

    return _win32_pid_alive(pid_int)


def _win32_pid_alive(pid: int) -> bool | None:
    """win32 真实探测；任何不可用 / 权限 / 异常都降级为未知，不向外抛。"""
    try:
        import pywintypes
        import win32api
        import win32process
    except ImportError:
        return None

    handle = None
    try:
        handle = win32api.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        code = win32process.GetExitCodeProcess(handle)
        return int(code) == _STILL_ACTIVE
    except pywintypes.error as exc:
        if getattr(exc, "winerror", None) in _DEAD_WINERRORS:
            return False
        return None
    except Exception:  # noqa: BLE001 - 句柄被拒等未知状态交上层按存活口径继续
        return None
    finally:
        if handle is not None:
            try:
                win32api.CloseHandle(handle)
            except Exception:  # noqa: BLE001 - 清理失败不覆盖探测结论
                pass
