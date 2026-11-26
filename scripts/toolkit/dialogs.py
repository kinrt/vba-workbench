# -*- coding: utf-8 -*-
"""弹窗 / 模态框识别族：窗口枚举 + 纯函数分类（1.7 B2 回补）。

分层口径
--------
  · 本模块顶层只标准库；``win32gui`` / ``win32process`` 一律在枚举函数内延迟导入，
    没装 pywin32 / 无桌面会话的机器 ``import toolkit`` 不崩；
  · 分类判定全部是**纯函数**：输入伪造的 :class:`WindowInfo` 列表即可单测，
    不碰真实窗口；真实 ``win32gui`` 枚举只在 :func:`enumerate_windows` 里。

分类口径（错误码映射由 ``features.agent_patch`` 超时分支完成，本模块不抛 Failure）：

  · ``vbe_compile_error``：``#32770`` 且标题 / 正文含编译错误关键字
    （编译错误 / compile error / #If / #Else / #ElseIf / #End If）；
  · ``vbe_dialog``：VBE 通用模态框（标题含 Visual Basic / VBE），且不是编译错误；
  · ``modal_dialog``：Excel 侧 MsgBox / InputBox / 运行时错误框等非 VBE 模态；
  · ``unknown``：归类不出。归属不明必须返回 unknown / ``None``，不瞎猜。

按目标 PID 过滤：枚举到的其它 Excel 进程 / 其它程序的弹窗不会误报。
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable

# 分类结果常量（与 agent_patch 超时分支的错误码映射一一对应）
KIND_COMPILE_ERROR = "vbe_compile_error"
KIND_VBE_DIALOG = "vbe_dialog"
KIND_MODAL_DIALOG = "modal_dialog"
KIND_UNKNOWN = "unknown"

# 普通对话框类名：Win32 模态框统一注册成 #32770。
_DIALOG_CLASS = "#32770"

# 编译错误关键字（匹配前会归一化小写并去掉空白，所以写法不带空格）。
_COMPILE_KEYWORDS = (
    "编译错误", "compileerror",
    "#elseif", "#else", "#endif", "#if",
)
_VBE_TITLE_MARKS = ("visual basic", "vbe")
_EXCEL_TITLE_MARK = "microsoft excel"

_WS_RE = re.compile(r"\s+")


@dataclass(frozen=True)
class WindowInfo:
    """一个顶层可见窗口的快照；纯数据，分类函数只吃它。"""

    hwnd: int
    pid: int
    title: str
    class_name: str
    text: str = ""          # 子窗口正文拼接（空格分隔），标题读不到时靠它认关键字


@dataclass(frozen=True)
class DialogMatch:
    """一次弹窗识别结论；``kind`` 取本模块 ``KIND_*`` 常量。"""

    kind: str
    hwnd: int
    pid: int
    title: str
    text: str


def _normalize(text: str) -> str:
    """小写并去掉所有空白：容忍「# Else」「#End If」这类空格 / 大小写差异。"""
    return _WS_RE.sub("", str(text or "")).lower()


def classify_window(win: WindowInfo) -> str:
    """把单个窗口归成 ``KIND_*``；非对话框类或认不出一律 ``KIND_UNKNOWN``。"""
    if win.class_name != _DIALOG_CLASS:
        return KIND_UNKNOWN

    blob = _normalize(f"{win.title}\n{win.text}")
    if any(keyword in blob for keyword in _COMPILE_KEYWORDS):
        return KIND_COMPILE_ERROR

    title_low = (win.title or "").lower()
    if any(mark in title_low for mark in _VBE_TITLE_MARKS):
        return KIND_VBE_DIALOG
    if _EXCEL_TITLE_MARK in title_low:
        return KIND_MODAL_DIALOG
    return KIND_UNKNOWN


def detect_dialog(windows: Iterable[WindowInfo],
                  target_pid: int | None = None) -> DialogMatch | None:
    """从窗口列表里挑最具体的弹窗；``target_pid`` > 0 时只认该进程所属窗口。

    优先级：编译错误 > VBE 通用模态 > Excel 模态；都不沾边返回 ``None``。
    纯函数：列表可伪造，不做任何真实枚举。
    """
    pid_filter: int | None = None
    if target_pid is not None:
        try:
            pid_filter = int(target_pid)
        except (TypeError, ValueError):
            pid_filter = None
    if pid_filter is not None and pid_filter <= 0:
        pid_filter = None

    rank = {KIND_COMPILE_ERROR: 0, KIND_VBE_DIALOG: 1, KIND_MODAL_DIALOG: 2}
    best: DialogMatch | None = None
    for win in windows:
        if pid_filter is not None and win.pid != pid_filter:
            continue
        kind = classify_window(win)
        if kind == KIND_UNKNOWN:
            continue
        match = DialogMatch(kind=kind, hwnd=win.hwnd, pid=win.pid,
                            title=win.title, text=win.text)
        if best is None or rank[kind] < rank[best.kind]:
            best = match
    return best


def enumerate_windows() -> list[WindowInfo]:
    """枚举当前桌面可见顶层窗口；win32gui 不可用 / 枚举异常时返回空列表，不崩。

    每个窗口附带子窗口正文（``EnumChildWindows`` 收集子控件标题），供关键字匹配。
    """
    try:
        import win32process
        import win32gui
    except ImportError:
        return []

    wins: list[WindowInfo] = []

    def probe(hwnd: int, _extra: object) -> bool:
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd) or ""
            class_name = win32gui.GetClassName(hwnd) or ""
            pid = int(win32process.GetWindowThreadProcessId(hwnd)[1])
            body: list[str] = []
            if hasattr(win32gui, "EnumChildWindows"):
                def child(chwnd: int, _cextra: object) -> bool:
                    try:
                        text = (win32gui.GetWindowText(chwnd) or "").strip()
                    except Exception:  # noqa: BLE001 - 单个子控件读不到不影响整窗
                        return True
                    if text:
                        body.append(text)
                    return True
                win32gui.EnumChildWindows(hwnd, child, None)
            wins.append(WindowInfo(hwnd=int(hwnd), pid=pid, title=title,
                                   class_name=class_name, text=" ".join(body)))
        except Exception:  # noqa: BLE001 - 枚举失败不能盖住原超时分类
            pass
        return True

    try:
        win32gui.EnumWindows(probe, None)
    except Exception:  # noqa: BLE001
        return []
    return wins


def detect_target_dialog(target_pid: int | None = None) -> DialogMatch | None:
    """真实枚举窗口并按目标 PID 分类；枚举不可用时返回 ``None``（交上层回退超时口径）。"""
    return detect_dialog(enumerate_windows(), target_pid)
