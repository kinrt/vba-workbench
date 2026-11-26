# -*- coding: utf-8 -*-
"""
toolkit.excel —— Excel / VBE 的 COM 边界层（开表、连实例、认工程状态、重试）
============================================================
**只放"和 Excel 打交道"的机制，不放任何业务判断**（谁该写回、写哪些模块、
写前写后要核对什么，都是入口脚本的事）。

硬规矩
  · 本文件**顶层只允许标准库**：`win32com` / `pythoncom` 一律在函数里延迟导入。
    这样 audit.py / agent.py style 在不装 pywin32、没有 Excel 的机器上照样能 `import toolkit`。
  · COM 抛的原始错误对外没有可读性，一律翻译成人话再 `Failure`。

能力清单
  · `new_excel()` / `ExcelSession` / `attach_running_excel()`：自己开隐藏实例，或连已在跑的
  · `com_call()`：Excel 忙时会拒答（RPC_E_CALL_REJECTED），隔一会儿重试
  · `find_open_workbook()` / `require_open_workbook()`：按全路径找已打开的表，**绝不拿别的表顶替**
  · `open_disk_workbook()`：给批量抓取/写回用的隐藏实例打开磁盘文件
  · `vbe_state()` / `vbe_state_hint()` / `ensure_vbe_access()`：工程能不能读写、不能读写时怎么救
  · `find_component()` / `component_snapshot()` / `remove_module()` / `replace_module_code()` /
    `import_component()` / `backup_workbook()`：工程与整表的机械动作
    （`remove_module` 删完会回读确认名字释放，`import_component` 保证导入后的组件名不被补序号）
"""

from __future__ import annotations

import os
import sys
import time
from pathlib import Path

from .base import Failure, abs_path, entry_command, json_mode, msg, warn
from .files import backup_copy
from .repo import IMPORTABLE
from .ui import read_line
from .vba_parse import drop_attribute_lines


# ---------------------------------------------------------------------------
# 实例
# ---------------------------------------------------------------------------
# 自有实例归属快照：{id(xl): {"pid": int, "hwnd": int}}；只给 quit_owned 收尾前
# 复核用。CLI 单线程执行，不跨线程共享；读不到 PID/Hwnd 的字段记 0。
_OWNED_SNAPSHOTS: dict[int, dict[str, int]] = {}

# 已完成收尾的实例键集合：优先 (pid, hwnd)，读不到才退 id(xl)。这样同一 Excel
# 实例换了 Python 代理对象也不会重复 Quit；地址复用不会误伤新实例。
_QUIT_DONE: set[tuple] = set()


def _instance_snapshot(xl) -> dict[str, int]:
    """读实例 PID/Hwnd 快照；读不到记 0，不把权限问题误报成崩溃。"""
    pid = hwnd = 0
    try:
        hwnd = int(xl.Hwnd)
    except Exception:
        hwnd = 0
    if hwnd:
        try:
            import win32process
            pid = int(win32process.GetWindowThreadProcessId(hwnd)[1])
        except Exception:
            pid = 0
    return {"pid": pid, "hwnd": hwnd}


def _quit_key(xl, snapshot: dict[str, int] | None = None) -> tuple:
    """收尾幂等键：能读到 PID/Hwnd 就用稳定键，否则退回 Python 对象地址。"""
    snap = snapshot if snapshot is not None else _instance_snapshot(xl)
    pid = int(snap.get("pid", 0))
    hwnd = int(snap.get("hwnd", 0))
    if pid or hwnd:
        return ("instance", pid, hwnd)
    return ("pyid", id(xl))


def new_excel(hidden: bool = True, *, force_new: bool = False):
    """起一个 Excel 实例（hidden=True 时不可见），并关掉事件与宏提示。

    force_new=True 时绝不借用户已经开着的可见实例：`book open` 这类要改
    Visible/DisplayAlerts/EnableEvents/AutomationSecurity 的命令一旦借到用户会话，
    就会把用户手工会话的全局状态改掉，且后续没有地方负责还原。

    起实例后登记 PID/Hwnd 归属快照（N1）：quit_owned 收尾前复核，防止 PID 被复用、
    COM 代理半死时代理串到别的实例上误退出他人会话。
    """
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    if hidden or force_new:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = not hidden
        xl.DisplayAlerts = False
    else:
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = True
        xl.DisplayAlerts = False
        try:
            if int(xl.WindowState) == -4140:  # xlMinimized
                xl.WindowState = -4143        # xlNormal
        except Exception:
            pass

    # 禁止 Workbook_Open 等事件触发，避免打开时弹窗/自动跑宏
    try:
        xl.EnableEvents = False
    except Exception:
        pass
    try:
        xl.AutomationSecurity = 1  # msoAutomationSecurityLow
    except Exception:
        pass
    _OWNED_SNAPSHOTS[id(xl)] = _instance_snapshot(xl)
    return xl


def quit_owned(xl, *, save_changes: bool = False,
               keep_workbooks: list[str] | None = None) -> list[str]:
    """回收工具自己起的 Excel 实例：先逐簿关闭，再 Quit；返回问题清单，空=干净。

    为什么不能直接 ``Quit()``：宏写过工作表后工作簿是 dirty，AppGuard 又把
    DisplayAlerts 还原成 True，Quit 会弹出“是否保存”的隐藏对话框，进程留在后台并
    留下 ``~$`` 锁文件。这里先关簿（默认不保存）再退出，Quit 不再触发保存提示。

    归属复核（N1）：``new_excel`` 起实例时登记 PID/Hwnd 快照；收尾前比对当前实例，
    不一致则不碰簿、不 Quit，把「归属校验不通过」写进问题清单。无快照的旧调用路径
    （attach 借来的实例、桩对象）保持旧行为，不因缺快照报错。

    白名单（N2）：``keep_workbooks`` 列出必须保留的簿（宏在隔离实例里自己
    ``Workbooks.Open`` 的参考簿）；名单内的簿不关闭。匹配先按 ``FullName`` 归一
    比较（大小写不敏感 + 路径归一），命中不上再回退按 ``Name`` 大小写不敏感比较，
    避免不同目录下同名 xlsm 被一起误保留。关完后实例里仍有残留簿时 **不 Quit**，
    把剩余簿名（附 FullName）写进问题清单——宁可留着实例，也不吞掉别人数据。
    ``keep_workbooks=None`` 为 1.6.4 旧行为：关全部簿后退出。

    幂等：同一实例二次调用直接返回空清单，不重复 Quit、不抛异常。
    """
    problems: list[str] = []
    current = _instance_snapshot(xl)
    key = _quit_key(xl, current)
    if key in _QUIT_DONE:
        return problems

    # N1 归属令牌复核：有快照才比对；读不到 PID/Hwnd 的字段（值为 0）跳过该字段。
    snap = _OWNED_SNAPSHOTS.get(id(xl))
    if snap:
        mismatch: list[str] = []
        for field in ("pid", "hwnd"):
            old, now = snap.get(field, 0), current.get(field, 0)
            if old and now and old != now:
                mismatch.append(f"{field} 快照={old} 当前={now}")
        if mismatch:
            problems.append("归属校验不通过：" + "；".join(mismatch)
                            + "；未退出实例，避免误关他人会话")
            return problems

    try:
        xl.DisplayAlerts = False
    except Exception as e:  # noqa: BLE001 - 先尽量关提示，失败仍要继续关簿
        problems.append(f"关闭 DisplayAlerts 失败：{e}")

    try:
        workbooks = list(xl.Workbooks)
    except Exception as e:  # noqa: BLE001 - 读不到集合时仍尝试 Quit
        workbooks = []
        problems.append(f"读取 Workbooks 失败：{e}")

    keep_entries = list(keep_workbooks) if keep_workbooks else None

    def _kept(wb) -> bool:
        """白名单命中：先按归一化 FullName，回退按 Name；两侧大小写均不敏感。"""
        if keep_entries is None:
            return False
        full = str(getattr(wb, "FullName", "") or "")
        name = str(getattr(wb, "Name", "") or "")
        full_norm = _norm_keep_path(full) if full else ""
        name_low = name.lower()
        for entry in keep_entries:
            entry_str = str(entry)
            if full_norm and _norm_keep_path(entry_str) == full_norm:
                return True
            if entry_str.lower() == name_low:
                return True
        return False

    for wb in workbooks:
        name = getattr(wb, "Name", "?")
        if _kept(wb):
            continue  # 白名单参考簿保留，不关
        try:
            wb.Close(SaveChanges=save_changes)
        except Exception as e:  # noqa: BLE001 - 单簿失败不阻断其余收尾
            problems.append(f"关闭工作簿 {name} 失败：{e}")

    # N2 残留复核：白名单模式下还有簿就不退出，避免无保存关掉他人参考簿。
    remaining: list[str] = []
    try:
        for w in xl.Workbooks:
            rname = str(getattr(w, "Name", "?"))
            rfull = str(getattr(w, "FullName", "") or "")
            remaining.append(f"{rname}（{rfull}）" if rfull and rfull != rname else rname)
    except Exception:
        remaining = []
    if keep_entries is not None and remaining:
        problems.append("实例内仍有未回收工作簿：" + "、".join(remaining)
                        + "；保留参考簿未退出实例，避免吞掉他人数据")
        return problems

    try:
        xl.Quit()
    except Exception as e:  # noqa: BLE001 - 回收失败要上报，不能静默
        problems.append(f"Quit 失败：{e}")

    try:
        if int(xl.Workbooks.Count) > 0:
            problems.append(f"Quit 后仍有 {int(xl.Workbooks.Count)} 个工作簿未关闭")
    except Exception:
        pass

    _OWNED_SNAPSHOTS.pop(id(xl), None)
    _QUIT_DONE.add(key)
    return problems


class ExcelSession:
    """
    按需创建的隐藏 Excel 实例：第一次 app() 时才启动，用完自动 Quit。
    抓取/批量写回共用，避免每处各写一遍 try/finally 与 xl = xl or new_excel(...)。

    prefer_running=True：已有 Excel 在跑就借它，没有才自己起隐藏实例。
    VBA 工程的解锁状态是跟着 Excel 进程走的，自己新起的实例一律 50289，
    所以「用户在 VBE 里解过密」的批量场景必须借现有实例；借来的实例用完不 Quit。
    """

    def __init__(self, prefer_running: bool = False) -> None:
        self.xl = None
        self.prefer_running = prefer_running
        self.owned = False

    def app(self):
        if self.xl is None:
            if self.prefer_running:
                try:
                    self.xl = attach_running_excel()
                    self.owned = False
                    return self.xl
                except Exception:
                    pass
            self.owned = True
            self.xl = new_excel(hidden=True)
        return self.xl

    def close(self) -> None:
        if self.xl is not None and self.owned:
            quit_owned(self.xl)
        self.xl = None

    def __enter__(self) -> "ExcelSession":
        return self

    def __exit__(self, *exc) -> bool:
        self.close()
        return False


# 最近一次 ROT 枚举里“看得到 moniker、但 BindToObject 失败”的工作簿。
# 只给 attach 诊断用；CLI 单线程执行，不跨线程共享。
_LAST_ROT_UNREACHABLE: list[tuple[str, str]] = []


def reset_rot_unreachable() -> None:
    """清空最近一次 ROT 枚举的「可见但连不上」记录（测试隔离 / 重入收口用）。

    选这个收口方式（N7）：保留模块级列表与 ``_rot_excel_apps`` 入口先 clear 的既有
    行为，COM_CONTEXT 分流语义不变；另补显式 reset 接口供测试与重入场景主动清场，
    不再靠“下一次枚举顺带清空”这一隐式约定。
    """
    _LAST_ROT_UNREACHABLE.clear()


def _norm_path(value) -> str:
    """路径比较口径：绝对化 + Windows 大小写归一；坏路径退回原文小写。"""
    try:
        return os.path.normcase(str(Path(value).resolve()))
    except Exception:
        return os.path.normcase(str(value))


def _norm_keep_path(value) -> str:
    """白名单 FullName 比较口径：路径归一（``normpath``）+ 大小写归一（``normcase``）；坏路径退回原文归一。"""
    try:
        return os.path.normcase(os.path.normpath(str(value)))
    except Exception:
        return os.path.normcase(str(value))


def _target_unreachable(path: Path | None) -> list[tuple[str, str]]:
    """从最近一次 ROT 枚举结果里，挑出目标工作簿的“看得到但连不上”记录。"""
    if path is None:
        return []
    want = _norm_path(path)
    return [(display, reason) for display, reason in _LAST_ROT_UNREACHABLE
            if _norm_path(display) == want]


def _rot_excel_apps() -> list:
    """枚举 ROT 里的 Excel 实例；读不到就返回空，调用方再退回 GetActiveObject。

    为什么不能只用 GetActiveObject：它只返回一个"当前注册"的实例。用户先手动开了
    Excel、脚本又起过一个隐藏实例，或者一台机器有多个 Excel 进程时，它可能拿到的
    不是目标表所在的那个，于是出现"表明明开着却连不上"。
    """
    import pythoncom
    import win32com.client

    apps: list = []
    seen: set[int] = set()
    _LAST_ROT_UNREACHABLE.clear()
    try:
        ctx = pythoncom.CreateBindCtx(0)
        rot = pythoncom.GetRunningObjectTable()
        enum = rot.EnumRunning()
        while True:
            monikers = enum.Next(1)
            if not monikers:
                break
            mon = monikers[0]
            try:
                display = mon.GetDisplayName(ctx, None)
            except Exception:
                continue
            low = display.lower()
            # 只认工作簿 moniker（完整文件路径）。Application moniker 拿不到
            # 具体工作簿，对“找目标表在哪个实例”没有用。
            if not low.endswith((".xlsm", ".xlsx", ".xlsb", ".xls")):
                continue
            try:
                raw = mon.BindToObject(ctx, None, pythoncom.IID_IDispatch)
                wb = win32com.client.Dispatch(raw)
                app = wb.Application
                hwnd = int(app.Hwnd)
            except Exception as e:
                warn(f"ROT 里有工作簿但连不上：{display}（{e}）")
                _LAST_ROT_UNREACHABLE.append((display, str(e)))
                continue
            if hwnd in seen:
                continue
            seen.add(hwnd)
            apps.append(app)
    except Exception:
        pass
    return apps


def find_excel_with_workbook(path: Path | None) -> tuple[object | None, object | None]:
    """在所有能枚举到的 Excel 实例里，找包含目标工作簿的那个。

    返回 (Application, Workbook)。找不到返回 (None, None)。
    path=None 时只返回第一个可枚举实例、Workbook=None。
    """
    want = Path(path).resolve() if path is not None else None
    for app in _rot_excel_apps():
        if want is None:
            return app, None
        try:
            wb, _opened = find_open_workbook(app, want)
        except Exception:
            continue
        if wb is not None:
            return app, wb
    return None, None


def _mark_app_ready(xl) -> None:
    """刚连上的实例先验证一次它愿意接话；不擅自改可见性/警告开关。

    可见性和 DisplayAlerts 是用户会话状态：写回流程自己用 AppGuard 临时关、结束还原；
    `plan` 这类只读命令更不能顺手把隐藏实例弹出来。
    """
    com_call(lambda: xl.Workbooks.Count)


def attach_running_excel(path: Path | None = None, *, target_hint: str = "",
                         target_suggested: str = ""):
    """连上已在运行的 Excel，并尽量连到"目标表所在的那个实例"。

    path=None：保持旧行为，只取当前注册的实例（菜单/清单这类"看看有哪些表"用）。
    path=目标表：先枚举 ROT 找包含该工作簿的实例；找不到再退回 GetActiveObject。
                 这样不会因为用户先手动开了一个 Excel，就误连到别的实例。
    """
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()

    blocked: list[tuple[str, str]] = []
    if path is not None:
        app, _wb = find_excel_with_workbook(path)
        if app is not None:
            try:
                _mark_app_ready(app)
            except Failure:
                raise
            except Exception as e:
                raise Failure(
                    f"找到了目标表所在的 Excel，但它不接受当前进程的 COM 调用：{e}",
                    code="COM_CONTEXT", stage="attach",
                    suggested=target_suggested or "先检查 Excel 是否卡在弹窗/受保护视图；先关掉卡住的实例后重试附着；该表仍不可用时，换一张能打开的表再走 isolate",
                ) from e
            return app
        # ROT 里看得到目标工作簿，但 BindToObject 失败：这不是“没开”，必须报
        # COM_CONTEXT。否则会被后面的 GetActiveObject 误导成 TARGET_NOT_OPEN。
        blocked = _target_unreachable(path)

    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
    except Exception as e:
        if blocked:
            detail = blocked[0]
            raise Failure(
                f"目标工作簿在 ROT 里，但当前进程连不上：{detail[0]}（{detail[1]}）",
                code="COM_CONTEXT", stage="attach",
                hint="Excel 可能卡在弹窗/受保护视图/权限不同的会话；先关掉该实例后重试附着，"
                     "隔离副本同样连不上这个卡住的会话",
                suggested=target_suggested or "先关掉卡住的 Excel 实例后重试附着；该表仍不可用时，换一张能打开的表再走 isolate",
            ) from e
        raise Failure(
            "没找到正在运行的 Excel。单表命令的出口按动词而定，见下面的提示。"
            "不要用 Start-Process 绕过工具开表。",
            code="NO_EXCEL", stage="attach",
            hint=target_hint or "先打开目标工作簿，或用 run --isolate；不要用 Start-Process 绕过工具开表",
            suggested=target_suggested or entry_command("sync.py", "list", "--json"),
        ) from e

    # 拿到对象不等于 Excel 愿意接话：刚连上就验证一次（忙的时候会拒答）
    try:
        _mark_app_ready(xl)
    except Failure:
        raise
    except Exception as e:
        raise Failure(
            f"找到了 Excel，但它不接受当前进程的 COM 调用：{e}",
            code="COM_CONTEXT", stage="attach",
            suggested=target_suggested or "先检查 Excel 是否在受保护视图/不同权限会话；先关掉卡住的实例后重试附着；该表仍不可用时，换一张能打开的表再走 isolate",
        ) from e

    if path is not None:
        try:
            wb, opened = find_open_workbook(xl, path)
        except Failure:
            raise
        except Exception as e:
            raise Failure(
                f"连上了 Excel，但读不到它的工作簿列表：{e}",
                code="COM_CONTEXT", stage="attach",
                suggested=target_suggested or "确认 Excel 没有卡在弹窗/受保护视图；先关掉卡住的实例后重试附着；该表仍不可用时，换一张能打开的表再走 isolate",
            ) from e
        if wb is None:
            if blocked:
                detail = blocked[0]
                raise Failure(
                    f"目标工作簿在 ROT 里，但当前进程连不上：{detail[0]}（{detail[1]}）",
                    code="COM_CONTEXT", stage="attach",
                    hint="Excel 可能卡在弹窗/受保护视图/权限不同的会话；先关掉该实例后重试附着，"
                         "隔离副本同样连不上这个卡住的会话",
                    suggested=target_suggested or "先关掉卡住的 Excel 实例后重试附着；该表仍不可用时，换一张能打开的表再走 isolate",
                )
            lines = [f"目标表没有打开：{path}"]
            if opened:
                lines.append("当前 Excel 打开的是：")
                lines.extend(f"  - {p}" for p in opened)
            raise Failure(
                "\n".join(lines),
                code="TARGET_NOT_OPEN", stage="attach",
                hint=target_hint,
                suggested=target_suggested or
                "表未打开：先打开目标工作簿，或用 book --action open --project <项目> --apply",
            )
    return xl


def classify_com_failure(error: Exception | str) -> str:
    """把「Excel 已退出」和「Excel 只是忙」分开；其余按上下文错误处理。

    COM 错误码别凭印象写：0x80010001 才是 RPC_E_CALL_REJECTED（忙），
    0x8000FFFF 是 E_UNEXPECTED（上下文异常），0x8001010E 是 RPC_E_WRONG_THREAD
    （跨线程，不是 Excel 退出）。这两类都不能再归到忙/退出，否则 CLI 会把
    真正的原因掩盖成 EXCEL_BUSY / EXCEL_CLOSED。
    """
    text = f"{error}".lower()
    if ("call was rejected by callee" in text or "被调用方拒绝" in f"{error}"
            or "-2147418111" in text or "80010001" in text):
        return "EXCEL_BUSY"
    if ("rpc server is unavailable" in text or "rpc 服务器不可用" in f"{error}"
            or "远程过程调用失败" in f"{error}" or "remote procedure call failed" in text
            or "-2147023174" in text or "800706ba" in text
            or "-2147023170" in text or "800706be" in text
            or "-2147417848" in text or "80010108" in text):
        return "EXCEL_CLOSED"
    return ""


def _com_failure_kind(error: Exception) -> str:
    """兼容旧内部调用名；分类逻辑只保留一份。"""
    return classify_com_failure(error)


def com_call(fn, tries: int = 5, gap: float = 0.4):
    """
    Excel 忙的时候会拒绝外部调用（表现为 RPC_E_CALL_REJECTED 或属性读不到，
    如 AttributeError: Excel.Application.Workbooks），隔一会儿重试几次再放弃。
    进程已经退出/RPC 断开时不再空转重试，并给出可分流、可处置的错误码。
    """
    last: Exception | None = None
    for _ in range(max(tries, 1)):
        try:
            return fn()
        except Exception as e:
            last = e
            if _com_failure_kind(e) == "EXCEL_CLOSED":
                break
            time.sleep(gap)
    kind = _com_failure_kind(last) if last is not None else ""
    if kind == "EXCEL_CLOSED":
        raise Failure(f"Excel 已退出或 RPC 连接已断开：{last}",
                      code="EXCEL_CLOSED", stage="attach",
                      suggested="重新打开目标工作簿后重跑；先查宏里有没有 Quit/Close 或 Excel 是否崩溃")
    if kind == "EXCEL_BUSY":
        raise Failure(f"Excel 一直拒绝 COM 调用（试了 {tries} 次）：{last}",
                      code="EXCEL_BUSY", stage="attach",
                      suggested="先关掉弹窗或等计算结束；该表仍不可用时，换一张能打开的表再走 isolate")
    raise Failure(f"Excel 一直不响应（试了 {tries} 次）：{last}",
                  code="COM_CONTEXT", stage="attach",
                  suggested="先检查 Excel 是否卡在弹窗/受保护视图；该表仍不可用时，换一张能打开的表再走 isolate")


# ---------------------------------------------------------------------------
# 工作簿
# ---------------------------------------------------------------------------
def find_open_workbook(xl, path: Path | None) -> tuple[object | None, list[str]]:
    """
    在已打开的 Excel 里按【全路径】找目标工作簿，不依赖"谁在前台"。
    返回 (工作簿或 None, 当前打开的全部全路径)，用于在不匹配时报清楚"现在开了哪些"。
    path 为 None 时才退回当前活动工作簿。
    """
    opened: list[str] = []
    want = Path(path).resolve() if path is not None else None
    for wb in com_call(lambda: list(xl.Workbooks)):
        try:
            full = str(wb.FullName)
        except Exception:
            continue
        opened.append(full)
        if want is not None:
            try:
                if Path(full).resolve() == want:
                    return wb, opened
            except Exception:
                continue
    if want is None:
        return xl.ActiveWorkbook, opened
    return None, opened


def require_open_workbook(xl, path: Path, project_name: str, purpose: str,
                          *, hint: str = "", suggested: str = ""):
    """
    取"配置里解析到的那一张表"的已打开工作簿；没打开就报错，
    绝不拿别的工作簿顶替（一个往错的表写代码，一个把错的表导进项目源码目录）。
    """
    wb, opened = find_open_workbook(xl, path)
    if wb is None:
        lines = [f"「{project_name}」的表格没有打开：{path}",
                 f"请先在 Excel 里打开它，再回来{purpose}。"]
        if opened:
            lines.append("当前 Excel 打开的是：")
            lines.extend(f"  - {p}" for p in opened)
        raise Failure("\n".join(lines), code="TARGET_NOT_OPEN", stage="attach",
                      hint=hint or "表没打开：push 默认走磁盘；pull-open/run/watch 需要先打开，或 run --isolate",
                      suggested=suggested)
    try:
        active = xl.ActiveWorkbook
        active_full = str(active.FullName) if active is not None else ""
    except Exception:
        active_full = ""
    if active_full and active_full == str(wb.FullName):
        if not json_mode():
            msg(f"目标工作簿：{wb.Name}（当前活动）")
    else:
        wb.Activate()
        if not json_mode():
            msg(f"目标工作簿：{wb.Name}（已从其它工作簿切过来）")
    return wb


def open_disk_workbook(xl, path: Path, *, read_only: bool):
    """
    批量抓取/批量写回共用：用隐藏实例打开磁盘上的表。

    前提（不满足就在打开这一步失败，而不是导到一半报错）：
      - 没有【打开密码】：本工具不传密码，Excel 会直接拒绝打开；
      - VBA 工程没加密：否则后面读 VBProject 会失败；
      - 写回时文件不能是"只读"属性：否则 wb.Save() 存不回去。
    """
    if not path.is_file():
        raise Failure(f"文件不存在：{path}", code="TARGET_NOT_FOUND", stage="open")

    if not read_only and (path.stat().st_mode & 0o200) == 0:
        raise Failure(
            f"文件是只读属性，写回后存不回去：{path.name}\n"
            f"  在资源管理器里右键 → 属性 → 取消「只读」，再重试",
            code="FILE_READONLY", stage="open",
            hint="只读属性会阻止 Save；取消只读后重跑 push/pull。若文件在网盘/共享盘，先确认没有他人占用。",
            suggested=f'attrib -R "{path}"')

    try:
        # 关键字参数：UpdateLinks=0 不更新链接、Notify=False 不弹任何、IgnoreReadOnlyRecommended=True
        # 注意 pywin32 的关键字名是 Filename（不是 FileName），大小写敏感
        return xl.Workbooks.Open(
            Filename=str(path),
            UpdateLinks=0,
            ReadOnly=read_only,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
            AddToMru=False,
        )
    except Exception as e:
        raise Failure(
            f"打开失败：{path.name}\n"
            f"  常见原因：① 表设置了【打开密码】（本工具不支持带密码的表，需先去密码）"
            f" ② VBA 工程被加密 ③ 正被 Excel / 网盘同步进程占用 ④ 文件来自网络处于受保护视图\n"
            f"  原始错误：{e}",
            code="WORKBOOK_OPEN_FAILED", stage="open",
            suggested="检查文件密码/占用/受保护视图；表已开着时先关闭或用已打开会话",
        ) from e


def backup_workbook(wb, backup_root: str | Path, *, tag: str = "") -> Path | None:
    """整表备份进 <backup_root>/<时间戳><tag>/；工作簿还没落过盘（Path 为空）就跳过。"""
    if not wb.Path:
        return None
    dest = backup_copy(Path(wb.FullName), abs_path(backup_root), tag=tag)
    if not json_mode():
        msg(f"备份：{dest}")
    return dest


# ---------------------------------------------------------------------------
# VBA 工程状态
# ---------------------------------------------------------------------------
def vbe_state(wb) -> tuple[str, str]:
    """
    VBA 工程到底能不能读写，返回 (状态, 说明)：
      ok        - 能列出模块
      protected - 工程还锁着：VBProject.Name 照样读得到，一碰 Components 就报
                  「该工程已被保护，不能执行操作」→ 没在 VBE 里输工程密码
      denied    - 信任中心没勾「信任对 VBA 工程对象模型的访问」（AccessVBOM）
    只看 VBProject.Name 是不够的，这正是"整表 46 个模块都报工程里没有"的来源。
    """
    try:
        _ = wb.VBProject.Name
    except Exception as e:
        return "denied", f"VBProject 读不到：{e}"
    last: Exception | None = None
    for _ in range(3):
        try:
            return "ok", f"{int(wb.VBProject.VBComponents.Count)} 个模块"
        except Exception as e:
            last = e
            text = str(e)
            if "已被保护" in text or "protected" in text.lower():
                return "protected", "VBA 工程还锁着（还没在 VBE 里输工程密码）"
            time.sleep(0.4)
    return "denied", f"VBProject 读不到：{last}"


def vbe_state_hint(state: str) -> str:
    """把 vbe_state 翻成"用户下一步该做什么"。"""
    if state == "protected":
        return ("在 VBE 里「工具 → VBAProject 属性 → 保护」输一次工程密码解开，"
                "再重跑自动写回")
    return ("到「文件 → 选项 → 信任中心 → 宏设置」勾上"
            "「信任对 VBA 工程对象模型的访问」，再重跑自动写回")


def has_vbproject_access(wb) -> bool:
    """能不能读写 wb.VBProject：信任中心 AccessVBOM + 工程未加密，缺一个就读不到。"""
    return vbe_state(wb)[0] == "ok"


def ensure_vbe_access(wb, title: str, *, interactive: bool | None = None, max_retry: int = 10) -> bool:
    """
    读写工程前的统一关卡：状态不对就把原因和下一步打出来，然后问「回车重试」。
    你在 Excel/VBE 里当场解锁、勾好信任中心，按回车就能接着跑，不必退出重来。
      interactive=None → 看 sys.stdin.isatty()：后台/管道/批量一律不等人
      max_retry：最多问几次；输入流 EOF 或一直没修好时不会无限打屏
    返回 True 可以继续；False 表示这次放弃这张表。
    """
    if interactive is None:
        interactive = sys.stdin.isatty()
    state, detail = vbe_state(wb)
    if state == "ok":
        return True
    for n in range(1, max_retry + 1):
        warn(f"「{title}」的 VBA 工程现在读不了：{detail}")
        msg(f"  {vbe_state_hint(state)}")
        if not interactive:
            msg("  非交互运行，没法等你处理：这次先不碰这张表")
            return False
        msg(f"  处理好后按回车重试（第 {n}/{max_retry} 次）；输入 q 放弃这次操作")
        line = read_line("  > ").strip().lower()
        if line in {"q", "quit", "退出"}:
            warn(f"已放弃：这次不碰「{title}」")
            return False
        state, detail = vbe_state(wb)
        if state == "ok":
            msg(f"  「{title}」工程已可读（{detail}），继续")
            return True
    warn(f"「{title}」试了 {max_retry} 次还是读不了，这次放弃")
    return False


# ---------------------------------------------------------------------------
# 工程内的机械动作
# ---------------------------------------------------------------------------
def find_component(vbproj, name: str):
    try:
        return vbproj.VBComponents.Item(name)
    except Exception:
        return None


def component_snapshot(vbproj) -> dict[str, int] | None:
    """{模块名: 类型} 快照；工程被锁/读不到时返回 None。"""
    try:
        return {c.Name: c.Type for c in vbproj.VBComponents}
    except Exception:
        return None


def remove_module(vbproj, name: str, retries: int = 3) -> bool:
    """
    删除可导入类型的模块（标准 / 类 / 窗体），**删完回读确认名字真的释放了**。

    为什么要回读：VBE 的 `Remove` 是"提交式"的，函数返回时那个名字未必已经从命名空间里
    消失；这时候紧接着 `Import` 同名牌，VB 会自动补序号，工程里就冒出 `表格工具1`
    （2026-11-18 一天内撞了两次，4 个模块一起变后缀名）。
    所以这里删完要 poll 一下，名字还在就短睡再确认；确认不掉就返回 False，
    调用方**必须放弃 Import**——否则一定再造一个后缀模块出来。
    """
    comp = find_component(vbproj, name)
    if comp is None or comp.Type not in IMPORTABLE:
        return False
    vbproj.VBComponents.Remove(comp)
    for i in range(retries):
        if find_component(vbproj, name) is None:
            return True
        time.sleep(0.05 * (i + 1))
    return False


def import_component(vbproj, path, expect_name: str, retries: int = 2):
    """
    Import 一个源码文件，并确保落下来的组件名**就是** expect_name；拿不到就返回 None。

    兜底三层：
      1. Import 抢到名字 → 直接用；
      2. 被自动补了序号（`名字1`）→ 试着把 `comp.Name` 改回 expect_name，改成功就用；
      3. 改名也不行（原名还占着）→ 删掉这份抢名的组件，等一下重试；
         重试用完仍失败就返回 None，由调用方报错收口。
    绝不把 `xxx1` 留在工程里。
    """
    for _ in range(retries + 1):
        comp = vbproj.VBComponents.Import(str(path))
        if comp.Name == expect_name:
            return comp
        try:
            comp.Name = expect_name
            if comp.Name == expect_name:
                return comp
        except Exception:
            pass
        try:
            vbproj.VBComponents.Remove(comp)
        except Exception:
            return None
        time.sleep(0.1)
    return None



def replace_module_code(comp, code: str) -> None:
    """
    覆盖模块正文（先清空再整体写入）。

    文档模块不能 Remove，只能这样改；调用方负责判断 comp 是不是文档模块、
    以及代码里的导出头是否已剥掉（vba_parse.strip_export_header）。
    正文里的过程级 Attribute 行在写入前统一去掉（写进 CodeModule 会编译不过）。
    """
    cm = comp.CodeModule
    code = drop_attribute_lines(code)
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)
    if code.strip():
        cm.AddFromString(code)
