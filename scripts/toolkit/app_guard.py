# -*- coding: utf-8 -*-
"""
toolkit.app_guard —— Excel 应用状态"进出还原"
============================================================================================
为什么要有它：宏跑到一半出错、或自己忘了成对复位，会把 Application 留成半关闭状态
（EnableEvents=False / ScreenUpdating=False / Calculation=手动或自动），表现就是
「界面点哪都不响应」「增删数据一直正在计算」。脚本侧跑完宏后统一还原一次，
比指望每个宏自己收尾更可靠。

【怎么用（脚本里）】
    from toolkit import app_guard

    with app_guard.guard(app) as g:          # 进：记快照；出：还原
        app.Run("'" + wb.Name + "'!某宏")
    if g.changes:                            # 还原时实际动过哪些项
        print("已还原：", g.describe())

    # 只还原、不进快照（收尾用）
    fixed = app_guard.restore(app)           # -> {"EnableEvents": (False, True), ...}

【还原口径（默认 calculation="manual"）】
    EnableEvents   = True          必须开：否则界面/事件全哑
    ScreenUpdating = True          必须开：否则屏幕不刷新
    DisplayAlerts  = True          必须开：否则错误静默
    DisplayStatusBar = True        必须开：进度条要能显示
    Calculation    = 手动(-4135)   本项目口径：表格长期手动，避免增删数据全表重算
    StatusBar      = False         清掉残留提示
    CutCopyMode    = False         清掉残留的"复制中"虚线框

    calculation 可传 "manual" / "auto" / None（None = 保持原样不动）。

【注意】
    · 只还原"应用级"状态，不动工作簿内容，不保存工作簿。
    · 宏故意保持手动计算是正常的，本模块不会把它改回自动（除非 calculation="auto"）。
"""

from __future__ import annotations

XL_MANUAL = -4135
XL_AUTOMATIC = -4105

# 属性名 -> 还原目标值（Calculation 单独处理）
_BOOL_TARGETS = {
    "EnableEvents": True,
    "ScreenUpdating": True,
    "DisplayAlerts": True,
    "DisplayStatusBar": True,
}

_ATTR_CN = {
    "EnableEvents": "事件响应",
    "ScreenUpdating": "屏幕刷新",
    "DisplayAlerts": "警告提示",
    "DisplayStatusBar": "状态栏显示",
    "Calculation": "计算模式",
    "StatusBar": "状态栏文本",
    "CutCopyMode": "剪切状态",
}


_OFF_TEXT = ("", "false", "0", "none")


def _is_off(value) -> bool:
    """StatusBar / CutCopyMode 的“已关闭”归一：False/0/None/空串以及 COM 的 "FALSE"。"""
    if value is False or value == 0 or value is None:
        return True
    if isinstance(value, str):
        return value.strip().lower() in _OFF_TEXT
    return False


class AppGuard:
    """with 语句载体：__enter__ 记快照，__exit__ 还原并记录改动了什么。"""

    def __init__(self, app, *, calculation="manual"):
        self.app = app
        self.calculation = calculation
        self.snapshot: dict = {}
        self.changes: dict = {}
        self.errors: list[str] = []

    # -- with 协议 ------------------------------------------------------------
    def __enter__(self):
        self.snapshot = self._read()
        return self

    def __exit__(self, exc_type, exc, tb):
        self.restore()
        return False        # 不吞异常，让调用方按原错误处理

    # -- 内部 ----------------------------------------------------------------
    def _read(self) -> dict:
        snap = {}
        for attr in list(_BOOL_TARGETS) + ["Calculation"]:
            try:
                snap[attr] = getattr(self.app, attr)
            except Exception as e:      # 应用已退出/COM 断链
                self.errors.append(f"读取 {attr} 失败：{e}")
        return snap

    def restore(self) -> dict:
        """按口径还原，返回 {属性: (原值, 现值)}（只含真的变了的项）。"""
        changed = {}
        for attr, target in _BOOL_TARGETS.items():
            try:
                cur = getattr(self.app, attr)
                if cur != target:
                    setattr(self.app, attr, target)
                    changed[attr] = (cur, target)
            except Exception as e:
                self.errors.append(f"还原 {attr} 失败：{e}")

        if self.calculation is not None:
            target = XL_MANUAL if str(self.calculation).lower() in ("manual", "-4135") else XL_AUTOMATIC
            try:
                cur = self.app.Calculation
                if cur != target:
                    self.app.Calculation = target
                    changed["Calculation"] = (cur, target)
            except Exception as e:
                self.errors.append(f"还原 Calculation 失败：{e}")

        for attr, value in (("StatusBar", False), ("CutCopyMode", False)):
            try:
                cur = getattr(self.app, attr)
                if not _is_off(cur):
                    setattr(self.app, attr, value)
                    # COM 可能把 False 读回字符串 "FALSE"；按归一后的口径判断是否真的
                    # 落成“关闭”，避免把一次无效设置记成“已还原”。
                    readback = getattr(self.app, attr)
                    if _is_off(readback):
                        changed[attr] = (cur, value)
                    else:
                        # N4：赋值被 Excel 静默忽略（弹窗占用 / 受保护视图等），回读仍是
                        # 开启态——不能弱化成“无需还原”，记入 errors，与布尔目标失败口径对齐。
                        self._report_restore_failed(attr, readback)
            except Exception:
                pass        # StatusBar/CutCopyMode 读不到就跳过，不影响主流程

        self.changes = changed
        return changed

    def describe(self) -> str:
        """把改动拼成一行中文，供脚本打印；有还原失败时附在末尾一起体现。"""
        parts = []
        for attr, (old, new) in self.changes.items():
            name = _ATTR_CN.get(attr, attr)
            if attr == "Calculation":
                parts.append(f"{name} {_calc_text(old)}→{_calc_text(new)}")
            else:
                parts.append(f"{name} {old}→{new}")
        head = "，".join(parts) if parts else "无需还原"
        # 「还原未生效」条目附进行人文案；读取/赋值异常类错误仍走调用方的专用
        # 错误通道（macro.py 的「状态还原未完成」告警），不在此重复拼接。
        failed = [e for e in self.errors if "未生效" in e]
        if failed:
            head += "；" + "；".join(failed)
        return head

    def _report_restore_failed(self, attr, readback) -> None:
        """记录「还原未生效」；同一条文案去重，重复调用 restore 不刷屏。"""
        msg = f"还原 {attr} 未生效：回读仍为 {readback}"
        if msg not in self.errors:
            self.errors.append(msg)


def _calc_text(value) -> str:
    if value == XL_MANUAL:
        return "手动"
    if value == XL_AUTOMATIC:
        return "自动"
    return str(value)


def guard(app, *, calculation="manual") -> AppGuard:
    """便捷构造：`with app_guard.guard(app) as g: ...`"""
    return AppGuard(app, calculation=calculation)


def restore(app, *, calculation="manual") -> dict:
    """不记快照、只按口径还原一次（出错收尾用），返回改动明细。"""
    return AppGuard(app, calculation=calculation).restore()


def calc_mode(app) -> str:
    """当前计算模式的中文名（诊断用）。"""
    try:
        return _calc_text(app.Calculation)
    except Exception as e:
        return f"读不到（{e}）"
