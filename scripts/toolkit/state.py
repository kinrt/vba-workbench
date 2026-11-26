# -*- coding: utf-8 -*-
"""Excel 目标状态 × 命令用途的唯一提示矩阵。

这里只描述“这个状态在这个用途下该报什么码/给什么建议”，不碰 COM。
WorkbookSession 负责把探测结果落到某个 TargetState，然后查本表。
"""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

from .base import entry_command


class TargetState(str, Enum):
    MISSING = "missing"
    CLOSED = "closed"
    OPEN_CLEAN = "open_clean"
    OPEN_DIRTY = "open_dirty"
    OPEN_PROTECTED = "open_protected"
    OPEN_BUSY = "open_busy"
    READONLY = "readonly"
    COM_UNAVAILABLE = "com_unavailable"


class Purpose(str, Enum):
    PULL = "pull"
    PULL_OPEN = "pull_open"
    PUSH = "push"
    RUN = "run"
    ASSERT = "assert"
    WATCH = "watch"
    BOOK_OPEN = "book_open"


@dataclass(frozen=True)
class StateRule:
    code: str = ""
    stage: str = ""
    hint: str = ""
    suggested: str = ""
    retryable: bool = False
    warning: str = ""


# 只登记“需要拦截/需要提示”的组合；未登记组合由调用方按命令自身默认处理。
STATE_RULES: dict[tuple[TargetState, Purpose], StateRule] = {
    (TargetState.MISSING, Purpose.PULL): StateRule(
        "TARGET_NOT_FOUND", "resolve", "目标工作簿不存在；检查 mapping 的目录与文件名规则。"),
    (TargetState.MISSING, Purpose.PUSH): StateRule(
        "TARGET_NOT_FOUND", "resolve", "目标工作簿不存在；先确认 --project 解析到哪张表。"),
    (TargetState.MISSING, Purpose.RUN): StateRule(
        "TARGET_NOT_FOUND", "resolve", "目标工作簿不存在；先 list 核对项目解析。"),
    (TargetState.MISSING, Purpose.ASSERT): StateRule(
        "TARGET_NOT_FOUND", "resolve", "目标工作簿不存在；先 list 核对项目解析。"),
    (TargetState.CLOSED, Purpose.RUN): StateRule(
        "TARGET_NOT_OPEN", "attach", "表没开。先打开目标工作簿，或用 book open。",
        entry_command("agent.py", "book", "--action", "open", "--project", "甲", "--apply")),
    (TargetState.CLOSED, Purpose.ASSERT): StateRule(
        "TARGET_NOT_OPEN", "attach", "表没开。先打开目标工作簿，或用 book open。",
        entry_command("agent.py", "book", "--action", "open", "--project", "甲", "--apply")),
    (TargetState.CLOSED, Purpose.WATCH): StateRule(
        "TARGET_NOT_OPEN", "attach", "watch 只能盯已打开的表；先用 book open 打开。",
        entry_command("agent.py", "book", "--action", "open", "--project", "甲", "--apply")),
    (TargetState.CLOSED, Purpose.PULL_OPEN): StateRule(
        "TARGET_NOT_OPEN", "attach", "pull-open 需要已打开的表；先用 book open。",
        entry_command("agent.py", "book", "--action", "open", "--project", "甲", "--apply")),
    (TargetState.OPEN_DIRTY, Purpose.PULL): StateRule(
        warning="表已打开且有未保存改动；pull 会借当前已打开会话抓内存里的最新代码，不会替你保存工作簿。"),
    (TargetState.OPEN_DIRTY, Purpose.PULL_OPEN): StateRule(
        warning="表已打开且有未保存改动；pull-open 抓当前内存，不会替你保存。"),
    (TargetState.READONLY, Purpose.PUSH): StateRule(
        "FILE_READONLY", "open", "文件是只读的，无法保存。",
        'attrib -R "<工作簿绝对路径>"'),
    (TargetState.OPEN_PROTECTED, Purpose.PUSH): StateRule(
        "VBE_PROTECTED", "vbe", "VBA 工程被密码保护；先在 VBE 输入密码解锁。"),
    (TargetState.OPEN_PROTECTED, Purpose.RUN): StateRule(
        "VBE_PROTECTED", "vbe", "VBA 工程被密码保护；先在 VBE 输入密码解锁。"),
    (TargetState.OPEN_PROTECTED, Purpose.ASSERT): StateRule(
        "VBE_PROTECTED", "vbe", "VBA 工程被密码保护；先在 VBE 输入密码解锁。"),
    (TargetState.OPEN_BUSY, Purpose.PUSH): StateRule(
        "EXCEL_BUSY", "attach", "Excel 忙（弹窗/计算/模态框）；先处理弹窗再重试。", retryable=True),
    (TargetState.OPEN_BUSY, Purpose.RUN): StateRule(
        "EXCEL_BUSY", "attach", "Excel 忙（弹窗/计算/模态框）；先处理弹窗再重试。", retryable=True),
    (TargetState.COM_UNAVAILABLE, Purpose.PULL): StateRule(
        "COM_CONTEXT", "attach", "连不上 Excel COM；检查 pywin32/Excel 会话。"),
    (TargetState.COM_UNAVAILABLE, Purpose.PUSH): StateRule(
        "COM_CONTEXT", "attach", "连不上 Excel COM；检查 pywin32/Excel 会话。"),
}


def rule_for(state: TargetState, purpose: Purpose) -> StateRule | None:
    return STATE_RULES.get((state, purpose))
