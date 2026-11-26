# -*- coding: utf-8 -*-
"""
toolkit.ui —— 控制台排版与交互原语
============================================================
统一约定（README §6.2）：

  · 只认数字编号；0 恒等于退出当前层（返回上一层菜单）；回车只有在 pick_list_default
    给了 default_hint 时才算"选默认"，否则无效；全角数字 NFKC 归一
  · 隔行显示：清单/菜单顶部空一行，「0. 退出 / 0. 返回菜单」行前也空一行
  · 选择必须分级：先选范围，再按分类列文件；任何一层不一次性罗列几十个选项
  · 菜单一律走 run_menu；清单走 pick_list / pick_list_default / pick_grouped；
    边框一律 SEP（─*72）；**入口脚本不许自己 for + input + 重问**

排版：display_width / pad_display / rjust / grid / section。
交互：read_line / pause / choose / ask_yes / run_action。
"""

from __future__ import annotations

import unicodedata
from pathlib import Path

from .base import Failure, assume_yes, fail, interactive_ok, msg
from .repo import MODULE_GROUP_ORDER

SEP = "─" * 72


# ---------------------------------------------------------------------------
# 排版
# ---------------------------------------------------------------------------
def display_width(text: str) -> int:
    """终端显示宽度：中文/全角算 2 列，其余算 1 列。"""
    return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in text)


def pad_display(text: str, width: int) -> str:
    """按显示宽度补空格（中文算 2 列），让清单里的箭头对齐。"""
    return text + " " * max(0, width - display_width(text))


def rjust(text: str, width: int) -> str:
    """按显示宽度右对齐（数字列用）。"""
    return " " * max(0, width - display_width(text)) + text


def grid(headers, rows, right=()) -> None:
    """按显示宽度对齐的通用表格；right 里放需要右对齐的列号。"""
    if not rows:
        msg("  （无）")
        return
    n = len(headers)
    widths = [display_width(str(h)) for h in headers]
    for r in rows:
        for i in range(n):
            widths[i] = max(widths[i], display_width(str(r[i])))

    def cell(v, i):
        text = str(v)
        if i in right:
            return rjust(text, widths[i]) + "  "
        return pad_display(text, widths[i]) + "  "

    msg("  " + "".join(cell(headers[i], i) for i in range(n)).rstrip())
    msg("  " + "─" * (sum(widths) + 2 * n))
    for r in rows:
        msg("  " + "".join(cell(r[i], i) for i in range(n)).rstrip())


def section(title: str) -> None:
    """报告分节标题：换行 + 边框 + 标题 + 边框。"""
    msg(f"\n{SEP}\n{title}\n{SEP}")


# ---------------------------------------------------------------------------
# 输入原语
# ---------------------------------------------------------------------------
def read_line(prompt: str) -> str:
    """
    读一行输入并去掉首尾空白；没有终端时直接报错（绝不静默卡死）。
    读不到（管道关闭）时返回空串。
    中文输入法开着时数字常是【全角】的（１而不是 1），int() 会直接抛 ValueError 把
    整个菜单崩掉，所以这里统一做 NFKC 归一：全角数字/空格/标点先转成半角。
    """
    if not interactive_ok():
        raise Failure(f"需要键盘输入但当前没有可交互终端：{prompt.strip()}（请改用子命令或加 --no-input）")
    try:
        raw = input(prompt)
    except EOFError:
        return ""
    return unicodedata.normalize("NFKC", raw).strip()


def pause() -> None:
    try:
        input("\n按回车返回菜单…")
    except EOFError:
        pass


def choose(prompt: str, count: int, exit_hint: str = "0 返回菜单") -> int | None:
    """
    统一的编号输入：只认 1..count 的数字；按 0 退出当前层（返回 None）；
    回车视为无效输入；输错编号提示一句后重问。三层菜单共用这一段，避免各写一套。
    """
    while True:
        v = read_line(prompt)
        if v == "0":
            return None
        if not v:
            print(f"  请输入编号；回车无效，按 {exit_hint}。")
            continue
        if v.isascii() and v.isdigit() and 1 <= int(v) <= count:
            return int(v)
        print(f"  没有编号 {v}，请重新输入（{exit_hint}）。")


def ask_yes(prompt: str, default_no: bool = True) -> bool:
    """y/N 提问；回车取默认（默认不覆盖，安全优先）。非交互一律按"不确认"处理。"""
    if assume_yes():
        return True
    if not interactive_ok():
        msg(f"[非交互] {prompt} → 不确认（跳过）")
        return False
    v = read_line(prompt + " (y/N) ").lower()
    if not v:
        return not default_no
    return v in ("y", "yes", "是", "1")


def run_action(func, *args, **kwargs) -> None:
    """执行一个功能，异常不让菜单崩掉；执行完按回车回菜单。"""
    try:
        func(*args, **kwargs)
    except Failure as e:
        fail(str(e))
    except KeyboardInterrupt:
        print("\n已取消本次操作")
    except Exception as e:
        fail(f"执行失败：{e}")
    pause()


# ---------------------------------------------------------------------------
# 清单与菜单
# ---------------------------------------------------------------------------
def group_files(files: list[Path]) -> list[tuple[str, list[Path]]]:
    """按 VBE 分类分组（MODULE_GROUP_ORDER 顺序，未知分类排后面；组内按模块名排序）。"""
    groups: dict[str, list[Path]] = {}
    for f in files:
        groups.setdefault(f.parent.name, []).append(f)
    order = [g for g in MODULE_GROUP_ORDER if g in groups] + \
            sorted(g for g in groups if g not in MODULE_GROUP_ORDER)
    return [(g, sorted(groups[g], key=lambda p: p.stem)) for g in order]


def pick_list(labels: list[str], title: str = "", *, prompt: str = "输入编号：",
              exit_hint: str = "返回菜单") -> int | None:
    """
    单层编号清单（一行一项，备注由调用方写进 labels）；返回 0 基下标，0/退出返回 None。
    隔行显示：清单顶部空一行，退出行前也空一行；exit_hint 是不带编号的纯文字。
    """
    msg("")
    if title:
        msg(title)
    for i, label in enumerate(labels, 1):
        msg(f"  {i}. {label}")
    msg(f"\n  0. {exit_hint}")
    k = choose(prompt, len(labels), f"0 {exit_hint}")
    return None if k is None else k - 1


def pick_list_default(labels: list[str], title: str = "", *, prompt: str = "请选择",
                      default_hint: str | None = None,
                      exit_hint: str = "返回菜单") -> tuple[str, int | None]:
    """
    同 pick_list 的排版，但支持【回车 = 默认项】（清单里有"最常用那几项"时用，
    省得每次都要挑）。返回三元组的第一项表示用户干了什么：

      ("pick", 0 基下标)  填了编号
      ("default", None)   直接回车（只有给了 default_hint 才可能）
      ("exit", None)      按 0

    default_hint 就是回车代表什么，例如"回车 = 已打开的 2 张"；传 None 表示本层
    不接受回车（提示"请输入编号"）。默认体现在调用方：它自己决定回车选哪些。
    """
    msg("")
    if title:
        msg(title)
    for i, label in enumerate(labels, 1):
        msg(f"  {i}. {label}")
    msg(f"\n  0. {exit_hint}")
    tail = f"0 {exit_hint}" if default_hint is None else f"{default_hint} · 0 {exit_hint}"
    while True:
        v = read_line(f"{prompt}（{tail}）：")
        if v == "0":
            return "exit", None
        if not v:
            if default_hint is not None:
                return "default", None
            msg(f"  请输入编号；回车无效，按 0 {exit_hint}。")
            continue
        if v.isascii() and v.isdigit() and 1 <= int(v) <= len(labels):
            return "pick", int(v) - 1
        msg(f"  没有编号 {v}，请重新输入（0 {exit_hint}）。")


def pick_grouped(groups: list[tuple[str, list[str]]], title: str = "", *,
                 prompt: str = "输入编号：", exit_hint: str = "返回菜单") -> int | None:
    """
    按分类分组的清单（编号跨组连续），返回扁平序列里的下标；0/退出返回 None。
    groups = [(组标题行, [每项显示文本...]), ...]，调用方自己维护同顺序的扁平列表。
    隔行显示同 pick_list：顶部空一行、退出行前空一行；exit_hint 是不带编号的纯文字。
    """
    msg("")
    if title:
        msg(title)
    n = 0
    for header, labels in groups:
        if not labels:
            continue
        msg(f"  {header}")
        for label in labels:
            n += 1
            msg(f"  {n}. {label}")
    msg(f"\n  0. {exit_hint}")
    k = choose(prompt, n, f"0 {exit_hint}")
    return None if k is None else k - 1


def run_menu(title_lines, items, handler, *, prompt: str = "  请选择：",
             exit_hint: str = "0 退出",
             divider: str | None = None) -> int:
    """
    一层编号菜单的骨架：空行 + 边框 + 标题行 + 边框 + 编号项（含备注行）
    + 隔行 +「0. 退出」（可选分隔线）+ 只认数字的输入。
    items = [(显示行, [备注行...]), ...]；handler 收 1 基编号（按 0 退出时收到 0），
    返回 None 表示继续显示菜单，返回 int 作为程序退出码。
    """
    while True:
        msg("")
        msg(SEP)
        for line in title_lines:
            msg(line)
        msg(SEP)
        for i, (label, notes) in enumerate(items, 1):
            msg(f"  {i}. {label}")
            for note in notes:
                msg(note)
        msg("\n  0. 退出")
        if divider:
            msg(divider)
        c = choose(prompt, len(items), exit_hint)
        code = handler(0 if c is None else c)
        if code is not None:
            return code
