# -*- coding: utf-8 -*-
"""sync -i 交互菜单本体。

流程：列项目 → 选项目 / 全局工具 / 新建项目 → 选动作 → 执行 → 回菜单。

边界规矩：
  · 菜单层**不碰 COM、不写入文件**，只做选择与参数装配；执行全交既有 handler；
  · 输入 ``reader`` / 输出 ``writer`` / 执行器 ``runner`` 均可注入，便于纯函数单测；
  · 终端把关（非 TTY 早退）仍由入口壳 ``ensure_interactive`` 负责，本模块不重复判断；
  · 只读动作选中即运行；写动作先自动预览，预览通过后回车才加 ``--apply``；
  · 每次动作结束都回到当前菜单，不退出进程，也不要求按回车确认返回。
"""
from __future__ import annotations

import re
import time
import unicodedata
from argparse import Namespace
from pathlib import Path
from typing import Any, Callable

from toolkit.base import ROOT, Failure, msg
from toolkit.commands import specs_for
from toolkit.context import AppContext, Progress
from toolkit.contract import MutationMode
from toolkit.registry import (ProjectMode, build_parser, load_handler)

# 第二级动作表：(动词, 展示备注)。顺序对齐常用方向：抓取在前、写回在后、维护类收尾。
# 需要额外必填参数的动词（promote --module、run --macro）不进菜单，避免在菜单里
# 再拼一套参数询问；这类请直接用命令行子命令。
PROJECT_ACTIONS: tuple[tuple[str, str], ...] = (
    ("pull", "抓取源码 VBE→src"),
    ("pull-open", "抓取已打开工作簿的未保存改动"),
    ("plan", "推荐模式与下一步（只读）"),
    ("push", "写回源码 src→VBE"),
    ("align", "换公共库占位壳"),
    ("gbk-export", "导出 GBK 副本"),
    ("vbamod", "核对模块清单（只读）"),
    ("watch", "常驻自动写回"),
)

GLOBAL_ACTIONS: tuple[tuple[str, str], ...] = (
    ("list", "列出各项目解析到哪个工作簿"),
    ("version", "核对公共库版本与各项目锁定版本"),
    ("doctor", "环境体检"),
    ("shells", "补 / 清公共库占位壳"),
    ("clean-backup", "清理 .backup 旧批次"),
)

GLOBAL_LABEL = "全局工具"
NEW_PROJECT_LABEL = "＋ 新建项目（拖入 .xlsm）"
_WORKBOOK_SUFFIXES = (".xlsm", ".xlsb")
_VERSION_RE = re.compile(r"[Vv](\d+)$")
_MENU_LINE = "─" * 52
_MENU_HEAVY = "═" * 52


def _section(writer: Callable[[str], None], title: str) -> None:
    """菜单分区：空行 + 标题 + 分隔线，避免选项和输入提示贴在一起。"""
    writer("")
    writer(f"  【{title}】")
    writer(f"  {_MENU_LINE}")


def _project_label(proj: dict[str, Any]) -> str:
    name = proj.get("name") or "(未命名)"
    src = proj.get("src") or proj.get("file") or ""
    return f"{name}（{src}）" if src else name


def _normalize(line: str) -> str:
    """全角数字/标点归一，去首尾空白（对齐 ui.read_line 规则）。"""
    return unicodedata.normalize("NFKC", str(line or "")).strip()


def _pick(labels: list[str], reader: Callable[[str], str],
          writer: Callable[[str], None], prompt: str,
          exit_hint: str, title: str = "") -> int | None:
    """单层编号选择：返回 0 基下标；按 0 / 退出返回 None。输错重问，不崩。

    选项区与输入区之间留空行并画一条分隔线，``0`` 单独放在线下方。
    """
    if title:
        _section(writer, title)
    else:
        writer("")
    for i, label in enumerate(labels, 1):
        writer(f"  {i}. {label}")
    writer("")
    writer(f"  {_MENU_LINE}")
    writer(f"  0. {exit_hint}")
    writer("")
    while True:
        try:
            v = _normalize(reader(prompt))
        except EOFError:
            return None
        if v == "0":
            return None
        if v.isascii() and v.isdigit() and 1 <= int(v) <= len(labels):
            return int(v) - 1
        writer(f"  没有编号 {v or '（空）'}，请重新输入（0 {exit_hint}）。")


def _ask_seconds(reader: Callable[[str], str],
                 writer: Callable[[str], None]) -> int | None:
    """watch 监听秒数；回车默认 600，0 取消。"""
    while True:
        try:
            v = _normalize(reader("监听秒数（回车 600，0 取消）："))
        except EOFError:
            return None
        if v == "":
            return 600
        if v == "0":
            return None
        if v.isascii() and v.isdigit() and int(v) > 0:
            return int(v)
        writer("  请输入大于 0 的整数秒。")


def _wait_enter(reader: Callable[[str], str], writer: Callable[[str], None],
                action_text: str) -> bool:
    """两级确认：回车执行，0 返回；其它输入重问。"""
    writer("")
    writer(f"  {_MENU_LINE}")
    writer(f"  回车 = {action_text}")
    writer("  0    = 返回")
    writer("")
    while True:
        try:
            v = _normalize(reader("> "))
        except EOFError:
            return False
        if v == "":
            return True
        if v == "0":
            return False
        writer("  只支持回车或 0。")


def _strip_quotes(text: str) -> str:
    text = (text or "").strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in "\"'“”":
        return text[1:-1].strip()
    return text


def _open_file_dialog() -> Path | None:
    """Windows 文件选择框；tkinter 不可用时返回 None，让用户改拖入路径。"""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        try:
            root.withdraw()
            root.attributes("-topmost", True)
            picked = filedialog.askopenfilename(
                title="选择要登记的工作簿",
                filetypes=[("Excel 工作簿", "*.xlsm *.xlsb"), ("所有文件", "*.*")],
            )
        finally:
            root.destroy()
        return Path(picked) if picked else None
    except Exception:
        return None


def _choose_workbook(reader: Callable[[str], str],
                     writer: Callable[[str], None]) -> Path | None:
    """拖入路径或弹出文件框；0 取消。"""
    _section(writer, "新建项目")
    writer("  把 .xlsm / .xlsb 拖进本窗口后回车")
    writer("  直接回车 = 打开文件选择框")
    writer("")
    writer(f"  {_MENU_LINE}")
    writer("  0. 取消")
    writer("")
    while True:
        try:
            raw = reader("> ")
        except EOFError:
            return None
        v = _strip_quotes(str(raw or ""))
        if v == "0":
            return None
        if not v:
            picked = _open_file_dialog()
            if picked is None:
                writer("  没有选择文件，请重新拖入，或按 0 取消。")
                continue
            path = picked
        else:
            path = Path(v).expanduser()
            if not path.is_absolute():
                path = Path.cwd() / path
        path = path.resolve()
        if not path.is_file():
            writer(f"  文件不存在：{path}")
            continue
        if path.suffix.lower() not in _WORKBOOK_SUFFIXES:
            writer(f"  只支持 .xlsm / .xlsb：{path.name}")
            continue
        return path


def _split_project_name(path: Path) -> tuple[str, str, bool]:
    """从文件名推导项目名和前缀：带 ` V<数字>` 时去掉版本号，按目录取最高版本。"""
    stem = path.stem.strip()
    m = _VERSION_RE.search(stem)
    if m:
        base = stem[:m.start()].rstrip() or stem
        return base, base, True
    return stem, stem, False


def _same_path(left: Any, right: Any) -> bool:
    try:
        return Path(left).expanduser().resolve() == Path(right).expanduser().resolve()
    except Exception:
        return str(left or "").lower() == str(right or "").lower()


def _project_name_for(config: dict[str, Any], base: str, path: Path,
                      prefix: str) -> str:
    """同名同目录沿用旧项目；同名不同目录自动加 -2 / -3。"""
    projects = list(config.get("projects") or [])
    for proj in projects:
        if proj.get("file") and _same_path(proj["file"], path):
            return str(proj.get("name") or base)
        if (proj.get("name") == base and proj.get("prefix") == prefix
                and _same_path(proj.get("search_dir"), path.parent)):
            return base
    name = base
    used = {str(proj.get("name") or "") for proj in projects}
    n = 2
    while name in used:
        name = f"{base}-{n}"
        n += 1
    return name


def _remember_project(config: dict[str, Any], proj: dict[str, Any]) -> None:
    """把 init 返回的项目并回当前配置，后续 pull/push 才能立即解析到它。"""
    rows = config.setdefault("projects", [])
    name = str(proj.get("name") or "")
    for i, row in enumerate(rows):
        if str(row.get("name") or "") == name:
            rows[i] = {**row, **proj}
            return
    rows.append(proj)


def _new_project(reader: Callable[[str], str], writer: Callable[[str], None],
                 runner: Callable[..., Any], config: dict[str, Any]) -> dict[str, Any] | None:
    """拖入工作簿后自动登记、自动抓取一次源码，并直接进入动作菜单。"""
    path = _choose_workbook(reader, writer)
    if path is None:
        writer("已取消新建项目。")
        return None

    base, prefix, versioned = _split_project_name(path)
    name = _project_name_for(config, base, path, prefix)
    extra = ["--dir", path.parent.as_posix(), "--prefix", prefix, "--allow-outside"]
    if not versioned:
        extra += ["--file", path.as_posix()]

    writer("")
    writer(f"已识别：{name}")
    writer(f"  目录：{path.parent}")
    if versioned:
        writer(f"  版本规则：{prefix} V<数字>{path.suffix.lower()}（自动取最高版本）")
    else:
        writer("  登记方式：按这个文件精确登记（文件名没有 V<数字>）")

    outcome = runner("init", name, apply=True, extra=tuple(extra))
    _report(outcome, writer)
    if int(getattr(outcome, "exit_code", 1)) != 0:
        writer("")
        writer("—— 新建失败，返回项目列表 ——")
        return None

    details = list(getattr(outcome, "details", []) or [])
    proj = details[0] if details and isinstance(details[0], dict) else {
        "name": name,
        "search_dir": path.parent.as_posix(),
        "prefix": prefix,
        "src": f"src/projects/{name}",
    }
    _remember_project(config, proj)

    writer("")
    writer("正在抓取源码……")
    pull = runner("pull", name, apply=True)
    _report(pull, writer)
    if int(getattr(pull, "exit_code", 1)) != 0:
        writer("项目已登记；源码抓取失败，可稍后重试。")
    else:
        writer("—— 已创建项目并抓取源码 ——")
    return proj


def _is_scoped(spec_by_verb: dict[str, Any], verb: str) -> bool:
    spec = spec_by_verb.get(verb)
    return spec is not None and spec.project != ProjectMode.NONE


def _default_runner(config: dict[str, Any], root: Any,
                    specs: tuple[Any, ...]) -> Callable[..., Any]:
    """默认执行器：按动词找登记 spec，装配 --project / --apply / --seconds。

    不在此层 new COM / 开表；handler 内部自己附着 Excel。
    """
    parser = build_parser(specs, prog="sync.py")
    spec_by_verb = {spec.verb: spec for spec in specs}

    def runner(verb: str, project: str, *, apply: bool = False,
               seconds: int | None = None,
               extra: tuple[str, ...] = ()) -> Any:
        spec = spec_by_verb[verb]
        argv = [verb]
        if project and _is_scoped(spec_by_verb, verb):
            argv += ["--project", project]
        if apply:
            argv.append("--apply")
        if seconds is not None:
            argv += ["--seconds", str(seconds)]
        argv += list(extra)
        args = parser.parse_args(argv)
        ctx = AppContext(
            tool="sync", args=args, config=config, root=root,
            json_mode=False, started=time.perf_counter(),
            progress=Progress(tool="sync", enabled=False))
        return load_handler(spec)(ctx)

    return runner


def _report(outcome: Any, writer: Callable[[str], None]) -> None:
    """把 handler 返回的 Outcome 翻成人话；错误行带 hint / 建议命令。"""
    for line in getattr(outcome, "human", []) or []:
        writer(str(line))
    for err in getattr(outcome, "errors", []) or []:
        writer(f"[错误] {getattr(err, 'message', err)}")
        hint = getattr(err, "hint", "")
        if hint:
            writer(f"说明：{hint}")
        suggested = getattr(err, "suggested", "")
        if suggested:
            writer(f"建议：{suggested}")


def _run_action(verb: str, spec: Any, project: str,
                reader: Callable[[str], str], writer: Callable[[str], None],
                runner: Callable[..., Any]) -> None:
    """执行一个动作；结束后返回动作菜单。"""
    if spec.mutation == MutationMode.NONE:
        outcome = runner(verb, project, apply=False)
        _report(outcome, writer)
        writer("")
        writer("—— 只读完成，返回动作菜单 ——")
        return

    if spec.mutation == MutationMode.CONTINUOUS:
        seconds = _ask_seconds(reader, writer)
        if seconds is None:
            writer("已取消本次操作。")
            return
        writer("")
        writer(f"将监听：{project}")
        writer(f"时长：{seconds} 秒")
        writer("源码变化后立即写回；启动后不再询问")
        if not _wait_enter(reader, writer, "开始监听"):
            return
        outcome = runner(verb, project, apply=True, seconds=seconds)
        _report(outcome, writer)
        writer("")
        writer("—— 监听结束，返回动作菜单 ——")
        return

    writer("")
    writer("—— 预览，未修改任何文件 ——")
    preview = runner(verb, project, apply=False)
    _report(preview, writer)
    if int(getattr(preview, "exit_code", 1)) != 0:
        writer("")
        writer("—— 预览失败，返回动作菜单 ——")
        return

    writer("")
    writer("—— 预览结束，未修改任何文件 ——")
    if not _wait_enter(reader, writer, "执行"):
        return

    writer("")
    outcome = runner(verb, project, apply=True)
    _report(outcome, writer)
    writer("")
    if int(getattr(outcome, "exit_code", 1)) != 0:
        writer("—— 执行失败，返回动作菜单 ——")
    else:
        writer("—— 完成，返回动作菜单 ——")


def _run_project_menu(project: dict[str, Any], reader: Callable[[str], str],
                      writer: Callable[[str], None], runner: Callable[..., Any],
                      actions: list[tuple[str, str]],
                      spec_by_verb: dict[str, Any]) -> None:
    """当前项目的动作菜单；0 返回项目列表，动作结束后继续留在本层。"""
    name = str(project.get("name") or "")
    labels = [label for _, label in actions]
    while True:
        idx = _pick(labels, reader, writer, "选动作编号：", "返回项目列表",
                    title=f"项目：{name} · 选择动作")
        if idx is None:
            return
        verb, _note = actions[idx]
        _run_action(verb, spec_by_verb[verb], name, reader, writer, runner)


def run_menu(*, config: dict[str, Any], root: Any,
             reader: Callable[[str], str] | None = None,
             writer: Callable[[str], None] | None = None,
             runner: Callable[..., Any] | None = None,
             specs: tuple[Any, ...] | None = None) -> int:
    """菜单主循环。返回进程退出码；用户按 0 退出时恒为 0。

    ``reader`` 收一行输入（默认终端 ``input``）；``writer`` 出一行人话
    （默认 ``msg``）；``runner(verb, project, apply=..., seconds=..., extra=...)``
    交既有 handler。
    """
    reader = reader or input
    writer = writer or msg
    specs = specs or specs_for("sync")
    runner = runner or _default_runner(config, root, specs)
    spec_by_verb = {spec.verb: spec for spec in specs}

    writer("")
    writer(f"  {_MENU_HEAVY}")
    writer("  vba-workbench 启动菜单")
    writer(f"  {_MENU_HEAVY}")
    if not config.get("projects"):
        writer("")
        writer("  mapping.json 里没有项目，可以直接新建一个。")

    project_actions = list(PROJECT_ACTIONS)
    global_actions = list(GLOBAL_ACTIONS)

    while True:
        projects = list(config.get("projects") or [])
        project_labels = [_project_label(p) for p in projects]
        project_labels += [GLOBAL_LABEL, NEW_PROJECT_LABEL]

        idx = _pick(project_labels, reader, writer, "选项目编号：", "退出",
                    title="选择项目")
        if idx is None:
            writer("已退出交互菜单。")
            return 0

        if idx == len(projects):
            _run_project_menu({"name": GLOBAL_LABEL}, reader, writer, runner,
                              global_actions, spec_by_verb)
            continue
        if idx == len(projects) + 1:
            created = _new_project(reader, writer, runner, config)
            if created is not None:
                _run_project_menu(created, reader, writer, runner,
                                  project_actions, spec_by_verb)
            continue

        _run_project_menu(projects[idx], reader, writer, runner,
                          project_actions, spec_by_verb)


def interactive_entry() -> int:
    """入口壳调用：加载配置后跑菜单；终端把关已在壳层 ``ensure_interactive`` 完成。

    首次没有 mapping.json 时不直接失败：用空配置进菜单，让用户选“新建项目”。
    """
    args = Namespace(command=None, json=False, apply=False, no_input=False)
    try:
        ctx = AppContext.build("sync", args)
    except Failure as exc:
        text = str(exc)
        if "找不到配置文件" not in text and "没有任何 projects" not in text:
            raise
        return run_menu(config={}, root=ROOT)
    return run_menu(config=ctx.config, root=ctx.root)
