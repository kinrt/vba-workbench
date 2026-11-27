# -*- coding: utf-8 -*-
"""audit -i 交互菜单本体。

边界：
  · 菜单层只做选择与参数装配，不碰 Excel、不写文件；执行全交既有 audit handler；
  · 输入 reader / 输出 writer / 执行器 runner 可注入，便于纯函数单测；
  · 终端把关（非 TTY 早退）由入口壳 ensure_interactive 负责；
  · 所有视角只读，动作结束回菜单，不要求按回车确认返回。
"""
from __future__ import annotations

import time
import unicodedata
from argparse import Namespace
from pathlib import Path
from typing import Any, Callable

from toolkit.base import msg
from toolkit.commands import specs_for
from toolkit.context import AppContext, Progress
from toolkit.registry import build_parser, load_handler
from toolkit.repo import files_under, module_name_of
from toolkit.ui import ask_yes, group_files

VIEW_ITEMS: tuple[tuple[str, str], ...] = (
    ("盘指定模块", "先选范围（公共库 / 某个项目），再按分类选模块，看成员级明细"),
    ("公共库总览", "src/common 全部模块的引用数 / 体积 / 建议"),
    ("项目依赖面", "选一个工作簿，看它用了哪些公共模块"),
    ("全仓总览", "公共库 + 结构问题 + 整体瘦身意见"),
    ("结构体检", "旧副本 / 孤儿壳 / 缺壳 / 重名过程 / 文档模块对账 / 规则回潮"),
    ("指标体检", "H01–H19 指标复跑"),
    ("导出 Markdown 报告", "选视角与范围，写到指定 .md"),
    ("显示设置", "排除自测 / 列出全部明细，会话内记住"),
)

VIEW_BY_INDEX = ("members", "common", "project", "overview",
                 "structure", "indicators")
MENU_LINE = "─" * 52
MENU_HEAVY = "═" * 52


def normalize(line: str) -> str:
    """全角数字 / 标点归一，去首尾空白。"""
    return unicodedata.normalize("NFKC", str(line or "")).strip()


def section(writer: Callable[[str], None], title: str) -> None:
    writer("")
    writer(f"  【{title}】")
    writer(f"  {MENU_LINE}")


def pick(labels: list[str], reader: Callable[[str], str],
         writer: Callable[[str], None], prompt: str,
         exit_hint: str, title: str = "") -> int | None:
    """单层编号选择；0 返回 None，输错重问。"""
    if title:
        section(writer, title)
    else:
        writer("")
    for i, label in enumerate(labels, 1):
        writer(f"  {i}. {label}")
    writer("")
    writer(f"  {MENU_LINE}")
    writer(f"  0. {exit_hint}")
    writer("")
    while True:
        try:
            value = normalize(reader(prompt))
        except EOFError:
            return None
        if value == "0":
            return None
        if value.isascii() and value.isdigit() and 1 <= int(value) <= len(labels):
            return int(value) - 1
        writer(f"  没有编号 {value or '（空）'}，请重新输入（0 {exit_hint}）。")


class MenuSettings:
    """会话内显示设置：只影响 --exclude / --all，不改判定规则。"""

    def __init__(self) -> None:
        self.exclude_selftest = True
        self.show_all = False

    def extra_args(self) -> tuple[str, ...]:
        args: list[str] = []
        if self.exclude_selftest:
            args += ["--exclude", "测试"]
        if self.show_all:
            args.append("--all")
        return tuple(args)

    def rule_line(self) -> str:
        return (f"  判定规则：排除自测={'开' if self.exclude_selftest else '关'}"
                f" · 展开明细={'开' if self.show_all else '关'}")


def common_dir(config: dict[str, Any], root: Path) -> Path:
    return root / str(config.get("defaults", {}).get("common_dir", "src/common"))


def project_dir(project: dict[str, Any], root: Path) -> Path:
    if project.get("src"):
        return root / str(project["src"])
    return root / "src" / "projects" / str(project.get("name") or "")


def scope_files(scope: dict[str, Any], config: dict[str, Any], root: Path,
                settings: MenuSettings) -> list[Path]:
    base = common_dir(config, root) if scope.get("kind") == "common" \
        else project_dir(scope["project"], root)
    return files_under(base, ["测试"] if settings.exclude_selftest else None)


def choose_scope(config: dict[str, Any], root: Path,
                 reader: Callable[[str], str],
                 writer: Callable[[str], None]) -> dict[str, Any] | None:
    projects = list(config.get("projects") or [])
    labels = [f"公共库（{common_dir(config, root)}）"]
    labels += [f"{p.get('name') or '(未命名)'}（{project_dir(p, root)}）" for p in projects]
    idx = pick(labels, reader, writer, "选范围编号：", "返回菜单", title="选择范围")
    if idx is None:
        return None
    if idx == 0:
        return {"kind": "common", "name": "公共库"}
    return {"kind": "project", "name": labels[idx], "project": projects[idx - 1]}


def choose_project(config: dict[str, Any], reader: Callable[[str], str],
                   writer: Callable[[str], None]) -> dict[str, Any] | None:
    projects = list(config.get("projects") or [])
    if not projects:
        writer("  mapping.json 里没有项目。")
        return None
    labels = [f"{p.get('name') or '(未命名)'}" for p in projects]
    idx = pick(labels, reader, writer, "选项目编号：", "返回菜单", title="选择项目")
    return None if idx is None else projects[idx]


def choose_module(scope: dict[str, Any], config: dict[str, Any], root: Path,
                  settings: MenuSettings, reader: Callable[[str], str],
                  writer: Callable[[str], None]) -> Path | None:
    files = scope_files(scope, config, root, settings)
    if not files:
        writer(f"  范围里没有可盘源码：{scope.get('name')}")
        return None
    groups = group_files(files)
    labels = [f"{name}（{len(items)}）" for name, items in groups]
    idx = pick(labels, reader, writer, "选分类编号：", "返回范围", title="选择分类")
    if idx is None:
        return None
    files = groups[idx][1]
    labels = [f"{module_name_of(f)}（{f.name}）" if f.name != f"{module_name_of(f)}{f.suffix}"
              else module_name_of(f) for f in files]
    idx = pick(labels, reader, writer, "选模块编号：", "返回分类", title="选择模块")
    return None if idx is None else files[idx]


def report(outcome: Any, writer: Callable[[str], None]) -> None:
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


def default_runner(config: dict[str, Any], root: Path,
                   specs: tuple[Any, ...]) -> Callable[..., Any]:
    """按动词找登记 spec，装配 --project / --md / --exclude / --all。"""
    parser = build_parser(specs, prog="audit.py")
    spec_by_verb = {spec.verb: spec for spec in specs}

    def runner(verb: str, *, project: str | None = None,
               extra: tuple[str, ...] = ()) -> Any:
        argv = [verb]
        if project:
            argv += ["--project", project]
        argv += list(extra)
        args = parser.parse_args(argv)
        ctx = AppContext(
            tool="audit", args=args, config=config, root=root,
            json_mode=False, started=time.perf_counter(),
            progress=Progress(tool="audit", enabled=False))
        return load_handler(spec_by_verb[verb])(ctx)

    return runner


def run_view(verb: str, project: str | None, settings: MenuSettings,
             runner: Callable[..., Any], writer: Callable[[str], None],
             extra: tuple[str, ...] = ()) -> None:
    outcome = runner(verb, project=project, extra=settings.extra_args() + extra)
    report(outcome, writer)
    writer("")
    writer("—— 完成，返回菜单 ——")


def run_members(config: dict[str, Any], root: Path, settings: MenuSettings,
                reader: Callable[[str], str], writer: Callable[[str], None],
                runner: Callable[..., Any]) -> None:
    scope = choose_scope(config, root, reader, writer)
    if scope is None:
        return
    module = choose_module(scope, config, root, settings, reader, writer)
    if module is None:
        return
    run_view("members", None, settings, runner, writer, ("--module", module.as_posix()))


def run_project(config: dict[str, Any], root: Path, settings: MenuSettings,
                reader: Callable[[str], str], writer: Callable[[str], None],
                runner: Callable[..., Any]) -> None:
    project = choose_project(config, reader, writer)
    if project is not None:
        run_view("project", str(project.get("name") or ""), settings, runner, writer)


def ask_output_path(verb: str, root: Path, reader: Callable[[str], str],
                    writer: Callable[[str], None]) -> Path | None:
    default = root / f"audit-{verb}.md"
    while True:
        try:
            value = normalize(reader(f"输出路径（回车 {default.name}，0 取消）："))
        except EOFError:
            return None
        if value == "0":
            return None
        path = default if not value else Path(value).expanduser()
        if not path.is_absolute():
            path = root / path
        if path.exists() and not ask_yes(f"文件已存在，覆盖 {path.name}？", default_no=True):
            writer("  已取消，换一个路径或按 0 退出。")
            continue
        return path


def run_export(config: dict[str, Any], root: Path, settings: MenuSettings,
               reader: Callable[[str], str], writer: Callable[[str], None],
               runner: Callable[..., Any]) -> None:
    view_labels = ["盘指定模块 members", "公共库总览 common", "项目依赖面 project",
                   "全仓总览 overview", "结构体检 structure", "指标体检 indicators"]
    idx = pick(view_labels, reader, writer, "选导出视角编号：", "返回菜单", title="导出视角")
    if idx is None:
        return
    verb = VIEW_BY_INDEX[idx]
    project: str | None = None
    extra: tuple[str, ...] = ()
    if verb == "members":
        scope = choose_scope(config, root, reader, writer)
        if scope is None:
            return
        module = choose_module(scope, config, root, settings, reader, writer)
        if module is None:
            return
        extra = ("--module", module.as_posix())
    elif verb == "project":
        chosen = choose_project(config, reader, writer)
        if chosen is None:
            return
        project = str(chosen.get("name") or "")
    path = ask_output_path(verb, root, reader, writer)
    if path is None:
        return
    writer("")
    writer(f"导出到：{path}")
    run_view(verb, project, settings, runner, writer, extra + ("--md", path.as_posix()))


def run_settings(settings: MenuSettings, reader: Callable[[str], str],
                 writer: Callable[[str], None]) -> None:
    while True:
        labels = [f"排除自测（--exclude 测试）：{'开' if settings.exclude_selftest else '关'}",
                  f"列出全部明细（--all）：{'开' if settings.show_all else '关'}"]
        idx = pick(labels, reader, writer, "选要切换的编号：", "返回菜单", title="显示设置")
        if idx is None:
            return
        if idx == 0:
            settings.exclude_selftest = not settings.exclude_selftest
        else:
            settings.show_all = not settings.show_all
        writer("")
        writer(f"  已更新：排除自测={'开' if settings.exclude_selftest else '关'}"
               f" · 展开明细={'开' if settings.show_all else '关'}")


def run_menu(*, config: dict[str, Any], root: Any,
             reader: Callable[[str], str] | None = None,
             writer: Callable[[str], None] | None = None,
             runner: Callable[..., Any] | None = None,
             specs: tuple[Any, ...] | None = None) -> int:
    reader = reader or input
    writer = writer or msg
    specs = specs or specs_for("audit")
    runner = runner or default_runner(config, root, specs)
    settings = MenuSettings()

    writer("")
    writer(f"  {MENU_HEAVY}")
    writer("  vba-workbench audit 菜单")
    writer(f"  {MENU_HEAVY}")
    writer(settings.rule_line())

    labels = [f"{label}  —  {note}" for label, note in VIEW_ITEMS]
    while True:
        idx = pick(labels, reader, writer, "选视角编号：", "退出", title="审计视角")
        if idx is None:
            writer("已退出交互菜单。")
            return 0
        if idx == 0:
            run_members(config, root, settings, reader, writer, runner)
        elif idx == 1:
            run_view("common", None, settings, runner, writer)
        elif idx == 2:
            run_project(config, root, settings, reader, writer, runner)
        elif idx == 3:
            run_view("overview", None, settings, runner, writer)
        elif idx == 4:
            run_view("structure", None, settings, runner, writer)
        elif idx == 5:
            run_view("indicators", None, settings, runner, writer)
        elif idx == 6:
            run_export(config, root, settings, reader, writer, runner)
        else:
            run_settings(settings, reader, writer)


def interactive_entry() -> int:
    """入口壳调用：加载配置后跑菜单；终端把关已在壳层完成。"""
    args = Namespace(command=None, json=False, no_input=False)
    ctx = AppContext.build("audit", args)
    return run_menu(config=ctx.config, root=ctx.root)
