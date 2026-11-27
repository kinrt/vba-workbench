# -*- coding: utf-8 -*-
"""sync diff：只读比较项目源码与工作簿 VBE 模块。

不写回、不保存、不增删模块；表没开时用隐藏实例只读打开，读自己的副本后关闭。
模块状态以源码为基准：add=源码有工程无，update=两边不同，remove=工程有源码无，
same=忽略大小写/空白后相同。过程级只列新增/删除/变化，不做字符级 diff。
"""
from __future__ import annotations

from pathlib import Path

from toolkit.base import entry_command, read_source_text
from toolkit.contract import ErrorCode, Failure, MutationMode, Outcome, Stage
from toolkit.excel import (ExcelSession, component_snapshot, find_open_workbook,
                           open_disk_workbook, vbe_state)
from toolkit.files import write_text_utf8
from toolkit.registry import ArgSpec, CommandSpec, ProjectMode
from toolkit.repo import module_index, module_name_of, resolve_target, source_files
from toolkit.vba_diff import compare_modules

from .common_subset import (find_project, resolve_common_subset, snapshot_module_code,
                            warn_unknown_modules)


_STATUS_LABEL = {"add": "＋ 新增", "update": "～ 修改", "remove": "－ 仅工程", "same": "＝ 一致"}


def _local_dir(root: Path, spec: str) -> Path:
    """测试/多环境用 ctx.root 解析项目相对目录；绝对路径原样保留。"""
    p = Path(spec)
    return p if p.is_absolute() else Path(root) / p


def _source_modules(proj: dict, defaults: dict, root: Path) -> tuple[dict[str, str], str]:
    """源码侧模块快照；公共库按项目 common_modules 子集取，项目源码覆盖同名。"""
    common_dir = _local_dir(root, str(defaults.get("common_dir", "src/common")))
    all_mods = module_index(common_dir)
    subset, reason, ignored = resolve_common_subset(proj, defaults)
    warn_unknown_modules(ignored)
    common_names = set(all_mods) if subset is None else set(subset)

    out: dict[str, str] = {}
    for f in source_files(common_dir):
        if module_name_of(f) in common_names:
            out[module_name_of(f)] = read_source_text(f)
    src_spec = str(proj.get("src") or "")
    if not src_spec:
        raise Failure("项目配置缺少 src，无法比较源码",
                      code=ErrorCode.PROJECT_NOT_FOUND, stage=Stage.RESOLVE,
                      hint="核对 mapping.json 的 src")
    for f in source_files(_local_dir(root, src_spec)):
        out[module_name_of(f)] = read_source_text(f)
    return out, reason


def _vbe_modules(wb) -> dict[str, str]:
    """工程侧模块正文；工程锁着/未放行时按 VBE_PROTECTED 失败，不静默空表。"""
    state, detail = vbe_state(wb)
    if state != "ok":
        raise Failure(f"VBA 工程读不了：{detail}",
                      code=ErrorCode.VBE_PROTECTED, stage=Stage.VBE,
                      suggested="在 Excel 里解开工程保护/勾选 AccessVBOM 后重跑")
    comps = component_snapshot(wb.VBProject)
    if comps is None:
        raise Failure("VBA 工程读不到组件清单",
                      code=ErrorCode.VBE_PROTECTED, stage=Stage.VBE)
    code = snapshot_module_code(wb.VBProject)
    return {name: code.get(name, "") for name in comps}


def _open_readonly(proj: dict, path: Path):
    session = ExcelSession(prefer_running=True)
    try:
        xl = session.app()
        wb, _opened = find_open_workbook(xl, path)
        owned_wb = False
        if wb is None:
            wb = open_disk_workbook(xl, path, read_only=True)
            owned_wb = True
        return session, wb, owned_wb
    except Exception:
        session.close()
        raise


def _write_md(rows: list[dict], summary: dict, path: str, project: str, reason: str) -> None:
    out = ["# 源码 vs 工程模块差异", "",
           f"- 项目：{project}", f"- 公共库范围：{reason}",
           f"- 汇总：新增 {summary['add']} / 修改 {summary['update']} / "
           f"仅工程 {summary['remove']} / 一致 {summary['same']}", "",
           "| 模块 | 状态 | 源码行 | 工程行 | 新增过程 | 删除过程 | 变化过程 |",
           "|---|---|---:|---:|---|---|---|"]
    out += [f"| {r['name']} | {r['status']} | {r['source_lines']} | {r['workbook_lines']} | "
            f"{', '.join(r['procedures']['added'])} | {', '.join(r['procedures']['removed'])} | "
            f"{', '.join(r['procedures']['changed'])} |" for r in rows]
    write_text_utf8(path, "\n".join(out) + "\n")


def diff(ctx) -> Outcome:
    args = ctx.args
    project_name = str(getattr(args, "project", "") or "")
    proj = find_project(ctx.config, project_name)
    path, why = resolve_target(proj)
    if path is None:
        raise Failure(f"「{project_name}」没解析到 xlsm：{why}",
                      code=ErrorCode.TARGET_NOT_FOUND, stage=Stage.RESOLVE,
                      hint="核对 mapping.json 的 search_dir / prefix / file")

    with ctx.progress.stage("diff"):
        source, reason = _source_modules(proj, ctx.config.get("defaults", {}), ctx.root)
        if not source:
            raise Failure(f"「{project_name}」源码目录里没有可比较模块",
                          code=ErrorCode.SOURCE_PREFLIGHT, stage=Stage.PARSE,
                          suggested=entry_command("sync.py", "pull", "--project", project_name, "--apply"))
        session = None
        wb = None
        owned_wb = False
        try:
            session, wb, owned_wb = _open_readonly(proj, Path(path))
            vbe = _vbe_modules(wb)
            rows, summary = compare_modules(source, vbe)
        finally:
            if owned_wb and wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
            if session is not None:
                session.close()

    show_all = bool(getattr(args, "all", False))
    shown = [r for r in rows if show_all or r["status"] != "same"]
    if getattr(args, "md", None):
        _write_md(rows, summary, str(args.md), project_name, reason)

    human = [f"源码 vs 工程模块差异：{project_name}",
             f"  公共库范围：{reason}",
             f"  模块 {summary['total']}：新增 {summary['add']} / 修改 {summary['update']} / "
             f"仅工程 {summary['remove']} / 一致 {summary['same']}"]
    for r in shown:
        label = _STATUS_LABEL[r["status"]]
        proc = []
        if r["procedures"]["added"]:
            proc.append("新增 " + "、".join(r["procedures"]["added"][:6]))
        if r["procedures"]["removed"]:
            proc.append("删除 " + "、".join(r["procedures"]["removed"][:6]))
        if r["procedures"]["changed"]:
            proc.append("变化 " + "、".join(r["procedures"]["changed"][:6]))
        suffix = f"  [{'; '.join(proc)}]" if proc else ""
        human.append(f"  {label}  {r['name']}  ({r['source_lines']} → {r['workbook_lines']} 行){suffix}")
    if not shown:
        human.append("  没有差异。")
    if summary["remove"]:
        human.append("  注：remove 是“工程有、源码无”，diff 不会删除工程模块；push 也不自动删。")

    return Outcome(
        action="sync.diff",
        ok=True,
        details=rows,
        stats={"project": project_name, "common_scope": reason,
               **summary, "shown": len(shown), "timings": ctx.stage_timings()},
        human=human,
        warnings=(["工程有源码无的模块未被自动删除；如需清理请人工确认"] if summary["remove"] else []),
        exit_code=0,
        elapsed_ms=ctx.elapsed_ms(),
    )


SPECS: tuple[CommandSpec, ...] = (
    CommandSpec(
        entry="sync",
        verb="diff",
        handler="features.sync_diff:diff",
        mutation=MutationMode.NONE,
        project=ProjectMode.REQUIRED,
        summary="只读比较源码与工程模块（新增/修改/仅工程/一致）",
        ref="references/sync.md#diff",
        exit_codes=(0, 1, 2),
        args=(
            ArgSpec("--all", kind="flag", help="连一致模块一起列出"),
            ArgSpec("--md", kind="path", help="另写一份 Markdown 报告"),
            ArgSpec("--out", kind="path", help="另写一份 JSON 到文件"),
        ),
        tests=("tests/test_sync_diff.py",),
    ),
)
