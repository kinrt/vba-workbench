# -*- coding: utf-8 -*-
"""agent_io 切片：probe / book / backup / logs / push 入口薄层。

边界：
  · handler 只返回 Outcome，不 print、不自带 JSON；
  · COM 附着/开表只走 ``adapters.excel_gateway``；
  · 项目解析与状态一览装配在 ``agent_probe``，日志汇总在 ``agent_logs``；
  · push 写回链：源码预检 → 备份 → 写回 → 回读 → 失败回滚，全部走 toolkit 既有原语。
"""
from __future__ import annotations

import re
import tempfile
from pathlib import Path
from typing import Any

from adapters import excel_gateway
from toolkit import app_guard
from toolkit.base import (ROOT, abs_path, entry_command,
                          Failure as BaseFailure, read_source_text)
from toolkit.check import check_module_set
from toolkit.contract import (ErrorCode, Failure as ContractFailure,
                              MutationMode, Outcome, Stage)
from toolkit.excel import (backup_workbook, component_snapshot, ensure_vbe_access,
                           find_component, import_component, quit_owned, remove_module,
                           replace_module_code, vbe_state)
from toolkit.files import backup_copy, gbk_import_copy, remove_tree
from toolkit.repo import (TYPE_FOR_EXT, VBXT_CLASS, VBXT_DOCUMENT,
                         VBXT_STANDARD, module_name_of,
                         resolve_target, source_files)
from toolkit.registry import ArgSpec, CommandSpec, ProjectMode
from toolkit.vba_parse import (drop_attribute_lines, drop_trailing_placeholder,
                               fold_lines, norm_code_lines, strip_export_header)
from features.common_subset import (  # noqa: E402
    close_owned as _close_owned,
    resolve_common_subset as _resolve_common_subset_shared,
    snapshot_module_code as _snapshot_module_code,
    warn_unknown_modules as _warn_unknown_modules,
)
from features.agent_probe import (  # noqa: F401  # 入口层重导出，既有符号契约不变
    _find,
    _wrap,
    probe,
)

# 文档模块名：ThisWorkbook / SheetN / ChartN …
DOC_NAME_RE = re.compile(
    r"^(ThisWorkbook|Sheet\d+|Chart\d+|DialogSheet\d+|MacroSheet\d+)$", re.IGNORECASE)
VB_NAME_RE = re.compile(
    r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"', re.IGNORECASE | re.MULTILINE)


# ---------------------------------------------------------------------------
# book：开 / 关工作簿
# ---------------------------------------------------------------------------
def _resolve(proj: dict) -> Path:
    """解析项目对应的 xlsm；解析不到报 TARGET_NOT_FOUND。

    留在入口薄层：测试既有桩直接替换本模块的 ``resolve_target`` 绑定，
    全局查找必须经过本模块命名空间。
    """
    path, _reason = resolve_target(proj)
    if path is None:
        raise ContractFailure(
            f"「{proj['name']}」解析不到 xlsm",
            code=ErrorCode.TARGET_NOT_FOUND, stage=Stage.CONFIG,
            hint="检查 mapping 的 search_dir / prefix，确认 xlsm 在目录里",
            retryable=False,
        )
    return path


# 日志汇总段依赖本层 ``_resolve``；放在其定义之后引入，避免循环导入时符号未绑定。
from features.agent_logs import logs  # noqa: E402,F401


def _excel_pid(xl) -> int:
    """取目标 Excel 进程 PID；读不到返回 0，不把权限问题误报成崩溃。"""
    try:
        import win32process
        return int(win32process.GetWindowThreadProcessId(int(xl.Hwnd))[1])
    except Exception:  # noqa: BLE001
        return 0


def _quit_if_empty(xl) -> tuple[bool, list[str]]:
    """关闭后实例已没有工作簿就退掉，避免累积可见空窗口；回收诊断不静默吞。

    返回 (是否走了退出, 问题清单)。空实例退出走 ``quit_owned`` 原语；读不到簿数量
    就不动实例，问题清单交调用方按 warnings/stats 通道上报。
    """
    try:
        if int(xl.Workbooks.Count) != 0:
            return False, []
    except Exception:  # noqa: BLE001 - 读不到簿数量就不动实例
        return False, []
    problems = quit_owned(xl)
    return True, problems


def book(ctx) -> Outcome:
    args = ctx.args
    proj = _find(ctx, getattr(args, "project", None))
    name = proj["name"]
    dry = not bool(getattr(args, "apply", False))

    if dry:
        path, _reason = resolve_target(proj)
        target = str(path) if path else "未解析"
        return Outcome(
            action="agent.book",
            ok=True,
            details=[{"dry": True, "project": name, "action": args.action}],
            stats={"dry": True, "project": name, "action": args.action, "target": target},
            human=[f"[dry] 将执行 book {args.action}（{name}）；加 --apply 才真正执行"],
            exit_code=0,
            elapsed_ms=ctx.elapsed_ms(),
        )

    path = _resolve(proj)
    errors: list[ContractFailure] = []
    human: list[str] = []
    warnings: list[str] = []
    cleanup_problems: list[str] = []

    if args.action == "open":
        xl = excel_gateway.new_excel(hidden=False, force_new=True)
        try:
            wb = excel_gateway.open_disk_workbook(xl, path, read_only=False)
            try:
                wb.Activate()
            except Exception as e:
                raise ContractFailure(
                    f"已打开，但无法把 Excel 窗口切到可见：{e}",
                    code=ErrorCode.COM_CONTEXT, stage=Stage.OPEN,
                    hint="检查 Excel 是否被最小化/隐藏，或换一个可见会话后重试",
                ) from e
            hwnd = int(getattr(xl, "Hwnd", 0) or 0)
            pid = _excel_pid(xl)
            human.append(f"已打开：{wb.FullName}（Excel 不执行 Workbook_Open，宏入口需手工点一次）")
            human.append(f"  可见：{bool(xl.Visible)}  hwnd={hwnd}  pid={pid}")
        except ContractFailure as e:
            errors.append(e)
            cleanup_problems.extend(quit_owned(xl))
        except BaseFailure as e:
            errors.append(_wrap(e))
            cleanup_problems.extend(quit_owned(xl))
        except Exception as e:  # noqa: BLE001 - 入口兜底
            errors.append(ContractFailure(str(e), code=ErrorCode.UNEXPECTED, stage=Stage.OPEN))
            cleanup_problems.extend(quit_owned(xl))
    else:
        try:
            xl = excel_gateway.attach_running_excel(
                path,
                target_hint="表没打开：无需关闭；要打开目标表用 book --action open",
                target_suggested=entry_command(
                    "agent.py", "book", "--action", "open",
                    "--project", name, "--apply"),
            )
            wb, _opened = excel_gateway.find_open_workbook(xl, path)
            if wb is None:
                raise ContractFailure(
                    f"没打开：{path.name}",
                    code=ErrorCode.TARGET_NOT_OPEN, stage=Stage.ATTACH,
                    hint="先在 Excel 里打开目标工作簿，或用 book --action open --apply",
                    suggested=entry_command("agent.py", "book", "--action", "open",
                                            "--project", name, "--apply"),
                )
            wb.Close(SaveChanges=False)
            quit_done, close_problems = _quit_if_empty(xl)
            cleanup_problems.extend(close_problems)
            human.append(f"已关闭（不保存）：{path.name}"
                         + ("；空 Excel 实例已退出" if quit_done else ""))
        except ContractFailure as e:
            errors.append(e)
        except BaseFailure as e:
            errors.append(_wrap(e))

    if cleanup_problems:
        warnings.append("实例回收未完成：" + "；".join(cleanup_problems[:3]))
        stats_cleanup = {"isolate_cleanup_problems": cleanup_problems}
    else:
        stats_cleanup = {"isolate_cleanup": "ok"}

    return Outcome(
        action="agent.book",
        ok=not errors,
        details=[{"project": name, "action": args.action, "ok": not errors}],
        errors=errors,
        stats={"project": name, "dry": False, **stats_cleanup},
        human=human,
        warnings=warnings,
        exit_code=1 if errors else 0,
        elapsed_ms=ctx.elapsed_ms(),
    )


# ---------------------------------------------------------------------------
# backup：整表备份到 .backup/<时间戳>-agent
# ---------------------------------------------------------------------------
def backup(ctx) -> Outcome:
    args = ctx.args
    proj = _find(ctx, getattr(args, "project", None))
    name = proj["name"]
    tag = (getattr(args, "tag", "") or "").strip()
    if tag and not tag.startswith("-"):
        tag = "-" + tag
    tag = tag or "-agent"
    dry = not bool(getattr(args, "apply", False))

    if dry:
        path, _reason = resolve_target(proj)
        target = str(path) if path else "未解析"
        return Outcome(
            action="agent.backup",
            ok=True,
            details=[{"dry": True, "project": name, "target": target}],
            stats={"dry": True, "project": name, "target": target, "tag": tag},
            human=[f"[dry] 将备份 {name} → .backup（tag={tag}）；加 --apply 才真正复制"],
            exit_code=0,
            elapsed_ms=ctx.elapsed_ms(),
        )

    path = _resolve(proj)
    errors: list[ContractFailure] = []
    human: list[str] = []
    dest: Path | None = None
    wb = None
    try:
        xl = excel_gateway.attach_running_excel(path)
        wb, _opened = excel_gateway.find_open_workbook(xl, path)
    except BaseFailure as e:
        if e.code not in ("NO_EXCEL", "TARGET_NOT_OPEN"):
            errors.append(_wrap(e))

    if not errors:
        try:
            if wb is not None:
                dest = backup_workbook(wb, ROOT / ".backup", tag=tag)
            elif path and Path(path).is_file():
                dest = backup_copy(Path(path), abs_path(ROOT / ".backup"), tag=tag)
                human.append(f"备份（表未开，直接复制文件）：{dest}")
            else:
                raise ContractFailure(
                    f"「{name}」没解析到可备份的 xlsm",
                    code=ErrorCode.TARGET_NOT_FOUND, stage=Stage.RESOLVE,
                    hint="核对 mapping.json 的 search_dir / prefix",
                    suggested=entry_command("sync.py", "list", "--json"),
                )
            human.append(f"备份完成：{dest}" if dest else "这张表还没落过盘，跳过备份")
        except ContractFailure as e:
            errors.append(e)
        except BaseFailure as e:
            errors.append(_wrap(e))

    return Outcome(
        action="agent.backup",
        ok=not errors,
        details=[{"project": name, "ok": not errors,
                  "backup": str(dest) if dest else ""}],
        errors=errors,
        stats={"project": name, "dry": False, "backup": bool(dest)},
        human=human,
        exit_code=1 if errors else 0,
        elapsed_ms=ctx.elapsed_ms(),
    )


# ---------------------------------------------------------------------------
# push：src→VBE 写回（预检 → 备份 → 写回 → 回读 → 失败回滚）
# ---------------------------------------------------------------------------
def _source_dirs(proj: dict, defaults: dict) -> list[Path]:
    """公共库目录 + 项目源码目录（与旧 sync 同口径）。"""
    dirs: list[Path] = []
    if defaults.get("common_dir"):
        dirs.append(abs_path(defaults["common_dir"]))
    if proj.get("src"):
        dirs.append(abs_path(proj["src"]))
    return dirs


def _common_dir(defaults: dict) -> Path:
    return abs_path(defaults.get("common_dir", "src/common"))


def _common_subset(proj: dict, defaults: dict) -> tuple[set[str] | None, str, list[str]]:
    """写回哪些公共模块：None=整个 common（未配置，旧行为）；set=只写子集。口径同 sync push。"""
    subset, reason, ignored = _resolve_common_subset_shared(proj, defaults)
    _warn_unknown_modules(ignored)
    return subset, reason, ignored


def _in_common_scope(folder: Path, f: Path, common_dir: Path,
                     subset: set[str] | None) -> bool:
    """公共库目录按子集收窄；项目源码目录与未配置（None）时全部纳入。"""
    if subset is None or folder != common_dir:
        return True
    return module_name_of(f) in subset


def _scope_module_count(proj: dict, defaults: dict) -> int:
    """本次写回实际涉及的模块数（项目源码全量 + 公共库子集），与写回/预检同口径。"""
    subset, _reason, _ignored = _common_subset(proj, defaults)
    common_dir = _common_dir(defaults)
    count = 0
    for folder in _source_dirs(proj, defaults):
        if not folder.is_dir():
            continue
        for f in source_files(folder):
            if _in_common_scope(folder, f, common_dir, subset):
                count += 1
    return count


def _preflight(proj: dict, defaults: dict, doc_modules: bool) -> list[str]:
    """源码硬预检（只读源码不碰 Excel）；返回致命项列表，空即通过。

    公共库按 common_modules 子集收窄：不在本次写回范围的模块不参与预检，
    避免未写回的模块因空正文/编码等问题报假致命项（口径同 sync push 写回范围）。

    有意保留：与 sync/run 预检的差异轴为 notes 软提醒、检查深度、子集过滤；
    不抽公共骨架，避免 sync 告警丢失或 agent 多出不适用的检查。
    """
    fatal: list[str] = []
    dirs = _source_dirs(proj, defaults)
    common_dir = _common_dir(defaults)
    subset, _reason, _ignored = _common_subset(proj, defaults)
    for folder in dirs:
        if not folder.is_dir():
            fatal.append(f"源码目录不存在：{folder}")

    files = 0
    for folder in dirs:
        if not folder.is_dir():
            continue
        local: dict[str, Path] = {}
        for f in source_files(folder):
            if not _in_common_scope(folder, f, common_dir, subset):
                continue
            files += 1
            name = module_name_of(f)
            is_doc = f.suffix.lower() == ".bas" and bool(DOC_NAME_RE.match(name))
            if name in local:
                fatal.append(f"同一目录内模块重名：{name}")
            local[name] = f
            if f.suffix.lower() == ".frm" and not f.with_suffix(".frx").is_file():
                fatal.append(f"窗体缺同名 .frx：{f.name}")
            try:
                text = read_source_text(f)
            except OSError as e:
                fatal.append(f"源码读不出：{f.name}（{e}）")
                continue
            m = VB_NAME_RE.search(text)
            if m and m.group(1).strip() != name:
                fatal.append(f"文件头 VB_Name 与文件名不一致：{f.name}")
            if not is_doc and f.suffix.lower() in (".bas", ".cls"):
                body = [ln.rstrip() for ln in norm_code_lines(text) if ln.strip()]
                if not body:
                    fatal.append(f"模块正文为空：{f.name}")
    if files == 0:
        fatal.append("源码目录里没有任何可写模块")
    return fatal


def _code_body(text: str, vbe: bool = False) -> list[str]:
    """比对用正文：剥过程级 Attribute；VBE 侧再抹尾部 () 占位。"""
    text = drop_attribute_lines(text)
    if vbe:
        text = drop_trailing_placeholder(text)
    return [ln for ln in norm_code_lines(text) if ln.strip()]


def _writeback(wb: Any, proj: dict, defaults: dict, doc_modules: bool,
               *, rollback_on_fail: bool = False
               ) -> tuple[list[str], dict[str, Path]]:
    """把 common + 项目源码写回工作簿；返回 (回读失败明细, 本次写回的模块)。

    策略与旧 sync 一致：
      · 已有同类型标准/类模块 → 原地改正文（不 Remove，避免 VBE 补序号）；
      · 新模块 / 类型不符 / 窗体 → 删旧后 Import（import_component 保证落名）；
      · 文档模块按 doc_modules 开关，跳过则不写；
      · 回读逐模块比对正文指纹；rollback_on_fail=True 时失败按写回前内存正文
        复原本次写过的模块（默认不回滚，与 sync push 口径一致）。
    """
    vbproj = wb.VBProject
    common_dir = _common_dir(defaults)
    subset, _reason, _ignored = _common_subset(proj, defaults)
    before = component_snapshot(vbproj)
    code_before = _snapshot_module_code(vbproj) if rollback_on_fail else {}
    written: dict[str, Path] = {}

    tmp_parent = Path(wb.FullName).parent if getattr(wb, "Path", "") else None
    tmp_dir: Path | None = None
    try:
        for folder in _source_dirs(proj, defaults):
            if not folder.is_dir():
                continue
            for f in source_files(folder):
                if not _in_common_scope(folder, f, common_dir, subset):
                    continue
                name = module_name_of(f)
                comp = find_component(vbproj, name)
                ctype = comp.Type if comp is not None else None
                is_doc = ctype == VBXT_DOCUMENT or (
                    ctype is None and bool(DOC_NAME_RE.match(name))
                    and name.lower().startswith(("thisworkbook", "sheet", "chart")))
                if is_doc:
                    if not doc_modules:
                        continue
                    replace_module_code(comp, strip_export_header(read_source_text(f)))
                    written[name] = f
                    continue

                in_place = (ctype in (VBXT_STANDARD, VBXT_CLASS)
                            and TYPE_FOR_EXT.get(f.suffix.lower()) == ctype)
                if in_place:
                    replace_module_code(comp, strip_export_header(read_source_text(f)))
                    written[name] = f
                    continue

                if comp is not None and not remove_module(vbproj, name):
                    # 旧模块删不掉就放弃 Import，避免造出 name1 后缀模块
                    continue
                if tmp_dir is None:
                    base = str(tmp_parent) if tmp_parent and tmp_parent.is_dir() else None
                    tmp_dir = Path(tempfile.mkdtemp(prefix="vba-import-", dir=base))
                new_comp = import_component(vbproj, gbk_import_copy(f, tmp_dir), name)
                if new_comp is not None:
                    written[name] = f
    finally:
        if tmp_dir is not None:
            remove_tree(tmp_dir)

    # 模块清单核对：写回只该改正文，不该增删模块（差异只记录，不阻断）
    after = component_snapshot(vbproj)
    check_module_set(before, after, src_names=list(written))

    # 回读指纹比对
    bad: list[str] = []
    for name, path in written.items():
        comp = find_component(vbproj, name)
        if comp is None:
            bad.append(f"{name}（工程里找不到）")
            continue
        try:
            cm = comp.CodeModule
            got = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
            want = read_source_text(path)
        except Exception as e:  # noqa: BLE001 - 回读/读源码失败也算校验失败
            bad.append(f"{name}（回读失败：{e}）")
            continue
        if fold_lines(_code_body(want)) != fold_lines(_code_body(got, vbe=True)):
            bad.append(f"{name}（正文与源码不一致）")

    # 失败回滚（仅 --rollback-on-fail）：已有模块复原写回前正文；新增模块移除。
    # 复原失败必须记进 bad 明细（与 sync push 口径一致），不能静默吞掉——否则用户
    # 会以为回滚成功，内存里仍留着写回失败的代码。
    if bad and rollback_on_fail:
        rollback_failed: list[str] = []
        for name in written:
            comp = find_component(vbproj, name)
            if comp is None:
                rollback_failed.append(f"{name}（工程里找不到）")
                continue
            if name in code_before:
                try:
                    replace_module_code(comp, code_before[name])
                except Exception as e:  # 逐条记失败，不中断其它模块
                    rollback_failed.append(f"{name}（{e}）")
            else:
                remove_module(vbproj, name)
        if rollback_failed:
            bad.append("回滚未完成：" + "；".join(rollback_failed))

    return bad, written


def _attach_or_open(proj: dict, target: Path, *,
                    auto_open: bool = True) -> tuple[Any, Any, str]:
    """连上目标工作簿；返回 (wb, owned_xl, mode)。已打开借会话，没开走隐藏实例。

    auto_open=False（--open-only）：表没开就报 TARGET_NOT_OPEN，不自动开隐藏实例。

    有意保留：错误码/排障提示按入口不同，合并需同步改错误契约（见 1.6.x 统一化方案）。
    """
    try:
        xl = excel_gateway.attach_running_excel(
            Path(target),
            target_hint="表没打开：写回默认会自动开隐藏实例；检查文件占用/受保护视图",
            target_suggested=entry_command(
                "agent.py", "book", "--action", "open",
                "--project", proj["name"], "--apply"),
        )
        wb, _opened = excel_gateway.find_open_workbook(xl, Path(target))
        if wb is None:
            raise ContractFailure(
                f"目标表没有打开：{target}",
                code=ErrorCode.TARGET_NOT_OPEN, stage=Stage.ATTACH,
                hint="表没打开：写回默认会自动开隐藏实例；检查文件占用/受保护视图",
            )
        return wb, None, "open"
    except ContractFailure as e:
        raise
    except BaseFailure as e:
        if e.code not in ("NO_EXCEL", "TARGET_NOT_OPEN"):
            raise _wrap(e)
        if not auto_open:
            raise ContractFailure(
                f"目标表没有打开：{target}",
                code=ErrorCode.TARGET_NOT_OPEN, stage=Stage.ATTACH,
                hint="已指定 --open-only：不会自动开表；先在 Excel 里打开目标表，"
                     "或去掉 --open-only 让工具开隐藏实例写回",
            )
        owned_xl = excel_gateway.new_excel(hidden=True)
        wb = excel_gateway.open_disk_workbook(owned_xl, Path(target), read_only=False)
        if not ensure_vbe_access(wb, proj["name"], interactive=False):
            state, detail = vbe_state(wb)
            code = ErrorCode.VBE_PROTECTED if state == "protected" else ErrorCode.COM_CONTEXT
            raise ContractFailure(
                f"「{proj['name']}」的 VBA 工程读不了：{detail}",
                code=code, stage=Stage.VBE,
                suggested="先在 Excel 里解开工程保护/勾选 AccessVBOM，再重跑写回",
            )
        return wb, owned_xl, "disk"


def push(ctx) -> Outcome:
    args = ctx.args
    proj = _find(ctx, getattr(args, "project", None))
    name = proj["name"]
    defaults = ctx.config["defaults"]
    doc_modules = (bool(getattr(args, "doc_modules", False))
                   or bool(defaults.get("import_document_modules")))
    backup = not bool(getattr(args, "no_backup", False))
    save = not bool(getattr(args, "no_save", False))
    auto_open = not bool(getattr(args, "open_only", False))
    rollback = bool(getattr(args, "rollback_on_fail", False))
    switches = {"doc_modules": doc_modules, "backup": backup, "save": save,
                "auto_open": auto_open, "rollback_on_fail": rollback}
    mods = _scope_module_count(proj, defaults)
    dry = not bool(getattr(args, "apply", False))

    fatal = _preflight(proj, defaults, doc_modules)
    target, _reason = resolve_target(proj)
    target_s = str(target) if target else "未解析"

    if fatal:
        return Outcome(
            action="agent.push",
            ok=False,
            details=[],
            errors=[ContractFailure(
                "源码预检未通过，已放弃写回：" + "；".join(fatal[:3]),
                code=ErrorCode.SOURCE_PREFLIGHT, stage=Stage.PUSH,
                hint="按提示修源码/文件名后重跑；不要绕过预检直接写回",
                suggested=entry_command("agent.py", "probe", "--json"),
                retryable=False)],
            stats={"project": name, "target": target_s, "modules": mods,
                   "dry": dry, **switches},
            exit_code=1,
            elapsed_ms=ctx.elapsed_ms(),
        )

    switch_human = "、".join(
        f"{label}{'是' if switches[key] else '否'}"
        for key, label in (("doc_modules", "文档模块"), ("backup", "写前备份"),
                           ("save", "落盘保存"), ("auto_open", "自动开表"),
                           ("rollback_on_fail", "失败回滚")))

    if dry:
        return Outcome(
            action="agent.push",
            ok=True,
            details=[{"name": name, "ok": True, "target": target_s,
                      "modules": mods, "dry": True}],
            stats={"project": name, "target": target_s, "modules": mods,
                   "dry": True, **switches},
            human=[f"[dry] 将写回 {name}（{mods} 个模块）；加 --apply 才真正写回",
                   f"  开关：{switch_human}"],
            exit_code=0,
            elapsed_ms=ctx.elapsed_ms(),
        )

    target = _resolve(proj)
    errors: list[ContractFailure] = []
    human: list[str] = []
    warnings: list[str] = []
    wb = None
    owned_xl = None
    mode = ""
    backup_path: Path | None = None
    saved_ok = False
    written: dict[str, Path] = {}
    cleanup_problems: list[str] = []
    try:
        with ctx.progress.stage("attach"):
            wb, owned_xl, mode = _attach_or_open(proj, target, auto_open=auto_open)
        with ctx.progress.stage("backup"):
            if backup:
                backup_root = abs_path(defaults.get("backup_dir", ".backup"))
                backup_path = backup_copy(target, backup_root)
                human.append(f"  备份：{backup_path}")
        with ctx.progress.stage("writeback"):
            bad, written = _writeback(wb, proj, defaults, doc_modules,
                                       rollback_on_fail=rollback)
        with ctx.progress.stage("readback"):
            if bad:
                raise ContractFailure(
                    "写回后回读校验失败，未保存：" + "；".join(bad),
                    code=ErrorCode.POSTCHECK_FAILED, stage=Stage.PUSH,
                    hint="检查源码编码/Attribute 行；修好后重跑，不要手工保存这次结果",
                )
        # 按口径还原应用状态，再落盘保存；还原改动/失败并入 warnings（口径同 macro.py）。
        # 单一还原通道：guard.restore() 已覆盖布尔项、Calculation、StatusBar、CutCopyMode
        # 及「还原未生效」回读；不再额外调用模块级 restore，避免二次诊断重复刷屏。
        guard = app_guard.guard(wb.Application)
        try:
            guard.restore()
            if save:
                wb.Save()
                saved_ok = True
                human.append("已写入 VBA 工程并保存到 xlsm（不需要再按 Ctrl+S）")
            else:
                human.append("已写入 VBA 工程内存（--no-save：未保存到 xlsm）")
        except Exception as e:
            raise ContractFailure(
                f"写回成功，但保存失败：{e}；请回 Excel 手动 Ctrl+S",
                code=ErrorCode.SAVE_FAILED, stage=Stage.SAVE,
                suggested="检查文件是否只读、是否被网盘/Excel 占用，处理后重跑",
            ) from e
        if guard.changes:
            warnings.append(f"已还原 Excel 状态：{guard.describe()}")
        for line in guard.errors[:3]:
            warnings.append(f"状态还原未完成：{line}")
    except ContractFailure as e:
        errors.append(e)
    except BaseFailure as e:
        errors.append(_wrap(e))
    finally:
        cleanup_problems = _close_owned(wb, owned_xl)

    if cleanup_problems:
        warnings.append("隔离实例回收未完成：" + "；".join(cleanup_problems[:3]))
        stats_cleanup = {"isolate_cleanup_problems": cleanup_problems}
    else:
        stats_cleanup = {"isolate_cleanup": "ok"}

    result_text = "写回完成" if not errors else "失败"
    return Outcome(
        action="agent.push",
        ok=not errors,
        details=[{"name": name, "ok": not errors, "mode": mode,
                  "result": result_text, "modules": mods, "saved": saved_ok}],
        errors=errors,
        stats={"project": name, "mode": mode, "result": result_text, "dry": False,
               "saved": saved_ok, "backup": bool(backup) and bool(backup_path),
               "backup_path": str(backup_path) if backup_path else "",
               "written": sorted(written), **stats_cleanup},
        human=human,
        warnings=warnings,
        exit_code=1 if errors else 0,
        elapsed_ms=ctx.elapsed_ms(),
    )


SPECS = (
    CommandSpec(
        verb="probe", entry="agent", handler="features.agent_io:probe",
        mutation=MutationMode.NONE, project=ProjectMode.OPTIONAL,
        supports_json=True, exit_codes=(0, 1, 2),
        args=(ArgSpec("--strict", kind="flag",
                      help="Excel 没开或 COM 不接话也按失败退出（健康门禁）"),),
        summary="项目状态一览：磁盘/已打开/工程/模块数/测试工作表行数",
        ref="references/agent.md",
        tests=("contract.probe.success", "contract.probe.target-not-found"),
    ),
    CommandSpec(
        verb="book", entry="agent", handler="features.agent_io:book",
        mutation=MutationMode.REQUIRED, project=ProjectMode.OPTIONAL,
        supports_json=True, exit_codes=(0, 1, 2),
        args=(ArgSpec("--action", kind="choice", choices=("open", "close"), required=True,
                      help="open 打开可见实例 / close 关闭工作簿（不保存）"),),
        summary="开/关工作簿",
        ref="references/agent.md",
        tests=("contract.book.open", "contract.book.close-not-open"),
    ),
    CommandSpec(
        verb="backup", entry="agent", handler="features.agent_io:backup",
        mutation=MutationMode.REQUIRED, project=ProjectMode.OPTIONAL,
        supports_json=True, exit_codes=(0, 1, 2),
        args=(ArgSpec("--tag", kind="str", default="",
                      help="备份标签（默认 agent；加在时间戳后）"),),
        summary="整工作簿备份到 .backup/<时间戳>-agent",
        ref="references/agent.md",
        tests=("contract.backup.success", "contract.backup.target-not-found"),
    ),
    CommandSpec(
        verb="logs", entry="agent", handler="features.agent_io:logs",
        mutation=MutationMode.NONE, project=ProjectMode.OPTIONAL,
        supports_json=True, exit_codes=(0, 1, 2),
        args=(ArgSpec("--log", kind="str", default="", help="日志目录（默认项目 测试输出）"),
              ArgSpec("--tail", kind="int", default=40, help="日志尾部行数（默认 40，0=不打）")),
        summary="最新测试日志摘要",
        ref="references/agent.md",
        tests=("contract.logs.success", "contract.logs.not-found"),
    ),
    CommandSpec(
        verb="push", entry="agent", handler="features.agent_io:push",
        mutation=MutationMode.REQUIRED, project=ProjectMode.OPTIONAL,
        supports_json=True, exit_codes=(0, 1, 2),
        args=(
            ArgSpec("--doc-modules", kind="flag",
                    help="连文档模块（Sheet/ThisWorkbook）一起覆盖"),
            ArgSpec("--no-save", kind="flag", help="写回后不保存，只改内存"),
            ArgSpec("--no-backup", kind="flag", help="跳过写回前备份"),
            ArgSpec("--open-only", kind="flag",
                    help="只写已打开的表；表没开就失败（旧行为，不自动开隐藏实例）"),
            ArgSpec("--rollback-on-fail", kind="flag",
                    help="postcheck 失败时按写回前内存正文复原本次写过的模块（默认不回滚）"),
        ),
        summary="src→VBE 写回工作簿",
        ref="references/agent.md",
        tests=("contract.push.preview", "contract.push.project-not-found"),
    ),
)
