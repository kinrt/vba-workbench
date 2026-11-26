# -*- coding: utf-8 -*-
"""sync_read 切片：vbamod / plan / pull / pull-open（只读与预览类动词）。

边界规矩：
  · handler 只返回 Outcome，不 print、不自带 JSON；
  · COM 附着/打开一律走 adapters.excel_gateway；工程状态读 toolkit.excel 边界层；
  · 只读/预览默认不落盘，只有带 --apply 才写 src；
  · 命令名、参数名、JSON 键、退出码与 1.5.6 对齐。
"""
from __future__ import annotations

import re
from pathlib import Path

from adapters import excel_gateway
from toolkit.base import (ROOT, Failure as BaseFailure, SOURCE_ALL_EXTS,
                          SOURCE_EXTS, abs_path, entry_command, read_source_text)
from toolkit.check import Notes, check_export_set
from toolkit.contract import (ErrorCode, Failure, MutationMode, Outcome, Stage,
                              to_contract_failure)
from toolkit.excel import (ExcelSession, component_snapshot, ensure_vbe_access,
                           find_excel_with_workbook, has_vbproject_access,
                           require_open_workbook, vbe_state,
                           vbe_state_hint)
from toolkit.files import latest_backup, stamp_dir, transcode_exported_file, write_source
from toolkit.repo import (EXT_FOR_TYPE, TYPE_FOLDER, VBXT_DOCUMENT, VERSION_RE,
                          doc_stem, is_shell_file, module_index, module_name_of,
                          resolve_target, shell_text, source_files)
from toolkit.registry import ArgSpec, CommandSpec, ProjectMode
from toolkit.vba_parse import declared_procs, mask_strings, strip_vba_comments
from toolkit.vbabin import compare_vba_modules, vba_module_names

from features.common_subset import (  # noqa: E402
    find_project as _find_project,
    resolve_common_subset as _resolve_common_subset_shared,
    scope_human as _scope_human,
    scope_payload as _scope_payload,
    warn_unknown_modules as _warn_unknown_modules,
)


# ---------------------------------------------------------------------------
# 项目解析与公共库子集（纯本地规则，不碰 COM）
# ---------------------------------------------------------------------------
def _backup_root(cfg: dict) -> Path:
    return abs_path(cfg["defaults"].get("backup_dir", ".backup"))


def _resolve_common_subset(proj: dict, defaults: dict) -> tuple[set[str] | None, str]:
    """决定这个项目要写回哪些公共模块；未配置返回 None（整个 common）。"""
    subset, reason, ignored = _resolve_common_subset_shared(proj, defaults)
    _warn_unknown_modules(ignored)
    return subset, reason


def _common_subset_stems(proj: dict, defaults: dict) -> set[str]:
    """这个项目最终用到的公共模块名集合（None 视为全部）。"""
    all_mods = set(module_index(abs_path(defaults.get("common_dir", "src/common"))))
    subset, _ = _resolve_common_subset(proj, defaults)
    return all_mods if subset is None else subset


# ---------------------------------------------------------------------------
# 已打开工作簿连接（pull-open 显式入口与 pull 自动转会话共用；规则与 1.5.6 一致）
# ---------------------------------------------------------------------------
def _closed_hint(proj_name: str, purpose: str) -> tuple[str, str]:
    """表没开时按动词给可执行出口。返回 (hint, suggested_command)。"""
    book = entry_command("agent.py", "book", "--project", proj_name,
                         "--action", "open", "--apply")
    if purpose == "抓取":
        return (f"表没打开：pull 将走磁盘副本（sync.py pull --project {proj_name} --apply）；"
                f"要抓内存里的最新代码请先在 Excel 打开表，或显式用 pull-open",
                entry_command("sync.py", "pull", "--project", proj_name, "--apply"))
    return ("表未打开：需要先打开目标工作簿", book)


def _connect_open_workbook(proj: dict, purpose: str):
    """连上配置那张表的已打开工作簿，并确认 VBProject 可读；找不到绝不顶替。"""
    path, _ = resolve_target(proj)
    hint, cmd = _closed_hint(proj["name"], purpose)
    if path is None:
        raise Failure(
            f"「{proj['name']}」没解析到 xlsm：配置里 search_dir / prefix 找不到匹配文件",
            code=ErrorCode.TARGET_NOT_FOUND, stage=Stage.RESOLVE,
            hint="核对 mapping.json 的 search_dir 与 prefix；文件名必须是「prefix V<数字>.xlsm」",
            suggested=entry_command("sync.py", "list", "--json"))
    xl = excel_gateway.attach_running_excel(Path(path), target_hint=hint, target_suggested=cmd)
    wb = require_open_workbook(xl, Path(path), proj["name"], purpose,
                               hint=hint, suggested=cmd)
    if not ensure_vbe_access(wb, proj["name"]):
        state, detail = vbe_state(wb)
        # 旧版区分 protected / denied 两态；契约只登记 VBE_PROTECTED，
        # AccessVBOM 未勾选同属 VBE 权限类失败，一并归到 VBE_PROTECTED。
        raise Failure(
            f"「{proj['name']}」的 VBA 工程读不了，这次操作没做（处理好后重跑，或重新来一次）",
            code=ErrorCode.VBE_PROTECTED, stage=Stage.VBE,
            suggested="在 Excel 里解开工程保护/勾选 AccessVBOM 后重跑")
    return wb


# 附着失败里允许退回磁盘模式的错误码：只有“表确实没开”才退；
# COM_CONTEXT / EXCEL_BUSY / EXCEL_CLOSED / VBE_PROTECTED 一律原样上报，
# 不能静默读陈旧磁盘副本，把未保存改动丢掉。
_PULL_DISK_FALLBACK_CODES = ("NO_EXCEL", "TARGET_NOT_OPEN")


def _attach_pull_session(proj: dict, path: Path, purpose: str = "抓取") -> tuple[object | None, str]:
    """pull 专用附着：表已开 → 连已打开会话；明确没开 → 交回磁盘模式。

    返回 ``(wb, mode)``：mode 为 ``open`` 时 wb 是用户已打开的工作簿（借会话，
    不另开副本、不关窗）；mode 为 ``disk`` 时 wb 为 None，调用方走隐藏实例读磁盘。

    除 ``NO_EXCEL`` / ``TARGET_NOT_OPEN`` 外的附着/会话失败直接上抛，
    由调用方记失败行，禁止吞成普通磁盘成功。
    """
    hint, cmd = _closed_hint(proj["name"], purpose)
    try:
        xl = excel_gateway.attach_running_excel(
            Path(path), target_hint=hint, target_suggested=cmd)
    except BaseFailure as e:
        if e.code in _PULL_DISK_FALLBACK_CODES:
            return None, "disk"
        # 已打开会话路径失败：标注“本应走会话但失败”，不退回磁盘读陈旧副本。
        raise Failure(
            str(e), code=e.code, stage=e.stage,
            hint="本应借当前已打开会话抓取，但会话连不上；未退回磁盘读陈旧副本。"
                 "检查 Excel 是否卡在弹窗/受保护视图，处理好后重跑 pull。",
            suggested=e.suggested or cmd, retryable=e.retryable) from e
    wb = require_open_workbook(xl, Path(path), proj["name"], purpose,
                               hint=hint, suggested=cmd)
    if not ensure_vbe_access(wb, proj["name"]):
        _state, detail = vbe_state(wb)
        # 与 pull-open 同一关卡：权限类失败统一归 VBE_PROTECTED，不退磁盘。
        raise Failure(
            f"「{proj['name']}」的 VBA 工程读不了，这次操作没做（处理好后重跑，或重新来一次）",
            code=ErrorCode.VBE_PROTECTED, stage=Stage.VBE,
            suggested="在 Excel 里解开工程保护/勾选 AccessVBOM 后重跑")
    return wb, "open"


# ---------------------------------------------------------------------------
# 抓取（VBE 工程 → src）；dry=True 只预览不落盘
# ---------------------------------------------------------------------------
def _document_labels(wb) -> dict[str, str]:
    """{CodeName: 表页名}：工作表 / 图表 / 宏表 / 对话框表都收。"""
    out: dict[str, str] = {}
    try:
        for sh in wb.Sheets:
            try:
                out[str(sh.CodeName)] = str(sh.Name)
            except Exception:
                continue
    except Exception:
        return out
    return out


def _export_component(comp, out_dir: Path, dry: bool, sheet_label: str = "") -> Path:
    """导出一个 VBE 组件到分类目录，并清掉同名旧位置文件；GBK 就地转 UTF-8 + LF。"""
    name = comp.Name
    ext = EXT_FOR_TYPE.get(comp.Type, ".bas")
    sub = TYPE_FOLDER.get(comp.Type, "模块")
    stem = doc_stem(name, sheet_label) if comp.Type == VBXT_DOCUMENT else name
    target = out_dir / sub / (stem + ext)
    target.parent.mkdir(parents=True, exist_ok=True)

    for old in out_dir.rglob("*"):
        if not old.is_file() or is_shell_file(old):
            continue
        if module_name_of(old) == name and (old.suffix.lower() in SOURCE_ALL_EXTS or old.suffix == ""):
            if old.resolve() == target.resolve():
                continue
            if not dry:
                old.unlink()

    if not dry:
        comp.Export(str(target))
        transcode_exported_file(target)
    return target


def _export_workbook(wb, out_dir: Path, dry: bool, common_stems: set[str] | None,
                     human: list[str]) -> int:
    """工程全部模块按分类抓取到 out_dir；公共库模块只留占位，不抓正文。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    count = 0
    existing = {c.Name for c in wb.VBProject.VBComponents}
    common_stems = common_stems or set()
    labels = _document_labels(wb)

    for comp in wb.VBProject.VBComponents:
        if comp.Name in common_stems:
            human.append(f"    . 公共库模块 {comp.Name}：不抓取正文，保留占位标记")
            continue
        _export_component(comp, out_dir, dry, labels.get(comp.Name, ""))
        count += 1

    if not dry:
        for f in sorted(out_dir.rglob("*")):
            if not f.is_file() or f.suffix.lower() not in SOURCE_ALL_EXTS:
                continue
            if is_shell_file(f):
                continue
            if module_name_of(f) not in existing:
                f.unlink()
    return count


def _write_placeholder_shells(proj: dict, defaults: dict, dry: bool,
                              human: list[str], warnings: list[str]) -> int:
    """为项目用到的公共模块在项目目录放同名注释壳；已有真身绝不覆盖。"""
    src_dir = abs_path(proj["src"])
    if not src_dir.is_dir():
        return 0

    common_dir = abs_path(defaults.get("common_dir", "src/common"))
    all_mods = module_index(common_dir)
    subset, _reason = _resolve_common_subset(proj, defaults)
    listed = set(all_mods) if subset is None else set(subset)

    for f in sorted(src_dir.rglob("*")):
        if f.is_file() and is_shell_file(f) and f.stem in all_mods:
            listed.add(f.stem)

    for f in sorted(src_dir.rglob("*")):
        if f.is_file() and is_shell_file(f) and f.stem not in listed:
            if not dry:
                f.unlink()

    real_stems: dict[str, Path] = {}
    for f in src_dir.rglob("*"):
        if f.is_file() and f.suffix.lower() in SOURCE_EXTS and not is_shell_file(f):
            real_stems.setdefault(f.stem, f)

    written = 0
    for stem in sorted(listed):
        if stem in real_stems:
            warnings.append(
                f"项目里已有同名真实代码 {real_stems[stem].relative_to(src_dir).as_posix()}，"
                f"不覆盖；请用 promote 把它搬进公共库")
            continue
        src_file = all_mods[stem]
        try:
            rel_in_repo = src_file.relative_to(ROOT).as_posix()
        except ValueError:
            rel_in_repo = src_file.as_posix()
        folder_name = src_file.parent.name
        target = src_dir / folder_name / (stem + src_file.suffix)
        written += 1
        if dry:
            continue
        target.parent.mkdir(parents=True, exist_ok=True)
        write_source(target, shell_text(rel_in_repo, src_file.suffix.lower()))
    return written


def _verify_export_set(wb, out_dir: Path, common_stems: set[str] | None,
                        human: list[str], warnings: list[str], notes: Notes) -> None:
    """抓取后核对：工程每个模块是否都落成源码文件；缺项点出来。"""
    snap = component_snapshot(wb.VBProject)
    if snap is None:
        return
    have = {module_name_of(f) for f in source_files(out_dir)}
    skip = set(common_stems or ())
    missing = check_export_set(snap.keys(), have, skip)
    if missing:
        warnings.append(
            f"抓取后核对：工程里有 {len(missing)} 个模块在源码目录没有对应文件：{'、'.join(missing)}")
        notes.add("校验", f"抓取后 {len(missing)} 个模块没有源码文件：{'、'.join(missing)}")
    else:
        human.append(f"  校验：工程 {len(snap)} 个模块在源码目录都有文件（公共库 {len(skip)} 个占位不计）✔")
        notes.add("校验", f"工程 {len(snap)} 个模块在源码目录都有文件（公共库 {len(skip)} 个占位不计）✔")


# ---------------------------------------------------------------------------
# 动词 handler：每个返回 Outcome
# ---------------------------------------------------------------------------
def vbamod(ctx) -> Outcome:
    """核对各表模块清单（和最近一次备份比）；只读 vbaProject.bin，不需要 Excel。"""
    cfg = ctx.config
    args = ctx.args
    notes = Notes("核对模块清单提醒")
    human: list[str] = []
    warnings: list[str] = []
    errors: list[Failure] = []
    rows: list[dict] = []

    try:
        with ctx.progress.stage("解析项目"):
            name = getattr(args, "project", None) or None
            if name:
                proj = _find_project(cfg, name)
                path, reason = resolve_target(proj)
                if not path:
                    errors.append(Failure(
                        reason, code=ErrorCode.TARGET_NOT_FOUND, stage=Stage.RESOLVE,
                        hint="核对 mapping.json 的 search_dir 与 prefix；文件名必须是「prefix V<数字>.xlsm」",
                        suggested=entry_command("sync.py", "list", "--json")))
                    human.append(f"=== {proj['name']} ===")
                    human.append(f"  跳过：{reason}")
                    return _finish("sync.vbamod", rows, errors, human, warnings, notes, ctx)
                targets: list[tuple[str, Path]] = [(proj["name"], Path(path))]
            else:
                targets = []
                for proj in cfg["projects"]:
                    path, _ = resolve_target(proj)
                    if path:
                        targets.append((proj["name"], Path(path)))
            if not targets:
                errors.append(Failure(
                    "没有可解析的表：先跑 list 看项目清单",
                    code=ErrorCode.TARGET_NOT_FOUND, stage=Stage.RESOLVE,
                    suggested=entry_command("sync.py", "list", "--json")))
                return _finish("sync.vbamod", rows, errors, human, warnings, notes, ctx)

        with ctx.progress.stage("核对模块"):
            for proj_name, path in targets:
                names = vba_module_names(path)
                human.append("")
                human.append(f"=== {proj_name} ===")
                human.append(f"  文件：{path}")
                if names:
                    human.append(f"  模块 {len(names)} 个：{'、'.join(names)}")
                else:
                    human.append("  工程里没有模块")

                row: dict = {"name": proj_name, "file": str(path), "modules": len(names),
                             "module_names": names, "compared": False, "status": "未判定"}
                rows.append(row)
                backup = latest_backup(_backup_root(cfg), path.name)
                if backup is None:
                    notes.add("未判定", f"{proj_name}：没有备份，模块集合未核对（不能据此判“无异常”）")
                    continue
                diff = compare_vba_modules(backup, path)
                row.update(compared=True, same=bool(diff["same"]), status="已比对",
                           count_before=diff["count_before"], count_after=diff["count_after"],
                           lost=diff["lost"], added=diff["added"])
                if diff["same"]:
                    notes.add("校验", f"{proj_name}：模块集合与备份一致（{diff['count_after']} 个）✔")
                else:
                    if diff["lost"]:
                        notes.add("警告", f"{proj_name}：比备份少了模块 {'、'.join(diff['lost'])}")
                    if diff["added"]:
                        notes.add("警告", f"{proj_name}：比备份多了模块 {'、'.join(diff['added'])}")
    except BaseFailure as e:
        errors.append(to_contract_failure(e))

    return _finish("sync.vbamod", rows, errors, human, warnings, notes, ctx)


def plan(ctx) -> Outcome:
    """只读推荐：按目标表当前状态，给出 open / disk / isolate / open_required / unavailable 的结论。"""
    cfg = ctx.config
    args = ctx.args
    name = args.project
    intent = args.intent
    errors: list[Failure] = []
    human: list[str] = []
    warnings: list[str] = []

    with ctx.progress.stage("解析目标"):
        proj = _find_project(cfg, name)
        path, reason = resolve_target(proj)

    row: dict = {
        "name": name, "intent": intent, "ok": bool(path),
        "workbook": str(path) if path else "", "target_reason": reason,
        "open": False, "saved": None, "vbe_state": "", "probe": "",
        "mode": "", "needs_open": False, "will_save": False,
        "preserve_unsaved": False, "command": "", "risk": "",
    }

    if path:
        try:
            xl = excel_gateway.attach_running_excel(Path(path))
            wb, _opened = excel_gateway.find_open_workbook(xl, Path(path))
            if wb is not None:
                row["open"] = True
                try:
                    row["saved"] = bool(wb.Saved)
                except Exception:
                    pass
                try:
                    row["vbe_state"] = vbe_state(wb)[0]
                except Exception:
                    pass
        except BaseFailure as e:
            row["probe"] = getattr(e, "code", "") or ErrorCode.PROBE_FAILED.value
            row["probe_detail"] = str(e)

    if row["probe"] == "COM_CONTEXT":
        row["mode"] = "unavailable"
        row["needs_open"] = True
        row["command"] = ""
        row["risk"] = ("表可能开着，但当前进程连不上它的 COM 会话；不要直接走磁盘写回，"
                       "先按 probe_detail 检查弹窗/受保护视图/权限，或关表后改用 --isolate")
    elif not path:
        row["mode"] = "unavailable"
        row["command"] = ""
        row["risk"] = f"配置解析不到文件：{reason}"
    elif intent == "pull":
        row["mode"] = "open" if row["open"] else "disk"
        row["needs_open"] = False
        row["preserve_unsaved"] = bool(row["open"])
        # pull 检测到表已开会自动借已打开会话抓取；命令统一给 pull，
        # 不必手敲 pull-open（pull-open 仍保留为显式入口）。
        row["command"] = f"sync.py pull --project {name} --apply"
    elif intent == "push":
        row["mode"] = "open" if row["open"] else "disk"
        row["needs_open"] = False
        row["will_save"] = True
        row["command"] = f"sync.py push --project {name} --apply"
        if row["open"] and row["saved"] is False:
            row["risk"] = (
                "表已打开且有未保存改动。三条可执行处置："
                "① 先另存一份副本，把 mapping 的 file 指向副本后对副本 push（保数据也保 VBA）；"
                "② 先手动保存当前工作簿（数据与 VBA 一起落盘），再 push；"
                "③ 只想保留磁盘上旧数据：关表选「不保存」，再 push（内存里尚未落盘的 VBA 改动会丢）。"
                "注意 --no-save 只是“VBA 只改内存、不写 xlsm”，并不保护单元格改动；"
                "push 失败还会把 VBE 内存工程改脏，需要自动复原就加 --rollback-on-fail")
        elif not row["open"]:
            row["risk"] = ("表未打开：push 会自动开隐藏 Excel 写回并保存；"
                           "postcheck 失败默认只在内存留下半成品、不保存；"
                           "要按写回前正文复原本次模块，加 --rollback-on-fail")
    elif intent == "run":
        row["mode"] = "open" if row["open"] else "isolate"
        row["needs_open"] = False
        row["command"] = (f"sync.py run --project {name} --macro <宏名> --apply" if row["open"]
                         else f"sync.py run --project {name} --macro <宏名> --isolate --apply")
        if not row["open"]:
            row["risk"] = "隔离实例不触发 Workbook_Open；需要事件/界面的宏请先打开表"
    elif intent == "watch":
        row["mode"] = "open" if row["open"] else "open_required"
        row["needs_open"] = not row["open"]
        row["command"] = (f"sync.py watch --project {name} --apply" if row["open"]
                         else f"agent.py book --project {name} --action open --apply")
        row["risk"] = "" if row["open"] else "watch 只能盯已打开的表；先执行 command 打开目标工作簿，再跑 watch"
    else:
        row["mode"] = "unknown"
        row["risk"] = f"未知 intent：{intent}"

    if row["open"] and row["vbe_state"] in ("protected", "denied") and intent in ("pull", "push", "run", "watch"):
        extra = ("工程锁着：先解开 VBE 工程保护" if row["vbe_state"] == "protected"
                 else "信任中心未勾「信任对 VBA 工程对象模型的访问」：先改 Excel 信任设置")
        row["risk"] = (row["risk"] + "；" if row["risk"] else "") + extra

    if not path:
        errors.append(Failure(
            row["risk"], code=ErrorCode.TARGET_NOT_FOUND, stage=Stage.RESOLVE,
            hint="核对 mapping.json 的 search_dir 与 prefix；文件名必须是「prefix V<数字>.xlsm」",
            suggested=entry_command("sync.py", "list", "--json")))

    human.append(f"计划：{name} / {intent}")
    human.append(f"  表格：{row['workbook'] or '（未解析）'}")
    state_text = "已打开" if row["open"] else "未打开"
    if row["open"] and row["saved"] is False:
        state_text += " / 未保存"
    if row["vbe_state"]:
        state_text += f" / VBE={row['vbe_state']}"
    human.append(f"  当前状态：{state_text}")
    human.append(f"  推荐模式：{row['mode']}  needs_open={row['needs_open']}  will_save={row['will_save']}")
    human.append(f"  命令：{row['command'] or '（无可用命令）'}")
    if row["risk"]:
        warnings.append(f"  注意：{row['risk']}")

    stats = {"project": name, "intent": intent, "mode": row["mode"],
             "needs_open": row["needs_open"], "will_save": row["will_save"],
             "command": row["command"], "results": [row],
             "timings": ctx.stage_timings()}
    exit_code = 0 if row["ok"] else 1
    return Outcome(action="sync.plan", ok=row["ok"], details=[], errors=errors,
                   stats=stats, human=human, warnings=warnings,
                   exit_code=exit_code, elapsed_ms=ctx.elapsed_ms())


def pull(ctx) -> Outcome:
    """VBE→src：表已打开就借会话抓（含未保存改动），没开走隐藏实例读磁盘；默认预览，--apply 才写 src。"""
    cfg = ctx.config
    args = ctx.args
    dry = not bool(getattr(args, "apply", False))
    backup = not bool(getattr(args, "no_backup", False))
    defaults = cfg["defaults"]
    projects = cfg["projects"]
    only = getattr(args, "project", None)
    if only:
        proj = _find_project(cfg, only)
        projects = [proj]

    notes = Notes("抓取提醒")
    human: list[str] = []
    warnings: list[str] = []
    errors: list[Failure] = []
    rows: list[dict] = []
    exported = 0
    failed = 0

    session = ExcelSession()
    try:
        for proj in projects:
            with ctx.progress.stage(f"抓取 {proj['name']}"):
                path, reason = resolve_target(proj)
                human.append("")
                human.append(f"=== 抓取 {proj['name']} ===")
                if not path:
                    warnings.append(f"跳过：{reason}")
                    failed += 1
                    rows.append({"name": proj["name"], "ok": False, "reason": reason, "mode": ""})
                    errors.append(Failure(
                        reason, code=ErrorCode.TARGET_NOT_FOUND, stage=Stage.RESOLVE,
                        hint="核对 mapping.json 的 search_dir 与 prefix；文件名必须是「prefix V<数字>.xlsm」"))
                    continue
                human.append(f"  表格：{path}（{reason}）")

                out_dir = abs_path(proj["src"])
                human.append(f"  输出：{out_dir}")

                wb = None
                owned_copy = False  # True=本次隐藏实例开的磁盘副本，用完要关；False=借用户会话，不关窗
                mode = ""
                try:
                    if dry:
                        # 预览只做 ROT 轻量探测：不附着、不起新实例、不读表不写 src。
                        _xl, open_wb = find_excel_with_workbook(Path(path))
                        mode = "open" if open_wb is not None else "disk"
                        # dry 也要解析子集：预览即契约，先把将保留占位的公共模块范围摆出来。
                        subset, reason = _resolve_common_subset(proj, defaults)
                        human.append("  [dry] 将抓取该工作簿全部模块")
                        human.append(f"    {_scope_human(subset, reason, action='保留占位')}")
                        exported += 1
                        rows.append({"name": proj["name"], "ok": True, "dry": True,
                                     "mode": mode, **_scope_payload(subset)})
                        continue

                    wb, mode = _attach_pull_session(proj, Path(path))
                    if mode == "open":
                        # 已打开会话：抓内存里的最新代码（含未保存改动），不另开副本、不关窗。
                        if backup:
                            backups = abs_path(defaults.get("backup_dir", ".backup"))
                            dest = stamp_dir(backups, "-抓取前") / Path(wb.FullName).name
                            wb.SaveCopyAs(str(dest))
                            human.append(f"备份：{dest}")
                        human.append(f"是否已保存：{bool(wb.Saved)}")
                    else:
                        wb = excel_gateway.open_disk_workbook(
                            session.app(), Path(path), read_only=True)
                        owned_copy = True
                        if not has_vbproject_access(wb):
                            state, detail = vbe_state(wb)
                            warnings.append(f"「{proj['name']}」跳过：{detail}；{vbe_state_hint(state)}")
                            failed += 1
                            rows.append({"name": proj["name"], "ok": False, "reason": detail, "mode": mode})
                            # 旧版 protected/denied 两态无独立登记码；权限类失败统一归 VBE_PROTECTED。
                            errors.append(Failure(detail, code=ErrorCode.VBE_PROTECTED, stage=Stage.VBE))
                            continue

                    common_stems = _common_subset_stems(proj, defaults)
                    n = _export_workbook(wb, out_dir, dry, common_stems, human)
                    _write_placeholder_shells(proj, defaults, dry, human, warnings)
                    _verify_export_set(wb, out_dir, common_stems, human, warnings, notes)
                    human.append(f"  完成：抓取 {n} 个模块（{len(common_stems)} 个公共库模块保留占位，模式：{mode}）")
                    exported += 1
                    rows.append({"name": proj["name"], "ok": True, "modules": n, "mode": mode})
                except BaseFailure as e:
                    warnings.append(f"{proj['name']} 抓取失败：{e}")
                    failed += 1
                    rows.append({"name": proj["name"], "ok": False, "reason": str(e), "mode": mode})
                    errors.append(to_contract_failure(e))
                except Exception as e:  # 单表失败不中断整批，与 1.5.6 一致
                    warnings.append(f"{proj['name']} 抓取失败：{e}")
                    failed += 1
                    rows.append({"name": proj["name"], "ok": False, "reason": str(e), "mode": mode})
                    errors.append(Failure(str(e), code=ErrorCode.UNEXPECTED, stage=Stage.PULL))
                finally:
                    if wb is not None and owned_copy:
                        try:
                            wb.Close(SaveChanges=False)
                        except Exception:
                            pass
    finally:
        session.close()

    human.append("")
    human.append(f"抓取完成：成功 {exported} / 失败 {failed} / 共 {len(projects)} 个项目")
    if failed == 0:
        human.append("下一步：检查 src/ 下的文件（注意中文是否正常），然后 git add / commit")
    human.extend(notes.render("抓取提醒"))

    stats = {"projects": len(projects), "exported": exported, "failed": failed,
             "dry": dry, "results": rows, "timings": ctx.stage_timings()}
    exit_code = 0 if failed == 0 else 1
    return Outcome(action="sync.pull", ok=failed == 0, details=[], errors=errors,
                   stats=stats, human=human, warnings=warnings,
                   exit_code=exit_code, elapsed_ms=ctx.elapsed_ms())


def pull_open(ctx) -> Outcome:
    """VBE→src：从已打开的 Excel 抓（含未保存改动）；不保存工作簿。"""
    cfg = ctx.config
    args = ctx.args
    dry = not bool(getattr(args, "apply", False))
    backup = not bool(getattr(args, "no_backup", False))
    defaults = cfg["defaults"]

    notes = Notes("抓取提醒")
    human: list[str] = []
    warnings: list[str] = []
    errors: list[Failure] = []
    rows: list[dict] = []

    try:
        with ctx.progress.stage("连接已打开工作簿"):
            proj = _find_project(cfg, args.project)
            wb = _connect_open_workbook(proj, "抓取")

        saved_before = bool(wb.Saved)
        human.append(f"是否已保存：{saved_before}")
        out_dir = abs_path(proj["src"])

        if backup and not dry:
            backups = abs_path(defaults.get("backup_dir", ".backup"))
            dest = stamp_dir(backups, "-抓取前") / Path(wb.FullName).name
            wb.SaveCopyAs(str(dest))
            human.append(f"备份：{dest}")

        human.append(f"输出：{out_dir}")
        if dry:
            # dry 同样先解析子集：预览里就要能看到将保留占位的公共模块范围与未知模块告警。
            subset, reason = _resolve_common_subset(proj, defaults)
            human.append("  [dry] 将抓取该工作簿全部模块")
            human.append(f"    {_scope_human(subset, reason, action='保留占位')}")
            rows.append({"name": args.project, "ok": True, "dry": True,
                         **_scope_payload(subset)})
            stats = {"project": args.project, "dry": True, "modules": 0, "saved": False,
                     "workbook_saved": saved_before, "wrote_src": False,
                     **_scope_payload(subset),
                     "results": rows, "timings": ctx.stage_timings()}
            return Outcome(action="sync.pull-open", ok=True, details=[], errors=errors,
                           stats=stats, human=human, warnings=warnings,
                           exit_code=0, elapsed_ms=ctx.elapsed_ms())

        common_stems = _common_subset_stems(proj, defaults)
        n = _export_workbook(wb, out_dir, dry, common_stems, human)
        _write_placeholder_shells(proj, defaults, dry, human, warnings)
        _verify_export_set(wb, out_dir, common_stems, human, warnings, notes)
        human.append(f"完成：抓取 {n} 个模块（{len(common_stems)} 个公共库模块保留占位）")
        human.append("")
        human.append("当前工作簿没有被修改，也没有被保存。")
        human.append("下一步：检查 src 下的文件与中文，然后 git add / commit")
        human.extend(notes.render("抓取提醒"))
        rows.append({"name": args.project, "ok": True, "modules": n})
        stats = {"project": args.project, "dry": False, "modules": n, "saved": False,
                 "workbook_saved": saved_before, "wrote_src": True,
                 "results": rows, "timings": ctx.stage_timings()}
    except BaseFailure as e:
        errors.append(to_contract_failure(e))
        stats = {"project": args.project, "dry": dry, "results": rows,
                 "timings": ctx.stage_timings()}

    exit_code = 0 if not errors else 1
    return Outcome(action="sync.pull-open", ok=not errors, details=[], errors=errors,
                   stats=stats, human=human, warnings=warnings,
                   exit_code=exit_code, elapsed_ms=ctx.elapsed_ms())


def _finish(action: str, rows: list[dict], errors: list[Failure], human: list[str],
            warnings: list[str], notes: Notes, ctx) -> Outcome:
    """vbamod 收口：汇总计数、渲染提醒、生成 Outcome。

    结果集按 1.5.6 旧形状留在 ``stats.results``；``details`` 不承担命令结果行。
    """
    human.extend(notes.render("核对模块清单提醒"))
    stats = {"examined": len(rows),
             "changed": sum(1 for r in rows if r.get("same") is False),
             "unverifiable": sum(1 for r in rows if not r.get("compared")),
             "results": rows,
             "timings": ctx.stage_timings()}
    exit_code = 0 if not errors else 1
    return Outcome(action=action, ok=not errors, details=[], errors=errors,
                   stats=stats, human=human, warnings=warnings,
                   exit_code=exit_code, elapsed_ms=ctx.elapsed_ms())


SPECS = (
    CommandSpec(
        verb="vbamod", entry="sync", handler="features.sync_read:vbamod",
        mutation=MutationMode.NONE, project=ProjectMode.OPTIONAL,
        supports_json=True, exit_codes=(0, 1, 2),
        summary="核对各工作簿的模块清单（和最近一次备份比）",
        ref="references/sync-commands.md#vbamod",
        tests=("contract.vbamod.success", "contract.vbamod.not_found",
               "contract.vbamod.usage", "contract.vbamod.json"),
    ),
    CommandSpec(
        verb="plan", entry="sync", handler="features.sync_read:plan",
        mutation=MutationMode.NONE, project=ProjectMode.REQUIRED,
        args=(ArgSpec("--intent", kind="choice",
                      choices=("pull", "push", "run", "watch"), default="push",
                      help="打算做什么：pull / push / run / watch（默认 push）"),),
        supports_json=True, exit_codes=(0, 1, 2),
        summary="只读推荐：按当前状态选 open/disk/isolate/open_required/unavailable 模式",
        ref="references/sync-commands.md#plan",
        tests=("contract.plan.success", "contract.plan.target_missing",
               "contract.plan.usage", "contract.plan.json"),
    ),
    CommandSpec(
        verb="pull", entry="sync", handler="features.sync_read:pull",
        mutation=MutationMode.PREVIEW, project=ProjectMode.ANY,
        args=(ArgSpec("--no-backup", kind="flag", help="不写 .backup 直接改 src"),),
        supports_json=True, exit_codes=(0, 1, 2),
        summary="VBE→src：工作簿已打开时借用会话抓取，未打开时从磁盘抓取",
        ref="references/sync-commands.md#pull",
        tests=("contract.pull.success", "contract.pull.target_missing",
               "contract.pull.usage", "contract.pull.json"),
    ),
    CommandSpec(
        verb="pull-open", entry="sync", handler="features.sync_read:pull_open",
        mutation=MutationMode.PREVIEW, project=ProjectMode.REQUIRED,
        args=(ArgSpec("--no-backup", kind="flag", help="不写 .backup 直接改 src"),),
        supports_json=True, exit_codes=(0, 1, 2),
        summary="VBE→src：从已打开的 Excel 抓（含未保存改动）",
        ref="references/sync-commands.md#pull-open",
        tests=("contract.pull-open.success", "contract.pull-open.target_missing",
               "contract.pull-open.usage", "contract.pull-open.json"),
    ),
)
