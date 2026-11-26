# -*- coding: utf-8 -*-
"""sync_push 切片：init / push / shells 三个写路径动词的入口薄层。

流程段已按职责拆到同包内聚模块：
  · ``sync_preflight`` —— 写回前源码预检（只读，不碰 Excel）；
  · ``sync_shells``    —— 公共库占位壳生成 / 清除；
  · ``sync_init``      —— mapping.json 与 src 骨架登记（含 --lab 生成器接缝）。

push 写回链固定顺序：预检 → 备份 → 写回 → 回读 →（失败按开关）回滚 → 保存。
机械动作全部走 toolkit.excel / toolkit.repo / toolkit.files / toolkit.vba_parse；
COM 附着/开表只走 adapters.excel_gateway。

边界：
  · handler 只返回 Outcome，不 print、不自带 JSON；
  · 不回壳别的 CLI；--lab 生成测试表走 ``_run_lab_generator`` 接缝，便于测试替换。
"""
from __future__ import annotations

import re
import tempfile
from pathlib import Path
from typing import Any

from adapters import excel_gateway
from toolkit.base import (Failure as BaseFailure, abs_path, entry_command,
                          read_source_text)
from toolkit.check import check_module_set
from toolkit.contract import (ErrorCode, Failure, MutationMode, Outcome, Stage,
                               to_contract_failure)
from toolkit.excel import (component_snapshot, ensure_vbe_access, find_component,
                           has_vbproject_access, import_component, remove_module,
                           replace_module_code, vbe_state)
from toolkit.files import (backup_copy, gbk_import_copy, prune_backups, remove_tree)
from toolkit.repo import (DOC_MODULE_RE, IMPORTABLE, SOURCE_EXTS, TYPE_FOR_EXT,
                          VBXT_CLASS, VBXT_DOCUMENT, VBXT_STANDARD, is_shell_file,
                          module_name_of, resolve_target, source_files)
from toolkit.registry import ArgSpec, CommandSpec, ProjectMode
from toolkit.vba_parse import (drop_attribute_lines, drop_trailing_placeholder,
                               fold_lines, norm_code_lines, strip_export_header)

from features.common_subset import (  # noqa: E402
    close_owned as _close_owned,
    find_project as _find_project,
    resolve_common_subset as _resolve_common_subset_shared,
    restore_module_code as _restore_module_code,
    scope_human as _scope_human,
    scope_payload as _scope_payload,
    snapshot_module_code as _snapshot_module_code,
    warn_unknown_modules as _warn_unknown_modules,
)
from features.sync_init import (  # noqa: E402
    _default_config_path,
    _is_within,
    _run_lab_generator,
    init,
)
from features.sync_preflight import (  # noqa: E402
    _preflight_or_fail,
    _source_dirs,
    _source_preflight,
)
from features.sync_shells import (  # noqa: E402
    _write_shells,
    shells,
)


def _common_subset(proj: dict, defaults: dict) -> tuple[set[str] | None, str]:
    """这个项目要写回哪些公共模块；None 表示整个 common。"""
    subset, reason, ignored = _resolve_common_subset_shared(proj, defaults)
    _warn_unknown_modules(ignored)
    return subset, reason


# ---------------------------------------------------------------------------
# 备份与自动清理
# ---------------------------------------------------------------------------
def _backup_keep_settings(cfg: dict) -> dict:
    """备份保留口径（全部来自 mapping.json 的 defaults）。"""
    d = cfg.get("defaults", {})
    protect = d.get("backup_protect", [])
    if isinstance(protect, str):
        protect = [protect]
    return {
        "keep_recent": int(d.get("backup_keep", 10)),
        "keep_daily": int(d.get("backup_keep_daily", 1)),
        "min_keep": int(d.get("backup_min_keep", 3)),
        "auto": bool(d.get("backup_auto_clean", True)),
        "protect": list(protect or []),
    }


def _backup_workbook(cfg: dict, wb, target: Path) -> Path | None:
    """整表备份进配置 backup_dir；备份后按口径自动清理。清理失败只告警不影响写回。"""
    backup_root = abs_path(cfg["defaults"].get("backup_dir", ".backup"))
    dest = backup_copy(target, backup_root)
    settings = _backup_keep_settings(cfg)
    if dest is not None and settings["auto"]:
        try:
            prune_backups(backup_root, keep_recent=settings["keep_recent"],
                          keep_daily=settings["keep_daily"], min_keep=settings["min_keep"],
                          dry_run=False, protect=settings["protect"])
        except Exception:
            # 清理失败绝不能影响写回本身
            pass
    return dest


# ---------------------------------------------------------------------------
# VBE 机械动作（写回 / 回读 / 回滚）
# ---------------------------------------------------------------------------
def _code_body(text: str, vbe: bool = False) -> list[str]:
    """比对用正文：剥过程级 Attribute；VBE 侧再抹掉尾部 `()` 占位。"""
    text = drop_attribute_lines(text)
    if vbe:
        text = drop_trailing_placeholder(text)
    return [ln for ln in norm_code_lines(text) if ln.strip()]


def _has_trailing_placeholder(text: str) -> bool:
    """VBE 正文最末尾是否多了一行 `()`；中间游离 `()` 不算。"""
    lines = text.splitlines()
    while lines and not lines[-1].strip():
        lines.pop()
    return bool(lines and lines[-1].strip() == "()")


def _postcheck_sync(vbproj, written: dict[str, Path]) -> list[str]:
    """写回后逐模块回读比对；失败返回带行数的模块明细（空=全部一致）。"""
    bad: list[str] = []
    for name, path in written.items():
        comp = find_component(vbproj, name)
        if comp is None:
            bad.append(f"{name}（工程里找不到）")
            continue
        try:
            cm = comp.CodeModule
            got_text = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
            want_text = read_source_text(path)
            _ = _has_trailing_placeholder(got_text)
        except Exception as e:  # 回读/读源码失败也算校验失败
            bad.append(f"{name}（回读失败：{e}）")
            continue
        want_body = _code_body(want_text)
        got_body = _code_body(got_text, vbe=True)
        want = fold_lines(want_body)
        got = fold_lines(got_body)
        if want != got:
            first = next((i for i, (a, b) in enumerate(zip(want, got), 1) if a != b),
                         min(len(want), len(got)) + 1)
            bad.append(f"{name}（源码 {len(want_body)} 行 / VBE {len(got_body)} 行；"
                       f"首个差异第 {first} 行）")
    return bad


def _overwrite_document_module(vbproj, name: str, code: str) -> bool:
    """文档模块固定存在，不能 Remove，只能用 CodeModule 写回。"""
    comp = find_component(vbproj, name)
    if comp is None:
        return False
    if comp.Type != VBXT_DOCUMENT:
        return False
    replace_module_code(comp, code)
    return True


def _remove_importable(vbproj, name: str) -> bool:
    """删除旧模块（标准/类/窗体）；文档模块与非导入类型一律不碰。"""
    comp = find_component(vbproj, name)
    if comp is None or comp.Type not in IMPORTABLE:
        return False
    return remove_module(vbproj, name)


def _iter_source_files(folder: Path, used: set[str]) -> list[Path]:
    """递归收集源码文件；同名模块只取排序靠前的那个。"""
    result: list[Path] = []
    for f in source_files(folder):
        n = module_name_of(f)
        if n in used:
            continue
        used.add(n)
        result.append(f)
    return result


def _import_folder(vbproj, folder: Path, doc_modules: bool,
                   only: set[str] | None, tmp_parent: Path | None
                   ) -> tuple[int, int, dict[str, Path]]:
    """把文件夹下（含子目录）的 .bas/.cls/.frm 写回工程；返回 (写回数, 跳过数, 已写回)。

    旧模块删不掉或 Import 拿不到同名组件时直接抛 Failure（POSTCHECK_FAILED），
    由调用方在 Save 之前中止；不允许静默跳过导致模块从工程里消失。
    """
    if not folder.is_dir():
        return 0, 0, {}

    imported = skipped = 0
    written: dict[str, Path] = {}
    used: set[str] = set()
    tmp_dir: Path | None = None
    try:
        if tmp_parent is not None and Path(tmp_parent).is_dir():
            tmp_dir = Path(tempfile.mkdtemp(prefix=".vba-import-", dir=str(tmp_parent)))
        else:
            tmp_dir = Path(tempfile.mkdtemp(prefix="vba-import-"))

        for f in _iter_source_files(folder, used):
            if only is not None and module_name_of(f) not in only:
                continue
            # 公共库占位壳不能当模块写回，否则会把真公共模块覆盖成空壳
            if is_shell_file(f):
                skipped += 1
                continue
            name = module_name_of(f)
            comp = find_component(vbproj, name)
            ctype = comp.Type if comp is not None else None

            is_doc = (ctype == VBXT_DOCUMENT or (
                ctype is None and DOC_MODULE_RE.match(name)
                and name.lower().startswith(("thisworkbook", "sheet", "chart"))))
            if is_doc:
                if not doc_modules:
                    skipped += 1
                    continue
                if _overwrite_document_module(
                        vbproj, name, strip_export_header(read_source_text(f))):
                    imported += 1
                    written[name] = f
                else:
                    skipped += 1
                continue

            # 写回策略：.bas/.cls 同名同类型 → 原地改正文，不 Remove 不 Import，避免补序号；
            # .frm / 新模块 / 类型不符 → Import 落名。
            in_place = (ctype in (VBXT_STANDARD, VBXT_CLASS)
                        and TYPE_FOR_EXT.get(f.suffix.lower()) == ctype)
            if in_place:
                replace_module_code(comp, strip_export_header(read_source_text(f)))
                imported += 1
                written[name] = f
                continue
            old_present = comp is not None
            if not _remove_importable(vbproj, name) and old_present:
                # 旧模块删不掉：不能再 Import（否则 VBE 补序号造出 名字1 后缀）。
                # 本次写回未落地必须硬拦下、不保存，不能把旧正文/缺模块状态落盘。
                raise Failure(
                    f"写回被拦下：模块「{name}」旧组件删不掉，无法安全重导"
                    "（避免生成后缀模块）；本次未保存",
                    code=ErrorCode.POSTCHECK_FAILED, stage="postcheck",
                    suggested="在 VBE 里手工确认该模块状态后重跑；不要手工保存这次结果")
            new_comp = import_component(vbproj, gbk_import_copy(f, tmp_dir), name)
            if new_comp is None:
                # 旧模块已移除（或本应新建）但 Import 没拿到同名组件：工程里已少这个
                # 模块，必须拦下，绝不能走到 Save。
                raise Failure(
                    f"写回被拦下：模块「{name}」Import 后拿不到同名组件"
                    + ("（旧模块已移除）" if old_present else "") + "；本次未保存",
                    code=ErrorCode.POSTCHECK_FAILED, stage="postcheck",
                    suggested="检查源码编码/文件占用与 VBE 工程状态；修好后重跑，不要手工保存这次结果")
            imported += 1
            written[name] = f
    finally:
        if tmp_dir is not None:
            remove_tree(tmp_dir)

    return imported, skipped, written


def _preflight_sheet_modules(wb) -> None:
    """写回前硬拦：表页数 ≠ 文档模块数，或有文档模块没有表页指向它。"""
    try:
        sheets = [str(wb.Sheets(i).CodeName) for i in range(1, wb.Sheets.Count + 1)]
        docs = [c.Name for c in wb.VBProject.VBComponents if c.Type == VBXT_DOCUMENT]
    except Exception:
        # 读不到按未校验继续（旧版口径）
        return
    docs = [d for d in docs if d.lower() != "thisworkbook"]
    broken: list[str] = []
    if len(sheets) != len(docs):
        broken.append(f"{len(sheets)} 个表页 / {len(docs)} 个文档模块")
    for n in docs:
        if n not in sheets:
            broken.append(f"无宿主 {n}")
    for c in sheets:
        if c not in docs:
            broken.append(f"无代码 {c}")
    if broken:
        raise Failure(
            "写回已拦下：表页与文档模块不对应（H19）：" + "；".join(broken)
            + "。这种工程写回只会把空壳一起带下去，文档模块又删不掉——先换干净的 vbaProject.bin 再重植",
            code=ErrorCode.SOURCE_PREFLIGHT, stage="preflight")


# ---------------------------------------------------------------------------
# 引用依赖扫描（早期绑定检测；只提醒、不改退出码）
# ---------------------------------------------------------------------------
_REF_RULES: dict[str, str] = {
    "ADODB.": "Microsoft ActiveX Data Objects 6.1 Library",
    "ADOX.": "Microsoft ADO Ext. 6.0 for DDL and Security",
    "DAO.": "Microsoft DAO 3.6 Object Library",
    "Scripting.": "Microsoft Scripting Runtime",
    "MSForms.": "Microsoft Forms 2.0 Object Library",
    "Shell32.": "Microsoft Shell Controls And Automation",
    "SHDocVw.": "Microsoft Internet Controls",
    "MSComctlLib.": "Microsoft Windows Common Controls 6.0 (SP6)",
}
_EARLY_BOUND_RE = {
    pre: re.compile(r"\bAs\s+" + re.escape(pre), re.IGNORECASE) for pre in _REF_RULES
}


def _existing_references(vbproj) -> list[str]:
    """读工程当前已勾选的引用描述；工程被锁/读不到时返回空列表。"""
    descs: list[str] = []
    try:
        refs = vbproj.References
        for i in range(1, refs.Count + 1):
            try:
                descs.append(refs.Item(i).Description or "")
            except Exception:
                continue
    except Exception:
        return []
    return descs


def _ref_key(desc: str) -> str:
    """引用描述归一化关键词，兼容 2.8 / 6.1 / 不同语言版本。"""
    d = desc.lower()
    for token in ("activex data objects", "scripting runtime", "forms 2.0",
                  "shell controls", "internet controls", "common controls",
                  "dao 3.6", "ado ext"):
        if token in d:
            return token
    return d.strip()


def _missing_references(vbproj, need: dict[str, int]) -> list[str]:
    """对比源码需要与工程已勾选，返回缺少的引用描述列表；读不到引用不误报。"""
    have = {_ref_key(d) for d in _existing_references(vbproj)}
    if not have:
        return []
    return [desc for desc in need if _ref_key(desc) not in have]


def _scan_required_refs(folders: list[Path], only: set[str] | None) -> dict[str, int]:
    """静态扫描源码，返回 {需要勾选的引用: 早期绑定出现次数}。"""
    need: dict[str, int] = {}
    for folder in folders:
        if not folder.is_dir():
            continue
        for f in folder.rglob("*"):
            if not f.is_file() or f.suffix.lower() not in SOURCE_EXTS:
                continue
            if only is not None and module_name_of(f) not in only:
                continue
            text = read_source_text(f)
            for pre, pat in _EARLY_BOUND_RE.items():
                n = len(pat.findall(text))
                if n:
                    desc = _REF_RULES[pre]
                    need[desc] = need.get(desc, 0) + n
    return need


def _report_references(vbproj, folder: Path, only: set[str] | None, label: str) -> list[str]:
    """写回后扫描缺失引用，返回软提醒；只提醒、不改退出码。"""
    need = _scan_required_refs([folder], only)
    if not need:
        return []
    missing = _missing_references(vbproj, need)
    if not missing:
        return []
    notes = [f"[引用] {label} 缺少以下引用，请手工勾选后再用宏："]
    for desc in missing:
        notes.append(f"  - {desc}（早期绑定 {need[desc]} 处）")
    notes.append("  勾选方法：VBE → 工具 → 引用 → 在列表里搜索该名称 → 勾上 → 确定")
    return notes


def _writeback(wb, proj: dict, defaults: dict, *, doc_modules: bool,
               rollback_on_fail: bool) -> tuple[list[str], list[str]]:
    """把 common + 项目源码写回一个工作簿。

    返回 ``(bad, notes)``：bad 为硬拦截明细（少模块 / 回读缺失 / 正文回读不一致，空=通过）；
    notes 为缺失引用、多出非本次模块等软提醒。bad 非空时调用方必须在 Save 之前中止。
    rollback_on_fail=True 时，postcheck 失败按写回前内存正文复原本次写过的模块；
    默认关闭：失败不保存、也不自动改内存。
    """
    vbproj = wb.VBProject
    _preflight_sheet_modules(wb)
    notes: list[str] = []

    common_dir = abs_path(defaults.get("common_dir", "src/common"))
    subset, _reason = _common_subset(proj, defaults)

    before = component_snapshot(vbproj)
    code_before = _snapshot_module_code(vbproj) if rollback_on_fail else {}
    written_all: dict[str, Path] = {}

    tmp_parent: Path | None = None
    full_name = getattr(wb, "FullName", "") or ""
    if full_name:
        tmp_parent = Path(full_name).parent

    for folder in _source_dirs(proj, defaults):
        if not folder.is_dir():
            continue
        only = subset if folder == common_dir else None
        _imported, _skipped, written = _import_folder(
            vbproj, folder, doc_modules, only, tmp_parent)
        written_all.update(written)
        # 写完立刻比对引用，只针对本次实际写回的模块（不自动改引用设置）
        notes.extend(_report_references(vbproj, folder, only, folder.name))

    after = component_snapshot(vbproj)
    diff = check_module_set(before, after, list(written_all))
    blocking: list[str] = []
    if diff is not None:
        # 少模块是硬伤：写回前有的模块写回后不见了，落盘即永久丢失，必须拦下
        if diff.lost:
            blocking.append("写回后少了模块：" + "、".join(diff.lost))
        # 多出且不在本次写回名单里的模块（含疑似抢名后缀）：只作软提醒，不阻断
        stray_added = [n for n in diff.added if n not in written_all]
        if stray_added:
            notes.append("写回后多出非本次写回模块：" + "、".join(stray_added)
                          + "（仅提醒，未纳入本次写回）")
        if diff.suffix:
            pairs = "、".join(f"{n}（原名 {b}）" for n, b in diff.suffix)
            notes.append(f"疑似 Import 抢名残留：{pairs}（仅提醒）")

    # 回读覆盖：本次写回的每个模块都必须还在工程里，数量/对账不符一并拦下
    if after is not None:
        missing_written = [n for n in written_all if n not in after]
        if missing_written:
            blocking.append("回读不到本次写回模块：" + "、".join(missing_written))

    bad = blocking + _postcheck_sync(vbproj, written_all)
    if bad:
        if rollback_on_fail and code_before:
            snap = {n: code_before.get(n, "") for n in written_all}
            _restored, failed = _restore_module_code(vbproj, snap)
            if failed:
                bad.append("回滚未完成：" + "；".join(failed))
        return bad, notes
    return [], notes


# ---------------------------------------------------------------------------
# 附着 / 开表（唯一 COM 边界：adapters.excel_gateway）
# ---------------------------------------------------------------------------
def _attach_or_open(proj: dict, target: Path, *, auto_open: bool) -> tuple[Any, Any, str]:
    """连上目标工作簿；返回 (wb, owned_xl, mode)。

    已打开就借会话写回；表没开且允许自动开表时，起隐藏实例打开磁盘文件。
    只有 NO_EXCEL / TARGET_NOT_OPEN 才允许转磁盘；COM_CONTEXT、工程锁等原样上报。

    有意保留：错误码/排障提示按入口不同，合并需同步改错误契约（见 1.6.x 统一化方案）。
    """
    # push 自己的附着失败口径：不引用 run --isolate；按是否允许自动开表分流提示
    # 重试命令去掉 --open-only（允许工具自己开隐藏实例），入口按本包位置解析为绝对路径
    retry_cmd = entry_command("sync.py", "push", "--project", proj["name"], "--apply")
    if auto_open:
        attach_hint = ("push 未检测到正在运行的 Excel：工具会尝试自己开隐藏实例写回；"
                       "仍失败请检查 Excel 安装、文件占用与受保护视图")
        attach_suggested = "检查 Excel 弹窗/受保护视图；仍不行关掉表后重跑 push 从磁盘模式写回"
    else:
        attach_hint = ("已指定 --open-only：不会自动开表；去掉 --open-only 让工具自动开隐藏实例写回，"
                       "或先在 Excel 里打开目标表")
        attach_suggested = retry_cmd
    try:
        xl = excel_gateway.attach_running_excel(
            Path(target), target_hint=attach_hint, target_suggested=attach_suggested)
        wb, _opened = excel_gateway.find_open_workbook(xl, Path(target))
        if wb is None:
            if auto_open:
                wb_hint = "批量写回会自己开隐藏实例；若仍失败，检查文件占用/受保护视图"
                wb_suggested = "批量写回会自己开隐藏实例；若仍失败，检查文件占用/受保护视图"
            else:
                wb_hint = ("已指定 --open-only：不会自动开表，必须先在 Excel 里打开目标表；"
                           "或者去掉 --open-only 让工具开隐藏实例")
                wb_suggested = retry_cmd
            raise Failure(
                f"目标表没有打开：{target}",
                code=ErrorCode.TARGET_NOT_OPEN, stage=Stage.ATTACH,
                hint=wb_hint, suggested=wb_suggested,
            )
        return wb, None, "open"
    except BaseFailure as e:
        # --open-only 且没有正在运行的 Excel：语义是"不自动开表、表又没开"，
        # 归 TARGET_NOT_OPEN；COM_CONTEXT / VBE_PROTECTED 等其它码原样上抛。
        if not auto_open and e.code == "NO_EXCEL":
            raise Failure(
                f"未自动开表且目标表未在 Excel 中打开：{target}",
                code=ErrorCode.TARGET_NOT_OPEN, stage=Stage.ATTACH,
                hint=attach_hint, suggested=retry_cmd,
            ) from e
        if (not auto_open) or e.code not in ("NO_EXCEL", "TARGET_NOT_OPEN"):
            raise
        owned_xl = excel_gateway.new_excel(hidden=True)
        wb = excel_gateway.open_disk_workbook(owned_xl, Path(target), read_only=False)
        # 磁盘模式下先卡工程可读这一关；非交互不等人解锁
        if not ensure_vbe_access(wb, proj["name"], interactive=False):
            _state, detail = vbe_state(wb)
            # 1.5.6 错误码口径：VBA 权限类失败统一归 VBE_PROTECTED
            raise Failure(
                f"「{proj['name']}」的 VBA 工程读不了：{detail}",
                code=ErrorCode.VBE_PROTECTED, stage=Stage.VBE,
                suggested="先在 Excel 里解开工程保护/勾选 AccessVBOM，再重跑写回",
            )
        return wb, owned_xl, "disk"


# ---------------------------------------------------------------------------
# push：src→VBE 写回（预检 → 备份 → 写回 → 回读 → 保存）
# ---------------------------------------------------------------------------
def push(ctx) -> Outcome:
    args = ctx.args
    cfg = ctx.config
    defaults = cfg["defaults"]
    dry = not bool(getattr(args, "apply", False))
    doc = (bool(getattr(args, "doc_modules", False))
           or bool(defaults.get("import_document_modules")))
    backup = not bool(getattr(args, "no_backup", False))
    save = not bool(getattr(args, "no_save", False))
    auto_open = not bool(getattr(args, "open_only", False))
    rollback = bool(getattr(args, "rollback_on_fail", False))
    all_flag = bool(getattr(args, "all_project", False))
    name = getattr(args, "project", None)

    if all_flag:
        return _push_projects(ctx, cfg, defaults, doc, dry, backup, save, rollback)

    if not name:
        outcome = Outcome(action="sync.push", ok=False, exit_code=2)
        outcome.add_error(Failure("push 需要 --project <项目名> 或 --all-project",
                                 code=ErrorCode.USAGE, stage=Stage.CONFIG))
        return outcome

    try:
        proj = _find_project(cfg, name)
    except Failure as e:
        return Outcome(action="sync.push", ok=False, details=[], errors=[e],
                       exit_code=1, elapsed_ms=ctx.elapsed_ms())

    return _push_single(ctx, proj, cfg, defaults, doc, dry, backup, save,
                        auto_open, rollback)


def _push_single(ctx, proj, cfg, defaults, doc, dry, backup, save,
                 auto_open, rollback) -> Outcome:
    name = proj["name"]
    errors: list[Failure] = []
    human: list[str] = []
    warnings: list[str] = []

    with ctx.progress.stage("preflight"):
        target, reason = resolve_target(proj)
        try:
            warnings.extend(_preflight_or_fail(proj, defaults, doc))
        except Failure as e:
            errors.append(e)

    mods = len(source_files(abs_path(proj["src"]))) if proj.get("src") else 0

    if dry:
        # 预览容忍缺表：不连 Excel、不定位失败，只跑源码预检；目标解析不到记 "None"
        # （旧 cmd_single dry 口径）。TARGET_NOT_FOUND 只属于 apply 连接路径。
        # dry 同样先解析子集：预览即契约，先把本次将写回的公共模块范围摆出来。
        target_str = str(target)
        subset, reason = _common_subset(proj, defaults)
        scope = _scope_payload(subset)
        human.append(_scope_human(subset, reason))
        row = {"name": name, "ok": not errors, "target": target_str,
               "modules": mods, "dry": True, **scope}
        stats = {"project": name, "target": target_str,
                 "modules": mods, "dry": True, **scope, "results": [row]}
        return Outcome(action="sync.push", ok=not errors, details=[row],
                       errors=errors, stats=stats, human=human, warnings=warnings,
                       exit_code=1 if errors else 0, elapsed_ms=ctx.elapsed_ms())

    if target is None:
        errors.append(Failure(
            f"「{name}」没解析到 xlsm：{reason}",
            code=ErrorCode.TARGET_NOT_FOUND, stage=Stage.RESOLVE,
            hint="核对 mapping.json 的 search_dir 与 prefix；文件名必须是「prefix V<数字>.xlsm」",
            suggested=entry_command("sync.py", "list", "--json")))

    # 预检/定位失败：旧版在连 Excel、备份、写回之前就中止，这里同样不进入附着链
    if errors:
        row = {"name": name, "ok": False, "mode": "", "result": "失败",
               "modules": mods, "saved": False}
        stats = {"project": name, "mode": "", "result": "失败", "dry": False,
                 "saved": False, "backup": False, "backup_path": "",
                 "results": [row]}
        return Outcome(action="sync.push", ok=False, details=[row],
                       errors=errors, stats=stats, human=human, warnings=warnings,
                       exit_code=1, elapsed_ms=ctx.elapsed_ms())

    wb = None
    owned_xl = None
    mode = ""
    backup_path: Path | None = None
    saved_ok = False
    cleanup_problems: list[str] = []
    try:
        with ctx.progress.stage("attach"):
            wb, owned_xl, mode = _attach_or_open(proj, target, auto_open=auto_open)
        if backup:
            with ctx.progress.stage("backup"):
                backup_path = _backup_workbook(cfg, wb, target)
        with ctx.progress.stage("writeback"):
            bad, writeback_notes = _writeback(
                wb, proj, defaults, doc_modules=doc, rollback_on_fail=rollback)
            warnings.extend(writeback_notes)
        with ctx.progress.stage("readback"):
            if bad:
                raise Failure(
                    "写回后回读校验失败，未保存：" + "；".join(bad),
                    code=ErrorCode.POSTCHECK_FAILED, stage="postcheck",
                    suggested="检查源码编码/Attribute 行；修好后重跑，不要手工保存这次结果")
        # 写回后按口径还原一次应用状态（计算模式随工作簿一起保存）；失败明细并入 warnings
        try:
            from toolkit.app_guard import AppGuard
            guard = AppGuard(wb.Application)
            guard.restore()
            if guard.changes:
                warnings.append(f"已还原 Excel 状态：{guard.describe()}")
            for line in guard.errors[:3]:
                warnings.append(f"状态还原未完成：{line}")
        except Exception:
            pass
        if save:
            try:
                wb.Save()
                saved_ok = True
            except Exception as e:
                raise Failure(
                    f"写回成功，但保存失败：{e}；请回 Excel 手动 Ctrl+S",
                    code=ErrorCode.SAVE_FAILED, stage=Stage.SAVE,
                    suggested="检查文件是否只读、是否被网盘/Excel 占用，处理后重跑") from e
    except BaseFailure as e:
        errors.append(to_contract_failure(e))
    finally:
        cleanup_problems = _close_owned(wb, owned_xl)

    result_text = "写回完成" if not errors else "失败"
    row = {"name": name, "ok": not errors, "mode": mode, "result": result_text,
           "modules": mods, "saved": saved_ok}
    stats = {"project": name, "mode": mode, "result": result_text, "dry": False,
             "saved": saved_ok, "backup": bool(backup) and bool(backup_path),
             "backup_path": str(backup_path) if backup_path else "",
             "results": [row]}
    if cleanup_problems:
        warnings.append("隔离实例回收未完成：" + "；".join(cleanup_problems[:3]))
        stats["isolate_cleanup_problems"] = cleanup_problems
    else:
        stats["isolate_cleanup"] = "ok"
    return Outcome(action="sync.push", ok=not errors, details=[row],
                   errors=errors, stats=stats, human=human, warnings=warnings,
                   exit_code=1 if errors else 0, elapsed_ms=ctx.elapsed_ms())


def _push_projects(ctx, cfg, defaults, doc, dry, backup, save, rollback) -> Outcome:
    backup_root = abs_path(defaults.get("backup_dir", ".backup"))
    rows: list[dict[str, str]] = []
    errors: list[Failure] = []
    warnings: list[str] = []
    batch_cleanup: list[str] = []

    for proj in cfg["projects"]:
        path, reason = resolve_target(proj)
        if path is None:
            rows.append({"name": proj["name"], "status": f"跳过（{reason}）"})
            continue

        if dry:
            try:
                with ctx.progress.stage("preflight"):
                    warnings.extend(_preflight_or_fail(proj, defaults, doc))
                # 批量 dry 同样解析子集：行内带上将写回的公共模块范围。
                subset, _reason = _common_subset(proj, defaults)
                rows.append({"name": proj["name"], "status": "dry-run",
                             **_scope_payload(subset)})
            except Failure as e:
                rows.append({"name": proj["name"], "status": f"失败：{e}"})
                errors.append(e)
            continue

        wb = None
        owned_xl = None
        try:
            with ctx.progress.stage("preflight"):
                warnings.extend(_preflight_or_fail(proj, defaults, doc))
            with ctx.progress.stage("attach"):
                # 批量写回始终允许自动开隐藏实例（旧 cmd_all 口径）
                wb, owned_xl, _mode = _attach_or_open(proj, path, auto_open=True)
            if not has_vbproject_access(wb):
                _state, detail = vbe_state(wb)
                raise Failure(detail, code=ErrorCode.VBE_PROTECTED, stage=Stage.VBE,
                              suggested="解开工程保护/勾选 AccessVBOM 后重跑")
            if backup:
                with ctx.progress.stage("backup"):
                    backup_copy(path, backup_root)
            with ctx.progress.stage("writeback"):
                bad, writeback_notes = _writeback(
                    wb, proj, defaults, doc_modules=doc, rollback_on_fail=rollback)
                warnings.extend(writeback_notes)
            with ctx.progress.stage("readback"):
                if bad:
                    raise Failure(
                        "写回后回读校验失败，未保存：" + "；".join(bad),
                        code=ErrorCode.POSTCHECK_FAILED, stage="postcheck",
                        suggested="检查源码编码/Attribute 行；修好后重跑，不要手工保存这次结果")
            # 批量口径：逐张写回后固定保存（旧 cmd_all 不走 --no-save；保存失败落 UNEXPECTED）
            wb.Save()
            rows.append({"name": proj["name"], "status": "成功"})
        except BaseFailure as e:
            rows.append({"name": proj["name"], "status": f"失败：{e}"})
            errors.append(to_contract_failure(e))
        except Exception as e:  # 批量兜底：非 Failure 原样收口为 UNEXPECTED
            rows.append({"name": proj["name"], "status": f"失败：{e}"})
            errors.append(Failure(str(e), code=ErrorCode.UNEXPECTED, stage=Stage.PUSH))
        finally:
            batch_cleanup.extend(_close_owned(wb, owned_xl))

    ok = sum(1 for r in rows
             if r["status"] in ("成功", "dry-run") or r["status"].startswith("跳过"))
    stats = {"results": rows, "ok": ok, "total": len(rows), "dry": dry}
    if batch_cleanup:
        warnings.append("隔离实例回收未完成：" + "；".join(batch_cleanup[:3]))
        stats["isolate_cleanup_problems"] = batch_cleanup
    else:
        stats["isolate_cleanup"] = "ok"
    return Outcome(action="sync.push", ok=not errors, details=rows, errors=errors,
                   stats=stats, warnings=warnings,
                   exit_code=0 if not errors else 1,
                   elapsed_ms=ctx.elapsed_ms())


SPECS = (
    CommandSpec(
        verb="init", entry="sync", handler="features.sync_push:init",
        mutation=MutationMode.PREVIEW, project=ProjectMode.REQUIRED,
        supports_json=True, exit_codes=(0, 1, 2),
        args=(
            ArgSpec("--dir", kind="str", help="表所在目录（--lab 时默认 .lab）"),
            ArgSpec("--file", kind="path",
                    help="精确登记一个 .xlsm / .xlsb（给定时不再按版本号选表）"),
            ArgSpec("--prefix", kind="str", help="文件名基名（默认同项目名；给 --file 时取文件基名）"),
            ArgSpec("--src", kind="str", help="源码子目录（默认 src/projects/<项目名>）"),
            ArgSpec("--config", kind="str", help="写到指定的 mapping.json（默认自动定位）"),
            ArgSpec("--allow-outside", kind="flag", dest="allow_outside",
                    help="允许登记当前配置工作区之外的文件（菜单拖入时使用）"),
            ArgSpec("--lab", kind="flag", help="同时生成空白测试表（需要 Excel + pywin32）"),
        ),
        summary="生成/更新 mapping.json 与 src 骨架（默认只预览）",
        ref="references/sync-commands.md#init",
        tests=("contract.init.preview", "contract.init.apply"),
    ),
    CommandSpec(
        verb="push", entry="sync", handler="features.sync_push:push",
        mutation=MutationMode.REQUIRED, project=ProjectMode.ANY,
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
        summary="src→VBE：写回工作簿（工作簿未打开时自动开隐藏实例写回）",
        ref="references/sync-commands.md#push",
        tests=("contract.push.success", "contract.push.target_not_open"),
    ),
    CommandSpec(
        verb="shells", entry="sync", handler="features.sync_push:shells",
        mutation=MutationMode.PREVIEW, project=ProjectMode.NONE,
        supports_json=True, exit_codes=(0, 1, 2),
        args=(
            ArgSpec("--mode", kind="choice", choices=("create", "clean"),
                    default="create", help="create 生成（默认）/ clean 只删"),
        ),
        summary="公共库占位壳（同名一行注释，写回时跳过）",
        ref="references/sync-commands.md#shells",
        tests=("contract.shells.create", "contract.shells.clean"),
    ),
)
