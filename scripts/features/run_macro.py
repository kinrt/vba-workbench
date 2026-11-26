# -*- coding: utf-8 -*-
"""run 切片的宏执行装配段：参数解析 → 预检/备份/写回/回读 → 跑宏 → 汇总结果与 journal。

职责边界：
  · 源码预检、公共库 + 项目源码写回工程、会话获取（attach/isolate）、
    run journal 落盘，全部在本模块一份实现；
  · 本模块不自己拼 JSON 信封、不回壳别的 CLI；输出由入口层 ``render`` 渲染；
  · watch 常驻监听辅助不在本模块，见 ``run_watch``。
"""
from __future__ import annotations

import datetime as dt
import re
import tempfile
import time
import types
from pathlib import Path
from typing import Any

from adapters import excel_gateway
from toolkit import check, excel, files, repo, vba_parse
from toolkit.base import Failure as BaseFailure, abs_path, entry_command, read_source_text, warn
from toolkit.contract import (ErrorCode, Failure, Outcome, Stage,
                              to_contract_failure)
from toolkit.macro import RUN_FAIL_HINTS, MacroRunner
from features.common_subset import (  # noqa: E402
    find_project as _find_project,
    resolve_common_subset as _resolve_common_subset_shared,
    restore_module_code as _restore_module_code,
    scope_human as _scope_human,
    scope_payload as _scope_payload,
    snapshot_module_code as _snapshot_module_code,
    warn_unknown_modules as _warn_unknown_modules,
)

# 模块级睡眠钩子：测试可 monkeypatch 避免真睡；生产走 time.sleep。
_sleep = time.sleep


# --------------------------------------------------------------------------- 小工具
def _coerce_run_arg(text: str):
    """命令行参数 → VBA 认的类型：True/False → Boolean，整数 → Long，
    小数 → Double，其余原样当字符串（对象参数只能写包装宏）。"""
    low = str(text).strip().lower()
    if low in ("true", "false"):
        return low == "true"
    try:
        return int(str(text))
    except ValueError:
        pass
    try:
        return float(str(text))
    except ValueError:
        pass
    return str(text)


def _usage(message: str) -> Failure:
    return Failure(message, code=ErrorCode.USAGE, stage=Stage.CONTRACT,
                   hint="核对命令行参数；加 --help 查看用法。")


# --------------------------------------------------------------------------- 源码分析
def _source_dirs(proj: dict, defaults: dict) -> list[Path]:
    dirs: list[Path] = []
    if defaults.get("common_dir"):
        dirs.append(abs_path(defaults["common_dir"]))
    if proj.get("src"):
        dirs.append(abs_path(proj["src"]))
    return dirs


def _common_subset(proj: dict, defaults: dict) -> tuple[set[str] | None, str]:
    """决定写回哪些公共模块：None=全部；白名单=列出；auto=按引用。"""
    subset, reason, ignored = _resolve_common_subset_shared(proj, defaults)
    _warn_unknown_modules(ignored)
    return subset, reason


def _macro_declared(cfg: dict, proj: dict, macro: str) -> bool:
    """粗略判断宏是否在项目源码/公共子集里声明过（只用于提示文案）。"""
    name = macro.split("!")[-1].split(".")[-1]
    try:
        proj_src = abs_path(proj["src"])
        if proj_src.is_dir() and any(name in vba_parse.declared_procs(f)
                                     for f in repo.source_files(proj_src)):
            return True
        common_dir = abs_path(cfg["defaults"].get("common_dir", "src/common"))
        subset, _ = _common_subset(proj, cfg["defaults"])
        if common_dir.is_dir():
            return any(repo.module_name_of(f) in (subset or set())
                       and name in vba_parse.declared_procs(f)
                       for f in repo.source_files(common_dir))
    except Exception:  # noqa: BLE001 - 探测失败交分类器兜底
        pass
    return False


def _resolve_macro_ref(cfg: dict, proj: dict, wb, macro: str) -> tuple[str, str]:
    """组装 Application.Run 引用；模块名与过程名重名时自动写成 模块.过程。"""
    if "!" in macro:
        return macro, "原样（已带工作簿限定）"
    if "." in macro:
        return f"'{wb.Name}'!{macro}", "原样（已带模块名）"

    proj_src = abs_path(proj["src"])
    common_dir = abs_path(cfg["defaults"].get("common_dir", "src/common"))
    subset, _ = _common_subset(proj, cfg["defaults"])
    stems: list[str] = []
    if proj_src.is_dir():
        stems += [repo.module_name_of(f) for f in repo.source_files(proj_src)
                  if macro in vba_parse.declared_procs(f)]
    if common_dir.is_dir():
        stems += [repo.module_name_of(f) for f in repo.source_files(common_dir)
                  if repo.module_name_of(f) in (subset or set())
                  and macro in vba_parse.declared_procs(f)]
    if macro in stems:
        return f"'{wb.Name}'!{macro}.{macro}", "宏名与模块重名，已自动写成 模块.过程"
    return f"'{wb.Name}'!{macro}", "直接调用"


# --------------------------------------------------------------------------- 源码预检
def _source_preflight(proj: dict, defaults: dict, doc_modules: bool) -> list[str]:
    """写回前硬预检（只读源码）；返回致命项清单。

    有意保留：与 sync/agent 预检的差异轴为 notes 软提醒、检查深度、子集过滤；
    不抽公共骨架，避免 sync 告警丢失或 agent 多出不适用的检查。
    """
    fatal: list[str] = []
    dirs = _source_dirs(proj, defaults)
    for folder in dirs:
        if not folder.is_dir():
            fatal.append(f"源码目录不存在：{folder}")

    doc_name_re = re.compile(r"^(ThisWorkbook|Sheet\d+|Chart\d+|DialogSheet\d+|MacroSheet\d+)$",
                             re.IGNORECASE)
    vb_name_re = re.compile(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"', re.IGNORECASE | re.MULTILINE)
    total = 0

    for folder in dirs:
        if not folder.is_dir():
            continue
        local: dict[str, Path] = {}
        for f in repo.source_files(folder):
            if repo.is_shell_file(f):
                continue
            total += 1
            name = repo.module_name_of(f)
            is_doc = f.suffix.lower() == ".bas" and bool(doc_name_re.match(name))
            if name in local:
                fatal.append(f"同一目录内模块重名：{name}（{local[name]} / {f}）")
            local[name] = f

            if f.suffix.lower() == ".frm" and not f.with_suffix(".frx").is_file():
                fatal.append(f"窗体缺同名 .frx：{f}")

            try:
                text = read_source_text(f)
            except Exception as e:  # noqa: BLE001
                fatal.append(f"源码读不出：{f}（{e}）")
                continue

            m = vb_name_re.search(text)
            if m and m.group(1).strip() != name:
                fatal.append(f"文件头 VB_Name 与文件名不一致：{f}")

            if not is_doc and f.suffix.lower() in (".bas", ".cls"):
                body = [ln.rstrip() for ln in vba_parse.norm_code_lines(text) if ln.strip()]
                if not body:
                    fatal.append(f"模块正文为空：{f}（非文档模块/占位壳）")

    if total == 0:
        fatal.append("源码目录里没有任何可写模块（只有占位壳也算没有）")
    return fatal


# --------------------------------------------------------------------------- 写回核心
def _code_body(text: str, vbe: bool = False) -> list[str]:
    text = vba_parse.drop_attribute_lines(text)
    if vbe:
        text = vba_parse.drop_trailing_placeholder(text)
    return [ln for ln in vba_parse.norm_code_lines(text) if ln.strip()]


def _postcheck(vbproj, written: dict[str, Path]) -> list[str]:
    """逐模块回读比对；返回不一致明细。"""
    bad: list[str] = []
    for name, path in written.items():
        comp = excel.find_component(vbproj, name)
        if comp is None:
            bad.append(f"{name}（工程里找不到）")
            continue
        try:
            cm = comp.CodeModule
            got_text = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
            want_text = read_source_text(path)
        except Exception as e:  # noqa: BLE001
            bad.append(f"{name}（回读失败：{e}）")
            continue
        want = vba_parse.fold_lines(_code_body(want_text))
        got = vba_parse.fold_lines(_code_body(got_text, vbe=True))
        if want != got:
            first = next((i for i, (a, b) in enumerate(zip(want, got), 1) if a != b),
                         min(len(want), len(got)) + 1)
            bad.append(f"{name}（源码 {len(want)} 行 / VBE {len(got)} 行；首个差异第 {first} 行）")
    return bad


def _import_module_file(vbproj, file: Path, *, doc_modules: bool, dry: bool,
                        tmp_dir: Path | None) -> tuple[str, str]:
    """写回单个源码文件；返回 (状态 imported/skipped/failed, 模块名)。"""
    name = repo.module_name_of(file)
    comp = excel.find_component(vbproj, name)
    ctype = comp.Type if comp is not None else None
    is_doc = ctype == repo.VBXT_DOCUMENT or (
        ctype is None and repo.DOC_MODULE_RE.match(name)
        and name.lower().startswith(("thisworkbook", "sheet", "chart")))

    if is_doc:
        if not doc_modules:
            return "skipped", name
        if comp is None:
            return "skipped", name
        if not dry:
            excel.replace_module_code(
                comp, vba_parse.strip_export_header(read_source_text(file)))
        return "imported", name

    if not dry:
        in_place = (ctype in (repo.VBXT_STANDARD, repo.VBXT_CLASS)
                    and repo.TYPE_FOR_EXT.get(file.suffix.lower()) == ctype)
        if in_place:
            excel.replace_module_code(
                comp, vba_parse.strip_export_header(read_source_text(file)))
            return "imported", name
        if comp is not None and ctype in repo.IMPORTABLE:
            if not excel.remove_module(vbproj, name):
                return "failed", name
        copied = files.gbk_import_copy(file, tmp_dir)
        new_comp = excel.import_component(vbproj, copied, name)
        if new_comp is None:
            return "failed", name
    return "imported", name


def _sync_into(wb, proj: dict, defaults: dict, *, doc_modules: bool, dry: bool,
               rollback_on_fail: bool) -> list[str]:
    """公共库 + 项目源码写回工程；返回回读校验失败明细（空=通过）。"""
    folders = _source_dirs(proj, defaults)
    common_dir = abs_path(defaults.get("common_dir", "src/common"))
    subset, _ = _common_subset(proj, defaults)
    vbproj = wb.VBProject

    before = excel.component_snapshot(vbproj) if not dry else None
    code_before = _snapshot_module_code(vbproj) if rollback_on_fail and not dry else {}
    written: dict[str, Path] = {}
    tmp_dir: Path | None = None
    if not dry:
        tmp_dir = Path(tempfile.mkdtemp(prefix="vba-import-"))

    try:
        for folder in folders:
            if not folder.is_dir():
                continue
            only = subset if folder == common_dir else None
            used: set[str] = set()
            for f in repo.source_files(folder):
                name = repo.module_name_of(f)
                if name in used:
                    continue
                used.add(name)
                if only is not None and name not in only:
                    continue
                if repo.is_shell_file(f):
                    continue
                status, _ = _import_module_file(
                    vbproj, f, doc_modules=doc_modules, dry=dry, tmp_dir=tmp_dir)
                if status == "imported":
                    written[name] = f
                elif status == "failed":
                    warn(f"    写回失败：{name} 导入后拿不到同名组件")

        if dry:
            return []

        diff = check.check_module_set(before, excel.component_snapshot(vbproj))
        if diff is not None and not diff.ok:
            for level, text in diff.notes("警告"):
                warn(f"  [{level}] {text}")

        bad = _postcheck(vbproj, written)
        if bad and rollback_on_fail and code_before:
            snap = {n: code_before.get(n, "") for n in written}
            restored, failed = _restore_module_code(vbproj, snap)
            warn(f"  失败回滚：已按写回前内存正文复原 {len(restored)} 个模块")
            if failed:
                bad.append("回滚未完成：" + "；".join(failed))
        return bad
    finally:
        if tmp_dir is not None:
            files.remove_tree(tmp_dir)


# --------------------------------------------------------------------------- 会话获取
def _acquire(cfg: dict, proj: dict, *, isolate: bool, push: bool, macro: str):
    """取运行会话；返回 (wb, xl, mode, owned_xl|None)。只走 adapter/toolkit 机械动作。"""
    path, reason = repo.resolve_target(proj)
    if path is None:
        raise Failure(
            f"「{proj['name']}」没解析到 xlsm：{reason}",
            code=ErrorCode.TARGET_NOT_FOUND, stage=Stage.RESOLVE,
            hint="核对 mapping.json 的 search_dir 与 prefix；文件名必须是「prefix V<数字>.xlsm」",
            suggested=entry_command("sync.py", "list", "--json"),
        )

    if not isolate:
        xl = excel_gateway.attach_running_excel(
            path,
            target_hint="表没打开：先打开目标工作簿，或用 run --isolate",
            target_suggested=entry_command(
                "sync.py", "run", "--project", proj["name"],
                "--macro", macro, "--isolate", "--apply"),
        )
        wb = excel.require_open_workbook(
            xl, path, proj["name"], "运行宏",
            hint="表没打开：先打开目标工作簿，或用 run --isolate",
            suggested=entry_command(
                "sync.py", "run", "--project", proj["name"],
                "--macro", macro, "--isolate", "--apply"))
        return wb, xl, "attach", None

    # 隔离模式会自己开副本并保存；目标已打开时不能抢，避免文件锁/覆盖未保存改动。
    try:
        excel_gateway.attach_running_excel(path)
    except BaseFailure as e:
        if getattr(e, "code", "") not in ("NO_EXCEL", "TARGET_NOT_OPEN"):
            # 放行白名单与错误码归类不动：非「没开」口径（如 COM_CONTEXT）仍原样
            # 上抛；这里只换提示层——目标实例 COM 不可达时隔离副本同样不会启动，
            # 网关自带建议在隔离语境下（还叫人再开 isolate）不成立，不能照搬。
            raise Failure(
                str(e),
                code=getattr(e, "code", "") or ErrorCode.UNEXPECTED.value,
                stage=getattr(e, "stage", "") or Stage.OPEN.value,
                hint=("目标实例 COM 不可达：可能卡在弹窗/受保护视图，或运行在权限不同的会话；"
                      "隔离模式也不会启动——预检连不上这个实例，就不敢另开副本去抢文件锁。"
                      "先关掉卡住的实例再重试；只是想跑别的表时，再在那张表上走 --isolate。"),
                suggested=entry_command(
                    "sync.py", "run", "--project", proj["name"],
                    "--macro", macro, "--isolate", "--apply"),
                retryable=bool(getattr(e, "retryable", False)),
            ) from e
    else:
        raise Failure(
            "隔离模式要求目标表未打开；现在它在 Excel 里开着。"
            "请先关闭那个窗口，或去掉 --isolate 用已打开会话运行。",
            code=ErrorCode.TARGET_OPEN, stage=Stage.OPEN,
            suggested="关掉目标 Excel 窗口后重跑；需要保留未保存改动则不要用 --isolate",
        )

    xl = excel_gateway.new_excel(hidden=True)
    try:
        wb = excel_gateway.open_disk_workbook(xl, path, read_only=False)
        if push and not excel.ensure_vbe_access(wb, proj["name"], interactive=False):
            state, detail = excel.vbe_state(wb)
            code = ErrorCode.VBE_PROTECTED if state == "protected" else ErrorCode.COM_CONTEXT
            raise Failure(f"「{proj['name']}」的 VBA 工程读不了：{detail}",
                          code=code, stage=Stage.VBE)
    except Exception as e:
        cleanup_problems = excel.quit_owned(xl)
        if cleanup_problems:
            # 诊断挂到异常上，由 run() 收口进 warnings/stats，不再裸吞。
            try:
                e.cleanup_problems = cleanup_problems
            except Exception:  # noqa: BLE001 - 异常对象不允许挂属性就只走主错误通道
                pass
        raise
    return wb, xl, "isolate", xl


# --------------------------------------------------------------------------- run journal
def _journal_path(run_dir: Path, proj: str, macro: str) -> Path:
    d = Path(run_dir)
    d.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    safe = re.sub(r'[<>:"/\\|?*\s]+', "_", f"{proj}-{macro}")[:80]
    return d / f"{stamp}-{safe}.json"


def _write_journal(ctx: Any, *, proj: str, macro: str, mode: str, pushed: bool,
                   saved: bool, save_requested: bool, started_at: str,
                   exit_code: int, result: Any, error: str, log: str) -> None:
    """每次 run --apply 落一份结构化 journal（只记录，不支持 --resume）。"""
    run_dir = getattr(ctx, "journal_dir", None) or (ctx.root / ".run")
    path = _journal_path(run_dir, proj, macro)
    data = {
        "project": proj,
        "macro": macro,
        "mode": mode,
        "pushed": bool(pushed),
        "saved": bool(saved),
        "save_requested": bool(save_requested),
        "save_ignored": bool(save_requested and not saved and mode != "isolate"),
        "started_at": started_at,
        "ended_at": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
        "exit_code": int(exit_code),
        "result": result,
        "error": error or None,
        "log": log or "",
    }
    files.dump_json(path, data)


# --------------------------------------------------------------------------- run 动词
def run(ctx: Any) -> Outcome:
    """src→VBE（可选）→ 运行宏 → 取回结果；返回值与日志两条通道。"""
    args = ctx.args
    project = getattr(args, "project", None)
    macro = getattr(args, "macro", None)
    if not project or not macro:
        raise _usage("run 需要 --project <项目名> 与 --macro <宏名>")

    cfg = ctx.config
    proj = _find_project(cfg, project)
    argv = [_coerce_run_arg(a) for a in (getattr(args, "arg", None) or [])]

    dry = not bool(getattr(args, "apply", False))
    isolate = bool(getattr(args, "isolate", False))
    save = bool(getattr(args, "save", False))
    push = bool(getattr(args, "push", False))
    tail = int(getattr(args, "tail", 40))
    log_target = getattr(args, "log", None)
    log_optional = bool(getattr(args, "log_optional", False))
    rollback = bool(getattr(args, "rollback_on_fail", False))
    doc = bool(cfg.get("defaults", {}).get("import_document_modules", False))

    cmd_started = time.perf_counter()
    started_at = dt.datetime.now().astimezone().isoformat(timespec="seconds")
    mode = "isolate" if isolate else "attach"

    # 预览：不开表、不跑宏，只把将要执行的动作列出来。
    if dry:
        # dry 同样先解析子集：预览即契约，先把将写回的公共模块范围与未知模块告警摆出来。
        subset, reason = _common_subset(proj, cfg.get("defaults", {}))
        scope = _scope_payload(subset)
        human = [f"[dry] 将运行 {project} / {macro}，参数 {argv}；先写回={push}，"
                 f"隔离={isolate}，保存={save}。加 --apply 才真正执行",
                 _scope_human(subset, reason)]
        details = [{"name": project, "ok": True, "macro": macro, "dry": True}]
        stats = {"project": project, "macro": macro, "args": argv, "push": push,
                 "isolate": isolate, "save": save, "dry": True, **scope,
                 "results": [{"name": project, "ok": True, "macro": macro, "dry": True}]}
        return Outcome(action="sync.run", ok=True, details=details, stats=stats,
                       human=human, exit_code=0, elapsed_ms=ctx.elapsed_ms())

    errors: list[Failure] = []
    human: list[str] = []
    warnings: list[str] = []
    result: Any = None
    run_error = ""
    log_file = ""
    saved = False
    exit_code = 1
    stats: dict[str, Any] = {"project": project, "macro": macro, "mode": mode}
    owned_xl: Any = None
    exc_cleanup_problems: list[str] | None = None

    try:
        # --push 预检不碰 Excel：源码有问题时连表都不开。
        if push:
            with ctx.progress.stage("preflight"):
                fatal = _source_preflight(proj, cfg.get("defaults", {}), doc)
            if fatal:
                raise Failure("源码预检未通过，已放弃写回：" + "；".join(fatal[:3]),
                              code=ErrorCode.SOURCE_PREFLIGHT, stage="preflight",
                              suggested="按提示修源码/文件名后重跑；不要绕过预检直接写回")

        with ctx.progress.stage("attach"):
            wb, xl, mode, owned_xl = _acquire(
                cfg, proj, isolate=isolate, push=push, macro=macro)

        if push:
            with ctx.progress.stage("push"):
                backup_root = abs_path(cfg.get("defaults", {}).get("backup_dir", ".backup"))
                excel.backup_workbook(wb, backup_root, tag="运行前")
                bad = _sync_into(wb, proj, cfg.get("defaults", {}),
                                 doc_modules=doc, dry=False,
                                 rollback_on_fail=rollback)
                if bad:
                    raise Failure("写回后回读校验失败，未运行宏：" + "；".join(bad),
                                  code=ErrorCode.POSTCHECK_FAILED, stage="postcheck",
                                  suggested="检查源码编码/Attribute 行；修好后重跑")

        macro_ref = _resolve_macro_ref(cfg, proj, wb, macro)
        session = types.SimpleNamespace(workbook=wb)
        session.macro_declared = lambda m: _macro_declared(cfg, proj, m)

        with ctx.progress.stage("run"):
            runner = MacroRunner(session)
            run_result = runner.run(
                macro_ref[0], argv, isolate=isolate, save=save,
                log_optional=log_optional, log_target=log_target, tail=tail)

        result = run_result.return_value
        saved = run_result.saved
        log_file = run_result.log_file
        warnings.extend(run_result.warnings)
        if save and not isolate:
            warnings.append("--save 仅对 --isolate 生效，本次非隔离，未保存")

        run_sec = round(run_result.duration_ms / 1000, 2)
        human.append(f"运行结束，耗时 {run_result.duration_ms / 1000:.1f} 秒")
        if result is None:
            human.append("返回值：空（这是 Sub；结果看日志，或它自己写回的单元格）")
        else:
            human.append(f"返回值：{result!r}")

        if run_result.log_status == "found":
            human.append(f"--- 日志尾部（{log_file}）---")
            human.append(run_result.log_text)
            human.append("--- 结束 ---")
        elif run_result.log_status == "missing":
            human.append("没有取到日志：没找到新生成的 .log；用 --log 指定目录或文件，"
                         "或看宏自己的返回值")
        else:
            human.append("没有取到日志：该宏是 Sub / 按 --log-optional，不写日志属正常")

        stats["macro"] = macro_ref[0]
        stats.update({
            "args": argv, "pushed": bool(push), "mode": mode, "saved": saved,
            "log": log_file, "cost_sec": round(time.perf_counter() - cmd_started, 2),
            "run_sec": run_sec, "result": result,
            "log_status": run_result.log_status,
        })
        details = [{"name": project, "ok": True, "macro": macro_ref[0], "mode": mode,
                    "args": argv, "result": result}]
        exit_code = 0
    except BaseFailure as e:
        err = to_contract_failure(e)
        errors.append(err)
        run_error = str(e)
        human.append(f"[错误] {e}")
        exc_cleanup_problems = getattr(e, "cleanup_problems", None)
        if err.hint:
            human.append(f"说明：{err.hint}")
        if err.code == ErrorCode.RUN_FAILED.value:
            human.extend(f"  · {line}" for line in RUN_FAIL_HINTS)
        cost_sec = round(time.perf_counter() - cmd_started, 2)
        stats.update({
            "ok": False, "mode": mode, "cost_sec": cost_sec, "run_sec": 0.0,
            "error_code": err.code, "stage": err.stage,
            "suggested_command": err.suggested, "hint": err.hint,
            "errors": [str(e)],
            "results": [{"name": project, "ok": False, "macro": macro,
                         "mode": mode, "error": str(e)}],
        })
        details = stats.get("results", [])
    finally:
        try:
            _write_journal(ctx, proj=project, macro=macro, mode=mode, pushed=push,
                           saved=saved, save_requested=save, started_at=started_at,
                           exit_code=exit_code, result=result, error=run_error,
                           log=log_file)
        except Exception as je:  # noqa: BLE001 - 日志失败不能盖住运行结果
            warnings.append(f"run journal 写入失败：{je}")
        if owned_xl is not None:
            # N2 白名单：隔离实例由 new_excel 刚起、启动时为空；现存簿里除本次
            # 打开的目标表外，只能是宏运行中自行 Workbooks.Open 的参考簿，保留不关。
            keep: list[str] = []
            try:
                target_name = str(getattr(wb, "Name", ""))
                keep = [str(getattr(b, "Name", "")) for b in owned_xl.Workbooks
                        if str(getattr(b, "Name", "")) != target_name]
            except Exception:  # noqa: BLE001 - 读不到簿名单就退回旧收尾口径
                keep = []
            cleanup_problems = excel.quit_owned(owned_xl, keep_workbooks=keep or None)
            if cleanup_problems:
                warnings.append("隔离实例回收未完成：" + "；".join(cleanup_problems[:3]))
                stats["isolate_cleanup_problems"] = cleanup_problems
            else:
                stats["isolate_cleanup"] = "ok"
        else:
            # 隔离开表失败：诊断由 _acquire 挂在异常上，这里收口进 warnings/stats。
            if exc_cleanup_problems:
                warnings.append("隔离实例回收未完成：" + "；".join(exc_cleanup_problems[:3]))
                stats["isolate_cleanup_problems"] = exc_cleanup_problems

    return Outcome(
        action="sync.run", ok=not errors, details=details, errors=errors,
        stats=stats, human=human, warnings=warnings, exit_code=exit_code,
        elapsed_ms=ctx.elapsed_ms(),
    )
