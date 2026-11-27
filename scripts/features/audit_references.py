# -*- coding: utf-8 -*-
"""audit references：只读检查工作簿 VBA 工程的 References。

只读边界：不增删引用、不改工程、不保存工作簿；表没开时用隐藏实例只读打开，
读完关闭自己开的那一份。工程被锁/信任中心未放行时退出 3（ok=false 不归一为 1），
只有表路径/打开这类业务失败才退出 1。
"""
from __future__ import annotations

import re
from pathlib import Path

from adapters import excel_gateway
from toolkit.base import read_source_text
from toolkit.contract import ErrorCode, Failure, MutationMode, Outcome, Stage
from toolkit.excel import ExcelSession, find_open_workbook, open_disk_workbook, vbe_state
from toolkit.files import write_text_utf8
from toolkit.registry import ArgSpec, CommandSpec, ProjectMode
from toolkit.repo import rel_path, resolve_target, source_files
from toolkit.vba_parse import mask_strings, strip_vba_comments

from .common_subset import find_project


# 已知引用 → 代码里的类型前缀；只用于“被哪些模块用到”的轻量提示，不当作硬判定。
_PREFIX_HINTS = (
    ("scripting", "Scripting"),
    ("ado ext", "ADOX"),
    ("activex data objects", "ADODB"),
    ("forms", "MSForms"),
    ("internet controls", "SHDocVw"),
    ("common controls", "MSComctlLib"),
    ("dao", "DAO"),
    ("office", "Office"),
)


def _prefix_for(ref: dict) -> str:
    text = f"{ref.get('name', '')} {ref.get('description', '')}".lower()
    for key, prefix in _PREFIX_HINTS:
        if key in text:
            return prefix
    return ""


def _local_dir(root: Path, spec: str) -> Path:
    p = Path(spec)
    return p if p.is_absolute() else Path(root) / p


def _source_usage(proj: dict, defaults: dict, refs: list[dict], root: Path) -> None:
    """给每条引用填 used_by（源码里出现 `<Prefix>.` 的模块），只做提示。"""
    roots = []
    for key in ("src",):
        if proj.get(key):
            roots.append(_local_dir(root, proj[key]))
    if defaults.get("common_dir"):
        roots.append(_local_dir(root, defaults["common_dir"]))
    modules: dict[str, str] = {}
    for root_dir in roots:
        for f in source_files(root_dir):
            try:
                code = mask_strings(strip_vba_comments(read_source_text(f)))
            except Exception:
                continue
            modules[rel_path(f)] = code
    for ref in refs:
        prefix = _prefix_for(ref)
        ref["usage_prefix"] = prefix
        if not prefix:
            ref["used_by"] = []
            continue
        pat = re.compile(rf"(?<!\w){re.escape(prefix)}\s*\.", re.IGNORECASE)
        ref["used_by"] = [path for path, code in modules.items() if pat.search(code)]


def _ref_status(ref: dict) -> str:
    if ref.get("broken"):
        return "broken"
    if ref.get("built_in"):
        return "ok"
    full = str(ref.get("full_path") or "")
    if not full or not Path(full).is_file():
        return "missing"
    return "ok"


def _write_md(refs: list[dict], path: str, project: str) -> None:
    out = ["# VBA 引用体检", "", f"- 项目：{project}", "",
           "| 引用 | 状态 | 版本 | 路径 | 使用模块 |", "|---|---|---|---|---|"]
    out += [f"| {r['name']} | {r['status']} | {r.get('major', '')}.{r.get('minor', '')} | "
            f"{r.get('full_path', '')} | {', '.join(r.get('used_by', []))} |" for r in refs]
    write_text_utf8(path, "\n".join(out) + "\n")


def _outcome(project: str, refs: list[dict], warnings: list[str], readable: bool,
             ctx, detail: str = "") -> Outcome:
    ok = sum(1 for r in refs if r["status"] == "ok")
    missing = sum(1 for r in refs if r["status"] == "missing")
    broken = sum(1 for r in refs if r["status"] == "broken")
    strict = bool(getattr(ctx.args, "strict", False))
    code = 0 if (readable and not warnings) else (1 if strict else 3)
    errors: list[Failure] = []
    if code == 1:
        errors.append(Failure(f"引用体检有 {len(warnings)} 条提醒，按 --strict 失败",
                              code=ErrorCode.SOURCE_PREFLIGHT, stage=Stage.PARSE))

    if getattr(ctx.args, "md", None):
        _write_md(refs, str(ctx.args.md), project)

    human = [f"VBA 引用体检（只读）：{project}",
             f"  引用 {len(refs)}：正常 {ok} / 缺失 {missing} / 损坏 {broken}"
             + ("" if readable else f"  （工程不可读：{detail}）")]
    for r in refs:
        mark = {"ok": "OK", "missing": "缺失", "broken": "损坏"}[r["status"]]
        line = f"  [{mark}] {r['name']}"
        if r.get("full_path"):
            line += f"  {r['full_path']}"
        if r.get("used_by"):
            line += f"  被 {len(r['used_by'])} 个模块使用"
        human.append(line)
    for w in warnings:
        human.append(f"  ! {w}")

    return Outcome(
        action="audit.references",
        ok=code != 1,
        details=refs,
        errors=errors,
        stats={"project": project, "readable": readable, "total": len(refs),
               "ok": ok, "missing": missing, "broken": broken,
               "timings": ctx.stage_timings()},
        human=human,
        warnings=warnings,
        exit_code=code,
        elapsed_ms=ctx.elapsed_ms(),
    )


def _open_readonly(proj: dict, path: Path):
    """连已打开会话；表没开时用隐藏实例只读打开，返回 (session, wb, owned_wb)。"""
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


def references(ctx) -> Outcome:
    args = ctx.args
    project_name = str(getattr(args, "project", "") or "")
    proj = find_project(ctx.config, project_name)

    with ctx.progress.stage("references"):
        path, why = resolve_target(proj)
        if path is None:
            raise Failure(f"「{project_name}」没解析到 xlsm：{why}",
                          code=ErrorCode.TARGET_NOT_FOUND, stage=Stage.RESOLVE,
                          hint="核对 mapping.json 的 search_dir / prefix / file")
        session = None
        wb = None
        owned_wb = False
        try:
            session, wb, owned_wb = _open_readonly(proj, Path(path))
            state, detail = vbe_state(wb)
            if state != "ok":
                return _outcome(project_name, [], [f"VBA 工程读不了：{detail}"],
                                False, ctx, detail=detail)
            snap = excel_gateway.read_project_references(wb)
            if snap["error"]:
                return _outcome(project_name, [], [f"引用集合读不了：{snap['error']}"],
                                False, ctx, detail=snap["error"])
            refs = list(snap["references"])
            for ref in refs:
                ref["status"] = _ref_status(ref)
            _source_usage(proj, ctx.config.get("defaults", {}), refs, ctx.root)
            warnings: list[str] = []
            for ref in refs:
                if ref["status"] == "missing":
                    warnings.append(f"引用缺失：{ref['name']}（路径不存在：{ref.get('full_path') or '空'}）")
                elif ref["status"] == "broken":
                    warnings.append(f"引用损坏：{ref['name']}")
            return _outcome(project_name, refs, warnings, True, ctx)
        finally:
            if owned_wb and wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
            if session is not None:
                session.close()


SPECS: tuple[CommandSpec, ...] = (
    CommandSpec(
        entry="audit",
        verb="references",
        handler="features.audit_references:references",
        mutation=MutationMode.NONE,
        project=ProjectMode.REQUIRED,
        summary="只读检查 VBA 工程引用（缺失/损坏/路径；默认警告退出 3）",
        ref="references/audit.md#references",
        exit_codes=(0, 1, 2, 3),
        args=(
            ArgSpec("--strict", kind="flag",
                    help="有缺失/损坏或不可读时退出 1（默认退出 3）"),
            ArgSpec("--md", kind="path", help="另写一份 Markdown 报告"),
            ArgSpec("--out", kind="path", help="另写一份 JSON 到文件"),
        ),
        tests=("tests/test_audit_references.py",),
    ),
)
