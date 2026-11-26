# -*- coding: utf-8 -*-
"""audit_a 切片：members / common / project 三个只读视角。

分层（对齐 1.6 骨架）：
  · 本层只做参数适配与输出装配：解析命令参数、调 ``AuditEngine``、拼 human/stats/details；
  · 解析、引用扫描、可达性、依赖矩阵、汇总判定等规则口径全部在 ``audit_engine`` 一份；
  · 本层不复制阈值表与扫描循环；错误码映射（找不到模块/项目）属适配层职责。
"""
from __future__ import annotations

from pathlib import Path

from toolkit.base import abs_path, Failure as BaseFailure
from toolkit.contract import ErrorCode, Failure, MutationMode, Outcome, Stage
from toolkit.files import write_text_utf8
from toolkit.registry import ArgSpec, CommandSpec, ProjectMode
from toolkit.repo import find_source, module_name_of, rel_path
from toolkit.ui import SEP, display_width, pad_display, rjust
from toolkit.vba_parse import code_lines, parse_members

from .audit_engine import AuditEngine


# ---------------------------------------------------------------------------
# 引擎适配（委托至 engine；本层不保留解析/规则副本）
# ---------------------------------------------------------------------------
def _engine(ctx) -> AuditEngine:
    """按上下文构造共享审计引擎（规则单源在 audit_engine）。"""
    return AuditEngine(ctx.root, ctx.config)


def _resolve(spec: str) -> Path:
    """模块名/相对路径 → 源码路径；找不到映射为 TARGET_NOT_FOUND（适配层错误映射）。"""
    try:
        return find_source(spec)
    except BaseFailure as e:
        raise Failure(str(e), code=ErrorCode.TARGET_NOT_FOUND, stage=Stage.RESOLVE,
                      hint="给模块名，或 src 下的相对路径") from e


def _analyze(engine: AuditEngine, spec: str, projects=None, exclude=None) -> dict:
    """委托至 engine.analyze_target；模块名解析在适配层完成。"""
    target = _resolve(spec)
    return engine.analyze_target(target, projects, exclude)


def _project_files(engine: AuditEngine, name: str, exclude=None) -> list[Path]:
    """委托至 engine.project_files；缺失项目映射为 PROJECT_NOT_FOUND（只探存在性，不复制扫描）。"""
    files = engine.project_files(name, exclude)
    if files:
        return files
    known = "、".join(p.get("name", "") for p in engine.projects) or "（src/projects 为空）"
    raise Failure(f"没有项目「{name}」，可用：{known}",
                  code=ErrorCode.PROJECT_NOT_FOUND, stage=Stage.RESOLVE)


def _common_analyses(engine: AuditEngine, projects=None, exclude=None):
    """委托至 engine.common_analyses；跳过原因经 engine.last_skipped 取回。

    返回形状保持 ``(analyses, warnings)`` 不变（tests/test_new_issues.py 在此有 patch 缝）。
    """
    analyses = engine.common_analyses(projects, exclude)
    return analyses, list(engine.last_skipped)


def _own_modules(engine: AuditEngine, proj_files: list[Path], exclude=None) -> list[dict]:
    """本表专属模块清单（装配用）：剔除公共库同名模块，成员数走共享解析器计数。"""
    common_stems = {f.stem for f in engine.common_files(exclude)}
    own: list[dict] = []
    for f in proj_files:
        if module_name_of(f) in common_stems:
            continue
        try:
            members = len(parse_members(code_lines(f)))
        except BaseFailure:
            members = 0
        own.append({"name": f.name, "kind": engine._kind_of(f), "members": members, "path": f})
    return own


# ---------------------------------------------------------------------------
# 报告排版（只产人行，不打印）
# ---------------------------------------------------------------------------
def _grid_lines(headers, rows, right=()) -> list[str]:
    if not rows:
        return ["  （无）"]
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

    lines = ["  " + "".join(cell(headers[i], i) for i in range(n)).rstrip(),
             "  " + "─" * (sum(widths) + 2 * n)]
    for r in rows:
        lines.append("  " + "".join(cell(r[i], i) for i in range(n)).rstrip())
    return lines


def _table_lines(items, limit=None) -> list[str]:
    if not items:
        return ["  （无）"]
    rows = []
    for r in items[:limit]:
        who = "、".join(f"{Path(k).name}×{v}" for k, v in list(r["ext_files"].items())[:3])
        if len(r["ext_files"]) > 3:
            who += f" …共{len(r['ext_files'])}处"
        rows.append([r["member"], r["kind"], r["span"], r["ext"], r["internal"], r["suspect"],
                     (who or "-") + ("  ⚠动态调用" if r["dynamic"] else "")])
    lines = _grid_lines(["成员", "类型", "行数", "外部", "类内", "疑似", "引用方"],
                        rows, right=(2, 3, 4, 5))
    if limit and len(items) > limit:
        lines.append(f"  …另有 {len(items) - limit} 项（--all 看全表）")
    return lines


def _scope_line(res) -> str:
    return (f"  范围    ：{len(res['files'])} 个源文件（{res['scope']}，已跳过占位壳）"
            + (f"；排除关键字：{'、'.join(res['excluded'])}" if res["excluded"] else ""))


def _section(title: str) -> list[str]:
    return [SEP, title, SEP]


def build_members_lines(res, show_all: bool = False) -> list[str]:
    rows, dead, only_test, single = res["rows"], res["dead"], res["only_test"], res["single"]
    active = sum(1 for r in rows if r["verdict"] == "活跃")
    lines = ["", *_section(f"成员引用盘点：{res['cls']}"),
             f"  目标    ：{rel_path(res['target'])}",
             f"  口径    ：{'标准/文档模块（裸名调用 / Call / 模块名.成员）' if res['module_mode'] else '类或窗体（实例.成员 + With 块 + 工厂链）'}",
             _scope_line(res)]
    if res["facs"]:
        lines.append(f"  工厂函数：{'、'.join(sorted(res['facs']))}（返回本类实例，Variant 变量也按实例算）")
    lines += ["",
              f"  成员 {len(rows)} 个 = 字段 {sum(1 for r in rows if r['kind'] == '字段')}"
              f" + 方法 {sum(1 for r in rows if r['kind'] == '方法')}"
              f" + 属性 {sum(1 for r in rows if r['kind'] == '属性')}",
              f"  不可达 {len(dead)}  │  仅自测 {len(only_test)}  │  仅单文件 {len(single)}  │  活跃 {active}"
              f"  │  可删约 {sum(r['span'] for r in dead)}  行"]

    if dead:
        note = ("外部零引用，且类内没有可达路径走到它 —— 整条死链可一起删" if not res["module_mode"]
                else "外部零引用；但 Public Sub 仍可能是按钮/快捷键的宏入口，删前人工确认")
        lines += ["", *_section(f"① 不可达（{len(dead)}）—— {note}"),
                  *_table_lines(dead, limit=None if show_all else 60)]
    if only_test:
        lines += ["", *_section(f"② 只被自测模块引用（{len(only_test)}）—— 业务侧可删，删之前 测试* 模块要同步瘦身"),
                  *_table_lines(only_test, limit=None if show_all else 40)]
    if single:
        lines += ["", *_section(f"③ 只被单个文件引用（{len(single)}）—— 其实是那个模块的私有工具，可考虑下沉回去"),
                  *_table_lines(single, limit=None if show_all else 40)]
    if show_all:
        lines += ["", *_section(f"全部成员（{len(rows)}，按外部引用升序）"), *_table_lines(rows)]
    lines += ["", SEP,
              "  说明：「疑似」列是接收者不是本类/本模块的同名访问（如 fso.Path），不参与判定；"
              "「⚠动态调用」表示成员名出现在 Run/Evaluate 的字符串里，删之前人工确认。"]
    return lines


def build_common_lines(analyses, sums, projects, exclude, show_all: bool = False) -> list[str]:
    corpus = len(analyses[0]["files"]) if analyses else 0
    lines = ["", *_section(f"公共库总览（src/common，{len(analyses)} 个模块）"),
             f"  范围    ：{corpus} 个源文件（"
             f"{('项目 ' + '、'.join(projects)) if projects else '公共库 + 全部项目'}，已跳过占位壳）"
             + (f"；排除关键字：{'、'.join(exclude)}" if exclude else ""), ""]
    rows = [[s["module"], s["kind"], s["members"], s["dead"], s["test"], s["single"],
             s["active"], s["lines"], s["ref_files"]] for s in sums]
    lines += _grid_lines(["模块", "类型", "成员", "不可达", "仅自测", "仅单文件", "活跃", "可删行", "引用方"],
                         rows, right=(2, 3, 4, 5, 6, 7, 8))
    lines += ["",
              f"  合计：成员 {sum(s['members'] for s in sums)}，不可达 {sum(s['dead'] for s in sums)}"
              f"（约 {sum(s['lines'] for s in sums)} 行），仅自测 {sum(s['test'] for s in sums)}，"
              f"仅单文件 {sum(s['single'] for s in sums)}"]

    zero = [s for s in sums if s["ref_files"] == 0 and s["members"] and not s["infra"]]
    if zero:
        lines += ["", "  零外部引用的模块：" + "、".join(f"{s['module']}（{s['members']} 成员）" for s in zero),
                  "  → 没有任何表用到，可考虑整个退出公共库（先确认不是按钮/快捷键的宏入口）"]
    infra = [s for s in sums if s["ref_files"] == 0 and s["infra"]]
    if infra:
        lines += ["", "  保留：" + "、".join(s["module"] for s in infra)
                  + "（工具链的版本常量模块，sync.py version 要读它）"]
    if show_all:
        for res in analyses:
            lines += build_members_lines(res, show_all=False)
    return lines


def build_project_lines(name, proj_files, deps, own, engine: AuditEngine,
                        show_all: bool = False) -> list[str]:
    lines = ["", *_section(f"项目依赖面：{name}"),
             f"  源文件  ：{len(proj_files)} 个（已跳过占位壳）", ""]
    rows = [[d["module"], d["kind"], d["refs"], d["used"], d["files"], d["top"]] for d in deps]
    lines += _grid_lines(["公共模块", "类型", "引用次数", "用到成员", "涉及文件", "主要引用方"],
                         rows, right=(2, 3, 4))
    unused = [d["module"] for d in deps if not d["refs"]]
    if unused:
        lines += ["", f"  未被本表引用：{'、'.join(unused)}",
                  "  → 可以收窄 mapping.json 的 common_modules 子集，写回时少灌几个模块"]

    if own:
        lines += ["", *_section(f"本表专属模块（{len(own)}）—— 明细用 members <模块名>")]
        own_rows = [[o["name"], o["kind"], o["members"]] for o in own]
        lines += _grid_lines(["模块", "类型", "成员"], own_rows, right=(2,))
    if show_all:
        for o in own:
            try:
                lines += build_members_lines(engine.analyze_target(o["path"]), show_all=False)
            except BaseFailure as e:
                lines.append(f"  跳过 {o['name']}：{e}")
    return lines


# ---------------------------------------------------------------------------
# Markdown 报告产物（只写报告文件，不写表、不改 src）
# ---------------------------------------------------------------------------
def _rows_md(rows) -> list[str]:
    out = ["| 成员 | 类型 | 判定 | 行数 | 外部 | 类内 | 疑似 | 引用方 | 动态 |",
           "|---|---|---|--:|--:|--:|--:|---|---|"]
    for r in rows:
        who = "<br>".join(f"`{k}`×{v}" for k, v in r["ext_files"].items()) or "-"
        out.append(f"| {r['member']} | {r['kind']} | {r['verdict']} | {r['span']} | {r['ext']} | "
                   f"{r['internal']} | {r['suspect']} | {who} | {'⚠' if r['dynamic'] else ''} |")
    return out


def _write(path: str, lines: list[str]) -> Path:
    return write_text_utf8(path, "\n".join(lines) + "\n")


def write_members_md(res, path: str) -> None:
    out = [f"# {res['cls']} 成员引用盘点", "",
           f"- 目标：`{rel_path(res['target'])}`",
           f"- 口径：{'标准/文档模块（裸名调用）' if res['module_mode'] else '类或窗体（实例.成员 + With 块）'}",
           f"- 范围：{len(res['files'])} 个源文件（{res['scope']}，跳过占位壳）"
           + (f"；排除 {'、'.join(res['excluded'])}" if res["excluded"] else ""),
           f"- 成员 {len(res['rows'])}：不可达 {len(res['dead'])}、仅自测 {len(res['only_test'])}、"
           f"仅单文件 {len(res['single'])}、可删约 {sum(r['span'] for r in res['dead'])} 行",
           "", "## 瘦身候选", ""]
    cand = [r for r in res["rows"] if r["verdict"] != "活跃"]
    out += _rows_md(sorted(cand, key=lambda r: (r["verdict"], -r["span"], r["member"])))
    out += ["", "## 全部成员", ""] + _rows_md(res["rows"])
    _write(path, out)


def write_common_md(analyses, sums, path: str, projects, exclude) -> None:
    out = ["# 公共库盘点", "",
           f"- 范围：{('项目 ' + '、'.join(projects)) if projects else '公共库 + 全部项目'}"
           + (f"；排除 {'、'.join(exclude)}" if exclude else ""),
           f"- 模块 {len(sums)} 个，成员 {sum(s['members'] for s in sums)}，"
           f"不可达 {sum(s['dead'] for s in sums)}（约 {sum(s['lines'] for s in sums)} 行）", "",
           "| 模块 | 类型 | 成员 | 不可达 | 仅自测 | 仅单文件 | 活跃 | 可删行 | 引用方 |",
           "|---|--:|--:|--:|--:|--:|--:|--:|--:|"]
    for s in sums:
        out.append(f"| {s['module']} | {s['kind']} | {s['members']} | {s['dead']} | {s['test']} | "
                   f"{s['single']} | {s['active']} | {s['lines']} | {s['ref_files']} |")
    zero = [s for s in sums if s["ref_files"] == 0 and s["members"] and not s["infra"]]
    if zero:
        out += ["", "## 零外部引用的模块", ""]
        out += [f"- **{s['module']}**（{s['members']} 成员）" for s in zero]
        out += ["", "没有任何表用到，可考虑整个退出公共库；先确认不是按钮/快捷键的宏入口。"]
    infra = [s for s in sums if s["ref_files"] == 0 and s["infra"]]
    if infra:
        out += ["", "## 保留（工具链依赖）", ""]
        out += [f"- **{s['module']}**：版本常量模块，`sync.py version` 要读 `COMMON_VERSION`" for s in infra]
    for res in analyses:
        cand = [r for r in res["rows"] if r["verdict"] != "活跃"]
        if not cand:
            continue
        out += ["", f"## {res['cls']}", ""] + _rows_md(
            sorted(cand, key=lambda r: (r["verdict"], -r["span"], r["member"])))
    _write(path, out)


def write_project_md(name, proj_files, deps, own, path: str) -> None:
    out = [f"# 项目依赖面：{name}", "", f"- 源文件：{len(proj_files)} 个（跳过占位壳）", "",
           "| 公共模块 | 类型 | 引用次数 | 用到成员 | 涉及文件 | 主要引用方 |",
           "|---|--:|--:|--:|--:|---|"]
    for d in deps:
        out.append(f"| {d['module']} | {d['kind']} | {d['refs']} | {d['used']} | "
                   f"{d['files']} | {d['top']} |")
    unused = [d["module"] for d in deps if not d["refs"]]
    if unused:
        out += ["", "## 未被本表引用", "", "、".join(unused), "",
                "可收窄 mapping.json 的 common_modules 子集，写回时少灌几个模块。"]
    if own:
        out += ["", "## 本表专属模块", "", "| 模块 | 类型 | 成员 |", "|---|--:|--:|"]
        out += [f"| {o['name']} | {o['kind']} | {o['members']} |" for o in own]
    _write(path, out)


# ---------------------------------------------------------------------------
# handler：串适配与报告，回填 Outcome
# ---------------------------------------------------------------------------
def handle_members(ctx) -> Outcome:
    args = ctx.args
    exclude = getattr(args, "exclude", None)
    show_all = bool(getattr(args, "all", False))
    engine = _engine(ctx)
    with ctx.progress.stage("members 解析"):
        res = _analyze(engine, args.module, None, exclude)
    human = build_members_lines(res, show_all=show_all)
    if getattr(args, "md", None):
        write_members_md(res, args.md)
    return Outcome(
        action="audit.members",
        ok=True,
        details=res["rows"],
        errors=[],
        stats={"module": args.module, "members": len(res["rows"])},
        human=human,
        exit_code=0,
        elapsed_ms=ctx.elapsed_ms(),
    )


def _project_list(args) -> list[str] | None:
    """归一为旧版的列表口径（None 表示全量）。

    common 的 ``--project`` 是 append 可重复参数，解析后为列表；这里同时兜底
    直接塞单值字符串的桩调用，单值等价于 ``[值]``，行为与只传一次一致。
    """
    raw = getattr(args, "project", None)
    if raw is None:
        return None
    return [raw] if isinstance(raw, str) else list(raw)


def handle_common(ctx) -> Outcome:
    args = ctx.args
    projects = _project_list(args)
    exclude = getattr(args, "exclude", None)
    show_all = bool(getattr(args, "all", False))
    engine = _engine(ctx)
    with ctx.progress.stage("common 解析"):
        analyses, warnings = _common_analyses(engine, projects, exclude)

    if not analyses:
        common_dir = abs_path(ctx.config.get("defaults", {}).get("common_dir", "src/common"))
        warnings.insert(0, f"公共库未配置：{common_dir} 下没有可盘的 .bas。"
                           "这是配置状态，不是代码异常；只读体检按警告处理（exit 3）。")
        human = ["  配置位置：mapping.json 的 defaults.common_dir；"
                 "没有公共库的项目可继续用 overview / project / structure"]
        return Outcome(
            action="audit.common",
            ok=True,
            details=[],
            errors=[],
            stats={"configured": False, "common_dir": str(common_dir),
                   "config_key": "defaults.common_dir"},
            human=human,
            warnings=warnings,
            exit_code=3,
            elapsed_ms=ctx.elapsed_ms(),
        )

    sums = [engine.summary(r) for r in analyses]
    human = build_common_lines(analyses, sums, projects, exclude, show_all=show_all)
    if getattr(args, "md", None):
        write_common_md(analyses, sums, args.md, projects, exclude)
    return Outcome(
        action="audit.common",
        ok=True,
        details=sums,
        errors=[],
        stats={"configured": True, "modules": len(analyses)},
        human=human,
        warnings=warnings,
        exit_code=0,
        elapsed_ms=ctx.elapsed_ms(),
    )


def handle_project(ctx) -> Outcome:
    args = ctx.args
    name = args.project
    exclude = getattr(args, "exclude", None)
    show_all = bool(getattr(args, "all", False))
    engine = _engine(ctx)
    with ctx.progress.stage("project 解析"):
        proj_files = _project_files(engine, name, exclude)
        deps = engine.dep_rows(name, exclude)
        own = _own_modules(engine, proj_files, exclude)
    human = build_project_lines(name, proj_files, deps, own, engine, show_all=show_all)
    if getattr(args, "md", None):
        write_project_md(name, proj_files, deps, own, args.md)
    return Outcome(
        action="audit.project",
        ok=True,
        details=deps,
        errors=[],
        stats={"project": name, "modules": len(deps),
               "unused": [d["module"] for d in deps if not d["refs"]]},
        human=human,
        exit_code=0,
        elapsed_ms=ctx.elapsed_ms(),
    )


# ---------------------------------------------------------------------------
# 命令登记（entry="audit"；manifest 唯一来源）
# ---------------------------------------------------------------------------
def _shared_args() -> tuple[ArgSpec, ...]:
    return (
        ArgSpec("--exclude", kind="append",
                help="排除文件名含该关键字的源文件（可重复）"),
        ArgSpec("--all", kind="flag", dest="all", help="列出全部明细"),
        ArgSpec("--md", kind="path", help="另写一份 Markdown 报告"),
        ArgSpec("--out", kind="path", help="另写一份 JSON 到文件"),
    )


SPECS: tuple[CommandSpec, ...] = (
    CommandSpec(
        verb="members",
        entry="audit",
        handler="features.audit_a:handle_members",
        mutation=MutationMode.NONE,
        project=ProjectMode.NONE,
        args=(ArgSpec("--module", kind="str", required=True, help="要盘的模块"),
              *_shared_args()),
        supports_json=True,
        exit_codes=(0, 1, 2),
        summary="某模块被谁调用（默认展开 public 之外的成员）",
        ref="references/audit.md#members",
        tests=("audit_a.members.success", "audit_a.members.not_found",
               "audit_a.members.usage", "audit_a.members.json_pure"),
    ),
    CommandSpec(
        verb="common",
        entry="audit",
        handler="features.audit_a:handle_common",
        mutation=MutationMode.NONE,
        # --project 是 append 过滤项（可重复，纳入多张表）；不走 registry 按
        # ProjectMode 生成的单值通用位，显式登记可重复参数，口径对齐 1.5.6。
        project=ProjectMode.NONE,
        args=(
            ArgSpec("--project", kind="append",
                    help="只纳入这些项目（可重复）；公共库内部调用仍计入"),
            *_shared_args(),
        ),
        supports_json=True,
        exit_codes=(0, 1, 2, 3),
        summary="公共库全部模块的引用数/体积/建议",
        ref="references/audit.md#common",
        tests=("audit_a.common.success", "audit_a.common.not_configured_exit3",
               "audit_a.common.usage", "audit_a.common.json_pure"),
    ),
    CommandSpec(
        verb="project",
        entry="audit",
        handler="features.audit_a:handle_project",
        mutation=MutationMode.NONE,
        project=ProjectMode.REQUIRED,
        args=_shared_args(),
        supports_json=True,
        exit_codes=(0, 1, 2),
        summary="某个工作簿的依赖矩阵",
        ref="references/audit.md#project",
        tests=("audit_a.project.success", "audit_a.project.not_found",
               "audit_a.project.usage", "audit_a.project.json_pure"),
    ),
)
