# -*- coding: utf-8 -*-
"""audit 只读视角切片 B：structure / indicators / overview。

分层（架构方案 §4）：
  · audit_engine 收口规则判定与汇总（阈值/白名单/分组登记只在引擎一份）；
  · 本模块只编排：取参数 → 调引擎 → 拼 Outcome（details / stats / errors / human）；
  · 不直接调 COM、不自带 JSON、不打印；退出码语义与 1.5.6 完全一致（0 / 1 / 2 / 3）。
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from adapters import excel_gateway
from toolkit.base import entry_command
from toolkit.files import write_text_utf8
from toolkit.contract import ErrorCode, Failure, Outcome, Stage
from toolkit.registry import ArgSpec, CommandSpec, ProjectMode
from toolkit.repo import resolve_target

from .audit_engine import AuditEngine


def _as_list(value: Any) -> list[str]:
    """append 参数兜底：None / 单值都归一为列表。"""
    if value is None:
        return []
    if isinstance(value, list):
        return list(value)
    return [str(value)]


def _engine(ctx: Any) -> AuditEngine:
    return AuditEngine(ctx.root, ctx.config)


# ---------------------------------------------------------------------------
# structure：结构体检（旧副本 / 孤儿壳 / 缺壳 / 重名 / 文档模块对账 / 口径回潮 / 绑定）
# ---------------------------------------------------------------------------

def _structure_human(issues: dict, total: int) -> list[str]:
    lines: list[str] = ["结构体检（占位壳 / 旧副本 / 重名过程 / 表页↔文档模块）"]
    if not total:
        lines.append("  全部正常：没有旧副本、没有孤儿壳、没有缺壳、公共库内也没有重名过程，表页与文档模块一一对应。")
        return lines

    def block(title: str, items: list[str], hint: str) -> None:
        lines.append(f"  {title} {len(items) if items else 0} 处")
        lines.extend(f"     {x}" for x in items[:20])
        lines.append(f"     → {hint}")

    if issues["stale"]:
        block("① 旧副本", [f"{a}  ←→  公共库 {b.name}" for a, b in issues["stale"]],
              entry_command("sync.py", "align", "--project", "<表名>", "--apply"))
    if issues["orphan"]:
        block("② 孤儿壳", [str(a) for a in issues["orphan"]],
              entry_command("sync.py", "shells", "--mode", "clean", "--apply") + "（只删不建）")
    if issues["kind"]:
        block("③ 壳类型不符", [f"{a}（壳 {a.suffix}）←→ 公共库 {b.name}（{b.suffix}）"
                          for a, b in issues["kind"]],
              "删掉这个壳后跑 " + entry_command("sync.py", "shells", "--apply") + " 重建")
    if issues["missing"]:
        block("④ 缺壳", [f"{p} → {m}" for p, m in issues["missing"]],
              entry_command("sync.py", "shells", "--apply") + "（按需生成占位壳）")
    if issues["dup"]:
        block("⑤ 公共库内重名过程", [f"{n}：{'、'.join(owners)}" for n, owners in issues["dup"]],
              "改名，或确认调用方都用「模块名.过程」限定")
    if issues["shadow"]:
        block("⑥ 项目模块与公共库重名过程",
              [f"{n}：{a}  与公共库 {'、'.join(owners)}" for n, a, owners in issues["shadow"]],
              "要么删项目里的旧实现，要么改名；这也是盘点里「疑似同名」的来源")
    if issues["shadow_known"]:
        names = sorted({n for n, _f, _o in issues["shadow_known"]})
        lines.append(f"  （已登记）项目侧同名但无跨模块裸调的 {len(issues['shadow_known'])} 组："
                     + "、".join(names) + "  → 保留原实现，见 T28")
    if issues["doc_orphan"]:
        block("⑦ 源码有、工程无的文档模块", [f"{p} → {m}" for p, m in issues["doc_orphan"]],
              entry_command("sync.py", "pull", "--all-project", "--apply"))
    if issues["doc_ungrabbed"]:
        block("⑧ 工程有、源码无的表页模块", [f"{p} → {m}" for p, m in issues["doc_ungrabbed"]],
              entry_command("sync.py", "pull", "--all-project", "--apply") + "（补抓）")
    if issues.get("naked_new"):
        block("⑨ 裸 New 表格工具", [f"{p}:{i}  {s[:70]}" for p, i, s in issues["naked_new"]],
              "改成 Set 工具 = getTL(表页, \"表名\")（复用实例并登记日志宿主）")
    if issues.get("manual_accel"):
        block("⑩ 手写加速", [f"{p}:{i}  {s[:70]}" for p, i, s in issues["manual_accel"]],
              "改成 Call 开启加速(...) / Call 关闭加速；例外块上方标 ' [加速例外]")
    if issues.get("no_comment"):
        block("⑪ 缺过程注释", [f"{p}:{i}  {s[:70]}" for p, i, s in issues["no_comment"]],
              "在第一行上方补 ' 函数名／功能／参数 三行")
    if issues.get("suffix"):
        block("⑫ 后缀模块", [f"{p} → {n}（原名 {b} 还在）" for p, n, b in issues["suffix"]],
              "在 VBE 里删掉带序号那个")
    if issues.get("sheet_mismatch"):
        block("⑬ 表页与文档模块不对应（H19）",
              [f"{p} → [{k}] {o}：{d}" for p, k, o, d, _h in issues["sheet_mismatch"]],
              "数量不等先查表页/模块配对；空壳文档模块换干净 vbaProject.bin 后重植（Q89）")
    if issues.get("binding_target"):
        block("⑭ 快捷键/动作绑定目标",
              [f"{p} → {f}:{i}  {s}" for p, f, i, s in issues["binding_target"]],
              "核对目标是否 Public 零必填参数过程；窗体要在 窗体实例 登记；可 --ignore binding_target")
    return lines


def _write_structure_md(issues: dict, path: str) -> None:
    total = AuditEngine.issues_total(issues)
    out = ["# 结构体检", ""]
    if not total:
        out.append("全部正常：没有旧副本、孤儿壳、缺壳，也没有重名过程；表页与文档模块一一对应。")
        write_text_utf8(path, "\n".join(out))
        return
    if issues["stale"]:
        out += ["## 旧副本", "", "项目里是真实代码，公共库也有同名 → `sync.py align --project <表名> --apply`", ""]
        out += [f"- `{a}` ←→ 公共库 `{b.name}`" for a, b in issues["stale"]]
    if issues["orphan"]:
        out += ["", "## 孤儿壳", "", "壳还在，公共库已无此模块 → `sync.py shells --mode clean --apply`", ""]
        out += [f"- `{a}`" for a in issues["orphan"]]
    if issues["kind"]:
        out += ["", "## 壳类型不符", "", "删壳后跑 `sync.py shells --apply` 重建", ""]
        out += [f"- `{a}`（{a.suffix}）←→ 公共库 `{b.name}`（{b.suffix}）" for a, b in issues["kind"]]
    if issues["missing"]:
        out += ["", "## 缺壳", "", "表在用但目录里没有占位壳 → `sync.py shells --apply`", ""]
        out += [f"- {p} → {m}" for p, m in issues["missing"]]
    if issues["dup"]:
        out += ["", "## 公共库内重名过程", ""]
        out += [f"- {n}：{'、'.join(owners)}" for n, owners in issues["dup"]]
    if issues["shadow"]:
        out += ["", "## 项目模块与公共库重名过程", ""]
        out += [f"- {n}：`{a}` 与公共库 {'、'.join(owners)}" for n, a, owners in issues["shadow"]]
    if issues["doc_orphan"]:
        out += ["", "## 源码有、工程没有的文档模块（表页已删）", "",
                "跑 `sync.py pull --all-project --apply` 清掉残留文件", ""]
        out += [f"- {p} → {m}" for p, m in issues["doc_orphan"]]
    if issues["doc_ungrabbed"]:
        out += ["", "## 工程有、源码没有的表页模块（抓取漏了）", "",
                "确认表页存在后跑 `sync.py pull --all-project --apply` 补抓", ""]
        out += [f"- {p} → {m}" for p, m in issues["doc_ungrabbed"]]
    if issues.get("sheet_mismatch"):
        out += ["", "## 表页与文档模块不对应（H19）", "",
                "表页数取 `workbook.xml`，文档模块取工程流 `PROJECT` 的 `Document=`；"
                "空壳文档模块删不掉 → 换干净 vbaProject.bin 后重植（Q89）", ""]
        out += [f"- {p} · [{k}] {o}：{d}" for p, k, o, d, _h in issues["sheet_mismatch"]]
    write_text_utf8(path, "\n".join(out))


def structure(ctx: Any) -> Outcome:
    args = ctx.args
    exclude = _as_list(getattr(args, "exclude", None))
    ignore = _as_list(getattr(args, "ignore", None))
    strict = bool(getattr(args, "strict", False))

    with ctx.progress.stage("structure"):
        engine = _engine(ctx)
        issues = engine.structure_issues(exclude=exclude, ignore=ignore)

    total = AuditEngine.issues_total(issues)
    code = 0 if not total else (1 if strict else 3)
    rows = AuditEngine.structure_rows(issues)

    stats: dict[str, Any] = {"groups": len(rows), "issues": total,
                             "timings": ctx.stage_timings()}
    if total:
        stats["warned"] = total

    errors: list[Failure] = []
    if code == 1:
        errors.append(Failure(f"结构体检发现 {total} 处问题（详见 details）",
                              code=ErrorCode.SOURCE_PREFLIGHT, stage=Stage.PARSE))

    if getattr(args, "md", None):
        _write_structure_md(issues, str(args.md))

    return Outcome(
        action="audit.structure",
        ok=code != 1,
        details=rows,
        errors=errors,
        stats=stats,
        human=_structure_human(issues, total),
        exit_code=code,
        elapsed_ms=ctx.elapsed_ms(),
    )


# ---------------------------------------------------------------------------
# indicators：体检指标 H01–H19
# ---------------------------------------------------------------------------

def _probe_live(projects: list[dict]) -> dict:
    """运行期探测（H16/H17）：经网关附着已开 Excel，只读工程引用与模块行数。

    口径同 1.5.6：体检不替用户开表、不碰密码；没开 Excel 或表没打开的项目算"没检查"，
    由引擎统一标"未跑"，绝不拿别的表顶替。引用/组件事实从网关快照取，通过/警告判定在引擎。
    返回 {"checked", "total", "h16", "h17", "reason"}。
    """
    out: dict[str, Any] = {"checked": 0, "total": len(projects), "h16": [], "h17": [], "reason": ""}
    if not projects:
        out["reason"] = "配置里没有项目"
        return out
    try:
        xl = excel_gateway.attach_running_excel()
    except Exception as e:
        out["reason"] = f"没有正在运行的 Excel：{e}"
        return out

    for proj in projects:
        path, _why = resolve_target(proj)
        if not path:
            continue
        try:
            wb, _opened = excel_gateway.find_open_workbook(xl, path)
        except Exception as e:
            out["checked"] += 1
            out["h16"].append(f"{proj['name']}  工程读不到，无法判 H16：{e}")
            out["h17"].append(f"{proj['name']}  {e}")
            continue
        if wb is None:
            continue
        snap = excel_gateway.read_vbproject_probe(wb)
        out["checked"] += 1
        if snap["error"]:
            out["h16"].append(f"{proj['name']}  工程读不到，无法判 H16：{snap['error']}")
            out["h17"].append(f"{proj['name']}  {snap['error']}")
            continue
        comps = snap["components"]
        need_forms = any(c["type"] == 3 for c in comps)  # 3=UserForm
        if not need_forms:
            need_forms = any(c["type"] == 100 and c["controls"] > 0 for c in comps)
        if need_forms and not any(re.search(r"(?i)forms", n) for n in snap["refs"]):
            out["h16"].append(f"{proj['name']}  有窗体/控件但引用无 Microsoft Forms 2.0")
    return out


def _indicators_human(rows: list[dict]) -> list[str]:
    bad = [r for r in rows if r["state"] == "警告"]
    todo = [r for r in rows if r["state"] in ("未跑", "锁定")]
    passed = [r for r in rows if r["state"] == "通过"]
    lines = ["体检指标复跑（H01–H19）",
             f"  通过 {len(passed)} / 警告 {len(bad)} / 未跑 {len(todo)}（共 {len(rows)} 项）"]
    for r in bad:
        lines.append(f"\n  {r['id']} {r['name']}：{r['count']}")
        if r["detail"]:
            lines.append(f"     {r['detail']}")
        if r.get("suggested_command"):
            lines.append(f"     修复：{r['suggested_command']}")
        for s in r["samples"][:15]:
            lines.append(f"     {s}")
    for r in rows:
        if r["id"] == "H10" and r["state"] == "通过" and r["detail"]:
            lines.append(f"\n  H10 {r['name']}：{r['detail']}")
    for r in todo:
        lines.append(f"\n  {r['id']} {r['name']}（{r['state']}）：{r['detail']}")
    return lines


def _write_indicators_md(rows: list[dict], path: str) -> None:
    passed = sum(1 for r in rows if r["state"] == "通过")
    warned = sum(1 for r in rows if r["state"] == "警告")
    todo = sum(1 for r in rows if r["state"] in ("未跑", "锁定"))
    out = ["# 体检指标复跑（H01–H19）", "",
           f"- 通过 {passed}、警告 {warned}、未跑 {todo}", "",
           "| 编号 | 指标 | 状态 | 数量 | 说明 | 修复 |", "|---|---|---|---|---|---|"]
    out += [f"| {r['id']} | {r['name']} | {r['state']} | {r['count']} | {r['detail']} | "
            f"{r.get('suggested_command', '')} |" for r in rows]
    for r in [r for r in rows if r["samples"]]:
        out += ["", f"## {r['id']} {r['name']}（{len(r['samples'])} 处）", ""]
        out += [f"- {s}" for s in r["samples"]]
    write_text_utf8(path, "\n".join(out) + "\n")


def indicators(ctx: Any) -> Outcome:
    args = ctx.args
    exclude = _as_list(getattr(args, "exclude", None))
    ignore = _as_list(getattr(args, "ignore", None))
    strict = bool(getattr(args, "strict", False))

    with ctx.progress.stage("indicators"):
        engine = _engine(ctx)
        live = _probe_live(engine.projects)
        # 口径同 1.5.6：indicators 不接 --exclude（解析器注册但 handler 不下发）
        rows = engine.indicator_rows(ignore=ignore, live_probe=live)

    warned = sum(1 for r in rows if r["state"] == "警告")
    passed = sum(1 for r in rows if r["state"] == "通过")
    code = 0 if not warned else (1 if strict else 3)

    stats: dict[str, Any] = {"passed": passed, "warned": warned,
                             "timings": ctx.stage_timings()}
    errors: list[Failure] = []
    if code == 1:
        errors.append(Failure(f"{warned} 项指标为警告",
                              code=ErrorCode.SOURCE_PREFLIGHT, stage=Stage.PARSE))

    if getattr(args, "md", None):
        _write_indicators_md(rows, str(args.md))

    return Outcome(
        action="audit.indicators",
        ok=code != 1,
        details=rows,
        errors=errors,
        stats=stats,
        human=_indicators_human(rows),
        exit_code=code,
        elapsed_ms=ctx.elapsed_ms(),
    )


# ---------------------------------------------------------------------------
# overview：公共库汇总 + 结构问题 + 整体意见
# ---------------------------------------------------------------------------

def _overview_human(actions: list[tuple[int, str]], issues: dict, sums: list[dict],
                   projects: list[str], exclude: list[str]) -> list[str]:
    total = AuditEngine.issues_total(issues)
    lines = ["全仓总览：公共库 + 结构问题 + 整体意见",
             f"  公共库  ：{len(sums)} 个模块，成员 {sum(s['members'] for s in sums)}"
             f"（不可达 {sum(s['dead'] for s in sums)}，约 {sum(s['lines'] for s in sums)} 行）",
             f"  结构问题：{total} 处"
             f"（旧副本 {len(issues['stale'])}、孤儿壳 {len(issues['orphan'])}、"
             f"壳类型不符 {len(issues['kind'])}、缺壳 {len(issues['missing'])}、"
             f"重名过程 {len(issues['dup']) + len(issues['shadow'])}、"
             f"文档模块对账 {len(issues['doc_orphan']) + len(issues['doc_ungrabbed'])}、"
             f"表页↔文档模块 {len(issues.get('sheet_mismatch', []))}、"
             f"口径回潮 {len(issues.get('naked_new', [])) + len(issues.get('manual_accel', [])) + len(issues.get('no_comment', []))}、"
             f"后缀模块 {len(issues.get('suffix', []))}、"
             f"绑定目标 {len(issues.get('binding_target', []))}）",
             f"  统计范围：{('项目 ' + '、'.join(projects)) if projects else '公共库 + 全部项目'}"
             + (f"（排除 {'、'.join(exclude)}）" if exclude else ""),
             ""]
    for s in sums:
        lines.append(f"  {s['module']}\t{s['members']}\t{s['dead']}\t{s['test']}\t"
                     f"{s['single']}\t{s['active']}\t{s['lines']}")

    lines.append("")
    lines.append("整体意见（按影响排序；命令都能直接跑）")
    if not actions:
        lines.append("  没有需要处理的问题。")
    for i, (_weight, text) in enumerate(sorted(actions, key=lambda x: -x[0]), 1):
        lines.append(f"  {i}. {text}")
    return lines


def _write_overview_md(analyses: list[dict], issues: dict, sums: list[dict],
                       path: str, projects: list[str], exclude: list[str]) -> None:
    out = ["# 全仓盘点", "",
           f"- 统计范围：{('项目 ' + '、'.join(projects)) if projects else '公共库 + 全部项目'}"
           + (f"；排除 {'、'.join(exclude)}" if exclude else ""),
           f"- 公共库 {len(sums)} 个模块，成员 {sum(s['members'] for s in sums)}，"
           f"不可达 {sum(s['dead'] for s in sums)}（约 {sum(s['lines'] for s in sums)} 行）",
           f"- 结构问题 {AuditEngine.issues_total(issues)} 处", "",
           "| 模块 | 成员 | 不可达 | 仅自测 | 仅单文件 | 活跃 | 可删行 |",
           "|---|--:|--:|--:|--:|--:|--:|"]
    out += [f"| {s['module']} | {s['members']} | {s['dead']} | {s['test']} | {s['single']} | "
            f"{s['active']} | {s['lines']} |" for s in sums]
    write_text_utf8(path, "\n".join(out) + "\n")


def overview(ctx: Any) -> Outcome:
    args = ctx.args
    projects = _as_list(getattr(args, "project", None))
    exclude = _as_list(getattr(args, "exclude", None))
    ignore = _as_list(getattr(args, "ignore", None))
    strict = bool(getattr(args, "strict", False))

    with ctx.progress.stage("overview"):
        engine = _engine(ctx)
        analyses = engine.common_analyses(projects, exclude)
        issues = engine.structure_issues(exclude=exclude, ignore=ignore)

    total = AuditEngine.issues_total(issues)
    fatal = AuditEngine.issues_fatal(issues)
    code = 1 if fatal else 0
    if strict and total and code == 0:
        code = 1

    sums = [engine.summary(a) for a in analyses]
    actions = engine.overview_actions(analyses, issues)
    stats: dict[str, Any] = {
        "modules": len(analyses),
        "common_configured": bool(analyses),
        "structure": {k: len(v) for k, v in issues.items()},
        "issues": total,
        "fatal": fatal,
        "timings": ctx.stage_timings(),
    }

    errors: list[Failure] = []
    if code == 1:
        if fatal:
            errors.append(Failure(f"结构体检发现 {fatal} 处阻断问题（共 {total} 处，详见 structure）",
                                  code=ErrorCode.SOURCE_PREFLIGHT, stage=Stage.PARSE))
        else:
            errors.append(Failure(f"--strict：结构体检共 {total} 处提醒，按失败处理",
                                  code=ErrorCode.SOURCE_PREFLIGHT, stage=Stage.PARSE))

    if getattr(args, "md", None):
        _write_overview_md(analyses, issues, sums, str(args.md), projects, exclude)

    return Outcome(
        action="audit.overview",
        ok=code != 1,
        details=sums,
        errors=errors,
        stats=stats,
        human=_overview_human(actions, issues, sums, projects, exclude),
        exit_code=code,
        elapsed_ms=ctx.elapsed_ms(),
    )


# ---------------------------------------------------------------------------
# 命令登记
# ---------------------------------------------------------------------------

_COMMON_ARGS = (
    ArgSpec("--exclude", kind="append",
            help="排除文件名含该关键字的源文件（可重复）"),
    ArgSpec("--ignore", kind="append",
            help="忽略某分组/指标，如 manual_accel、H05a,H06（可重复）"),
    ArgSpec("--strict", kind="flag",
            help="仅警告也退出 1（默认只警告退出 3）"),
    ArgSpec("--all", kind="flag", help="列出全部明细"),
    ArgSpec("--md", kind="str", help="另写一份 Markdown 报告"),
    ArgSpec("--out", kind="str", help="另写一份 JSON 到文件"),
)

SPECS: tuple[CommandSpec, ...] = (
    CommandSpec(
        entry="audit",
        verb="structure",
        handler="features.audit_b:structure",
        summary="旧副本 / 孤儿壳 / 缺壳 / 重名过程 / 文档模块对账 / 旧规则回退",
        ref="references/audit.md",
        project=ProjectMode.NONE,
        exit_codes=(0, 1, 2, 3),
        args=_COMMON_ARGS,
        tests=("tests/test_audit_b.py",),
    ),
    CommandSpec(
        entry="audit",
        verb="indicators",
        handler="features.audit_b:indicators",
        summary="体检指标 H01–H19 复跑",
        ref="references/audit.md",
        project=ProjectMode.NONE,
        exit_codes=(0, 1, 2, 3),
        args=_COMMON_ARGS,
        tests=("tests/test_audit_b.py",),
    ),
    CommandSpec(
        entry="audit",
        verb="overview",
        handler="features.audit_b:overview",
        summary="公共库 + 结构的汇总意见",
        ref="references/audit.md",
        project=ProjectMode.NONE,
        exit_codes=(0, 1, 2, 3),
        args=(
            ArgSpec("--project", kind="append",
                    help="只纳入这些项目（可重复）；公共库内部调用仍计入"),
            *_COMMON_ARGS,
        ),
        tests=("tests/test_audit_b.py",),
    ),
)
