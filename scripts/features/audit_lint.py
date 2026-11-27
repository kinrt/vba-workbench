# -*- coding: utf-8 -*-
"""audit lint：四条高置信静态规则的只读适配层。

规则实现全部在 ``toolkit.vba_lint``；本层只做项目文件枚举、结果装配与人话/JSON 出口。
刻意不扩成 10 条，也不做自动修复：默认 warning 退出 3，--strict 才退出 1。
"""
from __future__ import annotations

from toolkit.base import read_source_text
from toolkit.contract import ErrorCode, Failure, MutationMode, Outcome, Stage
from toolkit.files import write_text_utf8
from toolkit.registry import ArgSpec, CommandSpec, ProjectMode
from toolkit.repo import module_name_of, rel_path
from toolkit.vba_lint import lint_text

from .audit_engine import AuditEngine


def _as_list(value) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return list(value)
    return [str(value)]


def _write_md(findings: list[dict], path: str, project: str, files: int) -> None:
    out = ["# audit lint（四条规则）", "",
           f"- 项目：{project}", f"- 文件：{files}",
           f"- 命中：{len(findings)}", "",
           "| 文件 | 行 | 过程 | 规则 | 说明 |", "|---|---:|---|---|---|"]
    out += [f"| `{f['file']}` | {f['line']} | {f.get('procedure', '')} | "
            f"{f['rule']} | {f['message']} |" for f in findings]
    write_text_utf8(path, "\n".join(out) + "\n")


def lint(ctx) -> Outcome:
    args = ctx.args
    project = str(getattr(args, "project", "") or "")
    exclude = _as_list(getattr(args, "exclude", None))
    strict = bool(getattr(args, "strict", False))

    with ctx.progress.stage("lint"):
        engine = AuditEngine(ctx.root, ctx.config)
        files = engine.project_files(project, exclude)
        if not files:
            known = "、".join(p.get("name", "") for p in engine.projects) or "（src/projects 为空）"
            raise Failure(f"项目「{project}」没有可 lint 的源码模块；可用：{known}",
                          code=ErrorCode.PROJECT_NOT_FOUND, stage=Stage.RESOLVE,
                          hint="核对 mapping.json 的 src，或先跑 sync.py pull --project <项目> --apply")
        findings: list[dict] = []
        for f in files:
            try:
                text = read_source_text(f)
            except Exception as e:  # 单文件读失败上报，不静默跳过
                raise Failure(f"源码读不出：{rel_path(f)}（{e}）",
                              code=ErrorCode.SOURCE_PREFLIGHT, stage=Stage.PARSE) from e
            findings.extend(lint_text(text, module=module_name_of(f), file=rel_path(f)))

    by_rule = {k: 0 for k in ("L01_UNUSED_LOCAL", "L02_DUPLICATE_DECL",
                              "L03_MISSING_RETURN", "L04_UNQUALIFIED_RANGE_CELLS")}
    for f in findings:
        by_rule[f["rule"]] = by_rule.get(f["rule"], 0) + 1

    code = 0 if not findings else (1 if strict else 3)
    errors: list[Failure] = []
    if code == 1:
        errors.append(Failure(f"lint 命中 {len(findings)} 处，按 --strict 失败",
                              code=ErrorCode.SOURCE_PREFLIGHT, stage=Stage.PARSE))

    if getattr(args, "md", None):
        _write_md(findings, str(args.md), project, len(files))

    human = ["静态 lint（4 条高置信规则）",
             f"  项目：{project}  文件 {len(files)}  命中 {len(findings)}",
             "  规则：未使用局部变量 / 同作用域重复声明 / 函数未赋值返回 / 裸 Range·Cells"]
    for f in findings[:60]:
        human.append(f"  - {f['file']}:{f['line']}  [{f['rule']}] {f['message']}")
    if len(findings) > 60:
        human.append(f"  …还有 {len(findings) - 60} 处，加 --md 看完整报告")

    return Outcome(
        action="audit.lint",
        ok=code != 1,
        details=findings,
        errors=errors,
        stats={"project": project, "files": len(files), "findings": len(findings),
               "by_rule": by_rule, "timings": ctx.stage_timings()},
        human=human,
        exit_code=code,
        elapsed_ms=ctx.elapsed_ms(),
    )


SPECS: tuple[CommandSpec, ...] = (
    CommandSpec(
        entry="audit",
        verb="lint",
        handler="features.audit_lint:lint",
        mutation=MutationMode.NONE,
        project=ProjectMode.REQUIRED,
        summary="四条高置信静态 lint（只读，默认 warning 退出 3）",
        ref="references/audit.md#lint",
        exit_codes=(0, 1, 2, 3),
        args=(
            ArgSpec("--exclude", kind="append",
                    help="排除文件名含该关键字的源文件（可重复）"),
            ArgSpec("--strict", kind="flag",
                    help="命中即退出 1（默认只警告退出 3）"),
            ArgSpec("--all", kind="flag", help="兼容 audit 公共参数；本动词总是输出全部"),
            ArgSpec("--md", kind="path", help="另写一份 Markdown 报告"),
            ArgSpec("--out", kind="path", help="另写一份 JSON 到文件"),
        ),
        tests=("tests/test_audit_lint.py",),
    ),
)
