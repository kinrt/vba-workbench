# -*- coding: utf-8 -*-
"""从 toolkit.commands 的 manifest 生成命令文档表。

设计原则（对齐设计方案 §4.6：文档从契约生成，manifest 是唯一源）：

- 每个入口（sync / audit / agent）一张 Markdown 表；
- 列：动词、参数、变更模式、退出码、--json、对应测试、参考；
- 同一输入两次生成的字节完全一致（幂等）；
- manifest 为空、入口未知、动词/参数登记不完整、重复登记一律非零退出，
  不允许静默出空表或残缺表。

本脚本只做文档生成，不写任何运行时业务逻辑；可被测试以函数方式调用，
也可在命令行直接跑：

    python tests/tools/gen_commands_md.py [--out references/commands.md] [--strict]

测试调用：``render_markdown(specs)`` / ``build_rows(specs)``，不写入文件。
"""
from __future__ import annotations

import argparse
import sys
import time
from dataclasses import dataclass
from pathlib import Path

PKG = Path(__file__).resolve().parents[2]
SCRIPTS = PKG / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

# 三个用户可见入口；入口名与 CommandSpec.entry 对齐。
ENTRIES = ("sync", "audit", "agent")
ENTRY_FILE = {"sync": "sync.py", "audit": "audit.py", "agent": "agent.py"}

MODE_LABEL = {
    "none": "只读",
    "preview": "默认预览，--apply 才写入文件",
    "required": "写操作，必须 --apply",
    "conditional": "主操作只读，附带开关写入文件",
    "continuous": "常驻运行（不支持 --json）",
}

# 骨架生成 parser 时统一附加的公共参数；文档表的「参数」列只列业务参数，
# 公共参数在表注里统一说明，避免每个动词重复一份。
COMMON_ARGS = ("--json", "--no-input", "--apply", "--project", "--all-project")


class ManifestError(RuntimeError):
    """manifest 不合法：空表、重复、未知入口、登记残缺。生成器直接失败退出。"""


@dataclass(frozen=True)
class Row:
    """一条文档表行；字段顺序即 Markdown 列顺序。"""

    entry: str
    verb: str
    args: str
    mode: str
    exit_codes: str
    json: str
    tests: str
    ref: str
    summary: str


def _validate(specs: tuple) -> None:
    """校验 manifest：非空、入口已知、字段齐全、动词不重复。"""
    if not specs:
        raise ManifestError(
            "manifest 为空：没有任何已登记命令。命令表不得静默生成空表；"
            "请先在 features 模块的 SPECS 里登记动词。"
        )
    seen: dict[str, str] = {}
    for spec in specs:
        key = f"{spec.entry}.{spec.verb}"
        if spec.entry not in ENTRIES:
            raise ManifestError(f"未知入口 {spec.entry!r}（动词 {spec.verb}）；合法入口：{ENTRIES}")
        if not spec.verb or not spec.handler:
            raise ManifestError(f"登记残缺：{key} 缺少 verb/handler")
        if key in seen:
            raise ManifestError(f"命令重复登记：{key}（{seen[key]}）")
        seen[key] = spec.handler
        codes = tuple(spec.exit_codes)
        if not codes:
            raise ManifestError(f"{key} 未登记退出码")
        for arg in spec.args:
            if not arg.name:
                raise ManifestError(f"{key} 存在无名参数")


def _format_arg(arg) -> str:
    """单个业务参数的文档写法；choice 附取值范围。"""
    text = arg.name
    if arg.kind == "flag":
        return f"`{text}`"
    if arg.kind == "choice" and arg.choices:
        return f"`{text}`（{'|'.join(arg.choices)}）"
    if arg.kind == "append":
        return f"`{text}`（可重复）"
    if arg.kind == "int":
        return f"`{text}`（整数）"
    if arg.kind == "float":
        return f"`{text}`（小数）"
    if arg.kind == "path":
        return f"`{text}`（路径）"
    return f"`{text}`"


def _project_label(spec) -> str:
    mode = str(spec.project.value if hasattr(spec.project, "value") else spec.project)
    return {
        "none": "—",
        "optional": "`--project` 可选",
        "required": "`--project` 必填",
        "any": "`--project`/`--all-project` 二选一",
    }.get(mode, mode)


def build_rows(specs: tuple) -> list[Row]:
    """把 manifest 规范成文档行；先校验再排序（入口、动词）。"""
    _validate(specs)
    rows: list[Row] = []
    for spec in specs:
        arg_parts = [_format_arg(a) for a in spec.args]
        proj = _project_label(spec)
        if proj != "—":
            arg_parts.append(proj)
        mode_value = str(spec.mutation.value if hasattr(spec.mutation, "value") else spec.mutation)
        rows.append(Row(
            entry=spec.entry,
            verb=spec.verb,
            args="；".join(arg_parts) or "—",
            mode=MODE_LABEL.get(mode_value, mode_value),
            exit_codes="/".join(str(c) for c in spec.exit_codes),
            json="是" if spec.supports_json else "否",
            tests=", ".join(spec.tests) if spec.tests else "—",
            ref=spec.ref or "—",
            summary=spec.summary or "",
        ))
    rows.sort(key=lambda r: (r.entry, r.verb))
    return rows


def _render_entry(entry: str, rows: list[Row]) -> list[str]:
    """渲染单个入口的小节；行内不含时间戳/路径，保证可重放。"""
    lines: list[str] = [
        f"### {ENTRY_FILE.get(entry, entry)}（{entry}）",
        "",
        "| 动词 | 参数 | 变更模式 | 退出码 | --json | 对应测试 | 参考 |",
        "|---|---|---|---|---|---|---|",
    ]
    for row in rows:
        verb = f"`{row.verb}`"
        if row.summary:
            verb = f"{verb}<br>{row.summary}"
        lines.append(
            f"| {verb} | {row.args} | {row.mode} | {row.exit_codes} | "
            f"{row.json} | {row.tests} | {row.ref} |"
        )
    lines.append("")
    return lines


def render_markdown(specs: tuple) -> str:
    """生成完整 Markdown 文本；纯函数，两次调用结果一致。"""
    rows = build_rows(specs)
    by_entry: dict[str, list[Row]] = {e: [] for e in ENTRIES}
    for row in rows:
        by_entry[row.entry].append(row)

    lines: list[str] = [
        "# 命令表（由 manifest 生成，请勿手改）",
        "",
        "> 本文件由 `tests/tools/gen_commands_md.py` 从 `scripts/toolkit/commands.py` 的",
        "> manifest 生成；动词、参数、退出码以登记为准。重新生成：",
        "",
        "```",
        "python tests/tools/gen_commands_md.py",
        "```",
        "",
        "公共参数（各动词由骨架统一附加，不在下表重复）：",
        "",
        "- `--no-input`：非交互，缺参直接失败；",
        "- `--apply`：写操作才有的写入开关，不加只预览；",
        "- `--project` / `--all-project`：项目选择，取值见各行；",
        "- `--json`：输出一行 JSON 信封；标注「否」的动词显式拒绝该开关。",
        "",
        f"已登记动词合计：{len(rows)} 个。",
        "",
    ]
    for entry in ENTRIES:
        entry_rows = by_entry[entry]
        if entry_rows:
            lines.extend(_render_entry(entry, entry_rows))
    return "\n".join(lines).rstrip() + "\n"


def load_specs():
    """导入 toolkit 并返回 manifest；导入失败按 ManifestError 抛出。"""
    try:
        from toolkit import commands as manifest
    except Exception as exc:  # noqa: BLE001 - 生成期要把导入失败显式报出来
        raise ManifestError(f"manifest 导入失败：{exc}") from exc
    return manifest.SPECS


def generate(out_path: Path, *, specs: tuple | None = None) -> int:
    """写命令表到 out_path；返回动词数。"""
    specs = load_specs() if specs is None else specs
    markdown = render_markdown(specs)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    # 幂等：先写到临时文件再替换；内容相同则字节不变。
    tmp = out_path.with_suffix(out_path.suffix + ".tmp")
    tmp.write_text(markdown, encoding="utf-8", newline="\n")
    tmp.replace(out_path)
    return len(build_rows(specs))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="从 manifest 生成命令文档表")
    parser.add_argument("--out", default=str(PKG / "references" / "commands.md"),
                        help="输出 Markdown 路径（默认 references/commands.md）")
    parser.add_argument("--strict", action="store_true",
                        help="空 manifest 也判失败（默认同样失败；保留供 CI 显式声明）")
    args = parser.parse_args(argv)

    started = time.perf_counter()
    try:
        count = generate(Path(args.out))
    except ManifestError as exc:
        print(f"[gen_commands_md] 失败：{exc}", file=sys.stderr)
        return 2
    elapsed = (time.perf_counter() - started) * 1000
    print(f"[gen_commands_md] 已生成 {args.out}；覆盖动词 {count} 个；"
          f"用时 {elapsed:.1f} ms")
    return 0


if __name__ == "__main__":
    sys.exit(main())
