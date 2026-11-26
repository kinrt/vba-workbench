# -*- coding: utf-8 -*-
"""audit.py 入口壳：manifest 驱动解析 → AppContext → dispatch → 单一 render 出口。

本壳不写任何动词特例：动词清单 / 参数 / 处理函数全部来自 features 模块登记的
SPECS（``toolkit.commands.specs_for("audit")``）。新增视角只需在对应 feature 模块
登记 ``CommandSpec``，本壳与 render 出口不需要跟着改。

退出码沿用 1.5.6：0 成功 / 1 业务失败 / 2 用法错（argparse 直接 SystemExit）/
3 只读体检仅有警告（不被归一成 1）。
"""
from __future__ import annotations

import sys
from pathlib import Path

TOOLS = Path(__file__).resolve().parent
sys.path.insert(0, str(TOOLS))

from toolkit.base import set_json_mode, set_tool_name  # noqa: E402
from toolkit.cli import (add_interactive_arg, add_version_arg, apply_no_input,  # noqa: E402
                         ensure_interactive, entry, run_cli, run_dispatch)
from toolkit.commands import specs_for
from toolkit.registry import build_parser, dispatch
from toolkit.contract import ErrorCode, Failure, Stage  # noqa: E402
from toolkit.context import AppContext

TOOL = "audit"


def _interactive() -> int:
    """-i 菜单入口：先把关终端；菜单本体由后续切片接入。"""
    ensure_interactive()
    raise Failure("交互菜单将在后续切片接入；请改用子命令（见 --help）。",
                  code=ErrorCode.UNSUPPORTED_OPERATION, stage=Stage.CONTRACT)


def main(argv=None) -> int:
    set_tool_name(TOOL)
    specs = specs_for(TOOL)
    parser = build_parser(
        specs, prog="audit.py",
        description="VBA 代码盘点（公共库瘦身 + 结构体检，只读）。"
                    "不带动词只打印帮助并退出 2；进菜单用 -i。")
    add_version_arg(parser)
    add_interactive_arg(parser)
    args = parser.parse_args(argv)

    # 解析完立即落 JSON 开关：进入 handler / 加载配置之前，warn 就已经收口。
    json_mode = (bool(getattr(args, "json", False))
                 and not bool(getattr(args, "_reject_json", False)))
    set_json_mode(json_mode)

    apply_no_input(args)

    if not args.command:
        if getattr(args, "interactive", False):
            return run_cli(_interactive)
        parser.print_help()
        print("\n没有视角：只打印帮助。需要菜单请加 -i。")
        return 2

    action = f"{TOOL}.{args.command}"
    state = {"ctx": None}

    def _body():
        state["ctx"] = AppContext.build(TOOL, args)
        return dispatch(specs, args, state["ctx"])

    return run_dispatch(
        action, json_mode, body=_body,
        out_path=getattr(args, "out", None),
        elapsed=lambda: state["ctx"].elapsed_ms() if state["ctx"] is not None else 0,
        unexpected_prefix="未预期异常：")


if __name__ == "__main__":
    entry(main)
