# -*- coding: utf-8 -*-
r"""agent.py 入口壳（1.6 数据驱动）

本文件不含任何动词特例：命令清单来自各 features 模块登记的 CommandSpec，
流程固定为 ``build_parser(specs_for("agent")) → AppContext.build → dispatch → render``。

· 裸跑（不带动词）：打印帮助并退出 2；未知动词由 argparse 退出 2
· ``--json``：输出一行 JSON 信封；``--apply`` 是唯一写开关（仅写类动词登记该参数）
· 动词清单与参数见 references/agent.md；业务 handler 在 scripts/features/
"""
from __future__ import annotations

import sys
from pathlib import Path

_SCRIPTS = Path(__file__).resolve().parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

from toolkit.base import set_json_mode, set_tool_name  # noqa: E402
from toolkit.cli import (add_version_arg, apply_no_input, entry,  # noqa: E402
                          run_dispatch)
from toolkit.commands import specs_for  # noqa: E402
from toolkit.context import AppContext  # noqa: E402
from toolkit.registry import build_parser, dispatch  # noqa: E402


def main(argv: list[str] | None = None) -> int:
    set_tool_name("agent")
    specs = specs_for("agent")
    parser = build_parser(
        specs, prog="agent.py",
        description="Agent 专用统一入口（无交互菜单，只出结论和退出码）。裸跑只打印帮助并退出 2。")
    add_version_arg(parser)
    args = parser.parse_args(argv)

    # 解析完立即落 JSON 开关：进入 handler / 加载配置之前，warn 就已经收口。
    json_mode = (bool(getattr(args, "json", False))
                 and not bool(getattr(args, "_reject_json", False)))
    set_json_mode(json_mode)

    apply_no_input(args)

    if not getattr(args, "command", None):
        parser.print_help()
        print("\n没有动词：只打印帮助。需要菜单请用 sync.py。")
        return 2

    action = f"agent.{args.command}"

    def _body():
        ctx = AppContext.build("agent", args)
        return dispatch(specs, args, ctx)

    return run_dispatch(action, json_mode, body=_body)


if __name__ == "__main__":
    entry(main)
