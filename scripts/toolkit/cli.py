# -*- coding: utf-8 -*-
"""
toolkit.cli —— 三个入口脚本共用的命令行骨架
============================================================
  · add_interactive_arg / add_no_input_arg：统一注册 -i 与 --no-input
  · 菜单只认显式 -i；裸跑（任何环境）都打印用法并退出 2
  · ensure_interactive：进菜单前确认有人能答话，否则给用法并退出 2
  · run_cli：菜单桩派发骨架 —— Failure 报错误返回 1，Ctrl+C 返回 130，用法错用 2
  · run_dispatch：三入口派发兜底 —— Failure / Ctrl+C / 未预期异常统一翻译后走 render 出口
  · entry：模块入口 —— sys.exit(main())，输出被管道提前关闭时静默退出 0

给 agent / 流水线用：裸跑不会卡在 input()；也可以显式 `--no-input`，
让所有 ask_yes 一律按"不确认"处理。
"""

from __future__ import annotations

import sys
import traceback
from typing import Callable

from .base import Failure, entry_command, fail, interactive_ok, msg, set_non_interactive, tool_version
from .contract import ErrorCode, Failure as ContractFailure, Outcome, Stage, to_contract_failure
from .render import render


def add_version_arg(parser, version: str | None = None) -> None:
    """统一注册 --version；不传版本时读包根 VERSION。"""
    version = version or tool_version()
    parser.add_argument("--version", action="version",
                        version=f"vba-workbench {version}")


def add_interactive_arg(parser, help: str = "进入中文交互菜单（需要终端；管道里跑会被拒绝）") -> None:
    parser.add_argument("-i", "--interactive", action="store_true", help=help)


def add_no_input_arg(parser, help: str = "非交互：任何 y/N 提问一律按默认（不确认）处理") -> None:
    parser.add_argument("--no-input", action="store_true", help=help)


def apply_no_input(args) -> None:
    """main 开头调用：把 --no-input 落到全局开关上。"""
    if getattr(args, "no_input", False):
        set_non_interactive(True)


def ensure_interactive() -> None:
    """进菜单前把关：没有终端就直说怎么用命令行，别卡在 input() 上。"""
    if not interactive_ok():
        raise Failure(
            "当前没有可交互终端（stdin 不是 TTY）：交互菜单用不了。"
            f"请改用子命令，例如 `{entry_command('sync.py', 'list')}`；"
            "全部子命令见 `--help`。"
        )


def run_cli(fn) -> int:
    """执行一个返回退出码的派发函数，异常统一翻译成人话。"""
    try:
        return fn()
    except Failure as e:
        fail(str(e))
        if getattr(e, "hint", ""):
            fail(f"说明：{e.hint}")
        if getattr(e, "suggested", ""):
            fail(f"建议：{e.suggested}")
        return 1
    except KeyboardInterrupt:
        msg("已中断")
        return 130
    except BrokenPipeError:
        # 输出被管道提前关闭时静默退出；其它 OSError 仍按未预期异常报出来
        return 0
    except Exception as e:  # noqa: BLE001 - 非 JSON 入口也必须有人话兜底
        traceback.print_exc()
        fail(f"未预期异常：{e}")
        fail("请保留上面 traceback 的最后 20 行；加 --json 重跑可拿到 UNEXPECTED 错误码。")
        return 1


def run_dispatch(action: str, json_mode: bool, *, body: Callable[[], Outcome],
                 out_path: str | None = None,
                 elapsed: Callable[[], int] | None = None,
                 unexpected_prefix: str = "") -> int:
    """三入口共用的派发兜底：``body()`` 正常返回 Outcome，异常统一翻译后走唯一 render 出口。

    与旧三入口各自的 try/except 逐条对齐：

    · ``Failure``：人话模式 stderr 打 ``[错误] <msg>`` + 可选 ``说明：<hint>`` /
      ``建议：<suggested>`` 并退出 1；JSON 模式造业务失败信封，错误码由
      ``to_contract_failure`` 归类（自带码沿用，缺码归 UNEXPECTED）。
    · ``KeyboardInterrupt``：打「已中断」并退出 130，不出信封。
    · 未预期异常：traceback 进 stderr；人话模式再给两句兜底并退出 1；
      JSON 模式造 UNEXPECTED 信封。

    入口间的差异只走参数：``action`` 信封动作名、``out_path``（audit ``--out``）、
    ``elapsed``（失败信封计时；audit 取 ``ctx.elapsed_ms()``，其余恒 0）、
    ``unexpected_prefix``（audit 信封消息沿用旧前缀，其余为裸异常文本）。
    """
    elapsed = elapsed or (lambda: 0)
    try:
        outcome = body()
    except Failure as e:
        if not json_mode:
            fail(str(e))
            if getattr(e, "hint", ""):
                fail(f"说明：{e.hint}")
            if getattr(e, "suggested", ""):
                fail(f"建议：{e.suggested}")
            return 1
        outcome = Outcome(
            action=action, ok=False, exit_code=1, elapsed_ms=elapsed(),
            errors=[to_contract_failure(e)])
    except KeyboardInterrupt:
        msg("已中断")
        return 130
    except Exception as e:  # noqa: BLE001 入口兜底：未预期异常也要给结构化错误
        traceback.print_exc()
        if not json_mode:
            fail(f"未预期异常：{e}")
            fail("请保留上面 traceback 的最后 20 行；加 --json 重跑可拿到 UNEXPECTED 错误码。")
            return 1
        message = f"{unexpected_prefix}{e}" if unexpected_prefix else str(e)
        outcome = Outcome(
            action=action, ok=False, exit_code=1, elapsed_ms=elapsed(),
            errors=[ContractFailure(
                message, code=ErrorCode.UNEXPECTED, stage=Stage.UNKNOWN,
                hint="保留 stderr 最后 20 行；加 --json 重跑并把 errors[0] 交给维护者")])

    return render(outcome, json_mode=json_mode, out_path=out_path)


def entry(main_fn) -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    try:
        sys.exit(main_fn())
    except OSError:
        # 输出被管道提前关闭（例如 | Select-Object -First N）时静默退出，不打印 traceback
        sys.exit(0)
