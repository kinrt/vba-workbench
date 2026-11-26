# -*- coding: utf-8 -*-
"""命令 manifest 的数据形状与通用 argparse 生成器。"""
from __future__ import annotations

import argparse
import importlib
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Callable

from .contract import MutationMode, Outcome


class ProjectMode(str, Enum):
    NONE = "none"          # 不接受 --project
    OPTIONAL = "optional"  # 可给 --project
    REQUIRED = "required"  # 必须给 --project
    ANY = "any"            # --project 与 --all-project 二选一，必须给一个


@dataclass(frozen=True)
class ArgSpec:
    """一个业务参数；公共参数由 build_parser 统一加。"""

    name: str                       # --module
    kind: str = "str"               # str / int / float / path / flag / choice / append
    dest: str | None = None
    required: bool = False
    default: Any = None
    choices: tuple[str, ...] = ()
    help: str = ""
    action: str | None = None


@dataclass(frozen=True)
class CommandSpec:
    """一个动词的完整登记：参数、变更模式、处理函数、文档/测试引用。"""

    verb: str
    entry: str
    handler: str                    # "features.core:version"
    mutation: MutationMode = MutationMode.NONE
    project: ProjectMode = ProjectMode.NONE
    args: tuple[ArgSpec, ...] = field(default_factory=tuple)
    supports_json: bool = True
    exit_codes: tuple[int, ...] = (0, 1, 2, 3)
    summary: str = ""
    ref: str = ""
    tests: tuple[str, ...] = field(default_factory=tuple)

    def parser_args(self) -> tuple[ArgSpec, ...]:
        return self.args


def _arg_dest(name: str) -> str:
    return name.lstrip("-").replace("-", "_")


def _add_arg(parser: argparse.ArgumentParser, spec: ArgSpec) -> None:
    kwargs: dict[str, Any] = {"help": spec.help}
    if spec.dest:
        kwargs["dest"] = spec.dest
    if spec.default is not None:
        kwargs["default"] = spec.default
    if spec.kind == "flag":
        kwargs["action"] = "store_true"
    elif spec.kind == "append":
        kwargs["action"] = "append"
    elif spec.kind == "int":
        kwargs["type"] = int
    elif spec.kind == "float":
        kwargs["type"] = float
    elif spec.kind == "choice":
        kwargs["choices"] = spec.choices
    elif spec.kind == "path":
        kwargs["type"] = str
    if spec.required:
        kwargs["required"] = True
    if spec.action:
        kwargs["action"] = spec.action
    parser.add_argument(spec.name, **kwargs)


def build_parser(specs: tuple[CommandSpec, ...], *, prog: str,
                 description: str = "") -> argparse.ArgumentParser:
    """从 manifest 生成 argparse；入口不再手写每个动词的参数。"""
    parser = argparse.ArgumentParser(prog=prog, description=description)
    sub = parser.add_subparsers(dest="command", metavar="<动词>")
    for spec in specs:
        p = sub.add_parser(spec.verb, help=spec.summary or spec.verb)
        if spec.project in (ProjectMode.OPTIONAL, ProjectMode.REQUIRED):
            p.add_argument("--project", required=spec.project == ProjectMode.REQUIRED,
                           help="项目名")
        elif spec.project == ProjectMode.ANY:
            group = p.add_mutually_exclusive_group(required=True)
            group.add_argument("--project", help="项目名")
            group.add_argument("--all-project", action="store_true", dest="all_project",
                               help="全部项目")
        if spec.mutation != MutationMode.NONE:
            p.add_argument("--apply", action="store_true", help="真正执行；不加只预览")
        if spec.supports_json:
            p.add_argument("--json", action="store_true", help="输出一行 JSON 信封")
        else:
            # 常驻命令显式拒绝 --json，保证退出 2 而不是静默忽略
            p.add_argument("--json", action="store_true", help=argparse.SUPPRESS)
            p.set_defaults(_reject_json=True)
        p.add_argument("--no-input", action="store_true", help="非交互，缺参直接失败")
        for arg in spec.parser_args():
            _add_arg(p, arg)
        p.set_defaults(_spec=spec)
    return parser


def load_handler(spec: CommandSpec) -> Callable[[Any], Outcome]:
    module_name, _, attr = spec.handler.partition(":")
    if not module_name or not attr:
        raise ValueError(f"handler 必须写成 module:function，当前 {spec.handler!r}")
    module = importlib.import_module(module_name)
    handler = getattr(module, attr)
    if not callable(handler):
        raise TypeError(f"handler 不可调用：{spec.handler}")
    return handler


def dispatch(specs: tuple[CommandSpec, ...], args: argparse.Namespace, ctx: Any) -> Outcome:
    spec = getattr(args, "_spec", None)
    if spec is None:
        raise RuntimeError("缺少命令 manifest：请用 build_parser(manifest) 解析参数")
    return load_handler(spec)(ctx)
