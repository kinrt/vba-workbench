# -*- coding: utf-8 -*-
"""运行上下文：配置、路径基准、JSON 模式、阶段计时、进度输出。"""
from __future__ import annotations

import argparse
import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator

from .base import ROOT, set_assume_yes, set_json_mode, set_tool_name
from .config import load_config


@dataclass
class Progress:
    """长任务阶段计时；只写 stderr，绝不污染 JSON stdout。"""

    tool: str
    enabled: bool = True
    timings: dict[str, int] = field(default_factory=dict)

    @contextmanager
    def stage(self, name: str) -> Iterator[None]:
        import sys

        started = time.perf_counter()
        if self.enabled:
            print(f"[{self.tool}] {name} ...", file=sys.stderr, flush=True)
        try:
            yield
        finally:
            elapsed = int((time.perf_counter() - started) * 1000)
            self.timings[name] = elapsed
            if self.enabled:
                print(f"[{self.tool}] {name} 完成 {elapsed} ms", file=sys.stderr, flush=True)


@dataclass
class AppContext:
    tool: str
    args: argparse.Namespace
    config: dict
    root: Path
    json_mode: bool
    started: float
    progress: Progress

    @classmethod
    def build(cls, tool: str, args: argparse.Namespace) -> "AppContext":
        json_mode = bool(getattr(args, "json", False)) and not bool(getattr(args, "_reject_json", False))
        set_tool_name(tool)
        set_json_mode(json_mode)
        if bool(getattr(args, "apply", False)):
            set_assume_yes(True)
        # init 是引导动词：首次建 mapping.json 时配置还不存在，旧版入口对
        # 它跳过 load_config，这里保持同一豁免，否则首建会被配置缺失挡住。
        verb = getattr(args, "command", None) or getattr(args, "cmd", None) or ""
        config = {} if (tool == "sync" and verb == "init") else load_config()
        progress = Progress(tool=tool, enabled=not json_mode)
        return cls(tool=tool, args=args, config=config, root=ROOT,
                   json_mode=json_mode, started=time.perf_counter(), progress=progress)

    def elapsed_ms(self) -> int:
        return int((time.perf_counter() - self.started) * 1000)

    def stage_timings(self) -> dict[str, int]:
        return dict(self.progress.timings)
