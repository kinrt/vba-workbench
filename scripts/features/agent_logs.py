# -*- coding: utf-8 -*-
"""agent_io 切片的日志汇总段：logs 动词——最新测试日志摘要与尾部截取。

职责边界：
  · 日志目录定位、GBK 读取、通过/失败计数、尾部行截取全部在本模块；
  · 项目解析委托 ``agent_probe`` 既有原语，不复制第二份；
  · 本模块不碰 COM、不写回工作簿。
"""
from __future__ import annotations

import re
import time
from pathlib import Path

from toolkit.base import entry_command
from toolkit.contract import ErrorCode, Failure as ContractFailure, Outcome, Stage
from features.agent_probe import _find  # noqa: E402

PASS_RE = re.compile(r"通过\s*[:：]?\s*(\d+)")
FAIL_RE = re.compile(r"失败\s*[:：]?\s*(\d+)")


def _log_dir(proj: dict) -> Path:
    """断言日志目录：<search_dir>/测试输出，没有就退回上级/测试输出。"""
    base = Path(proj.get("search_dir") or ".")
    for cand in (base / "测试输出", base.parent / "测试输出"):
        if cand.is_dir():
            return cand
    return base / "测试输出"


def _read_log_text(log: Path) -> str:
    try:
        return log.read_text(encoding="gbk", errors="replace")
    except OSError as e:
        raise ContractFailure(
            f"日志读不了：{log}（{e}）",
            code=ErrorCode.ASSERT_LOG_UNREADABLE, stage=Stage.ASSERT,
            hint="检查日志是否被占用/权限不足；关闭占用程序后重跑",
            retryable=False,
        )


def logs(ctx) -> Outcome:
    # 懒加载入口层 ``_resolve``：避免与 ``agent_io`` 形成导入期循环。
    from features.agent_io import _resolve
    args = ctx.args
    proj = _find(ctx, getattr(args, "project", None))
    _path = _resolve(proj)
    folder = Path(args.log) if getattr(args, "log", "") else _log_dir(proj)
    found = sorted(folder.glob("*.log"), key=lambda f: f.stat().st_mtime, reverse=True)

    if not found:
        error = ContractFailure(
            f"{folder} 下没有日志",
            code=ErrorCode.LOGS_NOT_FOUND, stage=Stage.LOGS,
            hint="先 run --apply（宏内写日志），或用 --log 指定目录；--tail 0 只看摘要",
            suggested=entry_command("agent.py", "assert",
                                    "--project", proj["name"], "--apply"),
        )
        return Outcome(
            action="agent.logs",
            ok=False,
            details=[],
            errors=[error],
            stats={"project": proj["name"], "logs": 0},
            exit_code=1,
            elapsed_ms=ctx.elapsed_ms(),
        )

    log = found[0]
    text = _read_log_text(log)
    passed, failed = PASS_RE.search(text), FAIL_RE.search(text)
    passed_n = passed.group(1) if passed else ""
    failed_n = failed.group(1) if failed else ""
    human: list[str] = [
        f"{log.name}  {time.strftime('%m-%d %H:%M', time.localtime(log.stat().st_mtime))}"]
    if passed and failed:
        human.append(f"通过 {passed.group(1)} / 失败 {failed.group(1)}")
    tail = max(0, int(getattr(args, "tail", 40) or 0))
    tail_lines = [ln.strip()[:160] for ln in text.splitlines()[-tail:]] if tail else []
    human.extend(f"  {ln}" for ln in tail_lines)

    return Outcome(
        action="agent.logs",
        ok=True,
        details=[{"project": proj["name"], "file": log.name,
                   "passed": passed_n, "failed": failed_n, "lines": tail_lines}],
        errors=[],
        stats={"project": proj["name"], "logs": len(found), "tail": tail},
        human=human,
        exit_code=0,
        elapsed_ms=ctx.elapsed_ms(),
    )
