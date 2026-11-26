# -*- coding: utf-8 -*-
"""唯一输出出口：把 ``Outcome`` 渲染成 JSON 信封或人话。"""
from __future__ import annotations

import json
import sys

from .base import cached_warnings, fail, make_envelope, note, reset_warnings
from .contract import ErrorCode, Failure, Outcome, Stage
from .files import dump_json


def _contract_violation(message: str) -> Failure:
    return Failure(message, code=ErrorCode.CONTRACT_VIOLATION, stage=Stage.CONTRACT)


def _merge_warnings(stats: dict, outcome_warnings: list[str],
                    cached: list[str]) -> dict:
    """把 Outcome.warnings 与 base 缓存的警告收口进 ``stats.warnings``。

    只在有警告时落键；既有同名列表键追加，非列表同名键保留不动，不删任何既有键。
    """
    merged = dict(stats)
    warnings = list(outcome_warnings) + list(cached)
    if not warnings:
        return merged
    existing = merged.get("warnings")
    if isinstance(existing, list):
        merged["warnings"] = list(existing) + warnings
    elif existing is None:
        merged["warnings"] = warnings
    return merged


def render(outcome: Outcome, *, json_mode: bool,
           out_path: str | None = None) -> int:
    """输出结果并返回**归一化后的退出码**。

    - JSON 模式：stdout 恰好一行 JSON，warnings 收进 ``stats.warnings``，不写 stderr；
    - 人话模式：human 行写 stdout，warnings 写 stderr；
    - ``out_path``：把与 stdout 完全相同的信封另存一份，避免入口自己再拼一份；
    - ``exit_code=1`` 却没有任何 errors、``exit_code=0`` 却有 errors，都视为契约破坏，
      由 renderer 补一条 ``CONTRACT_VIOLATION``，不允许再出现 ``FAILED``；
    - ``exit_code=2`` 是用法错的合法出口（与 1.5.6 对齐），原样放行、不归一化。
    - 人话兜底：非 0/3 出口且 human 行为空时，errors 逐条补打 ``[错误]`` 三行；
      已有 human 行的失败命令不重复打印。
    """
    errors = list(outcome.errors)
    exit_code = int(outcome.exit_code)

    if exit_code == 0 and errors:
        exit_code = 1
    if exit_code == 1 and not errors:
        errors.append(_contract_violation("命令返回失败，但没有登记错误对象"))
    if exit_code not in (0, 1, 2, 3):
        errors.append(_contract_violation(f"未登记的退出码：{exit_code}"))
        exit_code = 1

    if json_mode:
        # JSON 模式：警告一律不碰 stderr，收口进 stats.warnings；缓存随即清空，
        # 避免同一进程内重复 render 时警告被二次计入。
        stats = _merge_warnings(outcome.stats, outcome.warnings, cached_warnings())
        reset_warnings()
    else:
        for warning in outcome.warnings:
            print(f"[警告] {warning}", file=sys.stderr)
        stats = outcome.stats
        if out_path:
            stats = _merge_warnings(outcome.stats, outcome.warnings, cached_warnings())

    payload = None
    if json_mode or out_path:
        payload = make_envelope(
            outcome.action,
            bool(outcome.ok) and exit_code in (0, 3) and not errors,
            details=outcome.details,
            errors=[error.as_error() for error in errors],
            elapsed_ms=int(outcome.elapsed_ms or 0),
            stats=stats,
            exit_code=exit_code,
        )

    if out_path:
        p = dump_json(out_path, payload)
        if not json_mode:
            note(f"JSON 结果：{p}")

    if not json_mode:
        for line in outcome.human:
            print(line)
        if exit_code not in (0, 3) and not outcome.human and errors:
            # 出口兜底：handler 把失败收进 errors 却没补 human 行时，人话模式不能静默；
            # 文案对齐入口壳：[错误] <msg> + 可选 说明：<hint> / 建议：<suggested>。
            for err in errors:
                row = err.as_error() if hasattr(err, "as_error") else {"message": str(err)}
                fail(row.get("message") or str(err))
                if row.get("hint"):
                    fail(f"说明：{row['hint']}")
                if row.get("suggested_command"):
                    fail(f"建议：{row['suggested_command']}")
        return exit_code

    print(json.dumps(payload, ensure_ascii=False, default=str))
    return exit_code
