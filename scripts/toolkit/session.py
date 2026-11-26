# -*- coding: utf-8 -*-
"""工作簿会话骨架：状态、用途、提示矩阵的唯一入口。

真正的 COM 探测由 ``adapters.excel_gateway`` 提供；实现者只负责把探测结果
落到某个 ``TargetState``，然后通过 ``rule()`` 取统一的 code/hint。
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .contract import ErrorCode, Failure
from .state import Purpose, StateRule, TargetState, rule_for


@dataclass
class WorkbookSession:
    target: str
    state: TargetState
    purpose: Purpose
    workbook: Any = None

    def rule(self) -> StateRule | None:
        return rule_for(self.state, self.purpose)

    def require(self) -> None:
        """有拦截规则时抛统一 Failure；警告类规则由调用方决定是否继续。"""
        rule = self.rule()
        if not rule or not rule.code:
            return
        raise Failure(
            f"{self.target}: {rule.code}",
            code=rule.code,
            stage=rule.stage or "attach",
            hint=rule.hint,
            suggested=rule.suggested,
            retryable=rule.retryable,
        )

    def require_no_error(self) -> None:
        self.require()


def missing_target(target: str) -> WorkbookSession:
    return WorkbookSession(target=target, state=TargetState.MISSING, purpose=Purpose.PULL)
