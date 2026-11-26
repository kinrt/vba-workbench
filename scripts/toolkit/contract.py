# -*- coding: utf-8 -*-
"""架构契约层：错误、阶段、变更模式、结果对象。

这一层只放**跨命令共享的数据形状**，不放 Excel/文件/业务逻辑。
旧代码仍可继续使用 ``toolkit.base.Failure``；新功能统一从本模块导入 ``Failure``。
迁移完成后，``base.Failure`` 会退化为本模块的别名或直接删除。
"""
from __future__ import annotations

import re
import subprocess
from dataclasses import dataclass, field
from enum import Enum
from typing import Any

from .base import Failure as BaseFailure

_CODE_RE = re.compile(r"^[A-Z][A-Z0-9_]*$")


class Stage(str, Enum):
    """失败阶段。值保持小写，直接进 JSON 的 stage 字段。"""

    CONFIG = "config"
    PARSE = "parse"
    RESOLVE = "resolve"
    ATTACH = "attach"
    OPEN = "open"
    VBE = "vbe"
    PULL = "pull"
    PUSH = "push"
    RUN = "run"
    PREFLIGHT = "preflight"
    POSTCHECK = "postcheck"
    SAVE = "save"
    PATCH = "patch"
    LOGS = "logs"
    ASSERT = "assert"
    CONTRACT = "contract"
    INIT = "init"
    UNKNOWN = "unknown"


class MutationMode(str, Enum):
    """命令对磁盘/工作簿的影响方式；用来生成 ``--apply`` 门。"""

    NONE = "none"              # 纯只读
    PREVIEW = "preview"        # 默认预览，--apply 才写（init / clean-backup）
    REQUIRED = "required"      # 写动词，必须 --apply
    CONDITIONAL = "conditional"  # 主操作只读/内存，附带开关才落盘（run --save）
    CONTINUOUS = "continuous"  # 持续写（watch），必须有界且不支持 --json


class ErrorCode(str, Enum):
    """已登记的错误码。新增码必须先登记，再使用。"""

    ARG_MISMATCH = "ARG_MISMATCH"
    ASSERT_FAILED = "ASSERT_FAILED"
    ASSERT_LOG_UNREADABLE = "ASSERT_LOG_UNREADABLE"
    ASSERT_MACRO_ARGS = "ASSERT_MACRO_ARGS"
    ASSERT_MACRO_INVALID = "ASSERT_MACRO_INVALID"
    ASSERT_MACRO_NOT_CONFIGURED = "ASSERT_MACRO_NOT_CONFIGURED"
    ASSERT_NO_LOG = "ASSERT_NO_LOG"
    ASSERT_NO_TABLE = "ASSERT_NO_TABLE"
    COM_CONTEXT = "COM_CONTEXT"
    CONFIG_INVALID = "CONFIG_INVALID"
    CONFIG_NOT_FOUND = "CONFIG_NOT_FOUND"
    CONTRACT_VIOLATION = "CONTRACT_VIOLATION"
    EXCEL_BUSY = "EXCEL_BUSY"
    EXCEL_CLOSED = "EXCEL_CLOSED"
    FILE_READONLY = "FILE_READONLY"
    INTERRUPTED = "INTERRUPTED"
    LAB_TEMPLATE_FAILED = "LAB_TEMPLATE_FAILED"
    LOGS_NOT_FOUND = "LOGS_NOT_FOUND"
    MISSING_LAB_TEMPLATE = "MISSING_LAB_TEMPLATE"
    NO_EXCEL = "NO_EXCEL"
    PATCH_BLOCK_AMBIGUOUS = "PATCH_BLOCK_AMBIGUOUS"
    PATCH_FILE_NOT_FOUND = "PATCH_FILE_NOT_FOUND"
    PATCH_FORMAT = "PATCH_FORMAT"
    POSTCHECK_FAILED = "POSTCHECK_FAILED"
    PROBE_FAILED = "PROBE_FAILED"
    PROJECT_NOT_FOUND = "PROJECT_NOT_FOUND"
    PROJECT_NOT_SPECIFIED = "PROJECT_NOT_SPECIFIED"
    PACKAGE_INTEGRITY = "PACKAGE_INTEGRITY"
    PACKAGE_MANIFEST_MISSING = "PACKAGE_MANIFEST_MISSING"
    RUN_FAILED = "RUN_FAILED"
    SAVE_FAILED = "SAVE_FAILED"
    SCRUB_HIT = "SCRUB_HIT"
    SOURCE_PREFLIGHT = "SOURCE_PREFLIGHT"
    TARGET_NOT_ACTIVE = "TARGET_NOT_ACTIVE"
    TARGET_NOT_FOUND = "TARGET_NOT_FOUND"
    TARGET_NOT_OPEN = "TARGET_NOT_OPEN"
    TARGET_OPEN = "TARGET_OPEN"
    UNEXPECTED = "UNEXPECTED"
    UNSUPPORTED_OPERATION = "UNSUPPORTED_OPERATION"
    USAGE = "USAGE"
    VBE_COMPILE_ERROR = "VBE_COMPILE_ERROR"
    VBE_PROTECTED = "VBE_PROTECTED"
    VERSION_MISMATCH = "VERSION_MISMATCH"
    WATCHDOG_TIMEOUT = "WATCHDOG_TIMEOUT"
    WORKBOOK_OPEN_FAILED = "WORKBOOK_OPEN_FAILED"


@dataclass(frozen=True)
class SuggestedCommand:
    """可直接执行的建议命令。第一条 token 必须是可执行命令。"""

    argv: tuple[str, ...]

    def __post_init__(self) -> None:
        if not self.argv or not str(self.argv[0]).strip():
            raise ValueError("SuggestedCommand 至少需要一个 token")
        first = str(self.argv[0]).lower()
        if not (first.endswith(".py") or first.startswith("python")
                or first.startswith("powershell") or first == "attrib"):
            raise ValueError(f"SuggestedCommand 不是可执行命令：{self.argv[0]!r}")

    def __str__(self) -> str:
        return subprocess.list2cmdline([str(x) for x in self.argv])


class Failure(BaseFailure):
    """新架构统一业务失败：code / stage 必填，不再由兜底猜。"""

    def __init__(self, message: str, *, code: str | ErrorCode, stage: str | Stage,
                 hint: str = "", suggested: str | SuggestedCommand = "",
                 retryable: bool = False) -> None:
        code_value = code.value if isinstance(code, ErrorCode) else str(code)
        stage_value = stage.value if isinstance(stage, Stage) else str(stage)
        if not _CODE_RE.match(code_value):
            raise ValueError(f"非法错误码：{code_value!r}（要求大写下划线）")
        if stage_value not in {member.value for member in Stage}:
            raise ValueError(f"非法阶段：{stage_value!r}（必须是 Stage 登记成员）")
        suggested_text = str(suggested) if suggested else ""
        super().__init__(message, code=code_value, stage=stage_value, hint=hint,
                         suggested=suggested_text, retryable=retryable)

    def as_error(self) -> dict[str, Any]:
        return {
            "code": self.code or ErrorCode.CONTRACT_VIOLATION.value,
            "stage": self.stage or Stage.UNKNOWN.value,
            "message": str(self),
            "hint": self.hint,
            "suggested_command": self.suggested,
            "retryable": bool(self.retryable),
        }


def to_contract_failure(e: BaseFailure) -> Failure:
    """旧基类 Failure 统一归类为契约层 Failure（sync / agent / audit 三入口共用）。

    归类语义：自带 code 沿用；无 code 一律 UNEXPECTED；stage 存在原样保留，缺失填 UNKNOWN。
    CONFIG_NOT_FOUND 只留给真正“缺配置对象”且带明确 code 的路径，不再作为兜底。
    """
    if isinstance(e, Failure):
        return e
    return Failure(
        str(e),
        code=getattr(e, "code", "") or ErrorCode.UNEXPECTED.value,
        stage=getattr(e, "stage", "") or Stage.UNKNOWN.value,
        hint=getattr(e, "hint", "") or "",
        suggested=getattr(e, "suggested", "") or "",
        retryable=bool(getattr(e, "retryable", False)),
    )


@dataclass
class Outcome:
    """一个命令用例的返回值；handler 只返回它，不打印。"""

    action: str
    ok: bool = True
    details: list[dict[str, Any]] = field(default_factory=list)
    errors: list[Failure] = field(default_factory=list)
    stats: dict[str, Any] = field(default_factory=dict)
    human: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    exit_code: int = 0
    elapsed_ms: int = 0

    def add_error(self, error: Failure) -> "Outcome":
        self.errors.append(error)
        self.ok = False
        if self.exit_code == 0:
            self.exit_code = 1
        return self
