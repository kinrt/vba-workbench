# -*- coding: utf-8 -*-
"""宏执行骨架：sync.run 与 agent.assert 共用的唯一执行/分类入口。

本模块是执行路径的**唯一**分类器（开表/附着仍只走
``adapters.excel_gateway``）：

  · ``MacroRunner.run`` 负责：活动簿核对 → guard 快照/还原 → ``Application.Run``
    → 错误分类 → 日志收尾 → 隔离模式按需保存。
  · 分类器只有这一份：参数个数不符 / COM 忙断开 / 业务失败，
    sync.run 与 agent.assert 拿到同一个 code。

会话约定（``session``）：
  · ``session.workbook``：COM 工作簿对象，需有 ``Application`` / ``Name`` /
    ``FullName`` / ``Path`` / ``Activate()`` / ``Save()``；
  · 可选 ``session.macro_declared(macro) -> bool``：宏是否在源码里声明过，
    只用于提示文案，不改变错误码。

测试一律用假 session / 桩对象，不在本模块里 new COM、不真启动 Excel。
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from . import app_guard, files
from .contract import ErrorCode, Failure, Stage

# 宏默认写日志的子目录（相对工作簿目录）；与 1.5.6 口径一致。
RUN_LOG_DIR = "测试输出"

# 业务失败时给人的排查清单（与 1.5.6 一致；runner 只归集，人话由编排层展示）。
RUN_FAIL_HINTS: tuple[str, ...] = (
    "宏名对不上：确认模块里真有这个过程，且不是 Private（Private 过程 Run 调不到）",
    "模块名与过程名重名时，'书'!名字 会被解析成那个模块 —— 写成 模块.过程",
    "文件必须是以「启用宏」方式打开的 .xlsm；宏被安全设置挡住时 Run 直接报名字找不到",
    "宏里弹了 MsgBox / InputBox：模态框会把这次调用一直卡住——给它加一个「静默」参数",
    "工程里有编译错误（任何一处编译错都会让全部宏不可用）：先回 VBE 修好再跑",
)


@dataclass
class MacroResult:
    """一次宏执行的结论；编排层据此拼 Outcome。"""

    macro: str
    return_value: Any = None
    log_text: str = ""
    duration_ms: int = 0
    timed_out: bool = False
    warnings: list[str] = field(default_factory=list)

    # 执行结论与分类（追加字段，默认值保持桩形状兼容）。
    ok: bool = True
    error_code: str = ""
    error_stage: str = ""
    error_message: str = ""
    hint: str = ""
    log_status: str = ""          # found / missing / not_applicable / ""
    log_file: str = ""
    saved: bool = False


class MacroRunner:
    """实现者只在这里调用 COM；不要在 features 里另写 Application.Run 分类器。"""

    def __init__(self, session, *, gateway: Any | None = None) -> None:
        self.session = session
        self._gateway = gateway

    # ------------------------------------------------------------------ 分类
    def classify(self, err: object, *, declared: bool | None = None,
                 automation_security: int | None = None,
                 macro: str = "") -> tuple[str, str, str]:
        """把 Application.Run 的 COM 异常归成 (code, stage, hint)。

        分类顺序：先认 Excel 忙/断开（COM 上下文），再认参数个数，
        其余（含宏不存在 / 不可用 / 被禁用 / 业务报错）一律收口 RUN_FAILED。
        """
        text = str(err).lower()

        com_code = self._com_code(err)
        if com_code:
            hint = {
                "EXCEL_BUSY": "Excel 忙（弹窗/计算/模态框）；先处理弹窗再重试。",
                "EXCEL_CLOSED": "Excel 进程已退出或 RPC 断开；重新打开目标表后重试。",
                "COM_CONTEXT": "连不上 Excel COM；检查 pywin32/Excel 会话。",
            }.get(com_code, "")
            return com_code, Stage.ATTACH.value, hint

        if ("-2147352562" in text or "-2147352561" in text
                or "0x8002000e" in text or "0x8002000f" in text
                or "参数个数" in text or "参数不足" in text
                or "wrong number of arguments" in text
                or "parameter count" in text or "parameter not optional" in text):
            return ("ARG_MISMATCH", Stage.RUN.value,
                    f"宏 {macro} 收到的参数个数与过程签名不符；核对 --arg 的个数/顺序，"
                    "例：宏声明 (a As Long, b As Long) 时传 --arg 2 --arg 3，"
                    "可选参数可先省略，对象参数只能写包装宏")

        # 宏不存在 / 不可用 / 被禁用：1.5.6 兜底分支即 RUN_FAILED，不另造专用码。
        # 区分信息只落 hint（错误码仍统一 RUN_FAILED）：缺宏写清宏名/模块名与编译问题，
        # 禁用宏写清信任位置 / 宏安全设置。
        unavailable = ("此工作簿中不可用" in text or "macro may not be available" in text)
        disabled = ("宏都被禁用" in text or "macros may be disabled" in text)
        hint_lines: list[str] = []
        if unavailable:
            hint_lines.append(
                f"找不到宏「{macro}」：核对模块名/过程名拼写，过程不是 Private；"
                "工程里有编译错误时全部宏也会不可用，先回 VBE 修好编译问题")
        if disabled:
            hint_lines.append("宏被安全策略禁用：在「文件→选项→信任中心→宏设置」启用宏，"
                               "或把工作簿所在目录加入「信任位置」后重新打开")
        if automation_security == 1:
            hint_lines.append("当前是 automation 实例，AutomationSecurity=Low，宏未被安全策略拦截")
        hint_lines.append("看「运行宏提醒」逐条排查；需要干净环境可改用 run --isolate")
        return ("RUN_FAILED", Stage.RUN.value, "；".join(hint_lines))

    def _com_code(self, err: object) -> str:
        """忙/断开走适配器分类器；读不到适配器时按空串处理，不瞎判。"""
        gateway = self._gateway
        if gateway is None:
            from adapters import excel_gateway
            gateway = excel_gateway
        classify = getattr(gateway, "classify_com_failure", None)
        if classify is None:
            return ""
        try:
            return str(classify(err) or "")
        except Exception:  # noqa: BLE001 - 分类器自身异常不能盖住原宏错误
            return ""

    # ------------------------------------------------------------------ 执行
    def run(self, macro: str, args: list[Any] | None = None, *, isolate: bool = False,
            save: bool = False, log_optional: bool = True,
            log_target: str | Path | None = None, tail: int = 40) -> MacroResult:
        """运行宏并回收结果；COM 失败抛统一 ``Failure``，不返回裸异常。"""
        argv = list(args or [])
        workbook = getattr(self.session, "workbook", None)
        if workbook is None:
            raise Failure(
                "没有可运行的工作簿：目标表未打开，且未走隔离模式开副本。",
                code=ErrorCode.TARGET_NOT_OPEN, stage=Stage.ATTACH,
                hint="先打开目标工作簿，或用 book open；隔离模式下目标表必须未打开。",
            )
        app = workbook.Application

        # 硬规矩：xl.Run 只跑活动工作簿；跑前 Activate 并核对 ActiveWorkbook.Name。
        try:
            workbook.Activate()
        except Exception as e:  # noqa: BLE001
            raise Failure(f"切不到目标工作簿：{e}", code=ErrorCode.TARGET_NOT_ACTIVE,
                          stage=Stage.ATTACH) from e
        active = self._safe_active_workbook(app)
        if active is not None and getattr(active, "Name", None) != getattr(workbook, "Name", None):
            raise Failure(
                f"活动簿是 {getattr(active, 'Name', '?')}，不是目标 {getattr(workbook, 'Name', '?')}；"
                "Run 只跑活动簿，先在 Excel 里切到目标表。",
                code=ErrorCode.TARGET_NOT_ACTIVE, stage=Stage.ATTACH,
            )

        declared = self._declared(macro)
        security = self._automation_security(app)
        started = time.perf_counter()
        guard = app_guard.guard(app)
        try:
            with guard:
                ret = app.Run(macro, *argv)
        except Exception as e:  # noqa: BLE001 - 分类后统一抛 Failure
            code, stage, hint = self.classify(e, declared=declared,
                                              automation_security=security, macro=macro)
            message = f"宏 {macro} 没有跑完：{e}"
            if code == "ARG_MISMATCH":
                message = (f"宏 {macro} 参数个数不符：本次传了 {len(argv)} 个；"
                           "核对 --arg 与过程签名")
            raise Failure(message, code=code, stage=stage, hint=hint) from e

        duration_ms = int((time.perf_counter() - started) * 1000)
        warnings: list[str] = []
        if guard.changes:
            warnings.append(f"已还原 Excel 状态：{guard.describe()}")
        for line in guard.errors[:3]:
            warnings.append(f"状态还原未完成：{line}")

        # 日志收尾：优先显式 --log，否则取工作簿目录下日志目录里本次新落的那份。
        log_file, log_text, log_status = self._collect_log(
            workbook, started, ret=ret, log_target=log_target,
            log_optional=log_optional, tail=tail)

        saved = False
        if save and isolate:
            try:
                workbook.Save()
                saved = True
            except Exception as e:  # noqa: BLE001
                raise Failure(
                    f"运行完成，但保存失败：{e}",
                    code=ErrorCode.SAVE_FAILED, stage=Stage.SAVE,
                    hint="检查文件只读/占用；或去掉 --save 只取返回值/日志。",
                ) from e

        return MacroResult(
            macro=macro, return_value=ret, log_text=log_text,
            duration_ms=duration_ms, warnings=warnings,
            ok=True, log_status=log_status, log_file=str(log_file) if log_file else "",
            saved=saved,
        )

    # ------------------------------------------------------------------ 内部
    def _declared(self, macro: str) -> bool | None:
        hook = getattr(self.session, "macro_declared", None)
        if not callable(hook):
            return None
        try:
            return bool(hook(macro))
        except Exception:  # noqa: BLE001 - 探测失败交分类器兜底
            return None

    @staticmethod
    def _automation_security(app) -> int | None:
        try:
            return int(app.AutomationSecurity)
        except Exception:  # noqa: BLE001 - 属性不可读时交兜底
            return None

    @staticmethod
    def _safe_active_workbook(app):
        try:
            return app.ActiveWorkbook
        except Exception:  # noqa: BLE001
            return None

    def _collect_log(self, workbook, started: float, *, ret: Any,
                     log_target: str | Path | None, log_optional: bool,
                     tail: int) -> tuple[Path | None, str, str]:
        """按 1.5.6 口径找本次日志；返回 (路径, 尾部文本, 状态)。"""
        if tail <= 0:
            return None, "", ""

        target: Path | None = None
        if log_target:
            target = Path(log_target)
        else:
            book_dir = self._book_dir(workbook)
            if book_dir is not None:
                target = book_dir / RUN_LOG_DIR

        log_file: Path | None = None
        if target is not None:
            if target.is_file():
                log_file = target
            elif target.is_dir():
                log_file = files.newest_file(target, "*.log", since=started - 2)

        if log_file is None:
            # Sub 本来就可能不写日志；只有显式给了 --log 或非 Sub 才像故障。
            if log_optional or (ret is None and not log_target):
                return None, "", "not_applicable"
            return None, "", "missing"

        text = files.read_tail(log_file, tail)
        return log_file, text, "found"

    @staticmethod
    def _book_dir(workbook) -> Path | None:
        full_name = getattr(workbook, "FullName", "") or ""
        if full_name:
            return Path(full_name).parent
        path = getattr(workbook, "Path", "") or ""
        return Path(path).parent if path else None
