# -*- coding: utf-8 -*-
"""agent_io 切片的项目解析与状态一览装配段：probe 动词的支撑原语。

职责边界：
  · 项目名查找、xlsm 解析、测试表名取值、适配器 Failure 契约归一，全部在本模块；
  · 本模块只走 ``adapters.excel_gateway`` 与 toolkit 既有原语，不回壳别的 CLI，
    不 import 其它 features；
  · 写回链（backup/book/push）与日志汇总在 ``agent_io`` / ``agent_logs``。
"""
from __future__ import annotations

import re
import time
from pathlib import Path
from typing import Any

from adapters import excel_gateway
from toolkit.base import entry_command, Failure as BaseFailure
from toolkit.contract import ErrorCode, Failure as ContractFailure, Outcome, Stage
from toolkit.excel import vbe_state
from toolkit.repo import resolve_target

DEFAULT_TEST_SHEET = "测试表"


def _wrap(e: BaseFailure) -> ContractFailure:
    """把适配器抛出的旧基类 Failure 包成契约层 Failure；缺码缺阶段归到 UNEXPECTED。"""
    return ContractFailure(
        str(e),
        code=getattr(e, "code", "") or "UNEXPECTED",
        stage=getattr(e, "stage", "") or Stage.UNKNOWN.value,
        hint=getattr(e, "hint", ""),
        suggested=getattr(e, "suggested", ""),
        retryable=getattr(e, "retryable", False),
    )


def _find(ctx, name: str | None) -> dict:
    """按项目名取配置；找不到报 PROJECT_NOT_FOUND / PROJECT_NOT_SPECIFIED。"""
    items = ctx.config["projects"]
    if not name:
        raise ContractFailure(
            "请指定项目名",
            code=ErrorCode.PROJECT_NOT_SPECIFIED, stage=Stage.CONFIG,
            hint="加 --project 项目名；不知道名字先跑 agent probe --json",
            suggested=entry_command("agent.py", "probe", "--json"),
            retryable=False,
        )
    for proj in items:
        if proj["name"] == name or proj["name"].startswith(name):
            return proj
    raise ContractFailure(
        f"没有这个项目：{name}（可选：{'、'.join(p['name'] for p in items)}）",
        code=ErrorCode.PROJECT_NOT_FOUND, stage=Stage.CONFIG,
        hint="先跑 agent probe --json 看项目名，按输出里的名字重跑",
        suggested=entry_command("agent.py", "probe", "--json"),
        retryable=False,
    )


def _test_sheet(proj: dict, defaults: dict) -> str:
    return str(proj.get("assert_sheet") or defaults.get("assert_sheet")
               or DEFAULT_TEST_SHEET)


def probe(ctx) -> Outcome:
    args = ctx.args
    defaults = ctx.config["defaults"]
    errors: list[ContractFailure] = []

    xl = None
    excel_state = "unknown"
    excel_error: BaseFailure | None = None
    try:
        xl = excel_gateway.attach_running_excel()
        excel_state = "running"
    except BaseFailure as e:
        excel_error = e
        excel_state = "com_context" if e.code == "COM_CONTEXT" else "no_excel"

    with ctx.progress.stage("探测项目状态"):
        targets = (ctx.config["projects"] if not getattr(args, "project", None)
                    else [_find(ctx, args.project)])
        rows: list[dict] = []
        for proj in targets:
            path, _reason = resolve_target(proj)
            sheet = _test_sheet(proj, defaults)
            row: dict[str, Any] = {
                "project": proj["name"],
                "file": str(path) if path else "未解析",
                "disk": "-", "open": "-", "vbe": "-", "modules": "-",
                "test_sheet": "-", "rows": "-",
            }
            if path and path.is_file():
                stat = path.stat()
                row["disk"] = (f"{stat.st_size // 1024}KB/"
                               f"{time.strftime('%m-%d %H:%M', time.localtime(stat.st_mtime))}")
                wb = None
                if xl is not None:
                    wb, _opened = excel_gateway.find_open_workbook(xl, path)
                    row["open"] = "是" if wb is not None else "否"
                if wb is not None:
                    state, detail = vbe_state(wb)
                    row["vbe"] = {"ok": "可读写", "protected": "工程锁着",
                                  "denied": "信任中心未勾"}.get(state, state)
                    row["modules"] = re.sub(r"\D", "", detail) or "-"
                    try:
                        names = [str(s.Name) for s in wb.Sheets]
                        row["test_sheet"] = "有" if sheet in names else "无"
                        if sheet in names:
                            ws = wb.Sheets(sheet)
                            lo = ws.ListObjects(1) if ws.ListObjects.Count else None
                            row["rows"] = (int(lo.ListRows.Count) if lo is not None
                                           else int(ws.UsedRange.Rows.Count))
                    except Exception as e:  # noqa: BLE001 - 单表读不到不拖垮整份清单
                        row["test_sheet"] = f"读不到({e})"
            else:
                row["disk"] = "缺失"
            rows.append(row)

    # COM 不接话一定是故障；--strict 时“Excel 没开”也算健康门禁失败。
    if excel_error is not None and (getattr(args, "strict", False)
                                    or excel_error.code == "COM_CONTEXT"):
        errors.append(_wrap(excel_error))

    human: list[str] = []
    if xl is not None:
        human.append("Excel：已在运行")
    elif excel_error is not None:
        human.append(f"Excel：没在运行（{str(excel_error).splitlines()[0]}）")
    human.append("项目 | 文件 | 磁盘 | 已打开 | 工程 | 模块数 | 测试表 | 行数")
    for r in rows:
        human.append(" | ".join(str(r[k]) for k in (
            "project", "file", "disk", "open", "vbe", "modules", "test_sheet", "rows")))

    return Outcome(
        action="agent.probe",
        ok=not errors,
        details=rows,
        errors=errors,
        stats={"projects": len(rows), "excel": excel_state,
               "lines": human, "timings": ctx.stage_timings()},
        human=human,
        exit_code=1 if errors else 0,
        elapsed_ms=ctx.elapsed_ms(),
    )
