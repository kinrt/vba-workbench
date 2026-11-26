# -*- coding: utf-8 -*-
"""唯一 COM 适配器：打开/附着工作簿只允许从这里进。

旧代码里仍有直接调用；迁移完成的判据是 grep 只在本文命中。
"""
from __future__ import annotations

from typing import Any

from toolkit import excel


def attach_running_excel(*args: Any, **kwargs: Any):
    return excel.attach_running_excel(*args, **kwargs)


def find_open_workbook(*args: Any, **kwargs: Any):
    return excel.find_open_workbook(*args, **kwargs)


def open_disk_workbook(*args: Any, **kwargs: Any):
    return excel.open_disk_workbook(*args, **kwargs)


def new_excel(*args: Any, **kwargs: Any):
    return excel.new_excel(*args, **kwargs)


def classify_com_failure(*args: Any, **kwargs: Any):
    return excel.classify_com_failure(*args, **kwargs)


def read_vbproject_probe(wb):
    """只读探测已打开工作簿的 VBA 工程快照（H16/H17 运行期体检用）。

    边界约定：这里只做 COM 读取、回普通 dict，不做任何通过/警告判定；
    引用名、组件类型、模块行数、表内控件数都是事实，结论由 feature/engine 收口。
    工程被锁或信任中心未放行时，整体读不到则 ``error`` 填原始异常文本。

    返回 {"refs": [引用名…], "components": [{"type", "lines", "controls"}],
          "lines": 模块行总数, "error": 异常文本或 None}。
    """
    snap: dict[str, Any] = {"refs": [], "components": [], "lines": 0, "error": None}
    try:
        vbp = wb.VBProject
        snap["refs"] = [str(r.Name) for r in list(vbp.References)]
        total_lines = 0
        for comp in list(vbp.VBComponents):
            ctype = int(comp.Type)
            lines = int(comp.CodeModule.CountOfLines)
            controls = 0
            if ctype == 100:  # 100=文档模块（Sheet/ThisWorkbook），表内控件只在这上面数
                try:
                    controls = int(comp.Designer.Controls.Count)
                except Exception:
                    controls = 0
            total_lines += lines
            snap["components"].append({"type": ctype, "lines": lines, "controls": controls})
        snap["lines"] = total_lines
    except Exception as e:
        snap["error"] = str(e)
    return snap


__all__ = ["attach_running_excel", "find_open_workbook", "open_disk_workbook",
           "new_excel", "classify_com_failure", "read_vbproject_probe"]
