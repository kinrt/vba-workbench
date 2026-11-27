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


def _safe_attr(obj, name: str, default=None):
    """读 COM 属性：单个属性失败不拖垮整张引用清单。"""
    try:
        return getattr(obj, name)
    except Exception:
        return default


def read_project_references(wb):
    """只读返回 VBA 工程引用清单（名称/GUID/路径/版本/是否损坏）。

    与 ``read_vbproject_probe`` 分开：那个只给 H16/H17 汇总引用名；这里保留
    排障需要的详细字段，供 ``audit references`` 使用。工程读不到时 ``error``
    填原始异常文本，调用方决定是否降级成“不可读”警告。
    """
    snap: dict[str, Any] = {"references": [], "error": None}
    try:
        vbp = wb.VBProject
        for ref in list(vbp.References):
            full_path = str(_safe_attr(ref, "FullPath", "") or "")
            snap["references"].append({
                "name": str(_safe_attr(ref, "Name", "") or ""),
                "description": str(_safe_attr(ref, "Description", "") or ""),
                "guid": str(_safe_attr(ref, "GUID", "") or ""),
                "major": _safe_attr(ref, "Major", None),
                "minor": _safe_attr(ref, "Minor", None),
                "full_path": full_path,
                "built_in": bool(_safe_attr(ref, "BuiltIn", False)),
                "broken": bool(_safe_attr(ref, "IsBroken", False)),
            })
    except Exception as e:
        snap["error"] = str(e)
    return snap


__all__ = ["attach_running_excel", "find_open_workbook", "open_disk_workbook",
           "new_excel", "classify_com_failure", "read_vbproject_probe",
           "read_project_references"]
