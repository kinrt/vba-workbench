# -*- coding: utf-8 -*-
"""toolkit.vba_diff —— 源码与 VBE 模块的只读差异比较（纯文本）。

口径以源码为基准：
  · add    = 源码有、工程无
  · update = 两边都有但正文不同
  · remove = 工程有、源码无
  · same   = 两边都有且忽略大小写/空白后相同
只到模块与过程级，不做字符级 diff；删除是否执行由使用方决定，本模块不写任何东西。
"""
from __future__ import annotations

from .vba_lint import parse_proc_spans
from .vba_parse import fold_lines, norm_code_lines


def parse_procedures(code: str) -> dict[str, str]:
    """{过程名: 折叠后的过程正文}；同名过程只保留最后一个。"""
    lines = norm_code_lines(code)
    out: dict[str, str] = {}
    for p in parse_proc_spans(lines):
        out[p["name"]] = "\n".join(fold_lines(lines[p["start"]:p["end"] + 1]))
    return out


def compare_module(source: str, workbook: str) -> dict:
    """单个模块的比较结果；source/workbook 为空串表示该侧不存在。"""
    src_lines = fold_lines(norm_code_lines(source))
    wb_lines = fold_lines(norm_code_lines(workbook))
    src_procs = parse_procedures(source)
    wb_procs = parse_procedures(workbook)
    return {
        "same": src_lines == wb_lines,
        "source_lines": len(norm_code_lines(source)),
        "workbook_lines": len(norm_code_lines(workbook)),
        "procedures": {
            "added": sorted(set(src_procs) - set(wb_procs)),
            "removed": sorted(set(wb_procs) - set(src_procs)),
            "changed": sorted(n for n in set(src_procs) & set(wb_procs)
                              if src_procs[n] != wb_procs[n]),
            "same": sorted(n for n in set(src_procs) & set(wb_procs)
                           if src_procs[n] == wb_procs[n]),
        },
    }


def compare_modules(source: dict[str, str], workbook: dict[str, str]) -> tuple[list[dict], dict]:
    """比较两组模块，返回 (模块明细, 汇总)。status 取 add/update/remove/same。"""
    rows: list[dict] = []
    for name in sorted(set(source) | set(workbook)):
        if name not in workbook:
            status = "add"
            detail = {"same": False, "source_lines": len(norm_code_lines(source[name])),
                      "workbook_lines": 0,
                      "procedures": {"added": sorted(parse_procedures(source[name])),
                                     "removed": [], "changed": [], "same": []}}
        elif name not in source:
            status = "remove"
            detail = {"same": False, "source_lines": 0,
                      "workbook_lines": len(norm_code_lines(workbook[name])),
                      "procedures": {"added": [], "removed": sorted(parse_procedures(workbook[name])),
                                     "changed": [], "same": []}}
        else:
            detail = compare_module(source[name], workbook[name])
            status = "same" if detail["same"] else "update"
        rows.append({"name": name, "status": status, **detail})
    summary = {k: sum(1 for r in rows if r["status"] == k)
               for k in ("add", "update", "remove", "same")}
    summary["total"] = len(rows)
    return rows, summary
