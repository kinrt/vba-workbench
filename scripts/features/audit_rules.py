# -*- coding: utf-8 -*-
"""audit_rules —— audit 规则单一来源：阈值 / 白名单 / 分组登记 / 指标口径正则与纯辅助函数。

只放无副作用的数据与纯函数；不枚举源码、不碰 COM、不写文件。
feature 层与引擎切片不得复制阈值表，需要规则口径一律从本模块取。
"""
from __future__ import annotations

import re
from collections.abc import Iterable
from pathlib import Path
from typing import Any

from toolkit.repo import rel_path

# ---------------------------------------------------------------------------
# 规则单一来源：分组登记 / 白名单 / 阈值（旧版散落在 audit.py 多处，这里只留一份）
# ---------------------------------------------------------------------------

# 仅提醒、不计入结构问题总数的分组
NOTE_ONLY_GROUPS: frozenset[str] = frozenset({"shadow_known"})

# 会阻断写回/编译/运行的分组：overview 只对这几组退 1
FATAL_STRUCTURE_GROUPS: frozenset[str] = frozenset(
    {"sheet_mismatch", "binding_target", "dup", "shadow"})

# 已登记的项目侧同名过程（本模块自用、无跨模块裸调）：只登记不报警
SHADOW_KNOWN: frozenset[str] = frozenset(
    {"分析", "输出", "分词", "test1", "获取选择的数据"})

# H06 状态栏临时写入白名单（文件名）
STATUSBAR_ALLOW_FILES: frozenset[str] = frozenset(
    {"表格工具.cls", "表格工具.bas", "自定义通用函数.bas", "测试表格工具.bas",
     "数据同步.bas", "CommonVersion.bas"})

# H01 VBA 内建函数 / 语句裸名白名单（小写匹配，口径同 1.5.6）
BUILTIN_CALLS: frozenset[str] = frozenset(
    """msgbox debug.print array split join len trim ltrim rtrim mid left right instr instrrev replace
ucase lcase cstr clng cint cdbl csng cbool cdate cbyte cur cvar format now date time timer
isnumeric isdate isempty isnull isarray isobject ismissing iserror iserr dir filelen filedatetime
shell environ inputbox chr asc val abs int fix round rnd sgn sqr exp log sin cos tan atan
createobject getobject callbyname choose iif switch partition strreverse string space strcomp
filter ubound lbound erase redim freefile open close print write input line get put seek
kill mkdir rmdir chdir chdrive curdir filecopy name setattr getattr randomize doevents load unload
sendkeys beep appactivate savepicture loadpicture qbcolor rgb""".split())

# 备份体量阈值（H10）：批次条数 / 单份体积（MB）
BACKUP_BATCH_LIMIT = 60
BACKUP_SIZE_MB_LIMIT = 450

# 引用扫描用的正则（口径同旧版，不复制阈值表）
_ID = r"[\w一-龥]+"
_WITH_RE = re.compile(r"(?i)^\s*With\s+(\S+)")
_END_WITH_RE = re.compile(r"(?i)^\s*End With\b")
_DYN_RE = re.compile(r"(?i)\b(Application\.Run|CallByName|Evaluate|Execute)\b")
_DECL_SKIP_RE = re.compile(
    rf"(?i)^(Public |Private |Friend )?(Sub|Function|Property|Dim|Const|Static)\s+")
_SIG_PARAM_RE = re.compile(
    rf"(?i)^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function)\s+{_ID}\s*\((.*)\)")
_DIM_LOCAL_RE = re.compile(rf"(?i)^\s*(?:Dim|Static|Const|ReDim)\s+(.+)$")

# 指标扫描正则（H01/H03/H05b/H06/Declare）
_CALL_TARGET_RE = re.compile(r"(?i)^\s*Call\s+(?!Me\.)([^\s(,]+)")
_ALERTS_RE = re.compile(r"(?i)Application\.DisplayAlerts\s*=\s*([^\s']+)")
_STATUS_RE = re.compile(r"(?i)Application\.StatusBar\s*=")
_SETDATA_CALL_RE = re.compile(r"(?i)\bsetData\s*\(")
_DECLARE_RE = re.compile(r"(?i)^\s*(?:Public\s+|Private\s+)?Declare\s+(?:PtrSafe\s+)?"
                         r"(?:Sub|Function)\s+([^\s(]+)")
_PROC_HEAD_RE = re.compile(
    r"(?i)^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([^\s(]+)")


def _names_in_list(text: str) -> set[str]:
    """逗号分隔的声明列表 → 名字集合（剥掉 ByVal/ByRef/Optional/ParamArray 与类型、初值）。"""
    out: set[str] = set()
    for part in text.split(","):
        part = re.sub(r"(?i)^\s*(ByVal|ByRef|Optional|ParamArray)\s+", "", part).strip()
        m = re.match(rf"({_ID})", part)
        if m:
            out.add(m.group(1))
    return out


def _ignore_set(ignore: Iterable[str] | None) -> set[str]:
    """--ignore 分组/指标名：可重复传，也接受逗号分隔；统一小写。"""
    out: set[str] = set()
    for item in ignore or ():
        for piece in str(item).split(","):
            piece = piece.strip().lower()
            if piece:
                out.add(piece)
    return out


def _jsonable(value: Any) -> Any:
    """结构问题元素 → JSON 安全形式：Path 转字符串，元组转正列。"""
    if isinstance(value, Path):
        return str(rel_path(value))
    if isinstance(value, tuple):
        return [_jsonable(v) for v in value]
    if isinstance(value, list):
        return [_jsonable(v) for v in value]
    if isinstance(value, dict):
        return {k: _jsonable(v) for k, v in value.items()}
    return value
