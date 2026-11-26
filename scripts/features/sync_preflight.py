# -*- coding: utf-8 -*-
"""sync_push 切片内聚段：写回前源码预检（只读，不碰 Excel）。

从 ``features.sync_push`` 拆出的预检流程：目录收集、类头/续行/私有字段/编码等
静态检查；返回 ``(fatal, notes)``：fatal 非空时调用方必须放弃写回，notes 是软提醒，
只进 Outcome.warnings，不改变退出码。口径同 1.5.6 ``source_preflight``。

有意保留：与 run/agent 预检的差异轴为 notes 软提醒、检查深度（类头 VERSION、
Me. 私有字段、24 行续行）、子集过滤；不抽公共骨架，避免告警丢失或多出不适用检查。
"""
from __future__ import annotations

import re
from pathlib import Path

from toolkit.base import abs_path, read_source_text
from toolkit.contract import ErrorCode, Failure
from toolkit.repo import is_shell_file, module_name_of, rel_path, source_files
from toolkit.style_rules import has_option_explicit
from toolkit.vba_parse import norm_code_lines


def _source_dirs(proj: dict, defaults: dict) -> list[Path]:
    """公共库目录在前、项目源码目录在后（写回顺序同旧版）。"""
    dirs: list[Path] = []
    if defaults.get("common_dir"):
        dirs.append(abs_path(defaults["common_dir"]))
    if proj.get("src"):
        dirs.append(abs_path(proj["src"]))
    return dirs


_VB_NAME_RE = re.compile(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"',
                         re.IGNORECASE | re.MULTILINE)
_DOC_NAME_RE = re.compile(
    r"^(ThisWorkbook|Sheet\d+|Chart\d+|DialogSheet\d+|MacroSheet\d+)$", re.IGNORECASE)
_CLASS_HEADER_RE = re.compile(r"^\s*VERSION\s+1\.0\s+CLASS\s*$", re.IGNORECASE)
_CONTINUATION_LIMIT = 24
_PRIVATE_FIELD_RE = re.compile(
    r"^\s*Private\s+(?:WithEvents\s+)?"
    r"([A-Za-z_\u4e00-\u9fff][\w\u4e00-\u9fff]*)\s+As\s+", re.IGNORECASE)


def _me_private_field(text: str) -> tuple[str, int] | None:
    """类模块里 `Me.<私有字段>` 的软提醒；返回 (字段名, 引用行号)。"""
    fields: dict[str, int] = {}
    for i, line in enumerate(text.splitlines(), 1):
        m = _PRIVATE_FIELD_RE.match(line)
        if m:
            fields.setdefault(m.group(1).lower(), i)
    if not fields:
        return None
    for i, line in enumerate(text.splitlines(), 1):
        for m in re.finditer(r"(?i)\bMe\.([A-Za-z_\u4e00-\u9fff][\w\u4e00-\u9fff]*)", line):
            if m.group(1).lower() in fields:
                return m.group(1), i
    return None


def _source_has_code(text: str) -> bool:
    """正文里是否有真正的代码（排除导出头、空行、注释和 Option 语句）。"""
    for line in norm_code_lines(text):
        s = line.strip()
        if not s or s.startswith("'") or re.match(r"(?i)^Rem\b", s):
            continue
        if re.match(r"(?i)^Option\s+(Explicit|Base|Compare|Private)\b", s):
            continue
        if re.match(r"(?i)^Attribute\s+VB_", s):
            continue
        return True
    return False


def _missing_class_header(text: str) -> int | None:
    """`.cls` 第一条实质行必须是 VERSION 1.0 CLASS；返回行号，没有则 None。"""
    for i, line in enumerate(text.splitlines(), 1):
        if line.strip():
            return None if _CLASS_HEADER_RE.match(line.lstrip("\ufeff")) else i
    return 1


def _long_continuation(text: str) -> tuple[int, int] | None:
    """找单条语句超过 24 行续行的位置；返回 (起始行号, 物理行数)。"""
    start = 0
    length = 0
    for i, raw in enumerate(text.splitlines(), 1):
        s = raw.rstrip()
        stripped = s.strip()
        if not stripped or stripped.startswith("'"):
            continue
        if length == 0:
            start = i
        length += 1
        if length > _CONTINUATION_LIMIT:
            return start, length
        if re.search(r"\s_\s*$", s):
            continue
        length = 0
    return None


def _source_preflight(proj: dict, defaults: dict, doc_modules: bool) -> tuple[list[str], list[str]]:
    """写回前源码预检（只读，不碰 Excel）。

    返回 ``(fatal, notes)``：fatal 非空时调用方必须放弃写回；notes 是软提醒，
    只进 Outcome.warnings，不改变退出码。口径同 1.5.6 ``source_preflight``。
    """
    fatal: list[str] = []
    notes: list[str] = []
    seen_global: dict[str, Path] = {}
    bom_files: list[str] = []
    crlf_files: list[str] = []
    no_option_files: list[str] = []
    trunc_files: list[str] = []
    me_private_files: list[str] = []
    files = 0

    dirs = _source_dirs(proj, defaults)
    for folder in dirs:
        if not folder.is_dir():
            fatal.append(f"源码目录不存在：{rel_path(folder)}")

    for folder in dirs:
        if not folder.is_dir():
            continue
        local: dict[str, Path] = {}
        for f in source_files(folder):
            if is_shell_file(f):
                continue
            files += 1
            name = module_name_of(f)
            is_doc = f.suffix.lower() == ".bas" and bool(_DOC_NAME_RE.match(name))
            if name in local:
                fatal.append(f"同一目录内模块重名：{name}（{rel_path(local[name])} / {rel_path(f)}）")
            local[name] = f
            prev = seen_global.get(name)
            if prev is not None and prev.parent != f.parent:
                notes.append(f"跨目录同名模块：{name}（{rel_path(prev)} / {rel_path(f)}），"
                             f"后写的项目源码会覆盖公共库")
            else:
                seen_global[name] = f

            if f.suffix.lower() == ".frm":
                frx = f.with_suffix(".frx")
                if not frx.is_file():
                    fatal.append(f"窗体缺同名 .frx：{rel_path(f)}（应存在 {rel_path(frx)}）")

            try:
                raw = f.read_bytes()
                text = read_source_text(f)
            except Exception as e:  # 编码/文件问题都必须在写回前暴露
                fatal.append(f"源码读不出：{rel_path(f)}（{e}）")
                continue

            if raw.startswith(b"\xef\xbb\xbf"):
                bom_files.append(rel_path(f))
            if b"\r\n" in raw:
                crlf_files.append(rel_path(f))

            if f.suffix.lower() == ".cls":
                line_no = _missing_class_header(text)
                if line_no is not None:
                    fatal.append(f"类模块缺 VERSION 1.0 CLASS 类头：{rel_path(f)}（第 {line_no} 行）")
                me_field = _me_private_field(text)
                if me_field:
                    me_private_files.append(f"{rel_path(f)}:{me_field[1]}（字段 {me_field[0]}）")

            long_cont = _long_continuation(text)
            if long_cont:
                fatal.append(f"单语句续行超过 24 行：{rel_path(f)}"
                             f"（第 {long_cont[0]} 行起，共 {long_cont[1]} 行）")

            m = _VB_NAME_RE.search(text)
            if m and m.group(1).strip() != name:
                fatal.append(f"文件头 VB_Name 与文件名不一致：{rel_path(f)}"
                             f"（VB_Name={m.group(1).strip()}，文件名={name}）")

            if not is_doc and f.suffix.lower() in (".bas", ".cls") and not _source_has_code(text):
                fatal.append(f"模块正文为空：{rel_path(f)}（非文档模块/占位壳）")

            if f.suffix.lower() in (".bas", ".cls") and not is_doc and _source_has_code(text):
                if not has_option_explicit(f):
                    no_option_files.append(rel_path(f))

            body = [ln.rstrip() for ln in norm_code_lines(text) if ln.strip()]
            if body and re.search(r"\s_\s*$", body[-1]):
                trunc_files.append(rel_path(f))

    if files == 0:
        fatal.append("源码目录里没有任何可写模块（只有占位壳也算没有）")

    for label, items in (("BOM", bom_files), ("CRLF", crlf_files),
                         ("缺 Option Explicit", no_option_files),
                         ("末行以续行符 _ 结尾/疑似截断", trunc_files),
                         ("类模块用 Me. 引用私有字段", me_private_files)):
        if items:
            notes.append(f"{label}：{'、'.join(items)}")
    return fatal, notes


def _preflight_or_fail(proj: dict, defaults: dict, doc_modules: bool) -> list[str]:
    """硬预检不通过就抛 Failure；返回软提醒 notes（fatal 时 notes 一并带在返回外）。"""
    fatal, notes = _source_preflight(proj, defaults, doc_modules)
    if fatal:
        raise Failure("源码预检未通过，已放弃写回：" + "；".join(fatal[:3]),
                      code=ErrorCode.SOURCE_PREFLIGHT, stage="preflight",
                      suggested="按提示修源码/文件名后重跑；不要绕过预检直接写回")
    return notes
