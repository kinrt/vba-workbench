# -*- coding: utf-8 -*-
"""toolkit.vba_lint —— 四条高置信 VBA 静态规则的纯文本实现。

只做源码文本判断，不碰 Excel、不写文件、不打印；feature 层负责文件枚举与输出。
规则固定为四条，刻意不扩成 Rubberduck 式全量 lint：
  · L01_UNUSED_LOCAL            局部 Dim/Static 声明后未使用
  · L02_DUPLICATE_DECL          同一作用域重复声明
  · L03_MISSING_RETURN          Function / Property Get 没有给返回值赋值
  · L04_UNQUALIFIED_RANGE_CELLS 裸 Range(...) / Cells(...)（With 块内放行）
"""
from __future__ import annotations

import re

from .vba_parse import mask_strings, split_lines, strip_vba_lines

IDENT = r"[A-Za-z_\u4e00-\u9fff][\w\u4e00-\u9fff]*"

PROC_HEADER_RE = re.compile(
    rf"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
    rf"(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+(?P<name>{IDENT})",
    re.IGNORECASE)
PROC_END_RE = re.compile(r"^\s*End\s+(Sub|Function|Property)\s*$", re.IGNORECASE)
MODULE_DECL_RE = re.compile(
    rf"^\s*(?:Private|Public|Dim|Global)\s+"
    rf"(?!Sub\b|Function\b|Property\b|Type\b|Enum\b|Const\b|Declare\b)(?P<rest>.*)$",
    re.IGNORECASE)
LOCAL_DECL_RE = re.compile(r"^\s*(?:Dim|Static)\s+(?P<rest>.*)$", re.IGNORECASE)
WITH_START_RE = re.compile(r"^\s*With\s+", re.IGNORECASE)
WITH_END_RE = re.compile(r"^\s*End\s+With\s*$", re.IGNORECASE)
UNQUALIFIED_RANGE_RE = re.compile(r"(?<![\w.])Range\s*\(", re.IGNORECASE)
UNQUALIFIED_CELLS_RE = re.compile(r"(?<![\w.])Cells\s*\(", re.IGNORECASE)


def _split_commas(text: str) -> list[str]:
    """按顶层逗号拆声明；括号里的逗号（数组边界/参数）不拆。"""
    out: list[str] = []
    start = 0
    depth = 0
    for i, ch in enumerate(text):
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth = max(0, depth - 1)
        elif ch == "," and depth == 0:
            out.append(text[start:i])
            start = i + 1
    out.append(text[start:])
    return [x.strip() for x in out if x.strip()]


def _decl_names(line: str, pattern: re.Pattern[str]) -> list[str]:
    """从一条 Dim/Static/模块声明里取变量名；跳过 Const/Type/Declare 等。"""
    m = pattern.match(mask_strings(line))
    if not m:
        return []
    rest = m.group("rest")
    names: list[str] = []
    for item in _split_commas(rest):
        # Dim x As Long: Set x = ... —— 只取冒号前的声明部分
        item = item.split(":", 1)[0].strip()
        if not item or item.lower().startswith(("const ", "type ", "enum ", "declare ")):
            continue
        nm = re.match(IDENT, item)
        if nm:
            names.append(nm.group(0))
    return names


def parse_proc_spans(lines: list[str]) -> list[dict]:
    """返回 [{name, kind, start, end}]；start/end 为 0 基行号，不含 End 行则 end=末行。"""
    spans: list[dict] = []
    i = 0
    while i < len(lines):
        m = PROC_HEADER_RE.match(mask_strings(lines[i]))
        if not m:
            i += 1
            continue
        kind = re.sub(r"\s+", " ", m.group("kind")).title()
        name = m.group("name")
        end = _find_proc_end(lines, i, kind.split()[0])
        spans.append({"name": name, "kind": kind, "start": i, "end": end})
        i = end + 1
    return spans


def _find_proc_end(lines: list[str], start: int, kind_word: str) -> int:
    end_re = re.compile(rf"^\s*End\s+{re.escape(kind_word)}\s*$", re.IGNORECASE)
    for j in range(start + 1, len(lines)):
        if end_re.match(mask_strings(lines[j])):
            return j
    return len(lines) - 1


def _finding(rule: str, message: str, line: int, snippet: str,
             procedure: str = "") -> dict:
    return {
        "rule": rule,
        "severity": "warning",
        "line": line,
        "procedure": procedure,
        "message": message,
        "snippet": snippet.strip()[:140],
    }


def _uses_identifier(name: str, lines: list[str], skip: int) -> bool:
    """变量在过程体内是否被使用；声明行自身不算，字符串字面量先抹掉。"""
    pat = re.compile(rf"(?<!\w){re.escape(name)}(?!\w)", re.IGNORECASE)
    for i, line in enumerate(lines):
        if i == skip:
            continue
        if pat.search(mask_strings(line)):
            return True
    return False


def _used_after_colon(line: str, name: str) -> bool:
    """声明行用冒号接后续语句时，冒号后出现同名即算使用（如 Dim tmp: tmp = a）。"""
    masked = mask_strings(line)
    if ":" not in masked:
        return False
    tail = masked.split(":", 1)[1]
    return bool(re.search(rf"(?<!\w){re.escape(name)}(?!\w)", tail, re.IGNORECASE))


def _check_declarations(lines: list[str], start: int, end: int,
                         pattern: re.Pattern[str], scope: str,
                         proc_name: str) -> tuple[list[dict], dict[str, int]]:
    """扫一个作用域的声明；返回 (findings, {小写名: 首次行号})。"""
    findings: list[dict] = []
    seen: dict[str, int] = {}
    for i in range(start, end + 1):
        names = _decl_names(lines[i], pattern)
        for name in names:
            key = name.lower()
            if key in seen:
                findings.append(_finding(
                    "L02_DUPLICATE_DECL",
                    f"同一作用域重复声明「{name}」，首次在第 {seen[key]} 行",
                    i + 1, lines[i], proc_name))
            else:
                seen[key] = i + 1
    return findings, seen


def _lint_proc(lines: list[str], proc: dict) -> list[dict]:
    findings: list[dict] = []
    start, end = proc["start"], proc["end"]
    name = proc["name"]

    decl_findings, decl_lines = _check_declarations(
        lines, start + 1, end - 1, LOCAL_DECL_RE, "local", name)
    findings.extend(decl_findings)

    # L01：局部变量声明后未使用；声明行冒号后的同句用法也算（如 Dim tmp: tmp = a）
    for lname, decl_line in decl_lines.items():
        skip = decl_line - start - 2  # 声明行在 body 切片里的 0 基下标
        if _used_after_colon(lines[decl_line - 1], lname):
            continue
        if 0 <= skip < (end - start - 1) and _uses_identifier(lname, lines[start + 1:end], skip):
            continue
        findings.append(_finding(
            "L01_UNUSED_LOCAL",
            f"局部变量「{lname}」声明后未使用",
            decl_line, lines[decl_line - 1], name))

    # L03：Function / Property Get 必须有返回值赋值
    kind = proc["kind"].lower()
    if kind == "function" or kind == "property get":
        pat = re.compile(rf"^\s*(?:Set\s+)?{re.escape(name)}\s*=", re.IGNORECASE)
        assigned = any(pat.match(mask_strings(lines[i])) for i in range(start + 1, end))
        if not assigned:
            findings.append(_finding(
                "L03_MISSING_RETURN",
                f"{proc['kind']}「{name}」没有给返回值赋值（缺少 {name} = ...）",
                start + 1, lines[start], name))

    # L04：裸 Range/Cells；With 块内的相对引用放行
    with_depth = 0
    for i in range(start + 1, end):
        s = lines[i].strip()
        if WITH_END_RE.match(s):
            with_depth = max(0, with_depth - 1)
            continue
        if with_depth == 0:
            masked = mask_strings(lines[i])
            if UNQUALIFIED_RANGE_RE.search(masked):
                findings.append(_finding(
                    "L04_UNQUALIFIED_RANGE_CELLS",
                    "Range(...) 未限定工作表；改用 ws.Range(...) 或在 With ws 块内",
                    i + 1, lines[i], name))
            elif UNQUALIFIED_CELLS_RE.search(masked):
                findings.append(_finding(
                    "L04_UNQUALIFIED_RANGE_CELLS",
                    "Cells(...) 未限定工作表；改用 ws.Cells(...) 或在 With ws 块内",
                    i + 1, lines[i], name))
        if WITH_START_RE.match(s):
            with_depth += 1
    return findings


def lint_text(text: str, *, module: str = "", file: str = "") -> list[dict]:
    """对单个源码文本跑四条规则；返回 findings（file/module 由调用方回填）。"""
    lines = strip_vba_lines(text)
    raw = split_lines(text)
    spans = parse_proc_spans(lines)
    findings: list[dict] = []

    # 模块级重复声明：只扫过程之外的行
    module_lines = [i for i in range(len(lines))
                    if not any(p["start"] <= i <= p["end"] for p in spans)]
    module_findings, _ = _check_declarations(
        lines, 0, len(lines) - 1, MODULE_DECL_RE, "module", "")
    # _check_declarations 会扫到过程内；这里只留模块级行
    findings.extend([f for f in module_findings if f["line"] <= len(lines)
                     and (f["line"] - 1) in module_lines])

    for proc in spans:
        findings.extend(_lint_proc(lines, proc))

    for f in findings:
        idx = max(0, min(int(f["line"]) - 1, len(raw) - 1))
        f["file"] = file
        f["module"] = module
        f["snippet"] = raw[idx].strip()[:140]
    findings.sort(key=lambda x: (x["line"], x["rule"]))
    return findings
