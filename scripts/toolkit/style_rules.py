# -*- coding: utf-8 -*-
"""toolkit.style_rules —— VBA 源码风格检查与修复（只动 src 文本，不碰表格）

落实 README §11 硬规则 1（UTF-8 无 BOM + LF）与硬规则 11（调用括号风格），
并与 push 预检共用“缺 Option Explicit”判定（单一真相源）。
规则实现原样沿用 `style.py`（2026-11 括号统一战役沉淀，267 处实战验证），
本文件只把 CLI 与菜单剥掉，留下纯函数，由 `vba-workbench\\scripts\\agent.py style` 调用。

  1. eol    .bas/.cls/.frm 必须 UTF-8 无 BOM + LF；修复 = 去 BOM / CRLF 压成 LF
  2. noarg  无参调用不许带空括号：语句行 `Call X()` → `Call X`（声明行豁免）
  3. bare   带参调用必须括号：`MsgBox "a", 1` / `Run "x"` / `obj.M a` /
            行内 `Then Foo a` → `Call Foo(a)`；`Foo (x)` 外层双括号会合并；
            命名参数 `:=`、顶层冒号分隔语句、` Else` 不被吞进参数；
            续行 `_` 在末行补 `)`；行尾注释留在括号外

识别口径（启发式，有边界）：bare 的"裸调用"判定 = 全库声明过的过程名（Sub/Function）
+ 内置 MsgBox/InputBox/Run + `a.b.c` 点号方法形态；语言语句（Load/Unload/Debug.Print/
ChDir/FileCopy/Name/SendKeys/Beep 等）不支持 Call 括号形式，保持原样；赋值/表达式
（首参是 `=`、运算符）不动。
"""

from __future__ import annotations

import re
from pathlib import Path

from .base import entry_command, msg, read_source_text, warn
from .check import Notes
from .files import write_source
from .repo import DOC_MODULE_RE, module_name_of, rel_path, source_files_at
from .ui import section
from .vba_parse import (declared_procs, split_code_comment, split_lines,
                        strip_export_header, strip_vba_comments)

ID = r"[A-Za-z_\u4e00-\u9fff][\w\u4e00-\u9fff]*"
BOM = b"\xef\xbb\xbf"
BUILTIN = {"MsgBox", "InputBox", "Run"}

# 语言关键字/语句：行首出现则不是调用
KW = set("""If Then Else ElseIf End For Each Next While Wend Do Loop With Set Let Dim Static
Public Private Friend Sub Function Exit Select Case GoTo GoSub On Error Resume Return
Call Print Debug Attribute Begin Option Explicit Base Rem Const Type Enum Declare Lib
Load Unload ChDir ChDrive FileCopy Kill Name MkDir RmDir SavePicture SendKeys Beep Random
Open Close Get Put Line Width Circle Cls New Step To ByRef ByVal ParamArray Optional
RaiseEvent Implements Event Property Pause AppActivate Date Time Timer
""".split())
DOT_EXCLUDE = {"Debug.Print"}
OP_HEAD = re.compile(r"^(=|<>|<=|>=|<|>|<=|>=|\+|-|\*|/|\^|&|Is|Like|In|Mod|And|Or|Not|Xor|Eqv|Imp)\b|^=")
DECL_RE = re.compile(r"(?i)^(Public |Private |Friend |Static )?(Sub|Function)\s")
PAT_B = re.compile(rf"^(\s*)((?:Call\s+)?{ID}(?:\.{ID})+)\s+(\S.*)$")   # 点号方法 空格 参数
PAT_NOARG = re.compile(rf"^(\s*(?:.*\bThen\s+)?)Call\s+({ID}(?:\.{ID})*)\(\)(\s+(?:'.*)?)?$")


def scope_files(scope: str) -> list[Path]:
    """按范围收集 .bas/.cls/.frm（单个文件或目录）；scope 相对仓库根解析。"""
    return source_files_at(scope)


# ---------------------------------------------------------------------------
# 检查项 1：行尾 / BOM
# ---------------------------------------------------------------------------
def check_option_explicit(files: list[Path]) -> list[tuple[str, str]]:
    """缺 Option Explicit 的源码（.bas/.cls，文档模块与空壳不算）。

    与 sync 的 source_preflight 共用这一个判定，避免 style 说全过、push 预检
    又说 16 个文件缺 Option Explicit。
    """
    issues: list[tuple[str, str]] = []
    for p in files:
        if p.suffix.lower() not in (".bas", ".cls"):
            continue
        if p.suffix.lower() == ".bas" and DOC_MODULE_RE.match(module_name_of(p)):
            continue
        text = read_source_text(p)
        code = strip_vba_comments(strip_export_header(text))
        if not _has_vba_code(code):
            continue
        if not re.search(r"(?im)^\s*Option\s+Explicit\b", code):
            issues.append((rel_path(p), "缺 Option Explicit"))
    return issues


def has_option_explicit(path: Path) -> bool:
    """单个文件是否有 Option Explicit（写回预检逐文件调用）。"""
    text = read_source_text(path)
    return bool(re.search(r"(?im)^\s*Option\s+Explicit\b",
                          strip_vba_comments(strip_export_header(text))))


def _has_vba_code(code: str) -> bool:
    """排除只有导出头/注释/Option 的空壳，避免对占位壳误报。"""
    for line in code.splitlines():
        s = line.strip()
        if not s or s.startswith("'") or re.match(r"(?i)^Rem\b", s):
            continue
        if re.match(r"(?i)^Option\s+(Explicit|Base|Compare|Private)\b", s):
            continue
        if re.match(r"(?i)^Attribute\s+VB_", s):
            continue
        return True
    return False


def check_eol(files: list[Path]) -> list[tuple[str, str]]:
    issues = []
    for p in files:
        raw = p.read_bytes()
        rp = rel_path(p)
        if raw.startswith(BOM):
            issues.append((rp, "UTF-8 BOM（应无 BOM）"))
        crlf = raw.count(b"\r\n")
        if crlf:
            issues.append((rp, f"CRLF 行尾（应全 LF，共 {crlf} 处）"))
        cr = raw.count(b"\r") - crlf
        if cr and not crlf:
            issues.append((rp, f"裸 CR {cr} 处（应全 LF）"))
    return issues


def fix_eol(files: list[Path], apply: bool) -> int:
    n = 0
    for p in files:
        raw = p.read_bytes()
        if not raw.startswith(BOM) and b"\r" not in raw:
            continue
        text = raw.decode("utf-8-sig", errors="replace")
        n += 1
        if apply:
            write_source(p, text)   # write_source 去 BOM、CRLF/裸 CR 一律压成 LF
    return n


# ---------------------------------------------------------------------------
# 检查项 2：无参调用不许带空括号（Call X() -> Call X）
# ---------------------------------------------------------------------------
def scan_noarg(files: list[Path], apply: bool) -> list[tuple[str, int, str, str]]:
    changes = []
    for p in files:
        lines = split_lines(read_source_text(p))
        touched = False
        for i, line in enumerate(lines):
            s = line.strip()
            if not s or s.startswith("'") or s.lower().startswith("attribute vb_"):
                continue
            if DECL_RE.match(s):
                continue
            m = PAT_NOARG.match(line)
            if m:
                new = f"{m.group(1)}Call {m.group(2)}{m.group(3) or ''}"
                changes.append((rel_path(p), i + 1, s, new.strip()))
                if apply:
                    lines[i] = new
                    touched = True
        if apply and touched:
            write_source(p, "\n".join(lines))
    return changes


# ---------------------------------------------------------------------------
# 检查项 3：带参调用必须括号（裸参数 -> Call Foo(args)）
# ---------------------------------------------------------------------------
def top_level_split(args: str) -> tuple[str, str]:
    """按顶层冒号（非 :=）/ Else 拆语句：返回 (第一段, 其余原文或 '')。"""
    depth, inq = 0, False
    for idx, c in enumerate(args):
        if c == '"':
            inq = not inq
        elif not inq:
            if c == "(":
                depth += 1
            elif c == ")":
                depth -= 1
            elif c == ":" and depth == 0 and args[idx + 1:idx + 2] != "=":
                return args[:idx].rstrip(), args[idx:]
            elif depth == 0 and args[idx:idx + 5].lower() == " else" and (idx == 0 or args[idx - 1] == " "):
                if args[idx + 5:idx + 6] in ("", " "):
                    return args[:idx].rstrip(), args[idx:]
    return args, ""


def strip_outer_parens(a: str) -> str:
    """整体被一对括号包住则剥掉：(x, y) -> x, y。"""
    a = a.strip()
    if not (a.startswith("(") and a.endswith(")")):
        return a
    depth, inq = 0, False
    for idx, c in enumerate(a):
        if c == '"':
            inq = not inq
        elif not inq:
            if c == "(":
                depth += 1
            elif c == ")":
                depth -= 1
                if depth == 0:
                    return a[1:-1] if idx == len(a) - 1 else a
    return a


def is_cont(code: str) -> bool:
    """真正的续行 = 空格+下划线结尾（变量名 str_ 不算）。"""
    return re.search(r"(?:\s|^)_\s*$", code) is not None


def convert_line(code: str, pat_a, pat_c):
    """返回 (prefix, target, args) 或 None；code 已剥注释。"""
    m = pat_c.match(code)
    if m:
        prefix, target, args = m.groups()
        if OP_HEAD.match(args):
            return None
        return prefix, target, args
    m = pat_a.match(code)
    if m and not m.group(4):
        indent, callkw, name, args = m.group(1), m.group(2), m.group(3), m.group(5)
        if OP_HEAD.match(args):
            return None
        return indent, callkw + name, args
    m = PAT_B.match(code)
    if m:
        indent, target, args = m.group(1), m.group(2), m.group(3)
        t = target.replace("Call ", "").strip()
        if t in DOT_EXCLUDE or OP_HEAD.match(args) or t.split(".")[0] in KW:
            return None
        return indent, target, args
    return None


def collect_names(files: list[Path]) -> set[str]:
    names = set(BUILTIN)
    for p in files:
        names |= declared_procs(p)
    return names


def scan_bare(files: list[Path], apply: bool) -> list[tuple[str, int, str, str]]:
    names = collect_names(files)
    name_re = "|".join(re.escape(n) for n in sorted(names, key=len, reverse=True))
    pat_a = re.compile(rf"^(\s*)((?:Call\s+)?)({name_re})((?:\([^)]*\))?)\s+(\S.*)$")
    pat_c = re.compile(rf"^(.*\bThen\s+)(?:Call\s+)?((?:{ID}\.)*(?:{name_re}))\s+(\S.*)$")

    changes = []
    for p in files:
        lines = split_lines(read_source_text(p))
        touched = False
        i = 0
        while i < len(lines):
            raw = lines[i]
            s = raw.strip()
            if (not s or s.startswith("'") or s.lower().startswith("attribute vb_")
                    or s.startswith("Begin ") or DECL_RE.match(s)):
                i += 1
                continue
            code, _ = split_code_comment(raw)
            if not code.strip():
                i += 1
                continue
            # 续行组：首行是 " _" 结尾则吞并后续物理行
            j, c2 = i, code
            while is_cont(c2) and j + 1 < len(lines):
                j += 1
                c2, _ = split_code_comment(lines[j])
            group_end = j

            r = convert_line(code, pat_a, pat_c)
            if r:
                prefix, target, args = r
                target = re.sub(r"^Call\s+", "", target.strip())
                args = strip_outer_parens(args)
                first_args, rest = top_level_split(args)
                if group_end == i:
                    _, comment = split_code_comment(lines[i])
                    lines[i] = (f"{prefix}Call {target}({first_args})"
                                f"{rest}{('   ' + comment) if comment else ''}")
                else:
                    lines[i] = f"{prefix}Call {target}({args}"
                    last_code, last_comment = split_code_comment(lines[group_end])
                    lines[group_end] = last_code + ")" + (("   " + last_comment) if last_comment else "")
                changes.append((rel_path(p), i + 1, s, lines[i].strip()))
                touched = True
                i = group_end + 1
                continue
            i += 1
        if apply and touched:
            write_source(p, "\n".join(lines))
    return changes


# ---------------------------------------------------------------------------
# 对外：检查 / 修复 / 落盘后自检
# ---------------------------------------------------------------------------
def _report(title: str, items, fmt) -> int:
    section(title)
    if not items:
        msg("  ✔ 无问题")
        return 0
    for it in items:
        msg(fmt(it))
    msg(f"  ✘ 共 {len(items)} 处")
    return len(items)


def run_checks(files: list[Path], which: str = "all") -> int:
    """只读检查：逐项打清单，末尾用统一提醒汇总（与 sync / audit 同一套分级）。"""
    return run_checks_with_details(files, which)[0]


def run_checks_with_details(files: list[Path], which: str = "all") -> tuple[int, list[dict]]:
    """返回 (退出码, 结构化明细)，供 agent.style 的 details[] 复用同一次扫描。"""
    notes = Notes("风格检查提醒")
    details: list[dict] = []
    bad = 0
    if which in ("all", "eol"):
        items = check_eol(files)
        details.extend({"rule": "eol", "file": file, "message": text}
                       for file, text in items)
        bad += _report("行尾 / BOM 检查", items, lambda t: f"  {t[0]}: {t[1]}")
        notes.add("警告" if items else "校验",
                  f"行尾/BOM 不合规 {len(items)} 个文件" if items else "行尾/BOM 全部合规")
    if which in ("all", "explicit"):
        items = check_option_explicit(files)
        details.extend({"rule": "explicit", "file": file, "message": text}
                       for file, text in items)
        bad += _report("Option Explicit 检查", items, lambda t: f"  {t[0]}: {t[1]}")
        notes.add("警告" if items else "校验",
                  f"缺 Option Explicit {len(items)} 个文件" if items else "Option Explicit 全部合规")
    if which in ("all", "call"):
        noarg_items = scan_noarg(files, False)
        bare_items = scan_bare(files, False)
        details.extend({"rule": "call_noarg", "file": file, "line": line,
                        "old": old, "new": new}
                       for file, line, old, new in noarg_items)
        details.extend({"rule": "call_bare", "file": file, "line": line,
                        "old": old, "new": new}
                       for file, line, old, new in bare_items)
        noarg = _report("无参调用空括号检查（Call X()）", noarg_items,
                        lambda t: f"  {t[0]} L{t[1]}\n    - {t[2]}\n    + {t[3]}")
        bare = _report("裸参数调用检查（应 Call Foo(a,b)）", bare_items,
                       lambda t: f"  {t[0]} L{t[1]}\n    - {t[2]}\n    + {t[3]}")
        bad += noarg + bare
        notes.add("警告" if noarg + bare else "校验",
                  f"调用括号不合规 {noarg + bare} 处（无参 {noarg} / 裸参数 {bare}）"
                  if noarg + bare else "调用括号全部合规")
    if bad:
        warn(f"合计 {bad} 处问题；`{entry_command('agent.py', 'style', '--action', 'fix')}` 预览，加 --apply 落盘")
    else:
        msg("\n全部通过：src 风格符合 README §11 硬规则 1 / 11，且所有业务模块都有 Option Explicit")
    msg("")
    for line in notes.render("风格检查提醒"):
        msg(line)
    return (1 if bad else 0), details


def selfcheck(files: list[Path], which: str = "all") -> int:
    """落盘后自检：只报数，不重复打逐条清单（有残留就点出来让人回头看预览）。"""
    bad = 0
    if which in ("all", "eol"):
        bad += len(check_eol(files))
    if which in ("all", "explicit"):
        bad += len(check_option_explicit(files))
    if which in ("all", "call"):
        bad += len(scan_noarg(files, False)) + len(scan_bare(files, False))
    section("落盘后自检（同一批文件）")
    if bad:
        warn(f"  ✘ 还有 {bad} 处没修掉（多半是启发式覆盖不到的形态，看预览手工处理）")
    else:
        msg("  ✔ 全部通过，可以提交")
    return bad


def run_fix(files: list[Path], which: str = "all", apply: bool = False) -> int:
    mode = "已落盘" if apply else "预览（--apply 才写文件）"
    if which in ("all", "eol"):
        issues = check_eol(files)
        n = fix_eol(files, apply)
        section(f"行尾 / BOM 修复 [{mode}]")
        msg(f"  {'修复' if apply else '待修复'} {n} 个文件" if issues else "  ✔ 无需修复")
    if which in ("all", "call"):
        for title, fn in (("无参空括号 → Call X", scan_noarg),
                          ("裸参数 → Call Foo(a,b)", scan_bare)):
            items = fn(files, apply)
            section(f"{title} [{mode}]")
            if not items:
                msg("  ✔ 无需修复")
                continue
            for rp, ln, old, new in items:
                msg(f"  {rp} L{ln}\n    - {old}\n    + {new}")
            msg(f"  共 {len(items)} 处" + ("，已写回" if apply else "，未写回"))
    if apply:
        # 落盘后立刻对同一批文件复检：修完不复检等于不知道有没有残留
        return 1 if selfcheck(files, which) else 0
    msg("")
    msg("预览模式：确认清单无误后加 --apply 落盘")
    return 0
