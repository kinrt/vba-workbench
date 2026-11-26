# -*- coding: utf-8 -*-
"""
toolkit.vba_parse —— VBA 源码文本的通用解析与规范化（不含业务判断）
============================================================
**所有"把 VBA 代码当文本处理"的口径都住在这里一份**，sync / audit / style 都调它。

一、剥注释与按行读
  · `strip_vba_lines(text)`：逐行剥注释，**行号不错位**（注释行剥完仍是空行）
  · `strip_vba_comments(text)`：整串形式（不需要行号时用）
  · `split_lines(text)`：只按 \\r\\n / \\r / \\n 拆，**不把 \\x0c 当换行**（`str.splitlines` 会误拆）
  · `code_lines(path)`：读文件 + 剥注释 + 按 mtime/size 缓存（同一进程内反复扫同一批文件不重复 IO）
  · 口径：单引号 ' 之后是注释，但双引号字符串里的 ' 不算——`Application.Run "模块.过程"`
    要靠字符串认引用，不能连字符串一起删；`Rem` 开头的整行置空；剥完统一 rstrip。

二、字符串字面量
  · `split_code_comment(line)`：拆成 (代码, 行尾注释)，注释原样保留
  · `mask_strings(line)`：把 "..." 换成**等长空格**（字符串里的宏名、表名不是标识符调用；
    等长而不是删除，是为了让后面的列号/位置判定仍然对得上）

三、VBE 导出头
  · `strip_export_header(code)`：剥掉 `VERSION 1.0 CLASS` / `BEGIN…END` 块 /
    `Attribute VB_*` 及其续行，只处理文件开头；正文里的 End Sub 不会被误删。
    文档模块（ThisWorkbook / SheetN）写回前必须过这一道，否则 VBE 顶上多出四行。

四、成员与事件
  · `parse_members(lines)`：成员表 `{名字: {kind, access, line, span}}`（字段 / 属性 / 方法）
  · `declared_procs(path)` / `VBA_PROC_RE`：只要过程名的轻量版（风格工具判"哪些名字带括号"）
  · `is_event_proc(name)` / `EVENT_SUFFIX`：`CommandButton1_Click` 这类由宿主触发的过程

五、代码比对用的规范化
  · `norm_code_lines(text)`：剥导出头 + 统一换行 + rstrip + 去首尾空行
  · `fold_lines(lines)`：忽略大小写与缩进/空白（VBE 会自己改关键字大小写，不该算冲突）

六、代码口径扫描（回归体检用，audit / 未来的检查脚本都调这里）
  · `scan_rules(files)`：一次扫出裸 New / 手写加速 / 缺过程注释三类命中，
    返回 `{naked_new, manual_accel, no_comment}`，元素是 (文件, 行号, 片段)。
    白名单与例外记号取自 `toolkit.repo`（RULE_*），不在调用端各自写常量。
"""

from __future__ import annotations

import re
from pathlib import Path

from .base import read_source_text


# ---------------------------------------------------------------------------
# 一、剥注释与按行读
# ---------------------------------------------------------------------------
_LINE_BREAK_RE = re.compile(r"\r\n|\r|\n")


def split_lines(text: str) -> list[str]:
    """按 \\r\\n / \\r / \\n 拆分，不把 \\x0c（VBA 里合法）当换行。"""
    return _LINE_BREAK_RE.split(text)


def strip_vba_lines(text: str) -> list[str]:
    """剥注释，返回逐行列表（与原文一行对一行，行号不错位）。"""
    out: list[str] = []
    for ln in split_lines(text):
        in_str = False
        cut = -1
        for i, c in enumerate(ln):
            if c == '"':
                in_str = not in_str
            elif c == "'" and not in_str:
                cut = i
                break
        body = ln if cut < 0 else ln[:cut]
        if body.strip().lower().startswith("rem "):
            body = ""
        out.append(body.rstrip())
    return out


def strip_vba_comments(text: str) -> str:
    """剥注释的整串文本形式（标识符匹配用，不需要行号时走这个）。"""
    return "\n".join(strip_vba_lines(text))


# 同一进程里反复扫同一批文件（audit 的引用统计要扫全仓多遍）时按 mtime/size 缓存
_CODE_CACHE: dict[tuple[str, int, int], list[str]] = {}


def code_lines(path: Path, use_cache: bool = True) -> list[str]:
    """读源码文件 + 剥注释，按 (路径, mtime, 大小) 缓存；文件改过自动失效。"""
    p = Path(path)
    try:
        st = p.stat()
        key = (str(p), st.st_mtime_ns, st.st_size)
    except OSError:
        return strip_vba_lines(read_source_text(p))
    if use_cache:
        got = _CODE_CACHE.get(key)
        if got is not None:
            return got
    lines = strip_vba_lines(read_source_text(p))
    if use_cache:
        _CODE_CACHE[key] = lines
    return lines


# 过程声明行：Sub / Function / Property Get|Let|Set 后面的名字（VBA 标识符允许中文）
VBA_PROC_RE = re.compile(
    r"^(?:Public |Private |Friend )?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([^\s(]+)", re.M)


def declared_procs(path: Path) -> set[str]:
    """文件里声明的过程名（Sub / Function / Property），按剥注释后的文本匹配。"""
    return set(VBA_PROC_RE.findall(strip_vba_comments(read_source_text(path))))


# ---------------------------------------------------------------------------
# 二、字符串字面量
# ---------------------------------------------------------------------------
STR_LITERAL_RE = re.compile(r'"[^"\n]*"')


def split_code_comment(line: str) -> tuple[str, str]:
    """拆成 (代码, 行尾注释)；' 在双引号外即注释起点。"""
    inq = False
    for i, c in enumerate(line):
        if c == '"':
            inq = not inq
        elif c == "'" and not inq:
            return line[:i].rstrip(), line[i:]
    return line.rstrip(), ""


def mask_strings(line: str) -> str:
    """把字符串字面量换成等长空格（字符串里的名字不是标识符调用，但位置要保住）。"""
    return STR_LITERAL_RE.sub(lambda m: " " * (m.end() - m.start()), line)


# ---------------------------------------------------------------------------
# 三、VBE 导出头（.cls / .frm 的 VERSION+BEGIN…END+Attribute VB_*）
# ---------------------------------------------------------------------------
# 这些行只能在 .cls/.frm 文件里存在；一旦 CodeModule.AddFromString 写进模块正文
# 就是语法错误（VBE 顶上多出 VERSION / BEGIN / MultiUse / END 四行）。
EXPORT_VERSION_RE = re.compile(r"^\s*VERSION\s+[\d.]+(?:\s+[A-Za-z]+)?\s*$", re.IGNORECASE)
EXPORT_BEGIN_RE = re.compile(r"^\s*BEGIN\b", re.IGNORECASE)
EXPORT_END_RE = re.compile(r"^\s*END\s*$", re.IGNORECASE)
EXPORT_ATTR_RE = re.compile(r"^\s*Attribute\s+VB_", re.IGNORECASE)
EXPORT_CONT_RE = re.compile(r"^\s*_\"?")


def strip_export_header(code: str) -> str:
    """
    剥掉 VBA 导出头（VERSION/BEGIN…END 块 + Attribute VB_* 及其续行）。

    只处理文件开头的头部区：碰到第一行真正的代码就停止，正文里的 End Sub、
    End Function 之类不会被误删。返回值用 CRLF 连接（写回 VBE 前的文本形态）。
    """
    lines = code.splitlines()
    out: list[str] = []
    i, n = 0, len(lines)
    in_header = True
    while i < n:
        line = lines[i]
        if in_header:
            if EXPORT_VERSION_RE.match(line):
                i += 1
                continue
            if EXPORT_BEGIN_RE.match(line):
                # 跳过 BEGIN…END（内含 MultiUse / OleObjectBlob 之类的属性行）
                i += 1
                while i < n and not EXPORT_END_RE.match(lines[i]):
                    i += 1
                i += 1
                continue
            if EXPORT_ATTR_RE.match(line):
                i += 1
                # Attribute 太长时 VBA 会拆成续行（下一行以 _ 开头）
                while i < n and EXPORT_CONT_RE.match(lines[i]):
                    i += 1
                continue
            stripped = line.strip()
            if stripped == "" or stripped.startswith("'"):
                out.append(line)          # 头部空行 / 注释原样保留
                i += 1
                continue
            in_header = False
        out.append(line)
        i += 1
    return "\r\n".join(out)


# ---------------------------------------------------------------------------
# 四、成员与事件
# ---------------------------------------------------------------------------
# VBA 标识符：字母/下划线/中文开头，后面可以跟数字
MEMBERS_ID = r"[A-Za-z_\u4e00-\u9fff][\w\u4e00-\u9fff]*"

# 类模块字段/属性/方法声明（成员表用；kind 是给人看的分类）
MEMBER_FIELD_RE = re.compile(rf"(?i)^(?P<acc>Public|Friend|Private)\s+(?P<name>{MEMBERS_ID})\s+As\s+")
MEMBER_PROP_RE = re.compile(
    rf"(?i)^(?P<acc>Public |Friend |Private )?Property\s+(?:Get|Let|Set)\s+(?P<name>{MEMBERS_ID})")
MEMBER_PROC_RE = re.compile(
    rf"(?i)^(?P<acc>Public |Friend |Private )?(?P<kind>Sub|Function)\s+(?P<name>{MEMBERS_ID})")
MEMBER_END_RE = re.compile(r"(?i)^\s*(End Sub|End Function|End Property)\s*$")

# 事件过程：由 Excel/窗体触发，永远是入口，不能算死代码；重名检查时也要跳过
EVENT_SUFFIX = re.compile(
    r"(?i)_(Click|Change|Initialize|Terminate|Activate|Deactivate|Open|Close|BeforeClose|"
    r"BeforeSave|BeforeDoubleClick|BeforeRightClick|Calculate|SelectionChange|KeyDown|"
    r"KeyUp|KeyPress|MouseDown|MouseUp|MouseMove|Enter|Exit|GotFocus|LostFocus|"
    r"AfterUpdate|BeforeUpdate|DblClick|Resize|Layout|QueryClose|Error|AddControl|"
    r"RemoveControl|Scroll|Zoom|NewSheet|SheetActivate|SheetChange|SheetDeactivate|"
    r"SheetBeforeDoubleClick|SheetBeforeRightClick|SheetSelectionChange|WindowResize|Timer)$"
)


def is_event_proc(name: str) -> bool:
    """`CommandButton1_Click` / `Worksheet_Change` 这类宿主触发的事件过程。"""
    return bool(EVENT_SUFFIX.search(name))


def _member_span(lines: list[str], start: int) -> int:
    """过程/属性占了从声明行到 End 行共几行（start 是 1-based 行号）。"""
    for j in range(start, len(lines)):
        if MEMBER_END_RE.match(lines[j].strip()):
            return j - start + 2
    return 1


def parse_members(lines: list[str]) -> dict[str, dict]:
    """成员表：名字 → {kind, access, line, span}（kind 取 字段 / 属性 / 方法）。"""
    members: dict[str, dict] = {}
    for i, ln in enumerate(lines, 1):
        s = ln.strip()
        if not s:
            continue
        m = MEMBER_FIELD_RE.match(s)
        if m:
            members.setdefault(m.group("name"), {
                "kind": "字段", "access": m.group("acc").lower(), "line": i, "span": 1})
            continue
        m = MEMBER_PROP_RE.match(s)
        if m:
            members.setdefault(m.group("name"), {
                "kind": "属性", "access": (m.group("acc") or "public").strip().lower(),
                "line": i, "span": _member_span(lines, i)})
            continue
        m = MEMBER_PROC_RE.match(s)
        if m:
            members.setdefault(m.group("name"), {
                "kind": "方法", "access": (m.group("acc") or "public").strip().lower(),
                "line": i, "span": _member_span(lines, i)})
    return members


# ---------------------------------------------------------------------------
# 五、比对用规范化
# ---------------------------------------------------------------------------
ATTRIBUTE_LINE_RE = re.compile(r"^\s*Attribute\s+", re.IGNORECASE)


def drop_attribute_lines(code: str) -> str:
    """去掉正文里的 Attribute 行；写回 VBE 与回读比对必须用同一口径。"""
    lines = code.splitlines()
    return "\r\n".join(ln for ln in lines if not ATTRIBUTE_LINE_RE.match(ln))


def drop_trailing_placeholder(code: str) -> str:
    """去掉 CodeModule.AddFromString 在末尾留下的游离 `()` 行（VBE 特有占位）。"""
    lines = code.splitlines()
    while lines and not lines[-1].strip():
        lines.pop()
    if lines and lines[-1].strip() == "()":
        lines.pop()
    return "\r\n".join(lines)


def norm_code_lines(text: str) -> list[str]:
    """比对用：剥掉 VBA 导出头、行尾空白与首尾空行。"""
    text = strip_export_header(text).replace("\r\n", "\n").replace("\r", "\n")
    lines = [ln.rstrip() for ln in text.split("\n")]
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
    return lines


def fold_lines(lines: list[str]) -> list[str]:
    """再退一步：忽略大小写与缩进/空白（VBE 会自己改关键字大小写，不该算冲突）。"""
    return [" ".join(ln.split()).lower() for ln in lines]


# ---------------------------------------------------------------------------
# 六、代码口径扫描（裸 New / 手写加速 / 缺过程注释）
# ---------------------------------------------------------------------------
# 允许的白名单与例外记号都在 toolkit.repo 里，这里只负责扫
ACCEL_ASSIGN_RE = re.compile(r"(?i)Application\.(ScreenUpdating|Calculation|EnableEvents)\s*=")
ACCEL_CAPTURE_RE = re.compile(
    r"(?i)Application\.(ScreenUpdating|Calculation|EnableEvents)\s*=\s*([^\s']+)")
ACCEL_OFF = {
    "screenupdating": {"false", "0"},
    "calculation": {"-4135", "xlcalculationmanual", "xlmanual"},
    "enableevents": {"false", "0"},
}
ACCEL_ON = {
    "screenupdating": {"true", "-1"},
    "calculation": {"-4105", "xlcalculationautomatic", "xlautomatic"},
    "enableevents": {"true", "-1"},
}
PROC_END_RE = re.compile(r"(?i)^\s*End\s+(Sub|Function|Property)\s*$")


def _paired_accel_lines(lines: list[str]) -> set[int]:
    """安全执行式手写加速：关闭项在同一过程里、On Error 之后有成对还原，放行。

    只认「同一过程 + 有 On Error + 同一属性后面有还原」三件事都成立的情况；
    普通业务代码里手写关加速、末尾顺手还原但不走错误分支的，仍然报。
    """
    exempt: set[int] = set()
    start = None
    spans: list[tuple[int, int]] = []
    for i, line in enumerate(lines):
        if PROC_HEAD_RE.match(line):
            start = i
        elif start is not None and PROC_END_RE.match(line):
            spans.append((start, i))
            start = None
    if start is not None:
        spans.append((start, len(lines)))

    for start, end in spans:
        body = lines[start:end]
        if not any(re.match(r"(?i)^\s*On\s+Error\b", ln) for ln in body):
            continue
        assigns: list[tuple[int, str, str]] = []
        for off, line in enumerate(body):
            m = ACCEL_CAPTURE_RE.search(line)
            if not m:
                continue
            prop = m.group(1).lower()
            value = m.group(2).strip().lower().rstrip("&")
            assigns.append((off, prop, value))
        for off, prop, value in assigns:
            if value not in ACCEL_OFF.get(prop, set()):
                continue
            for on, prop2, value2 in assigns:
                if on > off and prop2 == prop and value2 in ACCEL_ON.get(prop, set()):
                    exempt.add(start + off + 1)
                    exempt.add(start + on + 1)
                    break
    return exempt
# 过程头：Sub / Function（Property 不算——Get/Let/Set 三段式本来就难在头上放注释）
PROC_HEAD_RE = re.compile(
    rf"(?i)^\s*(?:Public |Private |Friend )?(?:Static )?(?:Sub|Function)\s+({MEMBERS_ID})")


def has_leading_comment(lines: list[str], head_idx: int) -> bool:
    """过程头上是否紧邻注释块（head_idx 为 0 基行号）。

    只看过程头往上：先跳过空行，再要求紧邻行是 `'` / `Rem` 注释。
    连续多行 `'` 块注释算；中间夹空行也认；一旦遇到代码/`End Sub` 就停，
    不再回看任意多行，避免上一过程的注释误算到这一过程。
    """
    j = head_idx - 1
    while j >= 0 and not lines[j].strip():
        j -= 1
    if j < 0:
        return False
    s = lines[j].strip()
    return s.startswith("'") or bool(re.match(r"(?i)^Rem\b", s))


def scan_rules(files, exclude=None) -> dict:
    """
    扫三类"改代码时容易悄悄退回"的口径，返回 {(键): [(文件, 行号, 片段)]}：
      · naked_new    —— 裸 `New 表格工具`（应走 getTL；测试模块与 getTL 本体除外）
      · manual_accel —— 手写 Application.ScreenUpdating/Calculation/EnableEvents 赋值
                        （应走 开启加速／关闭加速；块上方三行内标 [加速例外] 则放行）
      · no_comment   —— 过程（Sub/Function）头上没有紧邻的注释块
                        （只跳过空行；连续 `'` / `Rem` 注释算，事件过程/隐式入口除外）
    exclude 传关键字元组，命中文件名或路径的跳过（与 audit 的 --exclude 同义）。
    """
    from .repo import (IMPLICIT_PROCS, RULE_ALLOW_ACCEL_FILES, RULE_ALLOW_NEW_FILES,
                       RULE_EXEMPT_MARK, RULE_NEW_MARK)

    naked_new: list[tuple[Path, int, str]] = []
    manual_accel: list[tuple[Path, int, str]] = []
    no_comment: list[tuple[Path, int, str]] = []
    keys = [k for k in (exclude or ()) if k]

    for f in files:
        if keys and any(k in str(f) for k in keys):
            continue
        try:
            lines = read_source_text(f).splitlines()
        except Exception:
            continue
        paired_accel = _paired_accel_lines(lines)
        for i, raw in enumerate(lines, 1):
            s = raw.strip()
            is_comment = s.startswith("'")
            if not is_comment and RULE_NEW_MARK in raw and f.name not in RULE_ALLOW_NEW_FILES:
                naked_new.append((f, i, s))
            if not is_comment and ACCEL_ASSIGN_RE.search(raw):
                m = ACCEL_CAPTURE_RE.search(raw)
                prop = m.group(1).lower() if m else ""
                value = m.group(2).strip().lower().rstrip("&") if m else ""
                # 只盯“关闭”赋值；还原（True/Automatic）本身不是违规，
                # 否则错误分支里重复的还原会被当成第二批手写加速。
                if (prop and value in ACCEL_OFF.get(prop, set())
                        and f.name not in RULE_ALLOW_ACCEL_FILES and i not in paired_accel):
                    ctx = " ".join(lines[max(0, i - 3):i])
                    if RULE_EXEMPT_MARK not in ctx:
                        manual_accel.append((f, i, s))
            m = None if is_comment else PROC_HEAD_RE.match(raw)
            if not m:
                continue
            nm = m.group(1)
            if is_event_proc(nm) or nm.lower() in IMPLICIT_PROCS:
                continue
            if not has_leading_comment(lines, i - 1):
                no_comment.append((f, i, s))
    return {"naked_new": naked_new, "manual_accel": manual_accel, "no_comment": no_comment}
