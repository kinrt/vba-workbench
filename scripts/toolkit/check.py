# -*- coding: utf-8 -*-
"""
check —— 通用校验层：提醒收集 + 模块集合对账
============================================================
为什么有这个模块：写回、抓取、运行宏、自动写回、结构审计、风格检查各自都会发现问题，
以前每处自己判一遍、自己打一遍，口径容易漂。这里把"判定"和"收集"收成一份，
各脚本只负责**取数据**（COM、读文件）和**怎么打印**（控制台 / md / JSON）。

只放纯逻辑：不 import win32com、不碰 Excel、不读仓库结构、不做修复。
凡是需要读 VBProject 的，由调用方读成 `{模块名: 类型}` 交给这里。

三层东西：
  1. Notes              —— 提醒收集器（失败 / 警告 / 预检 / 校验），同文本自动合并 ×N
  2. ModuleSetDiff      —— 模块集合"写前 / 写后"对比结果（added / lost）
  3. check_* 系列函数    —— 纯集合判定，返回名单，由调用方决定怎么报

典型用法（写回）：

    notes = Notes()
    d = check_module_set(before, after)
    notes.extend(d.notes())
    for line in notes.render("写回提醒"):
        msg(line)

典型用法（自动写回多轮累计）：

    notes.add("校验", "模块清单一致（19 个）")
    ...
    notes.render("自动写回提醒（本次会话）")     # 重复的条目合并成（×N）
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field

# 提醒分级：顺序即汇总时的打印顺序（失败最要紧，排最前）
NOTE_LEVELS = ("失败", "警告", "预检", "校验")


class Notes:
    """
    一次操作里的提醒收集器。只收集、不打印——打印交给调用方（msg / warn / 写 md）。
    同一条文本重复记多次只算一行（显示 ×N），自动写回一轮轮重复不会刷屏。
    """

    def __init__(self, title: str = "本次提醒汇总"):
        self.title = title
        self._rows: list[tuple[str, str]] = []

    def add(self, level: str, text: str) -> None:
        """记一条。level 建议取 NOTE_LEVELS 之一（别的值也能存，只是排在最后）。"""
        if text:
            self._rows.append((level, text))

    def extend(self, rows) -> None:
        """并入另一批 (level, text)。"""
        for row in rows:
            self.add(*row)

    def merge(self, other: "Notes") -> None:
        """并入另一个 Notes（多项目 / 多脚本合并用）。"""
        self.extend(other._rows)

    def clear(self) -> None:
        self._rows.clear()

    def __len__(self) -> int:
        return len(self._rows)

    def __bool__(self) -> bool:
        return bool(self._rows)

    def counts(self) -> dict[str, int]:
        """{级别: 条数}（按原始条数，不合并重复文本）。"""
        return dict(Counter(level for level, _ in self._rows))

    def has(self, *levels: str) -> bool:
        """有没有某几级（判断要不要以非 0 退出：has("失败", "警告")）。"""
        return any(level in levels for level, _ in self._rows)

    def rows(self) -> list[tuple[str, str]]:
        return list(self._rows)

    def render(self, title: str | None = None, *, clear: bool = True) -> list[str]:
        """
        渲染成行列表（不带换行符），调用方自己决定 msg / 写文件。
        空的时候给一句"无异常"，避免"没输出 = 没检查"的误判。
        """
        head = title or self.title
        if not self._rows:
            if clear:
                self.clear()
            return [f"── {head}：无异常 ──"]

        counts = Counter(level for level, _ in self._rows)
        order = [lv for lv in NOTE_LEVELS if counts[lv]]
        order += sorted(lv for lv in counts if lv not in NOTE_LEVELS)
        head_line = f"── {head}：" + " · ".join(f"{lv} {counts[lv]}" for lv in order) + " ──"

        lines = [head_line]
        for level in order:
            items = Counter(t for lv, t in self._rows if lv == level)
            lines.append(f"  [{level}]")
            for text, n in items.items():
                lines.append(f"     · {text}" + (f"（×{n}）" if n > 1 else ""))
        if clear:
            self.clear()
        return lines


@dataclass
class ModuleSetDiff:
    """模块集合对比结果：写回只该改正文，不该增删模块，所以两边都空才算 ok。"""

    added: list[str] = field(default_factory=list)
    lost: list[str] = field(default_factory=list)
    # 疑似 Import 抢名残留：[(后缀名, 原名)]，如 ("表格工具1", "表格工具")
    suffix: list[tuple[str, str]] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.added and not self.lost

    def notes(self, level: str = "警告") -> list[tuple[str, str]]:
        """转成 Notes 条目：一致记一条"校验"通过，有差异逐项记警告。"""
        if self.ok:
            return [("校验", "模块清单与写回前一致 ✔")]
        rows: list[tuple[str, str]] = []
        if self.added:
            rows.append((level, f"写回后多出模块：{'、'.join(self.added)}"))
        if self.suffix:
            pairs = "、".join(f"{n}（原名 {b}）" for n, b in self.suffix)
            rows.append((level, f"其中疑似 Import 抢名：{pairs}"))
        if self.lost:
            rows.append((level, f"写回后少了模块：{'、'.join(self.lost)}"))
        return rows


def check_module_set(before: dict | None, after: dict | None,
                     src_names=None) -> ModuleSetDiff | None:
    """
    比"写回前 / 写回后"的模块快照（dict 或任何可迭代的模块名都行）。
    任一边拿不到（None / 空）就返回 None——宁可"没校验"，也不要假装校验通过。
    src_names 给定时顺带算出 diff.suffix（源码里没有、且原名还在的那种多出模块）。
    """
    if not before or not after:
        return None
    return ModuleSetDiff(
        added=sorted(set(after) - set(before)),
        lost=sorted(set(before) - set(after)),
        suffix=check_suffix_modules(after, src_names or (), only=sorted(set(after) - set(before))),
    )


# 结尾带数字的模块名（`表格工具1` / `Sheet11` 都形似），要结合"原名还在"才能定性
SUFFIX_NAME_RE = re.compile(r"[\w\u4e00-\u9fff]+\d")


def check_suffix_modules(comp_names, src_names=(), only=None) -> list[tuple[str, str]]:
    """
    判"后缀模块"（VBE Import 抢名的典型残留），返回 [(后缀名, 原名)]：
      · 名字形似 `X<数字>`（SUFFIX_NAME_RE）
      · 去掉末尾那一位数字后的 `X` 也在同一批名字里（否则 `Sheet11` 这种正常名字会被误判）
      · 源码目录里没有 `X<数字>` 这个模块（源码里真有的，说明是刻意命名）
    only 给定时只在这批名字里找（写回场景直接传 diff.added，避免全量比较）。
    """
    pool = {n for n in comp_names if n}
    src = {n for n in src_names if n}
    targets = [n for n in (only if only is not None else pool) if n]
    found: list[tuple[str, str]] = []
    for n in targets:
        if n in src or not SUFFIX_NAME_RE.fullmatch(n):
            continue
        base = n[:-1]
        if base in pool:
            found.append((n, base))
    return sorted(found)


def check_doc_modules(src_docs, comp_names) -> tuple[list[str], list[str]]:
    """
    文档模块对账，返回 (孤儿, 未抓取)：
      孤儿   —— 源码里有、工程里没有（表页被删，写回只能跳过，该跑抓取清理文件）
      未抓取 —— 工程里有、源码里没有（写回不会新建，但可能忘了抓）
    两边传的都是模块名集合（不要传文件名，文档模块文件名带表页名）。
    """
    src = {n for n in src_docs if n}
    comp = {n for n in comp_names if n}
    orphan = sorted(src - comp)
    ungrabbed = sorted(comp - src)
    return orphan, ungrabbed


def check_sheet_modules(pages, doc_names) -> list[tuple[str, str, str]]:
    """
    表页 ↔ 文档模块一一对应（体检指标 H19），返回 [(类型, 对象, 细节)]，类型三种：
      数量   —— 表页数 ≠ 文档模块数（ThisWorkbook 不算；不等就必然有一张表没代码）
      无代码 —— 这个表页的 codeName 在文档模块里找不到（T56 的 Sheet10 空壳 / 宿主丢失）
      无宿主 —— 有个文档模块没有表页的 codeName 指向它（T56 的 Sheet10_orphan）

    pages 传 vbabin.sheet_pages 的返回 [{"name": 表页名, "code": codeName}, …]；
    doc_names 传 vbabin.vba_documents 的返回（工程流 PROJECT 里的 Document 登记，
    文件名不参与判定，所以 `Sheet10_orphan` 这类怪名也能被判成"无宿主"）。
    """
    page_list = list(pages or [])
    docs = sorted(n for n in doc_names or [] if n and n.lower() != "thisworkbook")
    codes = [str(p.get("code") or "") for p in page_list]
    found: list[tuple[str, str, str]] = []
    if len(page_list) != len(docs):
        found.append(("数量", f"{len(page_list)} 个表页 / {len(docs)} 个文档模块", "两边必须相等"))
    for p in page_list:
        name, code = str(p.get("name") or ""), str(p.get("code") or "")
        if not code:
            found.append(("无代码", name, "表页没有 codeName"))
        elif code not in docs:
            found.append(("无代码", f"{name}({code})", "工程里没有这个文档模块"))
    for n in docs:
        if n not in codes:
            found.append(("无宿主", n, "没有表页的 codeName 指向它"))
    return found


def check_export_set(module_names, file_names, skip=()) -> list[str]:
    """
    抓取后核对：工程里的模块是否都落成了源码文件，返回缺文件的那批。
    skip 传公共库模块名——它们在项目目录里只是占位壳（会被 source_files 跳过），
    不排除就会被误报成"没有源码文件"。
    """

    have = {n for n in file_names if n}
    ignored = {n for n in skip if n}
    return sorted(set(module_names) - have - ignored)
