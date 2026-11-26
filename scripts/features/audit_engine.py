# -*- coding: utf-8 -*-
"""audit_engine —— audit 只读视角（structure / indicators / overview）共用的规则引擎门面。

职责边界（架构方案 §4 / 附录 A.4 领域服务）：
  · 只做纯计算与规则判定：枚举源码、解析成员、对账、汇总；
  · 不打印、不碰 COM、不写文件；结果以普通 dict / 列表上交 feature 层；
  · 规则数据（阈值、白名单、分组登记、指标口径）只在 audit_rules 一份，feature 层不得复制阈值表。

内部切片（同包叶子，不建子包）：
  · audit_rules      规则数据 / 正则 / 纯辅助函数（单一来源）
  · audit_scan       范围枚举、读取缓存、成员可达性扫描与单模块盘点
  · audit_structure  structure 结构问题分组对账
  · audit_indicators indicators H01–H19 指标计算
本文件保留薄门面：``AuditEngine`` 组合上述 mixin，并按旧契约重导出规则常量与 helper。
"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Mapping

from toolkit.base import entry_command, read_source_text
from toolkit.contract import Failure
from toolkit.repo import VERSION_RE, rel_path

from .audit_rules import (BACKUP_BATCH_LIMIT, BACKUP_SIZE_MB_LIMIT,
                          BUILTIN_CALLS, FATAL_STRUCTURE_GROUPS,
                          NOTE_ONLY_GROUPS, SHADOW_KNOWN,
                          STATUSBAR_ALLOW_FILES, _ignore_set, _jsonable,
                          _names_in_list)
from .audit_scan import ScanMixin
from .audit_structure import StructureMixin
from .audit_indicators import IndicatorMixin

__all__ = [
    "AuditEngine",
    "BACKUP_BATCH_LIMIT", "BACKUP_SIZE_MB_LIMIT", "BUILTIN_CALLS",
    "FATAL_STRUCTURE_GROUPS", "NOTE_ONLY_GROUPS", "SHADOW_KNOWN",
    "STATUSBAR_ALLOW_FILES", "_ignore_set", "_jsonable", "_names_in_list",
]


class AuditEngine(IndicatorMixin, StructureMixin, ScanMixin):
    """audit 三个只读视角共用的纯计算引擎。

    构造时只吃仓库根与配置 dict；文件枚举、源码读取、成员解析都在方法内完成，
    结果上交 feature 层编排退出码与人话。
    """

    def __init__(self, root: Path | str, config: Mapping[str, Any] | None = None):
        self.root = Path(root)
        self.config = dict(config or {})
        self.projects: list[dict] = list(self.config.get("projects", []) or [])
        self.defaults: dict = dict(self.config.get("defaults", {}) or {})
        self._code_cache: dict[Path, list[str]] = {}
        self._raw_cache: dict[Path, list[str]] = {}
        # common_analyses 本轮跳过的模块原因（适配层取回拼 warnings；不改公开返回形状）
        self.last_skipped: list[str] = []

    # ------------------------------------------------------------------
    # overview：公共库汇总 + 结构意见
    # ------------------------------------------------------------------
    def common_analyses(self, projects: list[str] | None = None,
                        exclude: list[str] | None = None) -> list[dict]:
        out: list[dict] = []
        self.last_skipped = []
        for f in self.common_files(exclude):
            try:
                out.append(self.analyze_target(f, projects, exclude))
            except Failure as e:
                self.last_skipped.append(f"跳过 {f.name}：{e}")
        return out

    def summary(self, res: dict) -> dict:
        """一个模块的浓缩结果，给总览表用。"""
        ref_files: set[str] = set()
        for r in res["rows"]:
            ref_files |= set(r["ext_files"])
        return {
            "module": res["cls"],
            "kind": self._kind_of(res["target"]),
            "path": str(rel_path(res["target"])),
            "members": len(res["rows"]),
            "dead": len(res["dead"]),
            "test": len(res["only_test"]),
            "single": len(res["single"]),
            "active": sum(1 for r in res["rows"] if r["verdict"] == "活跃"),
            "lines": sum(r["span"] for r in res["dead"]),
            "ref_files": len(ref_files),
            "infra": bool(VERSION_RE.search(read_source_text(res["target"]))),
        }

    def overview_actions(self, analyses: list[dict], issues: dict) -> list[tuple[int, str]]:
        """整体意见（按影响排序前的权重表）。"""
        actions: list[tuple[int, str]] = []
        for res in analyses:
            s = self.summary(res)
            if not s["ref_files"] and not s["infra"]:
                actions.append((max(s["members"], 1),
                                f"「{s['module']}」{s['members']} 个成员零外部引用 —— 整个模块可退出公共库"
                                "（先确认不是按钮/快捷键宏入口）"))
            if s["dead"]:
                names = "、".join(r["member"] for r in res["dead"][:8])
                more = "…" if s["dead"] > 8 else ""
                actions.append((s["lines"],
                                f"「{s['module']}」删不可达 {s['dead']} 个成员，约 {s['lines']} 行：{names}{more}"))
            if s["single"]:
                actions.append((s["single"],
                                f"「{s['module']}」{s['single']} 个成员只被单个文件引用 —— 下沉回那张表或改 Private"))
            if s["test"]:
                actions.append((s["test"],
                                f"「{s['module']}」{s['test']} 个成员只被自测模块引用 —— 业务侧可删，测试模块同步瘦身"))
            dyn = [r["member"] for r in res["rows"] if r["dynamic"]]
            if dyn:
                actions.append((1, f"「{s['module']}」{len(dyn)} 个成员出现在 Run/Evaluate 字符串里 —— "
                                   f"删前人工确认：{'、'.join(dyn[:6])}"))
        if issues["stale"]:
            actions.append((10, f"旧副本 {len(issues['stale'])} 处 —— "
                            + entry_command("sync.py", "align", "--project", "<表名>", "--apply")))
        if issues["orphan"]:
            actions.append((5, f"孤儿壳 {len(issues['orphan'])} 处 —— "
                            + entry_command("sync.py", "shells", "--mode", "clean", "--apply")))
        if issues["kind"]:
            actions.append((3, f"壳类型不符 {len(issues['kind'])} 处 —— 删壳后 "
                            + entry_command("sync.py", "shells")))
        if issues["missing"]:
            actions.append((5, f"缺壳 {len(issues['missing'])} 处 —— "
                            + entry_command("sync.py", "shells")))
        if issues["dup"] or issues["shadow"]:
            actions.append((len(issues["dup"]) + len(issues["shadow"]),
                            f"重名过程 {len(issues['dup']) + len(issues['shadow'])} 个 —— "
                            "用 structure 看清单，改名或改用「模块名.过程」限定"))
        if issues["doc_orphan"]:
            actions.append((len(issues["doc_orphan"]),
                            f"文档模块对账：源码有 {len(issues['doc_orphan'])} 个模块工程里已没有（表页已删）"
                            " —— " + entry_command("sync.py", "pull", "--all-project", "--apply")
                            + " 清掉残留文件"))
        if issues["doc_ungrabbed"]:
            actions.append((len(issues["doc_ungrabbed"]),
                            f"文档模块对账：工程有 {len(issues['doc_ungrabbed'])} 个表页模块没源码"
                            " —— " + entry_command("sync.py", "pull", "--all-project", "--apply")
                            + " 补抓"))
        return actions
