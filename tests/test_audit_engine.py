# -*- coding: utf-8 -*-
"""AuditEngine 独立单测：规则判定 / 汇总 / 单一来源证据。

离线：tempfile + 源码文本；不启动 Excel、不碰 COM、不写文件。
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))

from features import audit_b  # noqa: E402
from features.audit_engine import (  # noqa: E402
    AuditEngine,
    BUILTIN_CALLS,
    FATAL_STRUCTURE_GROUPS,
    NOTE_ONLY_GROUPS,
    SHADOW_KNOWN,
    STATUSBAR_ALLOW_FILES,
    BACKUP_BATCH_LIMIT,
    BACKUP_SIZE_MB_LIMIT,
    _ignore_set,
)
from toolkit.base import PKG_ROOT  # noqa: E402

CFG = {
    "projects": [{"name": "甲表", "src": "src/projects/甲表"}],
    "defaults": {"common_dir": "src/common"},
}

CLEAN_COMMON = {
    "src/common/工具.bas": (
        "Attribute VB_Name = \"工具\"\n"
        "Option Explicit\n\n"
        "' 相加\n"
        "Public Function 相加(a As Long, b As Long) As Long\n"
        "    相加 = a + b\n"
        "End Function\n"
    ),
}

INDICATOR_WARN = {
    "src/common/脚本.bas": (
        "Attribute VB_Name = \"脚本\"\n"
        "Option Explicit\n"
        "' 主流程\n"
        "Public Sub 主流程()\n"
        "    Call 未定义的过程\n"
        "End Sub\n"
    ),
}

# 跨视角一致性夹具：公共库一个被项目调用的函数 + 一个无人调用的成员。
CONSISTENCY_SRC = {
    "src/common/工具.bas": (
        'Attribute VB_Name = "工具"\n'
        "Option Explicit\n\n"
        "' 相加\n"
        "Public Function AddOne(ByVal x As Long) As Long\n"
        "    AddOne = x + 1\n"
        "End Function\n"
        "Public Sub 无人调用()\n"
        "End Sub\n"
    ),
    "src/projects/甲表/调用.bas": (
        'Attribute VB_Name = "调用"\n'
        "Option Explicit\n"
        "Public Sub 入口()\n"
        "    Dim n As Long\n"
        "    n = AddOne(1)\n"
        "End Sub\n"
    ),
}

EMPTY_SRC = "Option Explicit\n' 只有注释，没有过程\n"


def _make_root(files: dict[str, str]) -> Path:
    d = Path(tempfile.mkdtemp())
    for rel, text in files.items():
        p = d / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(text, encoding="utf-8", newline="\n")
    return d


class RuleAccountingTest(unittest.TestCase):
    def test_total_excludes_note_only(self):
        issues = {g: [] for g in FATAL_STRUCTURE_GROUPS | NOTE_ONLY_GROUPS}
        issues["shadow_known"] = [("分析", "x.bas", ["公共"])]
        issues["dup"] = []
        self.assertEqual(AuditEngine.issues_total(issues), 0)

    def test_fatal_counts_only_fatal_groups(self):
        issues = {g: [] for g in ("stale", "orphan", "kind", "missing", "dup",
                                  "shadow", "shadow_known", "doc_orphan",
                                  "doc_ungrabbed", "sheet_mismatch", "naked_new",
                                  "manual_accel", "no_comment", "binding_target",
                                  "suffix")}
        issues["sheet_mismatch"] = [1]
        issues["binding_target"] = [1, 2]
        issues["dup"] = [1]
        issues["shadow"] = [1]
        issues["shadow_known"] = [1]
        issues["stale"] = [1]
        self.assertEqual(AuditEngine.issues_fatal(issues), 5)

    def test_ignore_set_splits_commas_and_lowercases(self):
        self.assertEqual(_ignore_set(["H05a,H06", " manual_accel ", "H05a"]),
                         {"h05a", "h06", "manual_accel"})


class EngineComputationTest(unittest.TestCase):
    def test_summary_shape(self):
        root = _make_root(CLEAN_COMMON)
        engine = AuditEngine(root, CFG)
        analyses = engine.common_analyses()
        self.assertEqual(len(analyses), 1)
        summary = engine.summary(analyses[0])
        for key in ("module", "kind", "path", "members", "dead", "test", "single",
                    "active", "lines", "ref_files", "infra"):
            self.assertIn(key, summary)
        self.assertEqual(summary["members"], 1)

    def test_indicator_h01_unknown_call_warns(self):
        root = _make_root(INDICATOR_WARN)
        engine = AuditEngine(root, CFG)
        rows = engine.indicator_rows()
        h01 = next(r for r in rows if r["id"] == "H01")
        self.assertEqual(h01["state"], "警告")
        self.assertGreaterEqual(h01["count"], 1)

    def test_structure_rows_are_jsonable(self):
        root = _make_root(CLEAN_COMMON)
        engine = AuditEngine(root, CFG)
        issues = engine.structure_issues()
        rows = AuditEngine.structure_rows(issues)
        import json
        json.dumps(rows)  # Path / tuple 都已转成 JSON 安全形式
        self.assertTrue(all("group" in r and "count" in r and "items" in r
                            for r in rows))

    def test_live_probe_missing_keeps_h16_h17_unrun(self):
        root = _make_root(CLEAN_COMMON)
        engine = AuditEngine(root, CFG)
        live = {"checked": 0, "total": 1, "h16": [], "h17": [], "reason": "未连接"}
        rows = engine.indicator_rows(live_probe=live)
        for idx in ("H16", "H17"):
            row = next(r for r in rows if r["id"] == idx)
            self.assertEqual(row["state"], "未跑")


class SingleSourceTest(unittest.TestCase):
    """规则单一来源：阈值/白名单/分组登记只在 engine 一份，feature 层不得复制。"""

    def test_rule_tables_live_in_engine(self):
        self.assertTrue(BUILTIN_CALLS)
        self.assertTrue(STATUSBAR_ALLOW_FILES)
        self.assertEqual(SHADOW_KNOWN, {"分析", "输出", "分词", "test1", "获取选择的数据"})
        self.assertEqual(FATAL_STRUCTURE_GROUPS,
                         {"sheet_mismatch", "binding_target", "dup", "shadow"})
        self.assertEqual(NOTE_ONLY_GROUPS, {"shadow_known"})
        self.assertEqual(BACKUP_BATCH_LIMIT, 60)
        self.assertEqual(BACKUP_SIZE_MB_LIMIT, 450)

    def test_feature_layer_has_no_rule_copy(self):
        source = Path(audit_b.__file__).read_text(encoding="utf-8")
        for symbol in ("BUILTIN_CALLS", "STATUSBAR_ALLOW_FILES", "SHADOW_KNOWN",
                       "FATAL_STRUCTURE_GROUPS", "NOTE_ONLY_GROUPS",
                       "BACKUP_BATCH_LIMIT", "BACKUP_SIZE_MB_LIMIT"):
            self.assertNotIn(symbol, source, f"feature 层复制了规则表 {symbol}")


class CrossViewConsistencyTest(unittest.TestCase):
    """同一输入、不同视角结论一致：members ↔ common ↔ project ↔ overview。"""

    def _engine(self, root: Path) -> AuditEngine:
        return AuditEngine(root, CFG)

    def test_members_and_common_agree_on_same_module(self):
        root = _make_root(CONSISTENCY_SRC)
        engine = self._engine(root)
        common_path = root / "src" / "common" / "工具.bas"

        res = engine.analyze_target(common_path)
        analyses = engine.common_analyses()
        again = next(a for a in analyses if a["cls"] == res["cls"])

        # 成员清单逐名一致
        self.assertEqual([r["member"] for r in again["rows"]],
                         [r["member"] for r in res["rows"]])
        # 浓缩摘要两次计算一致，且与行级数据对得上
        self.assertEqual(engine.summary(res), engine.summary(again))
        summary = engine.summary(res)
        self.assertEqual(summary["members"], len(res["rows"]))
        ref_files = {f for r in res["rows"] for f in r["ext_files"]}
        self.assertEqual(summary["ref_files"], len(ref_files))

    def test_members_and_project_dep_rows_agree(self):
        root = _make_root(CONSISTENCY_SRC)
        engine = self._engine(root)
        common_path = root / "src" / "common" / "工具.bas"
        res = engine.analyze_target(common_path)

        dep = {d["module"]: d for d in engine.dep_rows("甲表")}
        self.assertIn("工具", dep)
        # 沙箱根不等于全局 ROOT，ext_files 键为绝对路径串，按绝对路径对齐
        proj_rel = {str(p) for p in engine.project_files("甲表")}

        refs = used = 0
        ext_files: set[str] = set()
        for r in res["rows"]:
            pf = {f: v for f, v in r["ext_files"].items() if f in proj_rel}
            if pf:
                used += 1
                refs += sum(pf.values())
                ext_files |= pf.keys()

        self.assertEqual(dep["工具"]["refs"], refs)
        self.assertEqual(dep["工具"]["used"], used)
        self.assertEqual(dep["工具"]["files"], len(ext_files))

    def test_project_and_overview_used_module_sets_agree(self):
        root = _make_root(CONSISTENCY_SRC)
        engine = self._engine(root)
        # overview 按同一张表收窄口径 common_analyses(projects=[甲表])
        analyses = engine.common_analyses(projects=["甲表"])
        sums = {s["module"]: s for s in (engine.summary(a) for a in analyses)}

        dep_used = {d["module"] for d in engine.dep_rows("甲表") if d["refs"] > 0}
        overview_used = {m for m, s in sums.items() if s["ref_files"] > 0}
        self.assertEqual(dep_used, overview_used)

    def test_empty_module_failure_carries_preflight_code(self):
        from toolkit.contract import Failure
        root = _make_root({"src/common/空.bas": EMPTY_SRC})
        engine = self._engine(root)
        with self.assertRaises(Failure) as cm:
            engine.analyze_target(root / "src" / "common" / "空.bas")
        self.assertEqual(cm.exception.code, "SOURCE_PREFLIGHT")
        self.assertEqual(cm.exception.stage, "parse")

    def test_common_analyses_keeps_list_shape_and_records_skipped(self):
        root = _make_root({**CONSISTENCY_SRC, "src/common/空.bas": EMPTY_SRC})
        engine = self._engine(root)
        out = engine.common_analyses()
        self.assertIsInstance(out, list)  # 公开返回形状不变
        self.assertEqual(len(out), 1)
        self.assertEqual(len(engine.last_skipped), 1)
        self.assertTrue(engine.last_skipped[0].startswith("跳过 空.bas："))


class AbsoluteEntryCommandTest(unittest.TestCase):
    """audit 产出的命令必须是 entry_command() 拼的绝对入口，不能照抄包内相对路径。"""

    def test_h10_suggested_carries_absolute_script(self):
        with tempfile.TemporaryDirectory() as d:
            bk = Path(d) / ".backup"
            bk.mkdir()
            for i in range(61):
                (bk / f"b{i:02d}").mkdir()
            cfg = {"projects": [], "defaults": {"backup_dir": str(bk), "backup_keep": 10}}
            rows = AuditEngine(Path(d), cfg).indicator_rows(ignore="H05a,H06")
        h10 = next(r for r in rows if r["id"] == "H10")
        self.assertIn(str(PKG_ROOT / "scripts" / "sync.py"), h10["suggested_command"])
        self.assertIn(Path(sys.executable).name, h10["suggested_command"])

    def test_overview_actions_carry_absolute_script(self):
        issues = {k: [] for k in ("stale", "orphan", "kind", "missing", "dup",
                                  "shadow", "doc_orphan", "doc_ungrabbed")}
        issues["stale"] = [("x.bas", Path("y.bas"))]
        engine = AuditEngine(Path(tempfile.mkdtemp()), CFG)
        actions = engine.overview_actions([], issues)
        text = " ".join(t for _, t in actions)
        self.assertIn(str(PKG_ROOT / "scripts" / "sync.py"), text)


if __name__ == "__main__":
    unittest.main()
