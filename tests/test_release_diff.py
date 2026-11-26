# -*- coding: utf-8 -*-
"""10 号切片测试：基线录制 introspect、差分 harness 判定、安装副本升级。

不启动 Excel、不碰业务表；真实基线用例在标签/仓库不可用时自动跳过。
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))
sys.path.insert(0, str(PKG / "tests" / "tools"))
sys.path.insert(0, str(PKG / "tests" / "baseline"))

import diff_harness as harness  # noqa: E402
import record_1_5_6 as recorder  # noqa: E402
import upgrade_install as upgrade  # noqa: E402
from toolkit.contract import MutationMode  # noqa: E402
from toolkit.registry import CommandSpec, ProjectMode  # noqa: E402

BASELINE_DIR = PKG / "tests" / "baseline" / "1.5.6"


def _spec(verb: str, entry: str = "sync", *, args=(), json_ok=True,
          mutation=MutationMode.NONE, project=ProjectMode.NONE) -> CommandSpec:
    return CommandSpec(verb=verb, entry=entry, handler=f"features.core:{verb}",
                       mutation=mutation, project=project, args=tuple(args),
                       supports_json=json_ok, exit_codes=(0, 1, 2),
                       summary="示例", tests=())


def _old(verbs: dict[str, dict]) -> dict:
    return {"sync": verbs, "audit": {}, "agent": {}}


class RecorderIntrospectTest(unittest.TestCase):
    def test_introspect_synthetic_parser(self):
        parser = argparse.ArgumentParser(prog="sync.py")
        sub = parser.add_subparsers(dest="command")
        p = sub.add_parser("version")
        p.add_argument("--json", action="store_true")
        p.add_argument("--no-input", action="store_true")
        q = sub.add_parser("push")
        q.add_argument("--project")
        q.add_argument("--apply", action="store_true")
        q.add_argument("--mode", choices=("create", "clean"), default="create")

        surface = recorder.introspect_parser(parser)
        self.assertIn("version", surface["verbs"])
        self.assertIn("push", surface["verbs"])
        self.assertTrue(surface["verbs"]["version"]["json_ok"])
        self.assertTrue(surface["verbs"]["push"]["apply"])
        self.assertTrue(surface["verbs"]["push"]["project"])
        mode = next(a for a in surface["verbs"]["push"]["args"] if "mode" in a["names"])
        self.assertEqual(mode["kind"], "choice")
        self.assertEqual(mode["choices"], ["clean", "create"])

    def test_normalize_strips_paths_and_ts(self):
        text = "失败于 C:\\Users\\<用户>\\表格\\示例表A V1.xlsm @ 2026-11-22T10:00:00+08:00"
        out = recorder.normalize_text(text)
        self.assertNotIn("示例表A", out)
        self.assertIn("<path>", out)
        self.assertIn("<ts>", out)


class HarnessSurfaceTest(unittest.TestCase):
    def test_equivalent_surface_has_no_failed(self):
        specs = (_spec("version"), _spec("doctor"))
        old = harness._new_surface(specs)  # 同形构造旧侧
        rows = harness.compare_surface(old, harness._new_surface(specs))
        self.assertEqual([r for r in rows if r["status"] == harness.STATUS_FAILED], [])
        self.assertTrue(all(r["status"] == harness.STATUS_MATCH for r in rows))

    def test_detects_removed_and_added_args(self):
        new_specs = (_spec("push", args=()),)
        ledger = {"entries": {"sync": {"verbs": {"push": {"args": [{"names": ["module"]}],
                                                "json_ok": True, "apply": True,
                                                "project": True, "all_project": False}}},
                              "audit": {"verbs": {}}, "agent": {"verbs": {}}}}
        old = harness._old_surface(ledger)
        new = harness._new_surface(new_specs)
        rows = harness.compare_surface(old, new)
        fields = {r["field"] for r in rows}
        self.assertIn("arg-removed", fields)

    def test_unapproved_difference_fails(self):
        old = {"sync": {"version": {"args": [], "arg_detail": {}, "json_ok": True,
                                     "apply": False, "project": False, "all_project": False}},
               "audit": {}, "agent": {}}
        new = {"sync": {"version": {"args": [], "arg_detail": {}, "json_ok": False,
                                     "apply": False, "project": False, "all_project": False}},
               "audit": {}, "agent": {}}
        rows = harness.compare_surface(old, new)
        report = harness.DiffReport()
        deltas = {"approved_deltas": [], "pending": []}
        for row in rows:
            report.add(row, deltas)
        self.assertTrue(report.failed, "白名单外的 json 能力差异必须判失败")
        self.assertEqual(report.failed[0]["status"], harness.STATUS_FAILED)

    def test_approved_delta_absorbs_difference(self):
        old = {"sync": {"version": {"args": [], "arg_detail": {}, "json_ok": True,
                                     "apply": False, "project": False, "all_project": False}},
               "audit": {}, "agent": {}}
        new = {"sync": {"version": {"args": [], "arg_detail": {}, "json_ok": False,
                                     "apply": False, "project": False, "all_project": False}},
               "audit": {}, "agent": {}}
        rows = harness.compare_surface(old, new)
        deltas = {"approved_deltas": [{"entry": "sync", "verb": "version",
                                        "field": "json", "reason": "有意调整"}],
                  "pending": []}
        report = harness.DiffReport()
        for row in rows:
            report.add(row, deltas)
        self.assertEqual(report.failed, [])
        self.assertTrue(any(r["status"] == harness.STATUS_APPROVED for r in report.rows))

    def test_pending_gap_is_not_failure(self):
        old = {"sync": {"pull": {"args": [], "arg_detail": {}, "json_ok": True,
                                  "apply": True, "project": True, "all_project": True}},
               "audit": {}, "agent": {}}
        new = {"sync": {}, "audit": {}, "agent": {}}
        rows = harness.compare_surface(old, new)
        deltas = {"approved_deltas": [],
                  "pending": [{"entry": "sync", "verb": "pull",
                               "field": "command-missing-in-new"}]}
        report = harness.DiffReport()
        for row in rows:
            report.add(row, deltas)
        self.assertEqual(report.failed, [])
        self.assertTrue(any(r["status"] == harness.STATUS_PENDING for r in report.rows))

    def test_markdown_has_five_columns(self):
        report = harness.DiffReport()
        report.add({"entry": "sync", "verb": "version", "argv": ["version", "--json"],
                    "field": "exit_code", "old": 0, "new": 0,
                    "status": harness.STATUS_MATCH},
                   {"approved_deltas": [], "pending": []})
        md = report.to_markdown()
        for col in ("命令", "输入", "旧输出", "新输出", "差异结论"):
            self.assertIn(col, md)

    def test_alignment_marks_migrated_and_pending(self):
        ledger = {"entries": {"sync": {"verbs": {
            "version": {"args": [], "json_ok": True, "apply": False,
                        "project": False, "all_project": False},
            "pull": {"args": [{"names": ["no-backup"]}], "json_ok": True, "apply": True,
                     "project": True, "all_project": True}}},
            "audit": {"verbs": {}}, "agent": {"verbs": {}}}}
        specs = (_spec("version"),)
        md = harness.alignment_markdown(ledger, specs)
        self.assertIn("对齐（无差异）", md)
        self.assertIn("待迁移（pending）", md)
        self.assertIn("`version`", md)
        self.assertIn("`pull`", md)


# 已确认、待负责人裁决/修复的真差异（不写进 approved_deltas 白名单；07 切片复核）。
# 方向：旧 1.5.6 接受 --json（agent main 解析前预剥离，九动词全走通用信封分支），
# 新 manifest 这四个动词登记 supports_json=False。键：(entry, verb, field)。
# 出现此集合之外的 failed 行即判测试失败。
KNOWN_REAL_GAPS = {
    ("agent", "assert", "json"),
    ("agent", "metrics", "json"),
    ("agent", "patch", "json"),
    ("agent", "style", "json"),
}


class RealBaselineTest(unittest.TestCase):
    """真实旧 vs 新对比；基线标签/仓库不可用时跳过。"""

    @classmethod
    def setUpClass(cls):
        cls.repo = recorder.discover_repo()
        cls.available = False
        if cls.repo and BASELINE_DIR.is_dir():
            try:
                cls.old_scripts, _ = recorder.resolve_baseline_dir(None, cls.repo,
                                                                   recorder.DEFAULT_TAG)
                cls.available = True
            except RuntimeError:
                cls.available = False

    def test_ledger_recorded_and_version_present(self):
        if not self.available:
            self.skipTest("基线仓库/标签不可用")
        ledger = json.loads((BASELINE_DIR / "ledger.json").read_text(encoding="utf-8"))
        verbs = {v for e in ledger["entries"].values() for v in e["verbs"]}
        self.assertIn("version", verbs)
        self.assertIn("sync.version",
                      [s.get("action") for s in ledger["scenarios"] if s.get("action")])

    def test_real_version_scenario_matches(self):
        if not self.available:
            self.skipTest("基线仓库/标签不可用")
        ledger = json.loads((BASELINE_DIR / "ledger.json").read_text(encoding="utf-8"))
        deltas = json.loads((BASELINE_DIR / "approved_deltas.json").read_text(encoding="utf-8"))
        from toolkit import commands as manifest
        report = harness.run_harness(old_scripts=self.old_scripts, ledger=ledger,
                                     deltas=deltas, specs=manifest.SPECS)
        version_rows = [r for r in report.rows if r["verb"] == "version"]
        self.assertTrue(version_rows, "至少要有一条 sync.version 的真实对比记录")
        self.assertTrue(all(r["status"] == harness.STATUS_MATCH for r in version_rows),
                        f"version 行必须等价：{version_rows}")
        failed_keys = {(r["entry"], r["verb"], r["field"]) for r in report.failed}
        unknown = failed_keys - KNOWN_REAL_GAPS
        self.assertFalse(unknown, f"白名单/待办之外的未登记真差异：{sorted(unknown)}")


class UpgradeInstallTest(unittest.TestCase):
    def setUp(self):
        # 备份根重定向到本用例临时目录：不写真实 LOCALAPPDATA，跨用例不串批次。
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.backup_base = Path(self._tmp.name) / "user-backups" / upgrade.BACKUP_GROUP
        self._base_patch = patch.object(upgrade, "_user_backup_base", lambda: self.backup_base)
        self._base_patch.start()
        self.addCleanup(self._base_patch.stop)

    def _make_source(self, root: Path) -> None:
        (root / "scripts" / "toolkit").mkdir(parents=True)
        (root / "scripts" / "toolkit" / "commands.py").write_text("NEW", encoding="utf-8")
        (root / "references").mkdir(parents=True)
        (root / "references" / "sync-commands.md").write_text("# 新文档", encoding="utf-8")
        (root / "VERSION").write_text("1.6.0", encoding="utf-8")
        (root / "SKILL.md").write_text("# skill", encoding="utf-8")
        # 预检要求的关键入口骨架（桩内空文件即可；真实内容由实跑覆盖）
        (root / "requirements.txt").write_text("", encoding="utf-8")
        (root / "LICENSE").write_text("MIT", encoding="utf-8")
        for entry in upgrade.SMOKE_ENTRY_SCRIPTS:
            f = root / entry
            f.parent.mkdir(parents=True, exist_ok=True)
            f.write_text("# 入口桩", encoding="utf-8")

    def _make_old_target(self, root: Path) -> None:
        (root / "scripts" / "toolkit").mkdir(parents=True)
        (root / "scripts" / "toolkit" / "commands.py").write_text("OLD", encoding="utf-8")
        # 受保护：用户源码与配置，升级后必须原样保留
        (root / "src" / "projects" / "示例表A").mkdir(parents=True)
        (root / "src" / "projects" / "示例表A" / "示例模块.bas").write_text("ATTR X", encoding="utf-8")
        (root / "mapping.json").write_text('{"projects": []}', encoding="utf-8")

    def test_preview_does_not_write(self):
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src); self._make_old_target(dst)
            plan = upgrade.build_plan(src, dst)
            report = upgrade.execute(plan, apply=False)
            self.assertFalse(report["applied"])
            self.assertEqual((dst / "scripts" / "toolkit" / "commands.py").read_text(encoding="utf-8"), "OLD")
            self.assertTrue((dst / "src" / "projects" / "示例表A" / "示例模块.bas").exists())

    def test_apply_backs_up_and_copies_keeps_protected(self):
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src); self._make_old_target(dst)
            plan = upgrade.build_plan(src, dst)
            report = upgrade.execute(plan, apply=True)
            self.assertTrue(report["applied"])
            self.assertEqual((dst / "scripts" / "toolkit" / "commands.py").read_text(encoding="utf-8"), "NEW")
            # 旧文件已备份（用户级外置目录，不进安装树、不进 skills 发现根）
            backup = Path(report["backup_root"])
            self.assertTrue(backup.is_relative_to(self.backup_base))
            self.assertFalse(backup.is_relative_to(dst))
            self.assertEqual((backup / "scripts" / "toolkit" / "commands.py").read_text(encoding="utf-8"), "OLD")
            # 清单存在且覆盖全部交付文件、哈希与磁盘一致
            manifest_path = dst / upgrade.MANIFEST_NAME
            self.assertTrue(manifest_path.is_file())
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(manifest["version"], "1.6.0")
            self.assertTrue(manifest["source"])
            self.assertTrue(manifest["installed_at"])
            delivered = {str(i["rel"]).replace("\\", "/") for i in plan.copies}
            self.assertEqual(set(manifest["files"]), delivered)
            for rel, digest in manifest["files"].items():
                self.assertEqual(digest, upgrade._sha256(dst / rel))
            # 受保护项未动
            self.assertEqual((dst / "mapping.json").read_text(encoding="utf-8"), '{"projects": []}')
            self.assertEqual((dst / "src" / "projects" / "示例表A" / "示例模块.bas").read_text(encoding="utf-8"), "ATTR X")
            self.assertEqual(report["self_check"]["mismatches"], [])
            self.assertEqual(report["self_check"]["protected_touched"], [])

    def test_conflict_copy_excluded_from_plan(self):
        # 同步盘冲突副本不进交付清单；正常文件仍在，预览不落盘。
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src)
            (src / "references" / "示例册_Conflict.md").write_text("# 冲突副本", encoding="utf-8")
            (src / "references" / "sync-commands (1).md").write_text("# 编号副本", encoding="utf-8")
            plan = upgrade.build_plan(src, dst)
            rels = {item["target"].relative_to(dst) for item in plan.copies}
            self.assertNotIn(Path("references") / "示例册_Conflict.md", rels)
            self.assertNotIn(Path("references") / "sync-commands (1).md", rels)
            self.assertIn(Path("references") / "sync-commands.md", rels)
            report = upgrade.execute(plan, apply=False)
            self.assertFalse(report["applied"])
            self.assertEqual([p for p in dst.rglob("*") if p.is_file()], [])

    def test_protected_target_path_rejected(self):
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src)
            (dst / "src").mkdir()
            plan = upgrade.build_plan(src, dst)
            plan.copies.append({"source": src / "VERSION",
                                 "target": dst / "src" / "VERSION"})
            with self.assertRaises(upgrade.UpgradeError):
                plan.validate()

    # ---- I2：孤儿文件清理 + 受保护项保留 ----

    def test_orphan_listed_in_preview_and_removed_on_apply(self):
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src)
            # 旧版本遗留：已删模块与旧 reference，新清单都不含
            (dst / "scripts" / "toolkit").mkdir(parents=True)
            (dst / "scripts" / "toolkit" / "legacy_old.py").write_text("DEAD", encoding="utf-8")
            (dst / "references").mkdir(parents=True)
            (dst / "references" / "retired.md").write_text("# 旧册", encoding="utf-8")
            # 受保护：业务表与配置，升级后必须原样保留
            (dst / ".lab").mkdir(parents=True)
            (dst / ".lab" / "示例表A.xlsm").write_bytes(b"PK\x03\x04")
            (dst / "mapping.json").write_text('{"projects": []}', encoding="utf-8")

            plan = upgrade.build_plan(src, dst)
            orphan_rels = {item["rel"] for item in plan.orphans}
            self.assertIn(Path("scripts") / "toolkit" / "legacy_old.py", orphan_rels)
            self.assertIn(Path("references") / "retired.md", orphan_rels)
            self.assertNotIn(Path(".lab") / "示例表A.xlsm", orphan_rels)
            self.assertNotIn(Path("mapping.json"), orphan_rels)
            preview = upgrade.execute(plan, apply=False)
            self.assertEqual(preview["orphans"], 2)
            # 预览不落盘
            self.assertTrue((dst / "scripts" / "toolkit" / "legacy_old.py").exists())

            report = upgrade.execute(plan, apply=True)
            self.assertEqual(report["orphans"], 2)
            self.assertFalse((dst / "scripts" / "toolkit" / "legacy_old.py").exists())
            self.assertFalse((dst / "references" / "retired.md").exists())
            # 受保护项仍在
            self.assertTrue((dst / ".lab" / "示例表A.xlsm").exists())
            self.assertEqual((dst / "mapping.json").read_text(encoding="utf-8"), '{"projects": []}')
            self.assertEqual(report["self_check"]["mismatches"], [])

    # ---- I3：中途失败自动回滚 ----

    def test_mid_failure_rolls_back_to_pre_upgrade_state(self):
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src); self._make_old_target(dst)
            plan = upgrade.build_plan(src, dst)
            before = {item["rel"]: item["target"].read_bytes()
                      for item in plan.copies if item["target"].exists()}

            real_copy2 = shutil.copy2
            # 第二个文件复制时抛错，模拟磁盘占用/只读
            def flaky_copy2(a, b):
                if getattr(flaky_copy2, "count", 0) == 1:
                    raise OSError("模拟复制失败：磁盘被占用")
                flaky_copy2.count += 1
                return real_copy2(a, b)
            flaky_copy2.count = 0

            original = shutil.copy2
            shutil.copy2 = flaky_copy2
            try:
                with self.assertRaises(upgrade.UpgradeError) as ctx:
                    upgrade.execute(plan, apply=True)
            finally:
                shutil.copy2 = original
            self.assertIn("已回滚", str(ctx.exception))

            # 目标恢复升级前状态：已备份文件内容仍是 OLD
            for rel, data in before.items():
                self.assertEqual((dst / rel).read_bytes(), data, f"回滚后 {rel} 内容不一致")
            self.assertEqual((dst / "scripts" / "toolkit" / "commands.py").read_text(encoding="utf-8"), "OLD")
            # 受保护项未动
            self.assertEqual((dst / "mapping.json").read_text(encoding="utf-8"), '{"projects": []}')

    # ---- 冒烟自检：关键文件缺失 → 非零退出 + 明确报缺失项 ----

    def test_smoke_check_reports_missing_key_file(self):
        # 源包缺已交付的关键入口：预检拦下并点名缺失项，不带病进入备份/替换流程。
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src); self._make_old_target(dst)
            (src / "scripts" / "agent.py").unlink()
            plan = upgrade.build_plan(src, dst)
            with self.assertRaises(upgrade.UpgradeError) as ctx:
                upgrade.execute(plan, apply=True)
            self.assertIn("scripts/agent.py", str(ctx.exception))

    def test_smoke_check_blocks_when_source_key_file_absent(self):
        # 源包缺根级关键文件：预检拦下并点名缺失项；旧目标未被写盘。
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src); self._make_old_target(dst)
            (src / "VERSION").unlink()
            plan = upgrade.build_plan(src, dst)
            with self.assertRaises(upgrade.UpgradeError) as ctx:
                upgrade.execute(plan, apply=True)
            self.assertIn("VERSION", str(ctx.exception))
            # 预检在备份/替换之前拦下：旧目标未被写盘
            self.assertTrue((dst / "mapping.json").exists())

    def test_idempotent_second_apply_no_backup_no_orphan(self):
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src); self._make_old_target(dst)
            first = upgrade.execute(upgrade.build_plan(src, dst), apply=True)
            self.assertFalse(first["identical"])
            second = upgrade.execute(upgrade.build_plan(src, dst), apply=True)
            self.assertTrue(second["identical"])
            self.assertEqual(second["copies"], 0)
            self.assertEqual(second["backups"], 0)
            self.assertEqual(second["orphans"], 0)
            # 幂等不新增安装树内备份；外置备份仍只在用户级目录、不进 skill 树
            self.assertFalse((dst / ".backup" / "upgrade-").exists())
            batches = [p for p in self.backup_base.glob("upgrade-*") if p.is_dir()]
            self.assertEqual(len(batches), 1)

    # ---- 1.6.2 问题⑦：备份落点 + 旧备份只读检测 ----

    def test_backup_root_never_inside_skills_root(self):
        """构造 <tmp>\\skills\\vba-workbench：备份根必须落在 skills 发现根之外，且落点稳定。"""
        with tempfile.TemporaryDirectory() as raw:
            skills = Path(raw) / "skills"
            target = skills / "vba-workbench"
            target.mkdir(parents=True)
            # 临时停掉用例补丁，验证真实默认规则落到用户级目录
            self._base_patch.stop()
            try:
                root = upgrade.external_backup_root(target, "20261122-000000")
                self.assertFalse(root.is_relative_to(skills))
                self.assertFalse(root.is_relative_to(target))
                # 同一 target + 时间戳落点稳定，可回滚定位
                self.assertEqual(root, upgrade.external_backup_root(target, "20261122-000000"))
                self.assertIn(upgrade.BACKUP_APP_DIR, root.parts)
            finally:
                self._base_patch.start()

    def test_legacy_whole_backups_listed_readonly_not_deleted(self):
        """vba-workbench.bak-* 与旧升级批次只读列出；apply 后不删除、不新建同类目录。"""
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src); self._make_old_target(dst)
            whole_in_target = dst / f"{upgrade.WHOLE_BACKUP_PREFIX}20260101-000000"
            whole_in_parent = dst.parent / f"{upgrade.WHOLE_BACKUP_PREFIX}20260102-000000"
            whole_in_target.mkdir(); whole_in_parent.mkdir()
            legacy_in_tree = dst / ".backup" / "upgrade-20260101-000000"
            legacy_in_tree.mkdir(parents=True)
            legacy_sibling = (dst.parent / upgrade.BACKUP_DIR_NAME / upgrade.BACKUP_GROUP
                              / "upgrade-20260102-000000")
            legacy_sibling.mkdir(parents=True)

            plan = upgrade.build_plan(src, dst)
            preview = upgrade.execute(plan, apply=False)
            self.assertEqual(len(preview["whole_backups"]), 2)
            self.assertTrue(any("upgrade-20260101" in p for p in preview["legacy_backups"]))
            self.assertTrue(any("upgrade-20260102" in p for p in preview["legacy_sibling_backups"]))
            self.assertTrue(preview["cleanup_advice"])

            report = upgrade.execute(plan, apply=True)
            # 旧整包备份原样保留：未删除、未新建同类目录
            self.assertTrue(whole_in_target.is_dir())
            self.assertTrue(whole_in_parent.is_dir())
            whole_after = list(dst.glob(f"{upgrade.WHOLE_BACKUP_PREFIX}*")) + \
                list(dst.parent.glob(f"{upgrade.WHOLE_BACKUP_PREFIX}*"))
            self.assertEqual(len(whole_after), 2)
            self.assertEqual(len(report["whole_backups"]), 2)
            self.assertTrue(report["cleanup_advice"])

    def test_rollback_restores_from_new_backup_root(self):
        """rollback 缺省从新用户级备份根取最新批次还原。"""
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src); self._make_old_target(dst)
            plan = upgrade.build_plan(src, dst)
            upgrade.execute(plan, apply=True)
            (dst / "scripts" / "toolkit" / "commands.py").write_text("BROKEN", encoding="utf-8")
            result = upgrade.rollback(plan)
            self.assertTrue(Path(result["backup_root"]).is_relative_to(self.backup_base))
            self.assertEqual(
                (dst / "scripts" / "toolkit" / "commands.py").read_text(encoding="utf-8"), "OLD")

    def test_rollback_falls_back_to_legacy_location(self):
        """新用户级根无批次时，回退树内旧位置 .backup/upgrade-* 只读还原。"""
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src); self._make_old_target(dst)
            plan = upgrade.build_plan(src, dst)
            legacy = dst / ".backup" / "upgrade-20260101-000000"
            (legacy / "scripts" / "toolkit").mkdir(parents=True)
            (legacy / "scripts" / "toolkit" / "commands.py").write_text("OLD", encoding="utf-8")
            result = upgrade.rollback(plan)
            self.assertIn(".backup", result["backup_root"])
            self.assertEqual(
                (dst / "scripts" / "toolkit" / "commands.py").read_text(encoding="utf-8"), "OLD")

    # ---- P19：安装清单 .installed.json ----

    def test_manifest_written_and_updated_not_appended(self):
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src); self._make_old_target(dst)
            upgrade.execute(upgrade.build_plan(src, dst), apply=True)
            manifest_path = dst / upgrade.MANIFEST_NAME
            first = json.loads(manifest_path.read_text(encoding="utf-8"))
            first_at = first["installed_at"]
            # 第二次幂等 apply：清单被覆盖更新，不追加成多份文件
            upgrade.execute(upgrade.build_plan(src, dst), apply=True)
            manifests = [p for p in dst.rglob(upgrade.MANIFEST_NAME)]
            self.assertEqual(len(manifests), 1)
            second = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(set(second["files"]), set(first["files"]))
            self.assertNotEqual(second["installed_at"], "")
            self.assertEqual(second["version"], first["version"])

    def test_manifest_detects_tampered_file(self):
        # 篡改一个交付文件后重跑计划的清单校验，必须报哈希不符。
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src); self._make_old_target(dst)
            upgrade.execute(upgrade.build_plan(src, dst), apply=True)
            (dst / "scripts" / "toolkit" / "commands.py").write_text("TAMPERED", encoding="utf-8")
            plan = upgrade.build_plan(src, dst)
            result = upgrade.verify_manifest(plan)
            self.assertIn("scripts/toolkit/commands.py", result["hash_mismatch"])

    def test_external_backup_restores_overwritten_file(self):
        # 从外置备份批次把被覆盖文件还原回旧内容。
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src); self._make_old_target(dst)
            plan = upgrade.build_plan(src, dst)
            report = upgrade.execute(plan, apply=True)
            self.assertEqual((dst / "scripts" / "toolkit" / "commands.py").read_text(encoding="utf-8"), "NEW")
            rolled = upgrade.rollback(plan, Path(report["backup_root"]))
            self.assertGreaterEqual(rolled["restored"], 1)
            self.assertEqual((dst / "scripts" / "toolkit" / "commands.py").read_text(encoding="utf-8"), "OLD")

    def test_legacy_backup_preserved_not_deleted(self):
        # 旧位置残留的 upgrade-* 批次只读保留，不被当孤儿删除。
        with tempfile.TemporaryDirectory() as raw:
            src, dst = Path(raw) / "src", Path(raw) / "dst"
            src.mkdir(); dst.mkdir()
            self._make_source(src); self._make_old_target(dst)
            legacy = dst / ".backup" / "upgrade-20260101-000000"
            (legacy / "scripts" / "toolkit").mkdir(parents=True)
            (legacy / "scripts" / "toolkit" / "commands.py").write_text("LEGACY", encoding="utf-8")
            plan = upgrade.build_plan(src, dst)
            report = upgrade.execute(plan, apply=True)
            self.assertTrue((legacy / "scripts" / "toolkit" / "commands.py").exists())
            self.assertIn(str(legacy), report["legacy_backups"])


if __name__ == "__main__":
    unittest.main()
