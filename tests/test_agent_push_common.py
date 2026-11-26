# -*- coding: utf-8 -*-
"""本泳道新增契约：agent push 子集写回 / 未知模块四入口告警 / 护栏白名单。

不起 Excel、不写真业务表；COM 与 toolkit 原语全部打桩，临时目录只放 .bas 测试件。
"""
from __future__ import annotations

import contextlib
import io
import json
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))
sys.path.insert(0, str(PKG / "tests"))

from toolkit.base import (  # noqa: E402
    ROOT, cached_warnings, read_source_text, reset_warnings, set_json_mode)
from toolkit.commands import specs_for  # noqa: E402
from toolkit.context import AppContext, Progress  # noqa: E402
from toolkit.registry import build_parser  # noqa: E402
from toolkit.render import render  # noqa: E402
import features.agent_io as agent_io  # noqa: E402
import features.common_subset as common_subset  # noqa: E402
import features.core as core  # noqa: E402
import features.run as run_slice  # noqa: E402
import features.sync_push as sync_push  # noqa: E402
import features.sync_read as sync_read  # noqa: E402


BAS_TEMPLATE = (
    'Attribute VB_Name = "{name}"\n'
    "Option Explicit\n"
    "\n"
    "Public Sub Demo()\n"
    "End Sub\n"
)


def make_tree(d: Path, *, common: list[str], proj: list[str], common_modules):
    """落公共库 + 项目源码测试件；返回 (defaults, 项目配置)。"""
    common_dir = d / "src" / "common"
    proj_src = d / "src" / "projects" / "示例表A"
    common_dir.mkdir(parents=True, exist_ok=True)
    proj_src.mkdir(parents=True, exist_ok=True)
    for stem in common:
        (common_dir / f"{stem}.bas").write_text(
            BAS_TEMPLATE.format(name=stem), encoding="utf-8", newline="\n")
    for stem in proj:
        (proj_src / f"{stem}.bas").write_text(
            BAS_TEMPLATE.format(name=stem), encoding="utf-8", newline="\n")
    defaults = {"common_dir": str(common_dir), "backup_dir": str(d / ".backup")}
    proj_cfg = {"name": "示例表A", "src": str(proj_src),
                "common_modules": common_modules}
    return defaults, proj_cfg


class _FakeCodeModule:
    def __init__(self, text: str):
        self._text = text

    @property
    def CountOfLines(self) -> int:
        return len(self._text.splitlines())

    def Lines(self, start: int, count: int) -> str:
        return self._text


class _FakeComp:
    def __init__(self, name: str, text: str = ""):
        self.Name = name
        self.Type = 1  # vbext_ct_StdModule
        self.CodeModule = _FakeCodeModule(text)


class _FakeWB:
    def __init__(self, tmp: Path):
        self.VBProject = type("P", (), {"VBComponents": []})()
        self.Application = object()
        self.FullName = str(tmp / "示例表A V1.xlsm")
        self.Path = str(tmp)


class AgentPushSubsetTest(unittest.TestCase):
    """agent push：写回范围跟着 common_modules 子集收窄（预检/计数/写回同口径）。"""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

    def test_scope_count_follows_common_modules(self):
        defaults, proj = make_tree(self.tmp, common=["公共", "未用"],
                                   proj=["模块A"], common_modules=["公共"])
        # 项目源码全量 + 公共库子集；子集外的「未用」不计入
        self.assertEqual(agent_io._scope_module_count(proj, defaults), 2)

    def test_preflight_skips_modules_outside_subset(self):
        defaults, proj = make_tree(self.tmp, common=["公共", "未用"],
                                   proj=["模块A"], common_modules=["公共"])
        # 子集外的公共模块写成空正文：仍不应被预检判为致命项
        (self.tmp / "src" / "common" / "未用.bas").write_text("", encoding="utf-8")
        self.assertEqual(agent_io._preflight(proj, defaults, False), [])

    def test_writeback_writes_only_subset(self):
        defaults, proj = make_tree(self.tmp, common=["公共", "未用"],
                                   proj=["模块A"], common_modules=["公共"])
        wb = _FakeWB(self.tmp)
        comps: dict[str, _FakeComp] = {}
        written: dict[str, Path] = {}

        def find_comp(vb, name):
            return comps.get(name)

        def imp(vb, path, name, retries=2):
            written[name] = path
            comps[name] = _FakeComp(name, read_source_text(path))
            return comps[name]

        with mock.patch.object(agent_io, "component_snapshot", lambda vb: {}), \
             mock.patch.object(agent_io, "find_component", find_comp), \
             mock.patch.object(agent_io, "remove_module",
                               lambda vb, name, retries=3: True), \
             mock.patch.object(agent_io, "gbk_import_copy", lambda f, tmp: f), \
             mock.patch.object(agent_io, "replace_module_code",
                               lambda comp, code: None), \
             mock.patch.object(agent_io, "import_component", imp):
            bad, written_map = agent_io._writeback(wb, proj, defaults, False)

        self.assertEqual(bad, [])
        self.assertEqual(set(written_map), {"模块A", "公共"})
        self.assertNotIn("未用", written_map)

    def test_rollback_restore_failure_is_reported(self):
        """--rollback-on-fail 下复原失败的模块必须进 bad，不能静默吞掉。"""
        defaults, proj = make_tree(self.tmp, common=["公共工具"], proj=["模块A"],
                                   common_modules=["公共工具"])
        old_comp = _FakeComp("模块A", "Old body")
        wb = _FakeWB(self.tmp)
        wb.VBProject.VBComponents = [old_comp]
        comps = {"模块A": old_comp}
        written: dict[str, Path] = {}

        def find_comp(vb, name):
            return comps.get(name)

        def imp(vb, path, name, retries=2):
            written[name] = path
            # 回读正文置空 → 制造 postcheck 失败，进入回滚分支
            comps[name] = _FakeComp(name, "")
            return comps[name]

        def fail_restore(comp, code):
            # 只有回滚复原写回前正文时才失败；初次写回改正文必须放行
            if code == "Old body":
                raise RuntimeError("restore boom")

        with mock.patch.object(agent_io, "component_snapshot", lambda vb: {}), \
             mock.patch.object(agent_io, "find_component", find_comp), \
             mock.patch.object(agent_io, "remove_module",
                               lambda vb, name, retries=3: True), \
             mock.patch.object(agent_io, "gbk_import_copy", lambda f, tmp: f), \
             mock.patch.object(agent_io, "replace_module_code", fail_restore), \
             mock.patch.object(agent_io, "import_component", imp):
            bad, _written_map = agent_io._writeback(
                wb, proj, defaults, False, rollback_on_fail=True)

        joined = "；".join(bad)
        self.assertIn("回滚未完成", joined)
        self.assertIn("模块A", joined)
        self.assertIn("restore boom", joined)


class UnknownModuleWarningTest(unittest.TestCase):
    """mapping 写错公共库模块名：list/pull/push/run（含 agent push）同一句告警、只刷一遍。"""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        set_json_mode(True)
        reset_warnings()
        common_subset.clear_unknown_warning_cache()

    def tearDown(self):
        set_json_mode(False)
        reset_warnings()
        common_subset.clear_unknown_warning_cache()
        self._tmp.cleanup()

    def test_four_entries_warn_once_with_same_sentence(self):
        defaults, proj = make_tree(self.tmp, common=["公共"], proj=["模块A"],
                                   common_modules=["不存在的模块"])
        _subset, _reason, ignored = core._resolve_common_subset(proj, defaults)
        self.assertEqual(ignored, ["不存在的模块"])

        # run / push / read / agent 入口各自再解析一遍：同一句话只进一次缓存
        run_slice._common_subset(proj, defaults)
        sync_push._common_subset(proj, defaults)
        sync_read._resolve_common_subset(proj, defaults)
        agent_io._common_subset(proj, defaults)

        warnings = cached_warnings()
        self.assertEqual(len(warnings), 1)
        self.assertTrue(
            warnings[0].startswith("common_modules 里这些模块在公共库不存在，已忽略："))
        self.assertIn("不存在的模块", warnings[0])

    def test_known_modules_do_not_warn(self):
        defaults, proj = make_tree(self.tmp, common=["公共"], proj=["模块A"],
                                   common_modules=["公共"])
        core._resolve_common_subset(proj, defaults)
        sync_push._common_subset(proj, defaults)
        sync_read._resolve_common_subset(proj, defaults)
        self.assertEqual(cached_warnings(), [])

    def test_empty_ignored_list_is_noop(self):
        defaults, proj = make_tree(self.tmp, common=["公共"], proj=["模块A"],
                                   common_modules=None)
        common_subset.warn_unknown_modules([])
        self.assertEqual(cached_warnings(), [])


class DryScopePreviewTest(unittest.TestCase):
    """dry 预览：pull/push/run 都要暴露公共模块子集范围，未知模块告警在预览就可见且只报一遍。"""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        set_json_mode(True)
        reset_warnings()
        common_subset.clear_unknown_warning_cache()

    def tearDown(self):
        set_json_mode(False)
        reset_warnings()
        common_subset.clear_unknown_warning_cache()
        self._tmp.cleanup()

    def _cfg(self, common_modules):
        defaults, proj = make_tree(self.tmp, common=["公共"], proj=["模块A"],
                                   common_modules=common_modules)
        xlsm = self.tmp / "示例表A V1.xlsm"
        xlsm.write_bytes(b"PK\x03\x04")
        proj["file"] = str(xlsm)
        return {"defaults": defaults, "projects": [proj]}, proj

    def _ctx(self, argv: list[str], cfg: dict) -> AppContext:
        parser = build_parser(specs_for("sync"), prog="sync.py")
        args = parser.parse_args(argv)
        return AppContext(
            tool="sync", args=args, config=cfg, root=ROOT, json_mode=True,
            started=time.perf_counter(), progress=Progress(tool="sync", enabled=False))

    def _render_json(self, outcome) -> dict:
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            render(outcome, json_mode=True)
        return json.loads(buf.getvalue().strip())

    def test_push_dry_exposes_subset_and_warning(self):
        cfg, _proj = self._cfg(["公共", "不存在的模块"])
        outcome = sync_push.push(self._ctx(["push", "--project", "示例表A"], cfg))
        self.assertEqual(outcome.exit_code, 0)
        self.assertEqual(outcome.stats["common_scope"], "subset")
        self.assertEqual(outcome.stats["common_modules"], ["公共"])
        payload = self._render_json(outcome)
        stats = payload["stats"]
        self.assertEqual(stats["common_modules"], ["公共"])
        unknown = [w for w in stats.get("warnings", []) if "不存在的模块" in w]
        self.assertEqual(len(unknown), 1)
        self.assertIn("不存在的模块", unknown[0])
        self.assertTrue(any("公共" in h for h in outcome.human))

    @mock.patch.object(sync_read, "find_excel_with_workbook", return_value=(None, None))
    def test_pull_dry_exposes_subset_and_warning(self, _patched):
        cfg, _proj = self._cfg(["公共", "不存在的模块"])
        outcome = sync_read.pull(self._ctx(["pull", "--all-project"], cfg))
        row = outcome.stats["results"][0]
        self.assertEqual(row["common_scope"], "subset")
        self.assertEqual(row["common_modules"], ["公共"])
        self.assertEqual(len(cached_warnings()), 1)
        self.assertIn("不存在的模块", cached_warnings()[0])

    def test_run_dry_exposes_subset_and_warning(self):
        cfg, _proj = self._cfg(["公共", "不存在的模块"])
        outcome = run_slice.run(
            self._ctx(["run", "--project", "示例表A", "--macro", "示例求和"], cfg))
        self.assertEqual(outcome.stats["common_scope"], "subset")
        self.assertEqual(outcome.stats["common_modules"], ["公共"])
        self.assertEqual(len(cached_warnings()), 1)
        self.assertIn("不存在的模块", cached_warnings()[0])

    def test_unconfigured_dry_scope_is_all_without_warning(self):
        cfg, _proj = self._cfg(None)
        outcome = sync_push.push(self._ctx(["push", "--project", "示例表A"], cfg))
        self.assertEqual(outcome.stats["common_scope"], "all")
        self.assertEqual(outcome.stats["common_modules"], [])
        self.assertEqual(cached_warnings(), [])


class GuardWhitelistTest(unittest.TestCase):
    """护栏白名单：common_subset 各引用对已登记，真实源码扫描零违规。"""

    def test_common_subset_pairs_are_whitelisted(self):
        import test_import_layers as guard
        for pair in (("core", "common_subset"), ("run", "common_subset"),
                     ("sync_push", "common_subset"), ("sync_read", "common_subset"),
                     ("agent_io", "common_subset")):
            self.assertIn(pair, guard.FEATURE_WHITELIST)
        self.assertEqual(guard.scan_features(guard.SCRIPTS), [])


if __name__ == "__main__":
    unittest.main()
