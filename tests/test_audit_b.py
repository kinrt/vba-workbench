# -*- coding: utf-8 -*-
"""audit_b 切片测试：structure / indicators / overview 三个只读动词。

离线夹具：tempfile + mapping.json + 源码文本；不启动 Excel、不碰业务表。
四类覆盖：成功 / 业务失败（exit 1 且 errors[].code）/ 用法错（exit 2）/ --json 单行纯度；
另设专例证明 exit 3（仅警告）语义保留，不归一成 1。
"""
from __future__ import annotations

import contextlib
import io
import json
import sys
import tempfile
import unittest
from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))

from features import audit_b  # noqa: E402
from toolkit.commands import specs_for  # noqa: E402
from toolkit.contract import ErrorCode  # noqa: E402
from toolkit.registry import build_parser  # noqa: E402
from toolkit.render import render  # noqa: E402


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

DIRTY_FILES = {
    "src/common/Alpha.bas": (
        "Option Explicit\n"
        "Public Sub 没注释过程()\n"
        "End Sub\n"
    ),
    "src/common/重复.bas": (
        "Option Explicit\n"
        "' 公共侧说明\n"
        "Public Sub 重复过程()\n"
        "End Sub\n"
    ),
    "src/projects/甲表/幽灵.bas": "' 真实文件在：Ghost\n",
    "src/projects/甲表/重复.bas": (
        "Option Explicit\n"
        "Public Sub 重复过程()\n"
        "End Sub\n"
    ),
}

# H01：Call 了一个全仓都没定义的裸名 → 警告（非 strict 退出 3）
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

# shadow 阻断：项目侧不同名文件里定义了与公共库同名的 Public 过程
OVERVIEW_FATAL = {
    "src/common/公共.bas": (
        "Attribute VB_Name = \"公共\"\n"
        "Option Explicit\n"
        "' 公共过程说明\n"
        "Public Sub 公共过程()\n"
        "End Sub\n"
    ),
    "src/projects/甲表/项目侧.bas": (
        "Attribute VB_Name = \"项目侧\"\n"
        "Option Explicit\n"
        "Public Sub 公共过程()\n"
        "End Sub\n"
    ),
}

# 非阻断：孤儿壳；strict 才升级为 1，默认退出 0
OVERVIEW_ORPHAN = {
    "src/projects/甲表/孤儿.bas": "' 真实文件在：Ghost\n",
}


def _make_root(files: dict[str, str]) -> Path:
    d = Path(tempfile.mkdtemp())
    for rel, text in files.items():
        p = d / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(text, encoding="utf-8", newline="\n")
    return d


class _StubProgress:
    @contextmanager
    def stage(self, name: str):
        yield


def _make_ctx(root: Path, cfg: dict, **args):
    ns = SimpleNamespace(
        exclude=[], ignore=[], strict=False, all=False, md=None, out=None,
        project=[], json=False, no_input=False,
    )
    for k, v in args.items():
        setattr(ns, k, v)
    return SimpleNamespace(
        args=ns, config=cfg, root=root, json_mode=False,
        progress=_StubProgress(), elapsed_ms=lambda: 7,
        stage_timings=lambda: {"stage": 3},
    )


def _json_lines(outcome) -> list[str]:
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        code = render(outcome, json_mode=True)
    assert code == outcome.exit_code
    return buf.getvalue().strip().splitlines()


# ---------------------------------------------------------------------------
# H16/H17 运行期探测：假 COM 对象，不启动 Excel、不碰真表
# ---------------------------------------------------------------------------
class _FakeRef:
    def __init__(self, name: str) -> None:
        self.Name = name


class _FakeComp:
    def __init__(self, ctype: int, lines: int, controls: int = 0) -> None:
        self.Type = ctype
        self.CodeModule = SimpleNamespace(CountOfLines=lines)
        self.Designer = SimpleNamespace(Controls=SimpleNamespace(Count=controls))


class _FakeVBProject:
    def __init__(self, refs: list[str], comps: list[_FakeComp]) -> None:
        self.References = [_FakeRef(n) for n in refs]
        self.VBComponents = list(comps)


class _FakeWB:
    def __init__(self, refs: list[str], comps: list[_FakeComp]) -> None:
        self.VBProject = _FakeVBProject(refs, comps)


class _UnreadableWB:
    """工程被锁/未放行：一碰 VBProject 就抛异常。"""
    @property
    def VBProject(self):
        raise RuntimeError("该工程已被保护，不能执行操作")


def _live_cfg(root: Path) -> dict:
    """带可解析表路径的项目配置；脱敏表名 示例表A / V100。"""
    xlsm = root / "示例表A V100.xlsm"
    xlsm.write_bytes(b"")
    return {"projects": [{"name": "示例表A", "file": str(xlsm)}],
            "defaults": {"common_dir": "src/common"}}


# ---------------------------------------------------------------------------
# structure
# ---------------------------------------------------------------------------
class StructureTest(unittest.TestCase):
    def test_clean_exit_0(self):
        root = _make_root(CLEAN_COMMON)
        outcome = audit_b.structure(_make_ctx(root, CFG))
        self.assertEqual(outcome.exit_code, 0)
        self.assertEqual(outcome.errors, [])
        self.assertTrue(outcome.ok)
        self.assertEqual(outcome.stats["issues"], 0)
        self.assertIsInstance(outcome.details, list)
        self.assertTrue(all("group" in r and "count" in r and "items" in r
                            for r in outcome.details))

    def test_warning_only_exit_3(self):
        """只有警告时退出码必须是 3，不能归一成 1。"""
        root = _make_root(DIRTY_FILES)
        outcome = audit_b.structure(_make_ctx(root, CFG))
        self.assertEqual(outcome.exit_code, 3)
        self.assertEqual(outcome.errors, [])
        self.assertTrue(outcome.ok)  # 仅警告：ok=true，退出码 3 不归一为 1
        self.assertGreaterEqual(outcome.stats["issues"], 1)
        self.assertEqual(outcome.stats["warned"], outcome.stats["issues"])

    def test_strict_exit_1_with_code(self):
        root = _make_root(DIRTY_FILES)
        outcome = audit_b.structure(_make_ctx(root, CFG, strict=True))
        self.assertEqual(outcome.exit_code, 1)
        self.assertFalse(outcome.ok)
        self.assertEqual(len(outcome.errors), 1)
        self.assertEqual(outcome.errors[0].code, ErrorCode.SOURCE_PREFLIGHT.value)

    def test_usage_error_exit_2(self):
        specs = specs_for("audit")
        parser = build_parser(specs, prog="audit")
        with self.assertRaises(SystemExit) as cm:
            parser.parse_args(["structure", "--no-such-argument"])
        self.assertEqual(cm.exception.code, 2)

    def test_json_single_line(self):
        root = _make_root(CLEAN_COMMON)
        outcome = audit_b.structure(_make_ctx(root, CFG))
        lines = _json_lines(outcome)
        self.assertEqual(len(lines), 1)
        payload = json.loads(lines[0])
        self.assertEqual(payload["exit_code"], 0)
        self.assertIn("details", payload)
        self.assertIn("stats", payload)
        self.assertIn("errors", payload)

    def test_ignore_group_cleared(self):
        root = _make_root(DIRTY_FILES)
        outcome = audit_b.structure(_make_ctx(root, CFG, ignore=["no_comment"]))
        by_group = {r["group"]: r for r in outcome.details}
        self.assertEqual(by_group["no_comment"]["count"], 0)
        self.assertGreaterEqual(by_group["orphan"]["count"], 1)


# ---------------------------------------------------------------------------
# indicators
# ---------------------------------------------------------------------------
class IndicatorsTest(unittest.TestCase):
    ROW_KEYS = {"id", "name", "state", "count", "detail", "samples", "suggested_command"}

    def test_clean_exit_0(self):
        root = _make_root(CLEAN_COMMON)
        outcome = audit_b.indicators(_make_ctx(root, CFG))
        self.assertEqual(outcome.exit_code, 0)
        self.assertEqual(outcome.errors, [])
        self.assertTrue(outcome.ok)
        self.assertEqual(outcome.stats["warned"], 0)
        self.assertGreaterEqual(outcome.stats["passed"], 1)
        self.assertIsInstance(outcome.details, list)
        self.assertTrue(all(self.ROW_KEYS <= set(r) for r in outcome.details))

    def test_warning_only_exit_3(self):
        """H01 命中警告但未开 --strict：退出码 3，errors 为空（3 不归一为 1）。"""
        root = _make_root(INDICATOR_WARN)
        outcome = audit_b.indicators(_make_ctx(root, CFG))
        self.assertEqual(outcome.exit_code, 3)
        self.assertEqual(outcome.errors, [])
        self.assertTrue(outcome.ok)
        self.assertEqual(outcome.stats["warned"], 1)
        h01 = next(r for r in outcome.details if r["id"] == "H01")
        self.assertEqual(h01["state"], "警告")
        self.assertGreaterEqual(h01["count"], 1)

    def test_strict_exit_1_with_code(self):
        root = _make_root(INDICATOR_WARN)
        outcome = audit_b.indicators(_make_ctx(root, CFG, strict=True))
        self.assertEqual(outcome.exit_code, 1)
        self.assertFalse(outcome.ok)
        self.assertEqual(len(outcome.errors), 1)
        self.assertEqual(outcome.errors[0].code, ErrorCode.SOURCE_PREFLIGHT.value)

    def test_usage_error_exit_2(self):
        specs = specs_for("audit")
        parser = build_parser(specs, prog="audit")
        with self.assertRaises(SystemExit) as cm:
            parser.parse_args(["indicators", "--no-such-argument"])
        self.assertEqual(cm.exception.code, 2)

    def test_json_single_line(self):
        root = _make_root(CLEAN_COMMON)
        outcome = audit_b.indicators(_make_ctx(root, CFG))
        lines = _json_lines(outcome)
        self.assertEqual(len(lines), 1)
        payload = json.loads(lines[0])
        self.assertEqual(payload["exit_code"], 0)
        self.assertIn("details", payload)
        self.assertIn("stats", payload)
        self.assertIn("errors", payload)

    def test_row_ids_present(self):
        root = _make_root(CLEAN_COMMON)
        outcome = audit_b.indicators(_make_ctx(root, CFG))
        ids = [r["id"] for r in outcome.details]
        for expected in ("H01", "H02", "H02b", "H03", "H04", "H05a", "H05b", "H06",
                          "H07", "H08", "H09", "H10", "H11", "H12", "H13",
                          "H14", "H15", "H16", "H17", "H18", "H19"):
            self.assertIn(expected, ids)

    def _live_rows(self, root: Path, cfg: dict, wb=None, *, attach=None, find=None):
        attach = attach or (lambda: object())
        find = find or (lambda xl, path: (wb, []))
        with patch.object(audit_b.excel_gateway, "attach_running_excel", side_effect=attach), \
             patch.object(audit_b.excel_gateway, "find_open_workbook", side_effect=find):
            outcome = audit_b.indicators(_make_ctx(root, cfg))
        return {r["id"]: r for r in outcome.details}

    def test_h16_h17_pass_when_workbook_open(self):
        """已开表 + 有 References + CountOfLines 非空：H16/H17 必须出通过，不能挂未跑。"""
        root = _make_root(CLEAN_COMMON)
        cfg = _live_cfg(root)
        wb = _FakeWB(["Microsoft Forms 2.0 Object Library", "Excel"],
                     [_FakeComp(3, 120), _FakeComp(1, 320)])
        rows = self._live_rows(root, cfg, wb)
        self.assertEqual(rows["H16"]["state"], "通过")
        self.assertEqual(rows["H17"]["state"], "通过")
        self.assertEqual(rows["H16"]["count"], "1/1")
        self.assertEqual(rows["H17"]["count"], "1/1")

    def test_h16_warns_when_form_missing_forms_ref(self):
        """有 UserForm 但引用里没有 Forms：H16 警告，H17 仍通过。"""
        root = _make_root(CLEAN_COMMON)
        cfg = _live_cfg(root)
        wb = _FakeWB(["Excel"], [_FakeComp(3, 80), _FakeComp(1, 200)])
        rows = self._live_rows(root, cfg, wb)
        self.assertEqual(rows["H16"]["state"], "警告")
        self.assertEqual(rows["H17"]["state"], "通过")
        self.assertTrue(any("示例表A" in s and "Microsoft Forms 2.0" in s
                            for s in rows["H16"]["samples"]))

    def test_h16_warns_when_sheet_controls_without_forms_ref(self):
        """表内控件（文档模块带控件）同样要求 Forms 引用；缺则 H16 警告。"""
        root = _make_root(CLEAN_COMMON)
        cfg = _live_cfg(root)
        wb = _FakeWB(["Excel"], [_FakeComp(100, 60, controls=1), _FakeComp(1, 150)])
        rows = self._live_rows(root, cfg, wb)
        self.assertEqual(rows["H16"]["state"], "警告")
        self.assertEqual(rows["H17"]["state"], "通过")

    def test_h16_h17_warn_when_project_locked(self):
        """工程读不到：H16/H17 都记警告并带项目名，不能静默少报成通过。"""
        root = _make_root(CLEAN_COMMON)
        cfg = _live_cfg(root)
        rows = self._live_rows(root, cfg, _UnreadableWB())
        self.assertEqual(rows["H16"]["state"], "警告")
        self.assertEqual(rows["H17"]["state"], "警告")
        self.assertTrue(any("示例表A" in s for s in rows["H17"]["samples"]))

    def test_h16_h17_unrun_when_workbook_closed(self):
        """表没打开：如实标未跑，不能拿别的表顶替、不能反向误判通过。"""
        root = _make_root(CLEAN_COMMON)
        cfg = _live_cfg(root)
        rows = self._live_rows(root, cfg, find=lambda xl, path: (None, [str(path)]))
        self.assertEqual(rows["H16"]["state"], "未跑")
        self.assertEqual(rows["H17"]["state"], "未跑")
        self.assertIn("没有项目的工作簿处于打开状态", rows["H16"]["detail"])

    def test_h16_h17_unrun_when_no_running_excel(self):
        """没有正在运行的 Excel：未跑并带原因，find_open 不应被调用。"""
        root = _make_root(CLEAN_COMMON)
        cfg = _live_cfg(root)

        def _attach():
            raise RuntimeError("NO_EXCEL")

        with patch.object(audit_b.excel_gateway, "attach_running_excel", side_effect=_attach), \
             patch.object(audit_b.excel_gateway, "find_open_workbook") as fake_find:
            outcome = audit_b.indicators(_make_ctx(root, cfg))
        rows = {r["id"]: r for r in outcome.details}
        self.assertEqual(rows["H16"]["state"], "未跑")
        self.assertEqual(rows["H17"]["state"], "未跑")
        self.assertTrue(rows["H16"]["detail"].startswith("没有正在运行的 Excel"))
        fake_find.assert_not_called()


# ---------------------------------------------------------------------------
# overview
# ---------------------------------------------------------------------------
class OverviewTest(unittest.TestCase):
    SUM_KEYS = {"module", "kind", "path", "members", "dead", "test", "single",
                "active", "lines", "ref_files", "infra"}

    def test_clean_exit_0(self):
        root = _make_root(CLEAN_COMMON)
        outcome = audit_b.overview(_make_ctx(root, CFG))
        self.assertEqual(outcome.exit_code, 0)
        self.assertEqual(outcome.errors, [])
        self.assertTrue(outcome.ok)
        self.assertEqual(outcome.stats["modules"], 1)
        self.assertTrue(outcome.stats["common_configured"])
        self.assertEqual(outcome.stats["issues"], 0)
        self.assertEqual(outcome.stats["fatal"], 0)
        self.assertIsInstance(outcome.details, list)
        self.assertTrue(all(self.SUM_KEYS <= set(r) for r in outcome.details))

    def test_fatal_exit_1_with_code(self):
        """shadow 属阻断组：默认就退 1（overview 不因仅提醒退 3）。"""
        root = _make_root(OVERVIEW_FATAL)
        outcome = audit_b.overview(_make_ctx(root, CFG))
        self.assertEqual(outcome.exit_code, 1)
        self.assertFalse(outcome.ok)
        self.assertGreaterEqual(outcome.stats["fatal"], 1)
        self.assertEqual(len(outcome.errors), 1)
        self.assertEqual(outcome.errors[0].code, ErrorCode.SOURCE_PREFLIGHT.value)

    def test_non_fatal_default_exit_0(self):
        """孤儿壳只计数不阻断：默认退出 0。"""
        root = _make_root(OVERVIEW_ORPHAN)
        outcome = audit_b.overview(_make_ctx(root, CFG))
        self.assertEqual(outcome.exit_code, 0)
        self.assertEqual(outcome.errors, [])
        self.assertGreaterEqual(outcome.stats["issues"], 1)
        self.assertEqual(outcome.stats["fatal"], 0)

    def test_strict_non_fatal_upgrades_to_1(self):
        root = _make_root(OVERVIEW_ORPHAN)
        outcome = audit_b.overview(_make_ctx(root, CFG, strict=True))
        self.assertEqual(outcome.exit_code, 1)
        self.assertFalse(outcome.ok)
        self.assertEqual(len(outcome.errors), 1)
        self.assertEqual(outcome.errors[0].code, ErrorCode.SOURCE_PREFLIGHT.value)

    def test_usage_error_exit_2(self):
        specs = specs_for("audit")
        parser = build_parser(specs, prog="audit")
        with self.assertRaises(SystemExit) as cm:
            parser.parse_args(["overview", "--no-such-argument"])
        self.assertEqual(cm.exception.code, 2)

    def test_json_single_line(self):
        root = _make_root(CLEAN_COMMON)
        outcome = audit_b.overview(_make_ctx(root, CFG))
        lines = _json_lines(outcome)
        self.assertEqual(len(lines), 1)
        payload = json.loads(lines[0])
        self.assertEqual(payload["exit_code"], 0)
        self.assertIn("details", payload)
        self.assertIn("stats", payload)
        self.assertIn("errors", payload)

    def test_never_exit_3(self):
        """overview 口径：只有 0 / 1，不出现 3。"""
        for fixture in (CLEAN_COMMON, OVERVIEW_FATAL, OVERVIEW_ORPHAN):
            root = _make_root(fixture)
            for strict in (False, True):
                outcome = audit_b.overview(_make_ctx(root, CFG, strict=strict))
                self.assertIn(outcome.exit_code, (0, 1))


if __name__ == "__main__":
    unittest.main()
