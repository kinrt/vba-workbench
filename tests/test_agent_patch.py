# -*- coding: utf-8 -*-
"""agent_patch 切片契约测试：patch / assert / style / metrics。

不启动 Excel、不跑业务表；会话与工作簿一律用桩对象，测试件只放 tempfile。
参数解析走 ``registry.build_parser(commands.specs_for("agent"))``，不依赖入口壳
（``scripts/agent.py`` 归 06 切片）。
"""
import io
import json
import os
import shlex
import sys
import tempfile
import time
import unittest
from contextlib import redirect_stdout, redirect_stderr
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))

from toolkit.commands import specs_for  # noqa: E402
from toolkit.context import Progress  # noqa: E402
from toolkit.contract import Failure  # noqa: E402
from toolkit.registry import build_parser  # noqa: E402
from toolkit.render import render  # noqa: E402


class Ctx:
    """最小运行上下文：形状对齐 ``toolkit.context.AppContext``，测试可直接塞桩。"""

    def __init__(self, args, config: dict, root: Path, json_mode: bool = False,
                 runner=None):
        self.args = args
        self.config = config
        self.root = root
        self.json_mode = json_mode
        self.started = time.perf_counter()
        self.progress = Progress(tool="agent", enabled=not json_mode)
        self.runner = runner

    def elapsed_ms(self) -> int:
        return int((time.perf_counter() - self.started) * 1000)

    def stage_timings(self) -> dict:
        return dict(self.progress.timings)


def _parser():
    return build_parser(specs_for("agent"), prog="agent.py")


def _config() -> dict:
    return {"defaults": {"common_dir": "src/common"}, "projects": []}


def _clean_basename_file(d: Path, name: str, body: str) -> Path:
    sub = d / "src" / "模块"
    sub.mkdir(parents=True, exist_ok=True)
    p = sub / name
    p.write_text(body, encoding="utf-8", newline="\n")
    return p


CLEAN_BAS = (
    'Attribute VB_Name = "示例"\n'
    "Option Explicit\n\n"
    "Public Sub 示例()\n"
    '    Call MsgBox("x")\n'
    "End Sub\n"
)

DIRTY_BAS = (
    'Attribute VB_Name = "坏样例"\n'
    "Public Sub 坏样例()\n"
    '    MsgBox "x"\n'
    "End Sub\n"
)


class StyleTest(unittest.TestCase):
    def _run(self, d: Path, *argv: str):
        args = _parser().parse_args(["style", *argv])
        ctx = Ctx(args, _config(), d, json_mode=args.json)
        from features.agent_patch import handle_style
        return handle_style(ctx)

    def test_style_success(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _clean_basename_file(d, "示例.bas", CLEAN_BAS)
            outcome = self._run(d, "--action", "check", "--scope", "src")
            self.assertEqual(outcome.exit_code, 0, outcome.human)
            self.assertTrue(outcome.ok)
            self.assertEqual(outcome.errors, [])
            self.assertIn("lines", outcome.stats)

    def test_style_business_failure_has_code(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _clean_basename_file(d, "坏样例.bas", DIRTY_BAS)
            outcome = self._run(d, "--action", "check", "--scope", "src")
            self.assertEqual(outcome.exit_code, 1)
            self.assertFalse(outcome.ok)
            self.assertTrue(outcome.errors)
            self.assertEqual(outcome.errors[0].code, "SOURCE_PREFLIGHT")
            # N5：警告独立走 Outcome.warnings，不再以 [警告] 前缀混进 human 报告行。
            self.assertTrue(any("合计" in w for w in outcome.warnings))
            self.assertFalse(any(line.startswith("[警告]") for line in outcome.human))

    def test_style_fix_apply_with_which(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            f = _clean_basename_file(d, "坏样例.bas", DIRTY_BAS)
            outcome = self._run(d, "--action", "fix", "--which", "call", "--apply")
            self.assertEqual(outcome.exit_code, 0, outcome.human)
            text = f.read_text(encoding="utf-8")
            self.assertIn('Call MsgBox("x")', text)

    def test_style_which_filters_explicit_only(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            f = _clean_basename_file(d, "坏样例.bas", DIRTY_BAS)
            outcome = self._run(d, "--action", "check", "--which", "explicit")
            self.assertEqual(outcome.exit_code, 1)
            self.assertEqual(outcome.errors[0].code, "SOURCE_PREFLIGHT")
            # --which explicit 只做该维度检查，不改文件。
            self.assertEqual(f.read_text(encoding="utf-8"), DIRTY_BAS)

    def test_style_which_eol_only(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            p = d / "src" / "模块" / "示例.bas"
            p.parent.mkdir(parents=True)
            p.write_bytes(CLEAN_BAS.replace("\n", "\r\n").encode("utf-8"))
            outcome = self._run(d, "--action", "check", "--which", "eol")
            self.assertEqual(outcome.exit_code, 1)
            self.assertEqual(outcome.errors[0].code, "SOURCE_PREFLIGHT")
            self.assertEqual(outcome.details[0]["rule"], "eol")
            self.assertIn("示例.bas", outcome.details[0]["file"])
            self.assertIn("CRLF", outcome.details[0]["message"])
            # 同一文件只查 explicit 维度过关，证明 --which 真的只跑指定检查。
            outcome = self._run(d, "--action", "check", "--which", "explicit")
            self.assertEqual(outcome.exit_code, 0, outcome.human)

    def test_style_failure_suggested_keeps_which(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            p = _clean_basename_file(d, "示例.bas", CLEAN_BAS)
            p.write_bytes(CLEAN_BAS.replace("\n", "\r\n").encode("utf-8"))
            outcome = self._run(d, "--action", "check", "--which", "eol")
            self.assertEqual(outcome.exit_code, 1)
            suggested = outcome.errors[0].suggested
            self.assertIn("--which eol", suggested)
            self.assertIn("--apply", suggested)

    def test_style_usage_error_exits_2(self):
        with self.assertRaises(SystemExit) as cm:
            _parser().parse_args(["style", "--action", "not-a-choice"])
        self.assertEqual(cm.exception.code, 2)

    def test_style_json_is_single_line(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _clean_basename_file(d, "示例.bas", CLEAN_BAS)
            args = _parser().parse_args(["style", "--action", "check", "--scope", "src",
                                         "--json"])
            ctx = Ctx(args, _config(), d, json_mode=True)
            from features.agent_patch import handle_style
            outcome = handle_style(ctx)
            buf, err = io.StringIO(), io.StringIO()
            with redirect_stdout(buf), redirect_stderr(err):
                code = render(outcome, json_mode=True)
            lines = buf.getvalue().splitlines()
            self.assertEqual(len(lines), 1, buf.getvalue())
            payload = json.loads(lines[0])
            for key in ("ok", "action", "details", "errors", "elapsed_ms", "ts",
                        "stats", "exit_code"):
                self.assertIn(key, payload)
            self.assertEqual(payload["action"], "agent.style")
            self.assertEqual(payload["exit_code"], code)
            self.assertIn("lines", payload["stats"])
            self.assertTrue(payload["stats"]["lines"])
            self.assertEqual(err.getvalue(), "")


PATCH_HIT = (
    'Attribute VB_Name = "示例"\n'
    "Option Explicit\n\n"
    "Public Sub 示例()\n"
    '    Call MsgBox("x")\n'
    "End Sub\n"
)

PATCH_DUP = (
    'Attribute VB_Name = "双样例"\n'
    "Option Explicit\n\n"
    "Public Sub A()\n"
    '    Call MsgBox("x")\n'
    "End Sub\n\n"
    "Public Sub B()\n"
    '    Call MsgBox("x")\n'
    "End Sub\n"
)


def _write_patch(d: Path, text: str) -> Path:
    p = d / "patch.txt"
    p.write_text(text, encoding="utf-8", newline="\n")
    return p


class PatchTest(unittest.TestCase):
    def _run(self, d: Path, *argv: str):
        args = _parser().parse_args(["patch", *argv])
        ctx = Ctx(args, _config(), d, json_mode=args.json)
        from features.agent_patch import handle_patch
        return handle_patch(ctx)

    def test_patch_dry_run_success(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _clean_basename_file(d, "示例.bas", PATCH_HIT)
            pf = _write_patch(d, "@@ src/模块/示例.bas\n"
                                 '-    Call MsgBox("x")\n'
                                 '+    Call MsgBox("y")\n')
            outcome = self._run(d, "--file", str(pf))
            self.assertEqual(outcome.exit_code, 0, outcome.human)
            self.assertTrue(outcome.ok)
            self.assertTrue(outcome.details[0]["dry"])
            self.assertEqual(outcome.details[0]["changes"], 1)
            self.assertEqual(outcome.stats["changes"], 1)
            # dry-run 不落盘
            self.assertIn('Call MsgBox("x")', (d / "src" / "模块" / "示例.bas").read_text("utf-8"))

    def test_patch_apply_writes_file_and_skips_style(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            f = _clean_basename_file(d, "示例.bas", PATCH_HIT)
            pf = _write_patch(d, "@@ src/模块/示例.bas\n"
                                 '-    Call MsgBox("x")\n'
                                 '+    Call MsgBox("y")\n')
            outcome = self._run(d, "--file", str(pf), "--apply", "--no-style")
            self.assertEqual(outcome.exit_code, 0, outcome.human)
            self.assertTrue(outcome.details[0]["applied"])
            self.assertIn('Call MsgBox("y")', f.read_text(encoding="utf-8"))
            self.assertNotIn('Call MsgBox("x")', f.read_text(encoding="utf-8"))

    def test_patch_format_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _clean_basename_file(d, "示例.bas", PATCH_HIT)
            pf = _write_patch(d, "@@ src/模块/示例.bas\n"
                                 "-    Call MsgBox(\"x\")\n"
                                 "@@ 这一行看不懂\n")
            with self.assertRaises(Failure) as cm:
                self._run(d, "--file", str(pf))
            self.assertEqual(cm.exception.code, "PATCH_FORMAT")
            self.assertEqual(cm.exception.stage, "patch")

    def test_patch_block_ambiguous(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _clean_basename_file(d, "双样例.bas", PATCH_DUP)
            pf = _write_patch(d, "@@ src/模块/双样例.bas\n"
                                 '-    Call MsgBox("x")\n'
                                 '+    Call MsgBox("y")\n')
            with self.assertRaises(Failure) as cm:
                self._run(d, "--file", str(pf))
            self.assertEqual(cm.exception.code, "PATCH_BLOCK_AMBIGUOUS")
            self.assertEqual(cm.exception.stage, "patch")

    def test_patch_file_not_found(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            pf = _write_patch(d, "@@ src/模块/不存在.bas\n-    foo\n+    bar\n")
            with self.assertRaises(Failure) as cm:
                self._run(d, "--file", str(pf))
            self.assertEqual(cm.exception.code, "PATCH_FILE_NOT_FOUND")

    def test_patch_usage_error_exits_2(self):
        with self.assertRaises(SystemExit) as cm:
            _parser().parse_args(["patch"])
        self.assertEqual(cm.exception.code, 2)

    def test_patch_json_is_single_line(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _clean_basename_file(d, "示例.bas", PATCH_HIT)
            pf = _write_patch(d, "@@ src/模块/示例.bas\n"
                                 '-    Call MsgBox("x")\n'
                                 '+    Call MsgBox("y")\n')
            args = _parser().parse_args(["patch", "--file", str(pf), "--json"])
            ctx = Ctx(args, _config(), d, json_mode=True)
            from features.agent_patch import handle_patch
            outcome = handle_patch(ctx)
            buf, err = io.StringIO(), io.StringIO()
            with redirect_stdout(buf), redirect_stderr(err):
                code = render(outcome, json_mode=True)
            lines = buf.getvalue().splitlines()
            self.assertEqual(len(lines), 1, buf.getvalue())
            payload = json.loads(lines[0])
            for key in ("ok", "action", "details", "errors", "elapsed_ms", "ts",
                        "stats", "exit_code"):
                self.assertIn(key, payload)
            self.assertEqual(payload["action"], "agent.patch")
            self.assertEqual(payload["exit_code"], code)
            self.assertIn("lines", payload["stats"])


METRICS_BAS = (
    'Attribute VB_Name = "示例"\n'
    "Option Explicit\n\n"
    "Public Sub 示例()\n"
    '    Call MsgBox("x")\n'
    "End Sub\n"
)


class MetricsTest(unittest.TestCase):
    def _run(self, d: Path, *argv: str):
        args = _parser().parse_args(["metrics", *argv])
        ctx = Ctx(args, _config(), d, json_mode=args.json)
        from features.agent_patch import handle_metrics
        return handle_metrics(ctx)

    def test_metrics_success(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _clean_basename_file(d, "示例.bas", METRICS_BAS)
            outcome = self._run(d)
            self.assertEqual(outcome.exit_code, 0, outcome.human)
            self.assertTrue(outcome.ok)
            self.assertEqual(outcome.errors, [])
            self.assertEqual(outcome.stats["files"], 1)
            self.assertEqual(outcome.stats["procs"], 1)
            self.assertTrue(any("模块 / 过程" in line for line in outcome.human))

    def test_metrics_out_dry_run_not_written(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _clean_basename_file(d, "示例.bas", METRICS_BAS)
            out = d / "out" / "metrics.md"
            outcome = self._run(d, "--out", str(out))
            self.assertEqual(outcome.exit_code, 0)
            self.assertFalse(out.exists())
            self.assertTrue(any("dry-run" in line for line in outcome.human))

    def test_metrics_out_apply_writes_markdown(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _clean_basename_file(d, "示例.bas", METRICS_BAS)
            out = d / "out" / "metrics.md"
            outcome = self._run(d, "--out", str(out), "--apply")
            self.assertEqual(outcome.exit_code, 0, outcome.human)
            self.assertTrue(out.exists())
            text = out.read_text(encoding="utf-8")
            self.assertIn("模块 / 过程", text)
            self.assertIn("1 个源文件 / 1 个方法", text)

    def test_metrics_usage_error_exits_2(self):
        with self.assertRaises(SystemExit) as cm:
            _parser().parse_args(["metrics", "--bogus"])
        self.assertEqual(cm.exception.code, 2)

    def test_metrics_json_is_single_line(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _clean_basename_file(d, "示例.bas", METRICS_BAS)
            args = _parser().parse_args(["metrics", "--json"])
            ctx = Ctx(args, _config(), d, json_mode=True)
            from features.agent_patch import handle_metrics
            outcome = handle_metrics(ctx)
            buf, err = io.StringIO(), io.StringIO()
            with redirect_stdout(buf), redirect_stderr(err):
                code = render(outcome, json_mode=True)
            lines = buf.getvalue().splitlines()
            self.assertEqual(len(lines), 1, buf.getvalue())
            payload = json.loads(lines[0])
            for key in ("ok", "action", "details", "errors", "elapsed_ms", "ts",
                        "stats", "exit_code"):
                self.assertIn(key, payload)
            self.assertEqual(payload["action"], "agent.metrics")
            self.assertEqual(payload["exit_code"], code)
            self.assertIn("lines", payload["stats"])
            self.assertEqual(payload["stats"]["procs"], 1)


from unittest.mock import patch  # noqa: E402


# ---------------------------------------------------------------------------
# assert 桩对象：不碰真实 Excel，会话/工作簿/宏执行全部替身
# ---------------------------------------------------------------------------
class _Code:
    def __init__(self, text: str):
        self.text = text
        self.CountOfLines = len(text.splitlines())

    def Lines(self, start: int, count: int) -> str:
        return self.text


class _Comp:
    def __init__(self, name: str, text: str):
        self.Name = name
        self.CodeModule = _Code(text)


class _Comps:
    def __init__(self, comps: dict[str, _Comp]):
        self.comps = comps

    def __iter__(self):
        return iter(self.comps.values())

    def Item(self, name: str):
        return self.comps[name]


class _Project:
    def __init__(self, comps: dict[str, _Comp]):
        self.VBComponents = _Comps(comps)


class _WB:
    def __init__(self, name: str, comps: dict[str, _Comp] | None = None):
        self.Name = name
        self.VBProject = _Project(comps or {})
        self.Sheets = []

    def Activate(self) -> None:
        pass


class _XL:
    def __init__(self, wb: _WB):
        self._wb = wb

    @property
    def ActiveWorkbook(self) -> _WB:
        return self._wb


class _FakeRunner:
    """MacroRunner 替身：可直接回结果，或跑完写一份日志。"""

    def __init__(self, result=None, log_path: Path | None = None, log_text: str = ""):
        from toolkit.macro import MacroResult
        self.result = result or MacroResult(macro="测试表格工具.运行全部断言")
        self.log_path = log_path
        self.log_text = log_text

    def run(self, macro, args=None, *, isolate=False, save=False, log_optional=True):
        if self.log_path is not None:
            self.log_path.write_text(self.log_text, encoding="gbk", errors="replace")
        return self.result


class _ScriptedRunner:
    """按调用顺序回结果/抛 Failure；记录每次调用的 argv，便于核对回退参数。"""

    def __init__(self, script, log_path: Path | None = None, log_text: str = ""):
        from toolkit.macro import MacroResult
        self.script = list(script)
        self.calls: list[list] = []
        self.log_path = log_path
        self.log_text = log_text
        self._default = MacroResult(macro="测试表格工具.运行全部断言")

    def run(self, macro, args=None, *, isolate=False, save=False, log_optional=True):
        self.calls.append(list(args or []))
        if self.log_path is not None:
            self.log_path.write_text(self.log_text, encoding="gbk", errors="replace")
        step = self.script.pop(0)
        if isinstance(step, Exception):
            raise step
        return step if step is not None else self._default


def _assert_config(d: Path) -> dict:
    xlsm = d / "示例表 V1.xlsm"
    xlsm.write_bytes(b"")
    return {"defaults": {"common_dir": "src/common"},
            "projects": [{"name": "示例表", "file": str(xlsm),
                          "assert_macro": "测试表格工具.运行全部断言",
                          "assert_sheet": "测试表", "search_dir": str(d)}]}


ASSERT_COMP = {"测试表格工具": _Comp("测试表格工具",
                                     "Public Sub 运行全部断言\nEnd Sub\n")}


class AssertTest(unittest.TestCase):
    def _run(self, d: Path, *argv: str, runner=None, wb: _WB | None = None,
             opened=None):
        argv = ("--project", "示例表", *argv)
        args = _parser().parse_args(["assert", *argv])
        ctx = Ctx(args, _assert_config(d), d, json_mode=args.json, runner=runner)
        wb = wb or _WB("示例表 V1.xlsm", ASSERT_COMP)
        xl = _XL(wb)
        with patch("features.agent_patch.excel_gateway.attach_running_excel",
                   return_value=xl), \
             patch("features.agent_patch.excel_gateway.find_open_workbook",
                   return_value=(wb, opened if opened is not None else [wb.Name])):
            from features.agent_patch import handle_assert
            return handle_assert(ctx)

    def _log_dir(self, d: Path) -> Path:
        folder = d / "测试输出"
        folder.mkdir(parents=True, exist_ok=True)
        return folder

    def test_assert_success(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            folder = self._log_dir(d)
            log = folder / "run.log"
            runner = _FakeRunner(log_path=log, log_text="通过: 5\n失败: 0\n")
            outcome = self._run(d, "--apply", runner=runner)
            self.assertEqual(outcome.exit_code, 0, outcome.human)
            self.assertTrue(outcome.ok)
            self.assertEqual(outcome.stats["passed"], 5)
            self.assertEqual(outcome.stats["failed"], 0)

    def test_assert_clear_to_baseline(self):
        class _Row:
            def __init__(self, rows):
                self.rows = rows

            def Delete(self):
                self.rows.count -= 1

        class _ListRows:
            def __init__(self, count):
                self.count = count

            @property
            def Count(self):
                return self.count

            def __call__(self, index):
                return _Row(self)

        class _ListObject:
            def __init__(self):
                self.ListRows = _ListRows(5)

        class _ListObjects:
            Count = 1

            def __init__(self):
                self.lo = _ListObject()

            def __call__(self, index):
                return self.lo

        class _Sheet:
            Name = "测试表"
            ListObjects = _ListObjects()

        class _Sheets:
            def __init__(self, sheets):
                self._sheets = sheets

            def __iter__(self):
                return iter(self._sheets)

            def __call__(self, name):
                return next(s for s in self._sheets if s.Name == name)

        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            folder = self._log_dir(d)
            log = folder / "run.log"
            runner = _FakeRunner(log_path=log, log_text="通过: 5\n失败: 0\n")
            wb = _WB("示例表 V1.xlsm", ASSERT_COMP)
            sheet = _Sheet()
            wb.Sheets = _Sheets([sheet])
            outcome = self._run(d, "--clear", "--baseline", "2", "--apply",
                                runner=runner, wb=wb)
            self.assertEqual(outcome.exit_code, 0, outcome.human)
            self.assertEqual(sheet.ListObjects(1).ListRows.Count, 2)
            self.assertTrue(any("5 → 2" in line for line in outcome.human))

    def test_assert_zero_arg_success_no_retry(self):
        # 零参一次成功：不触发第二次调用
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            folder = self._log_dir(d)
            log = folder / "run.log"
            runner = _ScriptedRunner([None], log_path=log,
                                     log_text="通过: 5\n失败: 0\n")
            outcome = self._run(d, "--apply", runner=runner)
            self.assertEqual(outcome.exit_code, 0, outcome.human)
            self.assertEqual(runner.calls, [[]])

    def test_assert_arg_mismatch_retries_two_args(self):
        # 零参 ARG_MISMATCH → 两参 (False, True) 成功
        first = Failure("参数个数不符", code="ARG_MISMATCH", stage="run")
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            folder = self._log_dir(d)
            log = folder / "run.log"
            runner = _ScriptedRunner([first, None], log_path=log,
                                     log_text="通过: 3\n失败: 0\n")
            outcome = self._run(d, "--apply", runner=runner)
            self.assertEqual(outcome.exit_code, 0, outcome.human)
            self.assertTrue(outcome.ok)
            self.assertEqual(runner.calls, [[], [False, True]])

    def test_assert_two_attempts_both_mismatch(self):
        # 两支都参数个数不符：最终报 ASSERT_MACRO_ARGS，不把 ARG_MISMATCH 当最终码
        first = Failure("参数个数不符", code="ARG_MISMATCH", stage="run")
        second = Failure("参数个数不符", code="ARG_MISMATCH", stage="run")
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            runner = _ScriptedRunner([first, second])
            with self.assertRaises(Failure) as cm:
                self._run(d, "--apply", runner=runner)
            self.assertEqual(cm.exception.code, "ASSERT_MACRO_ARGS")
            self.assertEqual(cm.exception.stage, "assert")
            self.assertIn("零参", cm.exception.hint)
            self.assertNotEqual(cm.exception.code, "ARG_MISMATCH")
            self.assertEqual(len(runner.calls), 2)

    def test_assert_second_non_mismatch_propagates(self):
        # 第二次是非参数个数错误：原样上抛，不糊成 ASSERT_MACRO_ARGS
        first = Failure("参数个数不符", code="ARG_MISMATCH", stage="run")
        second = Failure("宏不可用", code="RUN_FAILED", stage="run")
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            runner = _ScriptedRunner([first, second])
            with self.assertRaises(Failure) as cm:
                self._run(d, "--apply", runner=runner)
            self.assertEqual(cm.exception.code, "RUN_FAILED")
            self.assertEqual(len(runner.calls), 2)

    def test_assert_failure_exit_1_with_code(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            folder = self._log_dir(d)
            log = folder / "run.log"
            runner = _FakeRunner(log_path=log, log_text="通过: 2\n失败: 1\nErr 9 某处\n")
            outcome = self._run(d, "--apply", runner=runner)
            self.assertEqual(outcome.exit_code, 1)
            self.assertFalse(outcome.ok)
            self.assertEqual(outcome.errors[0].code, "ASSERT_FAILED")
            self.assertEqual(outcome.stats["failed"], 1)
            # B5 复核：业务失败退出 1（旧版 2 口径不恢复），hint 必须 actionable、
            # suggested_command 是可直接执行的绝对入口命令。
            err = outcome.errors[0]
            self.assertTrue(err.hint)
            self.assertIn("agent logs", err.hint)
            self.assertTrue(err.suggested)
            self.assertIn(sys.executable, err.suggested)
            tokens = shlex.split(err.suggested, posix=False)
            self.assertEqual(Path(tokens[1].strip('"')).name, "agent.py")

    def test_assert_no_table_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            folder = self._log_dir(d)
            log = folder / "run.log"
            runner = _FakeRunner(log_path=log, log_text="没有找到名为 测试表\n")
            with self.assertRaises(Failure) as cm:
                self._run(d, "--apply", runner=runner)
            self.assertEqual(cm.exception.code, "ASSERT_NO_TABLE")
            self.assertEqual(cm.exception.stage, "assert")

    def test_assert_failed_uncountable_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            folder = self._log_dir(d)
            log = folder / "run.log"
            runner = _FakeRunner(log_path=log, log_text="杂项日志，没有计数\n")
            with self.assertRaises(Failure) as cm:
                self._run(d, "--apply", runner=runner)
            self.assertEqual(cm.exception.code, "ASSERT_FAILED")
            self.assertIn("agent logs", cm.exception.hint)

    def test_assert_no_log(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            self._log_dir(d)
            runner = _FakeRunner()
            with self.assertRaises(Failure) as cm:
                self._run(d, "--apply", "--timeout", "0.1", runner=runner)
            self.assertEqual(cm.exception.code, "ASSERT_NO_LOG")
            self.assertEqual(cm.exception.stage, "assert")
            suggested = cm.exception.suggested
            # 必须是可直接执行的绝对入口命令，不是包内相对路径
            self.assertNotIn("vba-workbench/scripts/", suggested)
            self.assertIn(sys.executable, suggested)
            tokens = shlex.split(suggested, posix=False)
            script = Path(tokens[1].strip('"'))
            self.assertTrue(script.is_file(), str(script))
            self.assertEqual(script.name, "sync.py")

    def test_assert_watchdog_timeout(self):
        from toolkit.macro import MacroResult
        result = MacroResult(macro="测试表格工具.运行全部断言", timed_out=True)
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            runner = _FakeRunner(result=result)
            with self.assertRaises(Failure) as cm:
                self._run(d, "--apply", runner=runner)
            self.assertEqual(cm.exception.code, "WATCHDOG_TIMEOUT")
            self.assertEqual(cm.exception.stage, "assert")

    def test_assert_macro_not_configured(self):
        wb = _WB("示例表 V1.xlsm", {})
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            runner = _FakeRunner()
            with self.assertRaises(Failure) as cm:
                self._run(d, "--apply", runner=runner, wb=wb)
            self.assertEqual(cm.exception.code, "ASSERT_MACRO_NOT_CONFIGURED")

    def test_assert_target_not_open(self):
        from toolkit.macro import MacroResult
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            runner = _FakeRunner(result=MacroResult(macro="x"))
            args = _parser().parse_args(["assert", "--project", "示例表", "--apply"])
            ctx = Ctx(args, _assert_config(d), d, runner=runner)
            with patch("features.agent_patch.excel_gateway.attach_running_excel",
                       return_value=_XL(_WB("x"))), \
                 patch("features.agent_patch.excel_gateway.find_open_workbook",
                       return_value=(None, [])):
                from features.agent_patch import handle_assert
                with self.assertRaises(Failure) as cm:
                    handle_assert(ctx)
            self.assertEqual(cm.exception.code, "TARGET_NOT_OPEN")

    def test_assert_usage_error_exits_2(self):
        with self.assertRaises(SystemExit) as cm:
            _parser().parse_args(["assert", "--timeout", "not-a-number"])
        self.assertEqual(cm.exception.code, 2)

    def test_assert_json_is_single_line(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            folder = self._log_dir(d)
            log = folder / "run.log"
            runner = _FakeRunner(log_path=log, log_text="通过: 5\n失败: 0\n")
            outcome = self._run(d, "--apply", "--json", runner=runner)
            buf, err = io.StringIO(), io.StringIO()
            with redirect_stdout(buf), redirect_stderr(err):
                code = render(outcome, json_mode=True)
            lines = buf.getvalue().splitlines()
            self.assertEqual(len(lines), 1, buf.getvalue())
            payload = json.loads(lines[0])
            for key in ("ok", "action", "details", "errors", "elapsed_ms", "ts",
                        "stats", "exit_code"):
                self.assertIn(key, payload)
            self.assertEqual(payload["action"], "agent.assert")
            self.assertEqual(payload["exit_code"], code)
            self.assertIn("lines", payload["stats"])
            self.assertEqual(payload["stats"]["passed"], 5)


# ---------------------------------------------------------------------------
# metrics 失败路径：源码读盘失败 / 基线落盘失败，经入口兜底成 UNEXPECTED 信封。
# 进程内调 agent.py main；读盘失败只桩读盘原语，真实「入口 → dispatch → 兜底 → render」
# 链路保留。当前契约下两类失败都归 UNEXPECTED/exit 1（专用归类留待后续切片）。
# ---------------------------------------------------------------------------
class MetricsFailureEntryTest(unittest.TestCase):
    def _run_entry(self, argv: list[str], tmp: Path, *, patch_read: bool = False):
        import agent as agent_entry  # noqa: E402
        import features.agent_patch as ap  # noqa: E402
        from toolkit.base import set_json_mode  # noqa: E402

        cfg = tmp / "mapping.json"
        cfg.write_text(json.dumps({"projects": []}, ensure_ascii=False),
                       encoding="utf-8")
        old_env = os.environ.get("VBA_WORKBENCH_CONFIG")
        os.environ["VBA_WORKBENCH_CONFIG"] = str(cfg)
        orig_read = ap.code_lines
        if patch_read:
            def _raise_read(p, use_cache=True):
                raise OSError("源文件读取失败（模拟）")
            ap.code_lines = _raise_read
        buf, err = io.StringIO(), io.StringIO()
        try:
            with redirect_stdout(buf), redirect_stderr(err):
                code = agent_entry.main(argv)
        finally:
            ap.code_lines = orig_read
            set_json_mode(False)
            if old_env is None:
                os.environ.pop("VBA_WORKBENCH_CONFIG", None)
            else:
                os.environ["VBA_WORKBENCH_CONFIG"] = old_env
        return code, buf.getvalue(), err.getvalue()

    def test_metrics_source_read_failure_is_unexpected_exit1(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            code, out, err = self._run_entry(["metrics", "--json"], d, patch_read=True)
            self.assertEqual(code, 1, err)
            payload = json.loads(out.strip())
            self.assertFalse(payload["ok"])
            self.assertEqual(payload["exit_code"], 1)
            self.assertEqual(payload["errors"][0]["code"], "UNEXPECTED")
            self.assertEqual(payload["errors"][0]["stage"], "unknown")
            hint = payload["errors"][0]["hint"]
            self.assertTrue(hint)
            self.assertIn("--json", hint)

    def test_metrics_report_write_failure_is_unexpected_exit1(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            blocker = d / "blocker.txt"
            blocker.write_text("x", encoding="utf-8")
            out = blocker / "metrics.md"  # 父级是已存在文件，mkdir 落盘失败
            code, stdout, err = self._run_entry(
                ["metrics", "--out", str(out), "--apply", "--json"], d)
            self.assertEqual(code, 1, err)
            payload = json.loads(stdout.strip())
            self.assertFalse(payload["ok"])
            self.assertEqual(payload["exit_code"], 1)
            self.assertEqual(payload["errors"][0]["code"], "UNEXPECTED")
            self.assertEqual(payload["errors"][0]["stage"], "unknown")
            self.assertTrue(payload["errors"][0]["hint"])


if __name__ == "__main__":
    unittest.main()
