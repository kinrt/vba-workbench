# -*- coding: utf-8 -*-
"""1.6.5 定向回归：N5（style stderr 警告并入报告行的前缀归一）+ 风险点 3 缺测。

覆盖四点（不启动 Excel，COM / 源码边界全部桩对象或临时目录测试件）：

1. 非 JSON 模式：风格报告含 warn 时，``human`` / ``stats.lines`` 能看到警告内容，
   且带 ``[警告]`` 前缀；
2. JSON 成功路径：``stats.lines`` 非空，渲染后捕获的 stderr 为空；
3. JSON 失败路径（脏文件 / 风格不合规）：信封可解析、``stats.lines`` 有内容、
   stderr 为空，警告收口进 ``stats.warnings`` 而不是混进行报告行；
4. 重定向期间异常（``fail()`` / traceback）不被吞，异常抛出后
   ``sys.stdout`` / ``sys.stderr`` 正确还原。

口径边界：不改退出码与 P1-4（exit 2 仅人话）、P2-6（audit 全仓 NONE）；
本文件只桩 ``features.agent_patch`` 链路，不碰其它切片既有用例。
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

from toolkit.base import (  # noqa: E402
    reset_warnings, set_json_mode,
)
from toolkit.commands import specs_for  # noqa: E402
from toolkit.context import Progress  # noqa: E402
from toolkit.registry import build_parser  # noqa: E402
from toolkit.render import render  # noqa: E402


class Ctx:
    """最小运行上下文：形状对齐 ``toolkit.context.AppContext``，测试可直接塞桩。"""

    def __init__(self, args, config: dict, root: Path, json_mode: bool = False):
        self.args = args
        self.config = config
        self.root = root
        self.json_mode = json_mode
        self.started = time.perf_counter()
        self.progress = Progress(tool="agent", enabled=not json_mode)

    def elapsed_ms(self) -> int:
        return int((time.perf_counter() - self.started) * 1000)

    def stage_timings(self) -> dict:
        return dict(self.progress.timings)


def _parser():
    return build_parser(specs_for("agent"), prog="agent.py")


def _config() -> dict:
    return {"defaults": {"common_dir": "src/common"}, "projects": []}


def _write_bas(d: Path, name: str, body: str) -> Path:
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


def _run_style(d: Path, *argv: str):
    args = _parser().parse_args(["style", *argv])
    ctx = Ctx(args, _config(), d, json_mode=args.json)
    from features.agent_patch import handle_style
    return handle_style(ctx)


class StyleWarningPrefixTest(unittest.TestCase):
    """N5 收口：警告独立走 Outcome.warnings，人话报告行不再混入 [警告] 前缀。"""

    def test_warn_line_moves_to_warnings_channel(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _write_bas(d, "坏样例.bas", DIRTY_BAS)
            outcome = _run_style(d, "--action", "check", "--scope", "src")

            self.assertEqual(outcome.exit_code, 1)
            self.assertFalse([ln for ln in outcome.human if ln.startswith("[警告]")])
            self.assertTrue(any("合计" in w for w in outcome.warnings))
            # 报告主体仍是同一条 human / stats.lines 通道
            self.assertEqual(outcome.human, outcome.stats["lines"])

    def test_error_line_keeps_error_prefix_not_relabeled(self):
        """人话模式：fail() 的 [错误] 行保持错误标记，不被补成 [警告]。"""
        from features import agent_patch
        from toolkit.base import fail

        def _check_with_error(files, which):
            fail("模拟风格错误")
            return 0, []

        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _write_bas(d, "示例.bas", CLEAN_BAS)
            args = _parser().parse_args(["style", "--action", "check", "--scope", "src"])
            ctx = Ctx(args, _config(), d)
            with mock.patch.object(agent_patch, "run_checks_with_details",
                                   side_effect=_check_with_error):
                outcome = agent_patch.handle_style(ctx)
            self.assertTrue(any("[错误] 模拟风格错误" in ln for ln in outcome.human))
            self.assertFalse(any("[警告] [错误]" in ln for ln in outcome.human))

    def test_normalize_dedupes_repeated_prefix(self):
        """纯函数口径：重复前缀归一为一个；缺前缀补齐；空行不动。"""
        from features.agent_patch import _normalize_warn_line

        self.assertEqual(_normalize_warn_line("[警告] [警告] 重复前缀"),
                         "[警告] 重复前缀")
        self.assertEqual(_normalize_warn_line("裸警告正文"), "[警告] 裸警告正文")
        self.assertEqual(_normalize_warn_line("  [警告]  缩进保留"), "[警告] 缩进保留")
        self.assertEqual(_normalize_warn_line(""), "")
        self.assertEqual(_normalize_warn_line("[警告]"), "[警告]")


class StyleJsonSuccessTest(unittest.TestCase):
    """缺测 2：JSON 成功路径，lines 非空、stderr 为空、单行信封可解析。"""

    def setUp(self):
        set_json_mode(True)
        reset_warnings()
        self.addCleanup(set_json_mode, False)
        self.addCleanup(reset_warnings)

    def test_success_lines_nonempty_stderr_clean(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _write_bas(d, "示例.bas", CLEAN_BAS)
            outcome = _run_style(d, "--action", "check", "--scope", "src", "--json")

            self.assertEqual(outcome.exit_code, 0, outcome.human)
            self.assertTrue(outcome.stats["lines"], "JSON 成功路径 stats.lines 不能为空")

            buf, err = io.StringIO(), io.StringIO()
            with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(err):
                code = render(outcome, json_mode=True)
            self.assertEqual(code, 0)
            lines = buf.getvalue().splitlines()
            self.assertEqual(len(lines), 1, buf.getvalue())
            payload = json.loads(lines[0])
            self.assertEqual(payload["action"], "agent.style")
            self.assertTrue(payload["stats"]["lines"])
            self.assertEqual(err.getvalue(), "", f"JSON 出口 stderr 被污染：{err.getvalue()!r}")


class StyleJsonFailureTest(unittest.TestCase):
    """缺测 3：JSON 失败路径（脏文件），信封可解析、lines 有内容、stderr 为空。"""

    def setUp(self):
        set_json_mode(True)
        reset_warnings()
        self.addCleanup(set_json_mode, False)
        self.addCleanup(reset_warnings)

    def test_failure_envelope_parseable_lines_nonempty(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _write_bas(d, "坏样例.bas", DIRTY_BAS)
            outcome = _run_style(d, "--action", "check", "--scope", "src", "--json")

            self.assertEqual(outcome.exit_code, 1)
            self.assertFalse(outcome.ok)
            self.assertEqual(outcome.errors[0].code, "SOURCE_PREFLIGHT")
            self.assertTrue(outcome.stats["lines"], "JSON 失败路径 stats.lines 不能为空")

            buf, err = io.StringIO(), io.StringIO()
            with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(err):
                code = render(outcome, json_mode=True)
            self.assertEqual(code, 1)
            lines = buf.getvalue().splitlines()
            self.assertEqual(len(lines), 1, buf.getvalue())
            payload = json.loads(lines[0])
            self.assertEqual(payload["action"], "agent.style")
            self.assertEqual(payload["exit_code"], 1)
            self.assertTrue(payload["stats"]["lines"])
            self.assertEqual(err.getvalue(), "", f"JSON 出口 stderr 被污染：{err.getvalue()!r}")

            # 警告收口进 stats.warnings，不混进 stats.lines 的报告行
            warnings = payload["stats"].get("warnings", [])
            self.assertTrue(any("合计" in w for w in warnings),
                            f"警告没收口进 warnings：{warnings!r}")
            self.assertFalse(any(ln.startswith("[警告]") for ln in payload["stats"]["lines"]),
                             f"警告前缀混进了报告行：{payload['stats']['lines']!r}")


class StyleRedirectExceptionTest(unittest.TestCase):
    """缺测 4：重定向期间异常不被吞，抛出后 stdout/stderr 正确还原。"""

    def _ctx(self, d: Path) -> Ctx:
        args = _parser().parse_args(["style", "--action", "check", "--scope", "src"])
        return Ctx(args, _config(), d)

    def test_exception_propagates_and_streams_restored(self):
        from features import agent_patch

        def _boom(files, which):
            # 先经 fail() 写一条 [错误]，再抛 traceback；两者都不许被静默吞掉
            from toolkit.base import fail
            fail("模拟风格检查失败")
            raise RuntimeError("boom: 风格检查链路异常")

        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _write_bas(d, "示例.bas", CLEAN_BAS)
            ctx = self._ctx(d)

            orig_stdout, orig_stderr = sys.stdout, sys.stderr
            with mock.patch.object(agent_patch, "run_checks_with_details", side_effect=_boom):
                with self.assertRaises(RuntimeError) as cm:
                    agent_patch.handle_style(ctx)
            self.assertIn("boom", str(cm.exception))

            # 重定向退出后流对象必须还原，不能留 StringIO 挂在 sys 上
            self.assertIs(sys.stdout, orig_stdout, "sys.stdout 未还原")
            self.assertIs(sys.stderr, orig_stderr, "sys.stderr 未还原")

    def test_failure_path_exception_also_restores_streams(self):
        """run_fix 路径抛异常同样不吞、流同样还原。"""
        from features import agent_patch

        def _boom(files, which, apply=False):
            raise RuntimeError("boom: 风格修复链路异常")

        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            _write_bas(d, "示例.bas", CLEAN_BAS)
            args = _parser().parse_args(["style", "--action", "fix", "--scope", "src"])
            ctx = Ctx(args, _config(), d)

            orig_stdout, orig_stderr = sys.stdout, sys.stderr
            with mock.patch.object(agent_patch, "run_fix", side_effect=_boom):
                with self.assertRaises(RuntimeError) as cm:
                    agent_patch.handle_style(ctx)
            self.assertIn("boom", str(cm.exception))
            self.assertIs(sys.stdout, orig_stdout)
            self.assertIs(sys.stderr, orig_stderr)


if __name__ == "__main__":
    unittest.main()
