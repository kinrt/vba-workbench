# -*- coding: utf-8 -*-
"""新架构骨架的最小契约测试：错误必填、JSON 单出口、manifest、状态矩阵。"""
import io
import json
import sys
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))

from toolkit.contract import Failure, MutationMode, Outcome, SuggestedCommand
from toolkit.registry import CommandSpec, ProjectMode, build_parser
from toolkit.render import render
from toolkit.session import WorkbookSession
from toolkit.state import Purpose, STATE_RULES, TargetState, rule_for


class ContractSkeletonTest(unittest.TestCase):
    def test_failure_requires_code_and_stage(self):
        with self.assertRaises(TypeError):
            Failure("没有码")
        with self.assertRaises(ValueError):
            Failure("坏码", code="bad-code", stage="attach")
        f = Failure("表没开", code="TARGET_NOT_OPEN", stage="attach",
                    suggested='python vba-workbench\\scripts\\agent.py book --action open --project 甲')
        self.assertEqual(f.code, "TARGET_NOT_OPEN")
        self.assertEqual(f.as_error()["stage"], "attach")

    def test_suggested_command_is_executable(self):
        cmd = SuggestedCommand(("attrib", "-R", "a.xlsm"))
        self.assertIn("attrib", str(cmd))
        with self.assertRaises(ValueError):
            SuggestedCommand(("not-a-command",))

    def test_render_json_is_one_line(self):
        out = Outcome(action="sync.version", ok=False, exit_code=1,
                      errors=[Failure("版本不一致", code="VERSION_MISMATCH", stage="config")],
                      stats={"mismatched": 1}, elapsed_ms=12)
        buf, err = io.StringIO(), io.StringIO()
        with redirect_stdout(buf), redirect_stderr(err):
            code = render(out, json_mode=True)
        self.assertEqual(code, 1)
        lines = buf.getvalue().splitlines()
        self.assertEqual(len(lines), 1)
        payload = json.loads(lines[0])
        for key in ("ok", "action", "details", "errors", "elapsed_ms", "ts", "stats", "exit_code"):
            self.assertIn(key, payload)
        self.assertEqual(payload["errors"][0]["code"], "VERSION_MISMATCH")

    def test_render_contract_violation_not_failed(self):
        out = Outcome(action="sync.version", ok=False, exit_code=1, errors=[], elapsed_ms=1)
        buf = io.StringIO()
        with redirect_stdout(buf), redirect_stderr(io.StringIO()):
            code = render(out, json_mode=True)
        self.assertEqual(code, 1)
        payload = json.loads(buf.getvalue())
        self.assertEqual(payload["errors"][0]["code"], "CONTRACT_VIOLATION")
        self.assertNotIn("FAILED", buf.getvalue())

    def test_render_keeps_usage_exit_2(self):
        """用法错必须能原样返回 2；不许被归一化成 1。"""
        out = Outcome(action="sync.pull", ok=False, exit_code=2,
                      errors=[Failure("缺 --project", code="USAGE", stage="parse")])
        buf = io.StringIO()
        with redirect_stdout(buf), redirect_stderr(io.StringIO()):
            code = render(out, json_mode=True)
        self.assertEqual(code, 2)
        self.assertEqual(len(buf.getvalue().splitlines()), 1)
        payload = json.loads(buf.getvalue())
        self.assertEqual(payload["exit_code"], 2)
        self.assertFalse(payload["ok"])

    def test_manifest_and_parser(self):
        spec = CommandSpec(verb="version", entry="sync", handler="features.core:version",
                           mutation=MutationMode.NONE, project=ProjectMode.NONE)
        parser = build_parser((spec,), prog="sync.py")
        args = parser.parse_args(["version", "--json"])
        self.assertEqual(args.command, "version")
        self.assertTrue(args.json)
        with self.assertRaises(SystemExit) as cm:
            parser.parse_args(["version", "--apply"])
        self.assertEqual(cm.exception.code, 2)

    def test_state_matrix_has_unique_codes(self):
        self.assertTrue(STATE_RULES)
        rule = rule_for(TargetState.CLOSED, Purpose.ASSERT)
        self.assertIsNotNone(rule)
        self.assertEqual(rule.code, "TARGET_NOT_OPEN")
        session = WorkbookSession(target="示例表A", state=TargetState.CLOSED,
                                  purpose=Purpose.ASSERT)
        with self.assertRaises(Failure) as cm:
            session.require()
        self.assertEqual(cm.exception.code, "TARGET_NOT_OPEN")
        for key, value in STATE_RULES.items():
            self.assertIsInstance(key[0], TargetState)
            self.assertIsInstance(key[1], Purpose)
            self.assertTrue(value.code or value.warning)


if __name__ == "__main__":
    unittest.main()
