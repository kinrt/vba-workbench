# -*- coding: utf-8 -*-
"""--json 失败也要给原因：项目名打错、目标缺失、公共库未配置都要有固定错误码。"""

import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
SYNC = PKG / "scripts" / "sync.py"
sys.path.insert(0, str(PKG / "scripts"))

from toolkit.base import Failure as BaseFailure  # noqa: E402
from toolkit.contract import (ErrorCode, Failure as ContractFailure, Outcome,  # noqa: E402
                              Stage, to_contract_failure)


class ErrorContractTest(unittest.TestCase):
    def run_sync(self, d: Path, cfg: dict, *argv: str):
        cfg_path = d / "mapping.json"
        cfg_path.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
        env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(cfg_path), PYTHONIOENCODING="utf-8")
        return subprocess.run([sys.executable, str(SYNC), *argv],
                              cwd=str(d), env=env, capture_output=True, text=True, encoding="utf-8")

    def test_unknown_project_reports_reason(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            cfg = {
                "projects": [{"name": "示例表A", "search_dir": str(d), "prefix": "示例表A",
                              "require_version": False, "src": "src/projects/示例表A"}],
            }
            r = self.run_sync(d, cfg, "push", "--project", "不存在项目", "--json")
            payload = json.loads(r.stdout.strip())
            self.assertEqual(r.returncode, 1)
            self.assertFalse(payload["ok"])
            self.assertTrue(payload["errors"], "失败时 errors 不能为空")
            self.assertIn("不存在项目", payload["errors"][0]["message"])
            self.assertIn("retryable", payload["errors"][0])
            self.assertEqual(payload["exit_code"], 1)
            self.assertEqual(payload["action"], "sync.push")

    def test_unknown_project_human_shows_error_not_silent(self):
        # 人话模式：handler 把失败收进 Outcome.errors 且 human 行为空时，render 出口兜底
        # 必须补打 [错误]，不能静默 exit 1（与 audit 业务失败同一类契约）。
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            cfg = {
                "projects": [{"name": "示例表A", "search_dir": str(d), "prefix": "示例表A",
                              "require_version": False, "src": "src/projects/示例表A"}],
            }
            r = self.run_sync(d, cfg, "push", "--project", "不存在项目", "--no-input")
            self.assertEqual(r.returncode, 1, r.stderr)
            self.assertIn("[错误]", r.stderr)
            self.assertNotIn("Traceback", r.stderr)

    def test_list_missing_target_reports_target_not_found(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            cfg = {
                "projects": [{"name": "实验室", "search_dir": str(d), "prefix": "实验室",
                              "require_version": True, "src": "src/projects/实验室"}],
            }
            r = self.run_sync(d, cfg, "list", "--json")
            self.assertEqual(r.returncode, 1, r.stderr)
            self.assertEqual(len(r.stdout.strip().splitlines()), 1, r.stdout)
            payload = json.loads(r.stdout.strip())
            self.assertFalse(payload["ok"])
            self.assertEqual(payload["errors"][0]["code"], "TARGET_NOT_FOUND")
            self.assertEqual(payload["errors"][0]["stage"], "resolve")
            self.assertIn("实验室", payload["errors"][0]["hint"])
            self.assertEqual(payload["details"][0]["workbook"], "")

    def test_version_without_common_reports_configured_false(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            cfg = {"projects": [{"name": "甲", "src": "src/projects/甲"}]}
            r = self.run_sync(d, cfg, "version", "--json")
            self.assertEqual(r.returncode, 0, r.stderr)
            payload = json.loads(r.stdout.strip())
            self.assertFalse(payload["stats"]["configured"])
            self.assertEqual(payload["stats"]["mismatched"], 0)
            self.assertIn("未配置", payload["details"][0]["reason"])


class FailureClassifyTest(unittest.TestCase):
    """同一基础 Failure 经三入口共用归类规则转换，code/stage 必须一致。"""

    def _entries(self, e):
        sync_err = to_contract_failure(e)
        agent_err = to_contract_failure(e)
        audit_out = Outcome(action="audit.members", ok=False,
                            errors=[to_contract_failure(e)], exit_code=1)
        return sync_err, agent_err, audit_out.errors[0]

    def test_no_code_failure_unified_unexpected(self):
        e = BaseFailure("配置层旧骨架，没带码")
        errs = self._entries(e)
        for err in errs:
            self.assertEqual(err.code, ErrorCode.UNEXPECTED.value)
            self.assertEqual(err.stage, Stage.UNKNOWN.value)
            self.assertEqual(str(err), "配置层旧骨架，没带码")

    def test_explicit_code_and_stage_preserved(self):
        e = BaseFailure("缺配置对象", code="CONFIG_NOT_FOUND", stage="config",
                        hint="去配 mapping")
        errs = self._entries(e)
        for err in errs:
            self.assertEqual(err.code, "CONFIG_NOT_FOUND")
            self.assertEqual(err.stage, "config")
            self.assertEqual(err.hint, "去配 mapping")

    def test_contract_failure_passthrough(self):
        e = ContractFailure("契约层失败", code=ErrorCode.TARGET_NOT_FOUND,
                            stage=Stage.RESOLVE)
        errs = self._entries(e)
        for err in errs:
            self.assertIs(err, e)


class StageEnumTest(unittest.TestCase):
    """Stage 必须是登记成员：新成员正例放行，未登记阶段构造即拒。"""

    def test_registered_members_accepted(self):
        for member in (Stage.PREFLIGHT, Stage.POSTCHECK, Stage.INIT):
            f = ContractFailure("阶段校验", code=ErrorCode.USAGE, stage=member)
            self.assertEqual(f.stage, member.value)
            f_str = ContractFailure("阶段校验", code=ErrorCode.USAGE,
                                    stage=member.value)
            self.assertEqual(f_str.stage, member.value)

    def test_unregistered_stage_rejected(self):
        with self.assertRaises(ValueError):
            ContractFailure("阶段校验", code=ErrorCode.USAGE, stage="overview")
        with self.assertRaises(ValueError):
            ContractFailure("阶段校验", code=ErrorCode.USAGE, stage="backup")
        with self.assertRaises(ValueError):
            ContractFailure("阶段校验", code=ErrorCode.USAGE, stage="")


class HumanFallbackRenderTest(unittest.TestCase):
    """render 人话出口兜底：errors 无 human 行时补 [错误]；已有 human 行不重复打印。"""

    def test_empty_human_fallback_prints_error_lines(self):
        from toolkit.render import render
        err = ContractFailure("找不到项目", code=ErrorCode.PROJECT_NOT_FOUND,
                               stage=Stage.RESOLVE, hint="检查项目名",
                               suggested="sync.py list")
        outcome = Outcome(action="sync.push", ok=False, exit_code=1,
                          errors=[err], stats={})
        out, errbuf = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(errbuf):
            code = render(outcome, json_mode=False)
        self.assertEqual(code, 1)
        text = errbuf.getvalue()
        self.assertIn("[错误] 找不到项目", text)
        self.assertIn("说明：检查项目名", text)
        self.assertIn("建议：sync.py list", text)

    def test_human_lines_present_do_not_repeat_errors(self):
        from toolkit.render import render
        err = ContractFailure("表头错误", code=ErrorCode.PROJECT_NOT_FOUND,
                              stage=Stage.RESOLVE, hint="去配表")
        outcome = Outcome(action="sync.push", ok=False, exit_code=1,
                          errors=[err], human=["预览：未写入任何文件"], stats={})
        out, errbuf = io.StringIO(), io.StringIO()
        with redirect_stdout(out), redirect_stderr(errbuf):
            code = render(outcome, json_mode=False)
        self.assertEqual(code, 1)
        self.assertIn("预览：未写入任何文件", out.getvalue())
        self.assertNotIn("[错误]", errbuf.getvalue())


if __name__ == "__main__":
    unittest.main()
