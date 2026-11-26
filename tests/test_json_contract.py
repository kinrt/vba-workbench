# -*- coding: utf-8 -*-
"""--json 契约：stdout 只能有一行合法 JSON，人话/警告全部走 stderr。"""

import json
import os
import subprocess
import sys
import tempfile
import time
import unittest
import contextlib
import io
from pathlib import Path
from types import SimpleNamespace

PKG = Path(__file__).resolve().parents[1]
SYNC = PKG / "scripts" / "sync.py"
AUDIT = PKG / "scripts" / "audit.py"
AGENT = PKG / "scripts" / "agent.py"
sys.path.insert(0, str(PKG / "scripts"))

from toolkit.base import ROOT, set_json_mode  # noqa: E402
from toolkit.commands import specs_for  # noqa: E402
from toolkit.context import AppContext, Progress  # noqa: E402
from toolkit.registry import build_parser  # noqa: E402
from toolkit.render import render  # noqa: E402
from toolkit import excel  # noqa: E402


class JsonContractTest(unittest.TestCase):
    def test_list_stdout_is_single_json_line(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            cfg = d / "mapping.json"
            cfg.write_text(json.dumps({
                "defaults": {"common_dir": "src/common"},
                "projects": [{"name": "示例表A", "search_dir": str(d), "prefix": "示例表A",
                              "require_version": False, "src": "src/projects/示例表A"}],
            }, ensure_ascii=False), encoding="utf-8")
            # 覆盖层存在时 config 会打「已合并本机覆盖层」：JSON 模式收口进 stats.warnings，
            # 不碰 stderr，保证 cmd 2>&1 后 stderr 不被过程提示污染
            (d / "mapping.local.json").write_text(
                json.dumps({"defaults": {"require_version": False}}), encoding="utf-8")

            env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(cfg), PYTHONIOENCODING="utf-8")
            r = subprocess.run([sys.executable, str(SYNC), "list", "--json"],
                               cwd=str(d), env=env, capture_output=True, text=True, encoding="utf-8")
            out = r.stdout.strip()
            self.assertEqual(len(out.splitlines()), 1, f"stdout 不止一行：{r.stdout!r}")
            payload = json.loads(out)
            self.assertEqual(payload["action"], "sync.list")
            self.assertEqual(payload["exit_code"], 1)
            self.assertIn("errors", payload)
            self.assertNotIn("已合并本机覆盖层", r.stderr, f"stderr 漏了过程提示：{r.stderr!r}")
            self.assertTrue(any("已合并本机覆盖层" in w for w in payload["stats"].get("warnings", [])),
                            f"覆盖层提示没收口进 warnings：{payload['stats']!r}")
            self.assertRegex(payload["ts"], r"T\d\d:\d\d:\d\d(?:[+-]\d\d:\d\d|Z)$")

    def test_require_open_workbook_json_keeps_stderr_clean(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "目标.xlsm"
            p.write_bytes(b"x")
            wb = SimpleNamespace(Name="目标.xlsm", FullName=str(p), Activate=lambda: None)
            xl = SimpleNamespace(Workbooks=[wb], ActiveWorkbook=wb)
            set_json_mode(True)
            self.addCleanup(set_json_mode, False)
            out, err = io.StringIO(), io.StringIO()
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                excel.require_open_workbook(xl, p, "示例表A", "跑宏")
            self.assertEqual(out.getvalue(), "")
            self.assertEqual(err.getvalue(), "")

    def _run(self, d, *argv):
        env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(d / "mapping.json"),
                   PYTHONIOENCODING="utf-8")
        return subprocess.run([sys.executable, str(SYNC), *argv],
                              cwd=str(d), env=env, capture_output=True,
                              text=True, encoding="utf-8")

    def _sandbox(self, d):
        d = Path(d)
        cfg = d / "mapping.json"
        cfg.write_text(json.dumps({
            "defaults": {"common_dir": "src/common"},
            "projects": [{"name": "示例表A", "search_dir": str(d), "prefix": "示例表A",
                          "require_version": False, "src": "src/projects/示例表A",
                          "common_version": "1.0.0"}],
        }, ensure_ascii=False), encoding="utf-8")
        return d

    # -- 1.6 切片内进程助手：解析 agent 参数、构造最小上下文、走唯一 render 出口 ----
    def _agent_parser(self):
        return build_parser(specs_for("agent"), prog="agent.py")

    def _agent_ctx(self, args, cfg: dict, root: Path | None = None) -> AppContext:
        return AppContext(
            tool="agent", args=args, config=cfg, root=root or ROOT,
            json_mode=bool(getattr(args, "json", False)),
            started=time.perf_counter(),
            progress=Progress(tool="agent", enabled=False))

    def _render(self, outcome) -> tuple[int, dict]:
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            code = render(outcome, json_mode=True)
        return code, json.loads(out.getvalue().strip())

    def test_version_json_has_envelope(self):
        """version --json 必须出 sync.version 信封，不能只打一行人话。"""
        with tempfile.TemporaryDirectory() as d:
            d = self._sandbox(d)
            r = self._run(d, "version", "--json")
            out = r.stdout.strip()
            self.assertEqual(len(out.splitlines()), 1, f"stdout 不是单行：{r.stdout!r}")
            payload = json.loads(out)
            self.assertEqual(payload["action"], "sync.version")
            self.assertIn("details", payload)
            self.assertIn("stats", payload)

    def test_version_json_warnings_collected_into_stats(self):
        """缺 COMMON_VERSION 的警告：JSON 下不写 stderr，收口进 stats.warnings。"""
        with tempfile.TemporaryDirectory() as d:
            d = self._sandbox(d)
            r = self._run(d, "version", "--json")
            out = r.stdout.strip()
            self.assertEqual(len(out.splitlines()), 1, f"stdout 不止一行：{r.stdout!r}")
            payload = json.loads(out)
            self.assertEqual(payload["action"], "sync.version")
            self.assertNotIn("[警告]", r.stderr, f"stderr 漏了警告：{r.stderr!r}")
            warnings = payload["stats"].get("warnings")
            self.assertIsInstance(warnings, list, "stats.warnings 必须是字符串数组")
            self.assertTrue(warnings, "缺 COMMON_VERSION 时 stats.warnings 不能为空")
            self.assertTrue(any("COMMON_VERSION" in w for w in warnings),
                              f"警告文本没收口：{warnings!r}")

    def test_version_human_warnings_still_to_stderr(self):
        """非 JSON 模式行为不变：警告仍打 stderr，stdout 不含警告文本。"""
        with tempfile.TemporaryDirectory() as d:
            d = self._sandbox(d)
            r = self._run(d, "version")
            self.assertIn("[警告]", r.stderr, "人话模式警告必须走 stderr")
            self.assertIn("COMMON_VERSION", r.stderr)
            self.assertNotIn("COMMON_VERSION", r.stdout, "警告文本不能混进 stdout")

    def test_legacy_config_warning_collected_in_json(self):
        """base.warn 通道（配置放工具包根）：JSON 下同样收口进 stats.warnings。"""
        legacy = PKG / "mapping.json"
        self.assertFalse(legacy.exists(), "基线包根不应预置 mapping.json")
        legacy.write_text(json.dumps(
            {"projects": [{"name": "示例表A", "search_dir": tempfile.gettempdir(),
                           "prefix": "示例表A", "require_version": False}]},
            ensure_ascii=False), encoding="utf-8")
        self.addCleanup(lambda: legacy.unlink(missing_ok=True))
        env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(legacy),
                   PYTHONIOENCODING="utf-8")
        r = subprocess.run([sys.executable, str(SYNC), "list", "--json"],
                           cwd=str(PKG), env=env, capture_output=True,
                           text=True, encoding="utf-8")
        out = r.stdout.strip()
        self.assertEqual(len(out.splitlines()), 1, f"stdout 不止一行：{r.stdout!r}")
        payload = json.loads(out)
        self.assertNotIn("[警告]", r.stderr, f"stderr 漏了配置位置警告：{r.stderr!r}")
        warnings = payload["stats"].get("warnings")
        self.assertIsInstance(warnings, list)
        self.assertTrue(any("配置放在工具包根" in w for w in warnings),
                        f"配置位置警告没收口：{warnings!r}")

    def test_every_command_but_watch_has_json(self):
        """每个动词都要自带 --json；watch 是常驻命令，必须没有 --json。"""
        verbs = ("list", "doctor", "version", "vbamod", "pull", "pull-open", "push",
                 "run", "shells", "align", "promote", "gbk-export", "clean-backup",
                 "init", "plan")
        for verb in verbs:
            r = subprocess.run([sys.executable, str(SYNC), verb, "--help"],
                               capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(r.returncode, 0, f"{verb} --help 失败：{r.stderr}")
            self.assertIn("--json", r.stdout, f"{verb} 没有 --json")
            self.assertIn("--apply", r.stdout, f"{verb} 没有 --apply") if verb not in (
                "list", "doctor", "version", "vbamod", "plan") else None
        r = subprocess.run([sys.executable, str(SYNC), "watch", "--help"],
                           capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(r.returncode, 0)
        self.assertNotIn("--json", r.stdout)

    def test_write_verbs_preview_by_default(self):
        """写盘动词不带 --apply 时只预览：push 不打 Excel、不写文件。"""
        with tempfile.TemporaryDirectory() as d:
            d = self._sandbox(d)
            src = d / "src" / "projects" / "示例表A"
            src.mkdir(parents=True)
            common = d / "src" / "common"
            common.mkdir(parents=True)
            (common / "公共.bas").write_text(
                'Attribute VB_Name = "公共"\nOption Explicit\n'
                'Public Function AddOne(ByVal x As Long) As Long\n    AddOne = x + 1\nEnd Function\n',
                encoding="utf-8")
            (src / "示例模块.bas").write_text(
                'Attribute VB_Name = "示例模块"\nOption Explicit\n'
                'Public Sub Hello()\nEnd Sub\n', encoding="utf-8")
            # 占位表壳：验证“有目标时预览照常通过、不连 Excel、不写文件”
            (d / "示例表A V100.xlsm").write_bytes(b"")
            r = self._run(d, "push", "--project", "示例表A", "--json")
            payload = json.loads(r.stdout.strip())
            self.assertTrue(payload["ok"])
            self.assertTrue(payload["stats"]["dry"])
            self.assertEqual(r.returncode, 0)

    def test_push_dry_missing_target_tolerates_and_exits_zero(self):
        """dry + 目标 xlsm 缺失：与 1.5.6 一致——预览容忍缺表，只跑源码预检，exit 0。"""
        with tempfile.TemporaryDirectory() as d:
            d = self._sandbox(d)
            src = d / "src" / "projects" / "示例表A"
            src.mkdir(parents=True)
            common = d / "src" / "common"
            common.mkdir(parents=True)
            (common / "公共.bas").write_text(
                'Attribute VB_Name = "公共"\nOption Explicit\n'
                'Public Function AddOne(ByVal x As Long) As Long\n    AddOne = x + 1\nEnd Function\n',
                encoding="utf-8")
            (src / "示例模块.bas").write_text(
                'Attribute VB_Name = "示例模块"\nOption Explicit\n'
                'Public Sub Hello()\nEnd Sub\n', encoding="utf-8")
            # 故意不放占位 xlsm：源码预检可过，预览不因缺表而失败
            r = self._run(d, "push", "--project", "示例表A", "--json")
            payload = json.loads(r.stdout.strip())
            self.assertEqual(r.returncode, 0)
            self.assertTrue(payload["ok"])
            self.assertTrue(payload["stats"]["dry"])
            self.assertNotEqual(
                (payload["errors"] or [{}])[0].get("code"), "TARGET_NOT_FOUND")

    def test_push_all_json_has_project_rows(self):
        """批量写回 JSON 必须每个项目一条，不能把 stats 丢在 confirm 参数里。"""
        with tempfile.TemporaryDirectory() as d:
            d = self._sandbox(d)
            r = self._run(d, "push", "--all-project", "--json")
            payload = json.loads(r.stdout.strip())
            self.assertEqual(r.returncode, 0)
            self.assertTrue(payload["details"], "批量 push 的 details 不能为空")
            self.assertEqual(payload["details"][0]["name"], "示例表A")
            self.assertIn("status", payload["details"][0])

    def test_shells_preview_marks_dry(self):
        with tempfile.TemporaryDirectory() as d:
            d = self._sandbox(d)
            r = self._run(d, "shells", "--json")
            payload = json.loads(r.stdout.strip())
            self.assertEqual(r.returncode, 0)
            self.assertTrue(payload["stats"]["dry"], "预览必须能看出没有落盘")
            self.assertEqual(payload["action"], "sync.shells")

    def test_clean_backup_json_has_numbers(self):
        with tempfile.TemporaryDirectory() as d:
            d = self._sandbox(d)
            r = self._run(d, "clean-backup", "--json")
            payload = json.loads(r.stdout.strip())
            self.assertEqual(payload["action"], "sync.clean-backup")
            for key in ("applied", "freed_bytes", "kept", "kept_bytes", "deleted_batches"):
                self.assertIn(key, payload["stats"])
            self.assertEqual(r.returncode, 0)

    def test_promote_error_goes_to_errors(self):
        with tempfile.TemporaryDirectory() as d:
            d = self._sandbox(d)
            (d / "src" / "projects" / "示例表A").mkdir(parents=True)
            r = self._run(d, "promote", "--project", "示例表A", "--module", "不存在的模块", "--json")
            payload = json.loads(r.stdout.strip())
            self.assertFalse(payload["ok"])
            self.assertTrue(payload["errors"], "失败时 errors 必须非空，agent 才知道原因")
            self.assertIn("message", payload["errors"][0])
            self.assertEqual(r.returncode, 1)


    def test_audit_out_matches_stdout_envelope(self):
        """--out 落盘必须带与 stdout 相同的 warnings，不能少 stats.warnings。"""
        with tempfile.TemporaryDirectory() as d:
            d = self._sandbox(d)
            out = d / "audit.json"
            env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(d / "mapping.json"),
                       PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
            r = subprocess.run(
                [sys.executable, str(AUDIT), "common", "--json", "--no-input",
                 "--out", str(out)],
                cwd=str(d), env=env, capture_output=True, text=True,
                encoding="utf-8")
            self.assertEqual(r.returncode, 3, r.stderr)
            stdout_payload = json.loads(r.stdout.strip())
            file_payload = json.loads(out.read_text(encoding="utf-8"))
            self.assertEqual(stdout_payload["stats"]["warnings"],
                             file_payload["stats"]["warnings"])
            self.assertEqual(stdout_payload["exit_code"], file_payload["exit_code"])

    def test_audit_subcommands_expose_json(self):
        """audit 六个视角都必须提供 --json 与 --out。"""
        for verb in ("members", "common", "project", "structure", "indicators", "overview"):
            r = subprocess.run([sys.executable, str(AUDIT), verb, "--help"],
                               capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(r.returncode, 0, f"audit {verb} --help 失败：{r.stderr}")
            self.assertIn("--json", r.stdout, f"audit {verb} 没有 --json")
            self.assertIn("--out", r.stdout, f"audit {verb} 没有 --out")

    def test_audit_usage_error_exits_2_without_json(self):
        """缺参数是用法错：只出人话、退出 2，不能 NameError 或输出 JSON。"""
        r = subprocess.run([sys.executable, str(AUDIT), "members", "--json"],
                           capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(r.returncode, 2, f"缺参数应退出 2：stdout={r.stdout!r} stderr={r.stderr!r}")
        self.assertFalse(r.stdout.strip().startswith("{"), "用法错不应输出 JSON 信封")

    def test_audit_overview_stdout_is_single_json_line(self):
        """沙箱 audit overview：人话/表格必须走 stderr，stdout 只能一行 JSON。"""
        with tempfile.TemporaryDirectory() as d:
            d = self._sandbox(d)
            src = d / "src" / "projects" / "示例表A"
            src.mkdir(parents=True)
            (src / "示例模块.bas").write_text(
                'Attribute VB_Name = "示例模块"\nOption Explicit\n'
                'Public Sub Hello()\nEnd Sub\n', encoding="utf-8")
            env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(d / "mapping.json"),
                       PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
            r = subprocess.run([sys.executable, str(AUDIT), "overview", "--json",
                                "--no-input"], cwd=str(d), env=env,
                               capture_output=True, text=True, encoding="utf-8",
                               timeout=60)
            out = r.stdout.strip()
            self.assertEqual(len(out.splitlines()), 1, f"audit overview stdout 不止一行：{r.stdout[:500]!r}")
            payload = json.loads(out)
            self.assertEqual(payload["action"], "audit.overview")
            self.assertIn("exit_code", payload)
            self.assertEqual(r.returncode, payload["exit_code"])

    def test_agent_exposes_json(self):
        r = subprocess.run([sys.executable, str(AGENT), "probe", "--help"],
                           capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(r.returncode, 0, f"agent probe --help 失败：{r.stderr}")
        self.assertIn("--json", r.stdout, "agent 没有全局 --json")

    def test_run_cli_unexpected_is_friendly(self):
        """非 JSON 未知异常不能裸崩：traceback 进 stderr，再给一句人话和退出 1。"""
        from toolkit.cli import run_cli
        err = io.StringIO()
        with contextlib.redirect_stderr(err):
            code = run_cli(lambda: (_ for _ in ()).throw(RuntimeError("boom")))
        text = err.getvalue()
        self.assertEqual(code, 1)
        self.assertIn("Traceback", text)
        self.assertIn("未预期异常", text)

    def test_audit_failure_envelope_has_exit_code(self):
        """audit 派发层抛 Failure 也必须遵守信封契约：带 exit_code。"""
        import audit
        from toolkit.contract import Failure
        from toolkit.base import set_json_mode
        old = audit.dispatch

        def _boom(specs, args, ctx):
            raise Failure("boom", code="BOOM", stage="parse")

        audit.dispatch = _boom
        self.addCleanup(setattr, audit, "dispatch", old)
        self.addCleanup(set_json_mode, False)
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            code = audit.main(["overview", "--json", "--no-input"])
        payload = json.loads(out.getvalue().strip())
        self.assertEqual(code, 1)
        self.assertEqual(payload["exit_code"], 1)
        self.assertEqual(payload["errors"][0]["code"], "BOOM")

    def test_audit_unexpected_envelope_has_exit_code_and_hint(self):
        import audit
        from toolkit.base import set_json_mode
        old = audit.dispatch

        def _boom(specs, args, ctx):
            raise RuntimeError("boom")

        audit.dispatch = _boom
        self.addCleanup(setattr, audit, "dispatch", old)
        self.addCleanup(set_json_mode, False)
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            code = audit.main(["overview", "--json", "--no-input"])
        payload = json.loads(out.getvalue().strip())
        self.assertEqual(code, 1)
        self.assertEqual(payload["exit_code"], 1)
        self.assertEqual(payload["errors"][0]["code"], "UNEXPECTED")
        self.assertIn("维护者", payload["errors"][0]["hint"])

    def test_probe_com_context_is_failure(self):
        """Excel 开着但 COM 不接话时，probe 不能 ok=true/exit 0。"""
        from features import agent_io
        from toolkit.base import Failure, set_json_mode
        cfg = {"defaults": {}, "projects": [{"name": "示例表", "search_dir": ""}]}
        args = self._agent_parser().parse_args(["probe", "--project", "示例表", "--json"])
        ctx = self._agent_ctx(args, cfg)
        old = agent_io.excel_gateway.attach_running_excel
        agent_io.excel_gateway.attach_running_excel = lambda: (_ for _ in ()).throw(
            Failure("COM 不响应", code="COM_CONTEXT", stage="attach",
                    hint="先关掉 Excel 的弹窗/VBE 窗口"))
        self.addCleanup(setattr, agent_io.excel_gateway, "attach_running_excel", old)
        self.addCleanup(set_json_mode, False)
        outcome = agent_io.probe(ctx)
        code, payload = self._render(outcome)
        self.assertEqual(code, 1)
        self.assertEqual(payload["ok"], False)
        self.assertEqual(payload["exit_code"], 1)
        self.assertEqual(payload["errors"][0]["code"], "COM_CONTEXT")
        self.assertEqual(payload["stats"]["excel"], "com_context")
        self.assertIn("弹窗", payload["errors"][0]["hint"])

    def test_vbe_error_dialog_detects_modal(self):
        """纯函数分类能认出 VBE 的 #32770 编译错误弹窗，且按目标 PID 过滤。"""
        from toolkit.dialogs import KIND_COMPILE_ERROR, WindowInfo, detect_dialog
        wins = [WindowInfo(
            hwnd=1, pid=123, title="Microsoft Visual Basic for Applications",
            class_name="#32770",
            text="编译错误：# ElseIf、# Else 或 # EndIf 的前面必须为 # If 子句")]
        match = detect_dialog(wins, target_pid=123)
        self.assertIsNotNone(match)
        self.assertEqual(match.kind, KIND_COMPILE_ERROR)
        self.assertIn("编译错误", match.text)
        self.assertIsNone(detect_dialog(wins, target_pid=456))

    def test_com_error_codes_not_misclassified(self):
        """0x8000FFFF/0x8001010E 不能当忙/断开；0x80010001 才是 RPC_E_CALL_REJECTED。"""
        from toolkit.excel import classify_com_failure
        self.assertEqual(classify_com_failure("-2147418111"), "EXCEL_BUSY")
        self.assertEqual(classify_com_failure("80010001"), "EXCEL_BUSY")
        self.assertEqual(classify_com_failure("-2147418113"), "")
        self.assertEqual(classify_com_failure("-2147417842"), "")
        self.assertEqual(classify_com_failure("-2147023174"), "EXCEL_CLOSED")

    def test_macro_run_error_classified(self):
        """缺宏/禁用宏统一收口 RUN_FAILED（不恢复旧分流码）；区分信息只落 hint，两条文案必须可区分。"""
        from toolkit.macro import MacroRunner
        runner = MacroRunner(None)
        missing = "无法运行“'看门狗验收.xlsm'!测试表格工具.运行全部断言”宏。可能是因为该宏在此工作簿中不可用"
        disabled = "无法运行宏，因为所有的宏都被禁用。"
        ambiguous = "无法运行宏。可能是因为该宏在此工作簿中不可用，或者所有的宏都被禁用。"

        missing_code, missing_stage, missing_hint = runner.classify(
            missing, macro="测试表格工具.运行全部断言")
        disabled_code, disabled_stage, disabled_hint = runner.classify(
            disabled, macro="测试表格工具.运行全部断言")

        self.assertEqual(missing_code, "RUN_FAILED")
        self.assertEqual(disabled_code, "RUN_FAILED")
        self.assertEqual(missing_stage, "run")
        self.assertEqual(disabled_stage, "run")
        # 不恢复 MACRO_NOT_AVAILABLE / MACRO_DISABLED 两个旧分流码
        self.assertNotEqual(missing_code, "MACRO_NOT_AVAILABLE")
        self.assertNotEqual(disabled_code, "MACRO_DISABLED")

        # 两条文案可区分：缺宏写清宏名/模块名与编译问题，禁用写清信任位置/宏安全设置
        self.assertNotEqual(missing_hint, disabled_hint)
        self.assertIn("测试表格工具.运行全部断言", missing_hint)
        self.assertIn("编译", missing_hint)
        self.assertIn("信任", disabled_hint)

        # 歧义句两类线索都在，仍收口 RUN_FAILED，hint 带上本次宏名
        amb_code, _, amb_hint = runner.classify(ambiguous, macro="测试表格工具.运行全部断言")
        self.assertEqual(amb_code, "RUN_FAILED")
        self.assertIn("测试表格工具.运行全部断言", amb_hint)

    def test_vbe_generic_dialog_is_not_compile_error(self):
        """通用 VBE 模态框不能因为标题像 Visual Basic 就报编译错误。"""
        from toolkit.dialogs import KIND_VBE_DIALOG, WindowInfo, detect_dialog
        wins = [WindowInfo(
            hwnd=1, pid=123, title="Microsoft Visual Basic for Applications",
            class_name="#32770", text="工程被锁定，是否解锁？")]
        match = detect_dialog(wins, target_pid=123)
        self.assertIsNotNone(match)
        self.assertEqual(match.kind, KIND_VBE_DIALOG)
        self.assertEqual(match.title, "Microsoft Visual Basic for Applications")

    def test_assert_target_not_active_has_code(self):
        """切不到活动簿时也要有 TARGET_NOT_ACTIVE，不能只甩一句人话。"""
        import types
        from features import agent_patch
        from toolkit.base import Failure

        class FakeWB:
            Name = "目标.xlsm"
            def Activate(self):
                return None

        class FakeXL:
            ActiveWorkbook = types.SimpleNamespace(Name="别的.xlsm")

        with tempfile.TemporaryDirectory() as d:
            xlsm = Path(d) / "示例表 V100.xlsm"
            xlsm.write_bytes(b"")
            cfg = {"defaults": {},
                   "projects": [{"name": "示例表", "file": str(xlsm), "search_dir": str(d)}]}
            args = self._agent_parser().parse_args(["assert", "--project", "示例表"])
            ctx = self._agent_ctx(args, cfg, root=Path(d))
            old_attach = agent_patch.excel_gateway.attach_running_excel
            old_find = agent_patch.excel_gateway.find_open_workbook
            agent_patch.excel_gateway.attach_running_excel = (
                lambda path, *, target_hint="", target_suggested="": FakeXL())
            agent_patch.excel_gateway.find_open_workbook = lambda xl, path: (FakeWB(), [])
            self.addCleanup(setattr, agent_patch.excel_gateway, "attach_running_excel", old_attach)
            self.addCleanup(setattr, agent_patch.excel_gateway, "find_open_workbook", old_find)
            with self.assertRaises(Failure) as ctx2:
                agent_patch.handle_assert(ctx)
            self.assertEqual(ctx2.exception.code, "TARGET_NOT_ACTIVE")

    def test_book_open_makes_excel_visible(self):
        """book open 不能开在隐藏实例里，必须把窗口切到可见。"""
        from features import agent_io

        class FakeXL:
            Visible = False
            WindowState = -4140
            Hwnd = 1

        class FakeWB:
            FullName = "示例表.xlsm"
            def Activate(self):
                return None

        xl = FakeXL()
        cfg = {"defaults": {}, "projects": [{"name": "示例表"}]}
        args = self._agent_parser().parse_args(
            ["book", "--action", "open", "--apply", "--project", "示例表"])
        ctx = self._agent_ctx(args, cfg)
        old_resolve = agent_io.resolve_target
        old_new = agent_io.excel_gateway.new_excel
        old_open = agent_io.excel_gateway.open_disk_workbook
        old_pid = agent_io._excel_pid
        agent_io.resolve_target = lambda proj: (Path("示例表.xlsm"), "按配置路径")

        def fake_new(hidden=False, *, force_new=False):
            # book open 必须自己起可见实例，不能借用户会话改全局开关
            self.assertFalse(hidden)
            self.assertTrue(force_new)
            xl.Visible = True
            xl.WindowState = -4143
            return xl

        agent_io.excel_gateway.new_excel = fake_new
        agent_io.excel_gateway.open_disk_workbook = lambda app, path, read_only=False: FakeWB()
        agent_io._excel_pid = lambda app: 123
        for name, val in (("resolve_target", old_resolve), ("new_excel", old_new),
                          ("open_disk_workbook", old_open), ("_excel_pid", old_pid)):
            self.addCleanup(setattr, agent_io, name, val)
        outcome = agent_io.book(ctx)
        self.assertEqual(outcome.exit_code, 0)
        self.assertTrue(xl.Visible)
        self.assertEqual(xl.WindowState, -4143)

    def _assert_timeout_run(self, *, alive, dialog):
        """跑一条 timed_out=True 的断言宏，返回超时二次分类抛出的 Failure。"""
        from features import agent_patch
        from toolkit.base import Failure
        from toolkit.macro import MacroResult

        class _Runner:
            def run(self, macro, args=None, *, isolate=False, save=False,
                    log_optional=True):
                return MacroResult(macro=macro, timed_out=True)

        class FakeWB:
            Name = "目标.xlsm"
            def Activate(self):
                return None

        class FakeXL:
            Hwnd = 1
            ActiveWorkbook = FakeWB()

        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        root = Path(tmp.name)
        xlsm = root / "示例表 V100.xlsm"
        xlsm.write_bytes(b"")
        cfg = {"defaults": {},
               "projects": [{"name": "示例表", "file": str(xlsm), "search_dir": str(root)}]}
        args = self._agent_parser().parse_args(["assert", "--project", "示例表"])
        ctx = self._agent_ctx(args, cfg, root=root)
        ctx.runner = _Runner()

        old_attach = agent_patch.excel_gateway.attach_running_excel
        old_find = agent_patch.excel_gateway.find_open_workbook
        old_pid = agent_patch._pid_from_xl
        old_alive = agent_patch._probe_pid_alive
        old_dialog = agent_patch._probe_dialog
        agent_patch.excel_gateway.attach_running_excel = (
            lambda path, *, target_hint="", target_suggested="": FakeXL())
        agent_patch.excel_gateway.find_open_workbook = lambda xl, path: (FakeWB(), [])
        agent_patch._pid_from_xl = lambda xl: 123
        agent_patch._probe_pid_alive = lambda pid: alive
        agent_patch._probe_dialog = lambda pid: dialog
        self.addCleanup(setattr, agent_patch.excel_gateway, "attach_running_excel", old_attach)
        self.addCleanup(setattr, agent_patch.excel_gateway, "find_open_workbook", old_find)
        self.addCleanup(setattr, agent_patch, "_pid_from_xl", old_pid)
        self.addCleanup(setattr, agent_patch, "_probe_pid_alive", old_alive)
        self.addCleanup(setattr, agent_patch, "_probe_dialog", old_dialog)

        with self.assertRaises(Failure) as cm:
            agent_patch.handle_assert(ctx)
        return cm.exception

    def test_assert_timeout_excel_closed_has_code(self):
        """超时 + PID 已死 → EXCEL_CLOSED（stage=assert），优先于超时码。"""
        err = self._assert_timeout_run(alive=False, dialog=None)
        self.assertEqual(err.code, "EXCEL_CLOSED")
        self.assertEqual(err.stage, "assert")

    def test_assert_timeout_alive_no_dialog_is_timeout(self):
        """超时 + PID 存活 + 认不出弹窗 → 回退 WATCHDOG_TIMEOUT（stage=assert）。"""
        err = self._assert_timeout_run(alive=True, dialog=None)
        self.assertEqual(err.code, "WATCHDOG_TIMEOUT")
        self.assertEqual(err.stage, "assert")

    def test_assert_timeout_compile_error_is_vbe_error(self):
        """超时 + 存活 + 编译错误框 → VBE_COMPILE_ERROR（stage=vbe），hint 点名标题正文。"""
        from toolkit.dialogs import DialogMatch, KIND_COMPILE_ERROR
        dialog = DialogMatch(kind=KIND_COMPILE_ERROR, hwnd=2, pid=123,
                             title="编译错误：语法错误", text="# Else 没有对应的 # If")
        err = self._assert_timeout_run(alive=True, dialog=dialog)
        self.assertEqual(err.code, "VBE_COMPILE_ERROR")
        self.assertEqual(err.stage, "vbe")
        self.assertIn("编译错误：语法错误", err.hint)
        self.assertIn("# Else", err.hint)

    def test_project_target_missing_has_code(self):
        """项目名对但 mapping 指不到 xlsm 时要报 TARGET_NOT_FOUND。"""
        from features import agent_patch
        from toolkit.base import Failure
        cfg = {"defaults": {}, "projects": [{"name": "示例表", "search_dir": ""}]}
        args = SimpleNamespace(project="示例表")
        ctx = self._agent_ctx(args, cfg)
        with self.assertRaises(Failure) as ctx2:
            agent_patch._pick(ctx, "示例表")
        self.assertEqual(ctx2.exception.code, "TARGET_NOT_FOUND")
        self.assertEqual(ctx2.exception.stage, "config")

    def test_com_call_excel_closed_has_code(self):
        """com_call 遇到 RPC 断开不能再返回空错误码。"""
        from toolkit.excel import com_call
        from toolkit.base import Failure
        def boom():
            raise Exception("(-2147023174, 'The RPC server is unavailable.', None, None)")
        with self.assertRaises(Failure) as ctx:
            com_call(boom)
        self.assertEqual(ctx.exception.code, "EXCEL_CLOSED")
        self.assertEqual(ctx.exception.stage, "attach")

    def test_log_unreadable_has_code(self):
        """日志被占用/读不了时不能落到 UNEXPECTED。"""
        from features import agent_patch
        from toolkit.base import Failure
        with tempfile.TemporaryDirectory() as d:
            with self.assertRaises(Failure) as ctx2:
                agent_patch._read_log_text(Path(d))
        self.assertEqual(ctx2.exception.code, "ASSERT_LOG_UNREADABLE")
        self.assertEqual(ctx2.exception.stage, "assert")

    def test_sync_unknown_project_has_code(self):
        """sync.push 的项目解析失败也要带 code/stage。"""
        from features.sync_push import _find_project
        from toolkit.base import Failure
        with self.assertRaises(Failure) as ctx2:
            _find_project({"projects": [{"name": "示例表"}]}, "不存在")
        self.assertEqual(ctx2.exception.code, "PROJECT_NOT_FOUND")
        self.assertEqual(ctx2.exception.stage, "resolve")

    def test_audit_ignore_manual_accel(self):
        """audit indicators 允许按分组忽略 manual_accel。"""
        from features.audit_engine import AuditEngine
        with tempfile.TemporaryDirectory() as d:
            src = Path(d) / "src" / "projects" / "示例表A"
            src.mkdir(parents=True)
            (src / "示例模块.bas").write_text(
                'Attribute VB_Name = "示例模块"\n'
                "' 测试手写加速\n"
                'Public Sub 手动关闭刷新()\n'
                '    Application.ScreenUpdating = False\n'
                'End Sub\n', encoding="utf-8")
            cfg = {"defaults": {},
                   "projects": [{"name": "示例表A", "src": "src/projects/示例表A"}]}
            engine = AuditEngine(Path(d), cfg)

            rows = engine.indicator_rows()
            h05a = next(r for r in rows if r["id"] == "H05a")
            self.assertEqual(h05a["count"], 1, h05a)

            rows = engine.indicator_rows(ignore=["manual_accel"])
            h05a = next(r for r in rows if r["id"] == "H05a")
            self.assertEqual(h05a["state"], "通过")
            self.assertEqual(h05a["count"], 0)

    def test_unknown_project_has_code(self):
        """项目名写错不能只返回空 error_code。"""
        from features import agent_patch
        from toolkit.base import Failure
        cfg = {"defaults": {}, "projects": [{"name": "示例表", "search_dir": ""}]}
        args = SimpleNamespace(project="不存在")
        ctx = self._agent_ctx(args, cfg)
        with self.assertRaises(Failure) as ctx2:
            agent_patch._pick(ctx, "不存在")
        self.assertEqual(ctx2.exception.code, "PROJECT_NOT_FOUND")

    def test_excel_modal_dialog_detected(self):
        """MsgBox/运行时错误框要能被识别为 Excel 模态框，不能只靠 VBE 弹窗。"""
        from toolkit.dialogs import KIND_MODAL_DIALOG, WindowInfo, detect_dialog
        wins = [WindowInfo(hwnd=1, pid=123, title="Microsoft Excel",
                           class_name="#32770", text="测试弹窗")]
        match = detect_dialog(wins, target_pid=123)
        self.assertIsNotNone(match)
        self.assertEqual(match.kind, KIND_MODAL_DIALOG)
        self.assertEqual((match.title, match.text), ("Microsoft Excel", "测试弹窗"))

    def test_probe_strict_no_excel_is_failure(self):
        """--strict 时 Excel 没开也算健康门禁失败，避免磁盘探针假绿。"""
        from features import agent_io
        from toolkit.base import Failure, set_json_mode
        cfg = {"defaults": {}, "projects": [{"name": "示例表", "search_dir": ""}]}
        args = self._agent_parser().parse_args(
            ["probe", "--project", "示例表", "--json", "--strict"])
        ctx = self._agent_ctx(args, cfg)
        old = agent_io.excel_gateway.attach_running_excel
        agent_io.excel_gateway.attach_running_excel = lambda: (_ for _ in ()).throw(
            Failure("没找到 Excel", code="NO_EXCEL", stage="attach"))
        self.addCleanup(setattr, agent_io.excel_gateway, "attach_running_excel", old)
        self.addCleanup(set_json_mode, False)
        outcome = agent_io.probe(ctx)
        code, payload = self._render(outcome)
        self.assertEqual(code, 1)
        self.assertEqual(payload["ok"], False)
        self.assertEqual(payload["exit_code"], 1)
        self.assertEqual(payload["errors"][0]["code"], "NO_EXCEL")
        self.assertEqual(payload["stats"]["excel"], "no_excel")

    @unittest.skip("B5 已复核为接受设计差异：旧 main 把 handler 返回整数 2 合成 ASSERT_FAILED 信封"
                   "（exit 2）的壳层口径已消失；assert 业务失败在新架构走 Outcome 错误对象、退出 1，"
                   "actionable hint 由 test_agent_patch.test_assert_failure_exit_1_with_code 覆盖；"
                   "旧 exit_code=2 断言不恢复。")
    def test_agent_assert_failed_has_actionable_hint(self):
        """assert 读不出计数时，ASSERT_FAILED 要告诉调用方下一步跑什么。"""
        import agent
        from toolkit.base import set_json_mode
        old = agent.cmd_assert
        agent.cmd_assert = lambda args: 2
        self.addCleanup(setattr, agent, "cmd_assert", old)
        self.addCleanup(set_json_mode, False)
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            code = agent.main(["assert", "--project", "示例表", "--apply",
                               "--json", "--no-input"])
        payload = json.loads([l for l in out.getvalue().splitlines() if l.startswith("{")][-1])
        self.assertEqual(code, 2)
        self.assertEqual(payload["errors"][0]["code"], "ASSERT_FAILED")
        self.assertIn("agent logs", payload["errors"][0]["hint"])


if __name__ == "__main__":
    unittest.main()
