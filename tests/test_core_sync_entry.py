# -*- coding: utf-8 -*-
"""01 切片契约：core 四动词（version / list / doctor / scrub）+ sync.py 通用入口壳。

四类覆盖（每个动词）：
  · 成功：exit 0、ok=true；
  · 业务失败：exit 1、errors[].code 有值；
  · 用法错：exit 2（NONE 动词不登记 --apply，argparse 直接拒）；
  · --json 单行纯度：stdout 恰好一行、可 json.loads、含信封七键。
测试只用桩与 tempfile，不启动 Excel、不跑业务表。
"""

import io
import json
import os
import subprocess
import sys
import tempfile
import time
import unittest
from argparse import Namespace
from contextlib import redirect_stdout, redirect_stderr
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
SCRIPTS = PKG / "scripts"
SYNC = SCRIPTS / "sync.py"
sys.path.insert(0, str(SCRIPTS))

from toolkit.context import AppContext, Progress  # noqa: E402
from toolkit.contract import Outcome  # noqa: E402
from toolkit.base import PKG_ROOT, Failure, entry_command, make_envelope  # noqa: E402
import features.core as core  # noqa: E402
from features import sync_interactive  # noqa: E402

ENVELOPE_KEYS = ("ok", "action", "details", "errors", "elapsed_ms", "stats", "exit_code")


def build_ctx(cfg: dict, root: Path) -> AppContext:
    args = Namespace(json=False, apply=False, no_input=False)
    return AppContext(tool="sync", args=args, config=cfg, root=Path(root),
                      json_mode=False, started=time.perf_counter(),
                      progress=Progress(tool="sync", enabled=False))


def run_sync(root: Path, cfg: dict, *argv: str) -> subprocess.CompletedProcess:
    cfg_path = root / "mapping.json"
    cfg_path.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
    env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(cfg_path), PYTHONIOENCODING="utf-8",
               PYTHONUTF8="1")
    return subprocess.run([sys.executable, str(SYNC), *argv],
                          cwd=str(root), env=env, capture_output=True,
                          text=True, encoding="utf-8")


def assert_json_purity(test: unittest.TestCase, r: subprocess.CompletedProcess, action: str) -> dict:
    lines = r.stdout.strip().splitlines()
    test.assertEqual(len(lines), 1, f"stdout 必须恰好一行：{r.stdout}\n{r.stderr}")
    payload = json.loads(lines[0])
    for key in ENVELOPE_KEYS:
        test.assertIn(key, payload, f"信封缺键 {key}")
    test.assertEqual(payload["action"], action)
    test.assertIsInstance(payload["details"], list)
    test.assertIsInstance(payload["errors"], list)
    test.assertIsInstance(payload["stats"], dict)
    return payload


def write_common_version(common_dir: Path, version: str = "1.0.0") -> None:
    common_dir.mkdir(parents=True, exist_ok=True)
    (common_dir / "CommonVersion.bas").write_text(
        f'Public Const COMMON_VERSION As String = "{version}"\n', encoding="utf-8")


class ShellEntryTest(unittest.TestCase):
    """入口壳通用行为：裸跑 2、未知动词 2、--version 打印版本。"""

    def test_bare_prints_help_exits_2(self):
        with tempfile.TemporaryDirectory() as d:
            r = run_sync(Path(d), {"projects": []}, )
            self.assertEqual(r.returncode, 2, r.stderr)
            self.assertIn("动词", r.stdout)

    def test_unknown_verb_exits_2(self):
        with tempfile.TemporaryDirectory() as d:
            r = run_sync(Path(d), {"projects": []}, "no-such-verb")
            self.assertEqual(r.returncode, 2, r.stdout + r.stderr)

    def test_version_flag_prints_version(self):
        with tempfile.TemporaryDirectory() as d:
            r = run_sync(Path(d), {"projects": []}, "--version")
            self.assertEqual(r.returncode, 0, r.stderr)
            self.assertIn("vba-workbench", r.stdout)


class InitBootstrapTest(unittest.TestCase):
    """init 是引导动词：配置还不存在时也必须走到 handler（旧版入口对它免加载配置）。"""

    def test_init_usage_error_not_masked_by_missing_config(self):
        """缺 --dir 应报用法错 exit 2；不能被"配置不存在"掩盖成 exit 1。"""
        with tempfile.TemporaryDirectory() as d:
            missing = Path(d) / "mapping.json"  # 故意不存在
            env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(missing),
                       PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
            r = subprocess.run([sys.executable, str(SYNC), "init",
                                "--project", "示例表A", "--json"],
                               cwd=str(PKG), env=env, capture_output=True,
                               text=True, encoding="utf-8")
            self.assertEqual(r.returncode, 2, r.stdout + r.stderr)
            self.assertIn("USAGE", r.stdout)

    def test_context_build_skips_config_for_init(self):
        """load_config 抛错时 init 仍能拿到上下文（config 为空字典）。"""
        from unittest import mock

        args = Namespace(command="init", json=False, apply=False, no_input=False)
        with mock.patch("toolkit.context.load_config",
                        side_effect=RuntimeError("配置不存在")):
            ctx = AppContext.build("sync", args)
        self.assertEqual(ctx.config, {})


class VersionHandlerTest(unittest.TestCase):
    def _cfg(self, root: Path, locked: str | None) -> dict:
        return {
            "defaults": {"common_dir": str(root / "src" / "common")},
            "projects": [{"name": "示例表A", "src": "src/projects/示例表A",
                          "common_version": locked}],
        }

    def test_version_success(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            write_common_version(root / "src" / "common")
            ctx = build_ctx(self._cfg(root, "1.0.0"), root)
            out = core.version(ctx)
            self.assertTrue(out.ok)
            self.assertEqual(out.exit_code, 0)
            self.assertTrue(out.stats["configured"])
            self.assertEqual(out.stats["mismatched"], 0)
            self.assertTrue(out.details[0]["match"])

    def test_version_mismatch_exit1(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            write_common_version(root / "src" / "common")
            ctx = build_ctx(self._cfg(root, "0.9.0"), root)
            out = core.version(ctx)
            self.assertFalse(out.ok)
            self.assertEqual(out.exit_code, 1)
            self.assertEqual(out.errors[0].code, "VERSION_MISMATCH")
            self.assertFalse(out.details[0]["match"])

    def test_version_rejects_apply_exit2(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            r = run_sync(root, self._cfg(root, "1.0.0"), "version", "--apply")
            self.assertEqual(r.returncode, 2, r.stdout + r.stderr)

    def test_version_json_single_line(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            write_common_version(root / "src" / "common")
            r = run_sync(root, self._cfg(root, "1.0.0"), "version", "--json")
            self.assertEqual(r.returncode, 0, r.stderr)
            payload = assert_json_purity(self, r, "sync.version")
            self.assertEqual(payload["stats"]["mismatched"], 0)

    def test_version_stats_keys_match_baseline(self):
        """stats.results / stats.errors 必须与 1.5.6 一致，不许删键或搬家。"""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            write_common_version(root / "src" / "common")
            r = run_sync(root, self._cfg(root, "0.9.0"), "version", "--json")
            self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
            stats = assert_json_purity(self, r, "sync.version")["stats"]
            self.assertIn("results", stats, "1.5.6 的 stats.results 被删了")
            self.assertEqual(set(stats["results"][0]), {"name", "locked", "match"})
            self.assertEqual(stats["results"][0]["name"], "示例表A")
            self.assertIn("errors", stats, "1.5.6 的 stats.errors 被删了")
            self.assertIsInstance(stats["errors"], list)
            self.assertIsInstance(stats["errors"][0], str)


class ListHandlerTest(unittest.TestCase):
    def _cfg(self, root: Path, prefix: str, has_target: bool) -> dict:
        src = root / "src" / "projects" / prefix
        src.mkdir(parents=True, exist_ok=True)
        if has_target:
            (root / f"{prefix} V1.xlsm").touch()
        return {
            "defaults": {"common_dir": str(root / "src" / "common")},
            "projects": [{"name": prefix, "search_dir": str(root), "prefix": prefix,
                          "require_version": False, "src": str(src)}],
        }

    def test_list_success(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            ctx = build_ctx(self._cfg(root, "甲", True), root)
            out = core.list_projects(ctx)
            self.assertTrue(out.ok)
            self.assertEqual(out.exit_code, 0)
            self.assertEqual(out.stats["total"], 1)
            self.assertEqual(out.stats["resolved"], 1)
            self.assertTrue(out.details[0]["ok"])
            self.assertTrue(out.details[0]["workbook"].endswith("甲 V1.xlsm"))
            self.assertEqual(out.details[0]["common_modules"], "all")

    def test_list_target_not_found_exit1(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            ctx = build_ctx(self._cfg(root, "乙", False), root)
            out = core.list_projects(ctx)
            self.assertFalse(out.ok)
            self.assertEqual(out.exit_code, 1)
            self.assertEqual(out.errors[0].code, "TARGET_NOT_FOUND")
            self.assertEqual(out.errors[0].stage, "resolve")
            self.assertEqual(out.details[0]["workbook"], "")

    def test_list_rejects_apply_exit2(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            r = run_sync(root, self._cfg(root, "甲", True), "list", "--apply")
            self.assertEqual(r.returncode, 2, r.stdout + r.stderr)

    def test_list_json_single_line(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            r = run_sync(root, self._cfg(root, "甲", True), "list", "--json")
            self.assertEqual(r.returncode, 0, r.stderr)
            payload = assert_json_purity(self, r, "sync.list")
            self.assertEqual(payload["stats"]["resolved"], 1)


class DoctorHandlerTest(unittest.TestCase):
    def _cfg(self, root: Path, healthy: bool) -> dict:
        proj = "示例表A"
        src = root / "src" / "projects" / proj
        common = root / "src" / "common"
        if healthy:
            src.mkdir(parents=True, exist_ok=True)
            common.mkdir(parents=True, exist_ok=True)
            (common / "CommonVersion.bas").write_text(
                'Public Const COMMON_VERSION As String = "1.0.0"\n', encoding="utf-8")
            (root / f"{proj} V1.xlsm").touch()
        else:
            # 目标表与源码目录都缺：解析与源码目录两项必 fail
            pass
        return {
            "defaults": {"common_dir": str(common)},
            "projects": [{"name": proj, "search_dir": str(root), "prefix": proj,
                          "require_version": False, "src": str(src)}],
        }

    def test_doctor_success(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            ctx = build_ctx(self._cfg(root, True), root)
            out = core.doctor(ctx)
            self.assertTrue(out.ok, f"失败项：{[e.message for e in out.errors]}")
            self.assertEqual(out.exit_code, 0)
            self.assertEqual(out.stats["failed"], 0)

    def test_doctor_failed_exit1(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            ctx = build_ctx(self._cfg(root, False), root)
            out = core.doctor(ctx)
            self.assertFalse(out.ok)
            self.assertEqual(out.exit_code, 1)
            self.assertGreaterEqual(out.stats["failed"], 1)
            self.assertTrue(out.errors, "失败时 errors 不能为空")
            for e in out.errors:
                self.assertTrue(e.code, "errors[].code 必须有值")

    def _common_check(self, out) -> dict:
        return next(c for c in out.details if c["name"] == "公共库目录")

    def _healthy_cfg(self, root: Path) -> dict:
        proj = "示例表A"
        src = root / "src" / "projects" / proj
        src.mkdir(parents=True, exist_ok=True)
        (root / f"{proj} V1.xlsm").touch()
        return {
            "defaults": {"common_dir": str(root / "src" / "common")},
            "projects": [{"name": proj, "search_dir": str(root), "prefix": proj,
                          "require_version": False, "src": str(src)}],
        }

    def test_doctor_common_dir_counts_nested_modules(self):
        """嵌套子目录里的模块也要计入（promote 落盘口径），不再误报"没有模块"。"""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            common = root / "src" / "common"
            common.mkdir(parents=True, exist_ok=True)
            (common / "顶层.bas").write_text('Attribute VB_Name = "顶层"\n', encoding="utf-8")
            nested = common / "嵌套目录"
            nested.mkdir(parents=True)
            (nested / "深层.bas").write_text('Attribute VB_Name = "深层"\n', encoding="utf-8")
            ctx = build_ctx(self._healthy_cfg(root), root)
            check = self._common_check(core.doctor(ctx))
            self.assertEqual(check["level"], "ok")
            self.assertIn("2 个模块", check["detail"])
            self.assertNotIn("目录存在但没有模块", check["detail"])

    def test_doctor_common_dir_empty_only_dirs_still_warns(self):
        """只有目录、没有模块文件时，仍要正确报"目录存在但没有模块"。"""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            common = root / "src" / "common"
            (common / "空目录").mkdir(parents=True)
            ctx = build_ctx(self._healthy_cfg(root), root)
            check = self._common_check(core.doctor(ctx))
            self.assertEqual(check["level"], "warn")
            self.assertIn("目录存在但没有模块", check["detail"])

    def test_doctor_rejects_apply_exit2(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            r = run_sync(root, self._cfg(root, True), "doctor", "--apply")
            self.assertEqual(r.returncode, 2, r.stdout + r.stderr)

    def test_doctor_json_single_line(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            r = run_sync(root, self._cfg(root, False), "doctor", "--json")
            self.assertIn(r.returncode, (0, 1), r.stderr)
            payload = assert_json_purity(self, r, "sync.doctor")
            self.assertIn("failed", payload["stats"])

    def test_doctor_check_fields_contract(self):
        """doctor 明细固定为 name/level/detail 三项，且 AccessVBOM/包文件/配置在位。"""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            ctx = build_ctx(self._cfg(root, True), root)
            out = core.doctor(ctx)

            self.assertIsInstance(out.stats["failed"], int)
            self.assertIsInstance(out.details, list)
            self.assertTrue(out.details)
            names = set()
            for check in out.details:
                self.assertEqual(set(check.keys()), {"name", "level", "detail"})
                self.assertIn(check["level"], {"ok", "warn", "fail"})
                self.assertTrue(check["name"])
                self.assertTrue(check["detail"])
                names.add(check["name"])

            self.assertIn("包文件齐全", names)
            self.assertIn("配置文件", names)
            self.assertIn("仓库目录可写", names)
            self.assertEqual(sum("AccessVBOM" in n for n in names), 1, names)
            if out.ok:
                self.assertEqual(out.errors, [])

    def test_doctor_default_skips_package_check(self):
        """不带 --verify-package：stats 不含 package 段，本地改动不影响默认 doctor。"""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            ctx = build_ctx(self._cfg(root, True), root)
            out = core.doctor(ctx)
            self.assertNotIn("package", out.stats)

    def test_doctor_verify_package_missing_manifest_exit1(self):
        """源码树/未安装副本无清单：专用码 PACKAGE_MANIFEST_MISSING，不崩、退出 1。"""
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            r = run_sync(root, self._cfg(root, True), "doctor", "--verify-package", "--json")
            self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
            payload = assert_json_purity(self, r, "sync.doctor")
            self.assertEqual(payload["errors"][0]["code"], "PACKAGE_MANIFEST_MISSING")
            self.assertFalse(payload["stats"]["package"]["present"])


class PackageVerifyHelperTest(unittest.TestCase):
    """--verify-package 对账口径直接覆盖：通过 / 篡改 / 缺件 / 多余 / 版本不符 / 缺清单。"""

    def _make_pkg(self, root: Path, *, version="1.6.1", tamper=False,
                  missing=False, extra=False, manifest_version=None) -> None:
        (root / "scripts" / "features").mkdir(parents=True)
        (root / "references").mkdir(parents=True)
        (root / "VERSION").write_text(version, encoding="utf-8")
        (root / "SKILL.md").write_text("# skill", encoding="utf-8")
        (root / "scripts" / "a.py").write_text("print(1)", encoding="utf-8")
        (root / "references" / "b.md").write_text("# doc", encoding="utf-8")
        files = {rel: core._sha256_file(root / rel)
                 for rel in ("scripts/a.py", "references/b.md", "SKILL.md", "VERSION")}
        if tamper:
            (root / "scripts" / "a.py").write_text("print(999)", encoding="utf-8")
        if missing:
            (root / "references" / "b.md").unlink()
        if extra:
            (root / "references" / "c.md").write_text("# extra", encoding="utf-8")
        manifest = {"version": manifest_version or version, "source": "fake-source",
                    "installed_at": "2026-01-01T00:00:00", "files": files}
        (root / core.PACKAGE_MANIFEST_NAME).write_text(
            json.dumps(manifest, ensure_ascii=False), encoding="utf-8")

    def test_verify_package_pass(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            self._make_pkg(root)
            stats, errors = core._verify_package(root)
            self.assertTrue(stats["present"])
            self.assertEqual(stats["checked"], 4)
            self.assertEqual(stats["missing"], [])
            self.assertEqual(stats["mismatch"], [])
            self.assertEqual(stats["extra"], [])
            self.assertTrue(stats["version_match"])
            self.assertEqual(errors, [])

    def test_verify_package_tamper_reports_file_and_code(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            self._make_pkg(root, tamper=True)
            stats, errors = core._verify_package(root)
            self.assertEqual(stats["mismatch"], ["scripts/a.py"])
            self.assertEqual(len(errors), 1)
            self.assertEqual(errors[0].code, "PACKAGE_INTEGRITY")
            self.assertIn("scripts/a.py", str(errors[0]))
            self.assertTrue(errors[0].hint)

    def test_verify_package_missing_file(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            self._make_pkg(root, missing=True)
            stats, errors = core._verify_package(root)
            self.assertEqual(stats["missing"], ["references/b.md"])
            self.assertTrue(any(e.code == "PACKAGE_INTEGRITY" for e in errors))

    def test_verify_package_extra_file(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            self._make_pkg(root, extra=True)
            stats, errors = core._verify_package(root)
            self.assertEqual(stats["extra"], ["references/c.md"])
            self.assertTrue(any(e.code == "PACKAGE_INTEGRITY" for e in errors))

    def test_verify_package_version_mismatch(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            self._make_pkg(root, manifest_version="9.9.9")
            stats, errors = core._verify_package(root)
            self.assertFalse(stats["version_match"])
            self.assertTrue(any(e.code == "PACKAGE_INTEGRITY" for e in errors))

    def test_verify_package_manifest_missing(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            stats, errors = core._verify_package(root)
            self.assertFalse(stats["present"])
            self.assertEqual(len(errors), 1)
            self.assertEqual(errors[0].code, "PACKAGE_MANIFEST_MISSING")


class InterruptSignalTest(unittest.TestCase):
    """派发期间 Ctrl+C：sync 入口必须返回 130、提示「已中断」，不出 JSON 信封。"""

    def _run_interrupt(self, tmp: Path, *argv: str):
        import sync as sync_entry  # noqa: E402
        from toolkit.base import set_json_mode  # noqa: E402

        cfg = tmp / "mapping.json"
        cfg.write_text(json.dumps({"projects": []}, ensure_ascii=False), encoding="utf-8")
        old_env = os.environ.get("VBA_WORKBENCH_CONFIG")
        os.environ["VBA_WORKBENCH_CONFIG"] = str(cfg)
        orig = sync_entry.dispatch

        def _boom(specs, args, ctx):
            raise KeyboardInterrupt

        sync_entry.dispatch = _boom
        buf, err = io.StringIO(), io.StringIO()
        try:
            with redirect_stdout(buf), redirect_stderr(err):
                code = sync_entry.main(list(argv))
        finally:
            sync_entry.dispatch = orig
            set_json_mode(False)
            if old_env is None:
                os.environ.pop("VBA_WORKBENCH_CONFIG", None)
            else:
                os.environ["VBA_WORKBENCH_CONFIG"] = old_env
        return code, buf.getvalue(), err.getvalue()

    def test_sync_keyboard_interrupt_exits_130_without_envelope(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            # 非 JSON：人话含「已中断」，stdout 不出信封
            code, out, err = self._run_interrupt(d, "list")
            self.assertEqual(code, 130, err)
            self.assertIn("已中断", out)
            self.assertFalse(out.strip().startswith("{"))
            # JSON 模式：同样不得吐半截信封，「已中断」走 stderr
            code2, out2, err2 = self._run_interrupt(d, "list", "--json")
            self.assertEqual(code2, 130, err2)
            self.assertFalse(out2.strip(), f"JSON 中断不得出 stdout：{out2!r}")
            self.assertIn("已中断", err2)


class AbsoluteEntryCommandTest(unittest.TestCase):
    """sync 系出口命令：entry_command 拼绝对入口；信封白名单认得它、不降级成 hint；
    本切片四个源文件静态不得再出现包内相对路径命令。"""

    SOURCE_FILES = (
        "scripts/features/sync_read.py",
        "scripts/features/sync_push.py",
        "scripts/features/sync_transform.py",
        "scripts/features/core.py",
    )

    def test_entry_command_uses_interpreter_and_absolute_script(self):
        # ① 命令含本包入口脚本绝对路径，且含当前解释器可执行文件名
        cmd = entry_command("sync.py", "list", "--json")
        script = PKG_ROOT / "scripts" / "sync.py"
        self.assertIn(str(script), cmd)
        self.assertIn(Path(sys.executable).name, cmd)
        self.assertIn("--json", cmd)
        self.assertTrue(Path(sys.executable).is_absolute())
        self.assertTrue(script.is_absolute())

    def test_make_envelope_keeps_absolute_suggested_command(self):
        # ② 绝对入口命令进信封后 suggested_command 非空、不被白名单回退成 hint
        cmd = entry_command("sync.py", "list", "--json")
        payload = make_envelope("sync.list", ok=False, errors=["解析失败"],
                                suggested_command=cmd)
        row = payload["errors"][0]
        self.assertTrue(row["suggested_command"])
        self.assertEqual(row["suggested_command"], cmd)
        self.assertIn(str(PKG_ROOT / "scripts" / "sync.py"), row["suggested_command"])
        self.assertIn(sys.executable, row["suggested_command"])

    def test_source_files_have_no_relative_package_command(self):
        # ③ 静态断言：四个源文件文本里不含 python vba-workbench（含反斜杠写法）
        for rel in self.SOURCE_FILES:
            text = (PKG_ROOT / rel).read_text(encoding="utf-8")
            self.assertNotIn("python vba-workbench", text,
                             f"{rel} 仍含包内相对路径命令")


class SyncInteractiveMenuTest(unittest.TestCase):
    """sync -i 菜单（C 方案）：选动作即预览，回车执行，结束回菜单。"""

    def _cfg(self) -> dict:
        return {"defaults": {"common_dir": "src/common"},
                "projects": [{"name": "示例表A", "src": "src/projects/A"},
                             {"name": "示例表B", "src": "src/projects/B"}]}

    @staticmethod
    def _reader(lines: list[str]):
        it = iter(lines)

        def _read(prompt: str) -> str:
            try:
                return next(it)
            except StopIteration:
                raise EOFError

        return _read

    def test_lists_projects_from_mapping(self):
        writer: list[str] = []
        code = sync_interactive.run_menu(
            config=self._cfg(), root=Path("."),
            reader=self._reader(["0"]), writer=writer.append,
            runner=lambda verb, project, **kw: None)
        self.assertEqual(code, 0)
        text = "\n".join(writer)
        self.assertIn("示例表A", text)
        self.assertIn("示例表B", text)
        self.assertIn("全局工具", text)
        self.assertIn("新建项目", text)
        self.assertIn("【选择项目】", text)
        zero = writer.index("  0. 退出")
        self.assertEqual(writer[zero - 1], "  " + "─" * 52)
        self.assertEqual(writer[zero + 1], "", "0 后面要留空行再出输入提示")

    def test_write_action_previews_then_enter_executes_and_returns_to_menu(self):
        calls: list[tuple[str, str, bool, int | None]] = []
        outcome = Outcome(action="sync.push", ok=True, exit_code=0, human=["写回完成"])

        def runner(verb: str, project: str, *, apply: bool = False,
                   seconds: int | None = None, extra: tuple[str, ...] = ()):
            calls.append((verb, project, apply, seconds))
            return outcome

        writer: list[str] = []
        code = sync_interactive.run_menu(
            config=self._cfg(), root=Path("."),
            reader=self._reader(["1", "4", "", "0", "0"]),
            writer=writer.append, runner=runner)
        self.assertEqual(code, 0)
        self.assertEqual(calls, [("push", "示例表A", False, None),
                                 ("push", "示例表A", True, None)])
        text = "\n".join(writer)
        self.assertIn("预览，未修改任何文件", text)
        self.assertIn("回车 = 执行", text)
        self.assertIn("完成，返回动作菜单", text)

    def test_preview_cancel_with_zero_does_not_execute(self):
        calls: list[tuple[str, str, bool, int | None]] = []
        outcome = Outcome(action="sync.push", ok=True, exit_code=0, human=["预览完成"])

        def runner(verb: str, project: str, *, apply: bool = False,
                   seconds: int | None = None, extra: tuple[str, ...] = ()):
            calls.append((verb, project, apply, seconds))
            return outcome

        writer: list[str] = []
        code = sync_interactive.run_menu(
            config=self._cfg(), root=Path("."),
            reader=self._reader(["1", "4", "0", "0", "0"]),
            writer=writer.append, runner=runner)
        self.assertEqual(code, 0)
        self.assertEqual(calls, [("push", "示例表A", False, None)])

    def test_preview_failure_has_no_execute_prompt_and_returns_to_menu(self):
        calls: list[tuple[str, str, bool, int | None]] = []
        bad = Outcome(action="sync.push", ok=False, exit_code=1,
                      errors=[Failure("源码预检失败", code="SOURCE_PREFLIGHT",
                                      stage="preflight")])

        def runner(verb: str, project: str, *, apply: bool = False,
                   seconds: int | None = None, extra: tuple[str, ...] = ()):
            calls.append((verb, project, apply, seconds))
            return bad

        writer: list[str] = []
        code = sync_interactive.run_menu(
            config=self._cfg(), root=Path("."),
            reader=self._reader(["1", "4", "0", "0"]),
            writer=writer.append, runner=runner)
        self.assertEqual(code, 0)
        self.assertEqual(calls, [("push", "示例表A", False, None)])
        text = "\n".join(writer)
        self.assertIn("预览失败", text)
        self.assertNotIn("回车 = 执行", text)

    def test_readonly_action_runs_without_preview_prompt(self):
        calls: list[tuple[str, str, bool, int | None]] = []
        outcome = Outcome(action="sync.plan", ok=True, exit_code=0, human=["推荐 push"])

        def runner(verb: str, project: str, *, apply: bool = False,
                   seconds: int | None = None, extra: tuple[str, ...] = ()):
            calls.append((verb, project, apply, seconds))
            return outcome

        writer: list[str] = []
        code = sync_interactive.run_menu(
            config=self._cfg(), root=Path("."),
            reader=self._reader(["1", "3", "0", "0"]),
            writer=writer.append, runner=runner)
        self.assertEqual(code, 0)
        self.assertEqual(calls, [("plan", "示例表A", False, None)])
        self.assertIn("只读完成，返回动作菜单", "\n".join(writer))

    def test_global_tools_are_separate_from_project_actions(self):
        calls: list[tuple[str, str, bool, int | None]] = []
        outcome = Outcome(action="sync.list", ok=True, exit_code=0, human=["列表完成"])

        def runner(verb: str, project: str, *, apply: bool = False,
                   seconds: int | None = None, extra: tuple[str, ...] = ()):
            calls.append((verb, project, apply, seconds))
            return outcome

        writer: list[str] = []
        code = sync_interactive.run_menu(
            config=self._cfg(), root=Path("."),
            reader=self._reader(["3", "1", "0", "0"]),
            writer=writer.append, runner=runner)
        self.assertEqual(code, 0)
        self.assertEqual(calls, [("list", "全局工具", False, None)])

    def test_watch_asks_seconds_then_enter_starts(self):
        calls: list[tuple[str, str, bool, int | None]] = []
        outcome = Outcome(action="sync.watch", ok=True, exit_code=0, human=["监听结束"])

        def runner(verb: str, project: str, *, apply: bool = False,
                   seconds: int | None = None, extra: tuple[str, ...] = ()):
            calls.append((verb, project, apply, seconds))
            return outcome

        writer: list[str] = []
        code = sync_interactive.run_menu(
            config=self._cfg(), root=Path("."),
            reader=self._reader(["1", "8", "120", "", "0", "0"]),
            writer=writer.append, runner=runner)
        self.assertEqual(code, 0)
        self.assertEqual(calls, [("watch", "示例表A", True, 120)])
        self.assertIn("回车 = 开始监听", "\n".join(writer))

    def test_cancel_at_action_level_returns_to_projects_then_exits(self):
        calls: list[tuple[str, str]] = []
        writer: list[str] = []
        code = sync_interactive.run_menu(
            config=self._cfg(), root=Path("."),
            reader=self._reader(["1", "0", "0"]), writer=writer.append,
            runner=lambda verb, project, **kw: calls.append((verb, project)))
        self.assertEqual(code, 0)
        self.assertEqual(calls, [], "返回项目列表后退出，不得执行任何 handler")

    def test_no_projects_can_create_project(self):
        writer: list[str] = []
        code = sync_interactive.run_menu(
            config={"projects": []}, root=Path("."),
            reader=self._reader(["0"]), writer=writer.append,
            runner=lambda verb, project, **kw: None)
        self.assertEqual(code, 0)
        text = "\n".join(writer)
        self.assertIn("新建项目", text)
        self.assertIn("全局工具", text)

    def test_new_project_from_dragged_versioned_workbook(self):
        calls: list[tuple[str, str, dict]] = []
        outcome = Outcome(action="sync.init", ok=True, exit_code=0,
                          human=["初始化完成"],
                          details=[{"name": "示例表A",
                                    "search_dir": "D:/表格",
                                    "prefix": "示例表A",
                                    "src": "src/projects/示例表A"}])
        pull_outcome = Outcome(action="sync.pull", ok=True, exit_code=0,
                               human=["抓取完成"])

        def runner(verb: str, project: str, **kw):
            calls.append((verb, project, kw))
            return outcome if verb == "init" else pull_outcome

        with tempfile.TemporaryDirectory() as d:
            book = Path(d) / "示例表A V3.xlsm"
            book.write_bytes(b"")
            writer: list[str] = []
            code = sync_interactive.run_menu(
                config={"projects": []}, root=Path(d),
                reader=self._reader(["2", str(book), "0", "0"]),
                writer=writer.append, runner=runner)

        self.assertEqual(code, 0)
        self.assertEqual(calls[0][0:2], ("init", "示例表A"))
        extra = calls[0][2]["extra"]
        self.assertIn("--dir", extra)
        self.assertIn("--prefix", extra)
        self.assertIn("--allow-outside", extra)
        self.assertNotIn("--file", extra, "带版本号时应按目录取最高版本")
        self.assertEqual(calls[1][0:2], ("pull", "示例表A"))
        self.assertTrue(calls[1][2]["apply"])
        text = "\n".join(writer)
        self.assertIn("已创建项目并抓取源码", text)
        self.assertIn("项目：示例表A", text)

    def test_new_project_without_version_registers_exact_file(self):
        calls: list[tuple[str, str, dict]] = []
        outcome = Outcome(action="sync.init", ok=True, exit_code=0, human=["初始化完成"])

        def runner(verb: str, project: str, **kw):
            calls.append((verb, project, kw))
            return outcome

        with tempfile.TemporaryDirectory() as d:
            book = Path(d) / "示例表B.xlsm"
            book.write_bytes(b"")
            code = sync_interactive.run_menu(
                config={"projects": []}, root=Path(d),
                reader=self._reader(["2", str(book), "0", "0"]),
                writer=lambda _line: None, runner=runner)

        self.assertEqual(code, 0)
        extra = calls[0][2]["extra"]
        self.assertIn("--file", extra)
        self.assertEqual(Path(extra[extra.index("--file") + 1]).resolve(),
                         book.resolve())

    def test_non_tty_shell_exits_1_clean(self):
        """非 TTY 走壳层 ensure_interactive：人话退出 1，不出 traceback（与既有 subprocess 用例同口径）。"""
        import sync as sync_entry
        from toolkit.base import set_non_interactive
        from toolkit.cli import run_cli

        set_non_interactive(True)
        self.addCleanup(set_non_interactive, False)
        buf, err = io.StringIO(), io.StringIO()
        with redirect_stdout(buf), redirect_stderr(err):
            code = run_cli(sync_entry._interactive)
        self.assertEqual(code, 1, err.getvalue())
        self.assertNotIn("Traceback", err.getvalue())
        self.assertIn("没有可交互终端", err.getvalue())


if __name__ == "__main__":
    unittest.main()
