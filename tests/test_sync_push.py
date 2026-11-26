# -*- coding: utf-8 -*-
"""sync_push 切片契约测试：init / push / shells。

测试不起 Excel、不写真业务表；COM 附着只走 adapters.excel_gateway 桩，
写回引擎的低层 COM 访问用内存桩对象替换，临时目录只放 .bas 测试件。
--json 纯度走 toolkit.render 唯一出口。
"""
import contextlib
import io
import json
import sys
import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))

from toolkit.base import ROOT, Failure as BaseFailure  # noqa: E402
from toolkit import excel as excel_toolkit  # noqa: E402
from toolkit.commands import specs_for  # noqa: E402
from toolkit.context import AppContext, Progress  # noqa: E402
from toolkit.contract import Failure  # noqa: E402
from toolkit.registry import build_parser  # noqa: E402
import features.sync_push as sp  # noqa: E402
import features.sync_init as sync_init  # noqa: E402
import features.core as core  # noqa: E402
import features.sync_read as sr  # noqa: E402
import features.run as run_slice  # noqa: E402


COMMON_BAS = (
    'Attribute VB_Name = "公共"\n'
    'Option Explicit\n'
    'Public Function AddOne(ByVal x As Long) As Long\n    AddOne = x + 1\nEnd Function\n'
)
VERSION_BAS = (
    'Attribute VB_Name = "版本"\n'
    'Option Explicit\n'
    'Public Const COMMON_VERSION As String = "1.6.2"\n'
)
PROJ_BAS = (
    'Attribute VB_Name = "模块A"\n'
    'Option Explicit\n'
    'Public Function Foo() As Long\n    Foo = 1\nEnd Function\n'
)


def make_parser():
    return build_parser(specs_for("sync"), prog="sync.py")


def make_ctx(args, cfg):
    return AppContext(
        tool="sync", args=args, config=cfg, root=ROOT,
        json_mode=bool(getattr(args, "json", False)),
        started=time.perf_counter(),
        progress=Progress(tool="sync", enabled=False),
    )


def base_cfg(d: Path) -> dict:
    proj_src = d / "src" / "projects" / "示例表A"
    return {
        "defaults": {"common_dir": str(d / "src" / "common"),
                     "backup_dir": str(d / ".backup"),
                     "import_document_modules": False,
                     "backup_auto_clean": True, "backup_keep": 10,
                     "backup_keep_daily": 1, "backup_min_keep": 3,
                     "backup_protect": []},
        "projects": [{"name": "示例表A", "search_dir": str(d), "prefix": "示例表A",
                      "require_version": False, "src": str(proj_src)}],
    }


def setup_tree(d: Path, *, common: bool = True, proj_code: bool = True) -> tuple[Path, Path]:
    common_dir = d / "src" / "common"
    proj_src = d / "src" / "projects" / "示例表A"
    if common:
        common_dir.mkdir(parents=True, exist_ok=True)
        (common_dir / "公共.bas").write_text(COMMON_BAS, encoding="utf-8")
    if proj_code:
        proj_src.mkdir(parents=True, exist_ok=True)
        (proj_src / "模块A.bas").write_text(PROJ_BAS, encoding="utf-8")
    return common_dir, proj_src


def render_json(outcome) -> dict:
    """走唯一输出出口渲染 JSON，返回解析后的信封与退出码。"""
    from toolkit.render import render
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        code = render(outcome, json_mode=True)
    line = buf.getvalue().strip()
    return json.loads(line), code, line


class CommonSubsetAutoTest(unittest.TestCase):
    """auto 公共模块挑选：引用模块 + 版本常量模块都应进入集合。"""

    def test_sync_push_auto_includes_version_module(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = base_cfg(d)
            common_dir, proj_src = setup_tree(d)
            (common_dir / "版本.bas").write_text(VERSION_BAS, encoding="utf-8")
            (proj_src / "模块B.bas").write_text(
                'Attribute VB_Name = "模块B"\n'
                'Option Explicit\n'
                'Public Function Bar() As Long\n    Bar = AddOne(1)\nEnd Function\n',
                encoding="utf-8")
            proj = cfg["projects"][0] | {"common_modules": "auto"}
            subset, reason = sp._common_subset(proj, cfg["defaults"])
            self.assertIsInstance(subset, set)
            self.assertEqual(subset, {"公共", "版本"})
            self.assertIn("2/2", reason)

    def test_core_resolve_auto_includes_version_module(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = base_cfg(d)
            common_dir, proj_src = setup_tree(d)
            (common_dir / "版本.bas").write_text(VERSION_BAS, encoding="utf-8")
            (proj_src / "模块B.bas").write_text(
                'Attribute VB_Name = "模块B"\n'
                'Option Explicit\n'
                'Public Function Bar() As Long\n    Bar = AddOne(1)\nEnd Function\n',
                encoding="utf-8")
            proj = cfg["projects"][0] | {"common_modules": "auto"}
            subset, reason, ignored = core._resolve_common_subset(proj, cfg["defaults"])
            self.assertIsInstance(subset, set)
            self.assertEqual(subset, {"公共", "版本"})
            self.assertEqual(ignored, [])
            self.assertIn("2/2", reason)

    def test_all_slices_share_one_common_subset_result(self):
        """core / pull / push / run 必须共用同一份公共库挑选结果。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = base_cfg(d)
            common_dir, proj_src = setup_tree(d)
            (common_dir / "版本.bas").write_text(VERSION_BAS, encoding="utf-8")
            (proj_src / "模块B.bas").write_text(
                'Attribute VB_Name = "模块B"\n'
                'Option Explicit\n'
                'Public Function Bar() As Long\n    Bar = AddOne(1)\nEnd Function\n',
                encoding="utf-8")
            proj = cfg["projects"][0] | {"common_modules": "auto"}
            push_set, _ = sp._common_subset(proj, cfg["defaults"])
            pull_set, _ = sr._resolve_common_subset(proj, cfg["defaults"])
            run_set, _ = run_slice._common_subset(proj, cfg["defaults"])
            core_set, _, _ = core._resolve_common_subset(proj, cfg["defaults"])
            expected = {"公共", "版本"}
            self.assertEqual(push_set, expected)
            self.assertEqual(pull_set, expected)
            self.assertEqual(run_set, expected)
            self.assertEqual(core_set, expected)


class ShellsTest(unittest.TestCase):
    """shells：create/clean 两种模式，默认预览。"""

    def test_create_apply_writes_shell(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = base_cfg(d)
            _, proj_src = setup_tree(d)
            args = make_parser().parse_args(["shells", "--apply"])
            outcome = sp.shells(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 0)
            self.assertTrue(outcome.ok)
            self.assertEqual(outcome.stats["created"], 1)
            self.assertFalse(outcome.stats["dry"])
            row = outcome.details[0]
            self.assertEqual(row["name"], "示例表A")
            self.assertEqual(row["action"], "create")
            self.assertEqual(row["count"], 1)
            shell = next(proj_src.rglob("公共.bas"))
            from toolkit.repo import is_shell_file
            self.assertTrue(is_shell_file(shell))

    def test_create_preview_does_not_write(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = base_cfg(d)
            _, proj_src = setup_tree(d)
            args = make_parser().parse_args(["shells"])
            outcome = sp.shells(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 0)
            self.assertTrue(outcome.stats["dry"])
            self.assertEqual(outcome.stats["created"], 1)  # 预览计数照样给出
            from toolkit.repo import is_shell_file
            self.assertFalse(any(is_shell_file(f) for f in proj_src.rglob("*.bas")))

    def test_clean_apply_removes_shell(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = base_cfg(d)
            _, proj_src = setup_tree(d)
            from toolkit.repo import shell_text
            shell = proj_src / "common" / "公共.bas"
            shell.parent.mkdir(parents=True, exist_ok=True)
            shell.write_text(shell_text("src/common/公共.bas", ".bas"), encoding="utf-8")
            args = make_parser().parse_args(["shells", "--mode", "clean", "--apply"])
            outcome = sp.shells(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 0)
            self.assertTrue(outcome.stats["cleaned"])
            self.assertEqual(outcome.details[0]["action"], "clean")
            self.assertEqual(outcome.details[0]["count"], 1)
            self.assertFalse(shell.exists())

    def test_bad_mode_is_usage_error_exit_2(self):
        with self.assertRaises(SystemExit) as cm:
            make_parser().parse_args(["shells", "--mode", "bogus"])
        self.assertEqual(cm.exception.code, 2)

    def test_json_envelope_single_line(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = base_cfg(d)
            setup_tree(d)
            args = make_parser().parse_args(["shells", "--json"])
            outcome = sp.shells(make_ctx(args, cfg))
            payload, code, line = render_json(outcome)
            self.assertEqual(code, 0)
            self.assertEqual(len(line.splitlines()), 1)
            for key in ("ok", "action", "details", "errors", "elapsed_ms", "stats", "exit_code"):
                self.assertIn(key, payload)
            self.assertEqual(payload["action"], "sync.shells")
            self.assertEqual(payload["stats"]["results"], payload["details"])


class InitTest(unittest.TestCase):
    """init：默认预览，--apply 才写 mapping.json 与 src 骨架。"""

    def _args(self, d: Path, cfg_path: Path, *extra: str) -> list[str]:
        return ["init", "--project", "示例表A", "--dir", str(d),
                "--config", str(cfg_path), *extra]

    def test_preview_does_not_write(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg_path = d / "mapping.json"
            args = make_parser().parse_args(self._args(d, cfg_path))
            outcome = sp.init(make_ctx(args, base_cfg(d)))
            self.assertEqual(outcome.exit_code, 0)
            self.assertTrue(outcome.ok)
            self.assertTrue(outcome.stats["dry"])
            self.assertEqual(outcome.stats["created_src"], [])
            self.assertFalse(cfg_path.exists())
            self.assertEqual(outcome.details[0]["name"], "示例表A")
            self.assertEqual(outcome.stats["results"][0]["name"], "示例表A")

    def test_apply_writes_config_and_src(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg_path = d / "mapping.json"
            args = make_parser().parse_args(self._args(d, cfg_path, "--apply"))
            outcome = sp.init(make_ctx(args, base_cfg(d)))
            self.assertEqual(outcome.exit_code, 0)
            self.assertTrue(outcome.ok)
            self.assertFalse(outcome.stats["dry"])
            self.assertTrue(cfg_path.exists())
            self.assertTrue((d / "src" / "common").is_dir())
            self.assertTrue((d / "src" / "projects" / "示例表A").is_dir())
            self.assertEqual(outcome.stats["created_src"],
                             ["src/common", "src/projects/示例表A"])

    def test_invalid_json_is_business_failure(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg_path = d / "mapping.json"
            cfg_path.write_text("{ 不是合法 JSON", encoding="utf-8")
            args = make_parser().parse_args(self._args(d, cfg_path, "--apply"))
            outcome = sp.init(make_ctx(args, base_cfg(d)))
            self.assertEqual(outcome.exit_code, 1)
            self.assertFalse(outcome.ok)
            self.assertEqual(outcome.errors[0].code, "CONFIG_INVALID")
            self.assertEqual(outcome.errors[0].stage, "config")

    def test_missing_project_is_usage_error_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            with self.assertRaises(SystemExit) as cm:
                make_parser().parse_args(["init", "--dir", str(d)])
            self.assertEqual(cm.exception.code, 2)

    def test_missing_dir_is_handler_usage_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg_path = d / "mapping.json"
            args = make_parser().parse_args(
                ["init", "--project", "示例表A", "--config", str(cfg_path)])
            outcome = sp.init(make_ctx(args, base_cfg(d)))
            self.assertEqual(outcome.exit_code, 2)
            self.assertEqual(outcome.errors[0].code, "USAGE")

    def test_json_envelope_single_line(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg_path = d / "mapping.json"
            args = make_parser().parse_args(self._args(d, cfg_path, "--json"))
            outcome = sp.init(make_ctx(args, base_cfg(d)))
            payload, code, line = render_json(outcome)
            self.assertEqual(code, 0)
            self.assertEqual(len(line.splitlines()), 1)
            for key in ("ok", "action", "details", "errors", "elapsed_ms", "stats", "exit_code"):
                self.assertIn(key, payload)
            self.assertEqual(payload["action"], "sync.init")

    def test_lab_apply_uses_generator_seam(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg_path = d / "mapping.json"
            called: list[Path] = []

            def fake_generator(script: Path, target: Path) -> None:
                called.append(target)
                target.write_text("lab workbook placeholder", encoding="utf-8")

            old = sync_init._run_lab_generator
            sync_init._run_lab_generator = fake_generator
            self.addCleanup(setattr, sync_init, "_run_lab_generator", old)
            args = make_parser().parse_args(
                ["init", "--project", "示例表A", "--config", str(cfg_path),
                 "--lab", "--apply"])
            outcome = sp.init(make_ctx(args, base_cfg(d)))
            self.assertEqual(outcome.exit_code, 0)
            self.assertEqual(len(called), 1)
            self.assertTrue(called[0].name.startswith("示例表A V1"))


class _FakeCodeModule:
    def __init__(self, text: str = ""):
        self._text = text

    @property
    def CountOfLines(self) -> int:
        return len(self._text.splitlines())

    def Lines(self, _start: int, _count: int) -> str:
        return self._text


class _FakeComp:
    def __init__(self, name: str, ctype: int, text: str = ""):
        self.Name = name
        self.Type = ctype
        self.CodeModule = _FakeCodeModule(text)


class PushTest(unittest.TestCase):
    """push：预检 → 备份 → 写回 → 回读 →（失败按开关）回滚 → 保存。"""

    GATEWAY_FUNCS = ("attach_running_excel", "find_open_workbook",
                     "new_excel", "open_disk_workbook")

    def setUp(self):
        self._orig_gw = {n: getattr(sp.excel_gateway, n) for n in self.GATEWAY_FUNCS}
        self._orig_seams = {
            "preflight": sp._preflight_or_fail,
            "backup": sp._backup_workbook,
            "writeback": sp._writeback,
            "postcheck": sp._postcheck_sync,
            "restore": sp._restore_module_code,
        }
        self._orig_helpers = {
            "component_snapshot": sp.component_snapshot,
            "snapshot_module_code": sp._snapshot_module_code,
            "find_component": sp.find_component,
            "replace_module_code": sp.replace_module_code,
            "import_component": sp.import_component,
            "remove_module": sp.remove_module,
            "gbk_import_copy": sp.gbk_import_copy,
            "has_vbproject_access": sp.has_vbproject_access,
            "ensure_vbe_access": sp.ensure_vbe_access,
            "vbe_state": sp.vbe_state,
            "backup_copy": sp.backup_copy,
            "prune_backups": sp.prune_backups,
        }

    def tearDown(self):
        for n, fn in self._orig_gw.items():
            setattr(sp.excel_gateway, n, fn)
        for n, fn in self._orig_seams.items():
            key = {"preflight": "_preflight_or_fail", "backup": "_backup_workbook",
                   "writeback": "_writeback", "postcheck": "_postcheck_sync",
                   "restore": "_restore_module_code"}[n]
            setattr(sp, key, fn)
        for n, fn in self._orig_helpers.items():
            setattr(sp, n, fn)

    def _fake_wb(self, saved_flag: list | None = None) -> SimpleNamespace:
        saved_flag = saved_flag if saved_flag is not None else []
        wb = SimpleNamespace()
        wb.Save = lambda: saved_flag.append(1)
        wb.Close = lambda SaveChanges=False: None
        wb.Application = SimpleNamespace()
        wb.VBProject = SimpleNamespace(Name="proj", VBComponents=[])
        return wb

    def _stub_gateway(self, wb=None, attach_error=None):
        xl = SimpleNamespace()

        def _attach(path, **kwargs):
            if attach_error is not None:
                raise attach_error
            return xl

        sp.excel_gateway.attach_running_excel = _attach
        sp.excel_gateway.find_open_workbook = lambda app, path: (wb, False)
        sp.excel_gateway.new_excel = lambda hidden=True: xl
        sp.excel_gateway.open_disk_workbook = lambda app, path, read_only=False: wb

    def _cfg_with_target(self, d: Path) -> dict:
        cfg = base_cfg(d)
        (d / "示例表A V1.xlsm").write_text("dummy workbook", encoding="utf-8")
        return cfg

    # --- 成功 / 预览 -------------------------------------------------------

    def test_preview_success_real_preflight(self):
        """预览：不连 Excel、不动文件；真实源码预检照常跑。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            args = make_parser().parse_args(["push", "--project", "示例表A"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 0)
            self.assertTrue(outcome.ok)
            self.assertTrue(outcome.stats["dry"])
            self.assertTrue(outcome.details[0]["target"].endswith(".xlsm"))
            self.assertEqual(outcome.details[0]["dry"], True)
            self.assertEqual(outcome.stats["results"], outcome.details)

    def test_apply_success_open_session(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            saved: list[int] = []
            wb = self._fake_wb(saved)
            self._stub_gateway(wb)
            sp._backup_workbook = lambda cfg, wb, target: Path(d) / ".backup" / "bk.xlsm"
            sp._writeback = lambda wb, proj, defaults, *, doc_modules, rollback_on_fail: ([], [])
            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 0)
            self.assertTrue(outcome.ok)
            self.assertEqual(saved, [1])
            self.assertTrue(outcome.stats["saved"])
            self.assertTrue(outcome.stats["backup"])
            self.assertTrue(outcome.stats["backup_path"])
            self.assertEqual(outcome.stats["mode"], "open")
            self.assertEqual(outcome.details[0]["result"], "写回完成")
            self.assertEqual(outcome.stats["results"][0]["modules"], 1)

    def test_call_sequence_preflight_attach_backup_writeback_readback_save(self):
        """编排顺序断言：预检 → 附着 → 备份 → 写回 → 回读 → 保存。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            order: list[str] = []
            saved: list[int] = []
            wb = self._fake_wb(saved)
            self._stub_gateway(wb)

            sp._preflight_or_fail = lambda proj, defaults, doc: (order.append("preflight") or [])
            sp._backup_workbook = lambda cfg, wb, target: (
                order.append("backup") or Path(d) / ".backup" / "bk.xlsm")

            def fake_writeback(wb, proj, defaults, *, doc_modules, rollback_on_fail):
                order.append("writeback")
                order.append("readback")  # 回读紧随写回
                return ([], [])

            sp._writeback = fake_writeback
            wb.Save = lambda: order.append("save")

            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 0)
            self.assertEqual(order, ["preflight", "backup", "writeback", "readback", "save"])

    # --- 业务失败（exit 1，带错误码）--------------------------------------

    def test_target_not_open_is_business_failure(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            err = Failure("目标表没有打开", code="TARGET_NOT_OPEN", stage="attach")
            self._stub_gateway(None, attach_error=err)
            sp._writeback = lambda wb, proj, defaults, *, doc_modules, rollback_on_fail: ([], [])
            # --open-only：表没开就失败，不自动起隐藏实例
            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply", "--open-only"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 1)
            self.assertFalse(outcome.ok)
            self.assertEqual(outcome.errors[0].code, "TARGET_NOT_OPEN")

    def test_attach_base_failure_falls_back_to_disk(self):
        """基类 Failure（TARGET_NOT_OPEN）也必须命中自动开 disk 分支，不能穿透成 exit 1。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            saved: list[int] = []
            wb = self._fake_wb(saved)
            self._stub_gateway(wb, attach_error=BaseFailure(
                "目标表没有打开", code="TARGET_NOT_OPEN", stage="attach"))
            sp.ensure_vbe_access = lambda wb, name, interactive=True: True
            sp._backup_workbook = lambda cfg, wb, target: Path(d) / ".backup" / "bk.xlsm"
            sp._writeback = lambda wb, proj, defaults, *, doc_modules, rollback_on_fail: ([], [])
            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 0, outcome.errors)
            self.assertTrue(outcome.ok)
            self.assertEqual(outcome.stats["mode"], "disk")
            self.assertTrue(outcome.stats["saved"])

    def test_attach_contract_com_context_not_fallback_to_disk(self):
        """反向保护：契约层 COM_CONTEXT 不属可转磁盘口径，原样上抛、不归一吞掉。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            err = Failure("COM 不接话", code="COM_CONTEXT", stage="attach")
            self._stub_gateway(None, attach_error=err)
            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 1)
            self.assertFalse(outcome.ok)
            self.assertEqual(outcome.errors[0].code, "COM_CONTEXT")

    def test_com_context_suggested_passed_for_push_not_isolate(self):
        """COM_CONTEXT 上抛：传给网关的 suggested 是 push 自身出口，不含 isolate。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            proj = cfg["projects"][0]
            target = d / "示例表A V1.xlsm"
            captured: dict = {}

            def _attach(path, *, target_hint="", target_suggested=""):
                captured["hint"] = target_hint
                captured["suggested"] = target_suggested
                raise Failure("COM 不接话", code="COM_CONTEXT", stage="attach")

            sp.excel_gateway.attach_running_excel = _attach
            sp.excel_gateway.find_open_workbook = lambda app, path: (None, False)
            with self.assertRaises(Failure) as cm:
                sp._attach_or_open(proj, target, auto_open=True)
            self.assertEqual(cm.exception.code, "COM_CONTEXT")
            self.assertNotIn("isolate", captured["suggested"])
            self.assertIn("磁盘模式", captured["suggested"])

            captured.clear()
            with self.assertRaises(Failure):
                sp._attach_or_open(proj, target, auto_open=False)
            self.assertNotIn("isolate", captured["hint"] + captured["suggested"])
            self.assertIn("--open-only", captured["hint"])
            # open-only 重试命令是绝对入口，不是包内相对路径
            self.assertNotIn("vba-workbench/scripts/", captured["suggested"])
            self.assertIn(sys.executable, captured["suggested"])

    def test_open_only_target_not_open_suggested_does_not_promise_auto_open(self):
        """--open-only 下表没开：提示文案是 open-only 口径（进 hint），不承诺自动开表。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            proj = cfg["projects"][0]
            target = d / "示例表A V1.xlsm"
            xl = SimpleNamespace()
            sp.excel_gateway.attach_running_excel = lambda path, **kwargs: xl
            sp.excel_gateway.find_open_workbook = lambda app, path: (None, False)
            with self.assertRaises(Failure) as cm:
                sp._attach_or_open(proj, target, auto_open=False)
            self.assertEqual(cm.exception.code, "TARGET_NOT_OPEN")
            self.assertEqual(cm.exception.stage, "attach")
            self.assertIn("--open-only", cm.exception.hint)
            self.assertIn("不会自动开表", cm.exception.hint)
            # suggested 是绝对入口重试命令，不是包内相对路径
            self.assertNotIn("vba-workbench/scripts/", cm.exception.suggested)
            self.assertIn(sys.executable, cm.exception.suggested)

    def test_open_only_no_excel_hint_uses_push_not_isolate(self):
        """--open-only 且无 Excel：归 TARGET_NOT_OPEN（不自动开表、表又没开），提示不带 run --isolate。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            proj = cfg["projects"][0]
            target = d / "示例表A V1.xlsm"
            captured: dict = {}

            def _attach(path, *, target_hint="", target_suggested=""):
                captured["hint"] = target_hint
                captured["suggested"] = target_suggested
                raise BaseFailure("没找到正在运行的 Excel", code="NO_EXCEL", stage="attach")

            sp.excel_gateway.attach_running_excel = _attach
            sp.excel_gateway.find_open_workbook = lambda app, path: (None, False)
            with self.assertRaises(BaseFailure) as cm:
                sp._attach_or_open(proj, target, auto_open=False)
            self.assertEqual(cm.exception.code, "TARGET_NOT_OPEN")
            self.assertEqual(cm.exception.stage, "attach")
            self.assertIn("--open-only", cm.exception.hint)
            self.assertIn("隐藏实例", cm.exception.hint)
            self.assertNotIn("isolate", cm.exception.hint + cm.exception.suggested)
            self.assertNotIn("vba-workbench/scripts/", cm.exception.suggested)
            self.assertIn(sys.executable, cm.exception.suggested)

    def test_open_only_no_excel_push_outcome_target_not_open(self):
        """端到端：--open-only + 网关 NO_EXCEL → errors[0] = TARGET_NOT_OPEN/attach，hint 含 --open-only。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            self._stub_gateway(None, attach_error=BaseFailure(
                "没找到正在运行的 Excel", code="NO_EXCEL", stage="attach"))
            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply", "--open-only"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 1)
            self.assertFalse(outcome.ok)
            self.assertEqual(outcome.errors[0].code, "TARGET_NOT_OPEN")
            self.assertEqual(outcome.errors[0].stage, "attach")
            self.assertIn("--open-only", outcome.errors[0].hint)

    def test_auto_open_no_excel_falls_back_to_disk(self):
        """默认允许自动开表：网关 NO_EXCEL 仍转磁盘隐藏实例，不穿透成 exit 1。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            saved: list[int] = []
            wb = self._fake_wb(saved)
            self._stub_gateway(wb, attach_error=BaseFailure(
                "没找到正在运行的 Excel", code="NO_EXCEL", stage="attach"))
            sp.ensure_vbe_access = lambda wb, name, interactive=True: True
            sp._backup_workbook = lambda cfg, wb, target: Path(d) / ".backup" / "bk.xlsm"
            sp._writeback = lambda wb, proj, defaults, *, doc_modules, rollback_on_fail: ([], [])
            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 0, outcome.errors)
            self.assertTrue(outcome.ok)
            self.assertEqual(outcome.stats["mode"], "disk")
            self.assertTrue(outcome.stats["saved"])

    def test_open_only_com_context_propagates_unchanged(self):
        """--open-only 下 COM_CONTEXT 不属可改写口径，原样上抛，不归一成 TARGET_NOT_OPEN。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            err = Failure("COM 不接话", code="COM_CONTEXT", stage="attach")
            self._stub_gateway(None, attach_error=err)
            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply", "--open-only"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 1)
            self.assertFalse(outcome.ok)
            self.assertEqual(outcome.errors[0].code, "COM_CONTEXT")

    def test_postcheck_failure_is_business_failure_not_saved(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            saved: list[int] = []
            wb = self._fake_wb(saved)
            self._stub_gateway(wb)
            sp._backup_workbook = lambda cfg, wb, target: None
            sp._writeback = lambda wb, proj, defaults, *, doc_modules, rollback_on_fail: (["公共.bas"], [])
            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 1)
            self.assertEqual(outcome.errors[0].code, "POSTCHECK_FAILED")
            self.assertEqual(saved, [])

    def _wire_engine(self, comps: dict[str, _FakeComp], *, import_result=None):
        """接真实写回引擎：低层 COM 访问全走内存桩；import_result 可按模块名返回 None 模拟失败。"""
        from toolkit.base import read_source_text
        sp.component_snapshot = lambda vbproj: {n: c.Type for n, c in comps.items()}
        sp.find_component = lambda vbproj, name: comps.get(name)
        sp.remove_module = lambda vbproj, name, retries=3: comps.pop(name, None) is not None

        def fake_import(vbproj, path, expect_name, retries=2):
            if import_result is not None and expect_name in import_result:
                return import_result[expect_name]
            comp = _FakeComp(expect_name, 1, read_source_text(path))
            comps[expect_name] = comp
            return comp

        sp.import_component = fake_import
        sp.replace_module_code = lambda comp, code: None
        sp.gbk_import_copy = lambda src, tmp: src

    def test_import_failure_after_remove_blocks_save(self):
        """Remove 后 Import 失败：硬拦 POSTCHECK_FAILED，不 Save、stats.saved 为假。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            comps = {"模块A": _FakeComp("模块A", 1, "Option Explicit\n")}
            saved: list[int] = []
            wb = self._fake_wb(saved)
            wb.Sheets = SimpleNamespace(Count=0)
            wb.VBProject = SimpleNamespace(VBComponents=list(comps.values()))
            self._stub_gateway(wb)
            sp._backup_workbook = lambda cfg, wb, target: None
            self._wire_engine(comps, import_result={"模块A": None})

            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertFalse(outcome.ok)
            self.assertEqual(outcome.exit_code, 1)
            self.assertEqual(outcome.errors[0].code, "POSTCHECK_FAILED")
            self.assertIn("模块A", str(outcome.errors[0]))
            self.assertEqual(saved, [])
            self.assertFalse(outcome.stats["saved"])

    def test_lost_module_in_snapshot_diff_blocks_save(self):
        """写回后清单少一项：硬拦 POSTCHECK_FAILED，不 Save。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            comps = {"模块A": _FakeComp("模块A", 1, PROJ_BAS),
                     "遗留模块": _FakeComp("遗留模块", 1, "Option Explicit\n")}
            saved: list[int] = []
            wb = self._fake_wb(saved)
            wb.Sheets = SimpleNamespace(Count=0)
            wb.VBProject = SimpleNamespace(VBComponents=list(comps.values()))
            self._stub_gateway(wb)
            sp._backup_workbook = lambda cfg, wb, target: None

            snapshots = [{"模块A": 1, "遗留模块": 1}, {"模块A": 1}]
            sp.component_snapshot = lambda vbproj: snapshots.pop(0)
            sp.find_component = lambda vbproj, name: comps.get(name)
            sp.remove_module = lambda vbproj, name, retries=3: True

            def fake_import(vbproj, path, expect_name, retries=2):
                from toolkit.base import read_source_text
                return comps.setdefault(
                    expect_name, _FakeComp(expect_name, 1, read_source_text(path)))

            sp.import_component = fake_import
            sp.replace_module_code = lambda comp, code: None
            sp.gbk_import_copy = lambda src, tmp: src

            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertFalse(outcome.ok)
            self.assertEqual(outcome.exit_code, 1)
            self.assertEqual(outcome.errors[0].code, "POSTCHECK_FAILED")
            self.assertIn("遗留模块", str(outcome.errors[0]))
            self.assertEqual(saved, [])
            self.assertFalse(outcome.stats["saved"])

    def test_real_engine_success_consistent_snapshot_saves(self):
        """正常路径回归：引擎真实写回、清单一致 → Save 被调用，exit 0。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            comps: dict[str, _FakeComp] = {}
            saved: list[int] = []
            wb = self._fake_wb(saved)
            wb.Sheets = SimpleNamespace(Count=0)
            wb.VBProject = SimpleNamespace(VBComponents=[])
            self._stub_gateway(wb)
            sp._backup_workbook = lambda cfg, wb, target: None
            self._wire_engine(comps)

            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 0)
            self.assertTrue(outcome.ok)
            self.assertEqual(saved, [1])
            self.assertTrue(outcome.stats["saved"])

    def test_dry_missing_target_passes_preflight_exit_zero(self):
        """dry 容忍缺表：源码可过预检 → exit 0 / ok=true / dry=true，目标记 "None"。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = base_cfg(d)  # 没有落盘 .xlsm
            setup_tree(d)
            args = make_parser().parse_args(["push", "--project", "示例表A"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 0)
            self.assertTrue(outcome.ok)
            self.assertTrue(outcome.stats["dry"])
            self.assertEqual(outcome.details[0]["target"], "None")
            self.assertEqual(outcome.details[0]["ok"], True)

    def test_dry_missing_target_preflight_fatal_exit_1(self):
        """dry + 缺 xlsm + 源码有致命：仍按 SOURCE_PREFLIGHT exit 1。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = base_cfg(d)  # 没有落盘 .xlsm
            setup_tree(d)
            proj_src = d / "src" / "projects" / "示例表A"
            (proj_src / "空模块.bas").write_text(
                'Attribute VB_Name = "空模块"\n', encoding="utf-8")
            args = make_parser().parse_args(["push", "--project", "示例表A"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 1)
            self.assertEqual(outcome.errors[0].code, "SOURCE_PREFLIGHT")
            self.assertTrue(outcome.stats["dry"])

    def test_apply_missing_target_is_target_not_found(self):
        """apply 缺 xlsm：仍落 TARGET_NOT_FOUND（apply 侧口径不变）。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = base_cfg(d)  # 没有落盘 .xlsm
            setup_tree(d)
            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 1)
            self.assertEqual(outcome.errors[0].code, "TARGET_NOT_FOUND")

    def test_project_not_found_is_business_failure(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = base_cfg(d)
            args = make_parser().parse_args(["push", "--project", "不存在的项目"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 1)
            self.assertEqual(outcome.errors[0].code, "PROJECT_NOT_FOUND")

    def test_preflight_failure_business_exit_1(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)

            def fail_preflight(proj, defaults, doc):
                raise Failure("源码预检失败", code="SOURCE_PREFLIGHT", stage="preflight")

            sp._preflight_or_fail = fail_preflight
            args = make_parser().parse_args(["push", "--project", "示例表A"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 1)
            self.assertEqual(outcome.errors[0].code, "SOURCE_PREFLIGHT")

    def test_apply_preflight_failure_aborts_before_attach(self):
        """apply 预检失败：不附着 Excel、不备份、不写回。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            touched: list[str] = []

            def fail_preflight(proj, defaults, doc):
                raise Failure("源码预检失败", code="SOURCE_PREFLIGHT", stage="preflight")

            sp._preflight_or_fail = fail_preflight
            sp._backup_workbook = lambda c, w, p: touched.append("backup")
            sp._writeback = lambda wb, proj, defaults, *, doc_modules, rollback_on_fail: (touched.append("writeback") or ([], []))
            sp.excel_gateway.attach_running_excel = lambda path: touched.append("attach")
            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 1)
            self.assertEqual(outcome.errors[0].code, "SOURCE_PREFLIGHT")
            self.assertEqual(touched, [])

    # --- 用法错（exit 2）--------------------------------------------------

    def test_missing_project_is_usage_error_exit_2(self):
        with self.assertRaises(SystemExit) as cm:
            make_parser().parse_args(["push"])
        self.assertEqual(cm.exception.code, 2)

    # --- 开关端到端生效 ----------------------------------------------------

    def test_no_backup_skips_backup(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            wb = self._fake_wb()
            self._stub_gateway(wb)
            backup_called: list[bool] = []
            sp._backup_workbook = lambda cfg, wb, target: backup_called.append(True)
            sp._writeback = lambda wb, proj, defaults, *, doc_modules, rollback_on_fail: ([], [])
            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply", "--no-backup"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 0)
            self.assertEqual(backup_called, [])
            self.assertFalse(outcome.stats["backup"])
            self.assertEqual(outcome.stats["backup_path"], "")

    def test_no_save_keeps_memory_only(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            saved: list[int] = []
            wb = self._fake_wb(saved)
            self._stub_gateway(wb)
            sp._backup_workbook = lambda cfg, wb, target: None
            sp._writeback = lambda wb, proj, defaults, *, doc_modules, rollback_on_fail: ([], [])
            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply", "--no-save"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 0)
            self.assertEqual(saved, [])
            self.assertFalse(outcome.stats["saved"])

    def test_rollback_flag_passed_to_engine(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            wb = self._fake_wb()
            self._stub_gateway(wb)
            captured: dict[str, bool] = {}

            def fake_writeback(wb, proj, defaults, *, doc_modules, rollback_on_fail):
                captured["rollback_on_fail"] = rollback_on_fail
                return ([], [])

            sp._writeback = fake_writeback
            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply", "--rollback-on-fail"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 0)
            self.assertTrue(captured["rollback_on_fail"])

    def test_backup_auto_cleanup_wired(self):
        """备份后自动清理接上：auto=True 调 prune_backups；auto=False 不调。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            pruned: list[dict] = []
            copied: list[Path] = []
            sp.backup_copy = lambda src, root, **kw: copied.append(src) or (root / "20261122-120000" / src.name)
            sp.prune_backups = lambda root, **kw: pruned.append(kw) or {"deleted": [], "leftover": []}

            cfg = base_cfg(d)
            dest = sp._backup_workbook(cfg, SimpleNamespace(), d / "示例表A V1.xlsm")
            self.assertIsNotNone(dest)
            self.assertEqual(len(pruned), 1)
            self.assertEqual(pruned[0]["keep_recent"], 10)
            self.assertFalse(pruned[0]["dry_run"])

            pruned.clear()
            cfg["defaults"]["backup_auto_clean"] = False
            sp._backup_workbook(cfg, SimpleNamespace(), d / "示例表A V1.xlsm")
            self.assertEqual(pruned, [])

    # --- 写回引擎真实编排（低层 COM 桩，不注入引擎接缝）---------------------

    def test_real_writeback_walks_engine_without_seam_injection(self):
        """不替换 _writeback 本体：真实走快照→写回→回读链路，只替 COM 访问层。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = base_cfg(d)
            setup_tree(d)
            comps: dict[str, _FakeComp] = {}
            calls: list[str] = []

            sp.component_snapshot = lambda vbproj: {n: c.Type for n, c in comps.items()}
            sp.find_component = lambda vbproj, name: comps.get(name)
            sp.remove_module = lambda vbproj, name, retries=3: True

            def fake_import(vbproj, path, expect_name, retries=2):
                from toolkit.base import read_source_text
                comp = _FakeComp(expect_name, 1, read_source_text(path))
                comps[expect_name] = comp
                calls.append("import:" + expect_name)
                return comp

            sp.import_component = fake_import
            sp.replace_module_code = lambda comp, code: calls.append("replace:" + comp.Name)
            sp.gbk_import_copy = lambda src, tmp: src

            wb = SimpleNamespace()
            wb.Sheets = SimpleNamespace(Count=0)
            wb.VBProject = SimpleNamespace(VBComponents=[])

            bad, _notes = sp._writeback(wb, cfg["projects"][0], cfg["defaults"],
                                 doc_modules=False, rollback_on_fail=False)
            self.assertEqual(bad, [])
            self.assertIn("import:公共", calls)
            self.assertIn("import:模块A", calls)

    def test_rollback_restores_snapshot_on_postcheck_failure(self):
        """postcheck 失败且 --rollback-on-fail：按写前快照复原本次写过的模块。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = base_cfg(d)
            setup_tree(d)
            comps: dict[str, _FakeComp] = {"模块A": _FakeComp("模块A", 1, "Option Explicit\n")}
            restored: list[str] = []

            sp.component_snapshot = lambda vbproj: {n: c.Type for n, c in comps.items()}
            sp._snapshot_module_code = lambda vbproj: {"模块A": "旧正文", "公共": "公共旧正文"}
            sp.find_component = lambda vbproj, name: comps.get(name)
            sp.remove_module = lambda vbproj, name, retries=3: True

            def fake_import(vbproj, path, expect_name, retries=2):
                comp = comps.setdefault(expect_name, _FakeComp(expect_name, 1, "Option Explicit\n"))
                return comp

            sp.import_component = fake_import
            sp.replace_module_code = lambda comp, code: None
            sp.gbk_import_copy = lambda src, tmp: src
            sp._restore_module_code = lambda vbproj, snap: (
                [restored.append(n) for n in snap], [])

            wb = SimpleNamespace()
            wb.Sheets = SimpleNamespace(Count=0)
            wb.VBProject = SimpleNamespace(VBComponents=[])

            bad, _notes = sp._writeback(wb, cfg["projects"][0], cfg["defaults"],
                                doc_modules=False, rollback_on_fail=True)
            self.assertTrue(bad)
            self.assertIn("模块A", restored)

    # --- 批量 -------------------------------------------------------------

    def test_all_projects_preview_rows(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            args = make_parser().parse_args(["push", "--all-project"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 0)
            self.assertEqual(outcome.stats["total"], 1)
            self.assertEqual(outcome.stats["ok"], 1)
            self.assertTrue(outcome.stats["dry"])
            self.assertEqual(outcome.details[0]["status"], "dry-run")

    def test_all_projects_business_failure_continues(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)

            def fail_preflight(proj, defaults, doc):
                raise Failure("源码预检失败", code="SOURCE_PREFLIGHT", stage="preflight")

            sp._preflight_or_fail = fail_preflight
            args = make_parser().parse_args(["push", "--all-project"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 1)
            self.assertTrue(outcome.errors)
            self.assertTrue(outcome.details[0]["status"].startswith("失败："))

    def test_single_save_failure_is_save_failed(self):
        """单表保存失败：旧版落 SAVE_FAILED（区别于批量兜底 UNEXPECTED）。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            wb = self._fake_wb()

            def boom():
                raise RuntimeError("文件只读")

            wb.Save = boom
            self._stub_gateway(wb)
            sp._backup_workbook = lambda c, w, p: None
            sp._writeback = lambda wb, proj, defaults, *, doc_modules, rollback_on_fail: ([], [])
            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 1)
            self.assertEqual(outcome.errors[0].code, "SAVE_FAILED")

    def test_vbe_protected_failure_code(self):
        """磁盘模式工程不可读写：归 VBE_PROTECTED（不使用未登记的 VBE_DENIED）。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            wb = self._fake_wb()
            open_err = Failure("目标表没有打开", code="TARGET_NOT_OPEN", stage="attach")
            self._stub_gateway(wb, attach_error=open_err)
            sp.ensure_vbe_access = lambda wb, title, interactive=None: False
            sp.vbe_state = lambda wb: ("protected", "工程已被保护")
            args = make_parser().parse_args(
                ["push", "--project", "示例表A", "--apply"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 1)
            self.assertEqual(outcome.errors[0].code, "VBE_PROTECTED")

    def test_batch_save_non_failure_lands_unexpected(self):
        """批量保存抛非 Failure：旧版兜底落 UNEXPECTED（不包装成 SAVE_FAILED）。"""
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            wb = self._fake_wb()

            def boom():
                raise RuntimeError("磁盘只读")

            wb.Save = boom
            self._stub_gateway(wb)
            sp.has_vbproject_access = lambda wb: True
            sp._writeback = lambda wb, proj, defaults, *, doc_modules, rollback_on_fail: ([], [])
            args = make_parser().parse_args(["push", "--all-project", "--apply"])
            outcome = sp.push(make_ctx(args, cfg))
            self.assertEqual(outcome.exit_code, 1)
            self.assertEqual(outcome.errors[0].code, "UNEXPECTED")
            self.assertTrue(outcome.details[0]["status"].startswith("失败："))

    # --- JSON 纯度 ---------------------------------------------------------

    def test_json_envelope_single_line(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            cfg = self._cfg_with_target(d)
            setup_tree(d)
            args = make_parser().parse_args(["push", "--project", "示例表A", "--json"])
            outcome = sp.push(make_ctx(args, cfg))
            payload, code, line = render_json(outcome)
            self.assertEqual(code, 0)
            self.assertEqual(len(line.splitlines()), 1)
            for key in ("ok", "action", "details", "errors", "elapsed_ms", "stats", "exit_code"):
                self.assertIn(key, payload)
            self.assertEqual(payload["action"], "sync.push")


class _FakeRef:
    def __init__(self, desc: str):
        self.Description = desc


class _FakeRefs:
    """工程引用集合桩：返回给定描述列表。"""

    def __init__(self, descs: list[str]):
        self._descs = descs

    @property
    def Count(self) -> int:
        return len(self._descs)

    def Item(self, i: int) -> _FakeRef:
        return _FakeRef(self._descs[i - 1])


class PreflightNotesTest(unittest.TestCase):
    """预检软提醒通道：notes 只进 warnings，不改退出码；fatal 优先阻断。"""

    @staticmethod
    def _code(name: str = "测试模块", extra: str = "") -> str:
        return (f'Attribute VB_Name = "{name}"\n'
                "Option Explicit\n\n"
                "Public Sub 测试()\n"
                "    Debug.Print 1\n"
                "End Sub\n" + extra)

    def _case(self, files: dict[str, bytes | str]):
        d = Path(tempfile.mkdtemp())
        self.addCleanup(lambda: __import__("shutil").rmtree(d, ignore_errors=True))
        proj_dir = d / "proj"
        common_dir = d / "common"
        proj_dir.mkdir()
        common_dir.mkdir()
        for name, content in files.items():
            p = proj_dir / name
            p.parent.mkdir(parents=True, exist_ok=True)
            if isinstance(content, bytes):
                p.write_bytes(content)
            else:
                p.write_text(content, encoding="utf-8", newline="\n")
        proj = {"name": "测试", "src": str(proj_dir)}
        defaults = {"common_dir": str(common_dir)}
        return d, proj, defaults

    def _cfg(self, d: Path, proj_dir: Path, common_dir: Path) -> dict:
        return {
            "defaults": {"common_dir": str(common_dir),
                         "backup_dir": str(d / ".backup"),
                         "import_document_modules": False,
                         "backup_auto_clean": True, "backup_keep": 10,
                         "backup_keep_daily": 1, "backup_min_keep": 3,
                         "backup_protect": []},
            "projects": [{"name": "测试", "search_dir": str(d), "prefix": "测试",
                          "require_version": False, "src": str(proj_dir)}],
        }

    def test_soft_notes_do_not_block_exit_zero_and_warnings_carried(self):
        """仅软提醒：退出码 0、warnings 有内容、fatal 为空。"""
        bom = b"\xef\xbb\xbf" + self._code("BOM模块").encode("utf-8")
        crlf = self._code("CRLF模块").replace("\n", "\r\n").encode("utf-8")
        d, proj, defaults = self._case({
            "BOM模块.bas": bom,
            "CRLF模块.bas": crlf,
            "缺选项.bas": 'Attribute VB_Name = "缺选项"\nPublic Sub X()\nEnd Sub\n',
            "截断.bas": self._code("截断", "Public Sub 未完() _\n"),
        })
        fatal, notes = sp._source_preflight(proj, defaults, True)
        self.assertEqual(fatal, [], fatal)
        self.assertTrue(notes, "软提醒不应为空")
        text = "\n".join(notes)
        self.assertIn("BOM", text)
        self.assertIn("CRLF", text)
        self.assertIn("Option Explicit", text)
        self.assertIn("续行符", text)

        # 端到端预览：软提醒进 Outcome.warnings，退出码仍为 0
        (d / "测试 V1.xlsm").write_text("dummy", encoding="utf-8")
        cfg = self._cfg(d, Path(proj["src"]), Path(defaults["common_dir"]))
        args = make_parser().parse_args(["push", "--project", "测试"])
        outcome = sp.push(make_ctx(args, cfg))
        self.assertEqual(outcome.exit_code, 0)
        self.assertFalse(outcome.errors)
        self.assertTrue(outcome.warnings)

    def test_me_private_field_is_soft_note(self):
        """类模块 Me.<私有字段> 只提醒，不阻断写回。"""
        cls = ("VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\n"
               'Attribute VB_Name = "私有字段类"\n'
               "Option Explicit\n\n"
               "Private mCount As Long\n\n"
               "Public Sub 加一()\n"
               "    Me.mCount = Me.mCount + 1\n"
               "End Sub\n")
        _d, proj, defaults = self._case({"私有字段类.cls": cls})
        fatal, notes = sp._source_preflight(proj, defaults, True)
        self.assertFalse(fatal, fatal)
        self.assertTrue(any("Me. 引用私有字段" in x and "私有字段类" in x for x in notes))

    def test_truncated_last_line_is_soft_note(self):
        """末行以续行符 _ 结尾（疑似截断）只提醒。"""
        _d, proj, defaults = self._case({
            "截断.bas": self._code("截断", "Public Sub 未完() _\n"),
        })
        fatal, notes = sp._source_preflight(proj, defaults, True)
        self.assertFalse(fatal, fatal)
        self.assertTrue(any("续行符" in x for x in notes))

    def test_missing_reference_is_soft_note(self):
        """写回后缺失引用扫描：只产出软提醒，不改退出码。"""
        d, proj, defaults = self._case({
            "用ADODB.bas": ('Attribute VB_Name = "用ADODB"\nOption Explicit\n'
                             "Public Sub 取数()\n"
                             "    Dim rs As ADODB.Connection\n"
                             "End Sub\n"),
        })
        proj_dir = Path(proj["src"])
        # 工程只勾了 Scripting Runtime，缺 ADODB
        vbproj = SimpleNamespace(References=_FakeRefs(["Microsoft Scripting Runtime"]),
                                 VBComponents=[])
        notes = sp._report_references(vbproj, proj_dir, None, proj_dir.name)
        self.assertTrue(notes)
        text = "\n".join(notes)
        self.assertIn("缺少以下引用", text)
        self.assertIn("ActiveX Data Objects", text)

        # 经 _writeback 编排：缺失引用 notes 随返回值带出，bad 为空
        wb = SimpleNamespace(VBProject=vbproj)
        orig_import_folder = sp._import_folder
        self.addCleanup(lambda: setattr(sp, "_import_folder", orig_import_folder))
        sp._import_folder = lambda *a, **k: (1, 0, {})
        bad, wb_notes = sp._writeback(
            wb, proj, defaults, doc_modules=False, rollback_on_fail=False)
        self.assertEqual(bad, [])
        self.assertTrue(any("ActiveX Data Objects" in x for x in wb_notes))

    def test_fatal_and_notes_together_fatal_wins_exit_1(self):
        """fatal 与 notes 同时存在：fatal 优先阻断，退出码 1。"""
        d, proj, defaults = self._case({
            "空模块.bas": 'Attribute VB_Name = "空模块"\n',
            "CRLF模块.bas": self._code("CRLF模块").replace("\n", "\r\n"),
        })
        fatal, notes = sp._source_preflight(proj, defaults, True)
        self.assertTrue(fatal)
        self.assertTrue(notes)  # 软提醒同时存在，但不改变阻断口径

        with self.assertRaises(Failure) as cm:
            sp._preflight_or_fail(proj, defaults, True)
        self.assertEqual(cm.exception.code, "SOURCE_PREFLIGHT")

        (d / "测试 V1.xlsm").write_text("dummy", encoding="utf-8")
        cfg = self._cfg(d, Path(proj["src"]), Path(defaults["common_dir"]))
        args = make_parser().parse_args(["push", "--project", "测试"])
        outcome = sp.push(make_ctx(args, cfg))
        self.assertEqual(outcome.exit_code, 1)
        self.assertEqual(outcome.errors[0].code, "SOURCE_PREFLIGHT")


class ExcelAttachSuggestedDefaultTest(unittest.TestCase):
    """共享层保护：不传 target_suggested 时三处 COM_CONTEXT 默认文案逐字不变；传了则覆盖。"""

    DEFAULTS = {
        "rot": "先检查 Excel 是否卡在弹窗/受保护视图；先关掉卡住的实例后重试附着；该表仍不可用时，换一张能打开的表再走 isolate",
        "ready": "先检查 Excel 是否在受保护视图/不同权限会话；先关掉卡住的实例后重试附着；该表仍不可用时，换一张能打开的表再走 isolate",
        "books": "确认 Excel 没有卡在弹窗/受保护视图；先关掉卡住的实例后重试附着；该表仍不可用时，换一张能打开的表再走 isolate",
    }

    def setUp(self):
        self._orig = {
            "find_excel_with_workbook": excel_toolkit.find_excel_with_workbook,
            "_mark_app_ready": excel_toolkit._mark_app_ready,
            "find_open_workbook": excel_toolkit.find_open_workbook,
        }
        self._orig_modules = {n: sys.modules.get(n) for n in
                              ("pythoncom", "win32com", "win32com.client")}
        fake_client = SimpleNamespace(GetActiveObject=lambda progid: SimpleNamespace())
        sys.modules["pythoncom"] = SimpleNamespace(CoInitialize=lambda: None)
        sys.modules["win32com"] = SimpleNamespace(client=fake_client)
        sys.modules["win32com.client"] = fake_client

    def tearDown(self):
        for n, fn in self._orig.items():
            setattr(excel_toolkit, n, fn)
        for n, mod in self._orig_modules.items():
            if mod is None:
                sys.modules.pop(n, None)
            else:
                sys.modules[n] = mod

    def test_default_suggested_retained_when_not_overridden(self):
        # ROT 命中实例但 COM 拒答（:223）
        excel_toolkit.find_excel_with_workbook = lambda path: (SimpleNamespace(), None)
        excel_toolkit._mark_app_ready = lambda xl: (_ for _ in ()).throw(RuntimeError("boom"))
        with self.assertRaises(excel_toolkit.Failure) as cm:
            excel_toolkit.attach_running_excel(Path("示例表A V1.xlsm"))
        self.assertEqual(cm.exception.code, "COM_CONTEXT")
        self.assertEqual(cm.exception.suggested, self.DEFAULTS["rot"])

        # GetActiveObject 后实例不 ready（:247）
        excel_toolkit.find_excel_with_workbook = lambda path: (None, None)
        with self.assertRaises(excel_toolkit.Failure) as cm:
            excel_toolkit.attach_running_excel(None)
        self.assertEqual(cm.exception.code, "COM_CONTEXT")
        self.assertEqual(cm.exception.suggested, self.DEFAULTS["ready"])

        # 实例 ready 但读不到工作簿列表（:259）
        excel_toolkit._mark_app_ready = lambda xl: None
        excel_toolkit.find_open_workbook = lambda xl, path: (_ for _ in ()).throw(RuntimeError("boom"))
        with self.assertRaises(excel_toolkit.Failure) as cm:
            excel_toolkit.attach_running_excel(Path("示例表A V1.xlsm"))
        self.assertEqual(cm.exception.code, "COM_CONTEXT")
        self.assertEqual(cm.exception.suggested, self.DEFAULTS["books"])

    def test_target_suggested_overrides_default(self):
        excel_toolkit.find_excel_with_workbook = lambda path: (None, None)
        excel_toolkit._mark_app_ready = lambda xl: (_ for _ in ()).throw(RuntimeError("boom"))
        with self.assertRaises(excel_toolkit.Failure) as cm:
            excel_toolkit.attach_running_excel(None, target_suggested="自定义出口命令")
        self.assertEqual(cm.exception.code, "COM_CONTEXT")
        self.assertEqual(cm.exception.suggested, "自定义出口命令")


if __name__ == "__main__":
    unittest.main()
