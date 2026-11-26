# -*- coding: utf-8 -*-
"""1.6.5 回归：run --isolate 预检遇到 COM_CONTEXT 时的提示口径。

范围只在提示层：放行白名单（NO_EXCEL / TARGET_NOT_OPEN）与错误码归类不动；
补的是「目标实例 COM 不可达时隔离模式也不启动」的人话与可执行建议命令。
约束：不启动 Excel；COM 边界用桩对象 / monkeypatch，测试件只落临时目录。
"""
from __future__ import annotations

import io
import json
import sys
import tempfile
import time
import types
import unittest
from contextlib import contextmanager
from pathlib import Path
from unittest import mock

PKG = Path(__file__).resolve().parent.parent / "scripts"
if str(PKG) not in sys.path:
    sys.path.insert(0, str(PKG))

from adapters import excel_gateway  # noqa: E402
from features import run as run_feature  # noqa: E402
from toolkit.base import Failure as BaseFailure, looks_like_command  # noqa: E402
from toolkit.contract import ErrorCode, Failure  # noqa: E402
from toolkit.registry import build_parser  # noqa: E402
from toolkit.render import render  # noqa: E402

SRC_CODE = "Public Sub 示例求和()\nEnd Sub"


# --------------------------------------------------------------------------- 桩对象
class _FakeApp:
    def __init__(self) -> None:
        self.ActiveWorkbook = None
        self.AutomationSecurity = 2
        self.Workbooks = []

    def Quit(self) -> None:
        pass


class _FakeWB:
    def __init__(self, app: _FakeApp, name: str, full_name: str) -> None:
        self.Application = app
        self._name = name
        self.FullName = full_name
        self.Path = str(Path(full_name).parent)
        self.VBProject = types.SimpleNamespace(VBComponents=[])

    @property
    def Name(self) -> str:
        return self._name


class _Prog:
    @contextmanager
    def stage(self, name: str):
        yield


class _Ctx:
    """最小上下文：字段对齐 AppContext，journal_dir 指向临时目录。"""

    def __init__(self, args: object, cfg: dict, root: Path,
                 journal_dir: Path | None = None) -> None:
        self.args = args
        self.config = cfg
        self.root = root
        self.json_mode = False
        self.started = time.perf_counter()
        self.progress = _Prog()
        self.journal_dir = journal_dir or (root / ".run")

    def elapsed_ms(self) -> int:
        return 0

    def stage_timings(self) -> dict:
        return {}


# --------------------------------------------------------------------------- 环境搭建
def _make_env(tmp: Path) -> tuple[dict, Path, Path]:
    src = tmp / "srcproj"
    common = tmp / "common"
    search = tmp / "search"
    backup = tmp / "backup"
    for d in (src, common, search, backup):
        d.mkdir(parents=True, exist_ok=True)
    (src / "示例模块.bas").write_text(SRC_CODE, encoding="utf-8")
    xlsm = search / "示例表A V1.xlsm"
    xlsm.write_bytes(b"x")
    cfg = {
        "defaults": {"common_dir": str(common), "backup_dir": str(backup)},
        "projects": [{
            "name": "示例表A",
            "src": str(src),
            "search_dir": str(search),
            "prefix": "示例表A",
        }],
    }
    return cfg, tmp, xlsm


def _parse(argv: list[str]):
    parser = build_parser(run_feature.SPECS, prog="sync")
    return parser.parse_args(argv)


def _patch_gateway(test: unittest.TestCase, raiser):
    test.addCleanup(setattr, excel_gateway, "attach_running_excel",
                    excel_gateway.attach_running_excel)
    excel_gateway.attach_running_excel = raiser


# --------------------------------------------------------------------------- 用例
class IsolateComContextHintTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = Path(tempfile.mkdtemp(prefix="vba-run-165-"))
        self.cfg, self.root, self.xlsm = _make_env(self._tmp)

    def _ctx(self, argv: list[str]) -> _Ctx:
        return _Ctx(_parse(argv), self.cfg, self.root, self.root / ".run")

    def test_com_context_hints_isolate_wont_start_and_suggests_command(self):
        """COM_CONTEXT 上抛：归类不动，人话点明「隔离模式也不启动」，suggested 可执行。"""
        proj = self.cfg["projects"][0]
        err = Failure(
            "ROT 看得到但 BindToObject 失败",
            code="COM_CONTEXT", stage="attach",
            hint="旧网关 hint：先试试 run --isolate",
            suggested="旧网关建议：仍不行再用 run --isolate",
        )
        _patch_gateway(self, lambda path, **kw: (_ for _ in ()).throw(err))

        with self.assertRaises(Failure) as cm:
            run_feature._acquire(self.cfg, proj, isolate=True, push=False,
                                 macro="示例求和")
        raised = cm.exception

        # 判定口径不动：错误码与阶段原样保留，不被兜底改写成别的码。
        self.assertEqual(raised.code, ErrorCode.COM_CONTEXT.value)
        self.assertEqual(raised.stage, "attach")

        # 人话必须讲清：目标实例 COM 不可达，且隔离模式也不会启动。
        self.assertIn("COM 不可达", raised.hint)
        self.assertIn("隔离模式也不会启动", raised.hint)
        # 隔离语境下照搬网关「再开 isolate」的旧话术，不能残留。
        self.assertNotIn("旧网关", raised.hint)
        self.assertNotIn("旧网关", raised.suggested)

        # suggested 是可执行命令：入口命令 + --isolate + 项目 + 宏齐全。
        self.assertTrue(looks_like_command(raised.suggested), raised.suggested)
        self.assertIn("--isolate", raised.suggested)
        self.assertIn("--project", raised.suggested)
        self.assertIn("示例表A", raised.suggested)
        self.assertIn("--macro", raised.suggested)
        self.assertIn("示例求和", raised.suggested)

    def test_no_excel_still_falls_back_to_isolate(self):
        """NO_EXCEL 仍是放行口径：预检不阻断，照常另开隔离副本。"""
        proj = self.cfg["projects"][0]
        err = BaseFailure("没装 Excel", code="NO_EXCEL", stage="attach")
        _patch_gateway(self, lambda path, **kw: (_ for _ in ()).throw(err))

        fake_app = _FakeApp()
        fake_wb = _FakeWB(fake_app, "示例表A V1.xlsm", str(self.xlsm))
        self.addCleanup(setattr, excel_gateway, "new_excel",
                        excel_gateway.new_excel)
        excel_gateway.new_excel = lambda hidden=True: fake_app
        self.addCleanup(setattr, excel_gateway, "open_disk_workbook",
                        excel_gateway.open_disk_workbook)
        excel_gateway.open_disk_workbook = lambda xl, path, read_only=False: fake_wb

        wb, _xl, mode, owned = run_feature._acquire(
            self.cfg, proj, isolate=True, push=False, macro="示例求和")
        self.assertEqual(mode, "isolate")
        self.assertIs(wb, fake_wb)
        self.assertIsNotNone(owned)

    def test_target_not_open_still_falls_back_to_isolate(self):
        """TARGET_NOT_OPEN 仍是放行口径：预检不阻断，照常另开隔离副本。"""
        proj = self.cfg["projects"][0]
        err = BaseFailure("目标表没有打开", code="TARGET_NOT_OPEN", stage="attach")
        _patch_gateway(self, lambda path, **kw: (_ for _ in ()).throw(err))

        fake_app = _FakeApp()
        fake_wb = _FakeWB(fake_app, "示例表A V1.xlsm", str(self.xlsm))
        self.addCleanup(setattr, excel_gateway, "new_excel",
                        excel_gateway.new_excel)
        excel_gateway.new_excel = lambda hidden=True: fake_app
        self.addCleanup(setattr, excel_gateway, "open_disk_workbook",
                        excel_gateway.open_disk_workbook)
        excel_gateway.open_disk_workbook = lambda xl, path, read_only=False: fake_wb

        wb, _xl, mode, owned = run_feature._acquire(
            self.cfg, proj, isolate=True, push=False, macro="示例求和")
        self.assertEqual(mode, "isolate")
        self.assertIs(wb, fake_wb)
        self.assertIsNotNone(owned)

    def test_json_envelope_is_single_line_with_hint_in_fields(self):
        """JSON 模式：失败信封单行可解析；提示进 errors[].hint / suggested_command，
        不混进 stdout 正文；stats 里留同一口径。"""
        err = Failure("ROT 看得到但 BindToObject 失败",
                      code="COM_CONTEXT", stage="attach")
        _patch_gateway(self, lambda path, **kw: (_ for _ in ()).throw(err))

        ctx = self._ctx(["run", "--project", "示例表A", "--macro", "示例求和",
                         "--isolate", "--apply", "--json"])
        outcome = run_feature.run(ctx)
        self.assertEqual(outcome.exit_code, 1)
        self.assertFalse(outcome.ok)

        buf = io.StringIO()
        with mock.patch("sys.stdout", buf):
            code = render(outcome, json_mode=True)
        self.assertEqual(code, 1)

        lines = [ln for ln in buf.getvalue().splitlines() if ln.strip()]
        self.assertEqual(len(lines), 1, f"stdout 不止一行：{buf.getvalue()!r}")
        payload = json.loads(lines[0])

        self.assertEqual(payload["exit_code"], 1)
        self.assertEqual(payload["errors"][0]["code"], "COM_CONTEXT")
        hint = payload["errors"][0]["hint"]
        suggested = payload["errors"][0]["suggested_command"]
        self.assertIn("COM 不可达", hint)
        self.assertIn("隔离模式也不会启动", hint)
        self.assertTrue(looks_like_command(suggested), suggested)
        self.assertIn("--isolate", suggested)

        # 同口径在 stats 也留一份（入口层既有形状）。
        self.assertIn("隔离模式也不会启动", payload["stats"]["hint"])
        self.assertIn("--isolate", payload["stats"]["suggested_command"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
