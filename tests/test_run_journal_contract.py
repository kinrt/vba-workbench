# -*- coding: utf-8 -*-
"""1.6.11 自动回归：run journal 的 save / isolate 口径。

约束：不启动 Excel、不碰业务表；attach 与宏执行全部用桩对象，journal 只落临时目录。
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
from toolkit.contract import ErrorCode  # noqa: E402
from toolkit.registry import build_parser  # noqa: E402
from toolkit.render import render  # noqa: E402

SRC_CODE = "Public Sub 示例求和()\nEnd Sub"


class _FakeApp:
    def __init__(self, ret: object = 5) -> None:
        self.ActiveWorkbook = None
        self.AutomationSecurity = 2
        self.Workbooks = []
        self._ret = ret
        self.calls: list[tuple[str, tuple]] = []

    def Run(self, macro: str, *args):
        self.calls.append((macro, tuple(args)))
        return self._ret


class _FakeWB:
    def __init__(self, app: _FakeApp, name: str, full_name: str) -> None:
        app.ActiveWorkbook = types.SimpleNamespace(Name=name)
        self.Application = app
        self._name = name
        self.FullName = full_name
        self.Path = str(Path(full_name).parent)
        self.VBProject = types.SimpleNamespace(VBComponents=[])
        self.saved = False
        self.activated = False

    @property
    def Name(self) -> str:
        return self._name

    def Activate(self) -> None:
        self.activated = True

    def Save(self) -> None:
        self.saved = True


class _Prog:
    @contextmanager
    def stage(self, name: str):
        yield


class _Ctx:
    def __init__(self, args: object, cfg: dict, root: Path,
                 journal_dir: Path | None = None) -> None:
        self.args = args
        self.config = cfg
        self.root = root
        self.json_mode = bool(getattr(args, "json", False))
        self.started = time.perf_counter()
        self.progress = _Prog()
        self.journal_dir = journal_dir or (root / ".run")

    def elapsed_ms(self) -> int:
        return 0

    def stage_timings(self) -> dict:
        return {}


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


def _patch_attach(test: unittest.TestCase, app: _FakeApp) -> None:
    test.addCleanup(setattr, excel_gateway, "attach_running_excel",
                    excel_gateway.attach_running_excel)
    excel_gateway.attach_running_excel = lambda path=None, **kw: app


class RunJournalContractTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = Path(tempfile.mkdtemp(prefix="vba-run-journal-"))
        self.cfg, self.root, self.xlsm = _make_env(self._tmp)

    def _ctx(self, argv: list[str]) -> _Ctx:
        return _Ctx(_parse(argv), self.cfg, self.root, self.root / ".run")

    def _journals(self) -> list[dict]:
        return [json.loads(p.read_text(encoding="utf-8"))
                for p in sorted((self.root / ".run").glob("*.json"))]

    def test_save_without_isolate_is_reported_and_journaled(self):
        """非隔离 --save：人话显式提示，journal 记 save_requested/save_ignored。"""
        app = _FakeApp(ret=5)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm))
        app.Workbooks = [wb]
        _patch_attach(self, app)

        ctx = self._ctx(["run", "--project", "示例表A", "--macro", "示例求和",
                         "--save", "--apply", "--json"])
        outcome = run_feature.run(ctx)

        self.assertEqual(outcome.exit_code, 0, outcome.errors)
        self.assertTrue(any("仅对 --isolate 生效" in w for w in outcome.warnings),
                        outcome.warnings)
        self.assertFalse(wb.saved, "非隔离 --save 不能真的保存工作簿")

        journal = self._journals()[0]
        self.assertEqual(journal["mode"], "attach")
        self.assertTrue(journal["save_requested"])
        self.assertTrue(journal["save_ignored"])
        self.assertFalse(journal["saved"])

        buf = io.StringIO()
        with mock.patch("sys.stdout", buf):
            code = render(outcome, json_mode=True)
        self.assertEqual(code, 0)
        payload = json.loads(buf.getvalue().strip())
        self.assertTrue(any("仅对 --isolate 生效" in w for w in payload["stats"]["warnings"]))

    def test_isolate_on_open_target_keeps_mode_isolate_in_journal(self):
        """--isolate 撞已开表：错误为 TARGET_OPEN，journal mode 不能退成 unknown。"""
        app = _FakeApp(ret=None)
        wb = _FakeWB(app, "示例表A V1.xlsm", str(self.xlsm))
        app.Workbooks = [wb]
        _patch_attach(self, app)

        ctx = self._ctx(["run", "--project", "示例表A", "--macro", "示例求和",
                         "--isolate", "--apply"])
        outcome = run_feature.run(ctx)

        self.assertEqual(outcome.exit_code, 1)
        self.assertEqual(outcome.errors[0].code, ErrorCode.TARGET_OPEN.value)
        journal = self._journals()[0]
        self.assertEqual(journal["mode"], "isolate")
        self.assertEqual(journal["exit_code"], 1)
        self.assertIn("隔离模式要求目标表未打开", journal["error"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
