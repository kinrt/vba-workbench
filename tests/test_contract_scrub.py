# -*- coding: utf-8 -*-
"""scrub 命令的交付件范围扫描：纯函数覆盖白名单、GBK、大小写与排除项。"""
from __future__ import annotations

import contextlib
import sys
import tempfile
import types
import unittest
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))

import features.core as core  # noqa: E402
from features.core import _scan_package, _scrub_deny_terms  # noqa: E402


class _Progress:
    @contextlib.contextmanager
    def stage(self, _name: str):
        yield


def _ctx(root: Path, deny: list[str] | None) -> types.SimpleNamespace:
    return types.SimpleNamespace(
        config={"scrub": {"deny": deny}} if deny is not None else {},
        args=types.SimpleNamespace(path=str(root)),
        progress=_Progress(),
        elapsed_ms=lambda: 1,
        stage_timings=lambda: {},
    )


class ScrubScanTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)
        (self.root / "scripts").mkdir()

    def _write(self, rel: str, text: str, encoding: str = "utf-8") -> None:
        path = self.root / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(text.encode(encoding))

    def test_clean_package_has_no_hit(self) -> None:
        self._write("SKILL.md", "hello\n")
        self._write("scripts/a.py", "print('ok')\n")
        rels, hits, skipped = _scan_package(self.root, ["secret"])
        self.assertEqual(len(rels), 2)
        self.assertEqual(hits, [])
        self.assertEqual(skipped, [])

    def test_gbk_bat_hit_is_found(self) -> None:
        self._write("scripts/启动.bat", "echo 秘密渠道\r\n", encoding="gbk")
        _, hits, _ = _scan_package(self.root, ["秘密渠道"])
        self.assertEqual(len(hits), 1)
        self.assertEqual(hits[0]["file"], "scripts/启动.bat")
        self.assertEqual(hits[0]["line"], 1)
        self.assertEqual(hits[0]["word"], "秘密渠道")

    def test_match_is_case_insensitive(self) -> None:
        self._write("SKILL.md", "KINRT\n")
        _, hits, _ = _scan_package(self.root, ["kinrt"])
        self.assertEqual(len(hits), 1)

    def test_tests_and_evals_are_outside_package_scope(self) -> None:
        self._write("SKILL.md", "hello\n")
        self._write("tests/report.md", "secret\n")
        self._write("evals/eval.json", "secret\n")
        rels, hits, _ = _scan_package(self.root, ["secret"])
        self.assertEqual(sorted(rels), ["SKILL.md"])
        self.assertEqual(hits, [])

    def test_non_text_asset_is_skipped(self) -> None:
        self._write("assets/logo.png", "secret\n")
        _, hits, skipped = _scan_package(self.root, ["secret"])
        self.assertEqual(hits, [])
        self.assertEqual(skipped, ["assets/logo.png"])

    def test_deny_terms_are_deduped_and_sorted(self) -> None:
        self.assertEqual(_scrub_deny_terms({"scrub": {"deny": ["B", "a", "b", ""]}}),
                         ["a", "B"])
        with self.assertRaises(ValueError):
            _scrub_deny_terms({})
        with self.assertRaises(ValueError):
            _scrub_deny_terms({"scrub": {"deny": []}})


class ScrubHandlerTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)
        (self.root / "scripts").mkdir()
        (self.root / "SKILL.md").write_text("hello\n", encoding="utf-8")

    def test_success(self) -> None:
        out = core.scrub(_ctx(self.root, ["secret"]))
        self.assertTrue(out.ok)
        self.assertEqual(out.exit_code, 0)
        self.assertEqual(out.details, [])
        self.assertEqual(out.stats["hits"], 0)

    def test_hit_is_business_failure(self) -> None:
        (self.root / "SKILL.md").write_text("kinrt\n", encoding="utf-8")
        out = core.scrub(_ctx(self.root, ["kinrt"]))
        self.assertFalse(out.ok)
        self.assertEqual(out.exit_code, 1)
        self.assertEqual(out.errors[0].code, "SCRUB_HIT")
        self.assertEqual(out.errors[0].stage, "preflight")
        self.assertEqual(out.details[0]["file"], "SKILL.md")

    def test_missing_deny_is_config_failure(self) -> None:
        out = core.scrub(_ctx(self.root, None))
        self.assertFalse(out.ok)
        self.assertEqual(out.exit_code, 1)
        self.assertEqual(out.errors[0].code, "CONFIG_INVALID")


if __name__ == "__main__":
    unittest.main()
