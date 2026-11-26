# -*- coding: utf-8 -*-
"""gen_commands_md 测试：幂等、表格格式、空 manifest/重复登记报错退出。"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))
sys.path.insert(0, str(PKG / "tests" / "tools"))

import gen_commands_md as gen  # noqa: E402
from toolkit.contract import MutationMode  # noqa: E402
from toolkit.registry import CommandSpec, ProjectMode  # noqa: E402


def _spec(verb: str = "version", entry: str = "sync", **overrides) -> CommandSpec:
    base = dict(
        verb=verb, entry=entry, handler=f"features.core:{verb}",
        mutation=MutationMode.NONE, project=ProjectMode.NONE,
        supports_json=True, exit_codes=(0, 1, 2),
        summary="示例动词", ref="references/x.md", tests=("contract.x.success",),
    )
    base.update(overrides)
    return CommandSpec(**base)


class GenCommandsMdTest(unittest.TestCase):
    def test_render_is_deterministic(self):
        specs = (_spec("version"), _spec("list", entry="sync"),
                 _spec("probe", entry="agent", project=ProjectMode.OPTIONAL))
        first = gen.render_markdown(specs)
        second = gen.render_markdown(specs)
        self.assertEqual(first, second)

    def test_generate_is_idempotent_on_disk(self):
        specs = (_spec("version"),)
        with tempfile.TemporaryDirectory() as d:
            out = Path(d) / "commands.md"
            gen.generate(out, specs=specs)
            first = out.read_bytes()
            gen.generate(out, specs=specs)
            second = out.read_bytes()
            self.assertEqual(first, second)
            self.assertEqual(first, second)
            self.assertNotIn(b".tmp", out.read_bytes())

    def test_table_columns_and_rows(self):
        specs = (_spec("version"), _spec("doctor", entry="sync"),
                 _spec("probe", entry="agent", project=ProjectMode.OPTIONAL))
        md = gen.render_markdown(specs)
        self.assertIn("| 动词 | 参数 | 变更模式 | 退出码 | --json | 对应测试 | 参考 |", md)
        self.assertIn("`version`", md)
        self.assertIn("`doctor`", md)
        self.assertIn("`probe`", md)
        # 退出码与 json 标记落表
        self.assertIn("0/1/2", md)
        self.assertIn("是", md)

    def test_empty_manifest_raises(self):
        with self.assertRaises(gen.ManifestError):
            gen.build_rows(())

    def test_duplicate_registration_raises(self):
        specs = (_spec("version"), _spec("version"))
        with self.assertRaises(gen.ManifestError):
            gen.build_rows(specs)

    def test_unknown_entry_raises(self):
        specs = (_spec("version", entry="browser"),)
        with self.assertRaises(gen.ManifestError):
            gen.build_rows(specs)

    def test_missing_handler_raises(self):
        specs = (_spec("version", handler=""),)
        with self.assertRaises(gen.ManifestError):
            gen.build_rows(specs)

    def test_real_manifest_loads_and_matches_docs(self):
        """当前工作树 manifest 可导入；已生成的 references/commands.md 与之一致（幂等门禁）。"""
        specs = gen.load_specs()
        md = gen.render_markdown(specs)
        doc = PKG / "references" / "commands.md"
        if doc.is_file():
            self.assertEqual(doc.read_text(encoding="utf-8"), md,
                             "references/commands.md 与当前 manifest 不一致，请重新生成")
        # 已登记动词必须能在表里找到
        for spec in specs:
            self.assertIn(f"`{spec.verb}`", md)

    def test_main_empty_manifest_exits_2(self):
        """空 manifest 走命令行必须非零退出，不能静默出空表。"""
        old = gen.load_specs
        gen.load_specs = lambda: ()
        try:
            with tempfile.TemporaryDirectory() as d:
                rc = gen.main(["--out", str(Path(d) / "x.md")])
            self.assertEqual(rc, 2)
        finally:
            gen.load_specs = old


if __name__ == "__main__":
    unittest.main()
