# -*- coding: utf-8 -*-
"""写回后回读：内容被截断时必须返回带行数的差异明细。"""

import sys
import tempfile
import unittest
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))

import features.sync_push as sp  # noqa: E402


class _FakeCode:
    def __init__(self, text: str):
        self.text = text
        self.CountOfLines = len(text.splitlines())

    def Lines(self, start: int, count: int) -> str:
        return self.text


class _FakeComp:
    def __init__(self, text: str):
        self.CodeModule = _FakeCode(text)


class _FakeComponents:
    def __init__(self, rows: dict[str, _FakeComp]):
        self.rows = rows

    def Item(self, name: str):
        return self.rows[name]


class _FakeProject:
    def __init__(self, rows: dict[str, _FakeComp]):
        self.VBComponents = _FakeComponents(rows)


class PostcheckTest(unittest.TestCase):
    def _src(self, text: str) -> Path:
        d = Path(tempfile.mkdtemp())
        self.addCleanup(lambda: __import__("shutil").rmtree(d, ignore_errors=True))
        p = d / "模块.bas"
        p.write_text(text, encoding="utf-8", newline="\n")
        return p

    def test_matching_text_passes(self):
        src = self._src('Attribute VB_Name = "模块"\nOption Explicit\n\nPublic Sub A()\nEnd Sub\n')
        vbe = "Option Explicit\n\nPublic Sub A()\nEnd Sub\n"
        bad = sp._postcheck_sync(_FakeProject({"模块": _FakeComp(vbe)}), {"模块": src})
        self.assertEqual(bad, [])

    def test_truncated_text_reports_line_counts_and_first_difference(self):
        src = self._src('Attribute VB_Name = "模块"\n'
                        "Option Explicit\n\nPublic Sub A()\n    Debug.Print 1\nEnd Sub\n")
        vbe = "Option Explicit\n\nPublic Sub A()\n"
        bad = sp._postcheck_sync(_FakeProject({"模块": _FakeComp(vbe)}), {"模块": src})
        self.assertTrue(bad)
        self.assertIn("源码 4 行", bad[0])
        self.assertIn("VBE 2 行", bad[0])
        self.assertIn("首个差异第 3 行", bad[0])

    def test_process_attribute_lines_are_ignored(self):
        src = self._src('Attribute VB_Name = "模块"\n'
                        "Option Explicit\n\nPublic Sub A()\n"
                        'Attribute A.VB_ProcData.VB_Invoke_Func = "P\\n14"\n'
                        "End Sub\n")
        vbe = "Option Explicit\n\nPublic Sub A()\nEnd Sub\n"
        bad = sp._postcheck_sync(_FakeProject({"模块": _FakeComp(vbe)}), {"模块": src})
        self.assertEqual(bad, [])

    def test_trailing_bare_parentheses_placeholder_is_ignored(self):
        src = self._src('Attribute VB_Name = "模块"\nOption Explicit\n\nPublic Sub A()\nEnd Sub\n')
        vbe = "Option Explicit\n\nPublic Sub A()\nEnd Sub\n()\n"
        bad = sp._postcheck_sync(_FakeProject({"模块": _FakeComp(vbe)}), {"模块": src})
        self.assertEqual(bad, [])

    def test_mid_module_bare_parentheses_is_not_ignored(self):
        """白名单只限模块最末；正文中间的 `()` 是异常，不能静默放过。"""
        src = self._src('Attribute VB_Name = "模块"\nOption Explicit\n\nPublic Sub A()\nEnd Sub\n\nPublic Sub B()\nEnd Sub\n')
        vbe = "Option Explicit\n\nPublic Sub A()\nEnd Sub\n()\n\nPublic Sub B()\nEnd Sub\n"
        bad = sp._postcheck_sync(_FakeProject({"模块": _FakeComp(vbe)}), {"模块": src})
        self.assertTrue(bad)

    def test_missing_component_is_reported(self):
        src = self._src('Attribute VB_Name = "模块"\nPublic Sub A()\nEnd Sub\n')
        bad = sp._postcheck_sync(_FakeProject({}), {"模块": src})
        self.assertTrue(any("工程里找不到" in x for x in bad))


if __name__ == "__main__":
    unittest.main()
