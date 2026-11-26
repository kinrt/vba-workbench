# -*- coding: utf-8 -*-
"""源码预检：硬错误必须阻断，软提醒只提示。"""

import sys
import tempfile
import unittest
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PKG / "scripts"))

import features.sync_push as sp  # noqa: E402
from toolkit import check_option_explicit  # noqa: E402


class SourcePreflightTest(unittest.TestCase):
    def _case(self, files: dict[str, bytes | str], *, common: bool = True):
        d = Path(tempfile.mkdtemp())
        self.addCleanup(lambda: __import__("shutil").rmtree(d, ignore_errors=True))
        proj_dir = d / "proj"
        common_dir = d / "common"
        proj_dir.mkdir()
        if common:
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
        fatal, notes = sp._source_preflight(proj, defaults, True)
        return fatal, notes

    @staticmethod
    def code(name: str = "测试模块", extra: str = "") -> str:
        return (f'Attribute VB_Name = "{name}"\n'
                "Option Explicit\n\n"
                "Public Sub 测试()\n"
                "    Debug.Print 1\n"
                "End Sub\n" + extra)

    def test_empty_module_is_fatal_but_document_module_is_not(self):
        fatal, _ = self._case({
            "空模块.bas": 'Attribute VB_Name = "空模块"\n',
            "Sheet1(说明).bas": 'Attribute VB_Name = "Sheet1"\n',
        })
        self.assertTrue(any("模块正文为空" in x and "空模块" in x for x in fatal))
        self.assertFalse(any("Sheet1" in x for x in fatal))

    def test_missing_frx_and_duplicate_are_fatal(self):
        form = ("VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 坏窗体\n"
                "   Caption         =   \"坏窗体\"\nEnd\nAttribute VB_Name = \"坏窗体\"\n")
        fatal, _ = self._case({
            "坏窗体.frm": form,
            "同名.bas": self.code("同名"),
            "同名.cls": "VERSION 1.0 CLASS\nBEGIN\nEND\n" + self.code("同名"),
        })
        self.assertTrue(any(".frx" in x and "坏窗体" in x for x in fatal))
        self.assertTrue(any("同一目录内模块重名" in x for x in fatal))

    def test_vb_name_mismatch_and_no_files_are_fatal(self):
        fatal, _ = self._case({
            "错名.bas": self.code("另一个名字"),
        })
        self.assertTrue(any("VB_Name" in x and "错名" in x for x in fatal))

        fatal2, _ = self._case({}, common=False)
        self.assertTrue(any("源码目录不存在" in x or "没有任何可写模块" in x for x in fatal2))

    def test_class_header_and_long_continuation_are_fatal(self):
        """.cls 缺类头、单语句续行超过 24 行，都要在写回前报 SOURCE_PREFLIGHT。"""
        fatal, _ = self._case({
            "缺类头.cls": self.code("缺类头"),
            "长续行.bas": self.code("长续行", "Public Sub 长()\n"
                                                  + "x = 1 & _\n" * 25
                                                  + "End Sub\n"),
        })
        self.assertTrue(any("VERSION 1.0 CLASS" in x and "缺类头" in x for x in fatal))
        self.assertTrue(any("续行超过 24 行" in x and "长续行" in x for x in fatal))

    def test_me_private_field_is_soft_note(self):
        """类模块里 Me.<私有字段> 只提醒，不阻断写回。"""
        cls = ("VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\n"
               'Attribute VB_Name = "私有字段类"\n'
               "Option Explicit\n\n"
               "Private mCount As Long\n\n"
               "Public Sub 加一()\n"
               "    Me.mCount = Me.mCount + 1\n"
               "End Sub\n")
        fatal, notes = self._case({"私有字段类.cls": cls})
        self.assertFalse(fatal, fatal)
        self.assertTrue(any("Me. 引用私有字段" in x and "私有字段类" in x for x in notes))

    def test_soft_warnings_do_not_block(self):
        bom = b"\xef\xbb\xbf" + self.code("BOM模块").encode("utf-8")
        crlf = self.code("CRLF模块").replace("\n", "\r\n").encode("utf-8")
        fatal, notes = self._case({
            "BOM模块.bas": bom,
            "CRLF模块.bas": crlf,
            "缺选项.bas": 'Attribute VB_Name = "缺选项"\nPublic Sub X()\nEnd Sub\n',
            "截断.bas": self.code("截断", "Public Sub 未完() _\n"),
        })
        self.assertFalse(fatal, fatal)
        text = "\n".join(notes)
        self.assertIn("BOM", text)
        self.assertIn("CRLF", text)
        self.assertIn("Option Explicit", text)
        self.assertIn("续行符", text)


    def test_style_and_preflight_share_option_explicit(self):
        """style 与预检必须用同一个判定，否则会“全过”和“16 个缺失”并存。"""
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            ok = d / "好.bas"
            ok.write_text('Attribute VB_Name = "好"\nOption Explicit\nPublic Sub A()\nEnd Sub\n',
                          encoding="utf-8", newline="\n")
            bad = d / "缺.bas"
            bad.write_text('Attribute VB_Name = "缺"\nPublic Sub B()\nEnd Sub\n',
                           encoding="utf-8", newline="\n")
            doc = d / "Sheet1(说明).bas"
            doc.write_text('Attribute VB_Name = "Sheet1"\nPublic Sub C()\nEnd Sub\n',
                           encoding="utf-8", newline="\n")
            miss = check_option_explicit([ok, bad, doc])
        self.assertEqual([x[0] for x in miss], [str(bad)])


if __name__ == "__main__":
    unittest.main()
