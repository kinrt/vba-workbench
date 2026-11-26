# -*- coding: utf-8 -*-
"""vbaProject.bin 读取：坏文件要给人话错误，不能抛 traceback。"""

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from toolkit.base import Failure  # noqa: E402
from toolkit.vbabin import read_vba_bin  # noqa: E402


class VbaBinTest(unittest.TestCase):
    def test_bad_zip_gives_failure(self):
        with tempfile.TemporaryDirectory() as d:
            f = Path(d) / "坏表.xlsm"
            f.write_text("not a zip", encoding="utf-8")
            with self.assertRaises(Failure):
                read_vba_bin(f)

    def test_zip_without_bin_gives_failure(self):
        import zipfile
        with tempfile.TemporaryDirectory() as d:
            f = Path(d) / "空壳.xlsm"
            with zipfile.ZipFile(f, "w") as z:
                z.writestr("xl/workbook.xml", "<workbook/>")
            with self.assertRaises(Failure):
                read_vba_bin(f)


if __name__ == "__main__":
    unittest.main()
