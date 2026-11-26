# -*- coding: utf-8 -*-
"""项目 → xlsm 解析：只认「基名 V<纯数字>」，取版本号最大的一份。"""

import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from toolkit.repo import resolve_target  # noqa: E402


class ResolveTargetTest(unittest.TestCase):
    def test_pick_max_and_ignore_noise(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            for name in ("示例表A V99.xlsm", "示例表A V100.xlsm",
                         "示例表A V100 (version 1).xlsm", "示例表A V100 - 副本.xlsm",
                         "示例表A 历史.xlsm"):
                (d / name).write_bytes(b"x")
            got, why = resolve_target({"name": "示例表A", "search_dir": str(d),
                                       "prefix": "示例表A", "require_version": True})
            self.assertEqual(got.name, "示例表A V100.xlsm")
            self.assertIn("V100", why)

    def test_same_size_and_mtime_versions_are_flagged(self):
        """同尺寸同 mtime 的版本候选不能静默只选最后一份。"""
        with tempfile.TemporaryDirectory() as d:
            d = Path(d)
            a, b = d / "示例表A V120.xlsm", d / "示例表A V121.xlsm"
            a.write_bytes(b"x")
            b.write_bytes(b"x")
            t = 1700000000000000000
            os.utime(a, ns=(t, t))
            os.utime(b, ns=(t, t))
            got, why = resolve_target({"name": "示例表A", "search_dir": str(d),
                                       "prefix": "示例表A", "require_version": True})
        self.assertEqual(got.name, "示例表A V121.xlsm")
        self.assertIn("同尺寸同 mtime", why)
        self.assertIn("重复副本", why)


        with tempfile.TemporaryDirectory() as d:
            (Path(d) / "示例表A 历史.xlsm").write_bytes(b"x")
            got, why = resolve_target({"search_dir": str(d), "prefix": "示例表A",
                                       "require_version": True})
            self.assertIsNone(got)
            self.assertIn("V<数字>", why)


if __name__ == "__main__":
    unittest.main()
