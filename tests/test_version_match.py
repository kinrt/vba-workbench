# -*- coding: utf-8 -*-
"""纯函数自测：版本号解析与项目基名匹配（不需要 Excel、不需要 COM）。"""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from toolkit.repo import parse_stem_version  # noqa: E402


class VersionMatchTest(unittest.TestCase):
    def test_plain_version(self):
        # 返回 (版本号, 命中的原文片段)
        self.assertEqual(parse_stem_version("示例表A V100"), (100, "V100"))

    def test_version_is_last_number(self):
        self.assertEqual(parse_stem_version("示例表C V2026"), (2026, "V2026"))

    def test_no_version_returns_none(self):
        self.assertIsNone(parse_stem_version("示例表A"))
        self.assertIsNone(parse_stem_version("示例表A 副本"))


if __name__ == "__main__":
    unittest.main()
