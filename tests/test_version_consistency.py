# -*- coding: utf-8 -*-
"""SKILL.md 的 metadata.version 必须等于 VERSION，避免包版本漂移。"""

import re
import unittest
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]


class VersionConsistencyTest(unittest.TestCase):
    def test_skill_metadata_matches_version_file(self):
        version = (PKG / "VERSION").read_text(encoding="utf-8").strip()
        skill = (PKG / "SKILL.md").read_text(encoding="utf-8")
        m = re.search(r'version:\s*"([^"]+)"', skill)
        self.assertIsNotNone(m, "SKILL.md frontmatter 里没有 metadata.version")
        self.assertEqual(m.group(1), version)


if __name__ == "__main__":
    unittest.main()
