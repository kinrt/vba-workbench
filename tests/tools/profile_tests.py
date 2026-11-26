# -*- coding: utf-8 -*-
"""只读跑一遍全量单测，输出最慢的 30 条及总耗时。

用法（在 tests 目录）：
  python tools/profile_tests.py
"""
from __future__ import annotations

import sys
import time
import unittest
from pathlib import Path

TESTS = Path(__file__).resolve().parents[1]
if str(TESTS) not in sys.path:
    sys.path.insert(0, str(TESTS))


class TimedResult(unittest.TestResult):
    def __init__(self) -> None:
        super().__init__()
        self.times: list[tuple[float, str]] = []
        self._started = 0.0

    def startTest(self, test: unittest.TestCase) -> None:
        self._started = time.perf_counter()
        super().startTest(test)

    def stopTest(self, test: unittest.TestCase) -> None:
        self.times.append((time.perf_counter() - self._started, str(test)))
        super().stopTest(test)


def main() -> int:
    suite = unittest.defaultTestLoader.discover(
        start_dir=str(TESTS), pattern="test_*.py", top_level_dir=str(TESTS))
    result = TimedResult()
    started = time.perf_counter()
    suite.run(result)
    elapsed = time.perf_counter() - started

    print(f"ran={result.testsRun} failures={len(result.failures)} "
          f"errors={len(result.errors)} skipped={len(result.skipped)} "
          f"elapsed={elapsed:.1f}s")
    for seconds, name in sorted(result.times, reverse=True)[:30]:
        print(f"{seconds:8.2f}s  {name}")
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    raise SystemExit(main())
