"""#1 写回失败的现场复原：快照只覆盖本次写过的模块，复原失败要逐条报。"""

import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

import features.sync_push as sp  # noqa: E402


class FakeCodeModule:
    def __init__(self, text: str = "") -> None:
        self.text = text

    @property
    def CountOfLines(self) -> int:
        return len(self.text.splitlines())

    def Lines(self, start: int, count: int) -> str:
        lines = self.text.splitlines()
        return "\n".join(lines[start - 1:start - 1 + count])

    def DeleteLines(self, start: int, count: int) -> None:
        lines = self.text.splitlines()
        del lines[start - 1:start - 1 + count]
        self.text = "\n".join(lines)

    def AddFromString(self, code: str) -> None:
        self.text = f"{self.text}\n{code}" if self.text else code


class FakeComponent:
    def __init__(self, name: str, text: str = "") -> None:
        self.Name = name
        self.CodeModule = FakeCodeModule(text)


class FakeComponents:
    def __init__(self, comps: list[FakeComponent]) -> None:
        self._comps = comps

    def __iter__(self):
        return iter(self._comps)

    def Item(self, name: str):
        for comp in self._comps:
            if comp.Name == name:
                return comp
        raise KeyError(name)


class FakeProject:
    def __init__(self, comps: list[FakeComponent]) -> None:
        self.VBComponents = FakeComponents(comps)


class RollbackSnapshotTest(unittest.TestCase):
    def test_restore_puts_back_written_modules_and_reports_missing(self) -> None:
        a = FakeComponent("A", "Sub A()\nEnd Sub")
        b = FakeComponent("B", "Sub B()\nEnd Sub")
        project = FakeProject([a, b])

        snap = sp._snapshot_module_code(project)
        self.assertEqual(snap["A"], "Sub A()\nEnd Sub")

        # 模拟写回已经覆盖内存
        a.CodeModule.AddFromString("' 写坏了")
        b.CodeModule.DeleteLines(1, 2)
        project.VBComponents._comps.remove(b)

        restored, failed = sp._restore_module_code(project, snap)

        self.assertEqual(restored, ["A"])
        self.assertEqual(a.CodeModule.text.splitlines(), ["Sub A()", "End Sub"])
        self.assertEqual(len(failed), 1)
        self.assertIn("B", failed[0])
        self.assertIn("找不到", failed[0])


class SyncIntoRollbackTest(unittest.TestCase):
    def test_sync_into_rolls_back_only_when_flag_is_on(self) -> None:
        comp = FakeComponent("M", "Sub M()\nEnd Sub")
        project = FakeProject([comp])

        class FakeWb:
            Path = ""
            FullName = ""
            VBProject = project

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            src = root / "src"
            src.mkdir()
            proj = {"name": "P", "src": str(src)}
            defaults = {"common_dir": str(src)}

            def fake_import(vbproj, folder, doc_modules, only, tmp_parent=None):
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString("' 坏代码")
                return 1, 0, {"M": src / "M.bas"}

            def fake_postcheck(vbproj, written):
                return ["模块（写坏了）"]

            with mock.patch.object(sp, "_preflight_sheet_modules", lambda wb: None), \
                    mock.patch.object(sp, "_common_subset", lambda *a: (None, "")), \
                    mock.patch.object(sp, "_source_dirs", lambda *a: [src]), \
                    mock.patch.object(sp, "_import_folder", fake_import), \
                    mock.patch.object(sp, "component_snapshot", lambda vbproj: {"M": 1}), \
                    mock.patch.object(sp, "check_module_set", lambda *a: None), \
                    mock.patch.object(sp, "_postcheck_sync", fake_postcheck):
                bad, _notes = sp._writeback(FakeWb(), proj, defaults, doc_modules=False,
                                            rollback_on_fail=True)
                self.assertEqual(bad, ["模块（写坏了）"])
                self.assertEqual(comp.CodeModule.text.splitlines(), ["Sub M()", "End Sub"])

            # 默认关闭时不能悄悄改内存
            comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
            comp.CodeModule.AddFromString("Sub M()\nEnd Sub")
            with mock.patch.object(sp, "_preflight_sheet_modules", lambda wb: None), \
                    mock.patch.object(sp, "_common_subset", lambda *a: (None, "")), \
                    mock.patch.object(sp, "_source_dirs", lambda *a: [src]), \
                    mock.patch.object(sp, "_import_folder", fake_import), \
                    mock.patch.object(sp, "component_snapshot", lambda vbproj: {"M": 1}), \
                    mock.patch.object(sp, "check_module_set", lambda *a: None), \
                    mock.patch.object(sp, "_postcheck_sync", fake_postcheck):
                _bad, _notes = sp._writeback(FakeWb(), proj, defaults, doc_modules=False,
                                              rollback_on_fail=False)
                self.assertEqual(comp.CodeModule.text.splitlines(), ["' 坏代码"])


if __name__ == "__main__":
    unittest.main()
