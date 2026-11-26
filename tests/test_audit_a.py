# -*- coding: utf-8 -*-
"""audit_a 切片契约：members / common / project 三个只读视角。

四类覆盖（对齐方案 §4.6 测试矩阵）：
  · 成功：exit 0，信封键与 1.5.6 一致；
  · 业务失败：exit 1，errors[].code 为已登记错误码；
  · 用法错：exit 2，stdout 不出 JSON 信封；
  · --json 单行纯度：stdout 恰好一行合法 JSON。

另单独锁定 common 的 exit 3 语义：公共库未配置时 ok=true、errors=[]、exit_code=3，
不许归一成 1。测试全部走临时目录 + 桩源码，不启动 Excel、不碰业务表。
"""
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stdout, redirect_stderr
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
AUDIT = PKG / "scripts" / "audit.py"

sys.path.insert(0, str(PKG / "scripts"))
from toolkit.commands import specs_for  # noqa: E402
from toolkit.registry import build_parser  # noqa: E402

COMMON_SRC = (
    'Attribute VB_Name = "公共"\n'
    "Option Explicit\n"
    "Public Function AddOne(ByVal x As Long) As Long\n"
    "    AddOne = x + 1\n"
    "End Function\n"
    "Public Sub 无人调用()\n"
    "End Sub\n"
)

PROJECT_SRC = (
    'Attribute VB_Name = "调用"\n'
    "Option Explicit\n"
    "Public Sub 入口()\n"
    "    Dim n As Long\n"
    "    n = AddOne(1)\n"
    "End Sub\n"
)

# 只有声明、解析不出任何成员的桩模块：走 SOURCE_PREFLIGHT 业务失败
EMPTY_SRC = "Option Explicit\n' 只有注释，没有过程\n"


def _write_sandbox(root: Path, *, common: str | None = COMMON_SRC,
                   project: str | None = PROJECT_SRC) -> None:
    (root / "src" / "common").mkdir(parents=True, exist_ok=True)
    (root / "src" / "projects" / "示例表A").mkdir(parents=True, exist_ok=True)
    if common is not None:
        (root / "src" / "common" / "公共.bas").write_text(common, encoding="utf-8", newline="\n")
    if project is not None:
        (root / "src" / "projects" / "示例表A" / "调用.bas").write_text(
            project, encoding="utf-8", newline="\n")
    (root / "mapping.json").write_text(json.dumps({
        "defaults": {"common_dir": "src/common"},
        "projects": [{"name": "示例表A", "src": "src/projects/示例表A",
                      "require_version": False}],
    }, ensure_ascii=False), encoding="utf-8", newline="\n")


class _AuditSandbox(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        _write_sandbox(self.root)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def run_audit(self, *argv: str) -> subprocess.CompletedProcess:
        env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(self.root / "mapping.json"),
                   PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
        return subprocess.run([sys.executable, str(AUDIT), *argv],
                              cwd=str(self.root), env=env, capture_output=True,
                              text=True, encoding="utf-8")

    @staticmethod
    def _parse(payload: str) -> dict:
        return json.loads(payload.strip())


class MembersTest(_AuditSandbox):
    def test_members_success(self):
        r = self.run_audit("members", "--module", "公共", "--json")
        self.assertEqual(r.returncode, 0, r.stderr)
        payload = self._parse(r.stdout)
        self.assertEqual(payload["action"], "audit.members")
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["exit_code"], 0)
        self.assertEqual(payload["stats"]["module"], "公共")
        self.assertEqual(payload["stats"]["members"], 2)
        self.assertEqual(len(payload["details"]), 2)
        self.assertEqual(payload["errors"], [])

    def test_members_business_failure_exit1_with_code(self):
        r = self.run_audit("members", "--module", "不存在的模块", "--json")
        self.assertEqual(r.returncode, 1, r.stderr)
        payload = self._parse(r.stdout)
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["exit_code"], 1)
        self.assertTrue(payload["errors"])
        self.assertEqual(payload["errors"][0]["code"], "TARGET_NOT_FOUND")
        self.assertEqual(payload["errors"][0]["stage"], "resolve")

    def test_members_usage_error_exit2_without_json(self):
        r = self.run_audit("members", "--json")
        self.assertEqual(r.returncode, 2, r.stderr)
        self.assertFalse(r.stdout.strip().startswith("{"))

    def test_members_json_output_is_single_line(self):
        r = self.run_audit("members", "--module", "公共", "--json")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertEqual(len(r.stdout.strip().splitlines()), 1, r.stdout)
        self._parse(r.stdout)

    def test_members_rejects_project_flag(self):
        # 1.5.6 members 没有 --project；传入必须在解析层 exit 2
        r = self.run_audit("members", "--module", "公共", "--project", "示例表A", "--json")
        self.assertEqual(r.returncode, 2, r.stderr)
        self.assertFalse(r.stdout.strip().startswith("{"))

    def test_members_empty_source_is_preflight_failure(self):
        (self.root / "src" / "common" / "空模块.bas").write_text(
            EMPTY_SRC, encoding="utf-8", newline="\n")
        r = self.run_audit("members", "--module", "空模块", "--json")
        self.assertEqual(r.returncode, 1, r.stderr)
        payload = self._parse(r.stdout)
        self.assertEqual(payload["exit_code"], 1)
        self.assertEqual(payload["errors"][0]["code"], "SOURCE_PREFLIGHT")
        self.assertEqual(payload["errors"][0]["stage"], "parse")


class CommonTest(_AuditSandbox):
    def test_common_success(self):
        r = self.run_audit("common", "--json")
        self.assertEqual(r.returncode, 0, r.stderr)
        payload = self._parse(r.stdout)
        self.assertEqual(payload["action"], "audit.common")
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["exit_code"], 0)
        self.assertTrue(payload["stats"]["configured"])
        self.assertEqual(payload["stats"]["modules"], 1)
        self.assertEqual(payload["details"][0]["module"], "公共")
        self.assertEqual(payload["errors"], [])

    def test_common_accepts_project_filter(self):
        # 1.5.6 common 接受可选 --project 过滤；新版不得在解析层 exit 2
        r = self.run_audit("common", "--project", "示例表A", "--json")
        self.assertNotEqual(r.returncode, 2, r.stderr)
        self.assertIn(r.returncode, (0, 1))
        payload = self._parse(r.stdout)
        self.assertEqual(payload["action"], "audit.common")

    def _write_project_b(self) -> None:
        """补第二张表 示例表B：同样调用公共 AddOne，供重复 --project 用例对照。"""
        b_dir = self.root / "src" / "projects" / "示例表B"
        b_dir.mkdir(parents=True, exist_ok=True)
        (b_dir / "调用B.bas").write_text(
            'Attribute VB_Name = "调用B"\n'
            "Option Explicit\n"
            "Public Sub 入口()\n"
            "    Dim n As Long\n"
            "    n = AddOne(1)\n"
            "End Sub\n",
            encoding="utf-8", newline="\n")
        cfg = json.loads((self.root / "mapping.json").read_text(encoding="utf-8"))
        cfg["projects"].append({"name": "示例表B", "src": "src/projects/示例表B",
                                "require_version": False})
        (self.root / "mapping.json").write_text(
            json.dumps(cfg, ensure_ascii=False), encoding="utf-8", newline="\n")

    def test_common_project_is_repeatable_append(self):
        """--project 可重复传入，解析后收集为列表（对齐 1.5.6 append 口径）。"""
        parser = build_parser(specs_for("audit"), prog="audit.py")
        multi = parser.parse_args(["common", "--project", "示例表A",
                                   "--project", "示例表B"])
        self.assertEqual(multi.project, ["示例表A", "示例表B"])
        single = parser.parse_args(["common", "--project", "示例表A"])
        self.assertEqual(single.project, ["示例表A"])
        none = parser.parse_args(["common"])
        self.assertIsNone(none.project)

    def test_common_repeated_project_audits_both(self):
        # 1.5.6 --project 为 append：重复传入两张表都纳入扫描，不得被后者覆盖收窄
        self._write_project_b()
        r = self.run_audit("common", "--project", "示例表A",
                           "--project", "示例表B", "--json")
        self.assertEqual(r.returncode, 0, r.stderr)
        payload = self._parse(r.stdout)
        self.assertEqual(payload["action"], "audit.common")
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["exit_code"], 0)
        self.assertTrue(payload["stats"]["configured"])
        # 公共 AddOne 被 A、B 两个项目文件引用，引用方计数为 2，证明两表都被审计
        self.assertEqual(payload["details"][0]["ref_files"], 2)
        self.assertEqual(payload["errors"], [])

    def test_common_single_project_stays_narrow(self):
        # 只传一次：等价于 [该项目]，范围仍收窄到这一张表，不扫磁盘上的其余项目
        self._write_project_b()
        r = self.run_audit("common", "--project", "示例表A", "--json")
        self.assertEqual(r.returncode, 0, r.stderr)
        payload = self._parse(r.stdout)
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["details"][0]["ref_files"], 1)
        self.assertEqual(payload["errors"], [])

    def test_common_without_project_still_runs(self):
        # 缺省 --project 仍按旧版全量口径工作
        r = self.run_audit("common", "--json")
        self.assertEqual(r.returncode, 0, r.stderr)
        payload = self._parse(r.stdout)
        self.assertTrue(payload["stats"]["configured"])

    def test_common_business_failure_exit1_with_code(self):
        # 配置文件损坏：派发前就落配置失败，退出 1；裸 Failure 无 code 统一归 UNEXPECTED
        (self.root / "mapping.json").write_text("{ 不是合法 JSON", encoding="utf-8")
        r = self.run_audit("common", "--json")
        self.assertEqual(r.returncode, 1, r.stderr)
        payload = self._parse(r.stdout)
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["exit_code"], 1)
        self.assertTrue(payload["errors"])
        self.assertEqual(payload["errors"][0]["code"], "UNEXPECTED")
        self.assertEqual(payload["errors"][0]["stage"], "unknown")

    def test_common_usage_error_exit2_without_json(self):
        r = self.run_audit("common", "--不存在的开关", "--json")
        self.assertEqual(r.returncode, 2, r.stderr)
        self.assertFalse(r.stdout.strip().startswith("{"))

    def test_common_json_output_is_single_line(self):
        r = self.run_audit("common", "--json")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertEqual(len(r.stdout.strip().splitlines()), 1, r.stdout)
        self._parse(r.stdout)

    def test_common_not_configured_keeps_exit3_semantics(self):
        # 清空公共库：只读警告，exit 3；ok=true、errors=[]，不许归一成 1
        for f in (self.root / "src" / "common").glob("*"):
            f.unlink()
        r = self.run_audit("common", "--json")
        self.assertEqual(r.returncode, 3, r.stderr)
        payload = self._parse(r.stdout)
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["errors"], [])
        self.assertEqual(payload["exit_code"], 3)
        self.assertFalse(payload["stats"]["configured"])
        self.assertEqual(payload["stats"]["config_key"], "defaults.common_dir")
        self.assertIn("common", payload["stats"]["common_dir"].replace("\\", "/"))
        self.assertEqual(payload["details"], [])


class ProjectTest(_AuditSandbox):
    def test_project_success(self):
        r = self.run_audit("project", "--project", "示例表A", "--json")
        self.assertEqual(r.returncode, 0, r.stderr)
        payload = self._parse(r.stdout)
        self.assertEqual(payload["action"], "audit.project")
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["exit_code"], 0)
        self.assertEqual(payload["stats"]["project"], "示例表A")
        self.assertEqual(payload["stats"]["modules"], 1)
        self.assertEqual(payload["stats"]["unused"], [])
        self.assertEqual(payload["details"][0]["module"], "公共")
        self.assertEqual(payload["errors"], [])

    def test_project_business_failure_exit1_with_code(self):
        r = self.run_audit("project", "--project", "不存在的表", "--json")
        self.assertEqual(r.returncode, 1, r.stderr)
        payload = self._parse(r.stdout)
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["exit_code"], 1)
        self.assertTrue(payload["errors"])
        self.assertEqual(payload["errors"][0]["code"], "PROJECT_NOT_FOUND")
        self.assertEqual(payload["errors"][0]["stage"], "resolve")

    def test_project_usage_error_exit2_without_json(self):
        r = self.run_audit("project", "--json")
        self.assertEqual(r.returncode, 2, r.stderr)
        self.assertFalse(r.stdout.strip().startswith("{"))

    def test_project_json_output_is_single_line(self):
        r = self.run_audit("project", "--project", "示例表A", "--json")
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertEqual(len(r.stdout.strip().splitlines()), 1, r.stdout)
        self._parse(r.stdout)


class InterruptSignalTest(unittest.TestCase):
    """派发期间 Ctrl+C：audit 入口必须返回 130、提示「已中断」，不出 JSON 信封。"""

    def _run_interrupt(self, tmp: Path, *argv: str):
        import audit as audit_entry  # noqa: E402
        from toolkit.base import set_json_mode  # noqa: E402

        cfg = tmp / "mapping.json"
        cfg.write_text(json.dumps({"projects": []}, ensure_ascii=False), encoding="utf-8")
        old_env = os.environ.get("VBA_WORKBENCH_CONFIG")
        os.environ["VBA_WORKBENCH_CONFIG"] = str(cfg)
        orig = audit_entry.dispatch

        def _boom(specs, args, ctx):
            raise KeyboardInterrupt

        audit_entry.dispatch = _boom
        buf, err = io.StringIO(), io.StringIO()
        try:
            with redirect_stdout(buf), redirect_stderr(err):
                code = audit_entry.main(list(argv))
        finally:
            audit_entry.dispatch = orig
            set_json_mode(False)
            if old_env is None:
                os.environ.pop("VBA_WORKBENCH_CONFIG", None)
            else:
                os.environ["VBA_WORKBENCH_CONFIG"] = old_env
        return code, buf.getvalue(), err.getvalue()

    def test_audit_keyboard_interrupt_exits_130_without_envelope(self):
        with tempfile.TemporaryDirectory() as tmp:
            d = Path(tmp)
            # 非 JSON：人话含「已中断」，stdout 不出信封
            code, out, err = self._run_interrupt(d, "members", "--module", "公共")
            self.assertEqual(code, 130, err)
            self.assertIn("已中断", out)
            self.assertFalse(out.strip().startswith("{"))
            # JSON 模式：同样不得吐半截信封，「已中断」走 stderr
            code2, out2, err2 = self._run_interrupt(
                d, "members", "--module", "公共", "--json")
            self.assertEqual(code2, 130, err2)
            self.assertFalse(out2.strip(), f"JSON 中断不得出 stdout：{out2!r}")
            self.assertIn("已中断", err2)


if __name__ == "__main__":
    unittest.main()
