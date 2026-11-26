# -*- coding: utf-8 -*-
"""分层护栏（I8）：三层依赖方向违规即失败，全部用 AST 解析，不用 grep。

口径：
1. features 之间不得互相 import；唯一白名单 audit_b -> audit_engine（规则单源设计）。
2. toolkit / adapters 不得反向 import features 或入口壳；toolkit 可 import adapters
   （macro.py 函数内懒加载属允许项），adapters 可 import toolkit。
3. 三个入口壳（sync / agent / audit）顶层只允许 import toolkit / features；
   不得直连 adapters，三者互相 import 一律失败。

失败信息含：文件 + 行号 + 违规 import 原文。
"""
import ast
import shutil
import tempfile
import unittest
from dataclasses import dataclass
from pathlib import Path

PKG = Path(__file__).resolve().parents[1]
SCRIPTS = PKG / "scripts"

# 入口壳模块名（scripts 顶层三个 .py）
SHELL_MODULES = ("sync", "agent", "audit")

# 跨 feature 引用白名单：(来源 feature, 目标 feature) -> 理由
# audit_b -> audit_engine：审计规则单源设计，audit_b 委托共享审计引擎，避免规则双份维护。
# common_subset 是只依赖 toolkit 的共享叶子切片：list/pull/push/run/transform 与 agent_io
# 都委托它做公共模块子集判定与项目名查找，规则不在各切片里复制第二份；它不反向 import 任何 features。
FEATURE_WHITELIST = {
    ("audit_b", "audit_engine"),
    # audit_a 已收敛为参数适配层，规则委托共享引擎（口径同 audit_b -> audit_engine）。
    ("audit_a", "audit_engine"),
    # audit_engine 内聚拆分（同包叶子，非业务跨片互引）：audit_rules 为规则单源，
    # scan/structure/indicators 委托它；门面 audit_engine 组合三个 mixin。
    ("audit_engine", "audit_rules"),
    ("audit_engine", "audit_scan"),
    ("audit_engine", "audit_structure"),
    ("audit_engine", "audit_indicators"),
    ("audit_scan", "audit_rules"),
    ("audit_structure", "audit_rules"),
    ("audit_indicators", "audit_rules"),
    ("core", "common_subset"),
    ("run", "common_subset"),
    ("sync_push", "common_subset"),
    ("sync_read", "common_subset"),
    ("sync_transform", "common_subset"),
    ("agent_io", "common_subset"),
    # sync_push 结构拆分：预检 / 占位壳 / init 三段委托同包子模块；
    # 子模块只回引 common_subset 共享叶子，不反向依赖入口薄层。
    ("sync_push", "sync_preflight"),
    ("sync_push", "sync_shells"),
    ("sync_push", "sync_init"),
    ("sync_shells", "common_subset"),
    # run / agent_io 内聚拆分（同包叶子，非业务跨片互引）：run_macro 为宏执行装配单源，
    # run_watch 委托它做单文件写回；agent_probe 为项目解析单源，agent_logs 委托它。
    ("run", "run_macro"),
    ("run", "run_watch"),
    ("run_macro", "common_subset"),
    ("run_watch", "common_subset"),
    ("run_watch", "run_macro"),
    ("agent_io", "agent_probe"),
    ("agent_io", "agent_logs"),
    ("agent_logs", "agent_probe"),
    ("agent_logs", "agent_io"),
}

# 底层（toolkit / adapters）禁止反向引用的目标包/壳名
FORBIDDEN_BOTTOM = ("features",) + SHELL_MODULES


@dataclass
class ImportRecord:
    path: Path
    lineno: int
    level: int
    module: str
    names: str
    line_text: str
    top_level: bool


def _source_lines(path: Path) -> list[str]:
    return path.read_text(encoding="utf-8").splitlines()


def _collect_records(path: Path, *, top_level_only: bool) -> list[ImportRecord]:
    """解析单个文件，抽出 Import / ImportFrom 记录。"""
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    for node in ast.walk(tree):
        for child in ast.iter_child_nodes(node):
            child.parent = node  # 便于判断 import 是否处于模块顶层
    lines = _source_lines(path)
    records: list[ImportRecord] = []

    def add(node: ast.AST, module: str, names: str, level: int, top_level: bool) -> None:
        lineno = getattr(node, "lineno", 1)
        line_text = lines[lineno - 1].strip() if lineno - 1 < len(lines) else ""
        records.append(ImportRecord(path, lineno, level, module, names, line_text, top_level))

    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            top_level = isinstance(node.parent, ast.Module)
            if top_level_only and not top_level:
                continue
            for alias in node.names:
                add(node, alias.name, alias.name, 0, top_level)
        elif isinstance(node, ast.ImportFrom):
            top_level = isinstance(node.parent, ast.Module)
            if top_level_only and not top_level:
                continue
            module = node.module or ""
            names = ", ".join(a.name for a in node.names)
            add(node, module, names, node.level, top_level)
    return records


def _segments(module: str) -> list[str]:
    return [p for p in module.split(".") if p]


def _rel(path: Path, root: Path) -> str:
    try:
        return str(path.relative_to(root))
    except ValueError:
        return str(path)


def _violation(root: Path, rec: ImportRecord, reason: str) -> str:
    return f"{_rel(rec.path, root)}:{rec.lineno}: {reason}: {rec.line_text}"


# ---------------------------------------------------------------------------
# 第一层：features 不得互相 import
# ---------------------------------------------------------------------------

def _feature_target(rec: ImportRecord, stems: set[str]) -> str | None:
    """返回被引用的 feature 名；不属于 feature 互引则返回 None。"""
    parts = _segments(rec.module)
    first_name = rec.names.split(",")[0].strip()

    if rec.level >= 1:
        # 同包相对导入：level 1 指向 features 内另一切片；level>=2 已逃出本包。
        # features 包无子包，level 1 的相对引用一律按跨片候选处理，再交白名单裁决。
        if rec.level != 1:
            return None
        return parts[0] if parts else first_name

    if parts and parts[0] == "features":
        candidate = parts[1] if len(parts) > 1 else first_name
        return candidate if candidate in stems else None
    if parts and parts[0] in stems:
        return parts[0]
    return None


def scan_features(root: Path) -> list[str]:
    features_dir = root / "features"
    if not features_dir.exists():
        return []
    stems = {p.stem for p in features_dir.glob("*.py") if p.name != "__init__.py"}
    violations: list[str] = []
    for path in sorted(features_dir.glob("*.py")):
        if path.name == "__init__.py":
            continue
        source = path.stem
        for rec in _collect_records(path, top_level_only=False):
            target = _feature_target(rec, stems)
            if target is None:
                continue
            if (source, target) in FEATURE_WHITELIST:
                continue
            violations.append(_violation(
                root, rec, f"feature 切片不得互相 import（{source} -> {target}）"))
    return violations


# ---------------------------------------------------------------------------
# 第二层：toolkit / adapters 不得反向 import features 或入口壳
# ---------------------------------------------------------------------------

def _forbidden_bottom(rec: ImportRecord) -> str | None:
    parts = _segments(rec.module)
    first_name = rec.names.split(",")[0].strip()

    if rec.level >= 1:
        # level>=2 的相对导入可逃出本包（如 ..features.x / ..agent）
        if rec.level < 2:
            return None
        if parts and parts[0] in FORBIDDEN_BOTTOM:
            return parts[0]
        if not parts and first_name in FORBIDDEN_BOTTOM:
            return first_name
        return None

    if parts and parts[0] in FORBIDDEN_BOTTOM:
        return parts[0]
    return None


def scan_bottom_up(root: Path) -> list[str]:
    violations: list[str] = []
    for layer in ("toolkit", "adapters"):
        layer_dir = root / layer
        if not layer_dir.exists():
            continue
        for path in sorted(layer_dir.glob("*.py")):
            if path.name == "__init__.py":
                continue
            for rec in _collect_records(path, top_level_only=False):
                bad = _forbidden_bottom(rec)
                if bad is not None:
                    violations.append(_violation(
                        root, rec, f"底层 {layer} 不得反向 import {bad}"))
    return violations


# ---------------------------------------------------------------------------
# 第三层：入口壳顶层只允许 import toolkit / features
# ---------------------------------------------------------------------------

def scan_shells(root: Path) -> list[str]:
    violations: list[str] = []
    for shell in SHELL_MODULES:
        path = root / f"{shell}.py"
        if not path.exists():
            continue
        for rec in _collect_records(path, top_level_only=True):
            parts = _segments(rec.module)
            first_name = rec.names.split(",")[0].strip()
            top = parts[0] if parts else first_name

            if top == "adapters":
                violations.append(_violation(root, rec, "入口壳不得直连 adapters"))
            elif top in SHELL_MODULES and top != shell:
                violations.append(_violation(
                    root, rec, f"入口壳之间不得互相 import（{shell} -> {top}）"))
    return violations


# ---------------------------------------------------------------------------
# 真实源码护栏
# ---------------------------------------------------------------------------

class ImportLayerGuardTest(unittest.TestCase):
    def test_features_do_not_cross_import(self):
        violations = scan_features(SCRIPTS)
        self.assertEqual(violations, [], "\n".join(violations))

    def test_bottom_layers_do_not_reverse_import(self):
        violations = scan_bottom_up(SCRIPTS)
        self.assertEqual(violations, [], "\n".join(violations))

    def test_shells_import_only_toolkit_or_features(self):
        violations = scan_shells(SCRIPTS)
        self.assertEqual(violations, [], "\n".join(violations))


# ---------------------------------------------------------------------------
# 护栏自身可信：临时目录造违规样本，同一套扫描函数必须确实检出
# ---------------------------------------------------------------------------

class GuardSelfCheckTest(unittest.TestCase):
    def setUp(self) -> None:
        self.root = Path(tempfile.mkdtemp())
        for layer in ("features", "toolkit", "adapters"):
            (self.root / layer).mkdir(parents=True, exist_ok=True)

    def tearDown(self) -> None:
        shutil.rmtree(self.root, ignore_errors=True)

    def test_layer1_detects_cross_feature_import(self):
        (self.root / "features" / "evil.py").write_text(
            "from .other import BadThing\n", encoding="utf-8")
        violations = scan_features(self.root)
        self.assertTrue(violations, "跨 feature 引用必须被检出")
        self.assertTrue(any("evil.py" in v and "from .other import BadThing" in v for v in violations))

    def test_layer2_detects_reverse_from_toolkit_and_adapters(self):
        (self.root / "toolkit" / "evil.py").write_text(
            "from features import other\n", encoding="utf-8")
        (self.root / "adapters" / "evil.py").write_text(
            "import features.other\n", encoding="utf-8")
        violations = scan_bottom_up(self.root)
        self.assertEqual(len(violations), 2, violations)
        self.assertTrue(any("toolkit" in v and "from features import other" in v for v in violations))
        self.assertTrue(any("adapters" in v and "import features.other" in v for v in violations))

    def test_layer2_detects_shell_reference_from_bottom(self):
        (self.root / "toolkit" / "evil.py").write_text(
            "import agent\n", encoding="utf-8")
        violations = scan_bottom_up(self.root)
        self.assertTrue(any("agent" in v for v in violations))

    def test_layer3_detects_adapters_leak_and_shell_cross_import(self):
        (self.root / "agent.py").write_text(
            "from adapters import excel_gateway\nimport sync\n", encoding="utf-8")
        violations = scan_shells(self.root)
        self.assertEqual(len(violations), 2, violations)
        self.assertTrue(any("adapters" in v for v in violations))
        self.assertTrue(any("agent -> sync" in v for v in violations))

    def test_allowed_patterns_are_not_flagged(self):
        # toolkit 函数内懒加载 adapters、壳顶层引 toolkit，均属允许口径
        (self.root / "toolkit" / "macro.py").write_text(
            "def run():\n    from adapters import excel_gateway\n", encoding="utf-8")
        (self.root / "agent.py").write_text(
            "from toolkit.render import render\n", encoding="utf-8")
        self.assertEqual(scan_bottom_up(self.root), [])
        self.assertEqual(scan_shells(self.root), [])

    def test_whitelisted_pair_is_not_flagged(self):
        # 白名单例外：audit_b -> audit_engine 不得被第一层拦下
        (self.root / "features" / "audit_b.py").write_text(
            "from .audit_engine import AuditEngine\n", encoding="utf-8")
        (self.root / "features" / "audit_engine.py").write_text(
            "class AuditEngine:\n    pass\n", encoding="utf-8")
        self.assertEqual(scan_features(self.root), [])


if __name__ == "__main__":
    unittest.main()
