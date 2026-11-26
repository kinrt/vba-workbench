# -*- coding: utf-8 -*-
"""旧 1.5.6 vs 新实现的差分 harness（命令面 + 解析层场景）。

对比口径（设计方案 A.6.2 三层对比的本切片子集；副作用层由 Excel 集成测试覆盖）：

| 层 | 比什么 |
|---|---|
| 命令集合 | 旧 ledger 动词集合 vs 新 manifest 动词集合（按入口） |
| 参数集合 | 每个动词的参数名、类型、取值范围、默认值 |
| JSON 能力 | 是否支持 ``--json``；信封键集合 |
| 退出码 | 同一解析层场景在旧/新副本上的退出码 |
| 输出结构 | 信封键集合、``action`` 前缀 ``<入口>.<动词>`` |

判定：

- ``match``：两边等价；
- ``approved``：命中 ``approved_deltas.json`` 白名单（有意修复/规范变更）；
- ``pending``：命中迁移待办登记（manifest 未齐，显式已知缺口，不算回归）；
- ``failed``：白名单与待办之外的差异，harness 整体判失败。

不启动 Excel、不碰业务表；场景在临时目录 + 桩配置里跑。

用法：

```
python tests/tools/diff_harness.py [--report tests/baseline/1.5.6/diff-report.json]
```
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

PKG = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(PKG / "scripts"))
sys.path.insert(0, str(PKG / "tests" / "baseline"))

import record_1_5_6 as recorder  # noqa: E402

BASELINE_DIR = PKG / "tests" / "baseline" / "1.5.6"
ENTRY_SCRIPT = recorder.ENTRY_SCRIPT
ENTRIES = recorder.ENTRIES

# 骨架 build_parser 统一附加的公共参数；不进业务参数集合 diff（单独比能力位）。
COMMON_OPTIONS = {"json", "no-input", "apply", "project", "all-project", "all_project"}

STATUS_MATCH = "match"
STATUS_APPROVED = "approved"
STATUS_PENDING = "pending"
STATUS_FAILED = "failed"


# ---------------------------------------------------------------------------
# 两侧命令面
# ---------------------------------------------------------------------------
def _old_surface(ledger: dict[str, Any]) -> dict[str, dict[str, dict[str, Any]]]:
    """从旧版 ledger 整理成 {entry: {verb: {args, json_ok, apply, project, all_project}}}。"""
    surface: dict[str, dict[str, dict[str, Any]]] = {}
    for entry, data in ledger.get("entries", {}).items():
        verbs: dict[str, dict[str, Any]] = {}
        for verb, info in data.get("verbs", {}).items():
            args = sorted(name for a in info.get("args", []) for name in a["names"]
                           if name not in COMMON_OPTIONS)
            detail = {a["names"][-1]: a for a in info.get("args", [])
                      if a["names"][-1] not in COMMON_OPTIONS}
            verbs[verb] = {
                "args": args,
                "arg_detail": detail,
                "json_ok": bool(info.get("json_ok")),
                "apply": bool(info.get("apply")),
                "project": bool(info.get("project")),
                "all_project": bool(info.get("all_project")),
            }
        surface[entry] = verbs
    return surface


def _new_surface(specs: tuple) -> dict[str, dict[str, dict[str, Any]]]:
    """从新 manifest 整理成与旧侧同形的命令面。

    口径：参数的生效默认值、类型、取值范围一律走 ``build_parser`` 生成的
    argparse 对象 introspect（与旧侧同一路径），不直接读 raw ArgSpec ——
    否则 flag 的 ``default=None`` 会被误判成与旧版 ``False`` 不一致，
    ``kind=path`` 也应按 argparse 的 ``type=str`` 对齐。
    """
    from toolkit.registry import build_parser

    exit_codes_by_verb = {(s.entry, s.verb): list(s.exit_codes) for s in specs}
    json_by_verb = {(s.entry, s.verb): bool(s.supports_json) for s in specs}
    surface: dict[str, dict[str, dict[str, Any]]] = {}
    for entry in ENTRIES:
        entry_specs = tuple(s for s in specs if s.entry == entry)
        parser = build_parser(entry_specs, prog=f"{entry}.py")
        intro = recorder.introspect_parser(parser)
        verbs: dict[str, dict[str, Any]] = {}
        for verb, info in intro["verbs"].items():
            args = sorted(name for a in info["args"] for name in a["names"]
                           if name not in COMMON_OPTIONS)
            detail = {a["names"][-1]: a for a in info["args"]
                      if a["names"][-1] not in COMMON_OPTIONS}
            verbs[verb] = {
                "args": args,
                "arg_detail": detail,
                # json 位以 manifest 登记为准：骨架对不支持 --json 的动词仍挂
                # SUPPRESS 选项，不能靠选项是否存在判定。
                "json_ok": json_by_verb.get((entry, verb), info["json_ok"]),
                "apply": bool(info["apply"]),
                "project": bool(info["project"]),
                "all_project": bool(info["all_project"]),
                "exit_codes": exit_codes_by_verb.get((entry, verb), []),
            }
        surface[entry] = verbs
    return surface


# ---------------------------------------------------------------------------
# 白名单 / 待办登记
# ---------------------------------------------------------------------------
def load_deltas(path: Path) -> dict[str, Any]:
    if not path.is_file():
        return {"approved_deltas": [], "pending": []}
    data = json.loads(path.read_text(encoding="utf-8"))
    data.setdefault("approved_deltas", [])
    data.setdefault("pending", [])
    return data


def _matches(entry: str, verb: str, field: str, input_argv: list[str],
             rule: dict[str, Any]) -> bool:
    if rule.get("entry", entry) != entry or rule.get("verb", verb) != verb:
        return False
    if rule.get("field") and rule["field"] != field:
        return False
    rule_input = rule.get("input")
    if rule_input is not None and rule_input != " ".join(input_argv):
        return False
    return True


def classify(row: dict[str, Any], deltas: dict[str, Any]) -> str:
    """按白名单/待办给差异行定性；等价行直接 match。"""
    if row["status"] == STATUS_MATCH:
        return STATUS_MATCH
    for rule in deltas.get("approved_deltas", []):
        if _matches(row["entry"], row["verb"], row["field"], row["argv"], rule):
            return STATUS_APPROVED
    for rule in deltas.get("pending", []):
        if _matches(row["entry"], row["verb"], row["field"], row["argv"], rule):
            return STATUS_PENDING
    return STATUS_FAILED


# ---------------------------------------------------------------------------
# 静态命令面对比
# ---------------------------------------------------------------------------
def _norm_default(kind: str, value: Any) -> Any:
    # append 的缺省 None 与 [] 同义（无可重复项），归一后再比。
    if kind == "append":
        return value or []
    return value


def _detail_diff(old_detail: dict, new_detail: dict) -> list[dict[str, Any]]:
    """逐参数比类型/取值/默认值；返回差异小行。"""
    rows: list[dict[str, Any]] = []
    for name in sorted(set(old_detail) | set(new_detail)):
        o = old_detail.get(name)
        n = new_detail.get(name)
        if o is None:
            rows.append({"field": "arg-added", "arg": name, "old": None, "new": n})
        elif n is None:
            rows.append({"field": "arg-removed", "arg": name, "old": o, "new": None})
        else:
            for key in ("kind", "choices"):
                if o.get(key) != n.get(key):
                    rows.append({"field": f"arg-{key}", "arg": name,
                                 "old": o.get(key), "new": n.get(key)})
            o_default = _norm_default(o.get("kind"), o.get("default"))
            n_default = _norm_default(n.get("kind"), n.get("default"))
            if o_default != n_default:
                rows.append({"field": "arg-default", "arg": name,
                             "old": o.get("default"), "new": n.get("default")})
    return rows


def compare_surface(old: dict, new: dict) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for entry in ENTRIES:
        old_verbs = old.get(entry, {})
        new_verbs = new.get(entry, {})
        for verb in sorted(set(old_verbs) - set(new_verbs)):
            rows.append({
                "entry": entry, "verb": verb, "argv": [], "field": "command-missing-in-new",
                "old": sorted(old_verbs[verb]["args"]), "new": [], "status": STATUS_FAILED,
            })
        for verb in sorted(set(new_verbs) - set(old_verbs)):
            rows.append({
                "entry": entry, "verb": verb, "argv": [], "field": "command-added-in-new",
                "old": [], "new": sorted(new_verbs[verb]["args"]), "status": STATUS_FAILED,
            })
        for verb in sorted(set(old_verbs) & set(new_verbs)):
            o, n = old_verbs[verb], new_verbs[verb]
            if o["args"] != n["args"]:
                rows.append({
                    "entry": entry, "verb": verb, "argv": [], "field": "args",
                    "old": o["args"], "new": n["args"], "status": STATUS_FAILED,
                })
            for sub in _detail_diff(o["arg_detail"], n["arg_detail"]):
                rows.append({
                    "entry": entry, "verb": verb, "argv": [],
                    "field": sub["field"], "arg": sub["arg"],
                    "old": sub["old"], "new": sub["new"], "status": STATUS_FAILED,
                })
            if o["json_ok"] != n["json_ok"]:
                rows.append({
                    "entry": entry, "verb": verb, "argv": [], "field": "json",
                    "old": o["json_ok"], "new": n["json_ok"], "status": STATUS_FAILED,
                })
            if o["project"] != n["project"] or o["all_project"] != n["all_project"]:
                rows.append({
                    "entry": entry, "verb": verb, "argv": [], "field": "project",
                    "old": {"project": o["project"], "all_project": o["all_project"]},
                    "new": {"project": n["project"], "all_project": n["all_project"]},
                    "status": STATUS_FAILED,
                })
            else:
                rows.append({
                    "entry": entry, "verb": verb, "argv": [], "field": "surface",
                    "old": {"args": o["args"], "json": o["json_ok"],
                            "project": o["project"], "all_project": o["all_project"]},
                    "new": {"args": n["args"], "json": n["json_ok"],
                            "project": n["project"], "all_project": n["all_project"]},
                    "status": STATUS_MATCH,
                })
    return rows


# ---------------------------------------------------------------------------
# 动态场景对比（解析层，临时目录 + 桩配置）
# ---------------------------------------------------------------------------
def _run_copy(scripts_dir: Path, entry: str, argv: list[str], sandbox: Path) -> dict[str, Any]:
    script = scripts_dir / ENTRY_SCRIPT[entry]
    if not script.is_file():
        return {"exit_code": None, "envelope_keys": [], "action": "", "note": "脚本缺失"}
    return recorder.run_parse_scenario(script, argv, sandbox)


def compare_scenarios(old_scripts: Path, new_scripts: Path,
                      scenarios: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    with tempfile.TemporaryDirectory() as raw:
        root = Path(raw)
        for case in scenarios:
            entry = case["entry"]
            argv = case["argv"]
            old_sandbox = root / "old"
            new_sandbox = root / "new"
            old_sandbox.mkdir(exist_ok=True)
            new_sandbox.mkdir(exist_ok=True)
            old_run = _run_copy(old_scripts, entry, argv, old_sandbox)
            new_run = _run_copy(new_scripts, entry, argv, new_sandbox)
            verb = argv[0] if argv and not argv[0].startswith("__") else ""
            base = {"entry": entry, "verb": verb, "argv": argv}

            if old_run["exit_code"] != new_run["exit_code"]:
                rows.append({**base, "field": "exit_code",
                             "old": old_run["exit_code"], "new": new_run["exit_code"],
                             "status": STATUS_FAILED})
            if old_run["envelope_keys"] != new_run["envelope_keys"]:
                rows.append({**base, "field": "envelope_keys",
                             "old": old_run["envelope_keys"], "new": new_run["envelope_keys"],
                             "status": STATUS_FAILED})
            old_action = _action_of(old_run)
            new_action = _action_of(new_run)
            if old_action and new_action and old_action != new_action:
                rows.append({**base, "field": "action",
                             "old": old_action, "new": new_action, "status": STATUS_FAILED})
            if not rows or rows[-1]["entry"] != entry or rows[-1]["argv"] != argv:
                rows.append({**base, "field": "scenario",
                             "old": {"exit_code": old_run["exit_code"],
                                     "envelope_keys": old_run["envelope_keys"]},
                             "new": {"exit_code": new_run["exit_code"],
                                     "envelope_keys": new_run["envelope_keys"]},
                             "status": STATUS_MATCH})
    return rows


def _action_of(run: dict[str, Any]) -> str:
    return run.get("action", "")


# ---------------------------------------------------------------------------
# 报告编排
# ---------------------------------------------------------------------------
@dataclass
class DiffReport:
    rows: list[dict[str, Any]] = field(default_factory=list)
    counts: dict[str, int] = field(default_factory=dict)

    def add(self, row: dict[str, Any], deltas: dict[str, Any]) -> None:
        row["status"] = classify(row, deltas)
        self.rows.append(row)
        self.counts[row["status"]] = self.counts.get(row["status"], 0) + 1

    @property
    def failed(self) -> list[dict[str, Any]]:
        return [r for r in self.rows if r["status"] == STATUS_FAILED]

    def to_json(self) -> dict[str, Any]:
        return {"counts": self.counts, "rows": self.rows}

    def to_markdown(self) -> str:
        lines = [
            "# 旧新差分报告（1.5.6 vs 当前工作树）",
            "",
            f"统计：match {self.counts.get(STATUS_MATCH, 0)} / "
            f"approved {self.counts.get(STATUS_APPROVED, 0)} / "
            f"pending {self.counts.get(STATUS_PENDING, 0)} / "
            f"failed {self.counts.get(STATUS_FAILED, 0)}",
            "",
            "| 命令 | 输入 | 旧输出 | 新输出 | 差异结论 |",
            "|---|---|---|---|---|",
        ]
        for row in self.rows:
            cmd = f"{row['entry']}.{row['verb'] or '—'}"
            inp = " ".join(row["argv"]) or "—"
            old = json.dumps(row["old"], ensure_ascii=False)
            new = json.dumps(row["new"], ensure_ascii=False)
            conclusion = {"match": "等价", "approved": "白名单内有意变更",
                          "pending": "迁移待办（已知缺口）", "failed": "未登记差异（回归）"}[row["status"]]
            if row.get("field"):
                conclusion += f"｜{row['field']}"
            lines.append(f"| {cmd} | {inp} | {old} | {new} | {conclusion} |")
        return "\n".join(lines) + "\n"


def alignment_markdown(ledger: dict[str, Any], specs: tuple) -> str:
    """逐项对齐清单：旧 ledger 动词/参数/JSON 位 vs 新 manifest。

    退出码：旧 argparse 不静态编码退出码，新 manifest 的 ``exit_codes`` 为准；
    业务退出码由契约测试与动态场景兜底。
    """
    old = _old_surface(ledger)
    new = _new_surface(specs)
    old_total = sum(len(v) for v in old.values())
    new_total = sum(len(v) for v in new.values())

    lines = [
        "# 命令对齐清单（1.5.6 argparse vs 1.6 manifest）",
        "",
        f"- 旧版动词：{old_total}；新版已登记：{new_total}；待迁移：{old_total - new_total}",
        "- 口径：参数集合已剔除骨架公共位（--json/--no-input/--apply/--project/--all-project）",
        "",
    ]
    for entry in ENTRIES:
        old_verbs, new_verbs = old.get(entry, {}), new.get(entry, {})
        lines += [f"## {ENTRY_SCRIPT[entry]}（{entry}）", "",
                  "| 动词 | 旧参数 | 新参数 | 旧 --json | 新 --json | 新退出码 | 结论 |",
                  "|---|---|---|---|---|---|---|"]
        for verb in sorted(set(old_verbs) | set(new_verbs)):
            o, n = old_verbs.get(verb), new_verbs.get(verb)
            if o and n:
                aligned = (o["args"] == n["args"] and o["json_ok"] == n["json_ok"]
                           and o["project"] == n["project"]
                           and o["all_project"] == n["all_project"])
                conclusion = "对齐（无差异）" if aligned else "存在差异（见差分报告）"
                old_args, new_args = ", ".join(o["args"]) or "—", ", ".join(n["args"]) or "—"
                exit_codes = "/".join(str(c) for c in n.get("exit_codes", ()))
                lines.append(f"| `{verb}` | {old_args} | {new_args} | "
                             f"{'是' if o['json_ok'] else '否'} | "
                             f"{'是' if n['json_ok'] else '否'} | {exit_codes} | {conclusion} |")
            elif o:
                lines.append(f"| `{verb}` | {', '.join(o['args']) or '—'} | — | "
                             f"{'是' if o['json_ok'] else '否'} | — | — | 待迁移（pending） |")
            else:
                lines.append(f"| `{verb}` | — | {', '.join(n['args']) or '—'} | — | "
                             f"{'是' if n['json_ok'] else '否'} | "
                             f"{'/'.join(str(c) for c in n.get('exit_codes', ()))} | "
                             "新版新增（需核对旧版是否存在） |")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def run_harness(*, old_scripts: Path | None = None, new_scripts: Path | None = None,
                ledger: dict[str, Any] | None = None,
                deltas: dict[str, Any] | None = None,
                specs: tuple | None = None) -> DiffReport:
    """编排静态 + 动态对比；参数可注入便于测试。"""
    started = time.perf_counter()
    new_scripts = new_scripts or (PKG / "scripts")
    deltas = deltas or load_deltas(BASELINE_DIR / "approved_deltas.json")

    if ledger is None:
        ledger_path = BASELINE_DIR / "ledger.json"
        ledger = json.loads(ledger_path.read_text(encoding="utf-8")) if ledger_path.is_file() else {"entries": {}, "scenarios": []}

    if old_scripts is None:
        old_scripts, _ = recorder.resolve_baseline_dir(None, None, recorder.DEFAULT_TAG)
    if specs is None:
        from toolkit import commands as manifest
        specs = manifest.SPECS

    report = DiffReport()
    print("[1/2] 对比命令面（旧 ledger vs 新 manifest）…")
    for row in compare_surface(_old_surface(ledger), _new_surface(specs)):
        report.add(row, deltas)

    print("[2/2] 回放解析层场景（旧/新同桩配置）…")
    for row in compare_scenarios(old_scripts, new_scripts, ledger.get("scenarios", [])):
        report.add(row, deltas)

    elapsed = (time.perf_counter() - started) * 1000
    print(f"      用时 {elapsed:.1f} ms；{report.counts}")
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="旧 1.5.6 vs 新实现差分 harness")
    parser.add_argument("--report", default=str(BASELINE_DIR / "diff-report.json"))
    parser.add_argument("--markdown", default=str(BASELINE_DIR / "release-diff.md"))
    parser.add_argument("--alignment", default=str(BASELINE_DIR / "commands-alignment.md"))
    parser.add_argument("--old-dir", default="")
    args = parser.parse_args(argv)

    try:
        old_scripts = None
        if args.old_dir:
            old_scripts, _ = recorder.resolve_baseline_dir(args.old_dir, None,
                                                           recorder.DEFAULT_TAG)
    except RuntimeError as exc:
        print(f"[diff] 基线不可用：{exc}", file=sys.stderr)
        return 2

    report = run_harness(old_scripts=old_scripts)
    report_path = Path(args.report)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report.to_json(), ensure_ascii=False, indent=2,
                                      sort_keys=True), encoding="utf-8", newline="\n")
    Path(args.markdown).write_text(report.to_markdown(), encoding="utf-8", newline="\n")

    ledger = json.loads((BASELINE_DIR / "ledger.json").read_text(encoding="utf-8"))
    from toolkit import commands as manifest
    Path(args.alignment).write_text(alignment_markdown(ledger, manifest.SPECS),
                                     encoding="utf-8", newline="\n")

    if report.failed:
        print(f"[diff] 存在 {len(report.failed)} 条未登记差异，判失败。", file=sys.stderr)
        return 1
    print("[diff] 无未登记差异。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
