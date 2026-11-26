# -*- coding: utf-8 -*-
"""1.5.6 基线录制：把旧版 argparse 面与确定性场景录成脱敏清单。

录制范围（不启动 Excel、不碰业务表，全部在临时目录 + 桩配置里跑）：

1. **命令面（静态 introspect）**：从旧版三个入口的 ``build_parser()`` 直接走
   argparse 对象，录下每个动词的参数、取值范围、是否支持 ``--json`` /
   ``--apply`` / ``--project``；不调用任何 handler。
2. **场景（子进程，解析层）**：裸跑、未知动词、``--help``、只读动词的
   ``--json`` 信封键集合；只录退出码与归一化后的输出结构。

基线来源优先级：

- 环境变量 ``VBA_WB_BASELINE_DIR``：已导出的旧版副本目录（含 ``vba-workbench/scripts``）；
- 否则用 ``git archive`` 从 ``VBA_WB_BASELINE_REPO``（默认按当前 git 顶层）
  导出标签 ``baseline-1.5.6`` 到临时目录；不新建工作树。

产物见 ``tests/baseline/1.5.6/``：``ledger.json``（命令面 + 场景）、
``approved_deltas.json``（已知缺陷白名单，初始为空）。
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

PKG = Path(__file__).resolve().parents[2]
BASELINE_DIR_DEFAULT = PKG / "tests" / "baseline" / "1.5.6"
DEFAULT_TAG = "baseline-1.5.6"
ENTRIES = ("sync", "audit", "agent")
ENTRY_SCRIPT = {"sync": "sync.py", "audit": "audit.py", "agent": "agent.py"}

# 旧版统一信封键（toolkit.base.make_envelope）；动态录制以实际抓到的为准。
EXPECTED_ENVELOPE_KEYS = ("ok", "action", "details", "errors", "elapsed_ms", "ts",
                          "stats", "exit_code")

GIT_CANDIDATES = (os.environ.get("GIT_EXECUTABLE", ""),
                  "C:\\Program Files\\Git\\cmd\\git.exe", "git")

_PATH_RE = re.compile(r"[A-Za-z]:[\\/][^\s\"']+")
_TS_RE = re.compile(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:[+-]\d{2}:?\d{2}|Z)?")


def normalize_text(text: str, sandbox: Path | None = None) -> str:
    """脱敏归一化：临时目录/绝对路径/时间戳换成占位符；只留结构。"""
    out = text
    if sandbox is not None:
        out = out.replace(str(sandbox), "<sandbox>")
        out = out.replace(str(sandbox).replace("\\", "/"), "<sandbox>")
    out = _TS_RE.sub("<ts>", out)
    out = _PATH_RE.sub("<path>", out)
    return out


def _git_exe() -> str:
    for cand in GIT_CANDIDATES:
        if cand and (Path(cand).is_file() or shutil.which(cand)):
            return cand
    return "git"


def discover_repo() -> str | None:
    """按当前工作树向上找 git 顶层；找不到返回 None。"""
    try:
        r = subprocess.run([_git_exe(), "rev-parse", "--show-toplevel"],
                            cwd=str(PKG.parent), capture_output=True, text=True)
        if r.returncode == 0:
            return r.stdout.strip()
    except OSError:
        pass
    return None


def extract_baseline(repo: str, tag: str, dest: Path) -> Path:
    """git archive 导出旧版 scripts 到 dest，返回脚本目录。"""
    dest.mkdir(parents=True, exist_ok=True)
    archive = subprocess.run(
        [_git_exe(), "-C", repo, "archive", tag, "vba-workbench/scripts"],
        capture_output=True)
    if archive.returncode != 0:
        raise RuntimeError(f"git archive {tag} 失败：{archive.stderr.decode('utf-8', 'replace')}")
    subprocess.run(["tar", "-x", "-C", str(dest)], input=archive.stdout, check=True)
    scripts = dest / "vba-workbench" / "scripts"
    if not scripts.is_dir():
        raise RuntimeError(f"基线导出后未找到脚本目录：{scripts}")
    return scripts


def _load_module(module_name: str, path: Path):
    """从旧版副本目录加载入口模块；隔离 sys.path，不污染当前包。"""
    old_path = list(sys.path)
    sys.path.insert(0, str(path))
    try:
        spec = importlib.util.spec_from_file_location(module_name, path / ENTRY_SCRIPT[module_name])
        if spec is None or spec.loader is None:
            raise RuntimeError(f"无法加载旧版模块 {module_name}")
        module = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = module
        spec.loader.exec_module(module)
        return module
    finally:
        sys.path[:] = old_path


def _option_kind(action) -> str:
    if action.choices:
        return "choice"
    if action.nargs in ("*", "+") or type(action).__name__ == "_AppendAction":
        return "append"
    if action.type is int:
        return "int"
    if action.type is float:
        return "float"
    if getattr(action, "const", None) is not None and action.nargs == 0:
        return "flag"
    return "str"


def introspect_parser(parser, entry: str | None = None) -> dict[str, Any]:
    """走 argparse 对象，抽出动词面与参数（纯静态，不执行 handler）。

    ``entry`` 用于入口级特殊口径（见下 ``agent`` 分支）。
    """
    verbs: dict[str, dict[str, Any]] = {}
    sub = None
    for action in getattr(parser, "_actions", []):
        if getattr(action, "choices", None) and isinstance(action.choices, dict):
            sub = action
            break
    global_options: list[str] = []
    for action in getattr(parser, "_actions", []):
        if not getattr(action, "option_strings", None):
            continue
        if sub is not None and action in getattr(sub, "_choices_actions", []):
            continue
        global_options.extend(action.option_strings)

    if sub is None:
        return {"global_options": sorted(global_options), "verbs": verbs}

    for verb, subparser in sub.choices.items():
        args: list[dict[str, Any]] = []
        flags: dict[str, bool] = {"json": False, "apply": False,
                                   "project": False, "all_project": False}
        for action in getattr(subparser, "_actions", []):
            opts = getattr(action, "option_strings", [])
            if not opts:
                continue
            if opts in (["-h"], ["-h", "--help"]) or "--help" in opts:
                continue  # argparse 内置帮助，不是业务参数
            names = [o.lstrip("-") for o in opts]
            if "--json" in opts:
                flags["json"] = True
            if "--apply" in opts:
                flags["apply"] = True
            if "--project" in opts:
                flags["project"] = True
            if "--all-project" in opts:
                flags["all_project"] = True
            default = action.default
            if default is argparse.SUPPRESS or default is None:
                default = None
            args.append({
                "names": names,
                "kind": _option_kind(action),
                "required": bool(getattr(action, "required", False)),
                "choices": sorted(action.choices) if action.choices else [],
                "default": default,
            })
        verbs[str(verb)] = {
            "args": args,
            "json_ok": flags["json"],
            "apply": flags["apply"],
            "project": flags["project"],
            "all_project": flags["all_project"],
        }

    # agent 入口特殊口径（1.5.6 实测，源 main() 解析前预剥离 --json）：
    # main() 先 `as_json = "--json" in argv` 并 `argv.remove("--json")`，
    # 再 build_parser().parse_args(argv)。--json 是全局开关、不经过子命令
    # 选项注册：除 probe 外的 8 个动词虽未在 parser 里挂 --json，CLI 实际
    # 都接受，并落到通用 JSON 信封分支。故按真实 CLI 接受面录 json_ok=True，
    # 不能按 argparse 选项是否存在判定。
    # sync / audit 没有预剥离：--json 未登记即 argparse 拒绝（退出 2），
    # 保持选项存在性判定，不套用本分支。
    if entry == "agent":
        for info in verbs.values():
            info["json_ok"] = True
    return {"global_options": sorted(global_options), "verbs": dict(sorted(verbs.items()))}


def _stub_config(sandbox: Path) -> Path:
    cfg = sandbox / "mapping.json"
    cfg.write_text(json.dumps({"defaults": {}, "projects": []}, ensure_ascii=False),
                   encoding="utf-8")
    return cfg


def run_parse_scenario(script: Path, argv: list[str], sandbox: Path,
                       python: str | None = None) -> dict[str, Any]:
    """跑一个解析层场景：不起 Excel；只录退出码与 JSON 信封键（若有）。"""
    cfg = _stub_config(sandbox)
    env = dict(os.environ, VBA_WORKBENCH_CONFIG=str(cfg),
               PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    started = time.perf_counter()
    proc = subprocess.run([python or sys.executable, str(script), *argv],
                          cwd=str(sandbox), env=env, capture_output=True,
                          text=True, encoding="utf-8", errors="replace")
    elapsed_ms = round((time.perf_counter() - started) * 1000)
    envelope_keys: list[str] = []
    action = ""
    last_line = next((ln for ln in reversed(proc.stdout.splitlines()) if ln.strip().startswith("{")), "")
    if last_line:
        try:
            payload = json.loads(last_line)
            if isinstance(payload, dict):
                envelope_keys = sorted(payload.keys())
                action = str(payload.get("action", ""))
        except json.JSONDecodeError:
            envelope_keys = []
    return {
        "argv": argv,
        "exit_code": proc.returncode,
        "envelope_keys": envelope_keys,
        "action": action,
        "stdout_lines": len(proc.stdout.splitlines()),
        "stderr": normalize_text(proc.stderr, sandbox)[:500],
        "elapsed_ms": elapsed_ms,
    }


@dataclass
class Recorder:
    """录制编排：静态命令面 + 解析层场景，输出 ledger。"""

    scripts_dir: Path
    out_dir: Path
    repo: str | None = None
    tag: str = DEFAULT_TAG
    scenarios: list[dict[str, Any]] = field(default_factory=list)

    def record_static(self) -> dict[str, Any]:
        ledger_entries: dict[str, Any] = {}
        for entry in ENTRIES:
            module = _load_module(entry, self.scripts_dir)
            parser = module.build_parser()
            ledger_entries[entry] = introspect_parser(parser, entry=entry)
        return ledger_entries

    def record_scenarios(self) -> None:
        """解析层场景：裸跑 / 未知动词 / 帮助 / 只读 JSON 信封。"""
        with tempfile.TemporaryDirectory() as raw:
            sandbox = Path(raw)
            for entry in ENTRIES:
                script = self.scripts_dir / ENTRY_SCRIPT[entry]
                self.scenarios.append({
                    "entry": entry, "case": "bare",
                    **run_parse_scenario(script, [], sandbox),
                })
                self.scenarios.append({
                    "entry": entry, "case": "unknown-verb",
                    **run_parse_scenario(script, ["__baseline_no_such_verb__"], sandbox),
                })
            # 只读动词的 JSON 信封：version 不碰 Excel；其余动词合并完成后再补。
            version_script = self.scripts_dir / ENTRY_SCRIPT["sync"]
            self.scenarios.append({
                "entry": "sync", "case": "version-json",
                **run_parse_scenario(version_script, ["version", "--json"], sandbox),
            })

    def build_ledger(self) -> dict[str, Any]:
        static = self.record_static()
        self.record_scenarios()
        return {
            "meta": {
                "baseline_tag": self.tag,
                "baseline_commit": _commit_of(self.repo, self.tag) if self.repo else "",
                "scope": "静态 argparse 面 + 解析层场景；不启动 Excel，不碰业务表",
                "expected_envelope_keys": list(EXPECTED_ENVELOPE_KEYS),
                "normalization": "临时目录→<sandbox>，绝对路径→<path>，时间戳→<ts>",
            },
            "entries": static,
            "scenarios": self.scenarios,
        }

    def write(self) -> dict[str, Any]:
        self.out_dir.mkdir(parents=True, exist_ok=True)
        ledger = self.build_ledger()
        (self.out_dir / "ledger.json").write_text(
            json.dumps(ledger, ensure_ascii=False, indent=2, sort_keys=True),
            encoding="utf-8", newline="\n")
        deltas = self.out_dir / "approved_deltas.json"
        if not deltas.exists():
            deltas.write_text(json.dumps({"approved_deltas": [], "pending": []},
                                          ensure_ascii=False, indent=2),
                              encoding="utf-8", newline="\n")
        return ledger


def _commit_of(repo: str, tag: str) -> str:
    try:
        r = subprocess.run([_git_exe(), "-C", repo, "rev-parse", tag],
                           capture_output=True, text=True)
        return r.stdout.strip() if r.returncode == 0 else ""
    except OSError:
        return ""


def resolve_baseline_dir(cli_dir: str | None, repo: str | None, tag: str) -> tuple[Path, Path]:
    """返回 (脚本目录, 临时导出目录或空)。"""
    env_dir = os.environ.get("VBA_WB_BASELINE_DIR", "").strip()
    if cli_dir:
        env_dir = cli_dir
    if env_dir:
        root = Path(env_dir)
        scripts = root / "vba-workbench" / "scripts" if (root / "vba-workbench").is_dir() else root
        return scripts, root
    repo = repo or os.environ.get("VBA_WB_BASELINE_REPO", "").strip() or discover_repo()
    if not repo:
        raise RuntimeError("找不到基线仓库：设置 VBA_WB_BASELINE_REPO 或 VBA_WB_BASELINE_DIR")
    tmp = Path(tempfile.mkdtemp(prefix="vbaw-baseline-"))
    return extract_baseline(repo, tag, tmp), tmp


def reconcile_pending(ledger: dict[str, Any], specs: tuple,
                       deltas_path: Path) -> list[dict[str, Any]]:
    """按 ledger 与当前 manifest 差集刷新 pending 待办；保留已登记白名单。"""
    existing = load_deltas_file(deltas_path)
    new_verbs = {(s.entry, s.verb) for s in specs}
    pending: list[dict[str, Any]] = []
    for entry, data in ledger.get("entries", {}).items():
        for verb in data.get("verbs", {}):
            if (entry, verb) not in new_verbs:
                pending.append({
                    "entry": entry, "verb": verb,
                    "field": "command-missing-in-new",
                    "reason": "1.6 迁移待办：对应 feature 切片尚未登记 manifest",
                })
    payload = {"approved_deltas": existing.get("approved_deltas", []),
               "pending": pending}
    deltas_path.parent.mkdir(parents=True, exist_ok=True)
    deltas_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2,
                                      sort_keys=True), encoding="utf-8", newline="\n")
    return pending


def load_deltas_file(path: Path) -> dict[str, Any]:
    if not path.is_file():
        return {"approved_deltas": [], "pending": []}
    data = json.loads(path.read_text(encoding="utf-8"))
    data.setdefault("approved_deltas", [])
    data.setdefault("pending", [])
    return data


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="录制 1.5.6 基线命令面与解析层场景")
    parser.add_argument("--out", default=str(BASELINE_DIR_DEFAULT))
    parser.add_argument("--repo", default="")
    parser.add_argument("--tag", default=DEFAULT_TAG)
    parser.add_argument("--baseline-dir", default="")
    parser.add_argument("--sync-pending", action="store_true",
                        help="只按 ledger 与当前 manifest 差集刷新 pending 待办，不重录")
    args = parser.parse_args(argv)

    if args.sync_pending:
        ledger_path = BASELINE_DIR_DEFAULT / "ledger.json"
        ledger = json.loads(ledger_path.read_text(encoding="utf-8"))
        sys.path.insert(0, str(PKG / "scripts"))
        from toolkit import commands as manifest
        pending = reconcile_pending(ledger, manifest.SPECS,
                                    BASELINE_DIR_DEFAULT / "approved_deltas.json")
        print(f"[pending] 待迁移动词 {len(pending)} 个 → approved_deltas.json")
        return 0

    started = time.perf_counter()
    print("[1/3] 解析基线来源 …")
    try:
        scripts, tmp = resolve_baseline_dir(args.baseline_dir or None,
                                            args.repo or None, args.tag)
    except RuntimeError as exc:
        print(f"[record] 失败：{exc}", file=sys.stderr)
        return 2
    print(f"      脚本目录：{scripts}")

    print("[2/3] 录制静态命令面 + 解析层场景 …")
    recorder = Recorder(scripts_dir=scripts, out_dir=Path(args.out),
                       repo=args.repo or discover_repo(), tag=args.tag)
    ledger = recorder.write()

    verb_count = sum(len(e["verbs"]) for e in ledger["entries"].values())
    elapsed = (time.perf_counter() - started) * 1000
    print(f"[3/3] 完成：动词 {verb_count} 个，场景 {len(ledger['scenarios'])} 条；"
          f"清单 → {args.out}；用时 {elapsed:.1f} ms")
    if str(tmp).startswith(tempfile.gettempdir()):
        shutil.rmtree(tmp, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
