# -*- coding: utf-8 -*-
"""A1 门禁：全量测试 + 编译 + doctor + scrub + 安装副本校验。

只读检查，不改源码、不安装、不碰 Excel。报告落 <测试工作区>/.devloop/<时间戳>/gate.json。
用法：
  python tests/tools/devloop_gate.py
  python tests/tools/devloop_gate.py --source <仓库> --target <安装副本> --work-config <mapping>
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path


def _tail(text: str, limit: int = 4000) -> str:
    text = (text or "").strip()
    return text if len(text) <= limit else text[-limit:]


def _one_json(text: str) -> dict:
    for line in reversed((text or "").strip().splitlines()):
        line = line.strip()
        if line.startswith("{") and line.endswith("}"):
            try:
                return json.loads(line)
            except json.JSONDecodeError:
                continue
    return {}


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _excel_pids() -> list[str]:
    """只读列出当前 EXCEL.EXE 的 PID，用于观察门禁是否留下实例。"""
    try:
        proc = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=15)
    except (OSError, subprocess.TimeoutExpired):
        return []
    pids: list[str] = []
    for line in (proc.stdout or "").splitlines():
        if not line.strip().startswith('"EXCEL.EXE"'):
            continue
        parts = line.split(",")
        if parts:
            pid = parts[0].strip().strip('"')
            if pid:
                pids.append(pid)
    return pids


def _run(name: str, cmd: list[str], cwd: Path, env: dict[str, str],
         timeout: int = 900) -> dict:
    started = time.perf_counter()
    try:
        proc = subprocess.run(
            cmd, cwd=str(cwd), env=env, capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=timeout)
        timed_out = False
    except subprocess.TimeoutExpired as exc:
        return {
            "name": name, "ok": False, "returncode": None, "timeout": True,
            "elapsed_ms": int((time.perf_counter() - started) * 1000),
            "stdout_tail": _tail(getattr(exc, "stdout", "") or ""),
            "stderr_tail": _tail(getattr(exc, "stderr", "") or ""),
        }
    elapsed = int((time.perf_counter() - started) * 1000)
    return {
        "name": name, "ok": proc.returncode == 0, "returncode": proc.returncode,
        "timeout": timed_out, "elapsed_ms": elapsed,
        "stdout_tail": _tail(proc.stdout), "stderr_tail": _tail(proc.stderr),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="vba-workbench A1 门禁")
    default_source = Path(__file__).resolve().parents[2]
    parser.add_argument("--source", default=str(default_source), help="源码仓库根")
    parser.add_argument("--target", default=os.environ.get("VBA_WORKBENCH_TARGET", ""),
                        help="安装副本根；不传则跳过副本相关检查")
    parser.add_argument("--releases", default="",
                        help="发布归档目录，默认 <源码父目录>/releases")
    parser.add_argument("--work-config", default="",
                        help="doctor/list 用的项目配置，默认 assets/mapping.example.json")
    parser.add_argument("--scrub-config", default="",
                        help="带 scrub.deny 的配置，默认 assets/mapping.example.json")
    parser.add_argument("--report-root", default="",
                        help="报告目录，默认 <源码>/tests/.devloop")
    parser.add_argument("--skip-tests", action="store_true", help="跳过全量单测（仅调试用）")
    args = parser.parse_args()

    source = Path(args.source).resolve()
    target = Path(args.target).expanduser().resolve() if args.target else None
    releases = (Path(args.releases).resolve() if args.releases
                else source.parent / "releases")
    work_config = (Path(args.work_config).resolve() if args.work_config
                   else source / "assets" / "mapping.example.json")
    scrub_config = (Path(args.scrub_config).resolve() if args.scrub_config
                    else source / "assets" / "mapping.example.json")
    report_root = (Path(args.report_root).resolve() if args.report_root
                   else source / "tests" / ".devloop")
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    out_dir = report_root / stamp
    out_dir.mkdir(parents=True, exist_ok=True)

    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["VBA_WORKBENCH_CONFIG"] = str(work_config)

    py = Path(sys.executable)
    sync = source / "scripts" / "sync.py"
    results: list[dict] = []
    started = time.perf_counter()
    excel_before = _excel_pids()

    if not args.skip_tests:
        r = _run("全量单测", [str(py), "-m", "unittest", "discover",
                              "-s", ".", "-p", "test_*.py"], source / "tests", env)
        ran = re.search(r"Ran (\d+) tests", r["stdout_tail"] + r["stderr_tail"])
        r["tests_ran"] = int(ran.group(1)) if ran else None
        results.append(r)

    results.append(_run("编译", [str(py), "-m", "compileall", "-q", "scripts", "tests"],
                        source, env))

    r = _run("环境体检", [str(py), str(sync), "doctor", "--json"], source, env)
    payload = _one_json(r["stdout_tail"])
    r["failed"] = (payload.get("stats") or {}).get("failed")
    version_check = next((d for d in payload.get("details", [])
                          if d.get("name") == "VERSION 与 SKILL 一致"), {})
    r["version"] = (version_check.get("detail", "").split("VERSION=")[-1]
                    .split("；")[0].strip())
    results.append(r)

    scrub_env = env.copy()
    scrub_env["VBA_WORKBENCH_CONFIG"] = str(scrub_config)
    results.append(_run("源码脱敏", [str(py), str(sync), "scrub", "--json"], source, scrub_env))
    if target and (target / "scripts" / "sync.py").is_file():
        results.append(_run("副本脱敏", [str(py), str(sync), "scrub", "--path", str(target),
                                          "--json"], source, scrub_env))

    if target and (target / "scripts" / "sync.py").is_file():
        r = _run("安装副本校验", [str(py), str(target / "scripts" / "sync.py"),
                                   "doctor", "--verify-package", "--json"], target, env)
        payload = _one_json(r["stdout_tail"])
        pkg = (payload.get("stats") or {}).get("package") or {}
        r["package"] = {
            "present": pkg.get("present"), "checked": pkg.get("checked"),
            "missing": pkg.get("missing"), "mismatch": pkg.get("mismatch"),
            "version_match": pkg.get("version_match"),
            "manifest_version": pkg.get("manifest_version"),
            "disk_version": pkg.get("disk_version"),
        }
        results.append(r)

    release_name = ""
    latest = releases / "latest.txt"
    if latest.is_file():
        release_name = latest.read_text(encoding="utf-8").strip()
    release_zip = releases / release_name if release_name else None
    release = {
        "name": "发布包", "ok": False, "returncode": None,
        "latest": release_name, "zip": str(release_zip) if release_zip else "",
    }
    if release_zip and release_zip.is_file():
        digest = _sha256(release_zip)
        sums = releases / "SHA256SUMS.txt"
        declared = ""
        if sums.is_file():
            for line in sums.read_text(encoding="utf-8").splitlines():
                parts = line.split()
                if len(parts) == 2 and parts[1] == release_name:
                    declared = parts[0]
        release.update({"sha256": digest, "declared": declared,
                        "bytes": release_zip.stat().st_size,
                        "ok": bool(declared) and declared.lower() == digest})
        if not release["ok"]:
            release["stderr_tail"] = "SHA256SUMS 未登记或与 zip 哈希不一致"
    else:
        release["stderr_tail"] = f"latest 指向的包不存在：{release_zip}"
    results.append(release)

    version = (source / "VERSION").read_text(encoding="utf-8").strip() if (source / "VERSION").is_file() else ""
    excel_after = _excel_pids()
    environment = {
        "excel_before": len(excel_before),
        "excel_after": len(excel_after),
        "excel_delta": len(excel_after) - len(excel_before),
        "excel_new_pids": sorted(set(excel_after) - set(excel_before)),
        "excel_gone_pids": sorted(set(excel_before) - set(excel_after)),
    }
    if environment["excel_delta"] > 0:
        environment["warning"] = "门禁结束后新增 Excel 实例，未自动清理；需确认是否测试残留"
    report = {
        "ok": all(item.get("ok") for item in results),
        "version": version,
        "source": str(source),
        "target": str(target) if target else "",
        "elapsed_ms": int((time.perf_counter() - started) * 1000),
        "timestamp": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
        "environment": environment,
        "results": results,
    }
    report_path = out_dir / "gate.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    test_check = next((r for r in results if r.get("name") == "全量单测"), {})
    doctor_check = next((r for r in results if r.get("name") == "环境体检"), {})
    package_check = next((r for r in results if r.get("name") == "安装副本校验"), {})
    print(json.dumps({
        "ok": report["ok"], "version": version,
        "elapsed_ms": report["elapsed_ms"], "report": str(report_path),
        "tests_ran": test_check.get("tests_ran"),
        "doctor_failed": doctor_check.get("failed"),
        "package_ok": package_check.get("ok"),
        "environment": environment,
        "checks": [{"name": r["name"], "ok": r["ok"], "returncode": r["returncode"]}
                   for r in results],
    }, ensure_ascii=False))
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
