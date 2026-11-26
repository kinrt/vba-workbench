# -*- coding: utf-8 -*-
"""core 切片：version / list / doctor / scrub（只读，不碰 Excel、不写盘）。"""
from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any

from toolkit.base import (ROOT, SOURCE_EXTS, abs_path, entry_command,
                          read_source_text)
from toolkit.config import config_source
from toolkit.contract import Failure as ContractFailure, MutationMode, Outcome, Stage
from toolkit.repo import (VERSION_RE, files_under, is_shell_file, module_index,
                          read_common_version, resolve_target)
from toolkit.registry import ArgSpec, CommandSpec, ProjectMode
from toolkit.vba_parse import declared_procs, mask_strings, strip_vba_comments


# ---------------------------------------------------------------------------
# 公共模块子集判定：统一在 features/common_subset.py，切片只做参数适配
# ---------------------------------------------------------------------------
from features.common_subset import (  # noqa: E402
    resolve_common_subset as _resolve_common_subset,
    warn_unknown_modules as _warn_unknown_modules,
)


# 旧实现已迁到 features/common_subset.py；切片只做参数适配，不复制判定规则。


# ---------------------------------------------------------------------------
# version：核对公共库版本与各项目锁定版本
# ---------------------------------------------------------------------------
def version(ctx) -> Outcome:
    with ctx.progress.stage("核对公共库版本"):
        defaults = ctx.config["defaults"]
        common_dir = abs_path(defaults.get("common_dir", "src/common"))
        actual = read_common_version(common_dir)
        warnings: list[str] = []
        human: list[str] = []
        if actual is None:
            warnings.append(f"在 {common_dir} 里没找到 COMMON_VERSION 常量")
        else:
            human.append(f"公共库源码版本：{actual}")

        bad = 0
        configured = actual is not None
        rows: list[dict] = []
        for proj in ctx.config["projects"]:
            locked = proj.get("common_version")
            if not locked:
                human.append(f"  {proj['name']}: 未锁定版本")
                reason = ("未配置 src/common 或缺少 COMMON_VERSION" if not configured
                          else "项目未锁定 common_version")
                rows.append({"name": proj["name"], "locked": None, "match": None, "reason": reason})
                continue
            state = "一致" if locked == actual else "不一致"
            if locked != actual:
                bad += 1
            human.append(f"  {proj['name']}: 锁定 {locked} -> {state}")
            rows.append({"name": proj["name"], "locked": locked, "match": locked == actual})

        errors = []
        if bad:
            errors.append(ContractFailure(
                f"{bad} 个项目的锁定版本与公共库 {actual or '缺失'} 不一致",
                code="VERSION_MISMATCH", stage=Stage.CONFIG,
                hint="统一 src/common/COMMON_VERSION 与项目 mapping 的 common_version",
            ))
        # stats 键必须与 1.5.6 逐字一致：results 是行数据副本，errors 是人话字符串列表
        stats: dict = {"common_version": actual, "configured": configured,
                       "common_dir": str(common_dir), "mismatched": bad,
                       "results": rows, "timings": ctx.stage_timings()}
        if bad:
            stats["errors"] = [str(error) for error in errors]
        return Outcome(
            action="sync.version",
            ok=not bad,
            details=rows,
            errors=errors,
            stats=stats,
            human=human,
            warnings=warnings,
            exit_code=0 if not bad else 1,
            elapsed_ms=ctx.elapsed_ms(),
        )


# ---------------------------------------------------------------------------
# list：每个项目解析到哪张 xlsm
# ---------------------------------------------------------------------------
def _list_human(rows: list[dict], total: int, resolved: int,
                config: str, root: str) -> list[str]:
    lines = ["项目配置解析结果：", f"  配置：{config}", f"  基准：{root}", ""]
    for r in rows:
        if r["ok"]:
            lines.append(f"  [OK] {r['name']}")
            lines.append(f"       表格：{r['workbook']}")
            lines.append(f"       原因：{r['reason']}")
        else:
            lines.append(f"  [--] {r['name']}")
            lines.append(f"       表格：（未解析到）{r['reason']}")
        lines.append(f"       源码：{r['src']}")
        if r["common_modules"] == "all":
            lines.append("       公共模块：全部写回")
        else:
            names = "、".join(r["common_modules"]) or "（无）"
            lines.append(f"       公共模块：{r['common_reason']} → {names}")
        if r["note"]:
            lines.append(f"       备注：{r['note']}")
        lines.append("")
    lines.append(f"共 {total} 个项目，可用 {resolved} 个")
    return lines


def list_projects(ctx) -> Outcome:
    with ctx.progress.stage("解析项目清单"):
        rows: list[dict] = []
        errors: list[ContractFailure] = []
        warnings: list[str] = []
        for proj in ctx.config["projects"]:
            path, reason = resolve_target(proj)
            subset, subset_reason, ignored = _resolve_common_subset(proj, ctx.config["defaults"])
            rows.append({
                "name": proj["name"],
                "ok": bool(path),
                "workbook": str(path) if path else "",
                "reason": reason,
                "src": proj.get("src", ""),
                "common_modules": sorted(subset) if subset is not None else "all",
                "common_reason": subset_reason,
                "note": proj.get("note", ""),
            })
            if ignored:
                _warn_unknown_modules(ignored)
            if not path:
                prefix = proj.get("prefix") or proj["name"]
                expected = proj.get("file") or f"{prefix} V<数字>.xlsm"
                errors.append(ContractFailure(
                    f"「{proj['name']}」解析不到 xlsm：{reason}",
                    code="TARGET_NOT_FOUND", stage=Stage.RESOLVE,
                    hint=(f"search_dir={proj.get('search_dir', '')}；prefix={proj.get('prefix', '')}；"
                          f"期望文件 {expected}；改 mapping 后重跑 list --json"),
                    suggested=entry_command("sync.py", "list", "--json"),
                    retryable=False,
                ))
        total = len(rows)
        resolved = sum(1 for r in rows if r["ok"])
        ok = total > 0 and resolved == total

    return Outcome(
        action="sync.list",
        ok=ok,
        details=rows,
        errors=errors,
        stats={"total": total, "resolved": resolved,
               "config": config_source(), "root": str(ROOT),
               "timings": ctx.stage_timings()},
        human=_list_human(rows, total, resolved, config_source(), str(ROOT)),
        warnings=warnings,
        exit_code=0 if ok else 1,
        elapsed_ms=ctx.elapsed_ms(),
    )


# ---------------------------------------------------------------------------
# doctor：环境体检（每项只有 OK / WARN / FAIL 三档）
# ---------------------------------------------------------------------------
# 包完整性校验（--verify-package）：交付根/根级文件口径与安装器白名单对齐。
PACKAGE_MANIFEST_NAME = ".installed.json"
PACKAGE_DELIVERED_ROOTS = ("scripts", "references", "assets")
PACKAGE_DELIVERED_FILES = ("SKILL.md", "VERSION", "LICENSE", "requirements.txt")
PACKAGE_EXCLUDE_DIRS = {"__pycache__", ".git", ".lab", ".run", ".backup",
                        "releases", "evals", "tests"}
PACKAGE_EXCLUDE_SUFFIX = (".pyc", ".pyo", ".tmp", ".log")

# scrub.deny 扫描：交付件白名单内的文本后缀；.bat 由 read_source_text 回退 GBK。
SCRUB_TEXT_SUFFIXES = ("", ".md", ".py", ".json", ".txt", ".bat", ".cmd",
                       ".ps1", ".yml", ".yaml", ".toml", ".ini", ".cfg")


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _installed_file_relset(pkg: Path) -> set[str]:
    """包内现存交付件相对路径集合（POSIX 风格）；口径对齐安装器白名单与排除项。"""
    rels: set[str] = set()
    for root in PACKAGE_DELIVERED_ROOTS:
        base = pkg / root
        if not base.is_dir():
            continue
        for path in base.rglob("*"):
            if not path.is_file():
                continue
            rel = path.relative_to(pkg)
            if set(rel.parts) & PACKAGE_EXCLUDE_DIRS:
                continue
            if path.suffix.lower() in PACKAGE_EXCLUDE_SUFFIX:
                continue
            rels.add(str(rel).replace("\\", "/"))
    for name in PACKAGE_DELIVERED_FILES:
        if (pkg / name).is_file():
            rels.add(name)
    return rels


def _verify_package(pkg: Path) -> tuple[dict[str, Any], list[ContractFailure]]:
    """显式包完整性校验：清单在位、字段齐全、files 对账、逐项哈希、VERSION 一致。

    返回 ``(stats.package, errors)``；缺清单给专用码 ``PACKAGE_MANIFEST_MISSING``，
    不抛异常崩出口。默认 doctor 不调用本函数，本地正常改动不会因此失败。
    """
    stats: dict[str, Any] = {
        "present": False, "checked": 0, "missing": [], "mismatch": [],
        "extra": [], "missing_fields": [], "version_match": None,
    }
    errors: list[ContractFailure] = []
    manifest_path = pkg / PACKAGE_MANIFEST_NAME
    if not manifest_path.is_file():
        errors.append(ContractFailure(
            f"未找到安装清单 {PACKAGE_MANIFEST_NAME}（源码树或未安装副本属正常）",
            code="PACKAGE_MANIFEST_MISSING", stage=Stage.CONFIG,
            hint="源码树/未安装副本本就没有清单；若这是已安装副本，请用安装器 --apply 刷新清单后再校验",
            retryable=False))
        return stats, errors

    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as exc:
        errors.append(ContractFailure(
            f"清单 {PACKAGE_MANIFEST_NAME} 解析失败：{exc}",
            code="PACKAGE_INTEGRITY", stage=Stage.CONFIG,
            hint="清单损坏时用安装器重装/--apply 刷新；保留 stderr 线索交维护者",
            retryable=False))
        return stats, errors

    stats["present"] = True
    required_fields = ("version", "source", "installed_at", "files")
    stats["missing_fields"] = [f for f in required_fields if f not in manifest]
    recorded = manifest.get("files")
    if not isinstance(recorded, dict):
        recorded = {}
    disk = _installed_file_relset(pkg)
    recorded_rels = set(recorded)
    missing = sorted(recorded_rels - disk)
    extra = sorted(disk - recorded_rels)
    mismatch: list[str] = []
    for rel in sorted(recorded_rels & disk):
        try:
            digest = _sha256_file(pkg / rel)
        except OSError:
            mismatch.append(rel)
            continue
        if recorded[rel] != digest:
            mismatch.append(rel)
    ver_file = pkg / "VERSION"
    disk_version = ver_file.read_text(encoding="utf-8").strip() if ver_file.is_file() else ""
    manifest_version = str(manifest.get("version", ""))
    version_match = bool(disk_version) and manifest_version == disk_version
    stats.update(checked=len(recorded_rels), missing=missing, mismatch=mismatch,
                 extra=extra, version_match=version_match,
                 manifest_version=manifest_version, disk_version=disk_version)

    def fail(message: str) -> None:
        errors.append(ContractFailure(
            message, code="PACKAGE_INTEGRITY", stage=Stage.CONFIG,
            hint="本机正常改动可忽略；要恢复出厂快照可用安装器重装或 --apply 刷新清单；"
                 "缺文件/哈希不符的相对路径见 message",
            retryable=False))

    if stats["missing_fields"]:
        fail(f"清单缺字段：{'、'.join(stats['missing_fields'])}")
    if missing:
        fail(f"清单登记但磁盘缺失 {len(missing)} 个：{'、'.join(missing)}")
    if mismatch:
        fail(f"哈希不符 {len(mismatch)} 个：{'、'.join(mismatch)}")
    if extra:
        fail(f"磁盘存在但清单未登记 {len(extra)} 个：{'、'.join(extra)}")
    if not version_match:
        fail(f"VERSION 与清单版本不一致：磁盘 {disk_version or '缺失'} / 清单 {manifest_version or '缺失'}")
    return stats, errors


def _collect_doctor(cfg: dict) -> tuple[list[dict], list[ContractFailure]]:
    """环境体检的可判定清单：每项只有 OK / WARN / FAIL 三档，agent 也能直接读。"""
    checks: list[dict] = []
    errors: list[ContractFailure] = []

    def add(name: str, ok: bool, detail: str, level: str = "fail",
            code: str = "", stage: str = "", hint: str = "") -> None:
        # level="warn" 是明确要报的警告；否则按 ok 判 ok/fail
        lvl = "warn" if level == "warn" else ("ok" if ok else "fail")
        checks.append({"name": name, "level": lvl, "detail": detail})
        if lvl == "fail":
            errors.append(ContractFailure(
                f"{name}：{detail}", code=code, stage=stage, hint=hint, retryable=False))

    add("平台", sys.platform.startswith("win"),
        f"sys.platform = {sys.platform}（需 Windows + 微软 Excel）",
        code="NO_EXCEL", stage=Stage.CONFIG, hint="需要 Windows + 微软 Excel")
    try:
        import win32com.client  # noqa: F401
        add("pywin32", True, "已安装", code="NO_EXCEL", stage=Stage.CONFIG,
            hint="python -m pip install pywin32")
    except Exception:
        add("pywin32", False, "未安装：python -m pip install pywin32",
            code="NO_EXCEL", stage=Stage.CONFIG, hint="python -m pip install pywin32")

    add("python", sys.version_info >= (3, 10),
        f"{sys.version.split()[0]}（{sys.executable}）；要求 3.10+",
        code="CONFIG_INVALID", stage=Stage.CONFIG, hint="需要 Python 3.10+")

    # 包自检：交付出去的目录也可能是从 zip 解压来的，VERSION / SKILL / 必需文件都得齐
    pkg = Path(__file__).resolve().parents[2]
    ver_file = pkg / "VERSION"
    skill = pkg / "SKILL.md"
    ver = ver_file.read_text(encoding="utf-8").strip() if ver_file.is_file() else ""
    m = re.search(r'version:\s*"?([0-9.]+)"?', skill.read_text(encoding="utf-8")) if skill.is_file() else None
    add("VERSION 与 SKILL 一致",
        bool(ver) and bool(m) and m.group(1) == ver,
        f"VERSION={ver or '缺失'}；SKILL={m.group(1) if m else '缺失'}",
        code="CONFIG_INVALID", stage=Stage.CONFIG,
        hint="核对包根 VERSION 与 SKILL.md 的 metadata.version")
    need = ["SKILL.md", "VERSION", "scripts/sync.py", "scripts/audit.py", "scripts/toolkit/base.py",
            "references/sync.md", "references/sync-commands.md", "LICENSE",
            "assets/mapping.example.json"]
    missing = [n for n in need if not (pkg / n).is_file()]
    add("包文件齐全", not missing,
        "必需文件 9 个都在" if not missing else f"缺：{'、'.join(missing)}",
        code="CONFIG_INVALID", stage=Stage.CONFIG, hint="重装 vba-workbench 补齐必需文件")

    import winreg

    found_office = False
    for ver in ("16.0", "15.0", "14.0"):
        key_path = rf"Software\Microsoft\Office\{ver}\Excel\Security"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as k:
                val, _ = winreg.QueryValueEx(k, "AccessVBOM")
            found_office = True
            add(f"Office {ver} AccessVBOM", val == 1,
                "已信任 VBA 工程对象模型" if val == 1 else f"= {val}，需在信任中心勾选",
                code="VBE_PROTECTED", stage=Stage.ATTACH,
                hint="在信任中心勾选「信任对 VBA 工程对象模型的访问」")
        except OSError:
            continue
    if not found_office:
        add("Office AccessVBOM", False, "未读到 Excel\\Security\\AccessVBOM：Excel 未装或版本键不同",
            code="NO_EXCEL", stage=Stage.ATTACH, hint="Excel 未装或版本键不同")

    add("配置文件", True, config_source())

    root = ROOT
    try:
        probe = root / ".write_probe.tmp"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        add("仓库目录可写", True, str(root),
            code="FILE_READONLY", stage=Stage.CONFIG, hint="检查目录只读属性与同步盘占用")
    except Exception as e:
        add("仓库目录可写", False, f"{root}（{e}）",
            code="FILE_READONLY", stage=Stage.CONFIG, hint="检查目录只读属性与同步盘占用")

    resolved = 0
    for proj in cfg["projects"]:
        path, reason = resolve_target(proj)
        if path:
            resolved += 1
            dup = "注意：同尺寸同 mtime" in reason
            add(f"项目解析·{proj['name']}", True,
                str(path) + (f"；{reason}" if dup else ""),
                level="warn" if dup else "ok",
                code="TARGET_NOT_FOUND", stage=Stage.RESOLVE)
        else:
            add(f"项目解析·{proj['name']}", False, reason,
                code="TARGET_NOT_FOUND", stage=Stage.RESOLVE)
    total = len(cfg["projects"])
    add("项目可用数", resolved == total, f"{resolved}/{total}",
        level="fail" if resolved == 0 else ("ok" if resolved == total else "warn"),
        code="TARGET_NOT_FOUND", stage=Stage.RESOLVE)

    for s in [abs_path(proj["src"]) for proj in cfg["projects"] if proj.get("src")]:
        add(f"源码目录·{s.name}", s.is_dir(), str(s),
            code="SOURCE_PREFLIGHT", stage=Stage.PARSE, hint="检查项目 src 配置与目录")

    common = abs_path(cfg["defaults"].get("common_dir", "src/common"))
    if not common.is_dir():
        add("公共库目录", False, f"{common}（不存在）", level="warn")
    else:
        # 递归统计，口径与 audit common 一致（files_under：rglob、跳过占位壳与噪声文件）；
        # promote 把模块落到 src/common/<模块>/ 子目录，只数顶层会误报"没有模块"。
        mods = files_under(common)
        add("公共库目录", bool(mods),
            f"{common}（{len(mods)} 个模块）" if mods else f"{common}（目录存在但没有模块）",
            level="ok" if mods else "warn")
    return checks, errors


def doctor(ctx) -> Outcome:
    with ctx.progress.stage("环境体检"):
        checks, errors = _collect_doctor(ctx.config)

    # 显式包完整性校验：默认关闭；只有带 --verify-package 才对账 .installed.json。
    package_stats: dict[str, Any] = {}
    if getattr(ctx.args, "verify_package", False):
        pkg = Path(__file__).resolve().parents[2]
        package_stats, pkg_errors = _verify_package(pkg)
        for e in pkg_errors:
            errors.append(e)
            checks.append({"name": "包完整性校验", "level": "fail", "detail": str(e)})

    bad = [c for c in checks if c["level"] == "fail"]
    marks = {"ok": "OK", "warn": "--", "fail": "!!"}
    human = ["环境体检", ""]
    for c in checks:
        human.append(f"  [{marks[c['level']]}] {c['name']}：{c['detail']}")
    human.append("")
    if bad:
        human.append(f"体检未通过：{len(bad)} 项不合格，先按上面 [!!] 处理")
    else:
        human.append("体检通过")

    stats: dict[str, Any] = {"failed": len(bad), "timings": ctx.stage_timings()}
    if package_stats:
        stats["package"] = package_stats

    return Outcome(
        action="sync.doctor",
        ok=not bad,
        details=checks,
        errors=errors,
        stats=stats,
        human=human,
        exit_code=0 if not bad else 1,
        elapsed_ms=ctx.elapsed_ms(),
    )


def _scrub_deny_terms(cfg: dict) -> list[str]:
    """读 mapping.json 的 scrub.deny；缺失、空表或类型不对都按配置错误处理。"""
    scrub = cfg.get("scrub") if isinstance(cfg, dict) else None
    deny = scrub.get("deny") if isinstance(scrub, dict) else None
    if not isinstance(deny, list):
        raise ValueError("mapping.json 缺少 scrub.deny 数组")
    seen: set[str] = set()
    terms: list[str] = []
    for raw in deny:
        word = str(raw).strip()
        key = word.casefold()
        if word and key not in seen:
            seen.add(key)
            terms.append(word)
    terms.sort(key=str.casefold)
    if not terms:
        raise ValueError("mapping.json 的 scrub.deny 是空表")
    return terms


def _is_scrub_text(path: Path) -> bool:
    return path.suffix.lower() in SCRUB_TEXT_SUFFIXES


def _scan_package(root: Path, deny_terms: list[str]) -> tuple[set[str], list[dict], list[str]]:
    """按交付件白名单逐词扫；返回（文件集合，命中项，跳过的非文本文件）。"""
    rels = _installed_file_relset(root)
    hits: list[dict] = []
    skipped: list[str] = []
    for rel in sorted(rels):
        path = root / rel
        if not _is_scrub_text(path):
            skipped.append(rel)
            continue
        try:
            text = read_source_text(path)
        except OSError:
            skipped.append(rel)
            continue
        for lineno, line in enumerate(text.splitlines(), 1):
            folded = line.casefold()
            for term in deny_terms:
                if term.casefold() not in folded:
                    continue
                excerpt = line.strip()
                if len(excerpt) > 160:
                    excerpt = excerpt[:160] + "…"
                hits.append({"file": rel, "line": lineno, "word": term, "text": excerpt})
    return rels, hits, skipped


def scrub(ctx) -> Outcome:
    with ctx.progress.stage("脱敏扫描"):
        try:
            terms = _scrub_deny_terms(ctx.config)
        except ValueError as exc:
            return Outcome(
                action="sync.scrub", ok=False,
                errors=[ContractFailure(
                    str(exc), code="CONFIG_INVALID", stage=Stage.CONFIG,
                    hint="在仓库根 mapping.json 的 scrub.deny 里登记真实路径/表名/用户名等词；"
                         "本机覆盖层 mapping.local.json 可增量补词",
                    retryable=False)],
                stats={"deny_terms": 0, "files": 0, "hits": 0,
                       "timings": ctx.stage_timings()},
                human=["脱敏扫描", f"  配置错误：{exc}"],
                exit_code=1, elapsed_ms=ctx.elapsed_ms(),
            )

        root = Path(getattr(ctx.args, "path", None) or Path(__file__).resolve().parents[2]).expanduser()
        if not root.is_dir():
            return Outcome(
                action="sync.scrub", ok=False,
                errors=[ContractFailure(
                    f"扫描根目录不存在：{root}", code="CONFIG_NOT_FOUND", stage=Stage.CONFIG,
                    hint="用 --path 指向已安装包或源码包根目录后重跑",
                    retryable=False)],
                stats={"root": str(root), "deny_terms": len(terms), "files": 0, "hits": 0,
                       "timings": ctx.stage_timings()},
                human=["脱敏扫描", f"  根目录不存在：{root}"],
                exit_code=1, elapsed_ms=ctx.elapsed_ms(),
            )

        rels, hits, skipped = _scan_package(root, terms)
        if not rels:
            return Outcome(
                action="sync.scrub", ok=False,
                errors=[ContractFailure(
                    f"根目录下没有可扫描的交付件：{root}", code="CONFIG_INVALID",
                    stage=Stage.CONFIG,
                    hint="确认 --path 指向 vba-workbench 包根，而不是 tests/ 或空目录",
                    retryable=False)],
                stats={"root": str(root), "deny_terms": len(terms), "files": 0, "hits": 0,
                       "timings": ctx.stage_timings()},
                human=["脱敏扫描", f"  没有可扫描的交付件：{root}"],
                exit_code=1, elapsed_ms=ctx.elapsed_ms(),
            )

    lines = ["脱敏扫描（交付件范围）", f"  根：{root}",
             f"  词表 {len(terms)} 个；文件 {len(rels)} 个"]
    if skipped:
        lines.append(f"  跳过非文本 {len(skipped)} 个")
    lines.append("")
    errors: list[ContractFailure] = []
    if hits:
        lines.append(f"命中 {len(hits)} 处：")
        for hit in hits[:40]:
            lines.append(f"  {hit['file']}:{hit['line']}  {hit['word']}  {hit['text']}")
        if len(hits) > 40:
            lines.append(f"  …另有 {len(hits) - 40} 处，见 details")
        first = f"{hits[0]['file']}:{hits[0]['line']}（{hits[0]['word']}）"
        errors.append(ContractFailure(
            f"脱敏扫描命中 {len(hits)} 处，首处 {first}", code="SCRUB_HIT",
            stage=Stage.PREFLIGHT,
            hint="把这些真实路径/表名/用户名改成占位符，或补进 scrub.deny 让后续继续拦截；"
                 "重跑 sync.py scrub --json",
            retryable=False))
    else:
        lines.append("脱敏通过：无命中")

    return Outcome(
        action="sync.scrub",
        ok=not hits,
        details=hits,
        errors=errors,
        stats={"root": str(root), "scope": "package", "files": len(rels),
               "deny_terms": len(terms), "hits": len(hits),
               "matched_terms": sorted({h["word"] for h in hits}, key=str.casefold),
               "skipped_non_text": len(skipped), "timings": ctx.stage_timings()},
        human=lines,
        exit_code=0 if not hits else 1,
        elapsed_ms=ctx.elapsed_ms(),
    )


SPECS = (
    CommandSpec(
        verb="version", entry="sync", handler="features.core:version",
        mutation=MutationMode.NONE, project=ProjectMode.NONE,
        supports_json=True, exit_codes=(0, 1, 2),
        summary="核对公共库版本与各项目锁定版本",
        ref="references/sync-commands.md#version",
        tests=("contract.version.success", "contract.version.mismatch"),
    ),
    CommandSpec(
        verb="list", entry="sync", handler="features.core:list_projects",
        mutation=MutationMode.NONE, project=ProjectMode.NONE,
        supports_json=True, exit_codes=(0, 1, 2),
        summary="每个项目解析到哪个 xlsm",
        ref="references/sync-commands.md#list",
        tests=("contract.list.success", "contract.list.target-not-found"),
    ),
    CommandSpec(
        verb="doctor", entry="sync", handler="features.core:doctor",
        mutation=MutationMode.NONE, project=ProjectMode.NONE,
        supports_json=True, exit_codes=(0, 1, 2),
        args=(ArgSpec("--verify-package", kind="flag",
                      help="按 .installed.json 逐项复核交付件哈希与版本（默认不校验）"),),
        summary="环境体检（不合格退出 1）",
        ref="references/sync-commands.md#doctor",
        tests=("contract.doctor.success", "contract.doctor.failed",
               "contract.doctor.verify-package"),
    ),
    CommandSpec(
        verb="scrub", entry="sync", handler="features.core:scrub",
        mutation=MutationMode.NONE, project=ProjectMode.NONE,
        supports_json=True, exit_codes=(0, 1, 2),
        args=(ArgSpec("--path", kind="path",
                      help="要扫描的包根（默认当前工具包根）"),),
        summary="按 scrub.deny 扫交付件，命中退出 1",
        ref="references/sync-commands.md#scrub",
        tests=("contract.scrub.success", "contract.scrub.hit",
               "contract.scrub.config-invalid"),
    ),
)
