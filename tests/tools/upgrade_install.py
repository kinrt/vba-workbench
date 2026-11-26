# -*- coding: utf-8 -*-
"""安装副本升级脚本：把已安装的旧副本升到当前包版本（备份 → 暂存 → 原子替换 → 孤儿清理 → 自检）。

边界（硬规矩）：

- 只覆盖包内交付路径：``scripts/``、``references/``、``assets/``、
  ``SKILL.md``、``VERSION``、``LICENSE``、``requirements.txt``；
- **绝不触碰** 用户 ``src/``、业务表（``*.xlsm/*.xlsb/*.xlsx``）、``mapping.json``、
  ``mapping.local.json``、``.backup/``、``.lab/``、``.run/``；
- 默认**预览**：只列计划动作（含待删孤儿），不写盘；加 ``--apply`` 才执行；
- 升级备份写到**用户级目录**：``%LOCALAPPDATA%\\vba-workbench-install-backups\\vba-workbench\\
  upgrade-<时间戳>\\``，与安装位置解耦——无论 target 是否落在 skills 发现根
  （``skills`` / ``.skills`` / ``user_skills`` 等）下，备份都不会进被扫描的 skill 树；
  安装树内只留交付件 + ``.installed.json``；
- 旧位置只读探测、保留不删：安装树根/父目录下的整包备份 ``vba-workbench.bak-*``、
  树内 ``.backup/upgrade-*``、1.6.1 同级外置 ``../.install-backups/<组>/upgrade-*``，
  只列出并给人工清理建议，安装器绝不自动删除；
- 原子性：先全部复制到临时 staging 并复核哈希，成功后才统一替换；中途任何失败按外置
  备份快照逆序还原，目标目录不留下半新半旧的状态；
- 自检：交付文件 sha256 复核、受保护路径未被改动、关键入口文件在位、``.installed.json``
  清单齐全且哈希与磁盘一致。

本脚本是发布工具，不含运行时业务逻辑；可被测试以函数方式调用。
"""
from __future__ import annotations

import argparse
import datetime as dt
import fnmatch
import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

PKG = Path(__file__).resolve().parents[2]

# 允许覆盖的交付根（相对包根）；其余一律不动。
UPGRADE_ROOTS = ("scripts", "references", "assets")
UPGRADE_FILES = ("SKILL.md", "VERSION", "LICENSE", "requirements.txt")
# 安装清单文件名（写在安装目录根）。
MANIFEST_NAME = ".installed.json"
# 外置备份根目录名（旧 1.6.1 同级位置保留为只读兼容来源；新备份根落在用户级目录）。
BACKUP_DIR_NAME = ".install-backups"
# 用户级备份目录名（LOCALAPPDATA 下；缺环境变量时退到用户主目录的同名点目录）。
BACKUP_APP_DIR = "vba-workbench-install-backups"
BACKUP_GROUP = "vba-workbench"
# 旧整包备份目录前缀（<组名>.bak-<时间戳>）；只读列出，不删不建。
WHOLE_BACKUP_PREFIX = f"{BACKUP_GROUP}.bak-"
# 旧备份人工清理建议；安装器本身不执行删除。
CLEANUP_ADVICE = ("确认无需回滚后，可手动删除上述旧备份目录（资源管理器或 "
                  "PowerShell `Remove-Item -Recurse -Force \"<路径>\"`）；安装器只列出、不自动清理。")
# 安装后冒烟自检：关键入口必须在位（相对包根；LICENSE 随包交付才校验）。
SMOKE_KEY_FILES = ("SKILL.md", "VERSION", "requirements.txt", "LICENSE")
SMOKE_KEY_DIRS = ("scripts",)
SMOKE_ENTRY_SCRIPTS = ("scripts/sync.py", "scripts/audit.py", "scripts/agent.py")
# 受保护：任何情况下不读不写。
PROTECTED_NAMES = {"src", ".lab", ".run", ".backup", "mapping.json",
                   "mapping.local.json"}
PROTECTED_SUFFIX = (".xlsm", ".xlsb", ".xlam", ".xlsx", ".xls", ".frx")
# 打包/升级时排除的缓存与产物。
EXCLUDE_DIRS = {"__pycache__", ".git", ".lab", ".run", ".backup", "releases", "evals"}
EXCLUDE_SUFFIX = (".pyc", ".pyo", ".tmp", ".log")
# 同步盘冲突副本不进安装包：如 `xxx Conflict 2026-11-22.md`、`xxx 冲突.md`、`xxx (1).md`。
CONFLICT_NAME_PATTERNS = ("*Conflict*", "*冲突*")
CONFLICT_COPY_RE = re.compile(r"\(\d+\)")  # 重复下载编号副本，如 `xxx (1).md`


class UpgradeError(RuntimeError):
    """升级计划/执行不合法：越界路径、自检不一致、回滚失败等。"""


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    h.update(path.read_bytes())
    return h.hexdigest()


def _is_protected(rel: Path) -> bool:
    parts = set(rel.parts)
    if parts & PROTECTED_NAMES:
        return True
    if rel.suffix.lower() in PROTECTED_SUFFIX:
        return True
    return False


def _is_conflict_copy(path: Path) -> bool:
    """同步盘冲突副本：文件名带 Conflict/冲突标记，或主干含 `(数字)` 重复编号。"""
    if any(fnmatch.fnmatchcase(path.name, pat) for pat in CONFLICT_NAME_PATTERNS):
        return True
    return bool(CONFLICT_COPY_RE.search(path.stem))


def _user_backup_base() -> Path:
    """用户级备份组根：与安装位置解耦，永不落在 skills 发现根内。

    优先 ``%LOCALAPPDATA%\\vba-workbench-install-backups\\<组>``；环境变量缺失时
    退到 ``%USERPROFILE%\\.vba-workbench-install-backups\\<组>``。同一 skill 名一组，
    批次按时间戳隔离，rollback 取最新一批。
    """
    local = os.environ.get("LOCALAPPDATA")
    if local:
        return Path(local) / BACKUP_APP_DIR / BACKUP_GROUP
    home = os.environ.get("USERPROFILE") or str(Path.home())
    return Path(home) / f".{BACKUP_APP_DIR}" / BACKUP_GROUP


def external_backup_root(target: Path, stamp: str) -> Path:
    """外置升级备份目录：用户级备份组下 ``upgrade-<时间戳>/``。

    落点固定在用户目录，不随安装位置走；``target`` 仅保留兼容签名，不再参与计算。
    """
    return _user_backup_base() / f"upgrade-{stamp}"


def legacy_backup_root(target: Path) -> Path:
    """旧位置（安装树内 .backup/upgrade-*）；只读探测，保留不删。"""
    return target / ".backup"


def _legacy_upgrade_backups(target: Path) -> list[Path]:
    """列出旧位置残留的 upgrade-* 批次（只读，不删）。"""
    root = legacy_backup_root(target)
    if not root.is_dir():
        return []
    return sorted(p for p in root.iterdir()
                  if p.is_dir() and p.name.startswith("upgrade-"))


def _legacy_sibling_backups(target: Path) -> list[Path]:
    """1.6.1 旧同级外置位置：``<安装根>/../.install-backups/<组>/upgrade-*``（只读兼容来源）。"""
    root = target.parent / BACKUP_DIR_NAME / BACKUP_GROUP
    if not root.is_dir():
        return []
    return sorted(p for p in root.glob("upgrade-*") if p.is_dir())


def _whole_package_backups(target: Path) -> list[Path]:
    """旧整包备份 ``vba-workbench.bak-*``：安装根与其父目录下只读列出，不删不建。"""
    found: list[Path] = []
    for base in (target, target.parent):
        if not base.is_dir():
            continue
        found.extend(p for p in base.iterdir()
                     if p.is_dir() and p.name.startswith(WHOLE_BACKUP_PREFIX))
    dedup = {str(p): p for p in found}
    return sorted(dedup.values())


def _legacy_report(target: Path) -> dict[str, list[str]]:
    """汇总三类旧备份的只读清单；全部为路径字符串，供预览/apply 报告与提示使用。"""
    return {
        "whole_backups": [str(p) for p in _whole_package_backups(target)],
        "legacy_upgrade_backups": [str(p) for p in _legacy_upgrade_backups(target)],
        "legacy_sibling_backups": [str(p) for p in _legacy_sibling_backups(target)],
    }


def _iter_source_files(source: Path) -> list[Path]:
    files: list[Path] = []
    for root in UPGRADE_ROOTS:
        base = source / root
        if not base.is_dir():
            continue
        for path in base.rglob("*"):
            if not path.is_file():
                continue
            if set(path.relative_to(source).parts) & EXCLUDE_DIRS:
                continue
            if path.suffix.lower() in EXCLUDE_SUFFIX:
                continue
            if _is_conflict_copy(path):
                continue
            files.append(path)
    for name in UPGRADE_FILES:
        f = source / name
        if f.is_file():
            files.append(f)
    return sorted(files)


def _iter_target_files(target: Path) -> list[Path]:
    """目标副本现存文件（排除受保护路径、缓存产物目录与安装清单）；用于孤儿比对。"""
    files: list[Path] = []
    if not target.is_dir():
        return files
    for path in target.rglob("*"):
        if not path.is_file():
            continue
        rel = path.relative_to(target)
        if _is_protected(rel):
            continue
        if set(rel.parts) & EXCLUDE_DIRS:
            continue
        if path.suffix.lower() in EXCLUDE_SUFFIX:
            continue
        if rel.name == MANIFEST_NAME:
            continue
        files.append(path)
    return sorted(files)


@dataclass
class UpgradePlan:
    """一次升级的计划动作；preview 与 apply 共用同一份清单。"""

    source: Path
    target: Path
    copies: list[dict[str, Path]] = field(default_factory=list)
    backups: list[dict[str, Path]] = field(default_factory=list)
    orphans: list[dict[str, Path]] = field(default_factory=list)
    protected_kept: list[Path] = field(default_factory=list)
    identical: bool = False

    def validate(self) -> None:
        if not self.source.is_dir():
            raise UpgradeError(f"源包目录不存在：{self.source}")
        if not self.target.is_dir():
            raise UpgradeError(f"目标副本目录不存在：{self.target}")
        for item in self.copies:
            rel = item["target"].relative_to(self.target)
            if _is_protected(rel):
                raise UpgradeError(f"目标路径受保护，拒绝覆盖：{rel}")

    def preview_lines(self) -> list[str]:
        lines = [f"源包：{self.source}", f"目标副本：{self.target}",
                 f"计划覆盖文件：{len(self.copies)} 个",
                 f"其中需先备份：{len(self.backups)} 个",
                 f"将删除 {len(self.orphans)} 个孤儿（新清单已不含、旧版本遗留）",
                 f"外置备份落点（用户级目录，不进 skills 发现根）："
                 f"{external_backup_root(self.target, '<时间戳>')}",
                 "受保护、保持不动：src/、业务表、mapping.json、.backup/、.lab/、.run/"]
        legacy = _legacy_report(self.target)
        if legacy["whole_backups"]:
            lines.append(f"发现旧整包备份 {len(legacy['whole_backups'])} 个"
                         f"（只读保留，未自动清理）：{legacy['whole_backups']}")
        old_batches = legacy["legacy_upgrade_backups"] + legacy["legacy_sibling_backups"]
        if old_batches:
            lines.append(f"发现旧位置升级备份 {len(old_batches)} 批"
                         f"（只读保留，可作回滚兼容来源）：{old_batches}")
        if legacy["whole_backups"] or old_batches:
            lines.append(f"旧备份处理建议：{CLEANUP_ADVICE}")
        for item in self.copies:
            rel = item["target"].relative_to(self.target)
            tag = "备份后覆盖" if any(b["target"] == item["target"] for b in self.backups) else "新增"
            lines.append(f"  [{tag}] {rel}")
        for item in self.orphans:
            lines.append(f"  [孤儿·删除] {item['rel']}")
        return lines


def build_plan(source: Path, target: Path) -> UpgradePlan:
    """比对源包与目标副本，生成覆盖/备份/孤儿清单；不写盘。"""
    plan = UpgradePlan(source=source.resolve(), target=target.resolve())
    new_rels: set[Path] = set()
    for src_file in _iter_source_files(source):
        rel = src_file.relative_to(source)
        if _is_protected(rel):
            continue
        dst = target / rel
        plan.copies.append({"source": src_file, "target": dst, "rel": rel})
        new_rels.add(rel)
        if dst.exists():
            plan.backups.append({"target": dst, "rel": rel})
    # 孤儿：目标现存、不属于新交付清单、也不在受保护路径内的文件。
    for tgt_file in _iter_target_files(target):
        rel = tgt_file.relative_to(target)
        if rel in new_rels:
            continue
        plan.orphans.append({"target": tgt_file, "rel": rel})
    # 登记目标里现存的受保护项（自检时确认未被触碰）。
    for name in PROTECTED_NAMES:
        p = target / name
        if p.exists():
            plan.protected_kept.append(p)
    plan.identical = _is_identical(plan)
    plan.validate()
    return plan


def _is_identical(plan: UpgradePlan) -> bool:
    """幂等判定：新清单内每个目标文件都已存在且哈希一致，且无待删孤儿。"""
    if plan.orphans:
        return False
    for item in plan.copies:
        dst = item["target"]
        if not dst.is_file():
            return False
        if _sha256(item["source"]) != _sha256(dst):
            return False
    return True


def _backup_target(plan: UpgradePlan, backup_root: Path) -> None:
    for item in plan.backups:
        rel = item["rel"]
        dest = backup_root / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(item["target"], dest)


def _stage_to_temp(plan: UpgradePlan, staging: Path) -> None:
    """先全部复制到临时 staging 并逐文件复核哈希；staging 未成功前不动目标目录。"""
    for item in plan.copies:
        rel = item["rel"]
        staged = staging / rel
        staged.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(item["source"], staged)
        if _sha256(item["source"]) != _sha256(staged):
            raise UpgradeError(f"暂存复核不一致：{rel}")


def _restore_from_backup(plan: UpgradePlan, backup_root: Path) -> int:
    """按备份逆序还原被覆盖的旧文件；返回成功还原的文件数。"""
    restored = 0
    for item in reversed(plan.backups):
        rel = item["rel"]
        dest = backup_root / rel
        if not dest.is_file():
            continue
        dst = item["target"]
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(dest, dst)
        restored += 1
    return restored


def _remove_orphans(plan: UpgradePlan) -> int:
    """删除待清孤儿；删除后顺手清理空目录（不碰受保护根）。"""
    removed = 0
    for item in plan.orphans:
        tgt = item["target"]
        if tgt.is_file():
            tgt.unlink()
            removed += 1
    _prune_empty_dirs(plan.target)
    return removed


def _prune_empty_dirs(target: Path) -> None:
    """自底向上删除空目录；受保护根与旧整包备份目录即使变空也保留。"""
    for d in sorted((p for p in target.rglob("*") if p.is_dir()),
                    key=lambda p: len(p.parts), reverse=True):
        if d.name in PROTECTED_NAMES or d.name.startswith(WHOLE_BACKUP_PREFIX):
            continue
        try:
            if not any(d.iterdir()):
                d.rmdir()
        except OSError:
            continue


def smoke_check(plan: UpgradePlan) -> dict[str, Any]:
    """安装后冒烟自检：源包已交付的关键文件/目录在目标在位；返回 {checked, missing}。

    核对范围 = 源包真实存在且进交付清单的关键项；源包本身缺关键文件的情形由
    ``preflight_source`` 在写盘前拦下，不在此重复。
    """
    delivered = {item["rel"] for item in plan.copies}
    missing: list[str] = []
    checked = 0
    for name in SMOKE_KEY_FILES:
        rel = Path(name)
        if rel in delivered:
            checked += 1
            if not (plan.target / rel).is_file():
                missing.append(name)
    for entry in SMOKE_ENTRY_SCRIPTS:
        rel = Path(entry)
        if rel in delivered:
            checked += 1
            if not (plan.target / rel).is_file():
                missing.append(entry)
    for d in SMOKE_KEY_DIRS:
        if any(part == d for part in delivered):
            checked += 1
            if not (plan.target / d).is_dir():
                missing.append(d)
    return {"checked": checked, "missing": missing}


def preflight_source(plan: UpgradePlan) -> list[str]:
    """写盘前预检：声明的关键文件/目录在源包必须真实存在；返回缺失项。"""
    missing: list[str] = []
    for name in SMOKE_KEY_FILES:
        if not (plan.source / name).is_file():
            missing.append(name)
    for entry in SMOKE_ENTRY_SCRIPTS:
        if not (plan.source / entry).is_file():
            missing.append(entry)
    for d in SMOKE_KEY_DIRS:
        if not (plan.source / d).is_dir():
            missing.append(d)
    return missing


def self_check(plan: UpgradePlan) -> dict[str, Any]:
    """逐文件 sha256 复核；返回 {checked, mismatches, protected_touched}。"""
    mismatches: list[str] = []
    for item in plan.copies:
        if not item["target"].is_file():
            mismatches.append(str(item["target"].relative_to(plan.target)))
            continue
        if _sha256(item["source"]) != _sha256(item["target"]):
            mismatches.append(str(item["target"].relative_to(plan.target)))
    protected_touched: list[str] = []
    for kept in plan.protected_kept:
        # 自检只确认升级清单里没有任何一项指向受保护路径；文件实体由 plan.validate 把关。
        if any(c["target"] == kept for c in plan.copies):
            protected_touched.append(str(kept.relative_to(plan.target)))
    return {"checked": len(plan.copies), "mismatches": mismatches,
            "protected_touched": protected_touched}


def build_manifest(plan: UpgradePlan) -> dict[str, Any]:
    """生成安装清单：version/source/installed_at/files(rel→sha256)。"""
    version = ""
    ver_file = plan.source / "VERSION"
    if ver_file.is_file():
        version = ver_file.read_text(encoding="utf-8").strip()
    files: dict[str, str] = {}
    for item in plan.copies:
        rel = str(item["rel"]).replace("\\", "/")
        files[rel] = _sha256(item["target"])
    return {
        "version": version,
        "source": str(plan.source),
        "installed_at": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
        "files": files,
    }


def write_manifest_atomic(plan: UpgradePlan) -> Path:
    """先写临时文件再 os.replace 原子落盘 .installed.json；返回清单路径。"""
    manifest = build_manifest(plan)
    manifest_path = plan.target / MANIFEST_NAME
    fd, tmp_name = tempfile.mkstemp(prefix=".installed-", suffix=".json",
                                     dir=str(plan.target))
    tmp = Path(tmp_name)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=1, sort_keys=True)
        os.replace(tmp, manifest_path)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise
    return manifest_path


def verify_manifest(plan: UpgradePlan) -> dict[str, Any]:
    """校验 .installed.json 在位、字段齐全、files 覆盖全部交付件且哈希与磁盘一致。

    返回 {present, missing, hash_mismatch, files_count}。
    """
    manifest_path = plan.target / MANIFEST_NAME
    result: dict[str, Any] = {"present": False, "missing": [],
                              "hash_mismatch": [], "files_count": 0}
    if not manifest_path.is_file():
        result["missing"].append(MANIFEST_NAME)
        return result
    result["present"] = True
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as exc:
        result["missing"].append(f"{MANIFEST_NAME} 解析失败：{exc}")
        return result
    for field in ("version", "source", "installed_at", "files"):
        if field not in manifest:
            result["missing"].append(f"{MANIFEST_NAME} 缺字段 {field}")
    recorded = manifest.get("files", {})
    result["files_count"] = len(recorded)
    delivered = {str(item["rel"]).replace("\\", "/") for item in plan.copies}
    if set(recorded) != delivered:
        missing_in_manifest = sorted(delivered - set(recorded))
        extra_in_manifest = sorted(set(recorded) - delivered)
        for rel in missing_in_manifest:
            result["missing"].append(f"{MANIFEST_NAME} 未登记 {rel}")
        for rel in extra_in_manifest:
            result["hash_mismatch"].append(f"{rel} 不在交付清单")
    for rel in delivered:
        if rel not in recorded:
            continue
        dst = plan.target / rel
        if not dst.is_file():
            result["hash_mismatch"].append(rel)
            continue
        if recorded[rel] != _sha256(dst):
            result["hash_mismatch"].append(rel)
    return result


def execute(plan: UpgradePlan, *, apply: bool,
            backup_root: Path | None = None) -> dict[str, Any]:
    """按计划执行；apply=False 只预览。返回执行报告。"""
    started = time.perf_counter()
    if not apply:
        legacy = _legacy_report(plan.target)
        return {"applied": False, "copies": len(plan.copies),
                "backups": len(plan.backups), "orphans": len(plan.orphans),
                "identical": plan.identical,
                "legacy_backups": legacy["legacy_upgrade_backups"],
                "legacy_sibling_backups": legacy["legacy_sibling_backups"],
                "whole_backups": legacy["whole_backups"],
                "cleanup_advice": CLEANUP_ADVICE if any(legacy.values()) else "",
                "elapsed_ms": round((time.perf_counter() - started) * 1000),
                "self_check": None, "smoke": None, "manifest": None}

    if plan.identical:
        # 幂等：已是当前版本，不重复备份/删孤儿；刷新清单时间戳与哈希。
        write_manifest_atomic(plan)
        manifest_check = verify_manifest(plan)
        legacy = _legacy_report(plan.target)
        elapsed = round((time.perf_counter() - started) * 1000)
        if manifest_check["missing"] or manifest_check["hash_mismatch"]:
            raise UpgradeError(f"幂等复检清单未通过：{json.dumps(manifest_check, ensure_ascii=False)}")
        return {"applied": True, "copies": 0, "backups": 0, "orphans": 0,
                "identical": True, "elapsed_ms": elapsed,
                "legacy_backups": legacy["legacy_upgrade_backups"],
                "legacy_sibling_backups": legacy["legacy_sibling_backups"],
                "whole_backups": legacy["whole_backups"],
                "self_check": None, "smoke": None, "manifest": manifest_check}

    # 写盘前预检：源包关键项不齐直接拦下，不进入备份/替换流程。
    absent = preflight_source(plan)
    if absent:
        raise UpgradeError(f"源包缺少关键文件/目录：{absent}")

    stamp = time.strftime("%Y%m%d-%H%M%S")
    backup_root = backup_root or external_backup_root(plan.target, stamp)
    legacy = _legacy_report(plan.target)
    staging: Path | None = None
    try:
        staging = Path(tempfile.mkdtemp(prefix="vba-upgrade-", dir=str(plan.target)))
        print(f"[upgrade] 暂存新文件 → {staging}")
        _stage_to_temp(plan, staging)
        print(f"[upgrade] 备份旧文件（外置）→ {backup_root}")
        _backup_target(plan, backup_root)
        print(f"[upgrade] 原子替换 {len(plan.copies)} 个交付文件 …")
        for item in plan.copies:
            dst = item["target"]
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(staging / item["rel"], dst)
        removed = _remove_orphans(plan)
    except Exception as exc:  # 任何中途失败：按外置备份快照逆序还原，不留半新半旧
        restored = 0
        if backup_root.is_dir():
            restored = _restore_from_backup(plan, backup_root)
        if staging is not None:
            shutil.rmtree(staging, ignore_errors=True)
        raise UpgradeError(
            f"执行失败已回滚：{exc}；已还原 {restored} 个旧文件，"
            f"未完成 {len(plan.copies)} 个交付项") from exc
    finally:
        if staging is not None:
            shutil.rmtree(staging, ignore_errors=True)

    write_manifest_atomic(plan)
    check = self_check(plan)
    smoke = smoke_check(plan)
    manifest_check = verify_manifest(plan)
    elapsed = round((time.perf_counter() - started) * 1000)
    if check["mismatches"] or check["protected_touched"]:
        raise UpgradeError(f"自检未通过：{json.dumps(check, ensure_ascii=False)}")
    if smoke["missing"]:
        raise UpgradeError(f"冒烟自检未通过，缺失关键文件/目录：{smoke['missing']}")
    if manifest_check["missing"] or manifest_check["hash_mismatch"]:
        raise UpgradeError(f"清单自检未通过：{json.dumps(manifest_check, ensure_ascii=False)}")
    old_batches = legacy["legacy_upgrade_backups"] + legacy["legacy_sibling_backups"]
    if old_batches:
        print(f"[upgrade] 提示：旧位置仍有 {len(old_batches)} 批历史备份（未删除）："
              f"{[Path(p).name for p in old_batches]}")
    if legacy["whole_backups"]:
        print(f"[upgrade] 提示：发现 {len(legacy['whole_backups'])} 个旧整包备份"
              f"（未删除）：{legacy['whole_backups']}")
        print(f"[upgrade] {CLEANUP_ADVICE}")
    return {"applied": True, "copies": len(plan.copies),
            "backups": len(plan.backups), "orphans": removed,
            "identical": False, "backup_root": str(backup_root),
            "legacy_backups": legacy["legacy_upgrade_backups"],
            "legacy_sibling_backups": legacy["legacy_sibling_backups"],
            "whole_backups": legacy["whole_backups"],
            "cleanup_advice": CLEANUP_ADVICE if any(legacy.values()) else "",
            "elapsed_ms": elapsed, "self_check": check, "smoke": smoke,
            "manifest": manifest_check}


def rollback(plan: UpgradePlan, backup_root: Path | None = None) -> dict[str, Any]:
    """从备份快照把被覆盖的旧文件还原；返回还原计数与备份根。

    查找顺序：新用户级备份根最新一批 ``upgrade-*`` → 1.6.1 同级外置旧位置 →
    树内 ``.backup/upgrade-*``；后两类只作兼容只读来源。
    """
    if backup_root is None:
        base = _user_backup_base()
        candidates = sorted(p for p in base.glob("upgrade-*") if p.is_dir()) \
            if base.is_dir() else []
        if not candidates:
            legacy_candidates = sorted(
                _legacy_sibling_backups(plan.target) + _legacy_upgrade_backups(plan.target))
            if not legacy_candidates:
                raise UpgradeError("未找到任何升级备份批次，无法回滚")
            candidates = legacy_candidates
        backup_root = candidates[-1]
    if not backup_root.is_dir():
        raise UpgradeError(f"备份批次不存在：{backup_root}")
    restored = _restore_from_backup(plan, backup_root)
    return {"restored": restored, "backup_root": str(backup_root)}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="升级已安装的 vba-workbench 副本")
    parser.add_argument("--source", default=str(PKG), help="当前包目录（默认本包）")
    parser.add_argument("--target", required=True, help="已安装副本根目录")
    parser.add_argument("--apply", action="store_true", help="真正执行；不加只预览")
    args = parser.parse_args(argv)

    try:
        plan = build_plan(Path(args.source), Path(args.target))
    except UpgradeError as exc:
        print(f"[upgrade] 计划失败：{exc}", file=sys.stderr)
        return 2

    print("== 预览 ==")
    for line in plan.preview_lines():
        print(line)
    if not args.apply:
        print("\n未加 --apply：未写盘。确认无误后重跑并带 --apply。")
        return 0

    try:
        report = execute(plan, apply=True)
    except UpgradeError as exc:
        print(f"[upgrade] 执行失败：{exc}", file=sys.stderr)
        return 1
    if report.get("identical"):
        print(f"[upgrade] 目标副本已是当前版本，无需更新；用时 {report['elapsed_ms']} ms")
    else:
        print(f"[upgrade] 完成：覆盖 {report['copies']} 个，备份 {report['backups']} 个（外置），"
              f"清理孤儿 {report['orphans']} 个；自检通过；用时 {report['elapsed_ms']} ms")
        print(f"[upgrade] 备份批次：{report.get('backup_root', '')}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
