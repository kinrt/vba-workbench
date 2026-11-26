# -*- coding: utf-8 -*-
"""sync_push 切片内聚段：公共库占位壳生成 / 清除。

为项目用到的公共模块放同名注释壳；写回时占位壳跳过，避免把真公共模块覆盖成空壳。
``--apply`` 才真正写入/删除；默认预览只计数。
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from features.common_subset import resolve_common_subset, warn_unknown_modules
from toolkit.base import ROOT, SOURCE_EXTS, abs_path
from toolkit.contract import Outcome
from toolkit.files import write_source
from toolkit.repo import is_shell_file, module_index, shell_text


def _common_subset(proj: dict, defaults: dict) -> tuple[set[str] | None, str]:
    """这个项目要写回哪些公共模块；None 表示整个 common。"""
    subset, reason, ignored = resolve_common_subset(proj, defaults)
    warn_unknown_modules(ignored)
    return subset, reason


def _write_shells(proj: dict, defaults: dict, clean: bool = False, dry: bool = False) -> int:
    """为项目用到的公共模块放同名注释壳；clean 只删。返回处理的壳数量。"""
    src_dir = abs_path(proj["src"])
    if not src_dir.is_dir():
        return 0

    common_dir = abs_path(defaults.get("common_dir", "src/common"))
    all_mods = module_index(common_dir)
    subset, _reason = _common_subset(proj, defaults)
    listed = set(all_mods) if subset is None else set(subset)

    # 已有壳代表"这张表确实用过这个公共模块"，继续保留
    for f in sorted(src_dir.rglob("*")):
        if f.is_file() and is_shell_file(f) and f.stem in all_mods:
            listed.add(f.stem)

    # 清掉失效旧壳（公共库已没有这个模块）；clean 模式删全部壳
    removed = 0
    for f in sorted(src_dir.rglob("*")):
        if f.is_file() and is_shell_file(f) and (clean or f.stem not in listed):
            removed += 1
            if not dry:
                f.unlink()
    if clean:
        return removed

    # 项目里留着同名真身时绝不覆盖（提示走 promote 搬进公共库）
    real_stems: dict[str, Path] = {}
    for f in src_dir.rglob("*"):
        if f.is_file() and f.suffix.lower() in SOURCE_EXTS and not is_shell_file(f):
            real_stems.setdefault(f.stem, f)

    written = 0
    for stem in sorted(listed):
        if stem in real_stems:
            continue
        src_file = all_mods[stem]
        try:
            rel_in_repo = src_file.relative_to(ROOT).as_posix()
        except ValueError:
            rel_in_repo = src_file.as_posix()
        folder_name = src_file.parent.name
        target = src_dir / folder_name / (stem + src_file.suffix)
        written += 1
        if dry:
            continue
        target.parent.mkdir(parents=True, exist_ok=True)
        write_source(target, shell_text(rel_in_repo, src_file.suffix.lower()))
    return written


def shells(ctx) -> Outcome:
    """为所有项目生成 / 清除公共库占位文件；--apply 才真正写入/删除。"""
    args = ctx.args
    cfg = ctx.config
    defaults = cfg["defaults"]
    clean = getattr(args, "mode", "create") == "clean"
    dry = not bool(getattr(args, "apply", False))

    rows: list[dict[str, Any]] = []
    total = 0
    for proj in cfg["projects"]:
        n = _write_shells(proj, defaults, clean=clean, dry=dry)
        rows.append({"name": proj["name"], "ok": True,
                     "action": "clean" if clean else "create", "count": n})
        if not clean:
            total += n

    return Outcome(
        action="sync.shells",
        ok=True,
        details=rows,
        stats={"results": rows, "created": 0 if clean else total,
               "cleaned": bool(clean), "dry": dry},
        human=[],
        exit_code=0,
        elapsed_ms=ctx.elapsed_ms(),
    )
