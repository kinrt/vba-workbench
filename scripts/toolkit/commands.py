# -*- coding: utf-8 -*-
"""应用层命令登记：每个 feature 模块只声明自己的 SPECS，registry 负责汇总。"""
from __future__ import annotations

from .registry import CommandSpec

# 10 个切片的固定顺序；新增模块只追加，不在入口里写 if verb ==
FEATURE_MODULES = (
    "features.core",
    "features.sync_read",
    "features.sync_push",
    "features.sync_transform",
    "features.run",
    "features.agent_io",
    "features.agent_patch",
    "features.audit_a",
    "features.audit_b",
    "features.release",
)


def load_specs() -> tuple[CommandSpec, ...]:
    import importlib

    rows: list[CommandSpec] = []
    seen: dict[str, str] = {}
    for module_name in FEATURE_MODULES:
        module = importlib.import_module(module_name)
        for spec in getattr(module, "SPECS", ()):
            key = f"{spec.entry}.{spec.verb}"
            if key in seen:
                raise ValueError(f"命令重复登记：{key}（{seen[key]} / {module_name}）")
            seen[key] = module_name
            rows.append(spec)
    return tuple(rows)


SPECS = load_specs()


def specs_for(entry: str) -> tuple[CommandSpec, ...]:
    return tuple(spec for spec in SPECS if spec.entry == entry)
