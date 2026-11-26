# -*- coding: utf-8 -*-
"""
toolkit.config —— mapping.json 的加载、校验与写回
============================================================
load_config 只读不猜：结构不对就抛 Failure，绝不在坏配置上继续跑。
同目录若有 `mapping.local.json`（本机覆盖层，不进公开包），先合并再校验：
`defaults` 按键覆盖、`projects` 按 name 合并（同名项目只覆盖写到的字段，新名字追加，
顶层 `_drop` 列出的项目剔除）。
save_config 统一 UTF-8 无 BOM + LF + 缩进 2（与手工编辑保持一致）。
配置位置由 `base.find_config()` 决定（环境变量 → 包根向上逐级 → %APPDATA%），
加载时把 defaults 里的 `require_version` 下发给 repo.resolve_target 作为缺省。
"""

from __future__ import annotations

import json

from .base import CONFIG_PATH, PKG_ROOT, Failure, entry_command, note, warn

LOCAL_NAME = "mapping.local.json"
from .repo import set_defaults


def _is_legacy() -> bool:
    """配置放在工具包根（应在仓库根，与 vba-workbench/ 平级）。"""
    return CONFIG_PATH.parent == PKG_ROOT


def config_source() -> str:
    """配置来源一句话（给 list / doctor 显示，排查"改了不生效"用）。"""
    if _is_legacy():
        return f"{CONFIG_PATH}（和工具同层，建议放到仓库根）"
    return str(CONFIG_PATH)


def local_config_path():
    """覆盖层位置：永远跟着主配置走（同一个目录、固定文件名）。"""
    return CONFIG_PATH.with_name(LOCAL_NAME)


def merge_local(base: dict, over: dict) -> dict:
    """
    合并覆盖层（纯函数，方便单测）：
      · defaults / 其它字典型键：按键覆盖
      · projects：按 name 合并，同名项目只覆盖写到的字段；新名字追加
      · 顶层 `_drop`：列出要剔除的项目名
    """
    out = dict(base)
    for key, val in over.items():
        if key in ("projects", "_drop"):
            continue
        if isinstance(val, dict) and isinstance(out.get(key), dict):
            out[key] = {**out[key], **val}
        else:
            out[key] = val

    drop = set(over.get("_drop") or [])
    projects = [p for p in (out.get("projects") or []) if p.get("name") not in drop]
    by_name = {p.get("name"): i for i, p in enumerate(projects)}
    for proj in over.get("projects") or []:
        name = proj.get("name")
        if name in by_name:
            projects[by_name[name]] = {**projects[by_name[name]], **proj}
        else:
            projects.append(dict(proj))
    out["projects"] = projects
    return out


def load_raw_config() -> dict:
    """原样读 JSON（不校验、不补默认），给"追加一条项目"这类原地改写用。"""
    if not CONFIG_PATH.exists():
        raise Failure(f"找不到配置文件：{CONFIG_PATH}")
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise Failure(f"mapping.json 不是合法 JSON：{e}")


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        raise Failure(f"找不到配置文件：{CONFIG_PATH}（用 `{entry_command('sync.py', 'init', '--project', '甲', '--dir', '<表目录>', '--apply')}` 生成，或抄 assets/mapping.example.json）")
    if _is_legacy():
        warn("配置放在工具包根：请把 mapping.json 移到仓库根（与 vba-workbench/ 平级）")

    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise Failure(f"mapping.json 不是合法 JSON：{e}")

    local = local_config_path()
    if local.exists() and isinstance(data, dict):
        try:
            over = json.loads(local.read_text(encoding="utf-8"))
        except json.JSONDecodeError as e:
            raise Failure(f"{LOCAL_NAME} 不是合法 JSON：{e}")
        if isinstance(over, dict):
            data = merge_local(data, over)
            note(f"已合并本机覆盖层：{local}")

    defaults = {
        "common_dir": "src/common",
        "backup_dir": ".backup",
        "import_document_modules": False,
        "require_version": True,
    }

    if isinstance(data, list):
        projects = data
    else:
        defaults.update(data.get("defaults") or {})
        projects = data.get("projects") or []

    if not projects:
        raise Failure("mapping.json 里没有任何 projects 配置")

    names = [p.get("name") or "" for p in projects]
    dup = sorted({n for n in names if names.count(n) > 1})
    if dup:
        raise Failure(f"项目名重复：{dup}")
    if any(not n for n in names):
        raise Failure("每个 project 都必须有 name")

    set_defaults(defaults)
    return {"defaults": defaults, "projects": projects,
            "scrub": data.get("scrub") if isinstance(data, dict) else {}}


def save_config(data: dict) -> None:
    """把整份 mapping.json 数据写回：UTF-8 无 BOM + LF + 缩进 2。"""
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n",
                           encoding="utf-8", newline="\n")
