# -*- coding: utf-8 -*-
"""features 内共享叶子：子集判定 + 项目查找 + VBE 内存快照/复原/关闭实例。

判定/查找/快照/复原规则只在这里维护，调用方只做参数适配，不再各自复制一份
``common_reference_reasons`` / ``common_version_stems`` / ``resolve_common_subset`` /
``find_project`` / ``snapshot_module_code`` / ``restore_module_code`` / ``close_owned``。
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from toolkit.base import (SOURCE_EXTS, abs_path, entry_command,
                          read_source_text, warn)
from toolkit.contract import ErrorCode, Failure, Stage
from toolkit.excel import (find_component, quit_owned, replace_module_code)
from toolkit.repo import VERSION_RE, is_shell_file, module_index
from toolkit.vba_parse import declared_procs, mask_strings, strip_vba_comments

# 同进程内同一批未知模块名只告警一次：run 一处解析、多个辅助取子集点会重复进入，
# 避免同一遍命令把同一句话刷两遍。
_WARNED_UNKNOWN: set[str] = set()


def clear_unknown_warning_cache() -> None:
    """清空未知模块告警去重缓存（测试隔离用；正常一次 CLI 进程只走一个出口）。"""
    _WARNED_UNKNOWN.clear()


def scope_payload(subset: set[str] | None) -> dict[str, Any]:
    """dry 预览的公共模块范围：``all``=未配置/全部写回；``subset``=子集名单（排序稳定）。"""
    if subset is None:
        return {"common_scope": "all", "common_modules": []}
    return {"common_scope": "subset", "common_modules": sorted(subset)}


def scope_human(subset: set[str] | None, reason: str, *, action: str = "写回") -> str:
    """dry 人话行：本次将处理的公共模块名单与判定来源（白名单/auto/未配置）。"""
    names = "全部公共模块" if subset is None else "、".join(sorted(subset))
    return f"本次将{action}公共模块：{names}（{reason}）"


def warn_unknown_modules(ignored: list[str]) -> None:
    """mapping 的 common_modules 里写了公共库不存在的模块名：统一文案与出口。

    走 ``toolkit.base.warn``：JSON 模式收进缓存、由 render 收口到 stats.warnings，
    不漏到 stderr；人话模式打 ``[警告]`` 前缀。四个入口（list/pull/push/run）同一句。
    """
    if not ignored:
        return
    key = "、".join(sorted(ignored))
    if key in _WARNED_UNKNOWN:
        return
    _WARNED_UNKNOWN.add(key)
    warn(f"common_modules 里这些模块在公共库不存在，已忽略：{key}")


def find_project(cfg: dict, name: str) -> dict:
    """按名字取项目配置；找不到报 ``PROJECT_NOT_FOUND``（resolve 阶段）。

    纯 dict 查找：不碰 ctx、不写盘、不打日志；sync read / push / run / transform
    四份切片共用这一份，不再各写一遍。
    """
    projects = cfg.get("projects", [])
    proj = next((p for p in projects if p.get("name") == name), None)
    if proj is None:
        names = ", ".join(p.get("name", "") for p in projects)
        raise Failure(
            f"没有名为「{name}」的项目，可用：{names}",
            code=ErrorCode.PROJECT_NOT_FOUND, stage=Stage.RESOLVE,
            suggested=entry_command("sync.py", "list", "--json"),
        )
    return proj


# ---------------------------------------------------------------------------
# VBE 内存快照 / 复原 / 自有实例关闭（写回切片共用；薄封装，策略仍在调用方）
# ---------------------------------------------------------------------------
def snapshot_module_code(vbproj: Any) -> dict[str, str]:
    """写回前把 VBE 内存里的模块正文存一份（只存能读到的，文档模块也存）。"""
    out: dict[str, str] = {}
    try:
        comps = list(vbproj.VBComponents)
    except Exception:
        return out
    for comp in comps:
        try:
            cm = comp.CodeModule
            out[str(comp.Name)] = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
        except Exception:
            continue
    return out


def restore_module_code(vbproj: Any, snapshots: dict[str, str]) -> tuple[list[str], list[str]]:
    """按快照复原；返回 (已复原, 失败)。逐条报失败，不中断其它模块。"""
    restored: list[str] = []
    failed: list[str] = []
    for name, text in snapshots.items():
        comp = find_component(vbproj, name)
        if comp is None:
            failed.append(f"{name}（工程里找不到）")
            continue
        try:
            replace_module_code(comp, text)
            restored.append(name)
        except Exception as e:  # 复原失败逐条报，不中断其它模块
            failed.append(f"{name}（{e}）")
    return restored, failed


def close_owned(wb: Any, owned_xl: Any) -> list[str]:
    """只关这次自己开的副本；借用户已打开的表绝不替他关窗。

    收尾统一走 ``toolkit.excel.quit_owned``：关簿/退出失败不再裸吞异常，问题清单
    返回给调用方按既有 warnings/stats 通道上报。实例内除本次传入簿之外的簿一律
    视为他人参考簿（N2）：有残留就不退出实例，避免无保存关掉宏自开的参考簿。

    返回收尾诊断列表；空列表表示干净。``owned_xl is None``（attach 借来会话）
    直接返回空，绝不碰借来的实例。
    """
    if owned_xl is None or wb is None:
        return []
    keep: list[str] = []
    try:
        target = str(getattr(wb, "Name", ""))
        keep = [str(getattr(b, "Name", "")) for b in owned_xl.Workbooks
                if str(getattr(b, "Name", "")) != target]
    except Exception:  # noqa: BLE001 - 读不到簿名单就退回旧行为：只关传入簿、退出实例
        keep = []
    try:
        return quit_owned(owned_xl, keep_workbooks=keep or None)
    except Exception as e:  # noqa: BLE001 - 收尾本身抛错也不能盖住主流程结果
        return [f"自有实例回收异常：{e}"]


def _ident_re(name: str) -> re.Pattern[str]:
    # \w 在 Python 里包含中文，这样「解析」不会命中「解析字符串」这种更长标识符
    return re.compile(r"(?<!\w)" + re.escape(name) + r"(?!\w)")


def common_version_stems(common_dir: Path) -> set[str]:
    """带 COMMON_VERSION 常量的公共模块：这类始终写回，用于版本校验。"""
    out: set[str] = set()
    for stem, f in module_index(common_dir).items():
        if VERSION_RE.search(read_source_text(f)):
            out.add(stem)
    return out


def common_reference_reasons(proj_src: Path,
                             candidates: dict[str, Path]) -> dict[str, str]:
    """auto 判定：{公共模块名: 依据说明}。

    四条依据，文字证据全部要求“剥掉注释、再抹掉字符串字面量”后再匹配：
      · 剥注释——公共库头部注释里互相提及会被当成引用，壳会越滚越多；
      · 抹字符串——Application.Run("宏名", getTL) 这类软调用的宏名只是字符串，
        不是标识符调用，不抹掉会把整模块灌进不连库的表。
      1. 模块名   —— 项目代码里出现该模块名（标识符级匹配）
      2. 过程名 X —— 项目代码调用 X，且 X 只在公共库这一个模块里声明、项目自己没有同名声明
      3. 闭包←Y  —— 已被选中的公共模块 Y，用它自己的代码按同样两条规则牵连出来
      4. 占位壳   —— 项目目录里已有该模块的同名壳（文件名即证据，不读壳里的注释内容）
    """
    picked: dict[str, str] = {}
    if not proj_src.is_dir():
        return picked

    helpers = [f for f in proj_src.rglob("*")
               if f.is_file() and f.suffix.lower() in SOURCE_EXTS]
    files = [f for f in helpers if not is_shell_file(f)]
    # 壳只取文件名（stem）当作"这张表声明要用该公共模块"的证据，绝不解析其注释正文
    shell_stems = {f.stem for f in helpers if is_shell_file(f)}
    texts = [mask_strings(strip_vba_comments(read_source_text(f))) for f in files]
    proj_defs: set[str] = set()
    for f in files:
        proj_defs |= declared_procs(f)

    # 过程名归属表：{过程名: 声明它的公共模块}
    proc_owner: dict[str, set[str]] = {}
    for stem, path in candidates.items():
        for proc in declared_procs(path):
            proc_owner.setdefault(proc, set()).add(stem)

    def hit(lines: list[str], stem: str) -> str | None:
        if any(_ident_re(stem).search(t) for t in lines):
            return "模块名"
        called = [proc for proc, owners in proc_owner.items()
                  if owners == {stem} and proc not in proj_defs
                  and any(_ident_re(proc).search(t) for t in lines)]
        if called:
            called.sort()
            shown = "、".join(called[:3]) + (" 等" if len(called) > 3 else "")
            return f"过程名 {shown}"
        return None

    for stem in candidates:
        reason = hit(texts, stem)
        if reason:
            picked[stem] = reason

    # 已存在的占位壳：文件名就是"这张表用了该公共模块"的声明，优先级低于文字证据
    for stem in candidates:
        if stem in shell_stems and stem not in picked:
            picked[stem] = "占位壳（项目里已声明使用它）"

    # 传递闭包：公共模块之间互相引用，同样只认"剥注释后的证据"
    changed = True
    while changed:
        changed = False
        for donor in picked.copy():
            donor_text = [mask_strings(strip_vba_comments(
                read_source_text(candidates[donor])))]
            for stem in candidates:
                if stem in picked:
                    continue
                reason = hit(donor_text, stem)
                if reason:
                    picked[stem] = f"闭包←{donor}（{reason}）"
                    changed = True
    return picked


def resolve_common_subset(proj: dict, defaults: dict
                          ) -> tuple[set[str] | None, str, list[str]]:
    """决定这个项目要写回哪些公共模块。

      未配置 common_modules   -> None（写回整个 common，旧行为）
      "auto"                 -> 按引用自动挑（始终带上版本常量模块）
      列出模块名              -> 只用列出的这些
    返回 (模块名集合 或 None, 说明, 白名单里被忽略的不存在模块)
    """
    mode = proj.get("common_modules")
    common_dir = abs_path(defaults.get("common_dir", "src/common"))
    all_mods = module_index(common_dir)

    if mode is None:
        return None, "全部写回（未配置 common_modules）", []

    if isinstance(mode, str) and mode.strip().lower() == "auto":
        picked = set(common_reference_reasons(abs_path(proj["src"]), all_mods))
        picked |= {s for s in common_version_stems(common_dir) if s in all_mods}
        return picked, f"按引用自动挑选（{len(picked)}/{len(all_mods)} 个）", []

    if isinstance(mode, list):
        wanted = {str(x) for x in mode}
        unknown = sorted(wanted - set(all_mods))
        picked = wanted & set(all_mods)
        return picked, f"按配置白名单（{len(picked)}/{len(all_mods)} 个）", unknown

    return None, "配置无法识别，按全部写回处理", []
