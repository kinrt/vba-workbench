# -*- coding: utf-8 -*-
"""sync_transform 切片：align / promote / gbk-export / clean-backup。

四个动词都只编排磁盘源码与备份目录，不碰 COM；handler 只返回 ``Outcome``，
不打印、不自带 JSON。行为与 1.5.6 的 ``sync.py`` 对应动词保持一致：

  · align          把项目里与公共库同名的真实代码换成一行占位壳（默认预览）
  · promote        把一个项目模块搬进公共库（默认预览）
  · gbk-export     把 src 镜像导出为 GBK + CRLF 副本到 .backup/GBK（默认预览）
  · clean-backup   按保留口径清理 .backup 旧批次（默认只预览，--apply 才真删）
"""
from __future__ import annotations

import datetime as dt
import re
from pathlib import Path

from toolkit.base import (ROOT, SOURCE_ALL_EXTS, SOURCE_EXTS, abs_path,
                          entry_command, read_source_text)
from toolkit.contract import Failure, MutationMode, Outcome, Stage
from toolkit.files import (backup_copy, backup_stats, gbk_import_copy,
                           move_into, prune_backups, remove_tree, write_source)
from toolkit.repo import (is_shell_file, module_index, module_name_of,
                          rel_path, shell_text)
from toolkit.registry import ArgSpec, CommandSpec, ProjectMode
from toolkit.check import Notes
from features.common_subset import find_project as _find_project  # noqa: E402

# 源码扩展名 -> 分类目录名（提升、分组展示共用这一份）
_GROUP_BY_SUFFIX = {".bas": "模块", ".cls": "类模块", ".frm": "窗体"}

# 只在源码里出现 “As 前缀.xxx” 才算早期绑定；CreateObject 是晚期绑定，不需要引用。
_REF_RULES: dict[str, str] = {
    "ADODB.": "Microsoft ActiveX Data Objects 6.1 Library",
    "ADOX.": "Microsoft ADO Ext. 6.0 for DDL and Security",
    "DAO.": "Microsoft DAO 3.6 Object Library",
    "Scripting.": "Microsoft Scripting Runtime",
    "MSForms.": "Microsoft Forms 2.0 Object Library",
    "Shell32.": "Microsoft Shell Controls And Automation",
    "SHDocVw.": "Microsoft Internet Controls",
    "MSComctlLib.": "Microsoft Windows Common Controls 6.0 (SP6)",
}
_EARLY_BOUND_RE = {
    pre: re.compile(r"\bAs\s+" + re.escape(pre), re.IGNORECASE) for pre in _REF_RULES
}


# ---------------------------------------------------------------------------
# 私有辅助（纯磁盘逻辑；对应 1.5.6 单体里的同名薄函数）
# ---------------------------------------------------------------------------
def _backup_root(cfg: dict) -> Path:
    """配置里的备份根目录（.backup 那层）。"""
    return abs_path(cfg["defaults"].get("backup_dir", ".backup"))


def _backup_keep_settings(cfg: dict) -> dict:
    """备份保留口径（全部来自 mapping.json 的 defaults）。"""
    d = cfg.get("defaults", {})
    try:
        settings = {
            "keep_recent": int(d.get("backup_keep", 10)),
            "keep_daily": int(d.get("backup_keep_daily", 1)),
            "min_keep": int(d.get("backup_min_keep", 3)),
            "auto": bool(d.get("backup_auto_clean", True)),
        }
    except (TypeError, ValueError) as e:
        raise Failure(f"备份保留口径配置不是整数：{e}", code="CONFIG_INVALID",
                      stage=Stage.CONFIG)
    protect = d.get("backup_protect", [])
    if isinstance(protect, str):
        protect = [protect]
    settings["protect"] = list(protect or [])
    return settings


def _prune_backup_dir(cfg: dict, *, dry_run: bool, keep_recent: int,
                      protect: list[str] | tuple[str, ...] | None) -> dict:
    """按配置口径清理 .backup（keep_recent 传了就临时覆盖配置值）。"""
    s = _backup_keep_settings(cfg)
    tokens = list(s["protect"]) + list(protect or [])
    return prune_backups(
        _backup_root(cfg),
        keep_recent=keep_recent, keep_daily=s["keep_daily"],
        min_keep=s["min_keep"], dry_run=dry_run, protect=tokens,
    )


def _report_prune(result: dict, *, tag: str, verbose: bool) -> tuple[list[str], list[str]]:
    """清理结果拆成人话行与警告行；返回 (human, warnings)。"""
    human: list[str] = []
    warnings: list[str] = []
    deleted = result.get("deleted", [])
    leftover = result.get("leftover", [])
    protected = result.get("protected", [])
    if protected:
        human.append("保护：受 backup_protect / --keep-tag 保护跳过 "
                     f"{len(protected)} 批：" + "、".join(protected[:8])
                     + ("…" if len(protected) > 8 else ""))
    if not deleted:
        if verbose:
            human.append(f"{tag}：没有可清理的（共 {result['scanned']} 批，"
                         f"{result['kept_bytes'] / 1048576:.0f} MB）")
        return human, warnings
    mb = result["freed"] / 1048576
    human.append(f"{tag}：删 {len(deleted)} 批旧备份，释放 {mb:.0f} MB；"
                 f"保留 {result['kept']} 批 / {result['kept_bytes'] / 1048576:.0f} MB")
    if verbose:
        for name, size in deleted:
            human.append(f"    - {name}  {size / 1048576:.1f} MB")
    if leftover:  # 同步盘/杀软占用时 remove_tree 会静默失败
        warnings.append(f"有 {len(leftover)} 批没删掉（可能被同步盘或 Excel 占用，"
                        f"下次清理会再试）：" + "、".join(leftover[:5])
                        + ("…" if len(leftover) > 5 else ""))
    return human, warnings


def _scan_required_refs(folders: list[Path]) -> dict[str, int]:
    """静态扫描源码，返回 {需要勾选的引用: 早期绑定出现次数}。"""
    need: dict[str, int] = {}
    for folder in folders:
        if not folder.is_dir():
            continue
        for f in folder.rglob("*"):
            if not f.is_file() or f.suffix.lower() not in SOURCE_EXTS:
                continue
            text = read_source_text(f)
            for pre, pat in _EARLY_BOUND_RE.items():
                n = len(pat.findall(text))
                if n:
                    desc = _REF_RULES[pre]
                    need[desc] = need.get(desc, 0) + n
    return need


def _rel(path: Path) -> str:
    """显示用路径：仓库内给相对路径，仓库外原样给绝对路径。"""
    try:
        return path.relative_to(ROOT).as_posix()
    except ValueError:
        return path.as_posix()


def _write_promoted_shell(proj: dict, defaults: dict, module_name: str) -> bool:
    """提升后在原位置补一个同名占位壳；项目里还留着同名真实代码时不覆盖。"""
    src_dir = abs_path(proj["src"])
    if not src_dir.is_dir():
        return False
    common_dir = abs_path(defaults.get("common_dir", "src/common"))
    all_mods = module_index(common_dir)
    if module_name not in all_mods:
        return False
    src_file = all_mods[module_name]
    for f in src_dir.rglob("*"):
        if (f.is_file() and f.suffix.lower() in SOURCE_EXTS
                and not is_shell_file(f) and f.stem == module_name):
            return False
    try:
        rel_in_repo = src_file.relative_to(ROOT).as_posix()
    except ValueError:
        rel_in_repo = src_file.as_posix()
    folder_name = src_file.parent.name
    target = src_dir / folder_name / (module_name + src_file.suffix)
    target.parent.mkdir(parents=True, exist_ok=True)
    write_source(target, shell_text(rel_in_repo, src_file.suffix.lower()))
    return True


# ---------------------------------------------------------------------------
# align：同名真实代码 -> 占位壳
# ---------------------------------------------------------------------------
def _align_project(proj: dict, defaults: dict, dry: bool) -> tuple[int, int, int, list[str], list[str]]:
    """单个项目换壳；返回 (换壳数, 保留数, 已是壳数, 人话行, 警告行)。"""
    human: list[str] = []
    warnings: list[str] = []
    src_dir = abs_path(proj["src"])
    if not src_dir.is_dir():
        raise Failure(f"项目源码目录不存在：{src_dir}", code="SOURCE_PREFLIGHT",
                      stage=Stage.RESOLVE)

    common_dir = abs_path(defaults.get("common_dir", "src/common"))
    all_mods = module_index(common_dir)

    same_name: list[Path] = []
    already = 0
    for f in sorted(src_dir.rglob("*")):
        if not f.is_file() or f.suffix.lower() not in SOURCE_EXTS:
            continue
        if f.stem not in all_mods:
            continue
        if is_shell_file(f):
            already += 1
            continue
        same_name.append(f)

    if not same_name:
        human.append(f"  {proj['name']}：没有与公共库同名的真实代码"
                     f"（已经是壳的 {already} 个）")
        return 0, 0, already, human, warnings

    human.append(f"  {proj['name']}：发现 {len(same_name)} 个与公共库同名的真实代码"
                 + ("（只列清单，不改文件）" if dry else ""))

    swapped = kept = 0
    for f in same_name:
        common_file = all_mods[f.stem]
        rel = f.relative_to(src_dir).as_posix()
        c_rel = _rel(common_file)
        try:
            n_proj = len(read_source_text(f).splitlines())
            n_comm = len(read_source_text(common_file).splitlines())
        except Exception:
            n_proj = n_comm = 0
        t_proj = dt.datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
        t_comm = dt.datetime.fromtimestamp(common_file.stat().st_mtime).strftime("%Y-%m-%d %H:%M")

        human.append("")
        human.append(f"同名：{rel}")
        human.append(f"  项目    {n_proj} 行（改于 {t_proj}）")
        human.append(f"  公共库  {n_comm} 行（改于 {t_comm}）→ {c_rel}")
        if f.suffix.lower() != common_file.suffix.lower():
            warnings.append(f"注意：扩展名不同（项目 {f.suffix} / 公共库 "
                            f"{common_file.suffix}），写回时按模块名匹配，壳的后缀保持项目侧")

        if dry:
            human.append("  [dry-run] 会换成一行的占位壳（写回时改用公共库版本）")
            swapped += 1
            continue

        # 交互确认：--apply 由入口置 assume_yes；非交互默认不确认（保留）
        from toolkit.ui import ask_yes
        if not ask_yes("  把正文换成一行的占位壳（写回时改用公共库版本）？"):
            kept += 1
            continue

        dest = backup_copy(f, abs_path(defaults.get("backup_dir", ".backup")),
                           tag="-换壳", subdir=proj["name"], rel_to=src_dir)
        write_source(f, shell_text(c_rel, f.suffix.lower()))
        human.append(f"  [换壳] 已改为占位；原文备份：{_rel(dest)}")
        swapped += 1

    return swapped, kept, already, human, warnings


def align(ctx) -> Outcome:
    """对齐公共库：同名代码换壳。未给 --project 则处理所有项目。"""
    cfg = ctx.config
    name = getattr(ctx.args, "project", None)
    dry = not bool(getattr(ctx.args, "apply", False))

    human: list[str] = []
    warnings: list[str] = []
    rows: list[dict] = []
    total_swapped = total_kept = 0

    with ctx.progress.stage("align"):
        if name:
            targets = [_find_project(cfg, name)]
        else:
            targets = cfg["projects"]
        for proj in targets:
            swapped, kept, already, ph, pw = _align_project(proj, cfg["defaults"], dry)
            human.extend(ph)
            warnings.extend(pw)
            total_swapped += swapped
            total_kept += kept
            rows.append({"name": proj["name"], "swapped": swapped,
                         "kept": kept, "already": already})

    human.append("")
    if dry:
        human.append(f"共 {total_swapped} 个同名文件待换壳（dry-run 未改动任何文件）")
    else:
        human.append(f"共换壳 {total_swapped} 个，保留 {total_kept} 个（未新建任何文件）。")
        if total_swapped:
            human.append(f"下一步：写回一次，让表里真正用上公共库版本"
                         f"（菜单 2，或 {entry_command('sync.py', 'push', '--project', '<项目名>', '--apply')}）")

    return Outcome(
        action="sync.align",
        ok=True,
        details=rows,
        stats={"swapped": total_swapped, "kept": total_kept, "dry": bool(dry),
               "timings": ctx.stage_timings()},
        human=human,
        warnings=warnings,
        exit_code=0,
        elapsed_ms=ctx.elapsed_ms(),
    )


# ---------------------------------------------------------------------------
# promote：项目模块 -> 公共库
# ---------------------------------------------------------------------------
def promote(ctx) -> Outcome:
    """把某个项目的模块搬到 src/common，变成所有表共享。只搬文件，引用依赖需手工勾选。"""
    cfg = ctx.config
    project_name = ctx.args.project
    module_name = ctx.args.module
    dry = not bool(getattr(ctx.args, "apply", False))

    human: list[str] = []
    warnings: list[str] = []

    with ctx.progress.stage("promote"):
        proj = _find_project(cfg, project_name)
        src_dir = abs_path(proj["src"])
        if not src_dir.is_dir():
            raise Failure(f"源码目录不存在：{src_dir}", code="SOURCE_PREFLIGHT",
                          stage=Stage.RESOLVE)

        found = [f for f in src_dir.rglob("*")
                 if f.is_file() and f.suffix.lower() in SOURCE_EXTS
                 and module_name_of(f) == module_name]
        if not found:
            raise Failure(f"在 {src_dir} 里没找到模块「{module_name}」",
                          code="TARGET_NOT_FOUND", stage=Stage.RESOLVE,
                          suggested=entry_command("sync.py", "list", "--json"))
        if len(found) > 1:
            warnings.append(f"发现多个同名文件，只用第一个："
                            f"{found[0].relative_to(src_dir)}")

        src_file = found[0]
        common_dir = abs_path(cfg["defaults"].get("common_dir", "src/common"))
        target_dir = common_dir / _GROUP_BY_SUFFIX.get(src_file.suffix.lower(), "模块")
        target = target_dir / src_file.name

        if target.exists():
            raise Failure(f"公共库已存在同名文件：{_rel(target)}，请先在 VSCode 里合并再搬",
                          code="SOURCE_PREFLIGHT", stage=Stage.RESOLVE,
                          hint="先合并两边内容，再执行提升")

        human.append(f"源文件：{_rel(src_file)}")
        human.append(f"目标：  {_rel(target)}")

        # 提示影响面：其他项目有没有引用这个模块名
        others: list[str] = []
        for p in cfg["projects"]:
            if p["name"] == project_name:
                continue
            d = abs_path(p["src"])
            if not d.is_dir():
                continue
            for f in d.rglob("*"):
                if f.is_file() and f.suffix.lower() in SOURCE_EXTS:
                    if module_name in read_source_text(f):
                        others.append(p["name"])
                        break
        if others:
            human.append(f"同时被其他项目引用：{'、'.join(others)}"
                         "（搬走后它们将从公共库拿到这一份）")

        if dry:
            human.append("[dry] 不实际移动")
            stats = {"module": module_name, "project": project_name, "dry": True,
                     "from_path": str(src_file), "to_path": str(target),
                     "others": others, "shell_written": False, "refs": [],
                     "timings": ctx.stage_timings()}

            details = [{"name": module_name, "ok": True, "dry": True}]
            return Outcome(action="sync.promote", details=details, stats=stats,
                           human=human, warnings=warnings, exit_code=0,
                           elapsed_ms=ctx.elapsed_ms())

        move_into(src_file, target_dir)
        frx = src_file.with_suffix(".frx")
        if frx.exists():
            move_into(frx, target_dir)
            human.append(f"同时移动了 {frx.name}")

        human.append("")
        human.append(f"已移动到公共库：{_rel(target)}")
        need = _scan_required_refs([target_dir.parent])
        if need:
            human.append("")
            human.append("这个目录现在需要的 VBE 引用（每张表都要有，手工勾选）：")
            for desc, n in sorted(need.items()):
                human.append(f"  - {desc}（早期绑定 {n} 处）")

        # 刚搬走的模块自己不再出现在项目源码里，auto 会判定“没引用”，
        # 所以必须强制给它在原位置留一个占位壳
        written = _write_promoted_shell(proj, cfg["defaults"], module_name)
        human.append("")
        if written:
            human.append("原位置已生成占位文件（同名同后缀，只有一行注释指向公共库），"
                         "VSCode 里直接能看到")
        else:
            warnings.append("占位文件没能生成：项目目录里可能还留着同名真实代码，"
                             "请检查后手工处理")

        human.append("")
        human.append("下一步：")
        human.append("  1. 在 VSCode 里确认公共库文件的内容与中文正常")
        human.append("  2. git add -A && git commit  （Git 会识别为重命名）")
        human.append("  3. 每张要用的表：VBE → 工具 → 引用 → 勾上依赖 → 保存")
        human.append(f"  4. {entry_command('sync.py', 'push', '--all-project', '--apply')}"
                     "  或交互菜单 4，把公共库灌进各表")

        stats = {"module": module_name, "project": project_name, "dry": False,
                 "from_path": str(src_file), "to_path": _rel(target),
                 "others": others, "shell_written": bool(written),
                 "refs": sorted(need), "timings": ctx.stage_timings()}
        details = [{"name": module_name, "ok": True, "to": _rel(target)}]

    return Outcome(action="sync.promote", details=details, stats=stats,
                   human=human, warnings=warnings, exit_code=0,
                   elapsed_ms=ctx.elapsed_ms())


# ---------------------------------------------------------------------------
# gbk-export：src 镜像 -> GBK + CRLF 副本
# ---------------------------------------------------------------------------
def gbk_export(ctx) -> Outcome:
    """导出 src 的 GBK 副本到 .backup/GBK；不改 src，占位壳跳过，.frx 随 .frm 原样复制。"""
    cfg = ctx.config
    defaults = cfg.get("defaults", {})
    all_flag = bool(getattr(ctx.args, "all_project", False))
    dry = not bool(getattr(ctx.args, "apply", False))

    human: list[str] = []
    warnings: list[str] = []
    rows: list[dict] = []
    exported = shells = 0

    with ctx.progress.stage("gbk-export"):
        if all_flag:
            groups: list[tuple[Path, str]] = [
                (abs_path(defaults.get("common_dir", "src/common")), "_公共库")]
            proj_root = abs_path(defaults.get("projects_dir", "src/projects"))
            if proj_root.is_dir():
                groups += [(p, p.name) for p in sorted(proj_root.iterdir()) if p.is_dir()]
        else:
            name = ctx.args.project
            proj = _find_project(cfg, name)
            groups = [(abs_path(proj["src"]), name)]

        out_root = ctx.root / ".backup" / "GBK"

        for root, prefix in groups:
            dest_root = out_root / prefix
            group_exported = group_shells = 0
            # 这是“当前 src 的镜像”，不是历史备份：每次先清掉该组旧副本，
            # 避免源里已删除/改名的模块还留在 GBK 目录里被仓库外 VBE 误导入。
            if not dry:
                remove_tree(dest_root)
                if dest_root.exists():
                    warnings.append(f"旧副本未能完全清空，可能被占用：{dest_root}")
            if not root.is_dir():
                if dry and dest_root.is_dir():
                    human.append(f"    [dry] 将清空 {dest_root}")
                warnings.append(f"源码目录不存在，跳过：{root}")
                rows.append({"name": prefix, "exported": 0, "shells_skipped": 0,
                              "out_dir": str(dest_root), "cleaned": not dry,
                              "skipped": "源码目录不存在"})
                continue
            if dry and dest_root.is_dir():
                human.append("    [dry] 将清空 " + str(dest_root) + " 后重建")
            for f in sorted(root.rglob("*")):
                if not f.is_file() or f.suffix.lower() not in SOURCE_ALL_EXTS:
                    continue
                if f.suffix.lower() == ".frx":
                    continue  # 由 gbk_import_copy 随 .frm 原样带到同名目录
                if is_shell_file(f):
                    shells += 1
                    group_shells += 1
                    continue
                rel = Path(prefix) / f.relative_to(root)
                dest_dir = (out_root / rel).parent
                if dry:
                    human.append(f"    [dry] 导出 {rel.as_posix()}")
                    exported += 1
                    group_exported += 1
                    continue
                dest_dir.mkdir(parents=True, exist_ok=True)
                gbk_import_copy(f, dest_dir)
                exported += 1
                group_exported += 1
            rows.append({"name": prefix, "exported": group_exported,
                         "shells_skipped": group_shells, "out_dir": str(dest_root),
                         "cleaned": not dry})

    if dry:
        human.append("")
        human.append(f"[dry] 将导出 {exported} 个文件（跳过占位壳 {shells} 个）到 {out_root}")
    else:
        human.append(f"已导出 {exported} 个文件（GBK + CRLF，跳过占位壳 {shells} 个）")
        human.append(f"产物目录：{out_root}")

    return Outcome(
        action="sync.gbk-export",
        ok=True,
        details=rows,
        stats={"exported": exported, "shells_skipped": shells,
               "out_dir": str(out_root), "dry": bool(dry),
               "timings": ctx.stage_timings()},
        human=human,
        warnings=warnings,
        exit_code=0,
        elapsed_ms=ctx.elapsed_ms(),
    )


# ---------------------------------------------------------------------------
# clean-backup：清理 .backup 旧批次（默认只预览）
# ---------------------------------------------------------------------------
def _tag_label(tag: str) -> str:
    """JSON 里空 tag 统一叫 default，人话仍显示「默认」。"""
    return "(默认)" if tag in ("", "default") else tag


def clean_backup(ctx) -> Outcome:
    """按保留口径清理 .backup；--apply 才真删，默认只列删除清单。"""
    cfg = ctx.config
    apply = bool(getattr(ctx.args, "apply", False))
    keep = getattr(ctx.args, "keep", None)
    keep_tag = list(getattr(ctx.args, "keep_tag", None) or [])

    human: list[str] = []
    warnings: list[str] = []

    with ctx.progress.stage("clean-backup"):
        s = _backup_keep_settings(cfg)
        keep_requested = keep
        keep_note = ""
        if keep is None:
            keep_recent = s["keep_recent"]
        elif keep < 1:
            keep_recent = 1
            keep_note = f"--keep {keep} 无效：每类至少保留 1 批，已按 1 执行"
            warnings.append(keep_note)
        else:
            keep_recent = keep

        root = _backup_root(cfg)
        bk = backup_stats(root)
        human.append("")
        human.append(f"备份目录：{rel_path(root)}   共 {bk['count']} 批 / "
                     f"{bk['bytes'] / 1048576:.0f} MB")
        if bk["by_tag"]:
            parts = [f"{_tag_label(t)} {n} 批"
                     for t, n in sorted(bk["by_tag"].items(), key=lambda x: -x[1])]
            human.append("  按来源：" + " · ".join(parts))
        human.append(f"  保留口径：每类最近 {keep_recent} 批 + 每天最新 {s['keep_daily']} 份 + "
                     f"全局最新 {s['min_keep']} 批不动（GBK / watch 目录不参与）")

        result = _prune_backup_dir(cfg, dry_run=not apply, keep_recent=keep_recent,
                                    protect=keep_tag)
        ph, pw = _report_prune(result, tag="已删除" if apply else "将要删除", verbose=True)
        human.extend(ph)
        warnings.extend(pw)

        if result.get("kept_by_tag"):
            human.append("  保留来源：" + " · ".join(
                f"{_tag_label(t)} {n} 批"
                for t, n in sorted(result["kept_by_tag"].items(), key=lambda x: -x[1])))
        if not apply and result["deleted"]:
            human.append("")
            human.append(f"以上只是预览，文件没动。确认无误："
                         f"{entry_command('sync.py', 'clean-backup', '--apply')}")

        notes = Notes("本次提醒汇总")
        notes.add("校验", f"清理后剩 {result['kept']} 批 / "
                          f"{result['kept_bytes'] / 1048576:.0f} MB")
        human.extend(notes.render(clear=False))

        stats = {
            "applied": bool(apply),
            "freed_bytes": result["freed"],
            "deleted_batches": len(result["deleted"]),
            "kept": result["kept"],
            "kept_bytes": result["kept_bytes"],
            "keep_recent": keep_recent,
            "keep_requested": keep_requested,
            "keep_note": keep_note,
            "keep_daily": s["keep_daily"],
            "min_keep": s["min_keep"],
            "auto_clean": bool(s["auto"]),
            "keep_scope": "per-tag",
            "by_tag": {(t or "default"): n for t, n in bk["by_tag"].items()},
            "kept_by_tag": {(t or "default"): n for t, n in result.get("kept_by_tag", {}).items()},
            "kept_names": result.get("kept_names", []),
            "leftover": result.get("leftover", []),
            "protected": result.get("protected", []),
        }
        details = [{"batch": n, "bytes": b} for n, b in result["deleted"]]

    stats["timings"] = ctx.stage_timings()
    return Outcome(action="sync.clean-backup", details=details, stats=stats,
                   human=human, warnings=warnings, exit_code=0,
                   elapsed_ms=ctx.elapsed_ms())


SPECS = (
    CommandSpec(
        verb="align", entry="sync", handler="features.sync_transform:align",
        mutation=MutationMode.PREVIEW, project=ProjectMode.OPTIONAL,
        supports_json=True, exit_codes=(0, 1, 2),
        summary="把项目里与公共库同名的旧真身换成占位壳",
        ref="references/sync-commands.md#align",
        tests=("transform.align.success", "transform.align.not_found",
               "transform.align.usage", "transform.align.json"),
    ),
    CommandSpec(
        verb="promote", entry="sync", handler="features.sync_transform:promote",
        mutation=MutationMode.PREVIEW, project=ProjectMode.REQUIRED,
        args=(ArgSpec("--module", kind="str", required=True, help="要提升的模块名"),),
        supports_json=True, exit_codes=(0, 1, 2),
        summary="把一个模块从项目源码搬进公共库",
        ref="references/sync-commands.md#promote",
        tests=("transform.promote.success", "transform.promote.not_found",
               "transform.promote.usage", "transform.promote.json"),
    ),
    CommandSpec(
        verb="gbk-export", entry="sync", handler="features.sync_transform:gbk_export",
        mutation=MutationMode.PREVIEW, project=ProjectMode.ANY,
        supports_json=True, exit_codes=(0, 1, 2),
        summary="导出 src 的 GBK 副本到 .backup/GBK（不改 src）",
        ref="references/sync-commands.md#gbk-export",
        tests=("transform.gbk_export.success", "transform.gbk_export.not_found",
               "transform.gbk_export.usage", "transform.gbk_export.json"),
    ),
    CommandSpec(
        verb="clean-backup", entry="sync", handler="features.sync_transform:clean_backup",
        mutation=MutationMode.PREVIEW, project=ProjectMode.NONE,
        args=(
            ArgSpec("--keep", kind="int", help="每类保留最近 N 批（默认取配置 backup_keep）"),
            ArgSpec("--keep-tag", kind="append", help="这批/这类备份永不清理（匹配目录名或 tag，可重复）"),
        ),
        supports_json=True, exit_codes=(0, 1, 2),
        summary="清理 .backup 旧批次（默认只预览）",
        ref="references/sync-commands.md#clean-backup",
        tests=("transform.clean_backup.success", "transform.clean_backup.bad_config",
               "transform.clean_backup.usage", "transform.clean_backup.json"),
    ),
)
