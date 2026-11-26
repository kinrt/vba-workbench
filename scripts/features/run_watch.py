# -*- coding: utf-8 -*-
"""run 切片的 watch 监听辅助段：目标解析、探活、比对、单文件写回/抓回、基线刷新。

职责边界：
  · 本模块只承载 watch 主循环之外的内聚辅助；会话附着/开表只走
    ``adapters.excel_gateway``，组件读写走 toolkit 既有原语；
  · watch 主循环（含睡眠钩子 ``_sleep`` 与交互闸门 ``interactive_ok``）
    留在入口层 ``features.run``，避免测试既有 monkeypatch 断链；
  · 写回单文件/回读校验委托 ``run_macro`` 既有实现，不复制第二份。
"""
from __future__ import annotations

import datetime as dt
import tempfile
from pathlib import Path
from typing import Any

from adapters import excel_gateway
from toolkit import check, excel, files, repo, vba_parse, watch as watch_toolkit
from toolkit.base import (SOURCE_ALL_EXTS, Failure as BaseFailure, abs_path,
                          entry_command, read_source_text)
from toolkit.contract import ErrorCode, Failure, Stage
from features.common_subset import (  # noqa: E402
    find_project as _find_project,
)
from features.run_macro import (  # noqa: E402
    _common_subset,
    _import_module_file,
    _postcheck,
    _source_dirs,
)

WATCH_SAME = "一致"


def _watch_targets(cfg: dict, name: str | None) -> list[dict]:
    """要盯的表：给了项目名只盯那一个；没给则盯 Excel 里打开着的已登记项目。"""
    defaults = cfg.get("defaults", {})
    out: list[dict] = []

    def add(proj: dict, wb, xl) -> None:
        out.append({
            "proj": proj, "name": proj["name"], "book": str(wb.Name),
            "wb": wb, "xl": xl, "folders": _source_dirs(proj, defaults),
            "subset": _common_subset(proj, defaults)[0], "baseline": {}, "pending": {},
        })

    if name:
        proj = _find_project(cfg, name)
        path, reason = repo.resolve_target(proj)
        if path is None:
            raise Failure(f"「{proj['name']}」解析不到 xlsm：{reason}",
                          code=ErrorCode.TARGET_NOT_FOUND, stage=Stage.RESOLVE)
        xl = excel_gateway.attach_running_excel(
            path,
            target_hint="表没打开：watch 只能盯已打开的表",
            target_suggested=entry_command(
                "agent.py", "book", "--action", "open",
                "--project", proj["name"], "--apply"),
        )
        wb = excel.require_open_workbook(xl, path, proj["name"], "开启自动写回")
        add(proj, wb, xl)
        return out

    for proj in cfg.get("projects", []):
        path, _reason = repo.resolve_target(proj)
        if not path:
            continue
        try:
            xl = excel_gateway.attach_running_excel(path)
            wb, _opened = excel_gateway.find_open_workbook(xl, path)
        except BaseFailure:
            continue
        if wb is None:
            continue
        if not excel.ensure_vbe_access(wb, proj["name"], interactive=False):
            continue
        add(proj, wb, xl)

    if not out:
        raise Failure("Excel 里没有能监视的表（工程锁着/没授权访问的见上面的提示）。"
                      "先把要同步的表打开、在 VBE 里解开工程密码；只想盯一张就写 watch <表名>。",
                      code=ErrorCode.TARGET_NOT_OPEN, stage=Stage.ATTACH,
                      hint="watch 只能盯已打开的表；先用 book open 打开。")
    return out


def _watch_compare(wb, files: list[Path]) -> list[tuple[Path, str, int, int]]:
    """逐个比对 src 与工程正文：返回 [(文件, 结论, 源码行数, 工程行数)]。"""
    rows: list[tuple[Path, str, int, int]] = []
    for file in files:
        src = vba_parse.norm_code_lines(read_source_text(file))
        comp = excel.find_component(wb.VBProject, repo.module_name_of(file))
        vbe_text = ""
        if comp is not None:
            try:
                vbe_text = comp.CodeModule.Lines(1, max(comp.CodeModule.CountOfLines, 1))
            except Exception:
                comp = None
        if comp is None:
            rows.append((file, "工程里没有", len(src), 0))
            continue
        vbe = vba_parse.norm_code_lines(vbe_text)
        if src == vbe:
            rows.append((file, WATCH_SAME, len(src), len(vbe)))
        elif vba_parse.fold_lines(src) == vba_parse.fold_lines(vbe):
            rows.append((file, "仅大小写/空白", len(src), len(vbe)))
        else:
            rows.append((file, "内容不同", len(src), len(vbe)))
    return rows


def _compare_set(target: dict, common_dir: Path) -> list[Path]:
    """参与比对的文件：项目 src 全部 + 公共库子集。"""
    subset = target["subset"]
    folders = target["folders"]
    return [f for f in sorted(target["baseline"])
            if not (watch_toolkit.root_of(f, folders) == common_dir
                    and subset is not None and repo.module_name_of(f) not in subset)]


def _watch_write(wb, file: Path, folders: list[Path], common_dir: Path,
                 subset: set[str] | None, doc_modules: bool, save: bool,
                 dry: bool) -> tuple[bool, str]:
    """写回单个改动文件；返回 (是否已写回, 提示文字)。"""
    stem = repo.module_name_of(file)
    root = watch_toolkit.root_of(file, folders)
    if root == common_dir and subset is not None and stem not in subset:
        return False, f"跳过 {file.name}：本项目公共模块子集里没有它（共 {len(subset)} 个）"
    if repo.DOC_MODULE_RE.match(stem) and excel.find_component(wb.VBProject, stem) is None:
        return False, (f"跳过 {file.name}：工程里没有 {stem}（表页可能已被删除）；"
                       "跑一次「抓取源码」可清掉这个文件")

    before = excel.component_snapshot(wb.VBProject) if not dry else None
    tmp_dir: Path | None = None
    if not dry:
        tmp_dir = Path(tempfile.mkdtemp(prefix="vba-watch-"))
    try:
        status, _ = _import_module_file(
            wb.VBProject, file, doc_modules=doc_modules, dry=dry, tmp_dir=tmp_dir)
    finally:
        if tmp_dir is not None:
            files.remove_tree(tmp_dir)
    if status != "imported":
        return False, f"未写回 {file.name}：{status}"

    if not dry:
        check.check_module_set(before, excel.component_snapshot(wb.VBProject))
        bad = _postcheck(wb.VBProject, {stem: file})
        if bad:
            return False, f"回读校验失败 {file.name}：已放弃保存（{'、'.join(bad)}）"

    if save and not dry:
        try:
            wb.Save()
        except Exception as e:  # noqa: BLE001
            return True, f"写回 {file.name}，但保存失败：{e}"
        return True, f"写回 {file.name}（已保存）"
    return True, f"写回 {file.name}"


def _watch_export_comp(comp, out_dir: Path, dry: bool) -> Path:
    """把一个 VBE 组件 Export 到 out_dir 对应分类目录（模块私有，逻辑同 1.5.6）。

    清掉同名旧位置文件；VBE 产物 GBK 就地转 UTF-8 + LF。
    """
    name = comp.Name
    ext = repo.EXT_FOR_TYPE.get(comp.Type, ".bas")
    sub = repo.TYPE_FOLDER.get(comp.Type, "模块")
    stem = repo.doc_stem(name) if comp.Type == repo.VBXT_DOCUMENT else name
    target = out_dir / sub / (stem + ext)
    target.parent.mkdir(parents=True, exist_ok=True)

    for old in out_dir.rglob("*"):
        if not old.is_file() or repo.is_shell_file(old):
            continue
        if repo.module_name_of(old) == name and old.suffix.lower() in SOURCE_ALL_EXTS:
            if old.resolve() == target.resolve():
                continue
            if not dry:
                old.unlink()

    if not dry:
        comp.Export(str(target))
        files.transcode_exported_file(target)
    return target


def _watch_pull_one(wb, file: Path, folders: list[Path], dry: bool) -> Path | None:
    """以表内为准：把工程里这一个模块抓回 src；工程里没有返回 None。"""
    comp = excel.find_component(wb.VBProject, repo.module_name_of(file))
    if comp is None:
        return None
    return _watch_export_comp(comp, watch_toolkit.root_of(file, folders), dry)


def _watch_refresh_baseline(targets: list[dict], path: Path, old: Path | None = None) -> None:
    """脚本自己改过 src 后刷新所有表基线，避免下一轮被当成用户改动写回去。"""
    try:
        st = path.stat()
    except OSError:
        return
    for target in targets:
        target["baseline"][path] = (st.st_mtime_ns, st.st_size)
        if old is not None:
            target["baseline"].pop(old, None)


def _watch_log_path(ctx: Any, label: str) -> Path:
    defaults = ctx.config.get("defaults", {}) if isinstance(ctx.config, dict) else {}
    backup_dir = abs_path(defaults.get("backup_dir", ".backup"))
    return backup_dir / "watch" / f"{label}-{dt.datetime.now():%Y%m%d-%H%M%S}.log"
