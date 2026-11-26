# -*- coding: utf-8 -*-
"""
toolkit.watch —— 文件监听的三件套（快照 / 落点 / 两轮确认）
============================================================
纯 `stat`，不含 Excel、不含业务判断；任何"源码一改就干活"的工具都能用。

  · `snap(folders)`：{源码文件: (mtime_ns, 字节数)}，一轮只做一次 stat
  · `root_of(file, folders)`：这个文件属于哪个源码根（不在任何一个里就返回它的父目录）
  · `changed(snap, baseline, pending)`：挑出"相对基线变了，且连续两轮快照一致"的文件

为什么用 mtime+大小而不是哈希：改动判定够用，长驻监听每轮成本才可控；
为什么要求连续两轮一致：编辑器保存是分步写盘的，只看到一次变化就写回容易写进半截文件。
"""

from __future__ import annotations

from pathlib import Path

from .repo import source_files

Snap = dict[Path, tuple[int, int]]


def snap(folders: list[Path]) -> Snap:
    """监听快照 {源码文件: (mtime_ns, 字节数)}；文件清单走 repo.source_files（跳过占位壳）。"""
    out: Snap = {}
    for folder in folders:
        for f in source_files(folder):
            try:
                st = f.stat()
            except OSError:
                continue
            out[f] = (st.st_mtime_ns, st.st_size)
    return out


def root_of(file: Path, folders: list[Path]) -> Path:
    """文件属于哪个源码根（公共库 or 项目 src）；不在任何一个里就退回它的父目录。"""
    for root in folders:
        try:
            file.relative_to(root)
            return root
        except ValueError:
            continue
    return file.parent


def changed(current: Snap, baseline: Snap, pending: Snap) -> list[Path]:
    """
    挑出这一轮该处理的文件：相对基线有新变化，且**连续两轮**快照一致。

    pending 只记"上一轮看到的样子"（调用方持有，跨轮复用）；
    状态又变就继续等，避免写进编辑器写一半的文件。
    """
    ready: list[Path] = []
    for file, state in current.items():
        if baseline.get(file) == state:
            pending.pop(file, None)
            continue
        if pending.get(file) != state:
            pending[file] = state
            continue
        pending.pop(file, None)
        ready.append(file)
    return ready
