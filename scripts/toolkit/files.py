# -*- coding: utf-8 -*-
"""
toolkit.files —— 文件底层操作（写入 / 导出 / 备份 / 移动）
============================================================
编码约定在这里落地一次，功能文件不再各写各的：

  · write_source     ：VBA 源码一律 UTF-8（无 BOM）+ LF（src 的存储格式）
  · transcode_exported_file：VBE Export 产物（GBK）就地转成上面的 UTF-8 + LF 约定
  · gbk_import_copy  ：UTF-8 源码转成 GBK+CRLF 临时文件，供 VBE Import（边界转码；
                        VBE 按 ANSI 解码、按 CRLF 断行，src 的 LF 必须在这一步补成 CRLF）
  · write_text_utf8  ：文本产物一律 UTF-8 无 BOM + LF，相对路径按仓库根解析
  · dump_json        ：JSON 产物（默认缩进 1）
  · stamp_dir        ：时间戳备份子目录 <backup_root>/<YYYYmmdd-HHMMSS><tag>
  · backup_copy      ：复制进时间戳备份目录（copy2 保留元数据）
  · move_into        ：移动文件到目标目录（目录不存在就建）
  · append_line      ：追加一行日志（UTF-8 无 BOM + LF）
  · newest_file      ：目录里最近修改的那份文件（可选"只要某时刻之后的"）
  · read_tail        ：读文件末尾 N 行（utf-8 → gbk 回退，换行统一 LF）
  · remove_tree      ：递归删临时目录（不抛异常）
  · backup_dirs      ：扫时间戳备份批次（跳过 GBK / watch 这类固定目录）
  · backup_stats     ：备份体量概览（批次数 / 字节 / 每类 tag 批次数）
  · prune_backups    ：备份自动清理（每类留最近 N 批 + 每天 1 份 + 最新 M 批不动）

这些原语都**不打印**；提示语由调用方自己发。
"""

from __future__ import annotations

import datetime as dt
import json
import re
import shutil
from pathlib import Path

from .base import Failure, abs_path, read_source_text


def write_source(path: Path, content: str) -> None:
    """VBA 源码统一 UTF-8（无 BOM）+ LF 写入（内容末尾没有换行则自动补一个）。

    传入内容里的 CRLF / 裸 CR 一律压成 LF：仓库存储只认 LF（agent 与编辑器写文件
    天然产 LF）。VBE 需要 CRLF 的场合走 gbk_import_copy，在写回边界补，不回流 src。
    """
    if not content.endswith(("\n", "\r")):
        content += "\n"
    data = content.replace("\r\n", "\n").replace("\r", "\n").encode("utf-8")
    path.write_bytes(data)


def transcode_exported_file(path: Path) -> None:
    """把 VBE Export 产出的 .bas/.cls/.frm（ANSI/GBK）就地转成本仓库的 UTF-8 + LF 约定
    （VBE 的 CRLF 只在写回时由 gbk_import_copy 补，不留在 src）。"""
    raw = path.read_bytes()
    try:
        text = raw.decode("gbk")
    except UnicodeDecodeError:
        text = read_source_text(path)
    write_source(path, text)


def gbk_import_copy(src_file: Path, tmp_dir: Path) -> Path:
    """
    把 src 里的 UTF-8 源码复制成 tmp_dir 下同名的 GBK + CRLF 文件，供 VBE Import；
    **边界必须是 CRLF**：VBE 遇 LF 会把整段连成一行，这里的换行转换不能省；
    .frm 有同名 .frx 时一并原样复制过去（VBE Import 按同目录同名找 .frx）。
    内容含 GBK 编不出的字符时抛 Failure——宁可停下也不写乱码文件进表。
    """
    text = read_source_text(src_file)
    if not text.endswith(("\n", "\r")):
        text += "\n"
    try:
        data = text.replace("\r\n", "\n").replace("\n", "\r\n").encode("gbk")
    except UnicodeEncodeError as e:
        raise Failure(f"{src_file.name} 含 GBK 无法编码的字符（{e.reason}），无法写回 VBE") from e
    tmp = tmp_dir / src_file.name
    tmp.write_bytes(data)
    if src_file.suffix.lower() == ".frm":
        frx = src_file.with_suffix(".frx")
        if frx.exists():
            shutil.copy2(frx, tmp_dir / frx.name)
    return tmp


def write_text_utf8(path: str | Path, text: str) -> Path:
    """文本产物统一 UTF-8 无 BOM + LF；相对路径按仓库根解析，父目录自动创建。"""
    p = abs_path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(text, encoding="utf-8", newline="\n")
    return p


def dump_json(path: str | Path, data, indent: int = 1) -> Path:
    """JSON 产物：UTF-8 无 BOM + LF，ensure_ascii=False。"""
    return write_text_utf8(path, json.dumps(data, ensure_ascii=False, indent=indent))


def stamp_dir(backup_root: Path, tag: str = "") -> Path:
    """时间戳备份子目录（顺带创建）；tag 形如 "-换壳"、"-抓取前"。"""
    d = backup_root / f"{dt.datetime.now().strftime('%Y%m%d-%H%M%S')}{tag}"
    d.mkdir(parents=True, exist_ok=True)
    return d


def backup_copy(src: Path, backup_root: Path, *, tag: str = "",
                subdir: str | None = None, rel_to: Path | None = None) -> Path:
    """
    把 src 复制进时间戳备份目录（copy2 保留修改时间），返回备份文件路径。
      subdir：备份目录里再分一层（如项目名）
      rel_to：按该目录解析相对路径，保留子目录结构；不给就只留文件名
    """
    dest_dir = stamp_dir(backup_root, tag)
    if subdir:
        dest_dir = dest_dir / subdir
    rel = src.relative_to(rel_to) if rel_to is not None else src.name
    dest = dest_dir / rel
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dest)
    return dest


def latest_backup(backup_root: Path, name: str) -> Path | None:
    """
    在备份目录里找某个文件最近的一份副本。

    排序按**备份目录名**（`20261117-231633` 这种时间戳，字典序就是时间序），
    因为 backup_copy 用 copy2 保留了源文件 mtime，按文件时间排会选错。
    """
    root = Path(backup_root)
    if not root.is_dir():
        return None
    found = [p for p in root.rglob(name) if p.is_file()]
    stamped = [p for p in found if p.parent.name[:8].isdigit()]
    if stamped:
        stamped.sort(key=lambda p: p.parent.name)
        return stamped[-1]
    found.sort(key=lambda p: p.stat().st_mtime)
    return found[-1] if found else None


def move_into(src: Path, dest_dir: Path) -> Path:
    """把 src 移进 dest_dir（目录不存在就建），返回目标路径。"""
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / src.name
    shutil.move(str(src), str(dest))
    return dest


def append_line(path: Path, text: str) -> None:
    """往文本文件尾部追加一行（UTF-8 无 BOM + LF；父目录自动建）。

    日志类文件（.backup/watch/*.log）都走这里：编码与换行只在 toolkit 里定一次，
    入口脚本不再自己 open("a", encoding=..., newline=...)。
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8", newline="\n") as f:
        f.write(text + "\n")


def remove_tree(path: Path) -> None:
    """递归删除一个临时目录；不存在或删不干净都不抛（临时目录清理用）。"""
    shutil.rmtree(path, ignore_errors=True)


def dir_size_bytes(path: Path) -> int:
    """目录里所有文件的字节数之和（目录不存在算 0，不抛异常）。"""
    if not path.is_dir():
        return 0
    total = 0
    for f in path.rglob("*"):
        try:
            if f.is_file():
                total += f.stat().st_size
        except OSError:
            continue
    return total


# 备份目录名：<YYYYmmdd-HHMMSS><tag>，tag 形如 ""（默认）/ "-抓取前" / "-自动写回" / "_单模块"
_BACKUP_STAMP_RE = re.compile(r"^(\d{8}-\d{6})(.*)$")


def backup_dirs(backup_root: Path) -> list[tuple[Path, dt.datetime, str]]:
    """
    扫 <backup_root> 下所有时间戳备份目录，按时间**升序**返回 (目录, 时间, tag)。

    只认名字能以 `YYYYmmdd-HHMMSS` 开头的目录：`.backup/GBK`、`.backup/watch`
    这类固定目录不是备份批，一律跳过（它们由别的机制管理）。
    """
    root = Path(backup_root)
    if not root.is_dir():
        return []
    found = []
    for d in root.iterdir():
        if not d.is_dir():
            continue
        m = _BACKUP_STAMP_RE.match(d.name)
        if not m:
            continue
        try:
            ts = dt.datetime.strptime(m.group(1), "%Y%m%d-%H%M%S")
        except ValueError:
            continue
        found.append((d, ts, m.group(2)))
    found.sort(key=lambda x: x[1])
    return found


def backup_stats(backup_root: Path) -> dict:
    """
    备份体量概览（体检 H10 用）：批次数、总字节、每类 tag 的批次数。

    返回 {"count": 批次数, "bytes": 总字节, "by_tag": {tag: 批次数}}；
    目录不存在返回 count=0。GBK / watch 不算备份批，不计入。
    """
    items = backup_dirs(backup_root)
    by_tag: dict[str, int] = {}
    total = 0
    for d, _ts, tag in items:
        by_tag[tag or "default"] = by_tag.get(tag or "default", 0) + 1
        total += dir_size_bytes(d)
    return {"count": len(items), "bytes": total, "by_tag": by_tag}


def prune_backups(backup_root: Path, *, keep_recent: int = 10, keep_daily: int = 1,
                  min_keep: int = 3, dry_run: bool = True,
                  protect: list[str] | tuple[str, ...] | None = None) -> dict:
    """
    备份自动清理：按"每类保留最近 N 批 + 每天留 1 份 + 全局最新 M 批不动"删旧备份。

    规则（三条取并集，命中任一条就保留）：
      · 每个 tag（""、"-抓取前"、"-自动写回"…）各自保留最近 keep_recent 批
      · 每个 tag 每天保留最新的 keep_daily 份 —— 老日子也留得下一天一份，能往回翻
      · 全局（不分 tag）最新 min_keep 批永不删 —— `latest_backup` 核对模块清单要用它
      · protect 里的 token 命中批次目录名或 tag（如 "-覆盖前"）时永不删，
        用来保住人工标记的回滚点
    安全边界：
      · 只处理名字是 `<YYYYmmdd-HHMMSS><tag>` 的目录，`GBK`/`watch`/其它目录一律不碰
      · dry_run=True 只算不删（默认），调用方要真删得显式传 False
      · 删除用 remove_tree（不抛异常），单个目录删不掉不影响其它批

    返回 {"scanned", "kept", "kept_bytes", "deleted": [(名字, 字节)], "freed",
    "protected": [名字], "leftover"}。
    `leftover` 是真删之后仍留在磁盘上的批次名：`remove_tree` 用 ignore_errors，
    同步盘（如 OneDrive、坚果云）临时占用时会删不掉，调用方该提示"下次再清"。
    """
    root = Path(backup_root)
    items = backup_dirs(root)
    if not items:
        return {"scanned": 0, "kept": 0, "kept_bytes": 0, "deleted": [], "freed": 0,
                "protected": [], "leftover": []}

    tokens = [str(t).strip() for t in (protect or []) if str(t).strip()]
    keep: set[Path] = {d for d, _ts, _tag in items[-max(min_keep, 1):]}
    protected: list[str] = []
    for d, _ts, tag in items:
        if any(tok in d.name or tok == tag for tok in tokens):
            keep.add(d)
            protected.append(d.name)
    by_tag: dict[str, list[tuple[Path, dt.datetime]]] = {}
    for d, ts, tag in items:
        by_tag.setdefault(tag, []).append((d, ts))

    for group in by_tag.values():
        group.sort(key=lambda x: x[1], reverse=True)      # 新 → 旧
        per_day: dict[dt.date, int] = {}
        for i, (d, ts) in enumerate(group):
            hit = i < keep_recent
            if keep_daily > 0:
                day = ts.date()
                if per_day.get(day, 0) < keep_daily:
                    hit = True
                per_day[day] = per_day.get(day, 0) + 1
            if hit:
                keep.add(d)

    deleted: list[tuple[str, int]] = []
    kept_names: list[str] = []
    kept_by_tag: dict[str, int] = {}
    kept_bytes = 0
    for d, _ts, tag in items:
        if d in keep:
            kept_names.append(d.name)
            kept_by_tag[tag or "default"] = kept_by_tag.get(tag or "default", 0) + 1
            kept_bytes += dir_size_bytes(d)
            continue
        size = dir_size_bytes(d)
        deleted.append((d.name, size))
        if not dry_run:
            remove_tree(d)

    return {
        "scanned": len(items),
        "kept": len(items) - len(deleted),
        "kept_bytes": kept_bytes,
        "kept_names": kept_names,
        "kept_by_tag": kept_by_tag,
        "deleted": deleted,
        "freed": sum(s for _n, s in deleted),
        "protected": protected,
        "leftover": [] if dry_run else [n for n, _s in deleted if (root / n).exists()],
    }


def newest_file(folder: Path, pattern: str = "*", *,
                since: float | None = None) -> Path | None:
    """
    目录里最近修改的那份文件；since 给了就只在「修改时间 >= since」里找。
    典型用法：刚跑完一个宏，去它写日志的目录里取这次新落的那一份——
      newest_file(表目录 / "测试输出", "*.log", since=开始时间)
    目录不在、或没有满足条件的文件，都返回 None（调用方自己提示）。
    """
    if not folder.is_dir():
        return None
    files = [p for p in folder.glob(pattern) if p.is_file()]
    if since is not None:
        files = [p for p in files if p.stat().st_mtime >= since]
    if not files:
        return None
    files.sort(key=lambda p: p.stat().st_mtime)
    return files[-1]


def read_tail(path: Path, lines: int = 40,
              encodings: tuple[str, ...] = ("utf-8", "gbk")) -> str:
    """
    读文本文件末尾 lines 行，换行统一成 LF。
    VBA 的 Print # 落的是 ANSI(GBK)，所以先试 utf-8、再退 gbk；
    文件不存在、行数为 0、编码都不认时返回空串（不抛，调用方决定是否提示）。
    """
    if lines <= 0 or not path.is_file():
        return ""
    text = ""
    for enc in encodings:
        try:
            text = path.read_text(encoding=enc)
            break
        except (UnicodeDecodeError, LookupError):
            continue
    if not text:
        return ""
    norm = text.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")
    return "\n".join(norm.split("\n")[-lines:])
