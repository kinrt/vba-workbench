# -*- coding: utf-8 -*-
"""发布打包：按安装器交付白名单生成 zip，并更新 SHA256SUMS / latest.txt。

默认只预览；带 --apply 才写 releases/。包内根目录固定为 vba-workbench/。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import tempfile
import zipfile
from pathlib import Path

import upgrade_install as ui


MAX_FILE_BYTES = 2 * 1024 * 1024
MAX_TOTAL_BYTES = 16 * 1024 * 1024
MAX_FILES = 1024


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _zip_names(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as zf:
        return zf.namelist()


def main() -> int:
    parser = argparse.ArgumentParser(description="vba-workbench 发布打包")
    parser.add_argument("--source", default="", help="源码仓库根")
    parser.add_argument("--releases", default="",
                        help="releases 目录，默认 <源码父目录>/releases")
    parser.add_argument("--apply", action="store_true", help="真正写 zip / SHA / latest")
    args = parser.parse_args()

    source = Path(args.source).resolve() if args.source else Path(__file__).resolve().parents[2]
    releases = (Path(args.releases).resolve() if args.releases
                else source.parent / "releases")
    version = (source / "VERSION").read_text(encoding="utf-8").strip()
    name = f"vba-workbench-{version}.zip"
    target = releases / name

    files = ui._iter_source_files(source)
    if not files:
        raise SystemExit("没有收集到任何交付文件")
    total = sum(p.stat().st_size for p in files)
    oversize = [str(p) for p in files if p.stat().st_size > MAX_FILE_BYTES]
    if len(files) > MAX_FILES:
        raise SystemExit(f"交付文件 {len(files)} 个，超过上限 {MAX_FILES}")
    if total > MAX_TOTAL_BYTES:
        raise SystemExit(f"交付总大小 {total} 字节，超过上限 {MAX_TOTAL_BYTES}")
    if oversize:
        raise SystemExit(f"单文件超过 2 MB：{oversize}")
    bad_suffix = [str(p) for p in files if p.suffix.lower() in (".xlsm", ".xlsb", ".xlsx", ".xls")]
    if bad_suffix:
        raise SystemExit(f"交付件里混入工作簿：{bad_suffix}")

    payload = {
        "ok": True, "version": version, "zip": str(target), "files": len(files),
        "bytes": total, "apply": bool(args.apply),
    }
    if not args.apply:
        print(json.dumps(payload, ensure_ascii=False))
        return 0

    releases.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix=f".{name}.", suffix=".tmp", dir=str(releases))
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in files:
                rel = path.relative_to(source).as_posix()
                zf.write(path, f"vba-workbench/{rel}")
        names = _zip_names(tmp)
        if len(names) != len(files):
            raise SystemExit(f"zip 条目数 {len(names)} 与交付文件 {len(files)} 不一致")
        if any(not n.startswith("vba-workbench/") for n in names):
            raise SystemExit("zip 根目录不是 vba-workbench/")
        if any("/__pycache__/" in n or n.endswith(".pyc") for n in names):
            raise SystemExit("zip 里混入 __pycache__/*.pyc")
        os.replace(tmp, target)
    finally:
        if tmp.exists():
            tmp.unlink()

    digest = _sha256(target)
    sums = releases / "SHA256SUMS.txt"
    lines = []
    if sums.is_file():
        lines = [ln for ln in sums.read_text(encoding="utf-8").splitlines()
                 if ln.strip() and name not in ln]
    lines.append(f"{digest}  {name}")
    sums.write_text("\n".join(lines) + "\n", encoding="utf-8")
    (releases / "latest.txt").write_text(name + "\n", encoding="utf-8")

    payload.update({"sha256": digest, "entries": len(names)})
    print(json.dumps(payload, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
