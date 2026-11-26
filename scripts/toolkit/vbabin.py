# -*- coding: utf-8 -*-
"""
toolkit.vbabin —— 只读解析 xlsm / xlsb 里的 vbaProject.bin，列模块清单、比对两份
============================================================================================
纯标准库（struct + zipfile），**不需要 Excel、不需要 VBE、不需要解开工程密码**。
用途：写回 / 抓取 / 换壳这类"动了 VBE"的操作之后，用它在文件层面核对
"模块有没有被删、有没有多出来"，而不是靠肉眼或靠 Excel 打开看。

【怎么用（脚本里）】
    from toolkit import compare_vba_modules, vba_module_names

    # 只要清单
    names = vba_module_names(r"C:\\表格\\示例表A V<数字>.xlsm")     # -> ["ThisWorkbook", "Sheet1", ...]

    # 比对写回前后两份文件（xlsm 直接给；也可以给导出的 vbaProject.bin）
    diff = compare_vba_modules(r".backup\\20260101-000000\\示例表A.xlsm", 写回后的表)
    diff["count_before"], diff["count_after"]   # 模块数
    diff["lost"], diff["added"]                 # 消失 / 新增的模块名（已排除编译缓存）
    diff["same"]                                # True 表示模块集合完全一致

【怎么用（命令行）】
    python -m toolkit.vbabin 示例表A.xlsm                  # 列这份文件里的模块
    python -m toolkit.vbabin 写回前.xlsm 写回后.xlsm        # 比对两份的模块集合

    日常更省事的是同步工具的菜单 11 / 在工具包根目录运行 `python scripts/sync.py vbamod --project 示例表A`：
    它会自己到 .backup 里取这张表最近一份备份来比，不用手填路径。

【它认得什么、不认得什么】
    · 认得：OLE 复合文档里的模块流。VBA 工程把每个模块存成一条**独立流**，
      流名就是模块名（中文也原样保留），所以清单是可靠的；
      内部流（dir / _VBA_PROJECT / PROJECT / PROJECTwm / PROJECTlk，
      以及带控制前缀的 \\x01CompObj / \\x03VBFrame，还有 f / o 子流）
      和编译缓存（__SRP_*）不参与模块清单。
    · 不认得：模块**内容**。比内容要比字符流（还没实现解压 MS-OVBA），
      所以本模块只回答"哪些模块在、哪些没了"，不回答"哪个模块的代码变了"。
    · 局限：理论上存在名为 `f` / `o` 的模块会被误判成内部流（VBA 里几乎不可能出现）。

【什么时候该用】
    · push / 批量写回 / watch 写完之后自查：模块数有没有变、有没有整块消失；
    · 排查"表里代码和源码不一致""写回把模块写没了"这类问题；
    · 写回前后的自动化回归：diff["same"] 为 False 就报警。
"""

from __future__ import annotations

import re
import struct
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from .base import Failure, msg

VBA_BIN_SUFFIX = "vbaproject.bin"
# 不是模块的流：工程内部流 + 编译缓存（每次重编译都会增删，不能当模块看）
INTERNAL_STREAMS = {"dir", "_VBA_PROJECT", "CompObj", "VBFrame", "f", "o"}
# 表页 xml 开头就有的 codeName（<sheetPr codeName="Sheet1"/>），只读前 4 KB 就够
CODE_NAME_RE = re.compile(r"""(?s)<sheetPr[^>]*?codeName\s*=\s*(?:"([^"]+)"|'([^']+)')""")
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


def _base_name(name: str) -> str:
    """去掉 OLE 流名的前导控制字符：`\\x01CompObj` → `CompObj`，`\\x03VBFrame` → `VBFrame`。"""
    return name.lstrip("".join(chr(i) for i in range(0x20)))


def _is_internal(name: str) -> bool:
    base = _base_name(name)
    # PROJECT / PROJECTwm / PROJECTlk（VBE 的锁定流）都算工程内部流
    return (base in INTERNAL_STREAMS or base.startswith("PROJECT")
            or base.startswith("__SRP") or base in ("", " "))


class _Ole:
    """最小 OLE 复合文档读取器：只做"按名字取流"，不写、不建。"""

    def __init__(self, data: bytes):
        if len(data) < 512 or data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            raise Failure("不是 OLE 复合文档：vbaProject.bin 文件头不对")
        self.data = data
        self.sector_size = 1 << struct.unpack_from("<H", data, 0x1E)[0]
        self.mini_size = 1 << struct.unpack_from("<H", data, 0x20)[0]
        self.nfat = struct.unpack_from("<I", data, 0x2C)[0]
        self.dir_start = struct.unpack_from("<I", data, 0x30)[0]
        self.mini_cutoff = struct.unpack_from("<I", data, 0x38)[0]
        self.minifat_start = struct.unpack_from("<I", data, 0x3C)[0]
        self.difat_start = struct.unpack_from("<I", data, 0x44)[0]
        self.ndifat = struct.unpack_from("<I", data, 0x48)[0]
        self.fat = self._read_fat()
        self.entries = self._read_dir()
        root = self.entries[0]
        self.roots = self._chain_bytes(root["start"], root["size"])
        self.minifat = self._read_minifat()

    # ---- 扇区寻址 ----
    def _sector(self, index: int) -> bytes:
        if self.sector_size == 512:
            start = 512 + index * 512          # 512 字节扇区：头本身占一个扇区
        else:
            start = self.sector_size + index * self.sector_size
        return self.data[start:start + self.sector_size]

    def _read_fat(self) -> list[int]:
        difat = list(struct.unpack_from("<109I", self.data, 0x4C))
        nxt, left = self.difat_start, self.ndifat
        while left > 0 and nxt < 0xFFFFFFFA:
            vals = struct.unpack("<%dI" % (self.sector_size // 4), self._sector(nxt))
            difat.extend(vals[:-1])
            nxt, left = vals[-1], left - 1
        fat: list[int] = []
        for index in difat[: self.nfat]:
            if index >= 0xFFFFFFFA:
                continue
            fat.extend(struct.unpack("<%dI" % (self.sector_size // 4), self._sector(index)))
        return fat

    def _chain(self, start: int) -> list[int]:
        out, cur, guard = [], start, 0
        while cur < 0xFFFFFFFA and guard < 200000:
            out.append(cur)
            cur = self.fat[cur] if cur < len(self.fat) else 0xFFFFFFFE
            guard += 1
        return out

    def _chain_bytes(self, start: int, size: int | None) -> bytes:
        if start >= 0xFFFFFFFA:
            return b""
        raw = b"".join(self._sector(s) for s in self._chain(start))
        return raw if size is None else raw[:size]

    def _read_minifat(self) -> list[int]:
        if self.minifat_start >= 0xFFFFFFFA:
            return []
        raw = self._chain_bytes(self.minifat_start, None)
        count = len(raw) // 4 * 4
        return list(struct.unpack("<%dI" % (count // 4), raw[:count]))

    def _read_dir(self) -> list[dict]:
        raw = self._chain_bytes(self.dir_start, None)
        out = []
        for off in range(0, len(raw) - 127, 128):
            e = raw[off:off + 128]
            name = e[:64].decode("utf-16-le", "ignore").rstrip("\x00")
            nlen = struct.unpack_from("<H", e, 0x40)[0]
            out.append({
                "name": name[: max(nlen - 2, 0)] if nlen else name,
                "type": e[0x42],
                "start": struct.unpack_from("<I", e, 0x74)[0],
                "size": struct.unpack_from("<I", e, 0x78)[0],
            })
        return out

    # ---- 对外：流名 / 取流 ----
    def streams(self) -> list[str]:
        return [e["name"] for e in self.entries if e["type"] == 2]

    def stream(self, name: str) -> bytes | None:
        for e in self.entries:
            if e["type"] != 2 or e["name"] != name:
                continue
            if e["size"] < self.mini_cutoff:
                out, index, guard = [], e["start"], 0
                while index < 0xFFFFFFFA and guard < 200000:
                    out.append(self.roots[index * self.mini_size:(index + 1) * self.mini_size])
                    index = self.minifat[index] if index < len(self.minifat) else 0xFFFFFFFE
                    guard += 1
                return b"".join(out)[: e["size"]]
            return self._chain_bytes(e["start"], e["size"])
        return None


def read_vba_bin(path: str | Path) -> bytes:
    """取 vbaProject.bin 的字节：给 xlsm/xlsb 就从 zip 里取，给 .bin 就直接读文件。"""
    p = Path(path)
    if not p.is_file():
        raise Failure(f"文件不存在：{p}")
    if p.suffix.lower() == ".bin":
        return p.read_bytes()
    try:
        with zipfile.ZipFile(p) as z:
            names = [n for n in z.namelist() if n.lower().endswith(VBA_BIN_SUFFIX)]
            if not names:
                raise Failure(f"这个文件里没有 VBA 工程（无 {VBA_BIN_SUFFIX}）：{p.name}")
            return z.read(names[0])
    except zipfile.BadZipFile:
        raise Failure(f"不是 xlsx/xlsm 压缩包，也不是 .bin：{p.name}")


def vba_models(data: bytes) -> tuple[list[str], list[str]]:
    """返回 (模块名清单, 编译缓存名清单)，都排好序。"""
    streams = _Ole(data).streams()
    modules = sorted({s for s in streams if not _is_internal(s)})
    cache = sorted({s for s in streams if s.startswith("__SRP")})
    return modules, cache


def vba_module_names(path: str | Path) -> list[str]:
    """只要模块名清单：vba_module_names("示例表A.xlsm") -> ["Sheet1", "表格工具", ...]"""
    return vba_models(read_vba_bin(path))[0]


def vba_documents(path: str | Path) -> list[str]:
    """
    文档型模块（表页/图表页/ThisWorkbook）清单，取自工程流 `PROJECT` 里的
    `Document=<名>/&H…` 行 —— 这是 VBE 自己登记的宿主关系，**不靠名字猜类型**，
    所以 `Sheet10_orphan` 这种"名字看着像表页、实际没宿主"的也能抓出来。
        vba_documents(表) -> ["ThisWorkbook", "Sheet1", "Sheet2", …]
    """
    data = read_vba_bin(path)
    proj = (_Ole(data).stream("PROJECT") or b"").decode("utf-8", "ignore")
    return sorted(set(re.findall(r"(?im)^Document=([^/\r\n]+)/", proj)))


def sheet_pages(path: str | Path) -> list[dict]:
    """
    表页清单（含 codeName），给"表页 ↔ 文档模块"对账用（H19）：
        sheet_pages("示例表A.xlsm") -> [{"name": "表页1", "code": "Sheet1", "chart": False}, ...]

    怎么读：`xl/workbook.xml` 的 `<sheets>` 拿表页名与 r:id，`xl/_rels/workbook.xml.rels`
    把 r:id 映到 `worksheets/sheetN.xml`，再从表页 xml 开头（前 4 KB）取
    `<sheetPr codeName="SheetN"/>` —— 这个 codeName 就是 VBA 里的文档模块名。
    纯 zip + 正则，**不开 Excel、不需要工程密码**；`.xlsb` 没有 xml 表页结构，直接抛 Failure。
    """
    p = Path(path)
    if not p.is_file():
        raise Failure(f"文件不存在：{p}")
    if p.suffix.lower() not in (".xlsm", ".xlsx", ".xltm", ".xltx"):
        raise Failure(f"只有 xlsx/xlsm 才有 xml 表页结构，{p.suffix} 不支持：{p.name}")
    try:
        with zipfile.ZipFile(p) as z:
            names = set(z.namelist())
            if "xl/workbook.xml" not in names:
                raise Failure(f"这个文件里没有 xl/workbook.xml：{p.name}")
            rel_map: dict[str, str] = {}
            if "xl/_rels/workbook.xml.rels" in names:
                rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
                for rel in rels:
                    rel_map[rel.get("Id") or ""] = (rel.get("Target") or "").lstrip("/")
            wb = ET.fromstring(z.read("xl/workbook.xml"))
            pages: list[dict] = []
            for sh in wb.findall(".//m:sheets/m:sheet", NS):
                target = rel_map.get(sh.get(f"{{{NS['r']}}}id") or "", "")
                full = target if target.startswith("xl/") else f"xl/{target}"
                code = None
                if full in names:
                    # 快路径：codeName 在表页开头；正则没命中再整份解析（表页大体量数据不上常路）
                    head = z.read(full)[:4096].decode("utf-8", "ignore")
                    m = CODE_NAME_RE.search(head)
                    if m:
                        code = m.group(1) or m.group(2)
                    else:
                        pr = ET.fromstring(z.read(full)).find("m:sheetPr", NS)
                        code = pr.get("codeName") if pr is not None else None
                pages.append({"name": sh.get("name") or "", "code": code,
                              "chart": "chartsheet" in full})
            return pages
    except zipfile.BadZipFile:
        raise Failure(f"不是 xlsx/xlsm 压缩包：{p.name}")


def compare_vba_modules(before: str | Path, after: str | Path) -> dict:
    """
    比对两份文件（xlsm 或 vbaProject.bin）的模块集合。返回:
        count_before / count_after : 模块数
        lost / added               : 只在前者 / 只在后者 的模块名（已排除 __SRP 缓存）
        cache_before / cache_after : 编译缓存数（每次重编译都会变，不算异常）
        same                       : 模块集合是否完全一致
    """
    mo_before, cache_before = vba_models(read_vba_bin(before))
    mo_after, cache_after = vba_models(read_vba_bin(after))
    lost = sorted(set(mo_before) - set(mo_after))
    added = sorted(set(mo_after) - set(mo_before))
    return {
        "count_before": len(mo_before),
        "count_after": len(mo_after),
        "lost": lost,
        "added": added,
        "cache_before": len(cache_before),
        "cache_after": len(cache_after),
        "same": not lost and not added,
    }


def _print_report(diff: dict, before: str, after: str) -> None:
    msg(f"写回前：{before}")
    msg(f"写回后：{after}")
    msg(f"模块数：{diff['count_before']} → {diff['count_after']}")
    msg(f"编译缓存（__SRP，会随重编译增减，不算异常）："
        f"{diff['cache_before']} → {diff['cache_after']}")
    if diff["same"]:
        msg("模块集合：完全一致")
        return
    if diff["lost"]:
        msg(f"[消失] {len(diff['lost'])} 个：{'、'.join(diff['lost'])}")
    if diff["added"]:
        msg(f"[新增] {len(diff['added'])} 个：{'、'.join(diff['added'])}")


def main(argv: list[str] | None = None) -> int:
    import sys

    args = list(sys.argv[1:] if argv is None else argv)
    if len(args) == 1:
        names = vba_module_names(args[0])
        msg(f"{len(names)} 个模块：")
        for n in names:
            msg(f"  {n}")
        return 0
    if len(args) == 2:
        _print_report(compare_vba_modules(args[0], args[1]), args[0], args[1])
        return 0
    msg("用法：")
    msg("  python -m toolkit.vbabin 表.xlsm            列模块清单")
    msg("  python -m toolkit.vbabin 前.xlsm 后.xlsm   比对模块集合")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
