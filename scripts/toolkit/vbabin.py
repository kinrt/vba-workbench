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
    · 认得：模块正文。模块流 = MODULEOFFSET 字节的 PerformanceCache + 一段
      MS-OVBA 压缩的 CompressedSourceCode；本模块按 MS-OVBA 2.4.1.3 解压后，
      统一换行并剔除 Attribute 元数据行再比较，所以能回答"哪个模块的代码变了"，
      不只回答"模块在不在"。
    · 局限：解压失败时只报"内容未比对"，不会把未比对当成"内容一致"；
      理论上存在名为 `f` / `o` 的模块会被误判成内部流（VBA 里几乎不可能出现）。

【什么时候该用】
    · push / 批量写回 / watch 写完之后自查：模块数有没有变、内容有没有变；
    · 排查"表里代码和源码不一致""写回把模块写没了"这类问题；
    · 写回前后的自动化回归：diff["same"] / diff["content_same"] 任一为 False 就报警。
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


def decompress_vba_stream(data: bytes) -> bytes:
    """按 MS-OVBA 2.4.1.3 解压 CompressedContainer（模块正文 / dir 流）。

    调用方通常从模块流的第一个 0x01 压缩容器开始传入；本函数不负责找
    PerformanceCache 与 CompressedSourceCode 的分界。
    """
    if not data:
        return b""
    if data[0] != 0x01:
        raise Failure("VBA 压缩流签名不是 0x01，无法解压")
    out = bytearray()
    pos = 1
    while pos < len(data):
        if pos + 2 > len(data):
            raise Failure("VBA 压缩流在 chunk 头处截断")
        chunk_start_byte = pos
        header = struct.unpack_from("<H", data, pos)[0]
        pos += 2
        chunk_size = (header & 0x0FFF) + 3
        signature = (header >> 12) & 0x07
        if signature != 0b011:
            # 末尾全 0 填充：不是坏流，直接结束
            if all(b == 0 for b in data[chunk_start_byte:]):
                break
            raise Failure("VBA 压缩流 chunk 签名不是 0b011")
        compressed = bool((header >> 15) & 0x01)
        chunk_end = min(len(data), chunk_start_byte + chunk_size)
        chunk_start = len(out)
        if not compressed:
            out.extend(data[pos:chunk_end])
            pos = chunk_end
            continue
        while pos < chunk_end:
            flag = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if (flag >> bit) & 1 == 0:
                    out.append(data[pos])
                    pos += 1
                    continue
                if pos + 2 > chunk_end:
                    raise Failure("VBA 压缩流 CopyToken 截断")
                token = struct.unpack_from("<H", data, pos)[0]
                pos += 2
                length, offset = unpack_copy_token(token, len(out), chunk_start)
                copy_source = len(out) - offset
                if copy_source < chunk_start:
                    raise Failure("VBA 压缩流 CopyToken 越界")
                # 允许重叠复制（RLE），必须逐字节 append
                for index in range(length):
                    out.append(out[copy_source + index])
        pos = chunk_end
    return bytes(out)


def unpack_copy_token(token: int, decompressed_current: int,
                      decompressed_chunk_start: int) -> tuple[int, int]:
    """MS-OVBA 2.4.1.3.19.2：返回 (length, offset)，offset 从 1 开始。"""
    difference = decompressed_current - decompressed_chunk_start
    bit_count = max(4, (difference - 1).bit_length() if difference > 1 else 0)
    length_mask = 0xFFFF >> bit_count
    offset_mask = (~length_mask) & 0xFFFF
    length = (token & length_mask) + 3
    offset = ((token & offset_mask) >> (16 - bit_count)) + 1
    return length, offset


def find_compressed_start(stream: bytes) -> tuple[int, bytes]:
    """在模块流里找 CompressedSourceCode 起点：PerformanceCache 之后第一个
    解压后以 `Attribute VB_Name` 开头的合法容器；找不到再退回首个可解压候选。
    """
    if not stream:
        return 0, b""
    fallback: tuple[int, bytes] | None = None
    for start in range(0, len(stream) - 2):
        if stream[start] != 0x01:
            continue
        header = struct.unpack_from("<H", stream, start + 1)[0]
        if (header >> 12) & 0x07 != 0b011:
            continue
        chunk_size = (header & 0x0FFF) + 3
        if not 3 <= chunk_size <= 4098:
            continue
        try:
            source = decompress_vba_stream(stream[start:])
        except (Failure, struct.error):
            continue
        if source.startswith(b"Attribute VB_Name"):
            return start, source
        if fallback is None:
            fallback = (start, source)
    if fallback is not None:
        return fallback
    raise Failure("模块流里找不到合法的 MS-OVBA 压缩容器")


def normalize_vba_source(data: bytes) -> bytes:
    """比较用归一：统一 CRLF / CR 为 LF，去掉 Attribute 元数据行与首尾 NUL。"""
    text = data.replace(b"\r\n", b"\n").replace(b"\r", b"\n")
    lines = [line for line in text.split(b"\n")
             if not line.lstrip(b"\xef\xbb\xbf").startswith(b"Attribute ")]
    return b"\n".join(lines).strip(b"\x00")


def module_sources_from_bin(data: bytes) -> dict[str, bytes]:
    """从 vbaProject.bin 取每个模块的正文（已解压、已归一换行）。"""
    ole = _Ole(data)
    modules, _ = vba_models(data)
    sources: dict[str, bytes] = {}
    for name in modules:
        stream = ole.stream(name)
        if not stream:
            sources[name] = b""
            continue
        _offset, source = find_compressed_start(stream)
        sources[name] = normalize_vba_source(source)
    return sources


def vba_module_sources(path: str | Path) -> dict[str, bytes]:
    """公开入口：xlsm / xlsb / vbaProject.bin -> {模块名: 解压后的正文 bytes}。"""
    return module_sources_from_bin(read_vba_bin(path))


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
    比对两份文件（xlsm 或 vbaProject.bin）的模块集合与模块正文。返回:
        count_before / count_after : 模块数
        lost / added               : 只在前者 / 只在后者 的模块名（已排除 __SRP 缓存）
        changed                    : 两边都有、但正文不同的模块名
        cache_before / cache_after : 编译缓存数（每次重编译都会变，不算异常）
        same                       : 模块集合是否完全一致（保持旧语义）
        content_same               : 模块集合与正文是否都一致；解压失败时为 None
        content_compared           : 正文是否完成比对
        content_error              : 正文未比对时的原因
    """
    before_data = read_vba_bin(before)
    after_data = read_vba_bin(after)
    mo_before, cache_before = vba_models(before_data)
    mo_after, cache_after = vba_models(after_data)
    lost = sorted(set(mo_before) - set(mo_after))
    added = sorted(set(mo_after) - set(mo_before))
    changed: list[str] = []
    content_same: bool | None = True
    content_error = ""
    try:
        sources_before = module_sources_from_bin(before_data)
        sources_after = module_sources_from_bin(after_data)
        for name in sorted(set(mo_before) & set(mo_after)):
            if sources_before.get(name, b"") != sources_after.get(name, b""):
                changed.append(name)
        if lost or added:
            content_same = False
        else:
            content_same = not changed
    except Failure as e:
        content_same = None
        content_error = str(e)
    return {
        "count_before": len(mo_before),
        "count_after": len(mo_after),
        "lost": lost,
        "added": added,
        "changed": changed,
        "cache_before": len(cache_before),
        "cache_after": len(cache_after),
        "same": not lost and not added,
        "content_same": content_same,
        "content_compared": not content_error,
        "content_error": content_error,
    }


def _print_report(diff: dict, before: str, after: str) -> None:
    msg(f"写回前：{before}")
    msg(f"写回后：{after}")
    msg(f"模块数：{diff['count_before']} → {diff['count_after']}")
    msg(f"编译缓存（__SRP，会随重编译增减，不算异常）："
        f"{diff['cache_before']} → {diff['cache_after']}")
    if diff["changed"]:
        msg(f"[内容变化] {len(diff['changed'])} 个：{'、'.join(diff['changed'])}")
    if diff["content_compared"]:
        if diff["same"] and not diff["changed"]:
            msg("模块集合与正文：完全一致")
        elif diff["same"]:
            msg("模块集合：一致；正文有变化（见上）")
        else:
            msg("模块集合：有差异")
    else:
        msg(f"模块正文未比对：{diff['content_error']}")
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
    msg("  python -m toolkit.vbabin 前.xlsm 后.xlsm   比对模块集合与正文")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
