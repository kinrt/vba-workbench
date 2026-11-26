# -*- coding: utf-8 -*-
"""audit_indicators —— indicators 视角：H01–H19 离线体检指标计算。

运行期探测（H16/H17）由 handler 经网关注入，引擎只判；纯计算，不碰 COM、不写文件。
"""
from __future__ import annotations

import re
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Iterable

from toolkit.base import entry_command, read_source_text
from toolkit.check import check_sheet_modules
from toolkit.files import latest_backup
from toolkit.repo import (DOC_MODULE_RE, module_name_of, read_common_version,
                          rel_path, resolve_target, source_files)
from toolkit.vba_parse import MEMBER_PROC_RE, code_lines, parse_members, scan_rules
from toolkit.vbabin import (compare_vba_modules, sheet_pages, vba_documents,
                             vba_module_names)

from .audit_rules import (BACKUP_BATCH_LIMIT, BACKUP_SIZE_MB_LIMIT,
                          BUILTIN_CALLS, STATUSBAR_ALLOW_FILES,
                          _ALERTS_RE, _CALL_TARGET_RE, _DECLARE_RE,
                          _PROC_HEAD_RE, _SETDATA_CALL_RE, _STATUS_RE,
                          _ignore_set)


class IndicatorMixin:
    """indicators H01–H19：离线规则扫描 + 备份体量 + 工程/表页对账。"""

    # ------------------------------------------------------------------
    # indicators：体检指标 H01–H19（离线纯计算；运行期探测由 handler 注入）
    # ------------------------------------------------------------------
    def _indicator_files(self) -> dict[Path, tuple[list[str], list[str]]]:
        """体检用源文件缓存：文件 → (原始行, 去注释行)。"""
        cache: dict[Path, tuple[list[str], list[str]]] = {}
        src = self.root / "src"
        if not src.is_dir():
            return cache
        for p in sorted(src.rglob("*")):
            if p.suffix.lower() not in (".bas", ".cls", ".frm") or not p.is_file():
                continue
            try:
                lines = read_source_text(p).splitlines()
            except Exception:
                continue
            cache[p] = (lines, code_lines(p))
        return cache

    @staticmethod
    def _top_level_args(text: str) -> list[str]:
        """按顶层逗号切参数（括号/引号里的逗号不算），H03 判参顺序用。"""
        out, buf, depth, quote = [], [], 0, False
        for ch in text:
            if ch == '"':
                quote = not quote
            if not quote:
                if ch in "([":
                    depth += 1
                elif ch in ")]":
                    depth -= 1
                elif ch == "," and depth == 0:
                    out.append("".join(buf).strip())
                    buf = []
                    continue
            buf.append(ch)
        out.append("".join(buf).strip())
        return [a for a in out if a != ""]

    @staticmethod
    def _call_args(line: str, start: int) -> str:
        """取 setData( … ) 的括号内文本（括号配平）。"""
        depth, quote = 0, False
        for i in range(start, len(line)):
            ch = line[i]
            if ch == '"':
                quote = not quote
            if quote:
                continue
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    return line[start + 1:i]
        return ""

    def _statusbar_violations(self, files: dict[Path, tuple[list[str], list[str]]]) -> list[str]:
        """H06：只允许进度/日志通道写状态栏。"""
        out: list[str] = []
        for p, (_raw, code) in files.items():
            if p.name in STATUSBAR_ALLOW_FILES or "测试" in p.name or "自测" in p.name:
                continue
            text = "\n".join(code)
            if (p.suffix.lower() == ".cls"
                    and re.search(r"(?im)^\s*(?:Public\s+|Private\s+)?Sub\s+结束\b", text)
                    and re.search(r"(?i)Application\.StatusBar\s*=\s*(?:False|\"\")", text)):
                continue
            for i, ln in enumerate(code, 1):
                if not _STATUS_RE.search(ln):
                    continue
                rhs = ln.split("=", 1)[1].strip().lower() if "=" in ln else ""
                if rhs.startswith("false") or rhs.startswith('""'):
                    continue
                out.append(f"{rel_path(p)}:{i}")
        return out

    def _controls_pairing(self, files: dict[Path, tuple[list[str], list[str]]]) -> list[tuple[Path, int, str]]:
        """H05b：DisplayAlerts 只关不恢复才算命中。"""
        hits: list[tuple[Path, int, str]] = []
        for p, (raw, _code) in files.items():
            proc_start, proc_name, values, body = 0, "", [], []

            def flush() -> None:
                if proc_name and values and values[-1].lower().startswith(("false", "0")) \
                        and not any("Kill" in b for b in body):
                    hits.append((p, proc_start, f"{proc_name}：DisplayAlerts 只关未恢复"))

            for i, ln in enumerate(raw, 1):
                head = MEMBER_PROC_RE.match(ln.strip())
                m = head or _DECLARE_RE.match(ln)
                if m:
                    flush()
                    proc_name = head.group("name") if head else m.group(1)
                    proc_start, values, body = i, [], []
                body.append(ln)
                a = _ALERTS_RE.search(ln)
                if a:
                    values.append(a.group(1))
            flush()
        return hits

    def indicator_rows(self, exclude: list[str] | None = None,
                       ignore: Iterable[str] | None = None,
                       live_probe: dict | None = None) -> list[dict]:
        """逐项算体检指标，返回 [{id, name, state, count, detail, samples, suggested_command}]。

        state ∈ 通过 / 警告 / 未跑 / 锁定；live_probe 由 handler 经网关取来，引擎只判。
        """
        files = self._indicator_files()
        rows: list[dict] = []
        ignored = _ignore_set(ignore)

        def add(idx: str, name: str, state: str, count: Any, detail: str = "",
                samples: list | None = None, suggested: str = "") -> None:
            rows.append({"id": idx, "name": name, "state": state, "count": count,
                         "detail": detail, "samples": samples or [],
                         "suggested_command": suggested})

        # 定义表（H01）：过程与 Declare 名（字段/常量不算过程）
        defined: dict[str, list[Path]] = defaultdict(list)
        public_defs: dict[str, list[Path]] = defaultdict(list)
        for p, (_raw, code) in files.items():
            for proc, meta in parse_members(code).items():
                if meta.get("kind") != "方法":
                    continue
                defined[proc.lower()].append(p)
                if meta.get("access") != "private":
                    public_defs[proc.lower()].append(p)
            for ln in code:
                m = _DECLARE_RE.match(ln)
                if m:
                    defined[m.group(1).lower()].append(p)
                    public_defs[m.group(1).lower()].append(p)

        # H01 Call 目标都有定义
        h01: list[str] = []
        for p, (_raw, code) in files.items():
            for i, ln in enumerate(code, 1):
                m = _CALL_TARGET_RE.match(ln)
                if not m:
                    continue
                name = m.group(1)
                if "." in name or name.lower() in BUILTIN_CALLS or name.lower() in defined:
                    continue
                if ")." in ln:
                    continue  # `Call Controls(x).SetFocus` 这类对象方法
                h01.append(f"{rel_path(p)}:{i}  Call {name}")
        add("H01", "Call 目标都有定义", "警告" if h01 else "通过", len(h01),
            "扫裸 Call 名是否在公共库/本项目/Declare 里有定义（对象方法已排除）", h01)

        # H02 公共库内跨标准模块同名（只算 Public）
        dup = []
        for name, owners in public_defs.items():
            std = {q for q in owners if q.parent.name == "模块" and q.parent.parent.name == "common"}
            if len(std) > 1:
                dup.append((name, sorted(rel_path(q) for q in std)))
        add("H02", "公共库跨模块无同名过程", "警告" if dup else "通过", len(dup),
            "同名过程裸调会二义性；同一模块内重名见 H02b", dup)

        # H02b 同一模块内过程重名
        h02b: list[str] = []
        for p, (_raw, code) in files.items():
            if p.suffix.lower() not in (".bas", ".cls", ".frm"):
                continue
            seen: dict[tuple[str, str], int] = {}
            for i, ln in enumerate(code, 1):
                s = ln.strip()
                if not s or s.startswith("'") or re.match(r"(?i)^Rem\b", s):
                    continue
                m = _PROC_HEAD_RE.match(s)
                if not m:
                    continue
                nm = m.group(2).lower()
                kind = re.sub(r"\s+", " ", m.group(1).lower())
                key = (nm, kind)
                if key in seen:
                    h02b.append(f"{rel_path(p)}:{i}  {m.group(2)}（{m.group(1)}，首次在第 {seen[key]} 行）")
                else:
                    seen[key] = i
        add("H02b", "同一模块内过程不重名", "警告" if h02b else "通过", len(h02b),
            "VBA 不支持过程重载；同一模块里同名 Sub/Function/Property 会编译失败", h02b)

        # H03 setData 参数顺序
        h03: list[str] = []
        for p, (_raw, code) in files.items():
            for i, ln in enumerate(code, 1):
                for m in _SETDATA_CALL_RE.finditer(ln):
                    args = self._top_level_args(self._call_args(ln, m.end() - 1))
                    if len(args) == 3 and args[2].startswith('"') and not args[1].startswith('"'):
                        h03.append(f"{rel_path(p)}:{i}  setData({', '.join(args)})")
        add("H03", "setData 参数顺序", "警告" if h03 else "通过", len(h03),
            "第 3 参是字段名字面量、第 2 参不是（顶层逗号切分，Array 批量写法不算）", h03)

        # H04 / H05a 走 scan_rules（与写回后校验同一套规则）
        rules = scan_rules(list(files), exclude)
        if "manual_accel" in ignored or "h05a" in ignored:
            rules["manual_accel"] = []
        add("H04", "裸 New 表格工具只在 getTL", "警告" if rules["naked_new"] else "通过",
            len(rules["naked_new"]), "白名单：测试模块与 getTL 本体",
            [f"{rel_path(f)}:{i}  {s[:60]}" for f, i, s in rules["naked_new"][:10]])
        add("H05a", "项目侧不手写加速（三项）", "警告" if rules["manual_accel"] else "通过",
            len(rules["manual_accel"]), "ScreenUpdating／Calculation／EnableEvents；块上方三行内标 [加速例外] 放行",
            [f"{rel_path(f)}:{i}  {s[:60]}" for f, i, s in rules["manual_accel"][:10]])
        alerts = self._controls_pairing(files)
        add("H05b", "DisplayAlerts 开关成对", "警告" if alerts else "通过", len(alerts),
            "同过程内 False 之后必须恢复 True；只关不恢复才算命中（删文件弹窗属正常用法）",
            [f"{rel_path(f)}:{i}  {s}" for f, i, s in alerts[:10]])

        # H06 状态栏只由日志/进度通道写
        h06 = self._statusbar_violations(files)
        if "h06" in ignored:
            h06 = []
        add("H06", "状态栏只由日志通道写", "警告" if h06 else "通过", len(h06),
            "允许清空式写入；测试/自测模块与自带 结束 的进度器类放行", h06)

        # H07 开启加速 有对应关闭
        h07: list[str] = []
        for p, (_raw, code) in files.items():
            if "src" not in p.parts or "projects" not in p.parts:
                continue
            text = "\n".join(code)
            on = len(re.findall(r"开启加速", text))
            off = len(re.findall(r"关闭加速|退出加速|恢复环境", text))
            if on and not off:
                h07.append(f"{rel_path(p)}  开启 {on} / 收尾 {off}")
        add("H07", "开启加速 都有收尾", "警告" if h07 else "通过", len(h07),
            "同一文件里出现 关闭加速 / 退出加速 / 恢复环境（后两者内部都会 关闭加速）", h07)

        # H08 无带数字后缀模块（源码视角）
        h08: list[str] = []
        for p in files:
            base = re.sub(r"\(.*\)$", "", p.stem)
            if re.fullmatch(r"(Sheet|Chart)\d+", base) or base == "CommonVersion":
                continue
            if re.search(r"\d+$", base):
                h08.append(str(rel_path(p)))
        add("H08", "无带数字后缀的模块", "警告" if h08 else "通过", len(h08),
            "`X1` 与 `X` 并存属 Import 抢名残留；工程侧见 structure ⑫", h08)

        projects = self.projects
        h09, h15, h19, vers = [], [], [], []
        common_dir = self.common_dir()
        actual_ver = read_common_version(common_dir)
        for proj in projects:
            path, _why = resolve_target(proj)
            if not path:
                continue
            locked = proj.get("common_version")
            if locked and actual_ver and locked != actual_ver:
                vers.append(f"{proj['name']}  锁定 {locked} / 源码 {actual_ver}")
            src = proj.get("src")
            src_names = {module_name_of(f) for f in source_files(self.root / str(src))} if src else set()
            try:
                comp = set(vba_module_names(path))
            except Exception as e:
                h09.append(f"{proj['name']}  读不到工程：{e}")
            else:
                lost = sorted(n for n in src_names if n not in comp)
                extra = sorted(n for n in comp if n not in src_names and DOC_MODULE_RE.match(n)
                               and n.lower() != "thisworkbook")
                if lost or extra:
                    h09.append(f"{proj['name']}  源码有工程无 {lost} / 工程有源码无 {extra}")
            try:
                with zipfile.ZipFile(path) as z:
                    xml = z.read("xl/workbook.xml").decode("utf-8", "ignore")
                m = re.search(r"<calcPr[^>]*calcMode=\"([^\"]+)\"", xml)
                mode = m.group(1) if m else "（缺 calcPr）"
                if mode != "manual":
                    h15.append(f"{proj['name']}  calcMode={mode}")
            except Exception as e:
                h15.append(f"{proj['name']}  读不到：{e}")
            try:
                bad = check_sheet_modules(sheet_pages(path), vba_documents(path))
                h19 += [f"{proj['name']}  [{t}] {o}：{d}" for t, o, d in bad]
            except Exception as e:
                h19.append(f"{proj['name']}  读不到：{e}")
        add("H09", "src 模块名与工程一致", "警告" if h09 else "通过",
            f"{len(projects) - len(h09)}/{len(projects)}", "源码模块名（不含占位壳）逐个能在工程里找到", h09)
        add("H14", "项目公共库版本一致", "警告" if vers else "通过",
            f"{len(projects) - len(vers)}/{len(projects)}",
            f"公共库源码 {actual_ver or '?'}", vers)
        add("H15", "计算模式保持手动", "警告" if h15 else "通过",
            f"{len(projects) - len(h15)}/{len(projects)}",
            "读磁盘 `<calcPr calcMode=…>`，不依赖 COM（隐藏实例读 Application.Calculation 不可靠）", h15)
        add("H19", "表页与文档模块一一对应", "警告" if h19 else "通过",
            f"{len(projects) - len(h19)}/{len(projects)}", "表页取 workbook.xml，文档模块取工程流 PROJECT 的 Document 登记", h19)

        # H10 备份体量（阈值取 audit_rules 规则表）
        bk = self.root / self.defaults.get("backup_dir", ".backup")
        batches, size, per_class = [], 0, Counter()
        if bk.is_dir():
            for d in bk.iterdir():
                if not d.is_dir() or d.name in ("GBK", "watch"):
                    continue
                batches.append(d)
                size += sum(f.stat().st_size for f in d.rglob("*") if f.is_file())
                files_in = [f for f in d.iterdir() if f.is_file()]
                label = "（空批）"
                if files_in:
                    label = re.sub(r"\s*V?\d{6,8}.*$", "", files_in[0].stem).strip() or "（默认）"
                per_class[label] += 1
        keep = int(self.defaults.get("backup_keep", 10) or 10)
        keep_daily = int(self.defaults.get("backup_keep_daily", 1) or 0)
        min_keep = int(self.defaults.get("backup_min_keep", 3) or 3)
        threshold = keep + max(min_keep, 0)
        over_class = [] if keep_daily > 0 else [
            f"{k} {v} 批（每类阈值 ≤{threshold}）"
            for k, v in per_class.items() if v > threshold]
        mb = round(size / 1048576, 1)
        h10_problem = (len(batches) > BACKUP_BATCH_LIMIT or mb > BACKUP_SIZE_MB_LIMIT or over_class)
        add("H10", "备份体量可控", "警告" if h10_problem else "通过",
            f"{len(batches)} 批 / {mb} MB",
            f"阈值 ≤{BACKUP_BATCH_LIMIT} 批 / ≤{BACKUP_SIZE_MB_LIMIT} MB；每类 ≤{threshold} 批"
            + ("（keep_daily>0，每类不卡数）" if keep_daily > 0 else "")
            + "；各类批数："
            + ("、".join(f"{k or '（默认）'} {v}" for k, v in per_class.most_common()) or "无"),
            over_class,
            suggested=entry_command("sync.py", "clean-backup", "--apply")
            if h10_problem else "")

        # H11 断言基线
        test_file = self.root / "src" / "common" / "模块" / "测试表格工具.bas"
        asserts = len(re.findall(r"Call 断言\(", read_source_text(test_file))) if test_file.is_file() else 0
        baselines: dict[str, list[str]] = {}
        md_files = list((self.root / "src").rglob("*.md")) if (self.root / "src").is_dir() else []
        md_files += list(self.root.glob("*.md"))
        for p in md_files:
            for m in re.finditer(r"断言基线\s*\*\*(\d{2,3})\*\*", read_source_text(p)):
                baselines.setdefault(m.group(1), []).append(str(rel_path(p)))
        mism = [f"{num} 出现在 {'、'.join(baselines[num])}" for num in baselines if int(num) != asserts]
        add("H11", "断言基线四处一致", "警告" if mism else "通过", asserts,
            "实测 `Call 断言(` 处数，与文档里标 `断言基线 **N**` 的数字比对", mism)

        # H12 壳文件
        h12: list[str] = []
        for p, (raw, _code) in files.items():
            if p.stat().st_size < 200:
                text = "\n".join(raw)
                if "真实文件在" not in text and "Attribute VB_Name" not in text:
                    h12.append(f"{rel_path(p)}  {p.stat().st_size}B")
        add("H12", "壳文件都带标记", "警告" if h12 else "通过", len(h12), "小于 200B 的源码必须写明真实文件在哪", h12)

        # H13 契约变更清单：缺文件是“未跑”，文件在但没有编号行才是“警告”
        ct = self.root / "src" / "common" / "契约与回归.md"
        if not ct.is_file():
            add("H13", "契约变更清单有条目", "未跑", "-", f"缺文件 {rel_path(ct)}", [])
        else:
            rows13 = [ln for ln in read_source_text(ct).splitlines() if re.match(r"\|\s*\d+\s*\|", ln)]
            add("H13", "契约变更清单有条目", "通过" if rows13 else "警告", len(rows13),
                "`契约与回归.md` 的编号行",
                [] if rows13 else [f"{rel_path(ct)} 存在但没有形如 | 12 | … 的编号行"])

        # H16/H17：运行期探测由 handler 经网关注入；没开表就标未跑，不猜
        live = live_probe or {"checked": 0, "total": len(projects), "h16": [], "h17": [], "reason": ""}
        if live["checked"] == 0:
            reason = live.get("reason") or "没有项目的工作簿处于打开状态"
            add("H16", "引用含 Microsoft Forms 2.0", "未跑", "-",
                f"{reason}；要读 `VBProject.References`，工程锁着读不到", [])
            add("H17", "VBA 工程可读", "未跑", "-",
                f"{reason}；要读 `CodeModule.CountOfLines`；Excel 重启后必上锁，先解一次密", [])
        else:
            add("H16", "引用含 Microsoft Forms 2.0", "警告" if live["h16"] else "通过",
                f"{live['checked']}/{live['total']}",
                "读已打开工作簿的 `VBProject.References`；只有存在窗体或表内控件时才要求包含 Microsoft Forms 2.0", live["h16"])
            add("H17", "VBA 工程可读", "警告" if live["h17"] else "通过",
                f"{live['checked']}/{live['total']}",
                "读已打开工作簿的 `CodeModule.CountOfLines`；读不到的通常是工程还锁着", live["h17"])

        # H18：当前工作簿 vs 最近备份的模块清单（无备份时退回源码清单）
        bk_root = self.root / self.defaults.get("backup_dir", ".backup")
        h18_checked, h18_no_bak, h18_bad, h18_notes = self._module_manifest_rows(projects, bk_root)
        if h18_checked == 0:
            add("H18", "push 后模块清单不变", "未跑", "-",
                "没有可比对的备份或源码：" + ("；".join(h18_notes) or "没有项目"), [])
        else:
            detail = f"备份/源码对账 {h18_checked}/{len(projects)}"
            if h18_no_bak:
                detail += f"（{h18_no_bak} 项无备份，按源码清单）"
            add("H18", "push 后模块清单不变", "警告" if h18_bad else "通过",
                f"{h18_checked}/{len(projects)}", detail + "；" + "；".join(h18_notes), h18_bad)

        return rows

    def _module_manifest_rows(self, projects: list[dict],
                              backup_root: Path) -> tuple[int, int, list[str], list[str]]:
        """H18：当前工作簿 vs 最近一份备份的模块清单；无备份退回源码清单。"""
        checked = missing = 0
        issues: list[str] = []
        notes: list[str] = []
        for proj in projects:
            path, why = resolve_target(proj)
            if not path:
                notes.append(f"{proj['name']}  解析不到表：{why}")
                continue
            bak = latest_backup(backup_root, path.name)
            if bak:
                try:
                    diff = compare_vba_modules(bak, path)
                    checked += 1
                    if not diff["same"]:
                        issues.append(
                            f"{proj['name']}  备份 {diff['count_before']} → 当前 {diff['count_after']}；"
                            f"丢 {diff['lost']} / 多 {diff['added']}")
                except Exception as e:
                    notes.append(f"{proj['name']}  备份比对失败：{e}")
                continue
            src = proj.get("src")
            if not src:
                notes.append(f"{proj['name']}  没有备份也没有 src，无法对账")
                continue
            try:
                comp = set(vba_module_names(path))
                real_names = {module_name_of(f) for f in source_files(self.root / str(src))}
                shell_names = {module_name_of(f) for f in source_files(self.root / str(src), include_shell=True)} - real_names
                lost = sorted(real_names - comp)
                extra = sorted(n for n in comp - real_names - shell_names
                               if not DOC_MODULE_RE.match(n))
                checked += 1
                missing += 1
                notes.append(f"{proj['name']}  无备份，按源码对账")
                if lost or extra:
                    issues.append(f"{proj['name']}  源码 {len(real_names)} + 壳 {len(shell_names)} / 当前 {len(comp)}；"
                                  f"缺 {lost} / 多 {extra}")
            except Exception as e:
                notes.append(f"{proj['name']}  源码对账失败：{e}")
        return checked, missing, issues, notes
