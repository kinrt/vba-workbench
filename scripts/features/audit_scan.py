# -*- coding: utf-8 -*-
"""audit_scan —— 源码范围枚举、读取缓存与成员可达性扫描（structure / overview 共用）。

以 mixin 形式提供给 AuditEngine：方法只吃引擎实例上的 root/config/缓存，
不直接构造引擎；纯计算，不打印、不碰 COM、不写文件。
"""
from __future__ import annotations

import re
from collections import Counter, defaultdict
from pathlib import Path

from toolkit.base import read_source_text
from toolkit.contract import ErrorCode, Failure, Stage
from toolkit.repo import IMPLICIT_PROCS, files_under, module_name_of, rel_path
from toolkit.vba_parse import (MEMBER_END_RE, MEMBER_PROC_RE, MEMBER_PROP_RE,
                               code_lines, is_event_proc, mask_strings,
                               parse_members)

from .audit_rules import (_DECL_SKIP_RE, _DIM_LOCAL_RE, _DYN_RE, _END_WITH_RE,
                          _ID, _SIG_PARAM_RE, _WITH_RE, _names_in_list)


class ScanMixin:
    """范围枚举 + 源码读取缓存 + 引用可达性分析 + 单模块/依赖盘点。"""

    # ------------------------------------------------------------------
    # 范围枚举（只读源码）
    # ------------------------------------------------------------------
    def common_dir(self) -> Path:
        return self.root / self.defaults.get("common_dir", "src/common")

    def common_files(self, exclude: list[str] | None = None) -> list[Path]:
        return files_under(self.common_dir(), exclude)

    def project_cfg(self, name: str) -> dict | None:
        for proj in self.projects:
            if proj.get("name") == name:
                return proj
        return None

    def project_files(self, name: str, exclude: list[str] | None = None) -> list[Path]:
        proj = self.project_cfg(name)
        if proj and proj.get("src"):
            d = self.root / str(proj["src"])
        else:
            d = self.root / "src" / "projects" / name
        return files_under(d, exclude)

    def _project_dirs(self) -> list[Path]:
        """仓库根下的项目源码目录（按名字排序）；不依赖 toolkit 的全局 ROOT，便于测试隔离。"""
        d = self.root / "src" / "projects"
        if not d.is_dir():
            return []
        return [p for p in sorted(d.iterdir()) if p.is_dir()]

    def corpus_files(self, projects: list[str] | None = None,
                     exclude: list[str] | None = None) -> list[Path]:
        """统计范围：公共库 + 指定项目（不给就是全部项目），跳过占位壳。"""
        files = self.common_files(exclude)
        names = projects if projects else [d.name for d in self._project_dirs()]
        for n in names:
            files += self.project_files(n, exclude)
        return files

    # ------------------------------------------------------------------
    # 源码读取（只读，带缓存）
    # ------------------------------------------------------------------
    def code_of(self, path: Path) -> list[str]:
        if path not in self._code_cache:
            self._code_cache[path] = code_lines(path)
        return self._code_cache[path]

    def raw_of(self, path: Path) -> list[str]:
        if path not in self._raw_cache:
            self._raw_cache[path] = read_source_text(path).splitlines()
        return self._raw_cache[path]

    # ------------------------------------------------------------------
    # 引用统计（成员可达性分析，structure / overview 共用）
    # ------------------------------------------------------------------
    def _factories(self, cls_name: str, files: list[Path]) -> set[str]:
        """返回「返回本类实例」的工厂函数名，用来认出 Variant 变量。"""
        out: set[str] = set()
        pat = re.compile(
            rf"(?i)^\s*(?:Public|Private|Friend)?\s*Function\s+({_ID})\s*\([^)]*\)\s*As\s+{cls_name}\b")
        for f in files:
            for ln in self.code_of(f):
                m = pat.match(ln)
                if m:
                    out.add(m.group(1))
        return out

    def _instance_names(self, lines: list[str], cls_name: str, facs: set[str]) -> set[str]:
        names: set[str] = set()
        for ln in lines:
            for grp in re.findall(rf"(?i)({_ID}(?:\s*,\s*{_ID})*)\s+As\s+{cls_name}\b", ln):
                for n in re.split(r"\s*,\s*", grp):
                    n = re.sub(r"(?i)^(ByVal|ByRef|ParamArray|Optional)\s+", "", n.strip())
                    if n and n != cls_name:
                        names.add(n)
            for n in re.findall(rf"(?i)Set\s+({_ID})\s*=\s*New\s+{cls_name}\b", ln):
                names.add(n)
            if facs:
                alt = "|".join(re.escape(x) + r"\s*\(" for x in facs)
                for n in re.findall(rf"(?i)(?:Set\s+)?({_ID})\s*=\s*{alt}", ln):
                    names.add(n)
        return names

    def _declared_names(self, path: Path) -> set[str]:
        """模块自己声明的名字：过程名 + 形参 + Dim/Const/Static。"""
        lines = self.code_of(path)
        names = set(parse_members(lines))
        for ln in lines:
            m = _SIG_PARAM_RE.match(ln)
            if m:
                names |= _names_in_list(m.group(1))
                continue
            m = _DIM_LOCAL_RE.match(ln)
            if m:
                names |= _names_in_list(m.group(1))
        names.discard("")
        return names

    @staticmethod
    def _is_implicit(name: str) -> bool:
        return name.lower() in IMPLICIT_PROCS or is_event_proc(name)

    def _scan_class(self, files: list[Path], target: Path, cls_name: str,
                    members: set[str], facs: set[str]):
        ext: dict[str, Counter] = defaultdict(Counter)
        sus: Counter = Counter()
        dyn: dict[str, set[Path]] = defaultdict(set)
        chain = ("|".join(re.escape(x) for x in facs)) if facs else ""
        for f in files:
            if f == target:
                continue
            lines = self.code_of(f)
            insts = self._instance_names(lines, cls_name, facs)
            with_stack: list[bool] = []
            for ln in lines:
                mw = _WITH_RE.match(ln)
                if mw:
                    t = mw.group(1)
                    with_stack.append(t in insts or any(t.startswith(x) for x in facs))
                    continue
                if _END_WITH_RE.match(ln):
                    if with_stack:
                        with_stack.pop()
                code = mask_strings(ln)
                if any(with_stack):
                    for m in re.finditer(rf"\.({_ID})\b", code):
                        if m.group(1) in members:
                            ext[m.group(1)][f] += 1
                for m in re.finditer(rf"(?i)({_ID})\.({_ID})\b", code):
                    recv, name = m.group(1), m.group(2)
                    if name not in members:
                        continue
                    if recv in insts:
                        ext[name][f] += 1
                    else:
                        sus[name] += 1
                if chain:  # getTL(...).成员 这种链式写法
                    for m in re.finditer(rf"(?i)\b(?:{chain})\s*\([^()\n]*\)\s*\.({_ID})\b", code):
                        if m.group(1) in members:
                            ext[m.group(1)][f] += 1
                if _DYN_RE.search(ln):  # 动态调用必须看原始字符串
                    for s in re.findall(r'"([^"\n]*)"', ln):
                        for name in members:
                            if re.search(rf"(?i)\b{re.escape(name)}\b", s):
                                dyn[name].add(f)
        return ext, sus, dyn

    def _scan_module(self, files: list[Path], target: Path, members: set[str], stem: str):
        ext: dict[str, Counter] = defaultdict(Counter)
        sus: Counter = Counter()
        dyn: dict[str, set[Path]] = defaultdict(set)
        for f in files:
            if f == target:
                continue
            lines = self.code_of(f)
            local = self._declared_names(f)
            for ln in lines:
                code = mask_strings(ln)
                for m in re.finditer(rf"(?i)(\.)?\b({_ID})\b", code):
                    dotted, tok = m.group(1), m.group(2)
                    if tok not in members:
                        continue
                    if dotted:
                        prev = re.findall(rf"({_ID})\.\s*$", code[:m.start(2)])
                        if prev and prev[0].lower() == stem.lower():
                            ext[tok][f] += 1
                        else:
                            sus[tok] += 1
                    else:
                        if _DECL_SKIP_RE.match(ln):
                            continue
                        if tok in local:
                            sus[tok] += 1  # 同名冲突：这个调用指向它自己的过程
                        else:
                            ext[tok][f] += 1
                if _DYN_RE.search(ln):
                    for s in re.findall(r'"([^"\n]*)"', ln):
                        for name in members:
                            if re.search(rf"(?i)\b{re.escape(name)}\b", s):
                                dyn[name].add(f)
        return ext, sus, dyn

    def _internal_graph(self, lines: list[str], names: set[str]) -> dict[str, set[str]]:
        """类/模块内部：每个过程引用了哪些成员。"""
        graph: dict[str, set[str]] = defaultdict(set)
        cur: str | None = None
        for ln in lines:
            s = ln.strip()
            m = MEMBER_PROC_RE.match(s) or MEMBER_PROP_RE.match(s)
            if m:
                cur = m.group("name")
                graph.setdefault(cur, set())
                continue
            if MEMBER_END_RE.match(s):
                cur = None
                continue
            if cur is None:
                continue
            for t in re.finditer(rf"(?i)\b({_ID})\b", mask_strings(ln)):
                tok = t.group(1)
                if tok in names and tok.lower() != cur.lower():
                    graph[cur].add(tok)
        return graph

    @staticmethod
    def _reachable(graph: dict[str, set[str]], entries: set[str], names: set[str]) -> set[str]:
        seen, stack = set(entries), list(entries)
        while stack:
            for nxt in graph.get(stack.pop(), ()):
                if nxt in names and nxt not in seen:
                    seen.add(nxt)
                    stack.append(nxt)
        return seen

    def analyze_target(self, target: Path, projects: list[str] | None = None,
                       exclude: list[str] | None = None) -> dict:
        """单个模块/类/窗体的成员盘点（纯计算；无成员时抛 Failure）。"""
        cls_name = module_name_of(target)
        files = self.corpus_files(projects, exclude)
        lines = self.code_of(target)
        members = parse_members(lines)
        if not members:
            raise Failure(f"{cls_name} 里没解析到成员。",
                          code=ErrorCode.SOURCE_PREFLIGHT, stage=Stage.PARSE)
        module_mode = target.suffix.lower() == ".bas"
        facs: set[str] = set() if module_mode else self._factories(cls_name, files)
        if module_mode:
            ext, sus, dyn = self._scan_module(files, target, set(members), cls_name)
        else:
            ext, sus, dyn = self._scan_class(files, target, cls_name, set(members), facs)
        graph = self._internal_graph(lines, set(members))

        names = set(members)
        entries = {n for n in names if ext.get(n)}
        entries |= {n for n in names if self._is_implicit(n)}
        live = self._reachable(graph, entries, names)

        rows = []
        for n, info in members.items():
            rows.append({
                "member": n, "kind": info["kind"], "access": info["access"],
                "line": info["line"], "span": info["span"],
                "ext": sum(ext.get(n, Counter()).values()),
                "ext_files": {str(rel_path(k)): v for k, v in ext.get(n, Counter()).items()},
                "internal": sum(1 for _c, s in graph.items() if n in s),
                "suspect": sus.get(n, 0),
                "dynamic": sorted(str(rel_path(p)) for p in dyn.get(n, set())),
                "reachable": n in live,
            })
        rows.sort(key=lambda r: (r["ext"], r["internal"], r["member"]))

        only_test = [r for r in rows if r["ext"]
                     and all(Path(k).stem.startswith("测试") for k in r["ext_files"])]
        only_test_names = {r["member"] for r in only_test}
        dead = [r for r in rows if not r["reachable"] and not self._is_implicit(r["member"])]
        dead_names = {r["member"] for r in dead}
        single = [r for r in rows if r["ext"] and len(r["ext_files"]) == 1
                  and r["member"] not in only_test_names and r["member"] not in dead_names]
        single_names = {r["member"] for r in single}
        for r in rows:
            r["verdict"] = ("不可达" if r["member"] in dead_names
                            else "仅自测" if r["member"] in only_test_names
                            else "仅单文件" if r["member"] in single_names else "活跃")
        return dict(target=target, cls=cls_name, files=files, module_mode=module_mode,
                    facs=facs, rows=rows, dead=dead, only_test=only_test, single=single,
                    scope=("项目 " + "、".join(projects)) if projects else "公共库 + 全部项目",
                    excluded=list(exclude or []))

    def dep_rows(self, project: str, exclude: list[str] | None = None) -> list[dict]:
        """某张表对每个公共模块的引用情况（只统计该项目的源码文件）。"""
        scan = self.project_files(project, exclude)
        all_files = self.corpus_files(None, None)
        out = []
        for cf in self.common_files(exclude):
            cls_name = module_name_of(cf)
            members = parse_members(self.code_of(cf))
            if not members:
                continue
            module_mode = cf.suffix.lower() == ".bas"
            if module_mode:
                ext, _sus, _dyn = self._scan_module(scan, cf, set(members), cls_name)
            else:
                facs = self._factories(cls_name, all_files)
                ext, _sus, _dyn = self._scan_class(scan, cf, cls_name, set(members), facs)
            per_file: Counter = Counter()
            for c in ext.values():
                per_file.update(c)
            top = "、".join(f"{Path(k).name}×{v}" for k, v in per_file.most_common(2)) or "-"
            out.append({
                "module": cls_name, "kind": self._kind_of(cf),
                "refs": sum(per_file.values()),
                "used": sum(1 for n in members if ext.get(n)),
                "files": len(per_file),
                "top": top,
            })
        out.sort(key=lambda d: (-d["refs"], d["module"]))
        return out

    @staticmethod
    def _kind_of(path: Path) -> str:
        if path.suffix.lower() == ".cls":
            return "类"
        if path.suffix.lower() == ".frm":
            return "窗体"
        if path.parent.name == "Microsoft Excel 对象":
            return "文档"
        return "模块"
