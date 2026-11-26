# -*- coding: utf-8 -*-
"""audit_structure —— structure 视角：结构问题分组对账（壳 / 重名 / 绑定 / 文档表页）。

以 mixin 形式提供给 AuditEngine；纯计算，规则口径取自 audit_rules，不复制阈值表。
"""
from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path
from typing import Iterable

from toolkit.base import read_source_text
from toolkit.check import (check_doc_modules, check_sheet_modules,
                           check_suffix_modules)
from toolkit.repo import (DOC_MODULE_RE, doc_label_of, files_under,
                          is_shell_file, module_name_of, resolve_target,
                          source_files)
from toolkit.vba_parse import is_event_proc, parse_members, scan_rules
from toolkit.vbabin import sheet_pages, vba_documents, vba_module_names

from .audit_rules import (FATAL_STRUCTURE_GROUPS, NOTE_ONLY_GROUPS,
                          SHADOW_KNOWN, _ignore_set, _jsonable)


class StructureMixin:
    """structure 结构问题分组：后缀/文档/表页/绑定/重名/影子过程汇总。"""

    # ------------------------------------------------------------------
    # structure：结构问题分组（阈值/白名单取自 audit_rules 规则表）
    # ------------------------------------------------------------------
    def _suffix_issues(self) -> list[tuple[str, str, str]]:
        found: list[tuple[str, str, str]] = []
        for proj in self.projects:
            src = proj.get("src")
            path, _reason = resolve_target(proj)
            if not path:
                continue
            try:
                names = vba_module_names(path)
            except Exception:
                continue
            src_names = set()
            if src:
                src_names = {module_name_of(f) for f in source_files(self.root / str(src))}
            found += [(proj["name"], n, base)
                      for n, base in check_suffix_modules(names, src_names)]
        return found

    def _doc_issues(self) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]:
        orphan: list[tuple[str, str]] = []
        ungrabbed: list[tuple[str, str]] = []
        for proj in self.projects:
            src = proj.get("src")
            if not src:
                continue
            path, _reason = resolve_target(proj)
            if not path:
                continue
            try:
                names = vba_module_names(path)
            except Exception:
                continue  # 文件被占用/不是 xlsm：跳过，别假装查过
            comp_docs = [n for n in names if DOC_MODULE_RE.match(n)]
            src_docs = {module_name_of(f) for f in source_files(self.root / str(src))
                        if DOC_MODULE_RE.match(module_name_of(f))}
            orph, ung = check_doc_modules(src_docs, comp_docs)
            orphan += [(proj["name"], n) for n in orph]
            ungrabbed += [(proj["name"], n) for n in ung]
        return orphan, ungrabbed

    def _sheet_issues(self) -> list[tuple[str, str, str, str, str]]:
        found: list[tuple[str, str, str, str, str]] = []
        hints = {"数量": "换干净工程或补回缺的文档模块，别带着写回",
                 "无代码": "这张表没代码（宿主丢了），要写回先补文档模块",
                 "无宿主": "模块没有表页，push 也删不掉 → 换工程时换掉"}
        for proj in self.projects:
            path, _reason = resolve_target(proj)
            if not path:
                continue
            try:
                pages = sheet_pages(path)
                docs = vba_documents(path)
            except Exception:
                continue  # 文件被占用/不是 xlsx：跳过，不假装查过
            for kind, obj, detail in check_sheet_modules(pages, docs):
                found.append((proj["name"], kind, obj, detail, hints.get(kind, "")))
        return found

    @staticmethod
    def _split_vba_args(text: str) -> list[str]:
        """按顶层逗号切 VBA 实参（忽略括号内与字符串里的逗号）。"""
        out: list[str] = []
        cur: list[str] = []
        depth = 0
        inq = False
        i = 0
        while i < len(text):
            ch = text[i]
            if ch == '"':
                if inq and i + 1 < len(text) and text[i + 1] == '"':
                    cur.append('""')
                    i += 2
                    continue
                inq = not inq
                cur.append(ch)
            elif not inq and ch in "([{":
                depth += 1
                cur.append(ch)
            elif not inq and ch in ")]}":
                depth = max(0, depth - 1)
                cur.append(ch)
            elif not inq and depth == 0 and ch == ",":
                out.append("".join(cur).strip())
                cur = []
            else:
                cur.append(ch)
            i += 1
        if cur or text.strip():
            out.append("".join(cur).strip())
        return out

    def _vba_calls(self, text: str, names: tuple[str, ...]) -> list[tuple[str, list[str], int]]:
        """抽取 `注册快捷键(...)` / `注册动作(...)` 的调用参数与行号。"""
        found: list[tuple[str, list[str], int]] = []
        pat = re.compile(r"(?i)\b(" + "|".join(names) + r")\s*\(")
        for m in pat.finditer(text):
            start = m.end() - 1
            depth = 0
            inq = False
            i = start
            while i < len(text):
                ch = text[i]
                if ch == '"':
                    if inq and i + 1 < len(text) and text[i + 1] == '"':
                        i += 2
                        continue
                    inq = not inq
                elif not inq:
                    if ch == "(":
                        depth += 1
                    elif ch == ")":
                        depth -= 1
                        if depth == 0:
                            found.append((m.group(1), self._split_vba_args(text[start + 1:i]),
                                          text.count("\n", 0, m.start()) + 1))
                            break
                i += 1
        return found

    @staticmethod
    def _vba_string(raw: str) -> str | None:
        s = raw.strip()
        m = re.fullmatch(r'"((?:[^"]|"")*)"', s, re.S)
        return m.group(1).replace('""', '"') if m else None

    @staticmethod
    def _proc_args(text: str) -> tuple[str, str, int]:
        """从过程头返回 (名字, 访问级别, 必填参数数)。"""
        m = re.match(r"(?i)^\s*(Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
                     r"(?:Sub|Function)\s+([^\s(]+)(.*)$", text, re.S)
        if not m:
            return "", "", 0
        name = m.group(2).strip()
        access = (m.group(1) or "public").strip().lower()
        rest = m.group(3).strip()
        if not rest.startswith("("):
            return name, access, 0
        depth, inq = 0, False
        buf = []
        for ch in rest:
            if ch == '"':
                inq = not inq
            buf.append(ch)
            if not inq:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                    if depth == 0:
                        break
        params = "".join(buf)[1:-1]
        required = 0
        for a in StructureMixin._split_vba_args(params):
            s = a.strip()
            if not s or re.match(r"(?i)^(Optional|ParamArray)\b", s):
                continue
            required += 1
        return name, access, required

    def _binding_issues(self, exclude: list[str] | None = None) -> list[tuple[str, Path, int, str]]:
        """静态核对快捷键/动作绑定的目标（只读源码，不看运行中的 OnKey 状态）。"""
        issues: list[tuple[str, Path, int, str]] = []
        keys = [k for k in (exclude or ()) if k]

        def keep(p: Path) -> bool:
            return not (keys and any(k in str(p) for k in keys))

        def files_of(folder: Path) -> list[Path]:
            return [f for f in files_under(folder) if keep(f)]

        for proj in self.projects:
            src = proj.get("src")
            if not src:
                continue
            pdir = self.root / str(src)
            project_files = files_of(pdir)
            common = files_of(self.common_dir()) if self.common_dir().is_dir() else []

            std_procs: dict[str, list[tuple[str, str, int]]] = defaultdict(list)
            form_methods: dict[str, dict[str, tuple[str, int]]] = defaultdict(dict)
            sheet_methods: dict[str, dict[str, tuple[str, int]]] = defaultdict(dict)
            forms: set[str] = set()
            registry: set[str] = set()
            texts: dict[Path, str] = {}
            for f in project_files + common:
                try:
                    text = read_source_text(f)
                except Exception:
                    continue
                texts[f] = text
                fname = module_name_of(f)
                if f.suffix.lower() == ".frm":
                    forms.add(f.stem)
                    for ln in text.splitlines():
                        nm, acc, req = self._proc_args(ln)
                        if nm:
                            form_methods[f.stem][nm.lower()] = (acc, req)
                elif f.suffix.lower() == ".bas" and DOC_MODULE_RE.match(fname):
                    label = doc_label_of(f) or fname
                    for ln in text.splitlines():
                        nm, acc, req = self._proc_args(ln)
                        if nm:
                            sheet_methods[label][nm.lower()] = (acc, req)
                elif f.suffix.lower() in (".bas", ".cls") and not DOC_MODULE_RE.match(fname):
                    for ln in text.splitlines():
                        nm, acc, req = self._proc_args(ln)
                        if nm:
                            std_procs[nm.lower()].append((fname, acc, req))

            # 窗体实例 登记表（同一个模块里的 Function 窗体实例）
            for f in project_files:
                text = texts.get(f, "")
                if not re.search(r"(?im)^\s*(?:Public\s+)?Function\s+窗体实例\b", text):
                    continue
                m = re.search(r"(?ims)^\s*(?:Public\s+)?Function\s+窗体实例\b.*?^\s*End\s+Function\b", text)
                if m:
                    registry.update(re.findall(r'(?i)\bCase\s+"([^"]+)"', m.group(0)))

            call_files = [f for f in project_files if f.suffix.lower() in (".bas", ".cls")]
            for f in call_files:
                text = texts.get(f, "")
                call_text = "\n".join(self.code_of(f))
                for verb, args, line in self._vba_calls(call_text, ("注册快捷键", "注册动作")):
                    if verb.lower() == "注册快捷键":
                        if len(args) < 2:
                            issues.append((proj["name"], f, line, "注册快捷键 缺键位或目标名"))
                            continue
                        target = self._vba_string(args[1])
                        if target is None:
                            continue
                        if "." in target or "!" in target:
                            continue
                        hits = std_procs.get(target.lower(), [])
                        if not hits:
                            issues.append((proj["name"], f, line,
                                           f"注册快捷键目标「{target}」不是本簿标准模块过程"))
                        elif not any(acc != "private" and req == 0 for _m, acc, req in hits):
                            detail = "；".join(f"{m}（{acc}，必填 {r}）" for m, acc, r in hits)
                            issues.append((proj["name"], f, line,
                                           f"注册快捷键目标「{target}」不是 Public 零必填参数过程：{detail}"))
                        continue

                    if len(args) < 2:
                        issues.append((proj["name"], f, line, "注册动作 缺动作名或目标"))
                        continue
                    target_raw = args[1].strip()
                    extra = args[2:]
                    form = self._vba_string(target_raw)
                    if form is not None:
                        if form in forms:
                            if registry and form not in registry:
                                issues.append((proj["name"], f, line,
                                               f"窗体「{form}」没有在 窗体实例 里登记"))
                            if extra:
                                meth = self._vba_string(extra[0])
                                if meth and meth.lower() not in form_methods[form]:
                                    issues.append((proj["name"], f, line,
                                                   f"窗体「{form}」没有方法「{meth}」"))
                                if meth:
                                    acc, req = form_methods[form].get(meth.lower(), ("", 0))
                                    if req > max(0, len(extra) - 1):
                                        issues.append((proj["name"], f, line,
                                                       f"窗体方法「{form}.{meth}」必填 {req} 个，只传了 {max(0, len(extra) - 1)} 个"))
                            continue
                        hits = std_procs.get(form.lower(), [])
                        if hits:
                            req = min(r for _m, _a, r in hits)
                            if req > len(extra):
                                issues.append((proj["name"], f, line,
                                               f"动作过程「{form}」必填 {req} 个，只传了 {len(extra)} 个"))
                        elif not form.startswith("_"):
                            issues.append((proj["name"], f, line,
                                           f"注册动作目标「{form}」既不是窗体也不是本簿过程"))
                        continue
                    if target_raw.lower().startswith("thisworkbook.sheets("):
                        m = re.search(r'Sheets\s*\(\s*"((?:[^"]|"")*)"\s*\)', target_raw, re.I)
                        if not m:
                            continue
                        label = m.group(1).replace('""', '"')
                        if label not in sheet_methods:
                            issues.append((proj["name"], f, line,
                                           f"表页「{label}」没有对应的文档模块源码"))
                            continue
                        if extra:
                            meth = self._vba_string(extra[0])
                            if meth and meth.lower() not in sheet_methods[label]:
                                issues.append((proj["name"], f, line,
                                               f"表页「{label}」没有方法「{meth}」"))
                            if meth:
                                _acc, req = sheet_methods[label].get(meth.lower(), ("", 0))
                                if req > max(0, len(extra) - 1):
                                    issues.append((proj["name"], f, line,
                                                   f"表页方法「{label}.{meth}」必填 {req} 个，只传了 {max(0, len(extra) - 1)} 个"))
                        continue
                    if target_raw.lower().startswith("array("):
                        for item in re.findall(r'"((?:[^"]|"")*)"', target_raw):
                            nm = item.replace('""', '"')
                            if nm.lower() not in std_procs:
                                issues.append((proj["name"], f, line,
                                               f"动作数组目标「{nm}」不是本簿过程"))
        return issues

    def regression_issues(self, exclude: list[str] | None = None) -> dict:
        """三条强制口径回归扫描：裸 New / 手写加速 / 缺过程注释（扫描下沉 toolkit.scan_rules）。"""
        return scan_rules(self.corpus_files(None, exclude))

    def structure_issues(self, exclude: list[str] | None = None,
                         ignore: Iterable[str] | None = None) -> dict:
        """结构问题分组汇总；规则口径与旧版 1.5.6 一致。"""
        common_stems = {f.stem: f for f in self.common_files()}
        issues: dict[str, list] = {
            "stale": [], "orphan": [], "kind": [], "missing": [],
            "dup": [], "shadow": [], "shadow_known": [],
            "doc_orphan": [], "doc_ungrabbed": [], "sheet_mismatch": [],
            "naked_new": [], "manual_accel": [], "no_comment": [],
            "binding_target": [],
        }
        issues["doc_orphan"], issues["doc_ungrabbed"] = self._doc_issues()
        issues["sheet_mismatch"] = self._sheet_issues()
        issues["suffix"] = self._suffix_issues()
        issues["binding_target"] = self._binding_issues(exclude)
        issues.update(self.regression_issues(exclude))

        ignored = _ignore_set(ignore)
        if "h05a" in ignored:
            issues["manual_accel"] = []
        for group in ignored:
            if group in issues:
                issues[group] = []

        # 语义级检查（重名）才吃 exclude，且只比标准模块（.bas）的过程
        common_decl: dict[str, list[str]] = defaultdict(list)
        for name, f in common_stems.items():
            if f.suffix.lower() != ".bas":
                continue
            if exclude and any(x.lower() in name.lower() for x in exclude):
                continue
            for proc, meta in parse_members(self.code_of(f)).items():
                if meta["kind"] == "方法":
                    common_decl[proc].append(name)
        for n, owners in common_decl.items():
            if len(owners) > 1 and not is_event_proc(n):
                issues["dup"].append((n, sorted(owners)))

        for pdir in self._project_dirs():
            present = {module_name_of(f) for f in source_files(pdir, include_shell=True)}
            for f in source_files(pdir, include_shell=True):
                if is_shell_file(f):
                    c = common_stems.get(f.stem)
                    if c is None:
                        issues["orphan"].append(f)
                    elif c.suffix.lower() != f.suffix.lower():
                        issues["kind"].append((f, c))
                elif f.stem in common_stems:
                    issues["stale"].append((f, common_stems[f.stem]))
            for d in self.dep_rows(pdir.name, exclude):
                if d["refs"] and d["module"] not in present:
                    issues["missing"].append((pdir.name, d["module"]))

            for f in files_under(pdir, exclude):
                if f.stem in common_stems or f.suffix.lower() != ".bas":
                    continue
                for proc, meta in parse_members(self.code_of(f)).items():
                    if meta["kind"] != "方法" or is_event_proc(proc):
                        continue
                    if proc in common_decl:
                        if proc in SHADOW_KNOWN:
                            issues["shadow_known"].append((proc, f, sorted(common_decl[proc])))
                        else:
                            issues["shadow"].append((proc, f, sorted(common_decl[proc])))
        return issues

    @staticmethod
    def issues_total(issues: dict) -> int:
        return sum(len(v) for k, v in issues.items() if k not in NOTE_ONLY_GROUPS)

    @staticmethod
    def issues_fatal(issues: dict) -> int:
        return sum(len(issues.get(k, [])) for k in FATAL_STRUCTURE_GROUPS)

    @staticmethod
    def structure_rows(issues: dict) -> list[dict]:
        """结构问题 → JSON/表格行：{group, count, items}。"""
        return [{"group": k, "count": len(v), "items": [_jsonable(x) for x in v]}
                for k, v in issues.items()]
