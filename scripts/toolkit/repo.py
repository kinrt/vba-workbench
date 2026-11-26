# -*- coding: utf-8 -*-
"""
toolkit.repo —— 仓库结构与公共库查询（不含任何业务判断）
============================================================
  · project_dirs / source_files：src 下目录与源码文件的清单
  · module_index：{模块名: 文件}（公共库、任意目录通用）
  · is_shell_file / shell_text：公共库占位壳的识别与生成
  · VERSION_RE / read_common_version：COMMON_VERSION 版本常量的严格匹配
  · find_source：按模块名或相对路径定位仓库里的真实源码文件

MODULE_GROUP_ORDER 是 VBE 工程树的四个分类目录名（分组展示用）。
"""

from __future__ import annotations

import hashlib
import re
from pathlib import Path

from .base import (ROOT, SHELL_HINT, SOURCE_EXTS, Failure, abs_path,
                   read_source_text, warn)

MODULE_GROUP_ORDER = ("模块", "类模块", "窗体", "Microsoft Excel 对象")

# mapping.json 的 defaults 里与解析相关的缺省值；load_config 读完配置后由 set_defaults 灌进来，
# resolve_target 在没有单项目覆盖时用它（配置只写一处，不在每个调用点重复默认值）。
_DEFAULTS: dict = {}


def set_defaults(defaults: dict | None) -> None:
    """由 toolkit.config.load_config 调用：登记 defaults 里的解析缺省。"""
    _DEFAULTS.clear()
    _DEFAULTS.update(defaults or {})

# ---------------------------------------------------------------------------
# VBE 模块类型（VBIDE 的 vbext_ComponentType 常量）与目录/扩展名映射
# ---------------------------------------------------------------------------
# sync 写回、vbabin 核对、audit 分类都要用同一套编号，别在调用端各自定义。
VBXT_STANDARD = 1
VBXT_CLASS = 2
VBXT_FORM = 3
VBXT_DOCUMENT = 100

# 能用 Import 整块导入的类型；文档模块（ThisWorkbook / SheetN）不能 Remove，
# 只能改 CodeModule 正文，所以不在这个集合里。
IMPORTABLE = {VBXT_STANDARD, VBXT_CLASS, VBXT_FORM}

# VBE 的 Export 不会自动补扩展名，必须自己按模块类型加
EXT_FOR_TYPE = {
    VBXT_STANDARD: ".bas",
    VBXT_CLASS: ".cls",
    VBXT_FORM: ".frm",
    VBXT_DOCUMENT: ".bas",
}

# 反查：源码扩展名 → 该文件**应该**落成的模块类型。
# 写回时用它判断"能不能原地改正文"——文件类型和工程里的组件类型一致才敢原地改；
# 不一致（比如 .cls 的源码撞上 .bas 的组件）只能删了重导，交给 import_component 兜底。
TYPE_FOR_EXT = {".bas": VBXT_STANDARD, ".cls": VBXT_CLASS, ".frm": VBXT_FORM}

# 抓取时按 VBE 工程树的四类自动分目录，和 VBE 左侧看到的一致
TYPE_FOLDER = {
    VBXT_DOCUMENT: "Microsoft Excel 对象",
    VBXT_FORM: "窗体",
    VBXT_STANDARD: "模块",
    VBXT_CLASS: "类模块",
}

# ---------------------------------------------------------------------------
# 文档模块的文件名规则：`Sheet4(SKU统计).bas`
# ---------------------------------------------------------------------------
# VBE 左侧显示 `Sheet4 (SKU统计)`，源码里只叫 `Sheet4.bas` 就对不上号。所以抓源码时
# 把表页名放进括号，写回时再把括号整体丢弃 —— 规则只住在这里一份，sync / audit 共用。
#   · 模块名 = 文件名主干里括号之前那段（且必须以 ThisWorkbook / SheetN / ChartN … 开头）
#   · 老布局 `Sheet4.bas` 照样认；导出时若已有旧名文件，由 export_component 顺手清掉
DOC_MODULE_RE = re.compile(
    r"^(ThisWorkbook|Sheet\d+|Chart\d+|DialogSheet\d+|MacroSheet\d+|Module\d*)$",
    re.IGNORECASE,
)
_DOC_STEM_RE = re.compile(
    r"^(?P<name>ThisWorkbook|Sheet\d+|Chart\d+|DialogSheet\d+|MacroSheet\d+|Module\d*)"
    r"\s*\((?P<label>[^()]*)\)$",
    re.IGNORECASE,
)
# Windows 文件名非法字符（Excel 表页名本身已排除 : \ / ? * [ ]，这里再兜一层）
_ILLEGAL_NAME_CHARS = '<>:"/\\|?*\t\n\r'
# 表页名本身就是个编号（`Sheet10`、`Chart3`）：既认不出是哪张表，还会和 CodeName 混淆
_GENERIC_TAB_RE = re.compile(r"^(Sheet|Chart|DialogSheet|MacroSheet)\d+$", re.IGNORECASE)


def module_name_of(path: str | Path) -> str:
    """
    文件名 → VBA 模块名：`Sheet4(SKU统计).bas` → `Sheet4`；普通模块原样返回。

    只对"文档模块名 + 括号"这种形态做截断，`工具(旧).bas` 这类普通名字不受影响。
    """
    stem = Path(path).stem
    m = _DOC_STEM_RE.match(stem)
    return m.group("name") if m else stem


def doc_label_of(path: str | Path) -> str:
    """取文件名里的表页名（`Sheet4(SKU统计).bas` → `SKU统计`）；没有就返回空串。"""
    m = _DOC_STEM_RE.match(Path(path).stem)
    return m.group("label").strip() if m else ""


def doc_stem(code_name: str, sheet_name: str = "") -> str:
    """
    文档模块 → 源码文件名主干：`doc_stem("Sheet4", "SKU统计")` → `Sheet4(SKU统计)`。

    没有表页名（ThisWorkbook、或读不到表页名时）原样返回 code_name，保持老布局。
    表页名本身就是个编号（有的表页被改名为 `Sheet15` 之类）时也不加标签——
    它既认不出是哪张表，又会和 CodeName 撞在一起（`Sheet22(Sheet15).bas`）。
    """
    if not sheet_name or _GENERIC_TAB_RE.match(sheet_name.strip()):
        return code_name
    label = "".join("_" if c in _ILLEGAL_NAME_CHARS else c for c in sheet_name)
    label = label.strip(" .")
    return f"{code_name}({label})" if label else code_name


def project_dirs() -> list[Path]:
    """src/projects 下的项目源码目录（按名字排序）。"""
    d = ROOT / "src" / "projects"
    return [p for p in sorted(d.iterdir()) if p.is_dir()] if d.is_dir() else []


def source_files(d: Path, include_shell: bool = False, dedupe: bool = False) -> list[Path]:
    """
    目录下的源码文件（排序）；默认跳过公共库占位壳。

    dedupe=True：同一个模块名只留排序靠前的那个（写回时避免导错版本）；
    被丢弃的文件不在这里报，调用端要提示就自己用 module_index 比对。
    """
    if not d.is_dir():
        return []
    out = []
    seen: set[str] = set()
    for f in sorted(d.rglob("*")):
        if not (f.is_file() and f.suffix in SOURCE_EXTS):
            continue
        if is_noise_file(f):
            continue
        if not include_shell and is_shell_file(f):
            continue
        if dedupe:
            n = module_name_of(f)
            if n in seen:
                continue
            seen.add(n)
        out.append(f)
    return out


def source_files_at(spec: str | Path) -> list[Path]:
    """范围可以是单个源码文件，也可以是目录（相对路径按仓库根解析）。"""
    p = spec if isinstance(spec, Path) else abs_path(spec)
    if p.is_file():
        return [p] if p.suffix.lower() in SOURCE_EXTS else []
    if not p.is_dir():
        raise Failure(f"范围不存在：{p}")
    return source_files(p)


def files_under(d: Path, exclude: list[str] | None = None) -> list[Path]:
    """目录下的真实源码（跳过占位壳），按模块名关键字排除（`测试` 之类）。"""
    files = source_files(d)
    if exclude:
        low = [x.lower() for x in exclude]
        files = [f for f in files
                 if not any(x in module_name_of(f).lower() for x in low)]
    return files


def rel_path(path: Path, posix: bool = False) -> str:
    """显示用路径：仓库内给相对路径，仓库外原样给绝对路径。"""
    try:
        p = path.relative_to(ROOT)
    except ValueError:
        return path.as_posix() if posix else str(path)
    return p.as_posix() if posix else str(p)


def module_index(d: Path) -> dict[str, Path]:
    """目录下所有可写回模块：{模块名: 文件}（同名取排序后的第一个）。"""
    out: dict[str, Path] = {}
    if not d.is_dir():
        return out
    for f in sorted(d.rglob("*")):
        if f.is_file() and f.suffix.lower() in SOURCE_EXTS and not is_noise_file(f):
            out.setdefault(module_name_of(f), f)
    return out


# 同步盘（如 OneDrive、坚果云）产出的冲突副本与编辑器临时文件，永远不是源码。
# 典型名：`Sheet3(模块名)_<设备名>_11月-18-200609-2026_Conflict.bas`
NOISE_MARKS = ("_conflict", "-conflict", "冲突")


def is_noise_file(path: Path) -> bool:
    """同步冲突副本 / 临时文件：名字含 `_Conflict`、以 `~` 开头或结尾的，一律跳过。"""
    name = path.name.lower()
    if any(m in name for m in NOISE_MARKS):
        return True
    return name.startswith("~") or name.endswith("~")


def is_shell_file(path: Path) -> bool:
    """判断是不是公共库占位壳：内容是单行注释且带固定提示语。"""
    if path.suffix.lower() not in SOURCE_EXTS:
        return False
    try:
        text = read_source_text(path).strip()
    except Exception:
        return False
    return text.startswith("'") and SHELL_HINT in text and "\n" not in text


# 源码扩展名 -> 中文类型名（生成壳注释用）
SHELL_TYPE_NAME = {".bas": "标准模块", ".cls": "类模块", ".frm": "窗体"}


# ---------------------------------------------------------------------------
# 代码口径白名单（audit / style / 任何检查脚本共用一份，别再各自写常量）
#   RULE_ALLOW_NEW_FILES  —— 允许出现裸 `New 表格工具` 的文件
#                            getTL 本体在 自定义通用函数.bas；
#                            测试表格工具.bas 的 A0/A/B/C 段是性能测试的被测对象
#   RULE_ALLOW_ACCEL_FILES —— 允许直接赋值 Application.ScreenUpdating/Calculation/EnableEvents
#                            同上：加速本体 + 性能对照
#   RULE_EXEMPT_MARK      —— 刻意要手写开关的代码块，上方三行内标这个记号即放行
#   IMPLICIT_PROCS        —— VBA 隐式入口过程（人工不可删、不算死代码）
# ---------------------------------------------------------------------------
RULE_ALLOW_NEW_FILES = {"测试表格工具.bas", "自定义通用函数.bas"}
RULE_ALLOW_ACCEL_FILES = {"自定义通用函数.bas", "测试表格工具.bas"}
RULE_NEW_MARK = "New 表格工具"
RULE_EXEMPT_MARK = "[加速例外]"
IMPLICIT_PROCS = {"class_initialize", "class_terminate", "main"}


def shell_text(rel_in_repo: str, suffix: str) -> str:
    """一行占位壳的内容：指向公共库里的真实文件。"""
    kind = SHELL_TYPE_NAME.get(suffix, "模块")
    return f"' 此文件是公共库{kind}，{SHELL_HINT}{rel_in_repo}"


# 版本常量：必须严格匹配 `COMMON_VERSION As String = "x.y.z"`，
# 只出现 "COMMON_VERSION" 字样的注释/说明不算（audit 的 infra 判定同用这一个正则）
VERSION_RE = re.compile(r'COMMON_VERSION\s+As\s+String\s*=\s*"([^"]+)"', re.IGNORECASE)


def read_common_version(common_dir: Path) -> str | None:
    """在公共库里找 COMMON_VERSION 常量；支持「按类型分目录」和「平铺」两种布局。"""
    if not common_dir.is_dir():
        return None
    for f in sorted(common_dir.rglob("*")):
        if not f.is_file() or f.suffix.lower() not in (".bas", ".cls"):
            continue
        m = VERSION_RE.search(read_source_text(f))
        if m:
            return m.group(1)
    return None


def find_source(spec: str) -> Path:
    """
    按「模块名」或「src 下相对路径 / 绝对路径」定位一个真实源码文件（跳过占位壳）。
    同名命中多个时优先 .cls，再按路径短的先；只警告不失败。
    """
    p = Path(spec)
    if p.is_absolute() and p.exists():
        return p
    cand = ROOT / spec
    if cand.exists():
        return cand.resolve()
    stem = Path(spec).stem
    hits = [f for f in (ROOT / "src").rglob("*")
            if f.is_file() and f.suffix in SOURCE_EXTS and module_name_of(f) == stem
            and not is_shell_file(f)]
    if not hits:
        raise Failure(f"找不到模块「{spec}」；给模块名，或 src 下的相对路径。")
    hits.sort(key=lambda f: (f.suffix != ".cls", len(f.parts)))
    if len(hits) > 1:
        warn(f"同名模块有 {len(hits)} 个，用第一个：{hits[0].relative_to(ROOT)}")
    return hits[0]


# ---------------------------------------------------------------------------
# 解析项目对应的 xlsm（sync 写回要用，audit 离线核对 vbaProject.bin 也要用）
# ---------------------------------------------------------------------------
# 文件名里的版本号：表名必须以 V<纯数字> 结尾（如 示例表B、示例表A V<数字>）。
# V 前可以有空格，V 后必须紧跟数字且到名字结尾——后面再有任何文字就不算版本
# （如「示例表A V<数字> (version 1)」「示例表B - 副本」都不认）。
# 数字按【整数】比较，越大越新——不当作日期解读。
DEFAULT_VERSION_PATTERNS = (r"[Vv](\d+)$",)


def parse_stem_version(stem: str, patterns: tuple[str, ...] = DEFAULT_VERSION_PATTERNS) -> tuple[int, str] | None:
    """
    从文件名（不含扩展名）里解析出版本号：取结尾 V 后面那串纯数字。
    返回 (版本号, 命中的原文片段，如 "V100")；认不出来返回 None。
    """
    for pat in patterns:
        if not pat:
            continue
        m = re.search(pat, stem, re.IGNORECASE)
        if not m:
            continue
        raw = m.group(1) if m.groups() else m.group(0)
        digits = "".join(c for c in raw if c.isdigit())
        if not digits:
            continue
        try:
            return int(digits), m.group(0).strip()
        except ValueError:               # 理论上到不了这里，纯保险
            continue
    return None


def _file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def resolve_target(proj: dict) -> tuple[Path | None, str]:
    """
    解析项目对应的 xlsm：
      1. 指定 file -> 直接用
      2. search_dir (+ prefix) -> 取文件名里版本号最大的一份
         版本号 = 表名结尾 V 后面那串纯数字（示例表B、示例表A V<数字>），
         按整数比较，越大越新；可用 version_pattern 自定义正则。
      3. require_version（默认开）：没有版本号的文件一律不算候选，宁可能报错
         也不猜——避免误选目录里同族的历史表或崩溃恢复副本。
         显式写 "require_version": false 才退回按修改时间取最新。
    """
    if proj.get("file"):
        p = abs_path(proj["file"])
        if p.exists():
            return p, "按配置路径"
        return None, f"文件不存在：{p}"

    search_dir = proj.get("search_dir")
    if not search_dir:
        return None, "缺少 file 或 search_dir 配置"

    d = abs_path(search_dir)
    if not d.is_dir():
        return None, f"目录不存在：{d}"

    prefix = proj.get("prefix") or ""
    custom = proj.get("version_pattern") or proj.get("date_pattern")
    pats = ((custom,) if custom else ()) + DEFAULT_VERSION_PATTERNS
    strict = proj.get("require_version", _DEFAULTS.get("require_version", True))

    versioned: list[tuple[int, float, Path, str]] = []
    others: list[tuple[float, Path]] = []

    for f in d.iterdir():
        if not f.is_file() or f.suffix.lower() not in (".xlsm", ".xlsb"):
            continue
        stem = f.stem
        if prefix and not stem.startswith(prefix):
            continue
        got = parse_stem_version(stem, pats)
        if got:
            ver, mark = got
            versioned.append((ver, f.stat().st_mtime, f, mark))
            continue
        others.append((f.stat().st_mtime, f))

    if versioned:
        # 先比版本号（整数），版本号相同的再比修改时间
        versioned.sort(key=lambda x: (x[0], x[1]))
        ver, _m, path, mark = versioned[-1]
        reason = f"按版本号 {mark} 取最新，候选 {len(versioned)} 个"
        # 同尺寸 + 同 mtime 的版本号候选：不能静默只挑最后一份
        try:
            size, mtime_ns = path.stat().st_size, path.stat().st_mtime_ns
            twins = [p for _v, _mt, p, _mk in versioned
                     if p != path and p.stat().st_size == size
                     and p.stat().st_mtime_ns == mtime_ns]
            if twins:
                names = "、".join(p.name for p in twins)
                same = all(_file_sha256(p) == _file_sha256(path) for p in twins)
                if same:
                    reason += f"；注意：同尺寸同 mtime 的重复副本 {names}（内容相同，建议删掉多余副本）"
                else:
                    reason += (f"；注意：同尺寸同 mtime 但内容不同的候选 {names}"
                               f"（当前静默选了 {path.name}，需要时请显式配置 file）")
        except Exception:
            pass
        return path, reason
    if others and not strict:
        others.sort(key=lambda x: x[0])
        return others[-1][1], f"按修改时间取最新（文件名里没有 V 版本号，候选 {len(others)} 个）"
    if others:
        names = "、".join(sorted(f.name for _m, f in others)[:3])
        more = "…" if len(others) > 3 else ""
        return None, (f"没有以 V<数字> 结尾的文件（缺版本号的 {len(others)} 个：{names}{more}）；"
                      f"要按修改时间兜底请给项目加 \"require_version\": false")
    return None, f"目录里没有匹配 {prefix or '*.xlsm'} 的文件"
