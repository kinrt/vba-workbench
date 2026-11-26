# -*- coding: utf-8 -*-
"""
toolkit.base —— 工具包根、配置位置、编码约定与最底层原语
============================================================
只放**没有业务含义**的东西：

  · 包根 PKG_ROOT / 默认仓库根 CODE_ROOT / 配置位置 / 路径基准（workspace）
  · 源码扩展名与编码兜底顺序
  · Failure 与三个输出原语（msg / warn / fail）
  · abs_path：相对路径一律按 workspace 解析
  · interactive_ok / set_non_interactive：有没有人能在键盘上答话
  · read_source_text：UTF-8 优先读 VBA 源码（GBK 兜底）

配置查找顺序（先到先用）：环境变量 `VBA_WORKBENCH_CONFIG` → 从包根逐级向上找 `mapping.json`
（最多 7 层，遇到磁盘根为止）→ `%APPDATA%\\VBA工程管家\\mapping.json`。
路径基准 workspace = 配置里的 `workspace` 键（可选）> 配置所在目录；
配置直接放在工具包根时一律回到包根的上一级。

本模块不依赖任何其它子模块；其余子模块都可以依赖它。
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

# base.py 在 <包根>/scripts/toolkit/base.py
PKG_ROOT = Path(__file__).resolve().parents[2]
# 默认仓库根：包根（vba-workbench/）的上一级；子包挂在仓库里时它就是仓库根
CODE_ROOT = PKG_ROOT.parent


def tool_version() -> str:
    """读包根 VERSION；三个入口的 --version 共用这一份。"""
    f = PKG_ROOT / "VERSION"
    return f.read_text(encoding="utf-8").strip() if f.is_file() else "unknown"

CONFIG_NAME = "mapping.json"
CONFIG_ENV = "VBA_WORKBENCH_CONFIG"
APP_CONFIG = Path(os.environ.get("APPDATA", "")) / "VBA工程管家" / CONFIG_NAME if os.environ.get("APPDATA") else None

# 配置向上找几层：包根 → 父 → 祖父 …（子包/独立克隆/临时拷贝都能覆盖）
CONFIG_WALK_DEPTH = 7

_NON_INTERACTIVE = False

# 当前入口名（sync / audit / agent）；make_envelope 用它把 action 统一成 <入口>.<动词>
_TOOL_NAME = ""

# --json：stdout 只留一行 JSON，人话/警告全部转 stderr
_JSON_MODE = False
# JSON 模式下收口的警告纯文本（不含 [警告] 前缀）；render 出口并进 stats.warnings
_WARN_CACHE: list[str] = []
# 自动确认：只有 --apply 的调用方才会走到这里（agent.py push 内部也传 --apply）
_ASSUME_YES = False


def set_non_interactive(flag: bool = True) -> None:
    """`--no-input` 或流水线调用时打开：此后一律不再向 stdin 提问。"""
    global _NON_INTERACTIVE
    _NON_INTERACTIVE = bool(flag)


def set_json_mode(flag: bool = True) -> None:
    """开 JSON 模式：之后所有 msg/warn 走 stderr，stdout 只留给 render 出口的一行 JSON。

    切回人话模式时顺手清空缓存，避免上一轮 JSON 留下的警告串到后续进程内调用。
    """
    global _JSON_MODE
    _JSON_MODE = flag
    if not flag:
        _WARN_CACHE.clear()


def set_tool_name(name: str) -> None:
    """入口启动时调用：action 统一成 <入口>.<动词>（sync.push / audit.overview）。"""
    global _TOOL_NAME
    _TOOL_NAME = str(name or "").strip().lower()


def json_mode() -> bool:
    return _JSON_MODE


def set_assume_yes(flag: bool = True) -> None:
    """自动确认：ask_yes 一律返回 True（写回这类"明知要动手"的调用用）。"""
    global _ASSUME_YES
    _ASSUME_YES = flag


def cached_warnings() -> list[str]:
    """JSON 模式下收口的警告副本；人话模式下恒为空。"""
    return list(_WARN_CACHE)


def reset_warnings() -> None:
    """清空警告缓存（测试隔离用；正常一次 CLI 进程只走一个 render 出口）。"""
    _WARN_CACHE.clear()


def assume_yes() -> bool:
    return _ASSUME_YES


def interactive_ok() -> bool:
    """能不能在键盘上答话：没有 stdin、被重定向、或显式 --no-input 都算不行。"""
    if _NON_INTERACTIVE:
        return False
    try:
        return bool(sys.stdin) and sys.stdin.isatty()
    except (OSError, ValueError):
        return False


def find_config() -> Path | None:
    """按固定顺序找配置；都没有返回 None。设了环境变量就只认它，找不到也如实返回。"""
    env = os.environ.get(CONFIG_ENV, "").strip()
    if env:
        return Path(env).expanduser().resolve()
    cands: list[Path] = []
    d = PKG_ROOT
    for _ in range(CONFIG_WALK_DEPTH):
        cands.append(d / CONFIG_NAME)
        if d.parent == d:
            break
        d = d.parent
    if APP_CONFIG is not None:
        cands.append(APP_CONFIG)
    for p in cands:
        if p.is_file():
            return p.resolve()
    return None


def _config_base(cfg: Path | None) -> Path:
    """路径基准：配置在工具包根 → 包根的上一级；否则 `workspace` 键 > 配置所在目录。"""
    if cfg is None:
        return CODE_ROOT
    if cfg.parent == PKG_ROOT:
        # 配置和工具同层时，src 一定在它的上一级
        return cfg.parent.parent
    try:
        data = json.loads(cfg.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return cfg.parent
    ws = data.get("workspace") if isinstance(data, dict) else None
    if ws:
        p = Path(ws)
        return p if p.is_absolute() else (cfg.parent / p).resolve()
    return cfg.parent


_FOUND_CONFIG = find_config()
# 找不到时指向默认位置（仓库根），报错直接指向该放的地方
CONFIG_PATH = _FOUND_CONFIG or (CODE_ROOT / CONFIG_NAME)
# workspace：配置里 src / .backup 等相对路径的解析基准
ROOT = _config_base(_FOUND_CONFIG)

SOURCE_EXTS = (".bas", ".cls", ".frm")
SOURCE_ALL_EXTS = (".bas", ".cls", ".frm", ".frx")

TEXT_ENCODINGS = ("utf-8-sig", "utf-8", "gbk")

# 占位壳里的固定提示语；is_shell_file 靠它认壳
SHELL_HINT = "真实文件在："


class Failure(Exception):
    """统一业务失败：消息给人看，code/stage/hint/suggested/retryable 给 agent 和流水线分流。"""

    def __init__(self, message: str = "", *, code: str = "", stage: str = "",
                 hint: str = "", suggested: str = "", retryable: bool = False) -> None:
        super().__init__(message)
        self.code = code
        self.stage = stage
        self.hint = hint
        self.suggested = suggested
        self.retryable = bool(retryable)


def msg(text: str) -> None:
    """人话输出。开了 --json 时改走 stderr，保证 stdout 只有一行 JSON。"""
    print(text, file=sys.stderr if _JSON_MODE else sys.stdout)


def note(text: str) -> None:
    """过程提示：人话模式走 stdout；JSON 模式收口进 warnings 缓存，不碰 stderr。

    与 msg 的区别：msg 在 JSON 模式下仍写 stderr；note 用于"已合并本机覆盖层"这类
    过程性提示，JSON 下由 render 出口并进 stats.warnings，保证 ``cmd 2>&1`` 后
    stderr 不被过程提示污染、stdout 仍是单行信封。
    """
    if _JSON_MODE:
        _WARN_CACHE.append(str(text))
        return
    print(text)


def warn(text: str) -> None:
    """警告一律走 stderr：stdout 留给结果/JSON，脚本按行解析时不会被 [警告] 打断。

    JSON 模式下不碰 stderr：文本收进缓存，由 render 出口并进 stats.warnings，
    保证 ``cmd 2>&1`` 之后 stdout 仍是单行合法 JSON。
    """
    if _JSON_MODE:
        _WARN_CACHE.append(str(text))
        return
    print(f"[警告] {text}", file=sys.stderr)


def fail(text: str) -> None:
    print(f"[错误] {text}", file=sys.stderr)


def abs_path(p: str | Path) -> Path:
    """相对路径按 workspace 解析；绝对路径原样返回。"""
    path = Path(p)
    if path.is_absolute():
        return path
    return (ROOT / path).resolve()


def entry_command(entry: str, verb: str, *args: str) -> str:
    """拼一条可直接复制执行的入口命令：当前解释器 + 本包入口脚本绝对路径 + 参数。

    ``entry`` 相对包根 ``scripts/`` 解析（如 ``"sync.py"`` / ``"agent.py"``），
    脚本路径由本包位置解析，不写死安装盘符；解释器、脚本、参数含空格时由
    ``subprocess.list2cmdline`` 统一加引号，cmd 与 PowerShell 都能直接粘贴。
    """
    script = (PKG_ROOT / "scripts" / entry).resolve()
    argv = [sys.executable, str(script), verb, *args]
    return subprocess.list2cmdline([str(a) for a in argv])


def looks_like_command(text: str) -> bool:
    """判断一条建议是不是可复制执行的命令（不是就只当人话放 hint）。

    认两类写法：``python ``/``python-``/``vba-workbench``/``powershell ``/``pwsh ``/
    ``attrib `` 这类老前缀，以及 ``entry_command()`` 产出的「解释器绝对路径 + 本包
    入口绝对路径」——安装后解释器可能落在任意盘符，不能只靠字面前缀。
    """
    s = str(text or "").strip()
    if not s:
        return False
    if s.startswith(("python ", "python-", "vba-workbench", "powershell ", "pwsh ", "attrib ")):
        return True
    exe = str(sys.executable)
    return s.startswith(exe + " ") or s.startswith(f'"{exe}" ')


def read_source_text(path: Path) -> str:
    """读取 .bas/.cls/.frm：优先 UTF-8（带/不带 BOM），兜底 GBK（遗留文件与 VBE 原始产物）。"""
    raw = path.read_bytes()
    for enc in TEXT_ENCODINGS:
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    warn(f"{path.name} 编码无法识别，按 GBK 强制解码")
    return raw.decode("gbk", errors="ignore")


# ---------------------------------------------------------------------------
# 机器可读输出（给 agent / 流水线用）：--json 时只吐一行 JSON，不再打印人话
# ---------------------------------------------------------------------------
def _qualified_action(action: str) -> str:
    action = str(action or "unknown")
    if "." in action or not _TOOL_NAME:
        return action
    return f"{_TOOL_NAME}.{action}"


def _as_error_objects(errors, *, code: str = "", stage: str = "", hint: str = "",
                      suggested_command: str = "", retryable: bool = False) -> list:
    """把历史字符串错误统一成 {code,stage,message,hint,suggested_command,retryable}。"""
    rows = []
    for item in (errors or []):
        if isinstance(item, dict):
            row = dict(item)
        else:
            row = {"message": str(item)}
        row.setdefault("code", code or "FAILED")
        row.setdefault("stage", stage or "unknown")
        row.setdefault("message", row.get("code", "FAILED"))
        row.setdefault("hint", hint or "")
        row.setdefault("suggested_command", suggested_command or "")
        row.setdefault("retryable", bool(retryable))
        rows.append(row)
    return rows


def make_envelope(action: str, ok: bool, details=None, errors=None,
                  elapsed_ms: int | None = None, stats: dict | None = None,
                  exit_code: int | None = None, **extra):
    """统一信封：ok/action/details/errors/elapsed_ms/ts/stats[/exit_code]。"""
    import datetime

    merged_stats = dict(stats or {})
    # 错误详情从 extra 里摘出来，其余命令统计统一进 stats
    code = str(extra.pop("error_code", "") or "")
    stage = str(extra.pop("stage", "") or "")
    hint = str(extra.pop("hint", "") or "")
    suggested = str(extra.pop("suggested_command", "") or "")
    retryable = bool(extra.pop("retryable", False))
    # 旧调用把 suggested 当“人话建议”传：不是可执行命令就只放 hint
    if suggested and not looks_like_command(suggested):
        hint = hint or suggested
        suggested = ""
    merged_stats.update(extra)

    error_rows = _as_error_objects(errors, code=code, stage=stage, hint=hint,
                                   suggested_command=suggested, retryable=retryable)
    if not ok and not error_rows:
        error_rows = _as_error_objects(["命令失败"], code=code or "FAILED",
                                       stage=stage, hint=hint,
                                       suggested_command=suggested, retryable=retryable)

    if isinstance(details, list):
        detail_rows = details
    elif details is None:
        detail_rows = []
    else:
        detail_rows = [details]

    payload = {
        "ok": bool(ok) and not error_rows,
        "action": _qualified_action(action),
        "details": detail_rows,
        "errors": error_rows,
        "elapsed_ms": int(elapsed_ms if elapsed_ms is not None else 0),
        "ts": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
        "stats": merged_stats,
    }
    if exit_code is not None:
        payload["exit_code"] = int(exit_code)
    return payload


