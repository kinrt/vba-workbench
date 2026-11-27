# -*- coding: utf-8 -*-
"""agent_patch 切片：patch / assert / style / metrics 四个动词。

架构位置（1.6）：本模块只做应用层编排——解析参数、调 toolkit 领域服务与
adapters 适配器、把结果收进 ``Outcome``；不打印、不自带 JSON、不直接碰 COM
附着/打开（附着走 ``adapters.excel_gateway``，宏执行走 ``toolkit.macro.MacroRunner``）。

命令名、参数名、JSON 键、退出码与 1.5.6 逐字对齐；旧版 ``stats.lines`` 这类
键继续留在 ``stats`` 里，不搬进 ``details``。
"""
from __future__ import annotations

import contextlib
import io
import re
import time
from pathlib import Path
from typing import Any

from adapters import excel_gateway  # noqa: F401  (保留给 assert 链路)
from toolkit.base import Failure as BaseFailure, entry_command
from toolkit import dialogs, watchdog
from toolkit.contract import Failure, MutationMode, Outcome, Stage, to_contract_failure
from toolkit.dialogs import KIND_COMPILE_ERROR, KIND_VBE_DIALOG
from toolkit.files import write_source, write_text_utf8  # noqa: F401  (metrics 链路用)
from toolkit.macro import MacroRunner
from toolkit.registry import ArgSpec, CommandSpec, ProjectMode
from toolkit.repo import rel_path, resolve_target, source_files
from toolkit.session import WorkbookSession
from toolkit.state import Purpose, TargetState
from toolkit.style_rules import run_checks_with_details, run_fix
from toolkit.vba_parse import code_lines, parse_members  # noqa: F401  (metrics 链路用)

# ---------------------------------------------------------------------------
# 常量：与 1.5.6 同口径
# ---------------------------------------------------------------------------
TESTSHEET = "测试表"
DEFAULT_ASSERT_MACRO = "测试表格工具.运行全部断言"
PASS_RE = re.compile(r"通过\s*[:：]?\s*(\d+)")
FAIL_RE = re.compile(r"失败\s*[:：]?\s*(\d+)")

# metrics 口径正则（原样沿用 1.5.6，不改判定口径）
BOOL_RE = re.compile(r"\bAs\s+Boolean\b", re.I)
MODVAR_RE = re.compile(r"(?i)^\s*(?:Public|Private|Dim|Global)\s+(?!Const\b)([^\s(]+)")
ONERR_RE = re.compile(r"On\s+Error\s+Resume\s+Next", re.I)
SHELL_RE = re.compile(r"(?i)^\s*(Worksheet_|Workbook_|UserForm_|BtnClickShowForm|ProcShowForm|Class_)")


# ---------------------------------------------------------------------------
# 共用小助手：项目解析 / 断言宏与表名 / 日志目录
# ---------------------------------------------------------------------------
def _projects(ctx) -> list[dict]:
    return list(ctx.config.get("projects") or [])


def _defaults(ctx) -> dict:
    return dict(ctx.config.get("defaults") or {})


def _pick(ctx, name: str | None) -> tuple[dict, Path]:
    """按项目名取配置与工作簿路径；不给名字要求恰好一本（口径同 1.5.6 ``pick``）。"""
    items = _projects(ctx)
    if name:
        for proj in items:
            if proj["name"] == name or proj["name"].startswith(name):
                path, _how = resolve_target(proj)
                if path is None:
                    raise Failure(
                        f"「{proj['name']}」解析不到 xlsm",
                        code="TARGET_NOT_FOUND", stage=Stage.CONFIG,
                        suggested="检查 mapping 的 search_dir / prefix，确认 xlsm 在目录里",
                        retryable=False,
                    )
                return proj, path
        raise Failure(
            f"没有这个项目：{name}（可选：{'、'.join(p['name'] for p in items)}）",
            code="PROJECT_NOT_FOUND", stage=Stage.CONFIG,
            suggested="先跑 agent probe --json 看项目名，按输出里的名字重跑",
            retryable=False,
        )
    raise Failure(
        "请指定项目名", code="PROJECT_NOT_SPECIFIED", stage=Stage.CONFIG,
        suggested="加 --project 项目名；不知道名字先跑 agent probe --json",
        retryable=False,
    )


def _assert_settings(ctx, proj: dict, args: Any) -> tuple[str, str]:
    """断言宏与测试表：命令行 > 项目键 > defaults > 模板默认值（口径同 1.5.6）。"""
    defaults = _defaults(ctx)
    macro = (getattr(args, "macro", "") or "").strip() \
        or proj.get("assert_macro") or defaults.get("assert_macro") or DEFAULT_ASSERT_MACRO
    sheet = (getattr(args, "sheet", "") or "").strip() \
        or proj.get("assert_sheet") or defaults.get("assert_sheet") or TESTSHEET
    macro = str(macro).strip()
    if "." not in macro:
        raise Failure(f"断言宏要写成「模块.过程」：{macro}",
                      code="ASSERT_MACRO_INVALID", stage=Stage.CONFIG,
                      hint="在 mapping 配 assert_macro，或用 --macro 模块.过程 临时覆盖",
                      suggested="--macro 测试表格工具.运行全部断言")
    module, proc = macro.rsplit(".", 1)
    if not module.strip() or not proc.strip():
        raise Failure(f"断言宏要写成「模块.过程」：{macro}",
                      code="ASSERT_MACRO_INVALID", stage=Stage.CONFIG,
                      hint="模块名或过程名为空；检查 mapping 的 assert_macro 或 --macro",
                      suggested="--macro 测试表格工具.运行全部断言")
    return macro, str(sheet).strip() or TESTSHEET


def _log_dir_of(proj: dict) -> Path:
    """断言日志目录：<search_dir>/测试输出，没有就退回上级目录的 测试输出。"""
    base = Path(proj.get("search_dir") or ".")
    for cand in (base / "测试输出", base.parent / "测试输出"):
        if cand.is_dir():
            return cand
    return base / "测试输出"


def _root_path(ctx, rel: str | Path) -> Path:
    p = Path(rel)
    return p if p.is_absolute() else ctx.root / p


def _human_lines(buf: io.StringIO) -> list[str]:
    return [line for line in buf.getvalue().splitlines() if line.strip()]


# ---------------------------------------------------------------------------
# style 报告行前缀归一（N5 过渡口径：产品裁决前不重新分流通道，只保留警告标记）
# ---------------------------------------------------------------------------
WARN_PREFIX = "[警告]"
# 重复的警告前缀（含尾随空白），并入报告行时剥到只剩一个
_WARN_PREFIX_RE = re.compile(r"^\s*(?:\[警告\]\s*)+")
# stderr 侧杂项标记：错误/警告前缀、traceback 头与帧行；JSON 出口不进 stats.lines
_STDERR_JUNK_RE = re.compile(
    r"^\s*(?:\[警告\]|\[错误\]|Traceback\s*\(|File\s*\")")
# 错误通道前缀：人话模式下保持错误标记，不误标成警告
_STDERR_ERROR_RE = re.compile(r"^\s*\[错误\]")


def _normalize_warn_line(line: str) -> str:
    """stderr 侧警告行并入报告：重复前缀归一为一个，缺前缀补齐；空行原样返回。

    - 行已带一个或多个 ``[警告]``：剥掉全部重复前缀，正文前补回单个前缀；
    - 行无前缀（漏带前缀的 stderr 警告）：补齐一个；
    - 正文为空（整行只有前缀）：原样返回，不补空壳。
    """
    if not line.strip():
        return line
    body = _WARN_PREFIX_RE.sub("", line).strip()
    if not body:
        return line
    return f"{WARN_PREFIX} {body}"


def _scope_files(ctx, scope: str) -> list[Path]:
    """按范围收集源码文件；相对路径按 ctx.root 解析（口径同 1.5.6 ``scope_files``）。"""
    p = _root_path(ctx, scope)
    if p.is_file():
        return [p] if p.suffix.lower() in (".bas", ".cls", ".frm") else []
    if not p.is_dir():
        raise Failure(f"范围不存在：{p}", code="CONFIG_INVALID", stage=Stage.CONFIG,
                      hint="--scope 给 src 下的目录或单个源码文件路径")
    return source_files(p)


# ---------------------------------------------------------------------------
# metrics：源码体检基线（口径与 1.5.6 一致，只读；--out 落盘要 --apply）
# ---------------------------------------------------------------------------
def handle_metrics(ctx) -> Outcome:
    args = ctx.args
    files = source_files(ctx.root / "src")
    procs, spans, thin, modvar, bools, onerr, longest = 0, [], 0, 0, 0, 0, []

    for f in files:
        lines = code_lines(f)
        members = parse_members(lines)
        onerr += sum(1 for ln in lines if ONERR_RE.search(ln))

        proc_lines: set[int] = set()
        for name, info in members.items():
            if info["kind"] != "方法":
                continue
            start, span = info["line"], info["span"]
            procs += 1
            proc_lines.update(range(start, start + span + 1))
            body = [ln for ln in lines[start: start + span - 1] if ln.strip()]
            spans.append(len(body))
            if len(body) >= 8:
                longest.append((len(body), rel_path(f).split("/")[-1], name))
            if len(body) <= 3 and sum(1 for ln in body
                                      if re.match(r"(?i)^\s*Call\s+", ln)
                                      or re.match(r"(?i)^\s*[A-Za-z_\u4e00-\u9fff][\w.]*\s*$", ln)) <= 1:
                if not SHELL_RE.match(name):
                    thin += 1

        for info in members.values():
            if info["kind"] == "方法":
                bools += len(BOOL_RE.findall(lines[info["line"] - 1]))
        for i, ln in enumerate(lines, 1):
            if i in proc_lines or not ln.strip():
                continue
            if MODVAR_RE.match(ln) and ln[:1] not in (" ", "\t"):
                modvar += 1

    spans.sort()
    median = spans[len(spans) // 2] if spans else 0
    p90 = spans[int(len(spans) * 0.9)] if spans else 0
    longest.sort(reverse=True)

    items = [("模块 / 过程", f"{len(files)} 个源文件 / {procs} 个方法"),
             ("语句数 中位 / P90 / 最长", f"{median} / {p90} / {spans[-1] if spans else 0}"),
             ("薄包装（≤3 句、最多一次调用，去事件壳）", str(thin)),
             ("模块级变量（列 0 声明）", str(modvar)),
             ("布尔形参位", str(bools)),
             ("On Error Resume Next 处数", str(onerr))]
    table = "| 指标 | 数值 |\n|---|---|\n" + "\n".join(f"| {k} | {v} |" for k, v in items)

    human = [table, "", "最长过程："]
    for n, mod, name in longest[:6]:
        human.append(f"  {n:>4} 句  {mod}.{name}")

    if getattr(args, "out", ""):
        if getattr(args, "apply", False):
            write_text_utf8(Path(args.out),
                            "# 体检基线（由 agent.py metrics 生成，勿手改）\n\n" + table + "\n")
            human.append(f"\n已写入 {args.out}")
        else:
            human.append(f"\ndry-run：未写入 {args.out}（加 --apply 落盘）")

    return Outcome(
        action="agent.metrics", ok=True, details=[], errors=[],
        stats={"lines": human, "files": len(files), "procs": procs, "median": median,
               "p90": p90, "longest": spans[-1] if spans else 0, "thin": thin,
               "modvar": modvar, "bools": bools, "onerr": onerr,
               "timings": ctx.stage_timings()},
        human=human, exit_code=0, elapsed_ms=ctx.elapsed_ms(),
    )


# ---------------------------------------------------------------------------
# assert：跑断言宏并核对日志（看门狗 / 活动簿核对 / 清表 / 日志口径原样保留）
# ---------------------------------------------------------------------------
def _macro_exists(wb, macro: str) -> bool | None:
    """目标工作簿里是否真有指定「模块.过程」；工程读不到返回 None。"""
    module, proc = macro.rsplit(".", 1)
    module, proc = module.strip(), proc.strip()
    try:
        names = [str(c.Name) for c in wb.VBProject.VBComponents]
    except Exception:  # noqa: BLE001 - 工程只读/受保护，无法判断
        return None
    if module not in names:
        return False
    try:
        comp = wb.VBProject.VBComponents.Item(module)
        code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
    except Exception:  # noqa: BLE001 - 模块存在但代码读不出，交给调用方回退
        return None
    pat = (r"(?m)^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
           rf"(?:Sub|Function)\s+{re.escape(proc)}\b")
    return bool(re.search(pat, code))


def _ensure_active(xl, wb) -> None:
    """xl.Run 只跑活动工作簿的宏：切过去并核对，别让日志写到别的书里（踩过假失败）。"""
    pid = _pid_from_xl(xl)
    for _ in range(5):
        # 模态框会挡住 Activate；同进程的前台 MsgBox 也算目标不可切，不等到 5 次空转。
        if pid:
            blocked = _probe_dialog(pid)
            if blocked is not None:
                snippet = _dialog_snippet(getattr(blocked, "text", "")
                                          or getattr(blocked, "title", ""))
                if getattr(blocked, "kind", "") == dialogs.KIND_COMPILE_ERROR:
                    raise Failure(
                        f"检测到 VBE 编译错误框「{getattr(blocked, 'title', '')}」（{snippet}）",
                        code="VBE_COMPILE_ERROR", stage=Stage.VBE,
                        hint=(f"弹窗标题：{getattr(blocked, 'title', '') or '（读不到标题）'}；"
                              f"正文：{snippet}。先修 VBA 编译错误，再重跑 agent assert"),
                        suggested="在 VBE 里修掉编译错误并编译通过后重跑",
                        retryable=False,
                    )
                raise Failure(
                    f"检测到 Excel 模态框「{getattr(blocked, 'title', '')}」挡住「{wb.Name}」"
                    f"（{snippet}）",
                    code="TARGET_NOT_ACTIVE", stage=Stage.ATTACH,
                    hint="先关掉挡在前面的 MsgBox/InputBox 或受保护视图，确认目标工作簿是活动簿后重跑",
                    suggested="先处理前台弹窗，再重跑 agent assert；不要把弹窗留在 Excel 进程里",
                    retryable=False,
                )
        try:
            wb.Activate()
        except Exception as exc:  # noqa: BLE001 - COM 拒绝切换统一报目标不可切
            raise Failure(
                f"切不到「{wb.Name}」：{exc}",
                code="TARGET_NOT_ACTIVE", stage=Stage.ATTACH,
                suggested="先关掉挡在前面的弹窗/受保护视图，确认目标工作簿是活动簿后重跑",
                retryable=False,
            ) from exc
        time.sleep(0.4)
        if str(xl.ActiveWorkbook.Name) == str(wb.Name):
            return
    raise Failure(
        f"切不到「{wb.Name}」，当前活动的是 {xl.ActiveWorkbook.Name}",
        code="TARGET_NOT_ACTIVE", stage=Stage.ATTACH,
        suggested="先关掉挡在前面的弹窗/受保护视图，确认目标工作簿是活动簿后重跑",
        retryable=False,
    )


def _clear_testsheet(wb, baseline: int, sheet: str = TESTSHEET) -> str:
    """把测试表的数据行删到基线（历史脏行会致 AutoFilter Err 1004）。"""
    if sheet not in [str(s.Name) for s in wb.Sheets]:
        return f"无测试表（{sheet}）"
    ws = wb.Sheets(sheet)
    if not ws.ListObjects.Count:
        return f"测试表没有 ListObject（{sheet}）"
    lo = ws.ListObjects(1)
    before = int(lo.ListRows.Count)
    while int(lo.ListRows.Count) > baseline:
        lo.ListRows(int(lo.ListRows.Count)).Delete()
    after = int(lo.ListRows.Count)
    return f"{before} → {after}" + ("" if before == after else f"（已清到基线 {baseline}）")


def _wait_new_log(folder: Path, before: set[str], seconds: float) -> Path | None:
    end = time.time() + seconds
    while time.time() < end:
        now = {f.name: f for f in folder.glob("*.log")}
        fresh = [v for k, v in now.items() if k not in before]
        if fresh:
            time.sleep(0.4)  # 让 VBA 把最后几行写完
            return max(fresh, key=lambda f: f.stat().st_mtime)
        time.sleep(0.5)
    return None


def _read_log_text(log: Path) -> str:
    try:
        return log.read_text(encoding="gbk", errors="replace")
    except OSError as e:
        raise Failure(f"日志读不了：{log}（{e}）",
                      code="ASSERT_LOG_UNREADABLE", stage=Stage.ASSERT,
                      suggested="检查日志是否被占用/权限不足；关闭占用程序后重跑",
                      retryable=False)


# 1.5.6 断言宏两参取值：含性能=False，静默=True（基线 agent.py: Run(macro_ref, False, True)）
ASSERT_PERF = False
ASSERT_SILENT = True


def _remap_attach_failure(e: BaseFailure) -> Failure:
    """附着阶段的关闭/忙，跑到断言阶段统一归 assert 口径（同 1.5.6 重映射）。

    先归一到契约 Failure：toolkit/adapters 抛的 base.Failure 也能被这里接住。
    """
    e = to_contract_failure(e)
    if e.code in ("EXCEL_CLOSED", "EXCEL_BUSY") and e.stage == Stage.ATTACH:
        return Failure(str(e), code=e.code, stage=Stage.ASSERT,
                       suggested=e.suggested, retryable=e.retryable)
    return e


# ---------------------------------------------------------------------------
# 超时二次分类：PID 存活优先，再认弹窗（1.7 B1/B2/B4 回补）
# ---------------------------------------------------------------------------
def _pid_from_xl(xl) -> int:
    """取目标 Excel 实例 PID；读不到返回 0（交探测/枚举降级口径，不抛）。"""
    try:
        hwnd = int(getattr(xl, "Hwnd", 0) or 0)
    except Exception:  # noqa: BLE001
        return 0
    return watchdog.pid_from_hwnd(hwnd) if hwnd > 0 else 0


def _probe_pid_alive(pid: int) -> bool | None:
    """存活探测 seam：测试可替换桩函数；默认走 watchdog.pid_alive 三态口径。"""
    return watchdog.pid_alive(pid)


def _probe_dialog(pid: int | None):
    """弹窗枚举 seam：测试可替换桩函数；默认走 dialogs.detect_target_dialog。"""
    return dialogs.detect_target_dialog(pid)


def _dialog_snippet(text: str, limit: int = 160) -> str:
    """正文片段：压掉多余空白，超长截断，给 hint 点名关键内容。"""
    snippet = " ".join(str(text or "").split())
    return snippet[:limit] if len(snippet) > limit else snippet


def _classify_watchdog_timeout(xl, *, timeout: float) -> Failure:
    """超时后的二次分类：实例退出优先，其次编译错误框，其余模态框点名弹窗，判不出回退通用超时。

    优先级（backlog-1.7 B4/B2 口径）：
      1. PID 已退出 → ``EXCEL_CLOSED``（``stage=assert``），优先于超时码；
      2. 编译错误弹窗 → ``VBE_COMPILE_ERROR``（``stage=vbe``）；
      3. 通用 VBE / Excel 模态框 → ``WATCHDOG_TIMEOUT``（``stage=assert``），hint 点名标题与正文；
      4. 存活未知 / 认不出弹窗 → 回退通用 ``WATCHDOG_TIMEOUT``。
    """
    pid = _pid_from_xl(xl)
    alive = _probe_pid_alive(pid) if pid > 0 else None
    if alive is False:
        return Failure(
            f"{timeout:g}s 内宏没有返回；目标 Excel 进程（PID {pid}）已经退出。",
            code="EXCEL_CLOSED", stage=Stage.ASSERT,
            hint="Excel 被宏关闭或崩溃了；重新打开目标工作簿后重跑，先查宏里有没有 Quit/Close",
            suggested="", retryable=False,
        )

    dialog = _probe_dialog(pid if pid > 0 else None)
    if dialog is not None and dialog.kind == KIND_COMPILE_ERROR:
        snippet = _dialog_snippet(dialog.text or dialog.title)
        return Failure(
            f"{timeout:g}s 内宏没有返回；检测到 VBE 编译错误弹窗「{dialog.title or snippet}」。",
            code="VBE_COMPILE_ERROR", stage=Stage.VBE,
            hint=f"弹窗标题：{dialog.title or '（读不到标题）'}；正文：{snippet}。"
                 "工程有编译错误（#If/#Else 配对、引用缺失等）会挡住 Application.Run；"
                 "切到 VBE 修掉编译错误后重跑，不要用 SendKeys 发字母",
            suggested="", retryable=False,
        )
    if dialog is not None:
        snippet = _dialog_snippet(dialog.text or dialog.title)
        label = ("VBE 模态框" if dialog.kind == KIND_VBE_DIALOG
                 else "Excel 模态框（MsgBox/InputBox 等）")
        return Failure(
            f"{timeout:g}s 内宏没有返回；检测到{label}「{dialog.title or snippet}」。",
            code="WATCHDOG_TIMEOUT", stage=Stage.ASSERT,
            hint=f"弹窗标题：{dialog.title or '（读不到标题）'}；正文：{snippet}。"
                 "宏被交互框卡住：把 MsgBox/InputBox 改成静默参数，或先手动关掉弹窗再重跑",
            suggested="先用 agent logs 看最后心跳；修掉卡住的宏后重跑，或加 --timeout 放宽",
            retryable=False,
        )

    return Failure(
        f"{timeout:g}s 内宏没有返回；已停止等待，但不会强行结束用户已打开的 Excel。"
        "请检查弹窗/死循环，必要时手动按 Esc 或关闭工作簿。",
        code="WATCHDOG_TIMEOUT", stage=Stage.ASSERT,
        suggested="先用 agent logs 看最后心跳；修掉卡住的宏后重跑，或加 --timeout 放宽",
        retryable=False,
    )


def _run_assert_macro(runner, macro: str):
    """先零参；参数个数不符再退 (含性能, 静默) 两参（口径同 1.5.6）。

    - 第一次非参数个数错误：原样上抛，不试探两参；
    - 两次都是参数个数不符：报 ``ASSERT_MACRO_ARGS``；
    - 第二次是其它错误：原样上抛，不把第一次的 ``ARG_MISMATCH`` 当最终结论。
    """
    try:
        return runner.run(macro, isolate=False, save=False, log_optional=True)
    except BaseFailure as first:
        first = _remap_attach_failure(first)
        if first.code != "ARG_MISMATCH":
            raise

    try:
        return runner.run(macro, args=[ASSERT_PERF, ASSERT_SILENT],
                          isolate=False, save=False, log_optional=True)
    except BaseFailure as second:
        second = _remap_attach_failure(second)
        if second.code == "ARG_MISMATCH":
            raise Failure(
                f"宏 {macro} 的参数个数不符：零参调用和 (含性能, 静默) 两参调用都失败。",
                code="ASSERT_MACRO_ARGS", stage=Stage.ASSERT,
                hint="断言宏要么声明为零参，要么接受 (含性能 As Boolean, 静默 As Boolean)；"
                     "别把两个业务参数当成含性能/静默。",
                suggested="", retryable=False,
            ) from second
        raise


def _make_assert_runner(path: Path, open_hint: str, open_cmd: str):
    """在调用线程里重新附着目标簿，返回可执行宏的 runner。

    真实 assert 的看门狗线程不能复用主线程的 COM 代理，所以在 worker 内
    ``CoInitialize`` 后按路径重新 attach；找不到目标簿仍走 TARGET_NOT_OPEN，
    不把断开的代理硬塞给 MacroRunner。
    """
    xl2 = excel_gateway.attach_running_excel(path, target_hint=open_hint,
                                             target_suggested=open_cmd)
    wb2, opened2 = excel_gateway.find_open_workbook(xl2, path)
    if wb2 is None:
        lines = [f"目标工作簿没打开：{path}"]
        lines += [f"  当前打开：{p}" for p in opened2]
        raise Failure("\n".join(lines), code="TARGET_NOT_OPEN", stage=Stage.ATTACH,
                      hint=open_hint, suggested=open_cmd)
    _ensure_active(xl2, wb2)
    session = WorkbookSession(target=str(path), state=TargetState.OPEN_CLEAN,
                              purpose=Purpose.ASSERT, workbook=wb2)
    return MacroRunner(session)


def _run_assert_watched(path: Path, macro: str, timeout: float,
                        open_hint: str, open_cmd: str, *,
                        runner_factory=None):
    """在 daemon worker 里执行断言宏，主线程只等 ``timeout``。

    ``runner_factory`` 是单测缝：默认走真实 attach；传入桩时不开 Excel。
    worker 内显式 ``CoInitialize`` / ``CoUninitialize``，宏超时后不 join 慢线程，
    与 watchdog 原语“只停止等待、不强行结束 Excel”一致。
    """
    def worker():
        pythoncom = None
        try:
            try:
                import pythoncom as _pythoncom  # type: ignore
                _pythoncom.CoInitialize()
                pythoncom = _pythoncom
            except Exception:  # noqa: BLE001 - 无 pywin32 时仍允许纯桩测试
                pythoncom = None
            factory = runner_factory or _make_assert_runner
            runner = factory(path, open_hint, open_cmd)
            return _run_assert_macro(runner, macro)
        finally:
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:  # noqa: BLE001 - 收尾失败不覆盖宏结果
                    pass

    return watchdog.run_with_watchdog(worker, timeout)


def handle_assert(ctx) -> Outcome:
    args = ctx.args
    proj, path = _pick(ctx, getattr(args, "project", None))
    macro, sheet = _assert_settings(ctx, proj, args)

    open_hint = ("先打开目标工作簿，或用 agent.py book --action open "
                 f"--project {proj['name']} --apply；不要用 Start-Process 绕过工具开表")
    open_cmd = entry_command("agent.py", "book", "--action", "open",
                             "--project", proj["name"], "--apply")

    xl = excel_gateway.attach_running_excel(path, target_hint=open_hint,
                                            target_suggested=open_cmd)
    wb, opened = excel_gateway.find_open_workbook(xl, path)
    if wb is None:
        lines = [f"「{proj['name']}」没打开：{path}"]
        lines += [f"  当前打开：{p}" for p in opened]
        raise Failure("\n".join(lines), code="TARGET_NOT_OPEN", stage=Stage.ATTACH,
                      hint=open_hint, suggested=open_cmd)

    if _macro_exists(wb, macro) is False:
        raise Failure(f"目标工作簿里没有断言宏：{macro}",
                      code="ASSERT_MACRO_NOT_CONFIGURED", stage=Stage.ASSERT,
                      hint="assert 默认模板只认 测试表格工具.运行全部断言；"
                           "非模板工程要在 mapping 配 assert_macro，或用 --macro 临时覆盖",
                      suggested=entry_command("agent.py", "assert",
                                               "--project", proj["name"],
                                               "--macro", "模块.过程", "--apply"))

    _ensure_active(xl, wb)

    human: list[str] = []
    if getattr(args, "clear", False):
        human.append(f"测试表（{sheet}）行数：{_clear_testsheet(wb, args.baseline, sheet)}")

    folder = Path(args.log) if getattr(args, "log", "") else _log_dir_of(proj)
    before = {f.name for f in folder.glob("*.log")} if folder.is_dir() else set()

    t0 = time.time()
    deadline = t0 + args.timeout if args.timeout > 0 else None
    try:
        if getattr(ctx, "runner", None) is not None:
            # 单测/注入 runner 走同步路径，保留 timed_out 桩语义。
            result = _run_assert_macro(ctx.runner, macro)
        else:
            # 真实路径必须走线程看门狗；不能再直接同步 Application.Run。
            result = _run_assert_watched(path, macro, args.timeout,
                                         open_hint, open_cmd)
    except watchdog.WatchdogTimeout:
        raise _classify_watchdog_timeout(xl, timeout=args.timeout)

    if getattr(result, "timed_out", False):
        # 兼容旧 runner 桩；真实路径由上面的 WatchdogTimeout 分支接管。
        raise _classify_watchdog_timeout(xl, timeout=args.timeout)

    human.append(f"Run 结束，用时 {time.time() - t0:.1f}s")

    # 宏已经返回，日志应在此前写完；只给 5 秒收尾，不把剩余总预算全耗在空等上。
    log_timeout = min(5.0, max(0.1, deadline - time.time())) if deadline else 5.0
    log = _wait_new_log(folder, before, log_timeout)
    if log is None:
        raise Failure(
            f"宏已返回，但 {log_timeout:.1f}s 内没有新日志（{folder}）",
            code="ASSERT_NO_LOG", stage=Stage.ASSERT,
            hint="确认断言宏会写日志（errPopup=False，写到 测试输出）；"
                 "只写工作表、不写日志的零参自测宏不适用 agent assert，改用 sync.py run",
            suggested=entry_command("sync.py", "run", "--project", proj["name"],
                                     "--macro", macro, "--apply"))

    text = _read_log_text(log)
    passed = PASS_RE.search(text)
    failed = FAIL_RE.search(text)
    human.append(f"日志：{log}")

    if passed and failed:
        n_pass, n_fail = int(passed.group(1)), int(failed.group(1))
        human.append(f"结论：通过 {n_pass} / 失败 {n_fail}")
        for line in text.splitlines():
            if re.search(r"失败|错误|Err \d", line) and "失败 0" not in line:
                human.append(f"  {line.strip()[:160]}")
        errors: list[Failure] = []
        if n_fail != 0:
            errors.append(Failure(
                f"断言失败 {n_fail} 项（通过 {n_pass}）",
                code="ASSERT_FAILED", stage=Stage.ASSERT,
                hint="先跑 agent logs 看失败行；修掉断言后重跑",
                suggested=entry_command("agent.py", "assert",
                                         "--project", proj["name"], "--apply"),
                retryable=False,
            ))
        return Outcome(
            action="agent.assert", ok=not errors, details=[], errors=errors,
            human=human,
            stats={"lines": human, "passed": n_pass, "failed": n_fail,
                   "macro": macro, "timings": ctx.stage_timings()},
            exit_code=0 if not errors else 1, elapsed_ms=ctx.elapsed_ms(),
        )

    human.append("日志里没找到「通过/失败」计数，最后 30 行如下：")
    for line in text.splitlines()[-30:]:
        human.append(f"  {line.strip()[:160]}")

    if any("没有找到名为 测试表" in line for line in human + [text]):
        code = "ASSERT_NO_TABLE"
        message = "工作簿没有配置的测试表，断言未执行"
    else:
        code = "ASSERT_FAILED"
        message = "日志里读不出通过/失败计数"
    raise Failure(
        message, code=code, stage=Stage.ASSERT,
        hint="先跑 agent logs 看最后心跳与日志尾部；确认测试表存在、断言宏把计数写进日志",
        suggested=entry_command("agent.py", "logs", "--project", proj["name"]),
        retryable=False,
    )


# ---------------------------------------------------------------------------
# patch：整行块替换（格式 / 消歧 / 源码缺失三档语义原样保留）
# ---------------------------------------------------------------------------
def _parse_patch_file(ctx, patch_file: Path) -> list[tuple[Path, str, str]]:
    """[(@@ 文件, 旧, 新)]；格式错就整批拒绝，不做半途替换（口径同 1.5.6）。"""
    groups: list[tuple[Path, str, str]] = []
    cur_file: Path | None = None
    old: list[str] | None = None
    new: list[str] | None = None

    def flush() -> None:
        nonlocal old, new
        if old is None:
            return
        if cur_file is None:
            raise Failure(
                "补丁里出现 -/+ 行却没有 `@@ 文件` 行",
                code="PATCH_FORMAT", stage=Stage.PATCH,
                hint="补丁第一行必须是 `@@ 相对路径`，下面才是成组的 -旧行 / +新行",
                suggested=entry_command("agent.py", "patch", "--file", str(patch_file)))
        if new is None:
            raise Failure(
                f"{rel_path(cur_file)} 的旧行后面没有 + 新行",
                code="PATCH_FORMAT", stage=Stage.PATCH,
                hint="每个 -旧块后面必须紧跟至少一行 +新行；不要只给旧行",
                suggested=entry_command("agent.py", "patch", "--file", str(patch_file)))
        groups.append((cur_file, "\n".join(old), "\n".join(new)))
        old = new = None

    for raw in patch_file.read_text(encoding="utf-8").splitlines():
        if raw.startswith("@@ "):
            flush()
            target = raw[3:].strip()
            cur_file = Path(target) if Path(target).is_absolute() else ctx.root / target
        elif raw.startswith("\\-") or raw.startswith("\\+"):
            # 转义：源码行本身以 -/+ 开头时，前面加一个反斜杠
            if raw.startswith("\\-"):
                old = (old or []) + [raw[1:]]
            else:
                new = (new or []) + [raw[1:]]
        elif raw.startswith("-"):
            old = (old or []) + [raw[1:]]
        elif raw.startswith("+"):
            new = (new or []) + [raw[1:]]
        elif raw.strip():
            raise Failure(
                f"补丁里有看不懂的行：{raw}",
                code="PATCH_FORMAT", stage=Stage.PATCH,
                hint="补丁只认 `@@ 文件`、`-旧行`、`+新行`；源码行本身以 -/+ 开头时前面加反斜杠",
                suggested=entry_command("agent.py", "patch", "--file", str(patch_file)))
    flush()
    if not groups:
        raise Failure(
            "补丁文件是空的",
            code="PATCH_FORMAT", stage=Stage.PATCH,
            hint="至少写一组 `@@ 文件` + `-旧行` + `+新行`",
            suggested=entry_command("agent.py", "patch", "--file", str(patch_file)))
    return groups


def _locate_block(text: str, block: str) -> list[int]:
    """旧文本的命中位置（按整行对齐，从 0 起的行号列表）。

    不用 str.count：子串匹配不锚行首，缩进 24 空格的行会被 32 空格的行顺带命中，
    “恰好 1 次”的校验因此失真（2026-11-20 踩过）。
    """
    lines = [ln.rstrip("\r") for ln in text.split("\n")]
    blk = [ln.rstrip("\r") for ln in block.split("\n")]
    n = len(blk)
    return [i for i in range(len(lines) - n + 1) if lines[i:i + n] == blk]


def _splice_block(text: str, start: int, old: str, new: str) -> str:
    """把从 start 行起、长度等于 old 行数的整块换成 new（保留文件原有行尾）。"""
    lines = text.split("\n")
    end = start + len(old.split("\n"))
    new_lines = new.split("\n")
    if lines[end - 1].endswith("\r"):          # CRLF 文件：末行补回 \r，别把行尾改坏
        new_lines[-1] = new_lines[-1] + "\r"
    return "\n".join(lines[:start] + new_lines + lines[end:])


def handle_patch(ctx) -> Outcome:
    args = ctx.args
    patch_file = Path(args.file)
    if not patch_file.is_absolute():
        patch_file = ctx.root / patch_file
    groups = _parse_patch_file(ctx, patch_file)

    texts: dict[Path, str] = {}
    plan: list[tuple[Path, str, str]] = []
    for path, old, new in groups:
        if path not in texts:
            if not path.is_file():
                raise Failure(
                    f"源码不存在：{path}",
                    code="PATCH_FILE_NOT_FOUND", stage=Stage.PATCH,
                    hint="检查 `@@` 后面的相对路径按仓库根解析；文件删了就先恢复或改清单",
                    suggested=entry_command("agent.py", "patch", "--file", str(patch_file)))
            texts[path] = path.read_text(encoding="utf-8")
        hits = _locate_block(texts[path], old)
        if len(hits) != 1:
            hit_text = "、".join(f"第 {i + 1} 行" for i in hits[:5]) \
                + ("…" if len(hits) > 5 else "")
            raise Failure(
                f"{rel_path(path)}：这段旧文本按整行命中 {len(hits)} 次（要求恰好 1 次），整批未落盘"
                + (f"\n命中位置：{hit_text}" if hit_text else "")
                + f"\n{old}",
                code="PATCH_BLOCK_AMBIGUOUS", stage=Stage.PATCH,
                hint="旧行必须整行唯一、缩进全一致；先 Grep 定位，再只改一处",
                suggested=entry_command("agent.py", "patch", "--file", str(patch_file)))
        texts[path] = _splice_block(texts[path], hits[0], old, new)
        plan.append((path, old, new))

    human = [f"  ✓ {rel_path(path)}：{old.splitlines()[0][:48]}…"
             for path, old, new in plan]

    if not getattr(args, "apply", False):
        human.append(f"dry-run：{len(plan)} 处全部可命中，未写入")
        return Outcome(
            action="agent.patch", ok=True,
            details=[{"dry": True, "cmd": "patch", "changes": len(plan)}],
            human=human,
            stats={"lines": human, "dry": True, "changes": len(plan),
                   "timings": ctx.stage_timings()},
            exit_code=0, elapsed_ms=ctx.elapsed_ms(),
        )

    for path, text in texts.items():
        write_source(path, text)
        human.append(f"  写入 {rel_path(path)}")

    errors: list[Failure] = []
    if not getattr(args, "no_style", False):
        files = source_files(ctx.root / "src")
        with ctx.progress.stage("patch-style-fix"):
            style_fix = run_fix(files, "all", True)
        with ctx.progress.stage("patch-style-check"):
            style_check = run_checks(files, "all")
        if style_fix or style_check:
            errors.append(Failure(
                "补丁落盘后的风格复检未通过",
                code="SOURCE_PREFLIGHT", stage=Stage.PATCH,
                hint="按 style 清单修复后重跑；补丁本身已落盘，仅风格复检不通过",
                suggested=entry_command("agent.py", "style", "--action", "fix", "--apply")))

    return Outcome(
        action="agent.patch", ok=not errors,
        details=[{"changes": len(plan), "applied": True}],
        errors=errors, human=human,
        stats={"lines": human, "changes": len(plan), "applied": True,
               "timings": ctx.stage_timings()},
        exit_code=0 if not errors else 1, elapsed_ms=ctx.elapsed_ms(),
    )


# ---------------------------------------------------------------------------
# style：风格检查 / 修复（规则在 toolkit.style_rules，本层只编排与收口）
# ---------------------------------------------------------------------------
def _style_fix_suggested(which: str, project: str, scope: str) -> str:
    """修复建议带上本次的 --which / 范围，避免默认 all + src 扫错对象。"""
    parts = ["style", "--action", "fix", "--which", which]
    if project:
        parts.extend(["--project", project])
    elif scope != "src":
        parts.extend(["--scope", scope])
    parts.append("--apply")
    return entry_command("agent.py", *parts)


def handle_style(ctx) -> Outcome:
    args = ctx.args
    action = args.action
    which = args.which
    details: list[dict] = []

    with ctx.progress.stage("style"):
        if getattr(args, "project", None):
            proj, _path = _pick(ctx, args.project)
            common_dir = _defaults(ctx).get("common_dir", "src/common")
            files = (source_files(_root_path(ctx, proj.get("src", "")))
                     + source_files(_root_path(ctx, common_dir)))
            if not files:
                warning = f"{args.project} 下没有找到 .bas/.cls/.frm"
                return Outcome(
                    action="agent.style", ok=True,
                    warnings=[warning], human=[warning],
                    stats={"lines": [warning], "action": action, "which": which,
                           "files": 0, "timings": ctx.stage_timings()},
                    exit_code=0, elapsed_ms=ctx.elapsed_ms(),
                )
        else:
            files = _scope_files(ctx, args.scope)
            if not files:
                warning = f"{args.scope} 下没有找到 .bas/.cls/.frm"
                return Outcome(
                    action="agent.style", ok=True,
                    warnings=[warning], human=[warning],
                    stats={"lines": [warning], "action": action, "which": which,
                           "files": 0, "timings": ctx.stage_timings()},
                    exit_code=0, elapsed_ms=ctx.elapsed_ms(),
                )

        out_buf = io.StringIO()
        err_buf = io.StringIO()
        # run_checks/run_fix 走 toolkit.base.msg：JSON 模式下 msg 固定写 stderr，
        # 只 redirect_stdout 会让人读报告全部漏到 stderr、stats.lines 反而为空。
        # 两个流一起截住，报告统一进 human / stats.lines，JSON 出口保持干净。
        with contextlib.redirect_stdout(out_buf), contextlib.redirect_stderr(err_buf):
            if action == "check":
                code, details = run_checks_with_details(files, which)
            else:
                code = run_fix(files, which, bool(getattr(args, "apply", False)))

    out_lines = _human_lines(out_buf)
    err_lines = _human_lines(err_buf)
    # N5：报告行（stdout）与警告（stderr）分两条通道。
    # · [错误] 行留在人话报告里；[警告] 行剥掉统一前缀后进 Outcome.warnings，
    #   render 再按人话/JSON 口径输出到 stderr 或 stats.warnings；
    # · traceback 等杂项不混进报告，避免污染单行 JSON 信封。
    warnings: list[str] = []
    human = list(out_lines)
    for line in err_lines:
        if _STDERR_JUNK_RE.match(line):
            if _STDERR_ERROR_RE.match(line):
                human.append(line)
                continue
            body = _WARN_PREFIX_RE.sub("", line).strip()
            if body:
                warnings.append(body)
            continue
        human.append(line)
    errors: list[Failure] = []
    if code != 0:
        # 1.5.6 的 style 失败没有独立登记码（旧信封兜底 FAILED，1.6 已移除 FAILED）；
        # 这里归到源码检查族的登记码 SOURCE_PREFLIGHT，提示与建议仍按 style 口径给。
        errors.append(Failure(
            "风格检查发现不合规项，详见报告清单",
            code="SOURCE_PREFLIGHT", stage=Stage.CONFIG,
            hint="按下方 details / 预览清单手工修，或按 suggested 原样执行修复并复检",
            suggested=_style_fix_suggested(which, args.project, args.scope),
        ))

    return Outcome(
        action="agent.style", ok=not errors, details=details, errors=errors,
        warnings=warnings,
        stats={"lines": human, "action": action, "which": which, "files": len(files),
               "timings": ctx.stage_timings()},
        human=human, exit_code=0 if not errors else 1,
        elapsed_ms=ctx.elapsed_ms(),
    )


SPECS: tuple[CommandSpec, ...] = (
    CommandSpec(
        verb="assert", entry="agent", handler="features.agent_patch:handle_assert",
        mutation=MutationMode.REQUIRED, project=ProjectMode.OPTIONAL,
        args=(
            ArgSpec("--macro", default="", help="断言宏「模块.过程」；缺省取 mapping/默认模板"),
            ArgSpec("--sheet", default="", help="测试表名；缺省取 mapping/测试表"),
            ArgSpec("--clear", kind="flag", help="跑之前把测试表清到基线行数"),
            ArgSpec("--baseline", kind="int", default=35, help="清表保留的基线行数"),
            ArgSpec("--timeout", kind="float", default=120, help="运行期硬超时（秒）"),
            ArgSpec("--log", default="", help="日志目录；缺省取项目 测试输出"),
        ),
        supports_json=True, exit_codes=(0, 1, 2),
        summary="运行断言宏并核对通过/失败计数（看门狗硬超时）",
        ref="references/agent.md",
        tests=("contract.assert.success", "contract.assert.no_table",
               "contract.assert.no_log", "contract.assert.timeout",
               "contract.assert.usage", "contract.assert.json"),
    ),
    CommandSpec(
        verb="patch", entry="agent", handler="features.agent_patch:handle_patch",
        mutation=MutationMode.REQUIRED, project=ProjectMode.OPTIONAL,
        args=(
            ArgSpec("--file", required=True, help="补丁清单文件"),
            ArgSpec("--no-style", kind="flag", help="落盘后跳过风格修复/复检"),
        ),
        supports_json=True, exit_codes=(0, 1, 2),
        summary="按整行块清单批量替换源码（dry-run 只核对命中）",
        ref="references/agent.md#按清单打补丁patch",
        tests=("contract.patch.success", "contract.patch.format",
               "contract.patch.ambiguous", "contract.patch.json"),
    ),
    CommandSpec(
        verb="metrics", entry="agent", handler="features.agent_patch:handle_metrics",
        mutation=MutationMode.CONDITIONAL, project=ProjectMode.NONE,
        args=(
            ArgSpec("--out", default="", help="把基线表写到该 Markdown 文件（落盘要 --apply）"),
        ),
        supports_json=True, exit_codes=(0, 1, 2),
        summary="源码体检基线（文件/过程/长度/薄包装/模块变量/布尔/On Error 规则）",
        ref="references/agent.md#源码体检基线metrics",
        tests=("contract.metrics.success", "contract.metrics.usage",
               "contract.metrics.json"),
    ),
    CommandSpec(
        verb="style", entry="agent", handler="features.agent_patch:handle_style",
        mutation=MutationMode.PREVIEW, project=ProjectMode.OPTIONAL,
        args=(
            ArgSpec("--action", kind="choice", choices=("check", "fix"), default="check",
                    help="check 只读检查；fix 预览修复（落盘要 --apply）"),
            ArgSpec("--which", kind="choice", choices=("all", "eol", "call", "explicit"),
                    default="all", help="只查哪一类：all / eol / call / explicit"),
            ArgSpec("--scope", default="src", help="范围（相对仓库根），默认整个 src"),
        ),
        supports_json=True, exit_codes=(0, 1, 2),
        summary="风格检查/修复（check 只读，fix 要 --apply）",
        ref="references/agent.md#风格检查--修复style-子命令",
        tests=("contract.style.success", "contract.style.failure",
               "contract.style.usage", "contract.style.json"),
    ),
)
