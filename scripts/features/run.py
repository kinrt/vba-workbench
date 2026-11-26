# -*- coding: utf-8 -*-
"""run 切片：run / watch 与 MacroRunner 的编排（动词入口薄层）。

职责边界：
  · ``run``：参数解析 →（可选）源码预检/备份/写回/回读校验 → ``MacroRunner``
    跑宏 → 取返回值与日志尾部 → 落 run journal；宏执行装配与结果汇总在
    ``run_macro`` 一份实现。
  · ``watch``：常驻监听；目标解析/探活/比对/写回走 toolkit 与 adapter 真实能力，
    只在 COM 边界经过 ``adapters.excel_gateway``；监听辅助在 ``run_watch``。
    ``watch`` 不支持 ``--json``。

本层不自己 new COM、不 subprocess 回壳、不另拼 JSON 信封；输出由 ``render`` 出口渲染。
"""
from __future__ import annotations

import datetime as dt
import time
from typing import Any

from adapters import excel_gateway  # noqa: F401  # 既有模块属性，测试桩直接引用
from toolkit import excel, files, ui, watch as watch_toolkit
from toolkit.base import abs_path, interactive_ok, msg, warn
from toolkit.contract import (ErrorCode, Failure, MutationMode, Outcome, Stage)
from toolkit.registry import ArgSpec, CommandSpec, ProjectMode

from features.run_macro import (  # noqa: F401  # 入口层重导出，既有符号契约不变
    _acquire,
    _common_subset,
    run,
)
from features.run_watch import (
    WATCH_SAME,
    _compare_set,
    _watch_compare,
    _watch_log_path,
    _watch_pull_one,
    _watch_refresh_baseline,
    _watch_targets,
    _watch_write,
)

# 模块级睡眠钩子：测试可 monkeypatch 避免真睡；生产走 time.sleep。
_sleep = time.sleep


# --------------------------------------------------------------------------- 入口闸门
def json_rejected(args: Any) -> bool:
    """常驻命令带 --json：解析层已置 ``_reject_json``，这里做派发前复核。"""
    return bool(getattr(args, "json", False)) and bool(getattr(args, "_reject_json", False))


def precheck(args: Any) -> int | None:
    """入口派发前的用法闸门；返回非 None 时入口按该退出码结束，不渲染 JSON。"""
    if json_rejected(args):
        return 2
    return None


def _usage(message: str) -> Failure:
    return Failure(message, code=ErrorCode.USAGE, stage=Stage.CONTRACT,
                   hint="核对命令行参数；加 --help 查看用法。")


# --------------------------------------------------------------------------- watch 主循环辅助（钩子消费方留在本层）
def _watch_probe(target: dict, tries: int = 3, gap: float = 0.4) -> str:
    """探活：alive / closed / busy（读不到 Workbooks 不算关闭，按忙处理）。"""
    book = target.get("book") or target["name"]
    xl = target.get("xl")
    wb = target["wb"]
    for _ in range(max(tries, 1)):
        try:
            if str(wb.Name) == book:
                return "alive"
        except Exception:
            pass
        try:
            books = list(xl.Workbooks)
        except Exception:
            _sleep(gap)
            continue
        for b in books:
            try:
                if str(b.Name) == book:
                    return "alive"
            except Exception:
                continue
        return "closed"
    return "busy"


def _watch_ask_one(file: Any, verdict: str, src_n: int, vbe_n: int) -> str:
    """逐个决定里的单模块：返回 src / vbe / skip；非交互一律 skip（不静默卡死）。"""
    if not interactive_ok():
        return "skip"
    msg(f"  {file.name}   {verdict}：源码 {src_n} 行 / 工程 {vbe_n} 行")
    pick = ui.choose("    以哪边为准？[1] 源码  [2] 表内  [3] 跳过：", 3,
                     exit_hint="跳过这一个")
    return {1: "src", 2: "vbe"}.get(pick, "skip")


# --------------------------------------------------------------------------- watch 动词
def watch(ctx: Any) -> Outcome:
    """常驻监听：源码一变就写回已打开的工程；不支持 --json。"""
    args = ctx.args
    gate = precheck(args)
    if gate is not None:
        # 用法错走 exit 2 闸门：只出人话，不渲染 JSON 信封（契约见 SKILL.md 退出码表）。
        usage = _usage("watch 是常驻命令，不支持 --json；机器读前置条件用 "
                        "plan --intent watch --json")
        outcome = Outcome(action="sync.watch", ok=False, exit_code=gate,
                          errors=[usage],
                          human=[f"[错误] {usage}", f"说明：{usage.hint}"],
                          elapsed_ms=ctx.elapsed_ms())
        return outcome

    cfg = ctx.config
    project = getattr(args, "project", None)
    interval = float(getattr(args, "interval", 1.0) or 1.0)
    save = not bool(getattr(args, "no_save", False))
    backup = not bool(getattr(args, "no_backup", False))
    doc_modules = bool(getattr(args, "doc_modules", False))
    startup = getattr(args, "startup", "ask")
    seconds = int(getattr(args, "seconds", 0) or 0)
    dry = not bool(getattr(args, "apply", False))
    defaults = cfg.get("defaults", {})
    common_dir = abs_path(defaults.get("common_dir", "src/common"))

    targets = _watch_targets(cfg, project)
    label = "全部表" if len(targets) > 1 else str(targets[0]["name"])
    log_path = _watch_log_path(ctx, label)

    def record(text: str, level: str = "note") -> None:
        stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        files.append_line(log_path, f"{stamp} {text}")
        if level == "problem":
            warn(f"[{stamp}] {text}")

    if backup and not dry:
        backup_root = abs_path(defaults.get("backup_dir", ".backup"))
        for target in targets:
            excel.backup_workbook(target["wb"], backup_root, tag="watch前")

    # 首轮基线与启动核对。
    written = skipped = 0
    for target in targets:
        target["baseline"] = watch_toolkit.snap(target["folders"])

    if not dry and startup != "skip":
        for target in targets:
            rows = _watch_compare(target["wb"], _compare_set(target, common_dir))
            diffs = [r for r in rows if r[1] != WATCH_SAME]
            if not diffs:
                continue

            if startup == "src":
                for file, _verdict, _src_n, _vbe_n in diffs:
                    ok, text = _watch_write(target["wb"], file, target["folders"],
                                            common_dir, target["subset"],
                                            doc_modules, save, dry)
                    record(text, "problem" if not ok else "note")
                    if ok:
                        written += 1
                    else:
                        skipped += 1
                continue

            if startup == "vbe":
                for file, _verdict, _src_n, _vbe_n in diffs:
                    path = _watch_pull_one(target["wb"], file, target["folders"], dry)
                    if path is None:
                        record(f"跳过 {file.name}：工程里没有这个模块", "problem")
                        continue
                    record(f"抓回源码 {path.name}（以表内为准）")
                    if not dry:
                        _watch_refresh_baseline(
                            targets, path, file if path != file else None)
                continue

            if startup == "each":
                if not interactive_ok():
                    record("非交互/--no-input：逐个询问不可用，两边先保持现状", "problem")
                    continue
                for file, verdict, src_n, vbe_n in diffs:
                    side = _watch_ask_one(file, verdict, src_n, vbe_n)
                    if side == "skip":
                        record(f"跳过 {file.name}：本次未选方向")
                        continue
                    if side == "vbe":
                        path = _watch_pull_one(target["wb"], file, target["folders"], dry)
                        if path is None:
                            record(f"跳过 {file.name}：工程里没有这个模块", "problem")
                            continue
                        record(f"抓回源码 {path.name}（以表内为准）")
                        if not dry:
                            _watch_refresh_baseline(
                                targets, path, file if path != file else None)
                        continue
                    ok, text = _watch_write(target["wb"], file, target["folders"],
                                            common_dir, target["subset"],
                                            doc_modules, save, dry)
                    record(text, "problem" if not ok else "note")
                    if ok:
                        written += 1
                    else:
                        skipped += 1

    t0 = dt.datetime.now()
    human: list[str] = [f"日志：{log_path}"]

    try:
        while True:
            if seconds and (dt.datetime.now() - t0).total_seconds() >= seconds:
                record(f"到 --seconds {seconds} 上限，停止监听")
                human.append(f"到 --seconds {seconds} 上限，停止监听")
                break
            _sleep(max(interval, 0.2))

            alive: list[dict] = []
            closed: list[str] = []
            busy: list[str] = []
            for target in targets:
                state = _watch_probe(target)
                if state == "closed":
                    closed.append(str(target.get("name", "?")))
                    continue
                if state == "busy":
                    busy.append(str(target.get("name", "?")))
                alive.append(target)
            for name in closed:
                record(f"「{name}」的工作簿已关闭（Excel 里已经没有这个工作簿），不再监听",
                       "problem")
            targets = alive
            if not targets:
                record(f"要盯的表都没了，停止监听：{'、'.join(closed)}", "problem")
                break
            if busy:
                record(f"Excel 这一轮读不到工作簿列表（忙或被弹窗挡住）："
                       f"{'、'.join(busy)} —— 本轮跳过，继续监听", "problem")
                continue

            for target in targets:
                current = watch_toolkit.snap(target["folders"])
                for file in watch_toolkit.changed(current, target["baseline"],
                                                  target["pending"]):
                    ok, text = _watch_write(target["wb"], file, target["folders"],
                                            common_dir, target["subset"],
                                            doc_modules, save, dry)
                    target["baseline"][file] = current[file]
                    if ok:
                        written += 1
                        record(text)
                    else:
                        skipped += 1
                        record(text, "problem")
    except KeyboardInterrupt:
        pass

    span = int((dt.datetime.now() - t0).total_seconds())
    record(f"已停止监听：写回 {written} 个，跳过/失败 {skipped} 个，运行 {span} 秒")
    human.append(f"已停止监听：写回 {written} 个，跳过/失败 {skipped} 个，运行 {span} 秒")
    stats = {"targets": len(targets), "written": written, "skipped": skipped,
             "seconds": span, "log": str(log_path), "startup": startup}
    return Outcome(action="sync.watch", ok=True, details=[], stats=stats, human=human,
                   exit_code=0, elapsed_ms=ctx.elapsed_ms())


# --------------------------------------------------------------------------- 命令登记
SPECS: tuple[CommandSpec, ...] = (
    CommandSpec(
        verb="run", entry="sync", handler="features.run:run",
        mutation=MutationMode.CONDITIONAL, project=ProjectMode.REQUIRED,
        supports_json=True, exit_codes=(0, 1, 2),
        args=(
            ArgSpec("--macro", kind="str", required=True,
                    help="要运行的宏（可写 模块.宏）"),
            ArgSpec("--arg", kind="append", default=[],
                    help="传给宏的参数，可重复；自动认 True/False/整数/小数，其余按字符串"),
            ArgSpec("--push", kind="flag", help="先写回源码再运行"),
            ArgSpec("--isolate", kind="flag",
                    help="自己开隐藏 Excel 跑（目标表必须未打开；不触发 Workbook_Open）"),
            ArgSpec("--save", kind="flag", help="隔离模式下运行完保存工作簿"),
            ArgSpec("--tail", kind="int", default=40,
                    help="日志尾部行数（默认 40，0=不打）"),
            ArgSpec("--log", kind="str", help="结果日志：目录或文件（默认 表目录\\测试输出）"),
            ArgSpec("--log-optional", kind="flag", dest="log_optional",
                    help="该宏本来不写日志：取不到日志按正常处理，不提示像失败"),
            ArgSpec("--rollback-on-fail", kind="flag", dest="rollback_on_fail",
                    help="--push 的 postcheck 失败时，按写回前内存正文复原本次模块"),
        ),
        summary="写回（可选）→运行宏→取返回值与日志尾部",
        ref="references/sync-commands.md#run",
        tests=("contract.run.success", "contract.run.failed", "contract.run.usage",
               "contract.run.json"),
    ),
    CommandSpec(
        verb="watch", entry="sync", handler="features.run:watch",
        mutation=MutationMode.CONTINUOUS, project=ProjectMode.OPTIONAL,
        supports_json=False, exit_codes=(0, 1, 2),
        args=(
            ArgSpec("--interval", kind="float", default=1.0,
                    help="轮询间隔（默认 1.0）"),
            ArgSpec("--startup", kind="choice",
                    choices=("ask", "src", "vbe", "each", "skip"), default="ask",
                    help="启动核对发现不一致时：ask 逐个问 / src 以源码为准 / "
                         "vbe 以表为准 / each 逐个决定 / skip 先不动"),
            ArgSpec("--no-save", kind="flag", dest="no_save",
                    help="写回后不保存，只改内存"),
            ArgSpec("--no-backup", kind="flag", dest="no_backup",
                    help="启动时不备份整表"),
            ArgSpec("--doc-modules", kind="flag", dest="doc_modules",
                    help="连文档模块一起覆盖"),
            ArgSpec("--seconds", kind="int", default=0,
                    help="最多监听 N 秒（0=一直盯，Ctrl+C 停）"),
        ),
        summary="常驻监听：源码一变就写回已打开的工程",
        ref="references/sync-commands.md#watch",
        tests=("contract.watch.success", "contract.watch.failed",
               "contract.watch.usage", "contract.watch.reject_json"),
    ),
)
