# vba-workbench 架构与文件地图

> 本文是 README 的深入附录：讲调用链、分层和每个 Python 文件的功能。
> 日常用法和第一次上手见仓库根 `README.md`（不随安装包交付）；agent 的操作入口看 [../SKILL.md](../SKILL.md)。

---

## 1. 一句话架构

这是一个**命令清单驱动（manifest-driven）的命令行工具**：

```
sync.py / agent.py / audit.py
        │
        ▼
toolkit.commands.specs_for(entry)
        │  每个 features 切片用 SPECS 登记“动词 → handler”
        ▼
toolkit.registry.build_parser
        │  从 manifest 生成 argparse
        ▼
toolkit.context.AppContext.build
        │  加载 mapping.json、设置 JSON/计时上下文
        ▼
toolkit.registry.dispatch
        │  按 handler 字符串导入函数
        ▼
features.<切片>.<handler>(ctx)
        │  业务编排；返回 Outcome
        ▼
toolkit.render
        │  JSON 信封或可读文本；统一退出码
        ▼
exit code
```

关键点：

* 入口脚本**不写业务特例**，只负责参数、上下文、异常兜底和渲染。

* 业务动词集中在 `scripts/features/`，每个切片自己登记 `SPECS`。

* 通用能力集中在 `scripts/toolkit/`，包括配置、仓库、文件、Excel COM、宏执行、风格检查和输出。

* `scripts/adapters/excel_gateway.py` 是 COM 边界的门面；真正的实现在 `toolkit/excel.py`。

* 所有命令默认 “只读 / 预览”，只有写类动词才认 `--apply`。

***

## 2. 五层结构

| 层        | 目录                                                        | 作用                                                       | 是否运行时必需     |
| -------- | --------------------------------------------------------- | -------------------------------------------------------- | ----------- |
| 入口层      | `scripts/sync.py`、`agent.py`、`audit.py`                   | 三个用户面：同步、Agent、盘点                                        | 是           |
| 命令登记层    | `toolkit/commands.py`、`registry.py`、`cli.py`、`context.py` | 汇总 SPECS、生成 argparse、建上下文、异常兜底                           | 是           |
| 业务切片层    | `scripts/features/`                                       | 29 个文件，按命令族拆：core /read/push /transform/run /agent/audit | 是           |
| 公共基础函数层 | `scripts/toolkit/` | 26 个文件：配置、repo、文件、Excel、宏、检查、风格、VBA 解析、UI、输出 | 是 |
| 外部边界层    | `scripts/adapters/`                                       | 2 个文件；Excel COM 网关                                       | 是（COM 相关命令） |
| 开发 / 验证层 | `tests/`、`tests/tools/`、`scripts/make_lab_workbook.py`    | 单测、基线、门禁、发布、造表                                           | 否           |

### 2.1 目录和文件数量

* `scripts/`：61 个 `.py`

  * 3 个入口脚本

  * 29 个 `features` 文件（含 `__init__.py`）

  * 26 个 `toolkit` 文件（含 `__init__.py`）

  * 2 个 `adapters` 文件

  * 1 个开发脚本 `make_lab_workbook.py`

* `tests/`：45 个 `test_*.py`

* `tests/tools/`：7 个开发工具

* `tests/baseline/`：9 个文件（1 个记录脚本 + 8 个基线数据 / 文档）

> 读代码时不要从 `tests/` 开始。测试是规格，不是主链。

***

## 3. 命令面：35 个动词

### 3.1 按切片分组

| 切片                        | 动词                                                  | 职责                              |
| ------------------------- | --------------------------------------------------- | ------------------------------- |
| `features.core`           | `version` / `list` / `doctor` / `scrub`             | 版本核对、项目解析、环境体检、交付件脱敏扫描；不碰 Excel |
| `features.sync_read`      | `vbamod` / `plan` / `pull` / `pull-open`            | 模块清单核对、模式推荐、VBE → src 抓取        |
| `features.sync_diff`      | `diff`                                              | 只读比较源码与工程模块，过程级差异报告             |
| `features.sync_push`      | `init` / `push` / `shells`                          | 初始化配置、src → VBE 写回、公共库占位壳       |
| `features.sync_transform` | `align` / `promote` / `gbk-export` / `clean-backup` | 公共库换壳、模块提升、GBK 副本、备份清理          |
| `features.run`            | `run` / `watch`                                     | 宏运行编排、常驻监听                      |
| `features.agent_io`       | `probe` / `book` / `backup` / `logs` / `push`       | Agent 状态、开关工作簿、备份、日志、写回         |
| `features.agent_patch`    | `assert` / `patch` / `metrics` / `style`            | 断言、补丁、源码体检、风格检查 / 修复            |
| `features.audit_a`        | `members` / `common` / `project`                    | 成员引用、公共库总览、项目依赖矩阵               |
| `features.audit_b`        | `structure` / `indicators` / `overview`             | 结构问题、H01–H19 指标、全仓总览            |
| `features.audit_lint`     | `lint`                                              | 四条高置信静态规则；规则实现在 `toolkit.vba_lint` |
| `features.audit_references` | `references`                                      | 只读引用体检；COM 读取在 `excel_gateway.read_project_references` |

### 3.2 入口与切片的关系

```
sync.py  → sync.*   → core / sync_read / sync_diff / sync_push / sync_transform / run
agent.py → agent.*  → agent_io / agent_patch
audit.py → audit.*  → audit_a / audit_b / audit_lint / audit_references → AuditEngine
```

三个入口**互不 import**。如果某个能力三个入口都要用，应下沉到 `toolkit`，而不是让入口互相调用。

***

## 4. 五条主调用链

### 4.1 `pull`：工作簿 → 源码

```
sync.py pull
  → features.sync_read.pull(ctx)
  → resolve_target(proj)                     # 找 xlsm
  → find_excel_with_workbook / open_disk_workbook
  → _export_workbook
      → _export_component                   # VBE Export，按 comp.Type 补扩展名
      → files.transcode_exported_file       # ANSI/GBK → UTF-8 + LF
  → _write_placeholder_shells               # 公共库模块在项目 src 留一行占位壳
  → check_export_set                        # 工程模块 ↔ 源码文件对账
  → Outcome → render
```

要点：

* 工作簿已打开：借用已打开会话，能抓到未保存改动，不关用户工作簿。

* 工作簿未打开：隐藏实例只读打开，抓完关闭自己打开的副本。

* 默认只预览；`--apply` 才写 `src/`。

* `.frm` 必须和 `.frx` 同目录同名。

### 4.2 `push`：源码 → 工作簿

```
sync.py push
  → features.sync_push.push(ctx)
  → _push_single / _push_projects
  → sync_preflight._source_preflight       # 只读预检，不通过连工作簿都不打开
  → _attach_or_open                        # 借用已打开工作簿 / 自动打开隐藏实例
  → backup_workbook                        # 写前整工作簿备份
  → _writeback
      → 公共模块先写
      → 项目模块后写（项目同名会覆盖公共库）
      → 文档模块 DeleteLines + AddFromString（不能 Remove）
  → _postcheck_sync                        # 回读模块集合与正文
  → AppGuard.restore                       # 还原 Excel 状态
  → wb.Save                                # 按开关保存
  → close_owned                            # 只关自己开的实例
  → Outcome → render
```

要点：

* 写回顺序固定为 `common → project`。

* 文档模块不能 `Remove`，只能清空行后写入。

* 写回前预检硬项：`.cls` 缺 `VERSION 1.0 CLASS`、单语句续行超过 24 行。

* 写回后必须回读校验，失败不保存。

### 4.3 `run`：写回（可选）→ 运行宏 → 取日志

```
sync.py run
  → features.run_macro.run(ctx)
  → _source_preflight（仅 --push）
  → _acquire                              # isolate 自己开 / attach 借会话
  → backup_workbook + _sync_into（仅 --push）
  → _resolve_macro_ref                    # 解析 模块.宏
  → toolkit.macro.MacroRunner.run
      → 活动簿核对
      → AppGuard 快照
      → Application.Run
      → 错误分类：参数 / COM 忙断开 / 业务失败 / 超时
      → 日志收尾
  → journal + Outcome
```

`MacroRunner` 是唯一的宏执行 / 分类入口。`sync.run` 和 `agent.assert` 都复用它，不要另写一套 `Application.Run` 分类器。

### 4.4 `watch`：常驻监听

```
sync.py watch
  → features.run.watch(ctx)
  → _watch_targets                         # 启动时确定监听集合
  → toolkit.watch.snap                     # 记录源码快照
  → 循环：
      → _watch_probe                       # 三态探活：活着 / 已关 / 本轮忙
      → toolkit.watch.changed              # 连续两轮快照一致才处理
      → _watch_write / _watch_pull_one
      → _watch_refresh_baseline
```

要点：

* 只监听启动那一刻已打开且已登记的工作簿。

* 中途打开的新工作簿不会自动纳入。

* “这一轮读不到” 不等于关闭；忙或弹窗挡住只跳过本轮。

### 4.5 `audit`：纯计算 + 报告

```
audit.py structure / indicators / overview
  → features.audit_b
  → features.audit_engine.AuditEngine
      ├─ audit_scan       枚举源码、成员、可达性
      ├─ audit_structure  壳 / 重名 / 绑定 / 文档模块对账
      └─ audit_indicators H01–H19
  → audit_rules           阈值、白名单、正则（唯一来源）
  → 可读文本 / Markdown / Outcome
```

`AuditEngine` 只做纯计算，不打印、不碰 COM、不写文件。排版和退出码由 `audit_a.py` / `audit_b.py` 负责。

### 4.6 `agent`：不是第四套业务，是另一个 API 面

```
agent.py
  ├─ agent_io     probe / book / backup / logs / push
  ├─ agent_patch  assert / patch / metrics / style
  └─ 共用 toolkit.excel / toolkit.macro / toolkit.style_rules
```

注意：`agent.push` 目前有自己的 `_writeback` 实现，和 `sync_push` 属于同一类能力但代码不共用。读的时候把两者当 “同族实现”，不要误以为 `agent.push` 会调用 `sync.push`。

***

## 5. 运行时文件地图

### 5.1 入口层（3 个）

| 文件                 | 作用              | 主要函数                      |
| ------------------ | --------------- | ------------------------- |
| `scripts/sync.py`  | 同步入口；支持 `-i` 菜单 | `main()`、`_interactive()` |
| `scripts/agent.py` | Agent 入口；无交互菜单  | `main()`                  |
| `scripts/audit.py` | 盘点入口；`-i` 目前未接入 | `main()`、`_interactive()` |

### 5.2 命令与上下文层（`toolkit` 中的 5 个核心文件）

| 文件                    | 作用                                                                      | 主要函数 / 类                                                      |
| --------------------- | ----------------------------------------------------------------------- | ------------------------------------------------------------- |
| `toolkit/commands.py` | 按固定顺序汇总 13 个切片的 `SPECS`；重复登记直接报错                                        | `load_specs()`、`specs_for(entry)`                             |
| `toolkit/registry.py` | 定义 `CommandSpec` / `ArgSpec` / `ProjectMode`；生成 argparse；按字符串加载 handler | `build_parser()`、`load_handler()`、`dispatch()`                |
| `toolkit/cli.py`      | 三入口共用骨架：`--version`、`--no-input`、`-i` 把关、异常翻译                           | `run_dispatch()`、`run_cli()`、`entry()`、`ensure_interactive()` |
| `toolkit/context.py`  | 运行上下文：配置、根路径、JSON 模式、阶段计时                                               | `AppContext.build()`、`Progress.stage()`、`stage_timings()`     |
| `toolkit/render.py`   | 唯一输出出口：Outcome → JSON 信封或可读文本                                             | `render()`                                                    |

### 5.3 `features` 业务切片（29 个）

| 文件                             | 作用                                            | 主要函数                                                                                                        |
| ------------------------------ | --------------------------------------------- | ----------------------------------------------------------------------------------------------------------- |
| `features/__init__.py`         | 包标记；按命令族分切片                                   | —                                                                                                           |
| `features/core.py`             | 只读基础命令：版本、项目清单、体检、脱敏扫描                        | `version()`、`list_projects()`、`doctor()`、`scrub()`                                                          |
| `features/sync_read.py`        | VBE → src；模块清单、模式推荐、抓取                        | `vbamod()`、`plan()`、`pull()`、`pull_open()`、`_export_workbook()`                                             |
| `features/sync_diff.py`        | `sync.py diff`：只读比较源码与 VBE 模块，报 `add/update/remove/same` 和过程级差异 | `diff()`                                                                                                    |
| `features/sync_push.py`        | src → VBE；初始化、写回、占位壳入口                        | `push()`、`_push_single()`、`_push_projects()`、`_writeback()`、`_attach_or_open()`                             |
| `features/sync_init.py`        | `init` 动词：生成 mapping.json 和 src 骨架            | `init()`                                                                                                    |
| `features/sync_preflight.py`   | 写回前源码预检；只读、不碰 Excel                           | `_source_preflight()`、`_preflight_or_fail()`、`_missing_class_header()`、`_long_continuation()`               |
| `features/sync_shells.py`      | 公共库占位壳生成 / 清理                                 | `shells()`、`_write_shells()`                                                                                |
| `features/sync_transform.py`   | 换壳、提升模块、GBK 副本、备份清理                           | `align()`、`promote()`、`gbk_export()`、`clean_backup()`                                                       |
| `features/common_subset.py`    | features 内共享叶子：公共库 auto 子集、VBE 快照 / 复原、关闭自有实例 | `resolve_common_subset()`、`snapshot_module_code()`、`restore_module_code()`、`close_owned()`、`find_project()` |
| `features/run.py`              | `run` / `watch` 的入口层；目标探活、监听循环                | `watch()`、`precheck()`、`json_rejected()`；`run` 从 `run_macro` 导入                                             |
| `features/run_macro.py`        | 宏运行编排：预检、备份、写回、执行、日志、journal                  | `run()`、`_acquire()`、`_sync_into()`、`_resolve_macro_ref()`、`_write_journal()`                               |
| `features/run_watch.py`        | watch 的辅助段：目标解析、比对、单文件写回 / 抓回、基线刷新            | `_watch_targets()`、`_watch_compare()`、`_watch_write()`、`_watch_pull_one()`                                  |
| `features/sync_interactive.py` | `sync -i` 菜单本体                                | `interactive_entry()`、`run_menu()`                                                                          |
| `features/agent_io.py`         | Agent IO 薄层：状态、开关簿、备份、日志、写回                   | `book()`、`backup()`、`push()`；内部 `_writeback()`、`_attach_or_open()`                                          |
| `features/agent_probe.py`      | `agent.probe`：项目 / 磁盘 / Excel / 工程 / 模块数一览    | `probe()`                                                                                                   |
| `features/agent_logs.py`       | `agent.logs`：最新测试日志摘要与尾部                      | `logs()`                                                                                                    |
| `features/agent_patch.py`      | 补丁、断言、风格、指标四个动词                               | `handle_patch()`、`handle_assert()`、`handle_style()`、`handle_metrics()`                                      |
| `features/audit_a.py`          | 成员引用 / 公共库 / 项目依赖三个视角                         | `handle_members()`、`handle_common()`、`handle_project()`、`build_*_lines()`、`write_*_md()`                    |
| `features/audit_b.py`          | 结构 / 指标 / 总览三个视角                              | `structure()`、`indicators()`、`overview()`、`_probe_live()`                                                   |
| `features/audit_menu.py`       | `audit -i` 交互菜单：范围 / 分类 / 模块选择，六视角与 Markdown 导出 | `interactive_entry()`、`run_menu()`                                                                          |
| `features/audit_lint.py`       | `audit lint`：项目枚举、四条静态规则出口                       | `lint()`                                                                                                    |
| `features/audit_references.py` | `audit references`：只读引用体检，状态判定与轻量 `used_by`        | `references()`                                                                                              |
| `features/audit_engine.py`     | audit 共用规则引擎门面；组合三个 mixin                     | `AuditEngine`、`common_analyses()`、`summary()`、`overview_actions()`                                          |
| `features/audit_scan.py`       | 源码范围枚举、缓存、成员可达性扫描                             | `ScanMixin`：`common_files()`、`project_files()`、`analyze_target()`、`dep_rows()`                              |
| `features/audit_structure.py`  | 结构问题分组对账                                      | `StructureMixin`：`structure_issues()`、`structure_rows()`、`binding_issues()`、`regression_issues()`           |
| `features/audit_indicators.py` | H01–H19 指标计算                                  | `IndicatorMixin`：`indicator_rows()`、`_statusbar_violations()`、`_controls_pairing()`                         |
| `features/audit_rules.py`      | 阈值、白名单、正则、纯辅助函数；规则唯一来源                        | `_names_in_list()`、`_ignore_set()`、`_jsonable()`                                                            |
| `features/release.py`          | 发布 / 文档切片的占位；当前 `SPECS = ()`，无运行时命令           | —                                                                                                           |

### 5.4 `toolkit` 公共基础函数（26 个）

| 文件                       | 作用                                                | 主要函数 / 类                                                                                                                        |
| ------------------------ | ------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------- |
| `toolkit/__init__.py`    | 统一再导出公共 API；约束依赖方向和编码约定                           | —                                                                                                                               |
| `toolkit/base.py`        | 包根、配置位置、编码、Failure、JSON / 警告收口、路径、源码读取、信封         | `read_source_text()`、`entry_command()`、`make_envelope()`、`msg()`、`warn()`、`find_config()`                                       |
| `toolkit/contract.py`    | 跨命令共享数据形状                                         | `ErrorCode`、`Stage`、`MutationMode`、`Outcome`、`Failure`、`to_contract_failure()`                                                  |
| `toolkit/config.py`      | `mapping.json` 加载、本机覆盖层合并、写回                      | `load_config()`、`load_raw_config()`、`merge_local()`、`save_config()`、`config_source()`                                           |
| `toolkit/repo.py`        | 仓库结构基础函数：版本、模块名、源码文件、占位壳、目标解析                       | `module_name_of()`、`doc_stem()`、`source_files()`、`is_shell_file()`、`shell_text()`、`resolve_target()`                            |
| `toolkit/files.py`       | 文本 / 备份基础函数：UTF-8 + LF、GBK 副本、备份、清理、尾读              | `write_source()`、`transcode_exported_file()`、`gbk_import_copy()`、`backup_copy()`、`prune_backups()`、`read_tail()`                |
| `toolkit/excel.py`       | Excel / VBE COM 边界：起实例、附着、打开工作簿、备份、工程权限、导入 / 删除 / 替换 | `new_excel()`、`attach_running_excel()`、`open_disk_workbook()`、`backup_workbook()`、`ensure_vbe_access()`、`replace_module_code()` |
| `toolkit/macro.py`       | 唯一宏执行 / 分类入口                                      | `MacroRunner`、`MacroResult`、`MacroRunner.run()`、`MacroRunner.classify()`                                                        |
| `toolkit/check.py`       | 写回前后模块集合、文档模块、工作表模块、后缀模块对账                         | `Notes`、`check_module_set()`、`check_doc_modules()`、`check_sheet_modules()`、`check_suffix_modules()`                             |
| `toolkit/style_rules.py` | 风格检查与修复；只动 src 文本                                 | `scope_files()`、`run_checks()`、`run_checks_with_details()`、`run_fix()`                                                          |
| `toolkit/vba_parse.py`   | VBA 源码解析：剥注释 / 字符串、过程声明、成员表、规范化、规则扫描              | `code_lines()`、`declared_procs()`、`parse_members()`、`mask_strings()`、`scan_rules()`                                             |
| `toolkit/vba_diff.py`    | 模块 / 过程级差异纯函数；不碰 Excel                              | `parse_procedures()`、`compare_module()`、`compare_modules()`                                                                 |
| `toolkit/vba_lint.py`    | 四条静态 lint 纯文本规则；过程跨度、声明、返回值、限定调用                 | `lint_text()`、`parse_proc_spans()`                                                                                          |
| `toolkit/vbabin.py`      | 只读解析 `vbaProject.bin`：模块名、文档模块、工作表、两份对比（模块集合 + 解压后的正文）            | `vba_module_names()`、`vba_documents()`、`sheet_pages()`、`vba_module_sources()`、`compare_vba_modules()`                                                  |
| `toolkit/watch.py`       | 文件监听三件套                                           | `snap()`、`root_of()`、`changed()`                                                                                                |
| `toolkit/watchdog.py`    | 线程看门狗、PID 存活探测                                    | `run_with_watchdog()`、`WatchdogTimeout`、`pid_from_hwnd()`、`pid_alive()`                                                         |
| `toolkit/dialogs.py`     | Excel/VBE 弹窗识别                                    | `classify_window()`、`detect_dialog()`、`detect_target_dialog()`                                                                  |
| `toolkit/app_guard.py`   | Excel 应用状态快照 / 还原                                 | `AppGuard`、`guard()`、`restore()`、`calc_mode()`                                                                                  |
| `toolkit/session.py`     | 工作簿会话形状                                           | `WorkbookSession`、`missing_target()`                                                                                            |
| `toolkit/state.py`       | 目标状态 × 命令用途提示矩阵                                   | `TargetState`、`Purpose`、`rule_for()`                                                                                            |
| `toolkit/ui.py`          | 控制台排版与交互基础函数                                        | `read_line()`、`choose()`、`ask_yes()`、`pick_list()`、`pick_grouped()`、`run_menu()`                                                |
| `toolkit/cli.py`         | 见 5.2                                             | 见 5.2                                                                                                                           |
| `toolkit/commands.py`    | 见 5.2                                             | 见 5.2                                                                                                                           |
| `toolkit/context.py`     | 见 5.2                                             | 见 5.2                                                                                                                           |
| `toolkit/registry.py`    | 见 5.2                                             | 见 5.2                                                                                                                           |
| `toolkit/render.py`      | 见 5.2                                             | 见 5.2                                                                                                                           |

### 5.5 `adapters` 与开发脚本（3 个）

| 文件                             | 作用                                                            | 主要函数                                                                     |
| ------------------------------ | ------------------------------------------------------------- | ------------------------------------------------------------------------ |
| `adapters/__init__.py`         | 适配器层包标记                                                       | —                                                                        |
| `adapters/excel_gateway.py`    | COM 网关门面；`toolkit.excel` 的薄转发；`read_vbproject_probe()` 只读工程快照 | `attach_running_excel()`、`open_disk_workbook()`、`read_vbproject_probe()` |
| `scripts/make_lab_workbook.py` | 生成空白测试工作簿；不是业务命令                                                | `main()`                                                                 |

***

## 6. 30 分钟阅读顺序

如果只想快速建立全局理解，按这个顺序读：

1. `README.md`「人怎么用」+ `SKILL.md` 前两节：知道人怎么用菜单、agent 怎么用能力。

2. `scripts/sync.py`（57 行）：看入口壳到底多薄。

3. `toolkit/commands.py` + `toolkit/registry.py`：理解 “动词怎么登记、怎么派发”。

4. `features/core.py`：从最简单的只读动词开始。

5. `features/sync_read.py` 的 `pull()`：理解 VBE → src。

6. `features/sync_push.py` 的 `push()` + `sync_preflight.py`：理解 src → VBE。

7. `toolkit/excel.py`：只在需要排查 COM / VBE 权限时细读。

8. `toolkit/macro.py` + `features/run_macro.py`：宏执行链。

9. `features/audit_b.py` + `audit_engine.py`：盘点链。

10. `features/agent_patch.py`：Agent 补丁 / 断言 / 风格链。

不建议一上来读 `audit_scan.py`、`audit_indicators.py`、`sync_transform.py`：它们规则密集、私有辅助多，适合带着具体问题查。

***

## 7. 改哪里：按需求定位

| 需求              | 先改                                            | 再补                                                   |
| --------------- | --------------------------------------------- | ---------------------------------------------------- |
| 新增一个命令          | 对应 `features/*.py` 的 `SPECS` + handler        | `tests/test_*.py`；`references/commands.md` 重新生成      |
| 改参数 / 校验        | `CommandSpec` 的 `ArgSpec` 或 handler 入口        | `test_cli_contract` / `test_json_contract`           |
| 改 Excel 开关 / 附着 | `toolkit/excel.py`                            | `tests/test_owned_lifecycle_1_6_5.py`                |
| 改 pull 导出       | `features/sync_read.py`                       | `tests/test_sync_read.py`                            |
| 改 push 写回       | `features/sync_push.py` + `sync_preflight.py` | `tests/test_sync_push.py`、`test_source_preflight.py` |
| 改宏运行 / 错误分类     | `toolkit/macro.py` + `features/run_macro.py`  | `tests/test_run.py`、`test_run_1_6_5.py`              |
| 改 watch 监听      | `features/run.py` + `run_watch.py`            | `tests/test_run.py`                                  |
| 改风格规则           | `toolkit/style_rules.py`                      | `tests/test_agent_patch.py`                          |
| 改 audit 规则      | `features/audit_rules.py` + 对应 mixin          | `test_audit_*.py`                                    |
| 改输出 / 退出码       | `toolkit/render.py` + `toolkit/contract.py`   | `test_json_contract.py`、`test_error_contract.py`     |

***

## 8. 为什么文件这么多，哪里可以整合

不是所有文件都需要读。运行时主链其实只有：

```
3 个入口 + 12 个命令切片的主文件 + 8 个 toolkit 基础函数
```

其余大量文件是：

* 同一命令族的辅助段（例如 `run_macro` / `run_watch`、`sync_preflight` / `sync_shells`）；

* audit 的规则 / 扫描 / 指标拆分；

* 测试和开发工具。

### 8.1 真正可以合并的候选

1. `adapters/excel_gateway.py`

   目前是 `toolkit.excel` 的薄转发，只有 47 行。可以合并进 `toolkit/excel.py`；保留它的理由是把 “外部边界” 显式命名，便于以后替换 COM 实现。若要减少文件数，这是第一候选。

2. `features/release.py`

   只有 5 行，`SPECS = ()`，当前没有运行时命令。可以暂时删除，等真正接入发布命令时再加。

3. `agent_io.push` 与 `sync_push` 的写回逻辑

   两者是同族能力，但 `agent_io.py` 自己有 `_writeback()` / `_attach_or_open()`。这是目前最值得关注的重复点：不是文件多，而是同一套写回规则有两份实现，长期容易漂移。建议后续抽一个共享写回核心，两个入口都调用它。

### 8.2 不建议合并的

* **三个入口不要合并**：`sync` / `agent` / `audit` 是三个用户面，参数、JSON 契约、交互策略不同。

* **`features` 不要按文件数硬合并**：现在的切片是按命令族拆的；合并回一个 `sync.py` 会回到 1.5.6 单体。

* **`toolkit` 不要塞进 features**：公共基础函数被多个切片复用，塞回去会造成循环 import。

* **`run.py` / `run_macro.py` / `run_watch.py` 不建议合并**：一个是入口编排，一个是宏执行，一个是监听循环，职责清晰。

* **`audit_a` / `audit_b` 不建议合并**：一个负责成员 / 公共库 / 项目，一个负责结构 / 指标 / 总览；`audit_engine` 已经是共用门面。

***

## 9. 测试与开发工具地图

测试不是运行时，但它是 “行为规格”。45 个测试文件按被测对象分组：

| 组                | 文件                                                                                                                                                                                                              |
| ---------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 入口 / 契约 / 架构     | `test_cli_contract.py`、`test_core_sync_entry.py`、`test_error_contract.py`、`test_json_contract.py`、`test_architecture_skeleton.py`、`test_import_layers.py`、`test_version_consistency.py`、`test_version_match.py` |
| 配置 / 解析 / 读取     | `test_config_lookup.py`、`test_example_config.py`、`test_local_overlay.py`、`test_init.py`、`test_resolve_target.py`、`test_source_preflight.py`、`test_sync_read.py`、`test_sync_transform.py`                        |
| 写回 / 回滚          | `test_sync_push.py`、`test_postcheck.py`、`test_rollback_snapshot.py`                                                                                                                                             |
| run /watch/ 生命周期 | `test_run.py`、`test_run_1_6_5.py`、`test_run_journal_contract.py`、`test_owned_lifecycle_1_6_5.py`                                                                                                                |
| agent            | `test_agent_io.py`、`test_agent_json.py`、`test_agent_patch.py`、`test_agent_patch_1_6_5.py`、`test_agent_push_common.py`、`test_contract_scrub.py`                                                                  |
| audit            | `test_audit_a.py`、`test_audit_b.py`、`test_audit_engine.py`、`test_audit_gate.py`、`test_vbabin.py`                                                                                                                |
| 工具 / 回归          | `test_app_guard_1_6_5.py`、`test_watchdog.py`、`test_release_diff.py`、`test_gen_commands_md.py`、`test_new_issues.py`、`test_regression_1_6_4.py`                                                                   |

`tests/tools/`：

| 文件                   | 作用                           |
| -------------------- | ---------------------------- |
| `devloop_gate.py`    | A1 门禁：单测、编译、doctor、脱敏、安装副本校验 |
| `diff_harness.py`    | 差分 / 对比辅助                    |
| `gen_commands_md.py` | 从 manifest 生成命令表             |
| `make_release.py`    | 按交付白名单打 zip / SHA / latest 清单    |
| `profile_tests.py`   | 列出最慢测试                       |
| `upgrade_install.py` | 安装 / 升级副本                    |

`tests/baseline/record_1_5_6.py` 是历史基线记录脚本，不是运行时。

***

## 10. 最容易混淆的几组文件

| 容易混淆                                            | 区别                                                               |
| ----------------------------------------------- | ---------------------------------------------------------------- |
| `sync.py` / `agent.py` / `audit.py`             | 三个独立入口；互不 import；不要互相调用                                          |
| `features.run` / `run_macro` / `run_watch`      | `run.py` 是入口和监听循环；`run_macro.py` 是宏执行编排；`run_watch.py` 是监听辅助     |
| `features.sync_push` / `features.agent_io`      | 两者都有写回能力；`sync_push` 是主写回，`agent_io.push` 是 Agent 面的同族实现         |
| `toolkit.excel` / `adapters.excel_gateway`      | `toolkit.excel` 是真正实现；`excel_gateway` 是薄门面                       |
| `features.audit_a` / `audit_b` / `audit_engine` | `a/b` 是视角和报告；`engine` 是纯计算门面；`scan/structure/indicators` 是 mixin |
| `features.common_subset`                        | 不是命令切片，是 features 内部的共享叶子；`core/read/run/push` 都用它判公共模块子集        |

***

## 11. 术语表

| 术语             | 含义                                                        |
| -------------- | --------------------------------------------------------- |
| 工作簿 | 一个 `.xlsm` 文件；项目配置里的 `file` / `search_dir` 指向它 |
| 工作表 | Excel 里的 Sheet；文档模块绑定到它 |
| 抓取 | VBE → `src/` 源码 |
| 写回 | `src/` 源码 → VBE |
| `mapping.json` | 工作仓库根的项目配置：项目名、搜索目录、源码目录、公共模块子集 |
| `src/`         | VBA 源码目录；`.bas/.cls/.frm/.frx` 以 UTF-8 无 BOM + LF 存储 |
| `VBE`          | Excel 的 Visual Basic Editor；`vbaProject.bin` 是它的二进制工程 |
| 占位壳            | 项目 `src` 里指向公共库的一行注释文件；写回时跳过 |
| `SPECS`        | 每个 feature 切片登记的命令清单 |
| `Outcome`      | handler 的统一返回对象：`ok/details/errors/stats/human/exit_code` |
| `--apply`      | 唯一写开关；不加只预览 |
| `auto` 子集      | 按引用闭包算出的 “这个项目实际用到哪些公共模块” |

***

## 12. 维护规则

* 新增命令只改 `features` 的 `SPECS` + handler，不在入口里写 `if verb ==`。

* 跨切片复用的逻辑下沉 `toolkit`，不在切片之间互相 import。

* 规则数据只放 `audit_rules.py`，不在报告层复制阈值。

* 源码编码只走 `files.write_source()`；GBK 只出现在 VBE 交换边界。
* 文件名不等于模块名；一律用 `repo.module_name_of()` / `doc_stem()`，不要 `f.stem`。

* 写回预检硬项只增不减；放宽规则先写反例测试。

* 长任务在 handler 里用 `ctx.progress.stage()` 做阶段计时；进度只写 stderr，避免污染 JSON。

***

## 13. 结论

这个项目 “文件多” 主要来自**分层 + 命令族拆分 + 测试覆盖**，不是无意义的复制。

真正的主链只有：

```
入口壳 → manifest → dispatch → feature handler → toolkit 基础函数 → Outcome → render
```

先掌握这条链，再按命令族查 `features`，最后才下钻 `toolkit` 和 audit mixin。这样 61 个脚本文件不需要一次读完。