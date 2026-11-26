# 1.6 发布验收报告

- 对应提交：`4963bc0`（分支 `refactor/v1.6`，P14 全量复测基线）
- 脱敏：项目名、表名、绝对路径一律以 `<project>`、`<module>`、`<path>`、`示例表A/B/C/D` 占位。

## 验收口径（1.6）

- 命令面按**主要功能等价**验收，不再要求与 1.5.6 逐字兼容：命令名、参数、JSON 主结构与退出码语义一致即可；内部错误码细分、信封新增字段允许差异。
- **已知降级**与**已接受设计差异**均在文中显式标注（见 §3.3、§6.1、§7），不视为 1.6 交付阻塞；合入验证 / 1.7 复核分别登记。
- 新增信封键与错误码由旧 `FAILED` 兜底改为具体登记码，均属**已接受的设计差异**（向上兼容新增 / 更精细定位）；不写 approved delta、不登记白名单，也不为逐字一致删键或改码。
- 按本口径，原先针对新增键与错误码细分的阻塞项为 0；H16/H17 探测降级经批次 3 全量复测通过（`3b40fad`/`a8285ed`，P14 全量 396 例 0 失败），已转为功能等价，当前待查差异 0（见 §3.3、§6.4）。
- 验收过程不启动 Excel、不碰业务表，录制 / 差分 / 预览均在临时目录与桩配置内完成。

## 1. 验收结论摘要

| 步骤 | 命令 | 结果 | 退出码 | 结论 |
|---|---|---|---:|---|
| 编译 | `python -m compileall -q scripts tests` | 无语法告警 | 0 | 通过 |
| 刷新 pending | `record_1_5_6.py --sync-pending` | pending 0（写入 `approved_deltas.json`） | 0 | 功能等价 |
| 旧新差分 | `diff_harness.py` | match 38 / pending 0 / failed 0（`report.failed` 为空） | 0 | 功能等价 |
| 命令表 | `gen_commands_md.py` | 覆盖 31 动词 | 0 | 功能等价 |
| 升级预览 | `upgrade_install.py --target <path>`（未加 `--apply`） | 计划交付 58 个文件；需先备份 0 个；预览后目标目录 0 文件 | 0 | 未落盘 |
| 全量 unittest | 32 个模块 / 396 例 | failures 0 / errors 0，`OK (skipped=9)`（通过 387） | 0 | 通过 |

> 规模口径（P14 全量复测实测）：**32 模块 / 396 例 / skipped 9**（通过 387）。沿革：复测前基线（31 模块 / 355 例）→ 批次 3 口径（32 模块 / 383 例；+1 模块、+28 例，增量来自 P1 冲突副本过滤用例、P6 孤儿清理 / 冒烟 / 回滚 / 幂等用例，以及 P8 新增模块 `tests/test_import_layers.py` 分层护栏 9 例）→ P14 当前口径（396 例；**+13 例来自 P12/P13 补的用例**，模块数仍 32、skipped 仍 9）。31/355 与 32/383 仅作沿革对照，非当前口径。

> 交付数口径沿革：58 为更早历史口径 → 60 为 P1 冲突副本过滤后口径 → 61 为 P6 补 `LICENSE` 入清单后口径 → 62 为 P7 新增 `references/commands.md` 后口径 → **58 为 P11 目录治理后的当前口径**。P11 起 `docs/` 旁支目录移除：`docs/commands.md` 短指针直接删除（命令表唯一落点 `references/commands.md`）；`release-acceptance-1.6.md`、`release-diff.md`、`commands-alignment.md` 三份移入 `tests/`（不在交付根）；`backlog-1.7.md` 移入 `references/`（交付仍是 1，净 0）。相对 62 的差额 4：−1 短指针、−3 三份 tests 文档。

- 台账 `approved_deltas.json`：approved 0 / pending 0。
- 升级预览受保护、保持不动：`src/`、业务表（`.xlsm/.xlsb/.xlsx/.xlam/.xls/.frx`）、`mapping.json`、`mapping.local.json`、`.backup/`、`.lab/`、`.run/`。
- 预览计划范围：`scripts/`、`references/`、`assets/`、`SKILL.md`、`VERSION`、`LICENSE`、`requirements.txt`；排除 `__pycache__` / `.git` / `.lab` / `.run` / `.backup` / `releases` / `evals` / `tests` 与 `.pyc/.pyo/.tmp/.log`。

## 2. 已登记跳过（skipped=9）

> 不计失败；均为 1.6 有意延后 / 撤销的能力。runner 实测计 skipped=9，构成已复核：**7 条硬 `@unittest.skip`（`test_cli_contract` 1 + `test_json_contract` 6）+ `test_watchdog` 类级 `skipUnless(_WATCHDOG_AVAILABLE)` 覆盖的 2 个方法**。
>
> 另有 3 条**环境条件用例**不在这 9 条内：`test_json_contract` 的 overview `skipUnless`（要求仓库根 `mapping.json`）与 `test_release_diff.RealBaselineTest` 的 2 条基线条件用例（要求基线仓库 / `baseline-1.5.6` 标签）。本环境 `mapping.json` 存在、标签可用，**这 3 条实际未跳过、正常运行**；CI 环境若缺配置或缺标签，skip 数会相应增加。

| # | 跳过用例 | 登记原因 | 归属 |
|---:|---|---|---|
| 1 | `test_cli_contract.CliContractTest.test_sync_interactive_non_tty_is_clean` | 硬 `@unittest.skip`：sync `-i` 已注册且非 TTY 按旧版 exit 1 人话；菜单本体（多表选择等）随 1.7 迁移，本用例覆盖的菜单分支暂未回补 | 01（延后切片，1.7） |
| 2 | `test_json_contract.JsonContractTest.test_vbe_error_dialog_detects_modal` | 硬 `@unittest.skip`：win32gui 窗口枚举 / VBE 编译错误弹窗识别（`_vbe_error_dialog` / `_window_pid`）未从旧单体迁入；断言与桩保持 1.5.6 口径 | 06 / 07（待回补） |
| 3 | `test_json_contract.JsonContractTest.test_macro_run_error_classified` | 硬 `@unittest.skip`：`MACRO_NOT_AVAILABLE` / `MACRO_DISABLED` 分流码已撤销，缺宏 / 禁用宏统一收口 `RUN_FAILED`；分类器改为 `toolkit.macro.MacroRunner.classify`；该收口属已接受设计差异，旧断言口径随 1.7 复核 | 07（已接受设计差异） |
| 4 | `test_json_contract.JsonContractTest.test_vbe_generic_dialog_is_not_compile_error` | 硬 `@unittest.skip`：win32gui 窗口枚举 / 通用 VBE 模态框识别（`_vbe_dialog_title` / `_window_pid`）未迁入 | 06 / 07（待回补） |
| 5 | `test_json_contract.JsonContractTest.test_assert_timeout_excel_closed_has_code` | 硬 `@unittest.skip`：线程看门狗 `run_with_watchdog` / `WatchdogTimeout` 与按 PID 存活改判 `EXCEL_CLOSED` 分支已移除；超时结论由 `MacroRunner.timed_out → WATCHDOG_TIMEOUT` 表达 | 07（待回补等价原语） |
| 6 | `test_json_contract.JsonContractTest.test_excel_modal_dialog_detected` | 硬 `@unittest.skip`：win32gui 枚举 Excel 模态框（`_excel_modal_dialog` / `_window_pid`）未迁入 | 06 / 07（待回补） |
| 7 | `test_json_contract.JsonContractTest.test_agent_assert_failed_has_actionable_hint` | 硬 `@unittest.skip`：旧 main 把 handler 返回整数 2 合成 `ASSERT_FAILED` 信封（exit 2）的壳层口径已消失；assert 业务失败在新架构走 Outcome 错误对象、退出 1，错误信息与 actionable hint 保留，等价路径由 `test_agent_patch` 覆盖；该退出码口径属已接受设计差异，旧断言随 1.7 复核 | 07（已接受设计差异） |
| 8 | `test_watchdog.WatchdogTest.test_timeout_does_not_wait_for_slow_call` | 类级 `skipUnless(_WATCHDOG_AVAILABLE)`：线程看门狗 `run_with_watchdog` / `WatchdogTimeout` 重构后未迁移，原语不可用，2 个方法整体跳过 | 07（待回补） |
| 9 | `test_watchdog.WatchdogTest.test_value_and_exception_pass_through` | 类级 `skipUnless(_WATCHDOG_AVAILABLE)`：同上，看门狗原语未迁移 | 07（待回补） |

## 3. 31 动词命令面逐项对比

> 口径：`diff_harness.py` 命令面（参数、`--json` 支持、项目 / 全量开关）+ 解析层场景回放；旧 1.5.6 ledger vs 当前 manifest。按主要功能等价判定。
> 判定三级：**功能等价** = 命令名 / 参数 / JSON 主结构 / 退出码语义一致；**已接受设计差异** = 差异在错误码细分或内部原语层，命令面功能等价，属保留的新设计；**待查差异** = 尚未解释、需复查的出入。
> 结论：命令面 31 行中，`agent.assert` 1 行为已接受设计差异（命令面功能等价，差异在错误码 / 内部原语，见 §6.1）；audit 6 行命令面参数 / JSON / 退出码功能等价，其中 H16/H17 实时探测经批次 3 全量复测通过（见 §3.3），转为功能等价；**待查差异 0**；其余 30 行功能等价。
> 写动词的落盘行为以默认预览（`--apply` 才写）为准。

### 3.1 sync（16）

| 命令/动词 | 输入（关键参数） | 旧 1.5.6 输出 | 新 1.6 输出 | 差异结论 |
|---|---|---|---|---|
| `sync.version` | 无参；`--json` | 无项目开关；支持 json；信封键 `ok/action/details/errors/elapsed_ms/ts/stats/exit_code` | 同左 | 功能等价 |
| `sync.list` | 无参 | 解析面一致；支持 json | 同左 | 功能等价 |
| `sync.doctor` | 无参 | 解析面一致；支持 json | 同左 | 功能等价 |
| `sync.vbamod` | `--project`（可选） | 解析面一致；支持 json | 同左 | 功能等价 |
| `sync.plan` | `--project`；`--intent pull/push/run/watch`（默认 push） | 解析面一致；支持 json | 同左 | 功能等价 |
| `sync.pull` | `--project`/`--all-project`、`--no-backup`、`--json` | 解析面一致；支持 json | 同左 | 功能等价 |
| `sync.pull-open` | `--project`、`--no-backup`、`--json` | 解析面一致；支持 json | 同左 | 功能等价 |
| `sync.init` | `--project`、`--dir/--prefix/--src/--config/--lab`、`--json`；默认预览，`--apply` 才写 | 解析面一致；支持 json | 同左 | 功能等价 |
| `sync.push` | `--project`/`--all-project`、`--doc-modules/--no-save/--no-backup/--open-only/--rollback-on-fail`、`--json` | 预览容忍缺目标簿（exit 0）；`TARGET_NOT_FOUND` 仅在 apply 连接路径报出 | 同左（dry 缺表口径与 1.5.6 一致） | 功能等价 |
| `sync.shells` | `--mode create/clean`（默认 create）；`--json` | 解析面一致；支持 json | 同左 | 功能等价 |
| `sync.align` | `--project`（可选） | 解析面一致；支持 json | 同左 | 功能等价 |
| `sync.promote` | `--project`、`--module`（必填）、`--json` | 解析面一致；支持 json | 同左 | 功能等价 |
| `sync.gbk-export` | `--project`/`--all-project`、`--json` | 解析面一致；支持 json | 同左 | 功能等价 |
| `sync.clean-backup` | `--keep`、`--keep-tag`（可重复）、`--json`；默认预览 | 解析面一致；支持 json | 同左 | 功能等价 |
| `sync.run` | `--project`、`--macro`（必填）、`--arg`（可重复）、`--push/--isolate/--save/--tail/--log/--log-optional/--rollback-on-fail`、`--json` | 解析面一致；支持 json | 同左 | 功能等价 |
| `sync.watch` | `--project`、`--interval/--startup ask/src/vbe/each/skip/--no-save/--no-backup/--doc-modules/--seconds`；不支持 `--json`（持续写动词） | 解析面一致；`json=false` | 同左 | 功能等价 |

### 3.2 agent（9）

| 命令/动词 | 输入（关键参数） | 旧 1.5.6 输出 | 新 1.6 输出 | 差异结论 |
|---|---|---|---|---|
| `agent.probe` | `--strict`、`--json` | 解析面一致；支持 json | 同左 | 功能等价 |
| `agent.book` | `--project`、`--action open/close`（必填）、`--json` | 解析面一致；支持 json | 同左 | 功能等价 |
| `agent.backup` | `--project`、`--tag`、`--json` | 解析面一致；支持 json | 同左 | 功能等价 |
| `agent.logs` | `--project`、`--log`、`--tail`、`--json` | 解析面一致；支持 json | 同左 | 功能等价 |
| `agent.push` | `--project`、`--json` | 解析面一致；支持 json | 同左 | 功能等价 |
| `agent.assert` | `--project`、`--macro/--sheet/--clear/--baseline/--timeout/--log`、`--json` | 命令面一致；支持 json。错误码 / 内部原语层差异（见 §6.1）：看门狗公共原语未外导、宏不可用 / 禁用宏分流码撤销、VBE / 模态框弹窗识别未迁入、旧 exit 2 合成信封口径消失 | 命令面一致；超时由 `WATCHDOG_TIMEOUT` 表达，缺宏 / 禁用宏统一 `RUN_FAILED`，失败走 Outcome 错误对象 exit 1；错误信息与 actionable hint 保留 | 已接受设计差异（退出码 1，原 2；错误码 / 内部原语层，命令面功能等价；差异项见 §6.1） |
| `agent.patch` | `--project`、`--file`（必填）、`--no-style`、`--json` | 解析面一致；支持 json | 同左 | 功能等价 |
| `agent.metrics` | `--out`、`--json` | 解析面一致；支持 json | 同左 | 功能等价 |
| `agent.style` | `--project`、`--action check/fix`、`--which all/eol/call/explicit`、`--scope`、`--json` | 解析面一致；支持 json | 同左 | 功能等价 |

### 3.3 audit（6）

| 命令/动词 | 输入（关键参数） | 旧 1.5.6 输出 | 新 1.6 输出 | 差异结论 |
|---|---|---|---|---|
| `audit.members` | `--module`（必填）、`--exclude`（可重复）、`--all/--md/--out`、`--json` | 解析面一致；支持 json | 同左 | 功能等价 |
| `audit.common` | `--project`（可重复）、`--exclude/--all/--md/--out`、`--json` | 解析面一致；支持 json | 同左 | 功能等价（`--project` append 口径修复见 §7 映射表 B5） |
| `audit.project` | `--project`、`--exclude/--all/--md/--out`、`--json` | 解析面一致；支持 json | 同左 | 功能等价 |
| `audit.structure` | `--exclude/--ignore/--strict/--all/--md/--out`、`--json` | 解析面一致；支持 json；仅警告退出 3，`--strict` 退出 1 | 同左 | 功能等价 |
| `audit.indicators` | `--exclude/--ignore/--strict/--all/--md/--out`、`--json` | 解析面一致；退出码口径同上 | 同左 | 功能等价 |
| `audit.overview` | `--project`（可重复）、`--exclude/--ignore/--strict/--all/--md/--out`、`--json` | 解析面一致；退出码口径同上 | 同左 | 功能等价 |

> **H16/H17 实时探测口径（批次 3 复测通过，功能等价）**：H16（有窗体 / 控件时引用含 Microsoft Forms 2.0）、H17（VBA 工程可读、`CountOfLines` 可取）在 1.6 初版为已知降级——`_probe_live` 附着后 `checked` 恒 0，已开表时两项体检也恒判"未跑"。修复提交 `3b40fad`（merge `a8285ed`）已合入主干：`_probe_live` 回补网关只读探测，已开表时真实出通过 / 警告（Forms 引用 / `CountOfLines`），未开表仍标未跑；**P14 全量复测（32 模块 / 396 例，0 失败 0 错误）通过，这两项转为功能等价**。命令面参数 / JSON / 退出码按主要功能等价验收。

### 3.4 解析层场景回放（7 例，全部功能等价）

| 场景 | 旧退出码 / 信封 | 新退出码 / 信封 | 结论 |
|---|---|---|---|
| sync 裸跑 | exit 2，无信封 | 同左 | 功能等价 |
| sync 未知动词 | exit 2，无信封 | 同左 | 功能等价 |
| audit 裸跑 | exit 2，无信封 | 同左 | 功能等价 |
| audit 未知动词 | exit 2，无信封 | 同左 | 功能等价 |
| agent 裸跑 | exit 2，无信封 | 同左 | 功能等价 |
| agent 未知动词 | exit 2，无信封 | 同左 | 功能等价 |
| `sync.version --json` | exit 1，信封键 `action/details/elapsed_ms/errors/exit_code/ok/stats/ts` | 同左 | 功能等价 |

> **待查差异：0**。H16/H17 实时探测降级已经批次 3 全量复测通过（`3b40fad`，merge `a8285ed`），转为功能等价，见 §3.3；台账无 pending、无 approved 待办。

## 4. 新增键清单

顶层信封键与 1.5.6 `expected_envelope_keys` 一致：`ok / action / details / errors / elapsed_ms / ts / stats / exit_code`（`exit_code` 仅失败 / 需要时出现）。`warnings` 经 stderr 透出（不改退出码），不在 JSON 信封顶层。下列子键为 1.6 向上兼容新增，按新口径属**已接受的设计差异**，保留不回退。

| 键 | 位置 | 含义 | 可见性 | 结论 |
|---|---|---|---|---|
| `stats.timings` | `stats` 内 | 分阶段耗时字典（阶段名 → 毫秒，`AppContext.stage_timings()` 采集） | 内部：进度 / 诊断用，调用方一般不依赖；随 `stats` 出现在信封 | 已接受的设计差异（向上兼容新增） |
| `errors[].code` | `errors[]` | 登记错误码（大写下划线）；旧版缺省兜底 `FAILED` | 用户可见（对脚本 / 自动化调用方） | 已接受的设计差异（向上兼容新增） |
| `errors[].stage` | `errors[]` | 失败阶段（小写）；旧版缺省 `unknown` | 内部偏可见：主要供定位，随信封透出 | 已接受的设计差异（向上兼容新增） |
| `errors[].message` | `errors[]` | 错误人话正文 | 用户可见 | 已接受的设计差异（向上兼容新增） |
| `errors[].hint` | `errors[]` | 说明 / 下一步提示 | 用户可见 | 已接受的设计差异（向上兼容新增） |
| `errors[].suggested_command` | `errors[]` | 可直接执行的建议命令；非命令口径降级进 `hint` | 用户可见 | 已接受的设计差异（向上兼容新增） |
| `errors[].retryable` | `errors[]` | 是否可重试 | 内部偏可见 | 已接受的设计差异（向上兼容新增） |

> 其余 `stats` 子键（`lines/files/projects/excel/macro/results/dry/modules/...`）按动词各异，挂在既有 `stats` 对象内，不算新增顶层键。

## 5. 错误码映射清单（旧 `FAILED` 兜底 → 新登记具体码）

> 旧版未归类时兜底 `code="FAILED"`、`stage="unknown"`；1.6 统一走契约层 `Failure`，按触发场景映射到 `ErrorCode` 枚举（`toolkit/contract.py`）。
> 具体码取代旧 `FAILED` 兜底属**已接受的设计差异**：定位更精细，保留新设计，不为逐字一致回退到兜底码；下表各行均按已接受口径登记，无待拍板项。
> 可见性：**用户可见** = 出现在 JSON `errors[0].code` 并影响退出线索；**内部** = 主要用于内部路由 / 阶段定位。

| 触发场景（动词 / 环节） | 旧兜底 | 新错误码 | 阶段 | 可见性 |
|---|---|---|---|---|
| 目标工作簿未打开 / 已关闭（book close、attach、push/pull-open 取不到活动簿） | `FAILED` | `TARGET_NOT_OPEN` | `attach` | 用户可见 |
| 源码预检不通过（push / patch / audit 解析 / 目录缺失 / VB_Name 不符 / 缺类头 / 缺 frx / 续行过长） | `FAILED` | `SOURCE_PREFLIGHT` | `preflight`/`push`/`patch`/`parse`/`resolve` | 用户可见 |
| 项目名未指定 | `FAILED` | `PROJECT_NOT_SPECIFIED` | `config` | 用户可见 |
| 项目名在配置中找不到 | `FAILED` | `PROJECT_NOT_FOUND` | `config`/`resolve` | 用户可见 |
| 项目解析到的目标簿找不到 | `FAILED` | `TARGET_NOT_FOUND` | `config`/`resolve`/`open` | 用户可见 |
| 切不到活动簿 / 目标未激活 | `FAILED` | `TARGET_NOT_ACTIVE` | `attach` | 用户可见 |
| COM 不接话 / 运行时上下文异常 | `FAILED` | `COM_CONTEXT` | `open`/`attach` | 用户可见 |
| Excel 忙 / 被占用 | `FAILED` | `EXCEL_BUSY` | `attach` | 用户可见 |
| Excel 实例已关闭 | `FAILED` | `EXCEL_CLOSED` | `attach` | 用户可见 |
| VBE 工程受保护 | `FAILED` | `VBE_PROTECTED` | `vbe`/`attach` | 用户可见 |
| VBE 编译错误弹窗 | `FAILED` | `VBE_COMPILE_ERROR`（**登记保留、1.6 不可达**；弹窗识别族原语未迁入，见 §6.1 D4，随 1.7 回补） | `vbe` | 内部（登记码，1.6 无 throw 落点） |
| 断言日志读不出 | `FAILED` | `ASSERT_LOG_UNREADABLE` | `assert` | 用户可见 |
| 断言找不到日志 | `FAILED` | `ASSERT_NO_LOG` | `assert` | 用户可见 |
| 断言找不到测试表 | `FAILED` | `ASSERT_NO_TABLE` | `assert` | 用户可见 |
| 断言宏未配置 | `FAILED` | `ASSERT_MACRO_NOT_CONFIGURED` | `assert` | 用户可见 |
| 断言宏名非法 | `FAILED` | `ASSERT_MACRO_INVALID` | `config`/`assert` | 用户可见 |
| 断言宏参数非法 | `FAILED` | `ASSERT_MACRO_ARGS`（**登记保留、1.6 不可达**；参数个数不符当前统一走 `ARG_MISMATCH`，专用断言路径随 1.7 回补） | `assert` | 内部（登记码，1.6 无 throw 落点） |
| 断言未通过 | `FAILED` | `ASSERT_FAILED` | `assert` | 用户可见 |
| 看门狗硬超时 | `FAILED` | `WATCHDOG_TIMEOUT` | `assert` | 用户可见 |
| 日志目录找不到 | `FAILED` | `LOGS_NOT_FOUND` | `logs` | 用户可见 |
| postcheck 正文比对不一致 | `FAILED` | `POSTCHECK_FAILED` | `push`/`postcheck` | 用户可见 |
| 保存失败 | `FAILED` | `SAVE_FAILED` | `save` | 用户可见 |
| 补丁文件找不到 | `FAILED` | `PATCH_FILE_NOT_FOUND` | `patch` | 用户可见 |
| 补丁格式错 | `FAILED` | `PATCH_FORMAT` | `patch` | 用户可见 |
| 补丁块歧义 | `FAILED` | `PATCH_BLOCK_AMBIGUOUS` | `patch` | 用户可见 |
| 运行宏失败（含缺宏 / 禁用宏，分流码已撤销） | `FAILED` | `RUN_FAILED` | `run` | 用户可见 |
| 目标簿打开失败 | `FAILED` | `WORKBOOK_OPEN_FAILED` | `open` | 用户可见 |
| 目标已处于打开状态（隔离模式冲突） | `FAILED` | `TARGET_OPEN` | `open` | 用户可见 |
| 配置文件找不到 | `FAILED` | `CONFIG_NOT_FOUND` | `config` | 用户可见 |
| 配置文件 / 口径非法（含 Python 版本、必需文件缺失、备份口径非整数） | `FAILED` | `CONFIG_INVALID` | `config` | 用户可见 |
| 版本不匹配 | `FAILED` | `VERSION_MISMATCH` | `config` | 用户可见 |
| 环境缺 Excel / pywin32 / Windows | `FAILED` | `NO_EXCEL` | `config`/`attach` | 用户可见 |
| 文件只读 / 同步盘占用 | `FAILED` | `FILE_READONLY` | `config`/`open` | 用户可见 |
| 实验室模板缺失 | `FAILED` | `MISSING_LAB_TEMPLATE` | `init` | 用户可见 |
| 实验室模板生成失败 | `FAILED` | `LAB_TEMPLATE_FAILED` | `init` | 用户可见 |
| 参数 / 选项冲突 | `FAILED` | `ARG_MISMATCH` | `config`/`contract` | 用户可见 |
| 用法 / 参数冲突（契约层） | `FAILED` | `USAGE` | `config`/`contract` | 用户可见（argparse 用法错直接退出 2） |
| 中断 | `FAILED` | `INTERRUPTED` | — | 用户可见（退出 130） |
| 契约自检违规 | `FAILED` | `CONTRACT_VIOLATION` | `contract` | 内部 |
| 不支持的操作（如延后接入的菜单 / 原语） | `FAILED` | `UNSUPPORTED_OPERATION` | `contract` | 用户可见 |
| 其余未预期异常 | `FAILED` | `UNEXPECTED` | `unknown` | 用户可见（兜底，附 stderr 线索） |

> 上表按 `ErrorCode` 枚举与 features / toolkit 代码中实际 `code=` 落点归并；同一码可跨多个阶段复用。

## 6. 架构观察

### 6.1 差异明细（命令面之外）

| 编号 | 差异点 | 旧 1.5.6 | 新 1.6 | 可见性 | 处置状态 |
|---|---|---|---|---|---|
| D1 | sync 交互菜单本体 | `-i` 进入多表选择菜单 | `-i` 参数已注册；非 TTY 人话 exit 1；菜单本体随 1.7 迁移，`_interactive` 桩提示改用子命令 | 用户可见（菜单入口暂缺） | 1.7 增强项，不阻塞 1.6 验收 |
| D2 | 宏不可用 / 禁用宏分流码 | `MACRO_NOT_AVAILABLE`、`MACRO_DISABLED` 独立码 | 撤销分流，统一收口 `RUN_FAILED`；分类器为 `MacroRunner.classify` | 用户可见（错误码变粗） | 已接受设计差异（保留新设计；旧断言口径随 1.7 复核） |
| D3 | 线程看门狗公共原语 | `run_with_watchdog` 上下文管理器 + `WatchdogTimeout` 异常类，按 PID 存活改判 `EXCEL_CLOSED` | 无外导公共原语；超时内联断言路径，结论经 `WATCHDOG_TIMEOUT` 表达 | 内部（公共 API）；超时结论用户可见 | 1.7 增强项，不阻塞 1.6 验收 |
| D4 | VBE / Excel 弹窗识别 | win32gui 枚举识别 `#32770` 编译错误弹窗、通用 VBE 模态框、Excel 模态框（`_vbe_error_dialog` / `_vbe_dialog_title` / `_excel_modal_dialog` / `_window_pid`） | 未迁入 scripts | 内部（识别原语）；误判结果影响用户可见结论 | 1.7 增强项，不阻塞 1.6 验收 |
| D5 | assert 失败退出码口径 | handler 返回整数 2 时 main 合成 `ASSERT_FAILED` 信封（exit 2） | 该壳层合成口径消失；assert 业务失败走 Outcome 错误对象、**exit 1**，错误信息与 actionable hint 保留；等价路径由 `test_agent_patch` 覆盖 | 用户可见（退出码口径） | **已接受设计差异**：assert 失败退出码 1（原 2）；旧 exit 2 口径随 1.7 复核，不计缺陷 |

> D1 / D3 / D4 为 1.7 回补的功能缺口；D2 / D5 为已接受的设计差异。回补项见 `references/backlog-1.7.md`。

### 6.2 三个入口行数

| 入口 | 本轮行数 | 参考值 | 备注 |
|---|---:|---:|---|
| `scripts/sync.py` | 119 | ≈88 | 偏多约 31 行；全部为壳骨架：`_tool_version`、`--version` 参数块、`_to_contract`/`_failure_outcome` 契约适配助手、`-i` 入口与 `_interactive` 桩；无动词 handler |
| `scripts/agent.py` | 97 | ≈97 | 符合 |
| `scripts/audit.py` | 114 | ≈114 | 符合 |

### 6.3 `scripts/` 构成

- 三个入口壳：`sync.py` / `agent.py` / `audit.py`；独立工具 `make_lab_workbook.py`；`启动菜单.bat`。
- `adapters/`：`excel_gateway.py`（COM 网关适配）。
- `features/`：`core`、`sync_read`、`sync_push`、`sync_transform`、`run`、`agent_io`、`agent_patch`、`audit_a`、`audit_b`、`audit_engine`、`release`。
- `toolkit/`：契约 / 解析 / 配置 / 上下文 / 渲染 / 注册表 / Excel / 文件 / 宏 / 会话 / 状态 / 风格规则 / 界面 / VBA 解析 / 二进制 / 监听等共享设施。
- 所有 `.py` 均只被三个壳、`make_lab_workbook`、adapters、features、toolkit 构成；未见游离脚本。
- 三个壳顶层 `def` 仅：契约适配助手（`_to_contract` / `_failure_outcome` / `_wrap_failure`）、`_maybe_dump_out`、`_interactive`、`_tool_version` 与 `main`；未发现旧单体 `cmd_*` / 动词业务 handler 残留。

### 6.4 批次 3 全量复测收口

> 本小节登记全量复测结论：批次 3 为唯一全量专责复测；P14 在主干 `refactor/v1.6` HEAD `4963bc0` 上复核，下表数字为 P14 实测口径。本报告只登记复测结论，不替代全量测试记录。

| 复测目标 | 实测结果 | 结论 |
|---|---|---|
| 编译 `compileall` | 退出 0，无语法告警 | 达标 |
| 旧新差分 `diff_harness.py` | match 38 / pending 0 / failed 0（`report.failed` 为空） | 达标 |
| 命令表生成物 | 31 动词；`references/commands.md` 重生后 git 零改动，逐字节一致 | 达标 |
| 版本一致性 | `sync.py --version` → `vba-workbench 1.6.0`；`VERSION` 与 `SKILL.md metadata.version` 三处一致 | 达标 |
| 升级预览 `upgrade_install.py`（未加 `--apply`） | 计划交付 58 个文件；需先备份 0；预览后目标目录 0 文件；零冲突副本、无物理 `*_Conflict.md`；清单无 `docs/` 项 | 达标（未落盘） |
| 全量 unittest | 32 模块 / 396 例，`OK (skipped=9)`，0 failure / 0 error（通过 387） | 达标 |

- `watch --json` 用法错：exit 2、只出人话、不出 JSON 信封；该口径为已接受设计差异（依据 `test_json_contract.py` 用法错用例与 `SKILL.md` 失败行既有契约），对应用例为绿。
- 跨模块污染（`test_sync_push` + `test_rollback_snapshot`）本次**未复现**：合并跑无失败；隔离单跑 45 OK / 2 OK。此前标注的"基线预存污染"记为**本次未触发**，不构成 1.6 阻塞。
- H16/H17（B3）复测通过，已转为功能等价（见 §3.3）。

**次要项清零（P11/P12/P13 已全部修复并合入）**：批次 3 复测时登记的 7 条次要项与 I9 目录治理，已在后续三批逐条清零，提交均已合入主干：

| # | 原次要项 | 修复提交 | 落点口径 | 状态 |
|---:|---|---|---|---|
| ① | `base.Failure` 三入口归类不一致（audit `CONFIG_NOT_FOUND` vs sync/agent `UNEXPECTED`） | `9696f47` | 新增 `contract.to_contract_failure`：自带 code 沿用；无 code 一律 `UNEXPECTED`；stage 存在原样保留，缺失填 `UNKNOWN`；坏配置既有断言随之更新 | 已修复并合入 |
| ② | `errors[].stage` 只校验小写、不校验成员归属 | `0a6e3fe` | `Stage` 补 `preflight` / `postcheck` / `init`，`Failure` 强制枚举成员校验；`backup` / `import` 为进度标签、非 Failure stage，未纳入枚举 | 已修复并合入 |
| ③ | `emit_json` 死码路径 | `8767627` | 删除函数本体与 `toolkit/__init__` 导出，无兼容壳 | 已修复并合入 |
| ④ | `PROBE_FAILED` 未登记为契约错误码令牌（仅作回退字符串出现） | `a687f56` | 登记进 `ErrorCode`，`sync_read` 回退处改用枚举 | 已修复并合入 |
| ⑤ | `sync_read` 直连 `open_disk_workbook`，未走网关适配 | `d1952b8` | pull 磁盘开表改走 `excel_gateway` 只读入口，patch 点同步 | 已修复并合入 |
| ⑥ | `agent.metrics` 缺失败路径用例 | `d364e05` | 读盘 / 落盘两类失败路径断言 `UNEXPECTED` / exit 1 / hint；归类偏粗属现状，专用错误码留后续，本轮未写成已优化 | 已修复并合入 |
| ⑦ | 退出 130 中断（`INTERRUPTED`）无断言用例；并发现 audit 入口未捕 `KeyboardInterrupt` 的产品缺陷 | `f20b785`（audit 补捕获，中断返回 130 + 已中断）、`86ea949`（sync/agent/audit 三入口中断 130 与无 JSON 信封断言），merge `4963bc0` | 三入口中断路径均有断言覆盖 | 已修复并合入 |
| ⑧ | I9 `docs/` 旁支目录治理 | `8afa397`、`c44f671`、`e680156`，merge `730fa7b` | 分流结果见 §7.2 I9 行 | 已修复并合入 |

> 保留未触的登记错误码（`FILE_READONLY`、`LAB_TEMPLATE_FAILED`、`PROJECT_NOT_SPECIFIED`、`TARGET_OPEN`、`WORKBOOK_OPEN_FAILED`、`VBE_COMPILE_ERROR`、`ASSERT_MACRO_ARGS`）：当前 1.6 无 throw 落点，属 1.7 待办，不是缺陷；其中 `VBE_COMPILE_ERROR`、`ASSERT_MACRO_ARGS` §5 已标"登记保留、1.6 不可达"。

## 7. 巡检发现与修复映射

> 依据巡检复核总表：阻断 B1–B5、重要 I1–I11，以及 P11/P12/P13 三批次要项逐条对应修复批次与提交。各修复分支已全部提交并合入主干（合入顺序与 merge 号见下表）；状态以主干实际合入为准，未编造提交号。
> 状态口径：**已修复并合入** = 修复提交已合入主干；H16/H17（B3）经 P14 全量复测通过、转为功能等价；本批巡检项无"待修复"遗留。

### 7.1 阻断项（B1–B5）

| 编号 | 巡检问题 | 修复批次 / 提交 | 状态 |
|---|---|---|---|
| B1 | `*_Conflict.md` 冲突副本会进安装包（文件名内嵌真实设备名） | P1 `ca03b3c`（merge `6a8d6e6`）：安装器 `upgrade_install.py` 增加冲突名过滤，源目录 4 份副本清理 | 已修复并合入 |
| B2 | 真实业务表名未脱敏（`references/agent.md`、`tests/MANUAL.md` 两处） | P1 `ca03b3c`（merge `6a8d6e6`）：两处改为 `示例表A` 占位 | 已修复并合入 |
| B3 | audit H16/H17 体检恒"未跑"（`_probe_live` 附着后 `checked` 恒 0，已开表也静默少报） | P4 `3b40fad`（merge `a8285ed`）：回补网关只读探测，已开表真实出通过 / 警告（Forms 引用 / `CountOfLines`），未开表仍标未跑；批次 3 全量复测通过 | 已修复并合入（复测通过，转功能等价） |
| B4 | `watch --json` 用法错退出码违约（EXIT=1、无信封；`precheck` 生产链 0 调用方） | P3 `13c8ea3`（merge `c70aead`）：入口接 `precheck` 闸门，用法错归一 exit 2；信封口径维持"用法错只出人话"（依据 `test_json_contract.py` 用法错用例与 `SKILL.md` 失败行既有契约） | 已修复并合入 |
| B5 | `audit common --project` 静默收窄范围（重复传参只剩最后一个，多表审计少审且无提示） | P2 `60b1f0a`（merge `50dc241`）：恢复 `action="append"`，按列表过滤；单值仍收窄、缺省全量，输出结构与退出码不变 | 已修复并合入 |

### 7.2 重要项（I1–I11）

| 编号 | 巡检问题 | 修复批次 / 提交 | 状态 |
|---|---|---|---|
| I1 | 移除旧模块后 Import 失败时，模块清单差异只告警不拦 Save，模块可能从工作簿丢失 | P5 `416498d`（merge `b40109d`）：写回链三处"只告警"改硬拦（旧模块删不掉、Remove 后 Import 失败、清单比对丢返回值），`wb.Save()` 之前抛 `POSTCHECK_FAILED` | 已修复并合入 |
| I2 | 升级无孤儿文件清理、无安装后冒烟自检，旧副本模块可能被 import | P6 `c4c3ff1`（merge `6f7ddb3`）：安装器补孤儿清理 + 安装后冒烟自检 | 已修复并合入 |
| I3 | `apply` 逐文件 `copy2` 非原子，中途失败不自动回滚 | P6 `c4c3ff1`（merge `6f7ddb3`）：staging 暂存 + sha256 复核 + 备份快照统一替换，失败自动回滚 | 已修复并合入 |
| I4 | 失效错误码仍当可用码写在文档（`MACRO_NOT_AVAILABLE` / `MACRO_DISABLED` / `VBE_DIALOG` / `MODAL_DIALOG` / `VBE_DENIED` 等，41 码枚举中不存在） | P7 `d38b99f`（merge `7363f71`）：references 各册失效码改为"收口码 + 1.7 回补"，登记保留码标"1.6 不可达" | 已修复并合入 |
| I5 | 交互菜单仍写成可用路径（`_pick_projects` / `watch_open_list` 代码中无定义，`-i` 均为桩） | P7 `d38b99f`（merge `7363f71`）：sync/audit 命令册菜单段改标"1.7 待迁移；1.6 不可用" | 已修复并合入 |
| I6 | `sync-maintain.md` 整体基于旧单体骨架（约 30 个 `cmd_*` / `MENU_ITEMS` / `_dispatch` / `sync_into` 符号在 1.6 不存在） | P7 `d38b99f`（merge `7363f71`）：维护册按 1.6 三层骨架重写 | 已修复并合入 |
| I7 | 2 个死码与验收文档不符（`ASSERT_MACRO_ARGS`、`VBE_COMPILE_ERROR` 仅在枚举、无 throw 落点，§5 仍列为有效码） | P7 `d38b99f`（merge `7363f71`）：文档标注"登记保留、1.6 不可达"，§5 码表口径同步 | 已修复并合入 |
| I8 | import 分层护栏未固化成测试（features 互引、toolkit 反向引 features 无"违规即失败"断言） | P8 `1b8e19f`（merge `b0d2d4e`）：新增 `tests/test_import_layers.py`，纯 AST 三层护栏 + 自证违规样本 | 已修复并合入 |
| I9 | `docs/` 旁支目录违 §2.3（5 份已跟踪且进安装计划；验收报告应放 `.lab` 不入库） | P11 `8afa397`（分流/删除）+ `c44f671`（引用改址、安装器交付根去 docs，`UPGRADE_ROOTS` 移除 docs）+ `e680156`（按实测回写交付数 58），merge `730fa7b`：`docs/commands.md` 删除；`backlog-1.7.md` → `references/`；`release-acceptance-1.6.md` → `tests/reports/`；`release-diff.md`、`commands-alignment.md` → `tests/baseline/1.5.6/`；`docs/` 目录本身移除；交付数 62 → 58 | 已修复并合入 |
| I10 | `EVAL.md` 过期（1.5.5 / 73 项 / 118 行 / 1.2.0 包名等多处过期，照它自测会假失败） | P7 `d38b99f`（merge `7363f71`）：EVAL.md 过期口径按复测前基线对齐，入口行数 119 / 97 / 114；复测后 32 模块 / 396 例以本报告 §1 为准，EVAL.md 数字本轮同步收口 | 已修复并合入 |
| I11 | 生成表落点不符规范（`gen_commands_md.py` 默认写 `docs/commands.md`，规范要求 `references/commands.md`，且 `SKILL.md` 未链该表） | P7 `d38b99f`（merge `7363f71`）：命令表唯一落点改为 `references/commands.md`，`docs/commands.md` 改为短指针，`SKILL.md` 补链接；短指针后于 P11 I9 删除 | 已修复并合入 |

> 备注：I2/I3 同属 P6 一条提交 `c4c3ff1`；I4/I5/I6/I7/I10/I11 同属 P7 一条提交 `d38b99f`。P7 生成的命令表基于合入前源码，合并时已按合并后源码重生成（`common` 的 `--project` 从单值纠正为可重复 append）；I9 由 P11 分支 `fix/p11-docs-governance` 收口（`8afa397`/`c44f671`/`e680156`，merge `730fa7b`）。

### 7.3 目录治理与次要项清零（P11 / P12 / P13）

| 批次 | 编号 | 巡检问题 | 修复提交 / merge | 状态 |
|---|---|---|---|---|
| P11 | I9 | `docs/` 旁支目录治理（分流删除、引用改址、安装器去 docs、交付 62 → 58） | `8afa397`、`c44f671`、`e680156`，merge `730fa7b` | I9 已修复并合入（分流结果详见 §7.2 I9 行） |
| P12 | ① | `base.Failure` 三入口归类不一致（audit `CONFIG_NOT_FOUND` vs sync/agent `UNEXPECTED`） | `9696f47`：新增 `contract.to_contract_failure`，无 code → `UNEXPECTED`，缺 stage → `UNKNOWN` | 已修复并合入 |
| P12 | ② | `errors[].stage` 只校验小写、不校验成员归属 | `0a6e3fe`：`Stage` 补 `preflight`/`postcheck`/`init`，`Failure` 强制枚举校验；`backup`/`import` 为进度标签、非 Failure stage | 已修复并合入 |
| P12 | ③ | `emit_json` 死码路径（零调用，本体 + 导出残留） | `8767627`：删函数本体与 `toolkit/__init__` 导出，无兼容壳 | 已修复并合入 |
| P12 | ④ | `PROBE_FAILED` 未登记为契约错误码令牌（仅作回退字符串） | `a687f56`：登记进 `ErrorCode`，回退处改用枚举 | 已修复并合入 |
| P12 | ⑤ | `sync_read` 直连 `open_disk_workbook`，未走网关适配 | `d1952b8`：pull 磁盘开表改走 `excel_gateway` 只读入口，patch 点同步 | 已修复并合入 |
| P13 | ⑥ | `agent.metrics` 缺失败路径用例（读盘 / 落盘失败无断言） | `d364e05`：两类失败断言 `UNEXPECTED` / exit 1 / hint；归类偏粗属现状，专用码留后续 | 已修复并合入 |
| P13 | ⑦ | 退出 130 中断无断言；并发现 audit 入口未捕 `KeyboardInterrupt` | `f20b785`（audit 补捕获，中断返回 130 + 已中断）、`86ea949`（三入口中断 130 与无 JSON 信封断言），merge `4963bc0` | 已修复并合入 |

> **本轮新发现缺陷**：audit 入口原先未捕 `KeyboardInterrupt`（sync / agent 已捕，audit 漏网）。P13 实测发现，`f20b785` 补齐捕获（中断返回 130、人话"已中断"），`86ea949` 补 sync / agent / audit 三入口中断断言；随 merge `4963bc0` 合入主干。
