# 1.6.3 发布验收报告（review/v1.6-local 本地评审分支）

- 对应提交：分支 `review/v1.6-local`；代码终态 `6ee66d0`（各批次合并完成点，1.5.6 覆盖台账 + release_diff skip 口径修正合入）；文档收口产生的提交只改文档，分支 HEAD 以 `git log` 为准。本报告数字按 `6ee66d0` 实测；文档提交不改代码与测试。本报告只登记本轮本地评审验收；1.6.2 八项修复与命令绝对化收口见 `release-acceptance-1.6.2.md`，1.6 分层架构与已接受设计差异见 `release-acceptance-1.6.md`，1.6.1 六项修复见 `release-acceptance-1.6.1.md`，不重复改写。
- 脱敏：项目名、表名、绝对路径一律以 `<project>`、`<module>`、`<path>`、`示例表A/B/C/D`、`V100` 占位。
- 口径：本轮只做文档收口、结构拆分复核与自动化回归；**不安装、不升级安装副本、不启动真实 Excel、不跑真实宏与写回**，COM 全部打桩。验收过程不碰业务表，预览与定向落盘均在临时目录 / 桩配置内完成。父仓库规范 v1.4 已合入 `refactor/v1.6`（merge `4a4694b`，本地未 push），本轮验收不改父仓库，状态见 §7。

## 1. 验收结论摘要

| 步骤 | 命令 / 口径 | 结果 | 退出码 | 结论 |
|---|---|---|---:|---|
| 编译 | `python -m compileall -q scripts tests` | 无语法告警 | 0 | 通过 |
| 全量 unittest | 33 个模块 / 494 例 | failures 0 / errors 0，`OK (skipped=4)`（通过 490） | 0 | 通过 |
| 命令表 | `gen_commands_md.py` 幂等门禁 | 覆盖 31 动词；生成物与 manifest 逐字节一致，未重新生成 | 0 | 功能等价 |
| 版本一致性 | `VERSION` / `SKILL.md metadata.version` / `--version` | 三处均为 `1.6.3` | 0 | 通过 |
| 架构骨架 / 分层护栏 | `test_architecture_skeleton.py`、`test_import_layers.py` | 拆分后依赖方向仍单向、无循环 | 0 | 通过 |

> 规模口径（代码终态 `6ee66d0` 实测）：**33 模块 / 494 例 / skipped 4**（通过 490，0 failure / 0 error）。规模沿革：1.6.2 终态 `52368a5` 为 32 模块 / 450 例 / skipped 9（通过 441）；1.6.3 定版段（`d0e9c2c` 口径）为 475 例 / 467 通过 / skipped 8；代码终态 494 例 / 490 通过 / skipped 4。450 / 475 仅作沿革对照，非当前口径。命令面仍为 **31 动词**（sync 16 / audit 6 / agent 9），manifest 未改；入口壳行数实测 `scripts/sync.py` 74、`scripts/agent.py` 59、`scripts/audit.py` 77；`SKILL.md` 122 行。

## 2. 本轮闭环项逐项验收

### 2.1 B1 线程看门狗公共原语外导 —— 已在 1.6.3-local 闭环

- **问题**：1.6 验收时看门狗逻辑内联在断言路径，缺少可被其它动词复用的公共原语，超时 / 关闭两条分支无法独立测试。
- **改法**：新增 `toolkit/watchdog.py`（157 行）：`WatchdogTimeout` 硬超时异常类型、`run_with_watchdog(fn, timeout, ...)` 守护线程执行（超时不再等待、返回值 / 线程内异常原样回传，时钟与休眠可注入）、`pid_alive(pid, ...)` 三态存活探测、`pid_from_hwnd(hwnd)` 句柄反查 PID。原语与 Excel/COM 无关；COM 对象不能跨线程复用，包 COM 调用时 `fn` 内部需重新 attach 会话。
- **验证证据**：提交 `0e5d047`（merge `334a1ff`）。`test_watchdog.py` 看门狗两条旧跳过用例恢复运行（`test_timeout_does_not_wait_for_slow_call`、`test_value_and_exception_pass_through`），另补 `test_timeout_disabled_runs_synchronously`；原语不抛业务失败码，分类由上层编排层完成。

### 2.2 B2 win32gui 弹窗 / 模态识别族 —— 已在 1.6.3-local 闭环

- **问题**：识别原语未迁入 scripts，VBE 编译错误弹窗、通用 VBE 模态框、Excel 模态框 3 条断言保持 skip；弹窗出现时只能以粗口径兜底。
- **改法**：新增 `toolkit/dialogs.py`（168 行）：`WindowInfo` / `DialogMatch` 纯数据快照，`classify_window` 纯函数分类（编译错误 / VBE 通用模态 / Excel 模态 / 未知），`detect_dialog` 按目标 PID 过滤并按「编译错误 > VBE 通用 > Excel 模态」优先级取最具体弹窗，`enumerate_windows` / `detect_target_dialog` 才做真实 win32gui 枚举。顶层只标准库，`win32gui` / `win32process` 延迟导入；归属不明返回 unknown / `None`，不瞎猜。
- **验证证据**：提交 `0e5d047`（merge `334a1ff`）。`test_watchdog.py` 弹窗分类族 8 例（编译错误识别 / 子正文关键字 / 通用 VBE 不误判 / Excel 模态 / 非对话框类未知 / PID 过滤 / 优先级 / 无可识别弹窗返回 None）；`test_json_contract.py` 三条旧跳过弹窗用例恢复：`test_vbe_error_dialog_detects_modal`、`test_vbe_generic_dialog_is_not_compile_error`、`test_excel_modal_dialog_detected`。
- **错误码映射**（由 `features.agent_patch` 超时分支完成，dialogs 模块不抛 Failure）：编译错误框 → `VBE_COMPILE_ERROR`（`stage=vbe`）；通用 VBE / Excel 模态框 → `WATCHDOG_TIMEOUT`（`stage=assert`），`hint` 点名弹窗标题与正文关键片段。旧 `VBE_DIALOG` / `MODAL_DIALOG` 细分码不在当前错误码枚举中。

### 2.3 B3 缺宏 / 禁用宏分流码 —— 复核完成，维持已接受设计差异（非修好）

- **复核结论**：**不恢复** `MACRO_NOT_AVAILABLE` / `MACRO_DISABLED` 两个旧分流码；缺宏与禁用宏统一收口 `RUN_FAILED`，分类器为 `toolkit.macro.MacroRunner.classify`。区分信息固定落到 `errors[].hint`：缺宏写清宏名 / 模块名与编译问题线索，禁用宏写清信任位置 / 宏安全设置，两条 hint 文案可区分；歧义句两类线索都在，仍收口 `RUN_FAILED`。
- **验证证据**：提交 `cf30b95`（merge `44487c8`）。旧跳过用例 `test_macro_run_error_classified` 复核后恢复运行，断言两条文案均为 `RUN_FAILED`、`stage=run`、hint 互不相同且各带对应线索，并显式断言不再落回两个旧分流码。
- **口径备注**：此项是「复核确认已接受设计差异」，不是把旧码修回来；调用方仍不能从错误码区分缺宏与禁用宏，只能读 hint。

### 2.4 B4 超时按 PID 存活改判 `EXCEL_CLOSED` —— 已在 1.6.3-local 闭环

- **问题**：看门狗等待期间实例被外部关闭，旧版按 PID 存活探测区分「真超时」与「实例已关闭」；该分支随看门狗原语一并移除后，实例关闭与硬超时都收敛到超时 / 关闭类结论，定位偏粗。
- **改法**：随 B1 一并回补 `pid_alive` 三态探测；`agent_patch` 超时二次分类优先级为：目标 PID 已退出 → `EXCEL_CLOSED`（`stage=assert`，**优先于超时码**）；仍存活（或探测未知）→ 交 dialogs 认弹窗；判不出实例状态也认不出弹窗 → 回退 `WATCHDOG_TIMEOUT`（`stage=assert`）。pywin32 缺失 / 句柄被拒 / 查询异常一律降级为「未知」，不把权限问题误报成崩溃。
- **验证证据**：提交 `0e5d047`（merge `334a1ff`）。旧跳过用例 `test_assert_timeout_excel_closed_has_code` 恢复运行（`alive=False` → `EXCEL_CLOSED`）；另补 `test_assert_timeout_alive_no_dialog_is_timeout`（存活 + 认不出弹窗 → `WATCHDOG_TIMEOUT`）、`test_assert_timeout_compile_error_is_vbe_error`（存活 + 编译错误框 → `VBE_COMPILE_ERROR`，hint 点名标题正文）。

### 2.5 B5 旧 main 层合成 `ASSERT_FAILED` exit 2 口径 —— 复核完成，维持已接受设计差异（非修好）

- **复核结论**：**不补回** exit 2 壳层合成。`agent assert` 业务失败走 Outcome 错误对象、**退出 1**，错误信息与 actionable hint 保留；旧 main 把 handler 返回整数 2 合成 `ASSERT_FAILED` 信封（exit 2）的口径属已接受设计差异。
- **验证证据**：提交 `cf30b95`（merge `44487c8`）。旧跳过用例 `test_agent_assert_failed_has_actionable_hint` **维持硬 `@unittest.skip`**（旧 `exit_code=2` 断言不恢复）；等价路径由 `test_agent_patch.test_assert_failure_exit_1_with_code` 覆盖。
- **口径备注**：此项是「复核确认已接受设计差异」，不是修好；依赖退出码 2 判定「断言未通过」的外部脚本 / 自动化需改判退出 1。

### 2.6 B6 `sync -i` 交互菜单本体 —— 已回补菜单本体

- **问题**：旧版 `sync -i` 进入多表选择交互菜单；1.6 仅注册 `-i` 参数、`_interactive` 为桩，交互式选项目 / 选模块入口暂缺。
- **改法**：新增 `features/sync_interactive.py`（190 行）：列项目（来自 mapping）→ 选项目 → 列动作（既有登记子命令）→ 确认 → 复用 registry 登记的既有 handler 执行。菜单层不碰 COM、不写盘，只做选择与参数装配；`reader` / `writer` / `runner` 均可注入；非 TTY 终端把关仍由入口壳 `ensure_interactive` 负责；菜单默认走预览口径，写盘动词不带 `--apply`，真正落盘走命令行。
- **验证证据**：提交 `cf30b95`（merge `44487c8`）。`test_core_sync_entry.py` 新增 `SyncInteractiveMenuTest` 6 例：`test_lists_projects_from_mapping`、`test_picks_project_and_action_then_delegates_to_handler`、`test_cancel_at_action_level_returns_to_projects_then_exits`、`test_cancel_confirmation_does_not_run_handler`、`test_no_projects_exits_1_with_human_message`、`test_non_tty_shell_exits_1_clean`；`test_cli_contract.test_sync_interactive_non_tty_is_clean` 旧跳过用例恢复运行。同步移除死函数 `macro_error_hint`（`toolkit/base.py`）。

### 2.7 1.5.6 基线覆盖台账 —— 已补齐

- **改法**：新增 `tests/baseline/COVERAGE.md`（三态定义：covered / missing、verified / approved-diff / pending、success / business_fail / usage_error / json_purity）、`tests/baseline/1.5.6/coverage.json`（31 动词 / 124 条案例证据，机读字段与本表一一对应）、`tests/baseline/README.md`；并修正 `test_release_diff.py:195` 基线不可用分支由 `self.skipUnless` 改为 `self.skipTest`。
- **验证证据**：提交 `c219a35`（台账）、`0b02fec`（skip 口径修正），merge `6ee66d0`。台账当前口径：verified 25 / approved-diff 2（sync.doctor、agent.push）/ pending 4（agent.assert、agent.patch、agent.metrics、agent.style 的 CLI `--json` 接受度，未登记白名单）；业务失败缺例 1 条（`sync.shells`）。三类 golden（Excel 状态 / 文件副作用 / 真实宏执行）本轮未录制，台账 §4 显式登记。

### 2.8 四处 features 结构拆分 —— 已合入，行为等价

| 切片批次 | 入口薄层（行数） | 新拆出切片（行数） | 提交 |
|---|---|---|---|
| audit 规则 / 扫描 / 结构 / 指标 | `audit_engine.py` 1349 → 142 | `audit_rules` 106 / `audit_scan` 333 / `audit_structure` 416 / `audit_indicators` 459 | `fe170ef`（merge `bbc9ea1`） |
| push 预检 / 占位壳 / init | `sync_push.py` 1232 → 771 | `sync_preflight` 208 / `sync_shells` 103 / `sync_init` 205 | `d04fd48`（merge `3637afa`） |
| run 宏执行 / watch 监听 | `run.py` 992 → 312 | `run_macro` 547 / `run_watch` 206 | `8c7d79f`（merge `790317b`） |
| agent probe / 日志汇总 | `agent_io.py` 888 → 692 | `agent_probe` 146 / `agent_logs` 94 | `8c7d79f`（merge `790317b`） |

- 拆分口径：入口薄层只做参数适配与装配，判定 / 扫描 / 监听 / 汇总按职责进切片；规则阈值表单一来源（`audit_rules`），mixin 切片只吃引擎实例上的 root/config/缓存；`test_import_layers.py` 分层护栏随拆分补充白名单条目。拆分未改命令面、未改信封契约，生成物一致性测试通过。

## 3. 全量回归结论

- 代码终态 `6ee66d0`；全量复跑：**33 模块 / 494 例 / 490 通过 / 0 failures / 0 errors / skipped 4 / exit 0**；`python -m compileall -q scripts tests` 退出 0，无语法告警。
- 命令表 31 动词生成物与 manifest 逐字节一致；版本三处一致（`VERSION`、`SKILL.md metadata.version`、`--version` 均 `1.6.3`）；架构骨架与分层护栏通过。
- **skipped 4 条逐条构成**（本评审环境实测）：

| # | 用例 | 跳过方式 / 原因 | 性质 |
|---|---|---|---|
| 1 | `test_json_contract.JsonContractTest.test_agent_assert_failed_has_actionable_hint` | 硬 `@unittest.skip`：B5 已复核为接受设计差异，旧 main 合成 `ASSERT_FAILED` 信封（exit 2）壳层口径不恢复；assert 业务失败退出 1，actionable hint 由 `test_agent_patch.test_assert_failure_exit_1_with_code` 覆盖 | 已接受设计差异，不恢复旧断言 |
| 2 | `test_json_contract.JsonContractTest.test_audit_overview_stdout_is_single_json_line` | `skipUnless`：要求仓库根存在 `mapping.json`；本评审 worktree 根无配置文件，环境条件不满足 | 环境条件，非产品问题；带配置环境复跑即恢复 |
| 3 | `test_release_diff.RealBaselineTest.test_ledger_recorded_and_version_present` | `skipTest`：基线仓库 / `baseline-1.5.6` 标签在本 worktree 不可用 | 环境条件，非产品问题；具备基线环境自动恢复对比 |
| 4 | `test_release_diff.RealBaselineTest.test_real_version_scenario_matches` | `skipTest`：同上，基线仓库 / 标签不可用 | 环境条件，非产品问题；具备基线环境自动恢复对比 |

- 与 1.6.2 口径差异说明：1.6.2 验收环境带 `mapping.json` 且基线标签可用，上述第 2–4 条当时实跑通过、未计入 skipped；本评审 worktree 未挂本机配置与基线仓库，故这 3 条按环境条件跳过。CI 环境缺配置或缺标签时 skip 数会相应增加。功能型跳过仅剩第 1 条（B5）。

## 4. 未闭环项清单

| 项 | 状态 | 说明 |
|---|---|---|
| B3 缺宏 / 禁用宏细分码 | **复核完成，维持已接受设计差异** | 不恢复 `MACRO_NOT_AVAILABLE` / `MACRO_DISABLED`；统一 `RUN_FAILED`，区分信息只在 hint。属口径收粗，不是缺陷修复 |
| B5 assert 旧 exit 2 壳层合成 | **复核完成，维持已接受设计差异** | 不恢复 exit 2；assert 业务失败退出 1。依赖旧退出码 2 的外部调用方需改判 |
| 4 个 agent 动词 CLI `--json` 接受度（assert / patch / metrics / style） | 台账 pending，待裁决 | 进程内单行信封渲染有测试；端到端 `--json` 行为登记在 `test_release_diff.KNOWN_REAL_GAPS`，待 07 切片复核裁决后补录台账 |
| `sync.shells` 业务失败用例 | 台账 missing | 仅有 create/clean 预览与非法 `--mode` 用法错，缺目标缺失 / 预检失败类业务失败证据 |
| 三类 golden | 未录制 | Excel 状态 golden、文件副作用 golden、真实宏执行 golden；本轮约定不启动 Excel、COM 全打桩，待带 Excel 会话的本地集成环境录制回填 |
| 环境条件用例 3 条 | 本环境跳过 | 见 §3 第 2–4 条；非产品缺口，换具备配置 / 基线标签的环境复跑即可 |

## 5. 按批次 commit 与改动摘要

| 批次 | commit（短 / 全） | 合并提交 | 改动摘要 |
|---|---|---|---|
| `batch/audit-split` | `fe170ef` / `fe170ef556e32400e50296bd88f959ad5d25d243` | `bbc9ea1` | `audit_engine` 1349 → 142 行门面；拆出 `audit_rules` 106（阈值 / 白名单 / 分组 / 正则纯函数）、`audit_scan` 333（范围枚举 / 读取缓存 / 成员可达性 mixin）、`audit_structure` 416（壳 / 重名 / 绑定 / 文档表页对账 mixin）、`audit_indicators` 459（H01–H19 离线指标）；`test_import_layers` +9 |
| `batch/push-split` | `d04fd48` / `d04fd4805241737248c739e2232d1f504c3f16ab` | `3637afa` | `sync_push` 1232 → 771；拆出 `sync_preflight` 208（写回前只读预检，返回 fatal/notes）、`sync_shells` 103（公共库占位壳生成 / 清除）、`sync_init` 205（mapping.json 与 src 骨架生成 / 更新）；`test_import_layers` +6，`test_init` / `test_sync_push` 适配 |
| `batch/interactive` | `cf30b95` / `cf30b95e04d8231b64c64bfc75abaf28740c5f09` | `44487c8` | 回补 `sync -i` 菜单本体 `features/sync_interactive.py` 190 行 + 6 条新用例；复核 B3 宏分类 hint 分流（统一 `RUN_FAILED`，hint 可区分）与 B5 assert exit 1；移除死函数 `macro_error_hint`；`macro.py` 分类器收口；`references/sync.md` / `sync-commands.md` 口径同步 |
| `batch/watchdog` | `0e5d047` / `0e5d047511e4f5a51f266d12cd44ce8b774ed139` | `334a1ff` | 新增 `toolkit/watchdog.py` 157（B1 看门狗原语、B4 PID 三态存活）、`toolkit/dialogs.py` 168（B2 弹窗识别族）；`agent_patch` 超时二次分类（`EXCEL_CLOSED` 优先 → 编译错误框 `VBE_COMPILE_ERROR` → 其余模态 `WATCHDOG_TIMEOUT`）；恢复 6 条旧跳过用例（watchdog 2 + json_contract 弹窗 / 关闭 4） |
| `batch/run-split` | `8c7d79f` / `8c7d79f9681c4ac0d659e98a74c597d741cae292` | `790317b` | `run` 992 → 312，拆出 `run_macro` 547（参数解析 → 预检 / 备份 / 写回 / 回读 → 跑宏 → journal）、`run_watch` 206（目标解析 / 探活 / 比对 / 单文件写回抓回 / 基线刷新）；`agent_io` 888 → 692，拆出 `agent_probe` 146（项目解析与状态一览）、`agent_logs` 94（日志目录定位 / GBK 读取 / 计数 / 尾部截取）；`test_import_layers` +11 |
| `batch/baseline-ledger` | `c219a35` / `c219a351e749d85ab60bd954af282c0d0b3dba0a` | `6ee66d0` | 补全 `tests/baseline/COVERAGE.md`、`tests/baseline/1.5.6/coverage.json`（31 动词 / 124 条证据）、`tests/baseline/README.md` 三态定义 |
| `batch/baseline-ledger` | `0b02fec` / `0b02fec1da2bfdf731d0760f02be608afae6cc0a` | `6ee66d0` | `test_release_diff.py:195` 基线不可用分支 `self.skipUnless` → `self.skipTest` |

## 6. 风险与人工验证缺口

- **真实 VBE 弹窗未实测**：dialogs 识别族目前只由纯函数 + 伪造 `WindowInfo` 桩覆盖；真实 win32gui 枚举、子窗口正文拼接、跨进程 PID 归属与真实编译错误框 / MsgBox / InputBox 的识别效果，未在真实 Excel / VBE 会话验证。
- **真实 Excel 会话未跑**：本轮 COM 全部打桩，写回、宏执行、回读校验、watch 增量、弹窗时序均未做端到端手工验收；`tests/MANUAL.md` 清单待在带 `.lab` 空白表的本机环境执行。
- **安装副本未升级**：源码树 `VERSION=1.6.3`，本机安装副本 `<安装副本>` 仍是 **1.6.2**；本轮明确不安装、不跑 `upgrade_install.py --apply`，安装副本与源码树的差异未做对账。
- **基线台账 pending 缺口未裁决**：4 个 agent 动词 CLI `--json` 接受度与 `sync.shells` 业务失败缺例仍登记在案，未补证据。
- **环境条件 skip**：本评审 worktree 无 `mapping.json`、未挂基线仓库 / 标签，3 条条件用例跳过；换环境复跑前不要把这 3 条计为产品缺陷。

## 7. 父仓库登记项（只合入规范 v1.4，其余未动）

- 父仓库 `SKILL管理规范.md` v1.3 → v1.4（Phase 0 六项止血，提交 `4d49fbf`）**已合入父仓库 `refactor/v1.6`**，merge commit `4a4694b`，**本地未 push**。
- 父仓库登记的子模块指针仍为 `d0e9c2c`（未随本评审 worktree 的 `6ee66d0` 前移）；安装副本与远端未动。
- 父仓库两个未跟踪文件仍在、未提交：`vba-workbench 整体架构优化方案 1.6-1.9.md`、`vba-workbench-1.6.0-rc1-6ecc08b.zip`。本轮验收不修改父仓库及其未跟踪文件，仅在此登记其存在与状态。
