# 1.7 待办清单（1.6 验收遗留）

> 来源：1.6 最终验收（HEAD `432eaa8`）登记跳过与差异项。
> 验收标准（1.6 变更）：基准为**主要功能等价**，不要求与 1.5.6 逐字兼容；新增信封键、错误码细分属**已接受设计差异**，保留新设计。本清单只记 1.7 能力增强项与已关闭项，不含 1.6 已功能等价的命令面 / 参数 / 退出码。
> 脱敏：项目名、工作簿名、工作表名、路径以 `<project>`、`<module>`、`<path>`、`示例表A/B/C/D` 占位。
> 状态说明：本文是 1.6.3-local 评审台账，保留 B1–B6 / P1–P2 的处理过程；1.6.8 全量 566 例（skipped=1）通过，1.6.9 新增 scrub 9 例定向通过，全量待专责。历史数字只用于追溯，不作为现行基线。

## 目录

- 一、延后回补的能力缺口（B1–B6）
- 二、已关闭项（P1–P2）
- 三、环境条件用例
- 四、未闭环项登记

## 一、延后回补的能力缺口（1.7 增强，不阻塞 1.6 验收）

### B1. 线程看门狗基础函数外提 —— 已在 1.6.3-local 闭环（commit `0e5d047`）

- **触发场景**：`agent assert` 运行断言宏时需要硬超时保护；旧版由公共上下文管理器
  `run_with_watchdog` 包裹宏执行，超时抛 `WatchdogTimeout`，并在等待期间按 PID 存活改判
  Excel 已关闭（见 B4）。
- **影响面**：当前看门狗逻辑内联在断言路径，超时结论经 `WATCHDOG_TIMEOUT` 表达；
  缺少可被其它动词（`sync run`、`agent metrics` 等）复用的公共基础函数，
  也无法独立测试超时 / 关闭两条分支。内部公共 API 缺口；超时结论本身用户可见。
- **闭环情况**：已新增 `toolkit/watchdog.py`：`WatchdogTimeout` 异常类型、
  `run_with_watchdog(fn, timeout, ...)` 守护线程基础函数（时钟 / 休眠可注入）、`pid_alive` 三态存活探测、
  `pid_from_hwnd` 句柄反查 PID；`test_watchdog.py` 看门狗两条旧跳过用例恢复运行，并补超时禁用同步执行用例。
- **判定备注**：基础函数本身不抛业务失败码，超时后的错误码分类由上层编排层完成；COM 对象不能跨线程复用，
  包 COM 调用时 `fn` 内部需重新 attach 会话。

### B2. win32gui 弹窗 / 模态识别族 —— 已在 1.6.3-local 闭环（commit `0e5d047`）

- **触发场景**：宏运行期间 Excel 或 VBE 弹出模态框（编译错误 `#32770` 弹窗、通用 VBE
  模态框、Excel 模态对话框），旧版通过 win32gui 窗口枚举识别标题与所属 PID，
  分别落到编译错误、误判排除、模态阻塞等结论。
- **影响面**：识别基础函数未迁入 scripts，相关断言（VBE 编译错误弹窗、通用 VBE 弹窗、
  Excel 模态框 3 条）保持 skip；弹窗出现时只能以 `RUN_FAILED` / `EXCEL_CLOSED` 等
  粗略规则兜底，可能误判原因。内部识别基础函数；误判结果影响用户可见结论与建议提示。
- **闭环情况**：已新增 `toolkit/dialogs.py`：`WindowInfo` / `DialogMatch` 纯数据快照，
  `classify_window` 纯函数分类（编译错误 / VBE 通用模态 / Excel 模态 / 未知），
  `detect_dialog` 按目标 PID 过滤并按「编译错误 > VBE 通用 > Excel 模态」取最具体弹窗，
  `enumerate_windows` / `detect_target_dialog` 才做真实 win32gui 枚举。
  `test_watchdog.py` 弹窗分类族 8 例 + `test_json_contract.py` 三条旧跳过弹窗用例恢复运行。
- **错误码映射**（由 `features.agent_patch` 超时分支完成）：编译错误框 → `VBE_COMPILE_ERROR`（`stage=vbe`）；
  通用 VBE / Excel 模态框 → `WATCHDOG_TIMEOUT`（`stage=assert`），`hint` 点名弹窗标题与正文关键片段。
  旧 `VBE_DIALOG` / `MODAL_DIALOG` 细分码不在当前错误码枚举中。

### B3. `MACRO_NOT_AVAILABLE` / `MACRO_DISABLED` 分流码 —— 复核完成，维持已接受设计差异

- **触发场景**：运行宏时目标宏在工程中不存在，或宏被安全设置 / 工程保护禁用。
  旧版用两个独立错误码区分这两种情形；1.6 撤销分流，统一收口 `RUN_FAILED`，
  分类器为 `toolkit.macro.MacroRunner.classify`。
- **影响面**：调用方无法从错误码区分"缺宏"与"禁用宏"，提示与重试策略只能走 `hint` 文案。
  用户可见（错误码变粗）。
- **复核结论**：**收口到 `RUN_FAILED` 的新设计已接受，不恢复旧分流码**。区分信息固定落到
  `errors[].hint`：缺宏写清宏名 / 模块名与编译问题线索，禁用宏写清信任位置 / 宏安全设置，
  两条 hint 文案可区分；歧义句两类线索都在，仍收口 `RUN_FAILED`。
  旧跳过用例 `test_macro_run_error_classified` 复核后恢复运行，显式断言不再落回两个旧分流码。
- **判定备注**：此项是复核确认已接受设计差异，不是把旧码修回来；调用方仍只能从 hint 区分缺宏与禁用宏。

### B4. 超时按 PID 存活改判 `EXCEL_CLOSED` —— 已在 1.6.3-local 闭环（commit `0e5d047`）

- **触发场景**：看门狗等待宏返回期间，Excel 进程 / 目标实例被外部关闭；旧版按 PID
  存活探测区分"真超时"与"实例已关闭"，后者走 `EXCEL_CLOSED` 而不是超时码。
- **影响面**：该分支随看门狗基础函数一并移除（见 B1）；实例关闭与硬超时目前都收敛到
  超时 / 关闭类结论，定位阶段偏粗。内部探测基础函数；关闭结论用户可见。
- **闭环情况**：随 B1 一并回补 `pid_alive` 三态探测；超时二次分类优先级为：
  目标 PID 已退出 → `EXCEL_CLOSED`（`stage=assert`，**优先于超时码**）；
  仍存活（或探测未知）→ 交 `toolkit.dialogs` 认弹窗；判不出实例状态也认不出弹窗 →
  回退 `WATCHDOG_TIMEOUT`（`stage=assert`）。`test_assert_timeout_excel_closed_has_code` 恢复运行，
  另补存活无弹窗超时、存活编译错误框两条分支用例。pywin32 缺失 / 句柄被拒 / 查询异常一律降级为「未知」，
  不把权限问题误报成崩溃。

### B5. 旧 main 层合成 `ASSERT_FAILED` exit 2 规则（已接受设计差异）—— 复核完成，维持已接受设计差异

- **触发场景**：旧版 `agent assert` 的 handler 返回整数 2 时，main 层壳把它合成
  `ASSERT_FAILED` 信封并以退出码 2 透出；1.6 该壳层规则消失，assert 业务失败走
  Outcome 错误对象、**退出 1**，错误信息与 actionable hint 保留，等价路径由 `test_agent_patch` 覆盖。
- **影响面**：依赖退出码 2 判定"断言未通过"的外部脚本 / 自动化需要改判。
  用户可见（退出码规则变化）。
- **复核结论**：**已接受设计差异**——assert 失败退出码 1（原 2），错误信息与 actionable hint 保留；
  **不补回 exit 2 壳层合成**。旧跳过用例 `test_agent_assert_failed_has_actionable_hint` 维持硬 `@unittest.skip`，
  旧 `exit_code=2` 断言不恢复；等价路径由 `test_agent_patch.test_assert_failure_exit_1_with_code` 覆盖。
- **判定备注**：此项是复核确认已接受设计差异，不是修好；依赖旧退出码 2 的外部调用方需改判退出 1。

### B6. `sync -i` 交互菜单本体 —— 已在 1.6.3-local 闭环（commit `cf30b95`）

- **触发场景**：旧版 `sync -i` 进入多表选择交互菜单（项目 / 模块选择、参数询问）；
  1.6 已注册 `-i` 参数，非 TTY 按旧版规则给可读提示退出 1，`_interactive` 为桩，提示改用子命令。
- **影响面**：交互式选项目 / 选模块的入口暂缺；脚本化子命令路径不受影响。
  用户可见（菜单入口暂缺）。
- **闭环情况**：已新增 `features/sync_interactive.py`：列项目（来自 mapping）→ 选项目 / 全局工具 /
  新建项目 → 列动作（既有登记子命令）→ 只读动作直接运行、写操作自动预览 → 回车执行 → 复用 registry
  登记的既有 handler。菜单层不碰 COM、不写入文件，只做选择与参数装配；`reader` / `writer` / `runner`
  均可注入；非 TTY 终端把关仍由入口壳 `ensure_interactive` 负责；执行时才给 handler 加 `--apply`，
  `watch` 先问监听秒数。新建项目从拖入的 `.xlsm` 自动推导项目名、目录和版本规则。
  `test_core_sync_entry.py` 新增 `SyncInteractiveMenuTest` 12 例，
  `test_cli_contract.test_sync_interactive_non_tty_is_clean` 旧跳过用例恢复运行。

### 补充：本轮结构拆分与基线台账（非 1.7 能力项，登记归属）

- audit 切片拆分：`audit_engine` 1349 → 142 门面，拆出 `audit_rules` / `audit_scan` /
  `audit_structure` / `audit_indicators`，commit `fe170ef`。
- push 切片拆分：`sync_push` 1232 → 771，拆出 `sync_preflight` / `sync_shells` / `sync_init`，
  commit `d04fd48`。
- run / agent_io 切片拆分：`run` 992 → 312（`run_macro` / `run_watch`），
  `agent_io` 888 → 692（`agent_probe` / `agent_logs`），commit `8c7d79f`。
- 1.5.6 命令覆盖台账：`tests/baseline/COVERAGE.md` + `1.5.6/coverage.json`（31 动词 / 124 条证据），
  commit `c219a35`；`test_release_diff.py` 基线不可用分支 `skipUnless` → `skipTest`，commit `0b02fec`。

## 二、已关闭项（判定规则变更：保留新设计）

### P1. `stats.timings` 等新增键

- **结论：已关闭，保留新设计。** 新增子键作为向上兼容的设计差异保留，不删键、
  不回退到逐字兼容；`timings` 等诊断键定位为内部字段，后续若需对外承诺稳定性，
  再单独立项。
- **理由**：1.6 验收基准改为主要功能等价，新增字段不构成阻塞。

### P2. 旧 `FAILED` 兜底码被具体码取代

- **结论：已关闭，保留新设计。** 具体登记码取代旧兜底码的方向保留，不回退到
  `FAILED`；B3 分流码收口、B5 退出码规则变化同属已接受设计差异。
- **理由**：错误码细分定位更精确，按新规则不要求与旧兜底码逐字一致；
  调用方"旧码 / 旧退出码 → 新码 / 新退出码"迁移对照可在 1.7 文档化，
  不作为 1.6 阻塞项；`UNEXPECTED` 仍保留为未预期异常兜底。

## 三、环境条件用例（非产品问题，仅记录）

> 1.6.3-local 评审时 skipped=**4**，构成如下；现行 1.6.8 全量 skipped=1（仅设计差异硬跳过），其余条件已恢复或另行处理。
>
> 1. **1 条已接受设计差异硬跳过**：`test_json_contract.test_agent_assert_failed_has_actionable_hint`
>    （B5，旧 exit 2 断言不恢复，等价路径由 `test_agent_patch.test_assert_failure_exit_1_with_code` 覆盖）；
> 2. **1 条缺仓库 `mapping.json`**：`test_json_contract.test_audit_overview_stdout_is_single_json_line`
>    （`skipUnless`，仓库根无配置文件时跳过；带配置环境复跑即恢复）；
> 3. **2 条基线标签不可用**：`test_release_diff.RealBaselineTest`
>    （`test_ledger_recorded_and_version_present`、`test_real_version_scenario_matches`，
>    基线仓库 / `baseline-1.5.6` 标签不可用时 `skipTest`；具备基线环境自动恢复对比）。
>
> 后 3 条为环境条件，不计产品缺陷；CI 环境缺配置或缺标签，skip 数会相应增加。

## 四、未闭环项登记（1.6.3-local 仍遗留）

| 项 | 状态 | 说明 |
|---|---|---|
| 4 个 agent 动词 CLI `--json` 端到端接受度 | pending，待裁决 | `agent.assert` / `agent.patch` / `agent.metrics` / `agent.style` 的 CLI `--json` 接受度为真差异，未登记白名单（`test_release_diff.KNOWN_REAL_GAPS`）；进程内单行信封渲染有测试，端到端行为待 07 切片复核裁决后补录台账 |
| `sync.shells` 业务失败用例 | 缺例 | 仅有 create/clean 预览与非法 `--mode` 用法错，缺目标缺失 / 预检失败类业务失败证据（台账 `business_fail` missing 动词） |
| 三类 golden 未录制 | 待带 Excel 会话环境录制 | Excel 状态 golden（CLOSED / OPEN_CLEAN / OPEN_DIRTY / OPEN_PROTECTED / OPEN_BUSY / READONLY / COM_UNAVAILABLE）、文件副作用 golden（文件树前后 / 备份批次 / journal / 写入文件哈希）、真实宏执行 golden（宏结果 / 断言重试计数 / 保存回滚链路）；本轮不启动 Excel、COM 全打桩，待本地集成环境录制回填 `tests/baseline/1.5.6/` |
