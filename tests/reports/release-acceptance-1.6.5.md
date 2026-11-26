# 1.6.5 发布验收报告（review/v1.6-local 本地评审分支）

- 对应提交：分支 `review/v1.6-local`；修复提交 **`6606880`**（`fix(1.6.5): 自有实例收尾纵深收口与 style 警告口径过渡`，14 files / +1695 −73）；文档条目同提交号回填于 CHANGELOG 1.6.5 段。
- 口径：本报告只登记 1.6.5 对 1.6.4 评审遗留项（N1–N7）的落地验收。1.6.4 定向修复见 `release-acceptance-1.6.4.md`，历史 1.6 / 1.6.1 / 1.6.2 / 1.6.3 报告未改动、不重复改写。
- 边界：验收过程不终止任何用户 Excel 进程（用户进程 PID 15220 全程未被终止）、不打开/保存用户工作簿；`tests/.last-full-run.log` 为本地运行工件，未纳入提交。

## 1. 验收结论摘要

| 步骤 | 命令 / 口径 | 结果 | 退出码 | 结论 |
|---|---|---|---|---:|
| 编译 | `python -m compileall -q scripts tests` | 无语法告警 | 0 | 通过 |
| 全量 unittest | `python -m unittest discover -s tests` | **Ran 542，OK，skipped 4，0 failure / 0 error**（76.98 s） | 0 | 通过 |
| N3（P1）自有实例收尾统一 | 全仓回收路径切 `quit_owned` | 5 处路径 + 2 处收尾全部落地；全仓仅剩 2 处 `Quit`（`quit_owned` 内部、开发辅助脚本） | — | 通过 |
| N1 归属快照 / N2 簿白名单 / N7 ROT 显式收口 | `toolkit/excel.py` | 落地并有 24 例专项用例 | — | 通过 |
| N4 还原未生效入错误 | `toolkit/app_guard.py` + `agent_io` / `sync_push` 收尾接线 | 落地，7 例专项用例 | — | 通过 |
| N5 style 警告口径 | `features/agent_patch.py` 双缓冲 + 前缀归一 | 落地（**过渡口径**），7 例专项用例 | — | 通过（口径待裁决） |
| N6 run --isolate 提示对齐 | `features/run_macro.py` + attach 侧三处默认 suggested | 只改提示层，判定白名单未动；4 例专项用例 | — | 通过 |
| 安装（用户口径：直接复制） | 交付件直接覆盖 `<安装副本>` | `VERSION=1.6.5`、交付件 73 件、`__pycache__` 0 | — | 通过 |

> 规模口径：全量 **542 例 / 538 通过 / skipped 4 / 0 failure / 0 error**；较 1.6.4 的 500 例 +42 例，全部来自本轮新增用例。命令面与 1.6.4 等价，未新增业务动词。

**总判定：可交付（可提交）**；N1–N7 全部落地，剩余为产品口径裁决项与端到端未复跑项，见 §5，不记为已完成。

## 2. 本轮闭环项逐项验收

### 2.1 N3（P1）自有实例收尾统一走 `quit_owned`

- **问题**：`common_subset.close_owned`、`agent_io` 的 `book open` 失败回收与 `_quit_if_empty`、`sync_push` 两处收尾、`run_macro` 隔离开表失败分支，仍是裸 `Quit()` + `except: pass`：dirty 工作簿可能触发隐藏保存提示、进程滞留、留 `~$` 锁文件，且失败无任何诊断。
- **改法**：全部改走 `toolkit.excel.quit_owned`，返回的问题清单按各自功能既有的 warnings / stats 通道上报（`isolate_cleanup=ok` / `isolate_cleanup_problems`），未新增错误码、未改退出码。
- **验证证据**：`tests/test_owned_lifecycle_1_6_5.py` 中回收诊断用例（close_owned 剩余簿 / Quit 失败、run_macro 失败回收与参考簿保护、sync_push 单项目与批量聚合、agent_io guard 诊断）；全仓 `*.py` 扫描 `.Quit()` 仅剩 `excel.py` 内 `quit_owned` 收尾调用与 `scripts/make_lab_workbook.py`（开发辅助脚本，已判定保留：引入 toolkit 依赖不划算、无诊断通道、且实例自建自关，见 §5）。

### 2.2 N1 归属快照复核

- **问题**：`quit_owned` 只信调用方传入的对象，收尾前不校验该 COM 代理是否仍指向本工具创建的实例。
- **改法**：`new_excel` 起实例登记 `pid` / `hwnd` 快照；`quit_owned` 收尾前重取比对，字段不一致则**不碰簿、不 Quit**，问题清单记「归属校验不通过」。无快照的旧调用路径（attach 借来的实例、桩对象）保持旧行为，不因缺快照报错。
- **验证证据**：快照不一致不 Quit 且报问题的专项用例；`test_regression_1_6_4` 的 quit_owned 两例意图未改、仍绿。

### 2.3 N2 簿白名单

- **问题**：`quit_owned` 无条件关掉实例内全部工作簿，宏在隔离实例里自行 `Workbooks.Open` 的参考簿会被无保存关闭。
- **改法**：新增 `keep_workbooks`：名单内簿保留不关；关完后名单外仍有簿则**不 Quit** 并上报剩余簿名——宁可留实例，不吞他人数据。隔离路径已传白名单（取「除目标表外的现存簿」，依据：隔离实例由 `new_excel` 新起、启动时簿集为空，现存簿中除本次打开的目标表外只能是宏侧自开的参考簿）。`keep_workbooks=None` 保持 1.6.4 旧行为（关全部簿后退出）。
- **验证证据**：白名单外仍有簿 → 不 Quit 且报剩余簿名的用例；隔离路径参考簿保护用例。

### 2.4 N7 ROT 不可达记录显式收口

- **问题**：模块级 `_LAST_ROT_UNREACHABLE` 只能靠「下一轮枚举顺带 clear()」这种隐式约定，测试隔离 / 同进程二次 attach 有残留风险。
- **改法**：新增 `reset_rot_unreachable()` 显式收口；枚举开头清空与 1.6.4 的 `COM_CONTEXT` 分流语义**不变**（选型理由：改动面最小、不碰分流口径）。
- **验证证据**：`BindToObject` 抛错真实落表 + 下一轮枚举清空旧条目 + 显式 reset 用例。

### 2.5 N4 还原未生效记入错误

- **问题**：`AppGuard.restore` 赋回关闭后回读仍是开启态（Excel 静默吞赋值，如弹窗占用、受保护视图）时，既不记 `changes` 也不进 `errors`，`describe()` 还报「无需还原」——失败被弱化。
- **改法**：回读仍 on → `errors` 追加「还原 <attr> 未生效：回读仍为 …」，按整条文案去重；`describe()` 末尾附未生效条目；`_is_off` 归一口径与两属性独立迭代不动。`agent_io` 与 `sync_push` 收尾处把 `AppGuard.errors`（截 3 条）接进 warnings，有改动附 `describe()` 行，既有人读输出未动。
- **验证证据**：`tests/test_app_guard_1_6_5.py` 7 例（`CutCopyMode="FALSE"` 归类、回读仍 on 入 errors、重复 restore 不刷屏、`_is_off` 边界参数化、自定义文本还原 changes/describe）；`test_regression_1_6_4` 的 AppGuard 两例仍绿。

### 2.6 N5 style 警告口径（过渡）

- **问题**：1.6.4 为修 JSON 模式 `stats.lines` 为空把 stdout/stderr 并入同一缓冲，副作用是非 JSON 模式下 `warn()` 的 `[警告]` 分流被并入报告行，通道语义与旧版不同。
- **改法**：不回退截入机制，改为 `out_buf` / `err_buf` 双缓冲：人话模式警告行并入报告时保留 / 补齐 / 去重为单个 `[警告]` 前缀，`[错误]` 行不被误标成警告；JSON 模式只收报告行，警告走 `stats.warnings`，错误 / traceback 不进信封。
- **验证证据**：`tests/test_agent_patch_1_6_5.py` 7 例（非 JSON 含警告带前缀、`[错误]` 不误标、JSON 成功 / 失败 `stats.lines` 非空且 stderr 为空、重定向异常不被吞且流正确还原）。
- **口径状态**：**产品口径未裁决**，本轮只做到「前缀可辨识」，未重新分流警告通道。

### 2.7 N6 run --isolate 阻断提示对齐

- **问题**：隔离预检只放行 `NO_EXCEL` / `TARGET_NOT_OPEN`，1.6.4 新增的 `COM_CONTEXT` 会阻断隔离副本，但提示还在叫人「再用 --isolate」，自相矛盾。
- **改法**：只改提示层——hint 说明三类成因（弹窗 / 受保护视图 / 异权会话）并明确「隔离模式也不会启动」，suggested 改为可复制的 `run --project … --macro … --isolate --apply`；attach 侧三处网关默认 suggested（rot / ready / books）同步改为「先关卡住实例重试附着；该表仍不可用时换一张能打开的表再走 isolate」。放行白名单、错误码归类、退出码均未动。
- **验证证据**：`tests/test_run_1_6_5.py` 4 例（`COM_CONTEXT` 提示措辞与可执行 suggested、两条放行路径口径未坏、JSON 信封单行可解析）；`tests/test_sync_push.py` 的默认 suggested 契约字面量随口径同步（断言仍是逐字比对，强度未放宽、未删用例）。

## 3. 全量回归结论

- 修复提交 `6606880`；`python -m compileall -q scripts tests` 退出 0。
- `python -m unittest discover -s tests` → **Ran 542 tests，OK（skipped=4），0 failure / 0 error**，76.98 s。
- skipped 4 逐条构成（与 1.6.4 一致）：

| # | 用例 / 原因 | 性质 |
|---|---|---|
| 1 | `test_json_contract` assert 失败人话提示（硬 skip）：B5 已复核为接受设计差异，不恢复旧 exit 2 壳层口径 | 已接受设计差异 |
| 2 | `test_json_contract.test_audit_overview_stdout_is_single_json_line`：要求仓库根存在 `mapping.json`，本 worktree 根无配置 | 环境条件 |
| 3 | `test_release_diff.RealBaselineTest.test_ledger_recorded_and_version_present`：基线标签在本 worktree 不可用 | 环境条件 |
| 4 | `test_release_diff.RealBaselineTest.test_real_version_scenario_matches`：同上 | 环境条件 |

## 4. 安装（用户口径：直接复制交付件，不做打包 / 哈希对照）

- **方式**：按用户口径放弃安装器与 manifest 核对，直接把 review 工作区的交付件复制覆盖到 `<安装副本>`（`SKILL.md`、`VERSION`、`LICENSE`、`requirements.txt` 四文件 + `assets` / `references` / `scripts` 三目录镜像同步）。
- **结果**：`VERSION` = **1.6.5**；交付件 **73** 件（SKILL.md 1 + VERSION 1 + LICENSE 1 + requirements.txt 1 + assets 1 + references 12 + scripts 56）；`__pycache__` 0 个；顶层仍为 `.backup` / `assets` / `references` / `scripts` / `.installed.json` / `LICENSE` / `requirements.txt` / `SKILL.md` / `VERSION`。
- **冒烟**：`sync.py --version` 与 `agent.py --version` 均输出 `vba-workbench 1.6.5`（退出 0）；`sync.py --help` 正常输出。
- **已知偏差（如实登记）**：`.installed.json` 仍记录 1.6.4 与旧文件哈希，按本轮「不做 manifest / 哈希对照」的口径**未刷新**，因此 `sync.py doctor --verify-package` 现在会报不一致。若后续要恢复包完整性门禁，需用安装器重装一次或重新生成清单。
- **边界**：用户 Excel 进程 PID 15220 未被终止；未打开/保存 `<测试工作区>\` 下任何工作簿。

## 5. 未完成 / 未验证事项

- **产品口径待裁决**：N5 警告通道是否独立恢复（当前过渡为「并入报告 + 前缀可辨识」）。
- **两处 suggested 未改**：`excel.py` 中 `EXCEL_BUSY` 与重试耗尽的 `COM_CONTEXT` 两处仍带「仍不行再用 --isolate 或批量模式」措辞；裁决认为是 BUSY 场景下隔离副本作为退路的合理表述，本轮未动，也无测试钉死。
- **开发辅助脚本保留裸 `Quit`**：`scripts/make_lab_workbook.py:41` 判定不改（零 toolkit 依赖、实例自建自关、无 warnings/stats 通道承接诊断）。
- **白名单匹配键**：按簿名 `Name` 大小写不敏感匹配，未用 `FullName`；同名多簿场景未细化。
- **快照键假设**：归属快照以 `id(xl)` 为键，适用 CLI 单线程短生命周期；长驻 / 跨线程模型需另议。
- **agent_io 收尾接缝**：为兼容既有测试对「模块级 `app_guard.restore` 调用序」的校验，还原走「guard 实例 + 模块级入口」两遍（幂等、不改变目标状态）；若要单边收口需联动改接缝测试。
- **端到端未复跑**：本轮未在带真实 Excel 会话的环境复跑写回 / watch 增量 / 真实宏执行 / VBE 弹窗时序；真实项目（示例项目）仅 1.6.4 阶段复跑过 P0-1。
- **未 push 远端**：按用户约束只提交到本地 `review/v1.6-local`。
