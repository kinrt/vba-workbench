# 1.6.4 发布验收报告（review/v1.6-local 本地评审分支）

- 对应提交：分支 `review/v1.6-local`；修复提交 `b52f6d1`（fix(1.6.4) 隔离实例收口与 attach 失败分流，13 files / +303 −30）；文档提交 `1bfc422`（docs: CHANGELOG 新增 1.6.4 条目并回填提交号）。本报告数字按本轮工作区实测；文档提交不改代码与测试。
- 口径：本报告只登记 1.6.4 定向修复验收；1.6.3 命令与数据流统一化见 `release-acceptance-1.6.3.md`，1.6.2 八项修复与命令绝对化见 `release-acceptance-1.6.2.md`，1.6.1 六项修复见 `release-acceptance-1.6.1.md`，1.6 分层架构见 `release-acceptance-1.6.md`，不重复改写。
- 独立评审：仓库外 `review-notes/diff-review-1.6.4.md`（只读静态评审，未重跑测试、未触碰用户 Excel 进程与工作簿）。验收过程不终止用户 Excel 进程、不打开/保存用户工作簿；运行日志为本地工件，未纳入提交。

## 1. 验收结论摘要

| 步骤 | 命令 / 口径 | 结果 | 退出码 | 结论 |
|---|---|---|---:|---|
| 编译 | `python -m compileall -q scripts tests` | 无语法告警 | 0 | 通过 |
| 全量 unittest | 全量 500 例 | 496 通过 / skipped 4 / 0 failure / 0 error，`OK` | 0 | 通过 |
| 独立评审 | `diff-review-1.6.4.md` | 可提交，0 阻断；4 条重点风险均判非阻断；7 条非阻断建议；缺测点 4 类 | — | 通过（附遗留） |
| 真实项目实测 | 示例项目，工作簿 `示例表 V1.xlsm`，复跑 P0-1 | `isolate_cleanup=ok`；无新增 Excel 实例；无 `~$` 锁文件 | 0 | 通过 |
| 安装升级 | `upgrade_install.py --apply`，目标 `<安装副本>` | 覆盖 73、备份 73（外置）、清理孤儿 0、自检通过、用时 1697 ms | 0 | 通过 |
| 安装副本核验 | `--version` / `doctor --verify-package` / `list --json` / 两条 `--help` | 5 项核验全部退出码 0；doctor `ok:true`，73 件在位、哈希一致、版本匹配；list 5/5 | 0 | 通过 |

> 规模口径：全量 **500 例 / 496 通过 / skipped 4 / 0 failure / 0 error**。命令面仍为 **31 动词**，manifest 未改。

**总判定：可交付（可提交，0 阻断）。** 评审 4 条重点风险与 7 条非阻断建议均不构成本轮阻断；遗留项见 §7，均标注归属与优先级，后续随 1.6.5 / 1.7.0 跟进。

## 2. 本轮闭环项逐项验收

### 2.1 隔离实例只回收自身拥有的实例（`quit_owned`）

- **问题**：旧实现对工具自起的实例直接 `Quit()` 且异常全吞；宏写过工作表后工作簿为 dirty，AppGuard 又把 `DisplayAlerts` 还原为 True，`Quit` 会弹出“是否保存”的隐藏对话框，进程滞留后台并留下 `~$` 锁文件；收尾失败也无诊断上报。
- **改法**：`toolkit/excel.py` 新增 `quit_owned(xl, save_changes=False)`：先置 `DisplayAlerts=False`、逐簿 `Close(SaveChanges=False)`，再 `Quit`；返回未清干净的诊断列表，调用方把诊断写进 warnings，不再静默吞异常。`ExcelSession.close()` 与 `run --isolate` 收尾改走该原语：收尾正常记 `stats.isolate_cleanup=ok`，有问题写 `stats.isolate_cleanup_problems` 并进 warnings。归属判断仍在调用方：attach 分支返回 `owned_xl=None`，不会把借来的用户实例传进回收路径。
- **验证证据**：提交 `b52f6d1`。`tests/test_regression_1_6_4.py` 新增 `test_quit_owned_closes_dirty_workbook_before_quit`（先关簿再退出、关闭不保存、`DisplayAlerts` 置 False）与 `test_quit_owned_reports_quit_failure`（`Quit` 抛错进诊断、关簿动作已执行）。

### 2.2 ROT 可见但绑定失败按 `COM_CONTEXT` 分流

- **问题**：ROT 里看得到目标工作簿、但 `BindToObject` 失败时，旧口径会被后续 `GetActiveObject` 误导，降级成“没打开”的 `TARGET_NOT_OPEN`，掩盖实例卡在弹窗 / 受保护视图 / 异权会话的真实情况。
- **改法**：枚举 ROT 时记录“看得到 moniker、`BindToObject` 失败”的工作簿；目标落在该清单时，后续 `GetActiveObject` 失败或附着后找不到表均报 `COM_CONTEXT/stage=attach`，hint 指向弹窗 / 受保护视图 / 权限不同的会话，suggested 指向关卡住实例或 `--isolate`，不再降级为 `TARGET_NOT_OPEN`。
- **验证证据**：提交 `b52f6d1`。回归用例 `test_rot_target_bind_failure_is_com_context`（注入不可达记录 → 抛 `COM_CONTEXT`，message 含 ROT 与目标工作簿名）。

### 2.3 未开表兜底文案中性化，附着入口按动词对齐

- **问题**：`TARGET_NOT_OPEN` 的 suggested 原先带“push 默认自动开隐藏实例”口径，与 `run` / `agent` / `watch` 各入口不符；各附着失败提示共用中性兜底串，出口不可执行。
- **改法**：`TARGET_NOT_OPEN` 的 suggested 改为“先打开目标工作簿，或用 `book --action open --project <项目> --apply`”；`run` 附着失败的 hint / suggested 改为指向 `run --isolate`（带项目与宏名），不再指向 `list`；`agent.py book --action close` 附着失败提示为“表没打开：无需关闭；要打开目标表用 `book --action open`”；`agent_io._attach_or_open` 与 `run_watch` 按各自入口给 hint 与可执行 open 命令。
- **验证证据**：提交 `b52f6d1`。回归用例 `test_target_not_open_fallback_has_no_push_copy`（兜底 suggested 不含 push 口径、含 `book --action open`）；`test_run.py` / `test_agent_io.py` 的 attach 桩改为接受关键字参数，匹配新增的 `target_hint/target_suggested` 可选关键字，签名向后兼容。

### 2.4 style 路径同时捕获 stderr

- **问题**：`agent_patch` 的 `style check/fix` 原先只 `redirect_stdout`；JSON 模式下 `msg` 固定写 stderr，报告会漏到 stderr、`stats.lines` 反而为空。
- **改法**：stdout / stderr 一并截入同一缓冲，报告统一进 human 与 `stats.lines`，JSON 出口保持干净；信封由 `render()` 在重定向退出之后才打印，机器可解析契约未破坏。
- **验证证据**：提交 `b52f6d1`。`test_agent_patch.py` 风格成功路径新增断言：`stats.lines` 非空、`stderr` 缓冲为空。

### 2.5 状态栏 / CutCopyMode 假值不再误报还原

- **问题**：COM 可能把 False 读回字符串 `"FALSE"`，旧口径把它当成“仍处于开启态”，记成一次已做的还原；归一只认 `False/0/None/""`，覆盖不到 COM 假值。
- **改法**：`AppGuard` 对 `StatusBar` / `CutCopyMode` 的“已关闭”做归一（False / 0 / None / 空串 / `"false"` / `"0"` / `"none"`，含字符串 `"FALSE"`）；归一后视为关闭、不记改动；设置后回读确认确实落成关闭才记 `changed`，避免把一次无效设置记成“已还原”。两个属性在同一循环内独立迭代、无共享状态。
- **验证证据**：提交 `b52f6d1`。回归用例 `test_string_false_is_not_a_change`（`"FALSE"` 不算改动）与 `test_real_statusbar_is_still_restored`（自定义状态栏文本仍被还原并记入 `changes`）。

### 2.6 回归测试与既有断言同步

- **改法**：新增 `tests/test_regression_1_6_4.py` 共 6 例（2.1 两条、2.2 一条、2.3 一条、2.5 两条）；`test_run.py` / `test_agent_io.py` / `test_agent_patch.py` 断言随签名与口径同步更新。测试件只落临时目录，COM 边界全部打桩或桩对象，不启动真实 Excel。
- **验证证据**：提交 `b52f6d1`；全量结果见 §3。

## 3. 全量回归结论

- 修复提交 `b52f6d1`；全量复跑：**500 例 / 496 通过 / 0 failures / 0 errors / skipped 4 / exit 0**；`python -m compileall -q scripts tests` 退出 0，无语法告警。
- **skipped 4 条逐条构成**（本评审环境实测）：

| # | 用例 | 跳过方式 / 原因 | 性质 |
|---|---|---|---|
| 1 | `test_json_contract` · assert 失败人话提示用例 | 硬 `@unittest.skip`：B5 已复核为接受设计差异，旧 main 合成 `ASSERT_FAILED` 信封（exit 2）壳层口径不恢复；assert 业务失败退出 1，等价路径由 `test_agent_patch` 覆盖 | 已接受设计差异，不恢复旧断言 |
| 2 | `test_json_contract.test_audit_overview_stdout_is_single_json_line` | 环境条件：要求仓库根存在 `mapping.json`，本评审 worktree 根无配置文件 | 环境条件，非产品问题；带配置环境复跑即恢复 |
| 3 | `test_release_diff.RealBaselineTest.test_ledger_recorded_and_version_present` | `skipTest`：基线仓库 / `baseline-1.5.6` 标签在本 worktree 不可用 | 环境条件，非产品问题；具备基线环境自动恢复对比 |
| 4 | `test_release_diff.RealBaselineTest.test_real_version_scenario_matches` | `skipTest`：同上，基线仓库 / 标签不可用 | 环境条件，非产品问题；具备基线环境自动恢复对比 |

- 与 1.6.3 口径一致：功能型跳过仅剩第 1 条（B5）；第 2–4 条为环境条件用例，换具备配置 / 基线标签的环境复跑即可，不计产品缺陷。

## 4. 独立评审结论

### 4.1 四条重点风险判定

| 风险点 | 结论 | 遗留 |
|---|---|---|
| 1. `quit_owned` 是否可能误关用户自己打开的 Excel 实例 | **非阻断：当前调用链不会误关用户实例**。归属判断在调用方：`ExcelSession.close` 仅在 `owned=True` 时调用；`run` 收尾仅在 isolate 分支返回 `owned_xl`，attach 分支返回 `None`；两条路径均不把借来的实例传进 `quit_owned` | N1（无归属令牌复核）、N2（关簿无白名单），见 §7 |
| 2. ROT / `_LAST_ROT_UNREACHABLE` 模块级全局状态是否会被复用污染 | **非阻断：当前单线程模型下无污染**。列表在每次 `_rot_excel_apps()` 入口先 `clear()` 再枚举；`attach_running_excel` 先枚举后读取，中间无嵌套 ROT 调用；看门狗守护线程只跑已取得代理上的宏调用，不触碰该全局 | N7（缺显式 reset 接口），见 §7 |
| 3. style 同时捕获 stderr 是否破坏机器可解析契约 | **非阻断：机器契约未破坏**。`run_checks/run_fix/section/Notes.render` 全部经 `msg/warn` 通道，无直接 `print` / `json.dumps`；被截进缓冲的都是人话行，进 `human` / `stats.lines` 字符串数组；JSON 模式 `warn` 本就收口进缓存、不写 stderr | N5（非 JSON 模式警告分流通道并入报告，口径待裁决），见 §7 |
| 4. `StatusBar="FALSE"` 归类改动与 CutCopyMode / 交互态是否互影响、是否吞真错误 | **非阻断：归类修正正确，未与 CutCopyMode 互相干扰；不存在吞真错误**。`_is_off` 归一覆盖 COM 假值；两属性独立迭代；属性读不到的既有 `except: pass` 分支保持不变 | N4（还原未生效既不记 change 也不进 errors，弱报），见 §7 |

### 4.2 冻结契约核对（未触碰）

- **P1-4（exit 2 只输出人话）—— 通过，未触碰**：本次改动未新增任何 `SystemExit` / `sys.exit(2)` / argparse 退出 2 路径；新增 / 改写的出口全部是 `Failure` 业务码。入口兜底 `run_dispatch` 对 `Failure` 一律归 exit 1；exit 2 仍只由 argparse / 裸跑帮助产生。未发现 JSON 模式混入人话信封、或 exit 2 带结构化输出的情况。
- **P2-6（audit 全仓 `ProjectMode.NONE` 不得改）—— 通过，未触碰**：改动文件不含 `audit_a.py` / `audit_b.py` / `core.py` 等 audit 规格文件；未改任何 `CommandSpec` 的 `project` 字段；既有 audit 全仓只读规格原样保留。`agent_patch.py` 的 metrics 规格 `ProjectMode.NONE` 为既有代码，本次仅改 style 重定向段。
- **版本 / 文档同步**：`VERSION`、`SKILL.md metadata.version`、`references/toolkit.md`（`quit_owned` 入清单与场景表）均到 1.6.4；CHANGELOG 1.6.4 段落与本次代码 diff 逐条对得上，未发现与代码矛盾的登记。

## 5. 安装副本升级核验

- **安装动作**：安装器 `upgrade_install.py --apply`，目标 `<安装副本>`，退出码 0：覆盖 73 个、备份 73 个（外置）、清理孤儿 0、自检通过、用时 1697 ms；本次备份批次 `upgrade-20261123-215133`，落外置目录 `<安装备份目录>\`。
- **文件对账**：安装前后文件数 74（73 交付件 + `.installed.json`）→ 74，未丢件、未增多；manifest 73 条 / 磁盘 73，missing=[]、mismatch=[]、extra=[]。
- **版本一致性**：`VERSION` 1.6.3 → **1.6.4**；`.installed.json` `installed_at` `2026-11-23T21:51:34+08:00`，`version` `1.6.4`；安装副本 `SKILL.md` `metadata.version = 1.6.4`。
- **5 项核验（全部退出码 0）**：

| # | 核验项 | 结果 |
|---|---|---|
| 1 | `sync.py --version` | 输出 `vba-workbench 1.6.4` |
| 2 | `sync.py doctor --verify-package` | `ok:true`；`stats.package` 全部通过、`version_match=true`、项目 5/5 |
| 3 | `sync.py list --json` | `total 5 / resolved 5` |
| 4 | `sync.py --help`、`agent.py --help` | 两条帮助均正常输出 |
| 5 | `agent.py --version` | 输出 `1.6.4` |

- **目录卫生**：安装副本无 `tests/`、无 `.log`、无 `__pycache__`、无冲突副本；交付树内 `.backup` 为空目录（无文件），不计交付件。
- **源码边界**：安装前后源码工作区 git status 逐行一致，安装未改动源码。
- **用户边界**：用户 Excel 进程 PID 15220 全程未被终止；未打开 / 保存用户工作簿。

## 6. 真实项目实测

- **项目**：示例项目；工作簿 `示例表 V1.xlsm`。
- **复跑口径**：P0-1（`run --isolate` 隔离路径）。
- **结果**：`isolate_cleanup=ok`；复跑后无新增 Excel 实例；无 `~$` 锁文件残留。
- **边界**：本轮真实项目实测只覆盖该隔离路径；写回、watch 增量、真实宏执行与弹窗时序未在本项目端到端复跑，见 §8。

## 7. 后续跟进（1.6.5 / 1.7.0）

以下条目均为评审登记的非阻断项与缺测点，**本轮未实现、未补测**，按归属与优先级跟进；不记为已完成。

### 7.1 非阻断建议（评审 §③，N1–N7）

| 编号 | 归属 | 优先级 | 建议 | 触发场景 |
|---|---|---|---|---|
| N1 | `toolkit/excel` | P2 | `new_excel` 记录 Hwnd/PID 快照，`quit_owned` 收尾前比对；不一致则只上报诊断、不调用 `Quit`，补一道归属纵深防御 | 自有实例异常退出后 PID 被复用、COM 代理半死、未来出现跨进程复用代理的改法 |
| N2 | `toolkit/excel` | P2 | 记录工具自己 `Open` 的簿名清单，`quit_owned` 只关清单内；或在隔离实例文档中明示“宏侧副簿不保留、不保存” | 隔离 / 磁盘副本里宏自行 `Workbooks.Open` 参考簿；批量 pull 复用隐藏实例混入非本项目簿 |
| N3 | `features/common_subset`、`features/agent_io` | P1 | 其余两条自有实例回收路径（`common_subset.close_owned` 被 `agent.push` / `sync.push` 复用；`agent book open` 失败分支）统一改用 `quit_owned`，诊断进 warnings/stats，不再裸 `Quit` + `except: pass` | `agent.push` / `sync.push` 走磁盘模式收尾；`agent book open` 打开失败回收 |
| N4 | `toolkit/app_guard` | P2 | 回读仍为 on 态时向 `self.errors.append` 记录“还原未生效”，与既有布尔目标失败口径对齐，不再弱化为“无需还原” | 状态栏 / 剪切态属性赋值被 Excel 静默忽略（弹窗占用、受保护视图等），回读仍是原值 |
| N5 | `features/agent_patch` | P2 | 确认产品上可接受警告并入报告行；否则并入时保留 `[警告]` 前缀，或 warn 仍按原通道发送、只截 msg 报告 | 人话模式跑 `agent.py style --action check/fix` 且存在风格警告时，警告分流通道并入报告 |
| N6 | `features/run_macro` + `references` | P2 | 在 isolate 提示语里说明“目标实例 COM 不可达时隔离模式也不启动”；评估是否允许 `COM_CONTEXT` 预检失败后仍走磁盘副本；文档 / 排障提示对齐新口径 | 目标表所在实例卡在弹窗 / 受保护视图 / 异权会话，同时想跑 `run --isolate` |
| N7 | `toolkit/excel` | P3 | 提供 `_reset_rot_unreachable()` 显式收口；或把该列表改为 `attach_running_excel` 局部状态、经 `find_excel_with_workbook` 返回值带出 | 未来 watch / 批量在同一进程内多次 attach；嵌套枚举；测试间全局残留 |

### 7.2 缺测点（评审 §⑤，4 类）

| 类别 | 归属风险点 | 优先级 | 待补用例要点 |
|---|---|---|---|
| 归属与回收 | 风险点 1（N1/N2 配套） | P2 | attach 成功（`owned=False`）时 `close()` 绝不 `Quit`；`run` attach 模式 `owned_xl is None`、finally 不 Quit；`quit_owned` 异常分支（`Workbooks` 读取失败、单簿 `Close` 失败、`Quit` 后仍有簿未关闭）诊断内容；归属令牌复核行为；宏侧副簿行为约定；重复调用 `quit_owned` / `close()` 幂等性 |
| ROT 全局状态 | 风险点 2 | P3 | 生产落表路径：`BindToObject` 抛错真实追加进 `_LAST_ROT_UNREACHABLE`；下一轮枚举 `clear()` 清空旧条目；成功 attach 后残留条目不影响后续 `TARGET_NOT_OPEN` 判定；路径归一不匹配（8.3 短路径、UNC、大小写、相对路径）的诊断口径；同进程重入 / 嵌套枚举隔离性 |
| style stderr 捕获 | 风险点 3（N5 配套） | P2 | 非 JSON 模式风格报告（含 warn）完整进入 human / `stats.lines`；JSON 失败路径（脏文件）信封 `stats.lines` 可解析、stderr 为空；重定向期间 `fail()` / traceback 不被吞、异常抛出后 sys.stderr 正确还原；确认 style 链路无子进程继承 stderr fd |
| app_guard 归类 | 风险点 4（N4 配套） | P2 | `CutCopyMode="FALSE"` 字符串归类为 off；回读仍为 on（还原未生效）行为断言；`_is_off` 边界参数化（`None / 0 / False / "" / "false" / "0" / "none" / 大写 "FALSE" / 非空文本 / 数字 1`）；StatusBar 自定义文本还原后 `describe()` 文案正确 |

## 8. 未完成 / 未验证事项

- **非阻断建议未落地**：§7.1 的 N1–N7 本轮均未改代码；其中 N3 为 P1 跟进项，其余按 P2 / P3 排期。
- **缺测点未补**：§7.2 四类用例本轮未补录，现有覆盖以桩对象与成功路径为主。
- **端到端真实会话未全覆盖**：本轮真实项目仅复跑示例项目 的 P0-1 隔离路径；写回、watch 增量、真实宏执行、VBE 弹窗时序、跨进程 PID 归属等场景未在带真实 Excel 会话的环境做端到端验收。
- **评审为静态结论**：独立评审基于只读 `git diff / status` 与源码通读，未重跑测试；测试结果取自本地全量运行工件，未在本报告内重放。
- **N5 通道口径待裁决**：人话模式警告并入报告行是否可接受，属产品口径确认项，裁决前不改实现。
