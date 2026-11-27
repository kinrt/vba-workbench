# 变更日志

按「对外可见的行为变化」记；每条带日期与提交号（内部重构只在提交信息里写）。

## 1.7.0 — 2026-11-27（开发中：真实环境验证前移）

**主题：把 `audit -i` 菜单、`vbamod` 正文 diff 与 `assert` 看门狗接线收进 1.7；G-01 落 Excel golden 证据工具，G-03 已做真实 Excel MsgBox 与跨进程 PID，其余不可自动构造项如实登记 not_run。**

- **`audit -i` 菜单**：新增 `features/audit_menu.py`，8 项菜单覆盖盘模块、公共库、项目依赖面、全仓、结构体检、指标体检、导出 Markdown、显示设置；只读动作结束回菜单，非 TTY 仍退出 1。
- **`vbamod` 正文 diff**：`toolkit/vbabin.py` 增加 MS-OVBA 2.4.1.3 解压与 `Attribute` 归一，`sync.py vbamod` 输出 `content_same / changed / content_compared / content_error`；模块集合一致不再等同于正文一致。
- **`assert` 看门狗接线**：真实 `agent assert` 改由 worker 线程执行，worker 内 `CoInitialize` 并重新 attach 目标簿；超时由 `run_with_watchdog` 抛出，再按目标 PID 分类 `EXCEL_CLOSED / VBE_COMPILE_ERROR / WATCHDOG_TIMEOUT`。注入 runner 的单测路径保持同步，旧 `timed_out` 桩语义不变。
- **Excel golden 证据工具**：新增 `tests/tools/excel_golden.py`，只做 `prepare / record / verify / report`；工作簿、源码、备份批次、`.run` journal、GBK 产物全部落仓库根 `.lab/golden`，原始证据不入库。已录 17 个 case：12 recorded / 5 not_run，`verify` 通过 17。
  - 已录：状态 4（closed / open-clean / open-dirty / readonly）、文件副作用 3（pull / push / gbk-export）、宏 2（isolate / assert）、弹窗 3（Excel MsgBox / 跨进程 PID / TARGET_NOT_ACTIVE）。
  - not_run：`OPEN_PROTECTED`（工程密码只能 VBE 手工设）、`OPEN_BUSY`（由 Excel MsgBox 覆盖）、`COM_CONTEXT`（不破坏本机会话强造）、`VBE_COMPILE_ERROR`（自动化路径实际 `RUN_FAILED`）、通用 VBE 模态（需 GUI 手工构造）。
- **`assert` 前台弹窗预检**：`_ensure_active` 在切目标簿前先按 PID 探测模态框；同进程有 MsgBox/InputBox 挡住时直接报 `TARGET_NOT_ACTIVE / stage=attach`，不再空转 5 次后误跑成功；编译错误框单独报 `VBE_COMPILE_ERROR / stage=vbe`。
- **版本口径**：`VERSION` 1.7.0；`SKILL.md` metadata 同步；`tests/baseline/1.5.6/coverage.json` 只回填 golden 状态与复跑命令，不向旧基线塞原始文件。
- **测试**：全量 `unittest discover` 609 例通过 / 1 skip / 0 failure（2026-11-27）。

## 1.6.13 — 2026-11-26（交互菜单重构与公开文档收口）

**主题：重做 `sync -i` 启动菜单，把新建项目改成拖入工作簿自动登记；同步补齐公开仓库的 README、架构说明与脱敏口径。**

- **菜单交互**：选中写操作后立即预览，不再先问“预览 / 执行”；预览通过后回车执行，`0` 返回；只读动作直接运行；动作结束自动回当前菜单，不再退出进程或要求二次回车返回。
- **菜单结构**：项目列表和动作菜单加分区标题、分隔线与输入前空行；默认输入行与 `0. 退出` 分开显示。
- **全局工具**：`list / version / doctor / shells / clean-backup` 从项目动作中拆出，在项目列表里单独选“全局工具”，不再要求先选项目。
- **新建项目**：支持把 `.xlsm` / `.xlsb` 拖进菜单窗口；带 ` V<数字>` 时自动推导项目名和前缀并按最高版本解析，没有版本号时按精确文件登记；同名不同目录自动加 `-2` / `-3`；建完自动抓取一次源码并进入该项目的动作菜单。
- **`init` 参数**：新增 `--file`（按精确工作簿登记）和 `--allow-outside`（显式允许登记当前配置工作区之外的工作簿）；命令行旧用法保持兼容。
- **README 与文档**：README 重写为“痛点—使用入口—能力边界—文档导航”结构，架构图改用 Mermaid；命令行效果改为文本示例，删除四张旧截图；新增 `references/architecture.md`，统一文档里的“工作簿 / 工作表”口径，补齐 `.gitignore` / `.gitattributes`。
- **测试基线**：2026-11-26 全量 591 例，`OK (skipped=1)`；菜单与 `init` 定向 24 例通过。

## 1.6.12 — 2026-11-24（四组自动回归补齐）

**主题：把 1.6.11 列出的四个测试缺口固化成自动断言；命令面、JSON 契约、写回行为一律未变。**

- **`run --save` 非隔离**：新增自动断言，确认人话/JSON 都给出“仅对 --isolate 生效”，journal 记 `save_requested=true / save_ignored=true / saved=false`，且不会真的保存工作簿。
- **`run --isolate` 撞已开表**：新增自动断言，确认错误码为 `TARGET_OPEN`，失败 journal 的 `mode=isolate`，不落 `unknown`。
- **`plan` 三态矩阵**：新增表关/表开 × pull/push/run/watch 八格矩阵，锁定 `mode / needs_open / will_save / command`。
- **`doctor` 明细契约**：新增自动断言，锁定 `details[]` 只有 `name / level / detail` 三键、`level` 仅 `ok/warn/fail`，并检查 `AccessVBOM / 包文件齐全 / 配置文件 / 仓库目录可写` 在位。
- **测试提速**：`audit overview` stdout 契约与 `manual_accel` 忽略用例改用最小沙箱，不再扫真实仓库；新增 `tests/tools/profile_tests.py` 列出最慢 30 条。全量 582 例 profile 耗时由 112.7 秒降至 65.5 秒，两条用例分别由 8.5 秒 / 7.4 秒降至 0.36 秒 / 0.03 秒。
- **A1 门禁脚本**：新增 `tests/tools/devloop_gate.py`，一次跑全量单测、编译、doctor、源码/副本脱敏、安装副本校验，报告落测试工作区 `.devloop/<时间戳>/gate.json`；2026-11-24 实测 7 项全过、约 86 秒（其中 profile 单测 65.5 秒）。门禁同时校验 `latest.txt` 指向的 zip 与 `SHA256SUMS.txt` 一致。
- **测试基线**：2026-11-24 全量 582 例，`OK (skipped=1)`，约 76 秒；定向新增 4 例全部通过。
- **安装副本**：1.6.12 已装到 `<安装副本>`；73 个交付件，哈希逐项一致；备份批次 `upgrade-20261124-153524`。
- **正式发布包**：新增 `tests/tools/make_release.py`，按安装器交付白名单生成 `releases/vba-workbench-1.6.12.zip`（73 件 / 907,099 字节 / SHA256 `01de547f1bce34fa6fe87bba7355da94e9db1e4524dc4167c93cedc642393e18`）；与安装副本 `.installed.json` 对账 73/73，无缺件、无多余、无哈希差异；`latest.txt` 已指 1.6.12。

## 1.6.11 — 2026-11-24（测试流程与文档路径校正）

**主题：把“测试命令/造表路径”校正到本机可执行口径，并修掉 1.6.10 改 `style` 明细后遗留的两处旧测试桩；命令面、JSON 契约、写回行为一律未变。**

- **全量测试命令校正**：Python 3.13 下 `unittest discover -s vba-workbench/tests` 会因起点不可导入直接报 `ImportError`；`references/release.md` 与《SKILL管理规范》统一改为从 `vba-workbench/tests` 目录执行 `python -m unittest discover -s . -p "test_*.py"`。
- **造表脚本路径校正**：`tests/MANUAL.md` 原写 `tests/make_lab_workbook.py`，实际安装包只有 `scripts/make_lab_workbook.py`，照抄会找不到文件；改为 `scripts/` 路径。
- **两处旧测试桩对齐**：`handle_style` 已改调 `run_checks_with_details`，`test_agent_patch_1_6_5` 里两处仍打桩 `run_checks`，导致全量 2 个 error；测试桩改到新函数并返回 `(exit_code, details)`，实现未动。
- **测试基线**：2026-11-24 全量 578 例，`OK (skipped=1)`，约 70 秒；1.6.10 定向 3 例仍通过。
- **安装副本**：1.6.11 已装到 `<安装副本>`；73 个交付件，哈希逐项一致；备份批次 `upgrade-20261124-152430`。

## 1.6.10 — 2026-11-24（1.6.9 验收遗留收口）

**主题：按「示例项目 1.6.9 实战验收」的 P0/P2/P3 项做小步修复，不改命令面。**

- **P0 `watch --startup` 写回计数**：启动 `src` / 交互 `each` 选源码写回原先只记日志、不计数，汇总恒为 `written=0`；现在写回成功计入 `written`，失败计入 `skipped`，与轮询写回同一口径。
- **P2 `run` JSON stderr**：`require_open_workbook` 的“目标工作簿”提示和 `backup_workbook` 的“备份”提示在 `--json` 下不再写 stderr；人话模式保持原提示，JSON 信封与 `stats.warnings` 都不再被过程话污染。
- **P2 `style` 结构化明细与修复建议**：`style --action check` 新增 `details[]`（`rule / file / line / old / new / message`，与扫描同一次完成）；失败 `suggested` 保留本次 `--which`，带项目时保留 `--project`，自定义范围保留 `--scope`，不再一律回落到默认 `all/src`。
- **P3 `clean-backup` timings**：`stats.timings` 改在阶段退出后取值，`clean-backup` 阶段耗时不再为空。
- **P3 H06 误报**：进度器模块 `表格工具.bas` 与既有 `表格工具.cls` 同列白名单；示例项目 indicators 由 16 通过 / 2 警告回到 17 通过 / 1 警告（剩下 H10 是备份积累到 23 批的数据状态提示）。
- **测试**：新增 3 例（启动写回计数、JSON 下目标簿提示不落 stderr、style details 与建议保留 `--which`）；`test_run` 33、`test_agent_patch` 37、`test_sync_transform` 23、`test_audit_gate` 11、`test_json_contract` 40（skipped=1）全部通过；全量仍交专责。
- **安装副本**：1.6.10 已装到 `<安装副本>`；73 个交付件，哈希逐项一致；备份批次 `upgrade-20261124-145637`。

## 1.6.9 — 2026-11-24（scrub 发布脱敏命令与 evals 字段校正）

**主题：把发布检查里一直靠人工的脱敏扫描做成只读命令 `sync.py scrub`，同时收掉 evals 里 `list` 字段与已撤销错误码的残留。**

- **新增 `sync.py scrub`**：读仓库根 `mapping.json` 的 `scrub.deny`，按安装器交付件白名单逐文件扫；`.bat` 走 `read_source_text` 的 GBK 回退，图片等非文本跳过。命中报 `SCRUB_HIT/stage=preflight` 并退出 1，零命中退出 0；`details[]` 每条 `file / line / word / text`，`stats` 给 `files / deny_terms / hits / matched_terms / skipped_non_text`。`--path` 可复核已安装副本，只读不写盘。
- **发布清单**：`references/release.md` 的脱敏步骤改为 `sync.py scrub --json`；扫描范围限定交付件，`tests/`、`evals/`、`.lab/`、`releases/` 不进扫描。
- **evals 校正**：`#2` 的 `list` 结果字段由 `target` 改为 `workbook`；`#17` 删除已撤销的 `VBE_DENIED`，统一 `VBE_PROTECTED`；新增 `#19` 覆盖 scrub；触发用例补 1 条发布脱敏查询。
- **测试**：新增 `test_contract_scrub.py` 9 例，覆盖命中、GBK、大小写、排除项、非文本跳过、词表去重与 handler 成功/命中/缺配置；`scrub` 作为计划内新增动词登记进 1.5.6 `approved_deltas`；命令表由 manifest 重新生成，32 动词。
- **安装副本**：1.6.9 已装到 `<安装副本>`；73 个交付件，安装器自检通过，73 件哈希逐项一致（missing 0 / mismatch 0）；备份批次 `upgrade-20261124-142510`。

## 1.6.8 — 2026-11-24（覆盖缺口收口、N5 独立警告通道与真实项目复跑）

**主题：把 1.6.7 留下的 8 个覆盖缺口补进正式分支（16 例），同时收掉 N5 style 警告通道、`agent_io` 双还原、开发脚本裸 `Quit` 三项低风险遗留；四个 agent 动词的 CLI `--json` 接受度补端到端证据；真实项目「示例项目」首次跑隔离宏（commit `630a911`）。**

- **8 个覆盖缺口 16 例**：`watch` 源码变更写回（两轮快照确认后落 VBE）、`align --apply`、`clean-backup --apply`、`agent backup --apply`、`agent patch --apply`、`agent metrics --out --apply`、`style fix --apply` / `--which`、`agent book --action close` 目标簿；分落 `test_run.py` +5、`test_sync_transform.py` +2、`test_agent_io.py` +3、`test_agent_patch.py` +6。
- **选项分支补测**：`agent logs --tail`（含 0）、`sync run --tail/--log/--log-optional` 与默认日志目录、`assert --clear/--baseline` 实际删行。
- **N5 独立警告通道**：`style` 人话报告只保留 stdout 报告行；stderr 的 `[警告]` 行剥前缀进 `Outcome.warnings`，由 render 统一输出到 stderr；JSON 仍只一行、警告进 `stats.warnings`。原先“警告并入 human”的过渡口径结束。
- **`agent_io` 单边还原**：写回保存前只走 `guard.restore()` 一次，去掉模块级二次 `app_guard.restore`，避免重复诊断；顺序契约改为 `guard.restore → wb.Save`。
- **开发脚本裸 `Quit` 收口**：`scripts/make_lab_workbook.py` 改用 `toolkit.excel.new_excel` + `quit_owned`，全仓除 `quit_owned` 内部外不再有裸 `Quit`。
- **四个 agent 动词 `--json` 台账**：`test_agent_json.py` 新增 metrics 正常、patch/assert 失败信封的进程内 CLI 证据，`style` 已有端到端证据；补 ledger 依据。
- **真实项目复跑**：「示例项目」`pull --json` 成功（1 模块，disk 模式）；`run --macro 功能自测.零参自测 --isolate --apply` 成功，`isolate_cleanup=ok`，未保存用户工作簿，用户 Excel 进程未受影响。
- **测试基线**：**566 例 / 564 通过 / skipped 2 / 0 failure / 0 error**，68.3 秒；`compileall` 通过。
- **安装副本**：**1.6.8 已装到 `<安装副本>`**；73 个交付件，`doctor --verify-package` 通过（checked 73 / missing 0 / mismatch 0 / version_match true）；升级备份 `upgrade-20261124-121755`。
- **文档口径统一（不改命令行为，`5980a3c`）**：`plan` 模式集合改为 `open/disk/isolate/open_required/unavailable`；`plan / pull / pull-open / vbamod` 的结果行统一说明在 `stats.results`，`init` 的 `details` 字段按实际登记项回写；`stage` 枚举删掉不存在的 `backup/import`；备份保护、`--json` 输出、`agent probe` 示例、`book` 示例和专用开关分组按代码校正；`commands.md` 从 manifest 重新生成，31 动词不变。

## 1.6.7 — 2026-11-24（归属收尾稳定键与真实 Excel 端到端复跑）

**主题：把 `quit_owned` 的幂等键从 Python 对象地址换成 Excel 实例稳定键（PID/Hwnd），避免长驻 / 多实例场景下地址复用导致漏关；首次在隔离 lab 工作簿上复跑真实 Excel 的 pull / push / run / watch（commit `c0c4499`）。**

- **归属收尾稳定键**：`quit_owned` 的已完成集合改用 `(pid, hwnd)` 作键，读不到 PID/Hwnd 才退回 `id(xl)`；同一实例换新 COM 代理不会重复 Quit，不同实例各收各的。归属快照比对仍按 Python 代理对象取，发现 PID/Hwnd 变化仍拒绝退出。
- **测试**：`test_owned_lifecycle_1_6_5` 新增 2 例——同 Hwnd 新代理不重复 Quit、不同 Hwnd 不互相挡回；该文件 31 例全绿。
- **真实 Excel 冒烟**：隔离 lab `示例表 V1.xlsm` 上 `pull`（3 模块）、`push`（3 模块 / saved=true / backup 落盘）、`run --isolate`（返回 `1.6.6-e2e`）、`watch --seconds 20`（写回 1 次，宏返回 `1.6.6-e2e-watch`）、`book close`（目标簿关闭，用户 Excel 进程保留）全部通过；`isolate_cleanup=ok`。
- 测试基线：**549 例 / 547 通过 / skipped 2 / 0 failure / 0 error**；`compileall` 通过。
- 安装副本：**1.6.7 已装到 `<安装副本>`**；73 个交付件，`doctor --verify-package` 通过（version_match true）。

## 1.6.3-local 开发记录（历史，已并入 1.6.4–1.6.8）

**说明：本段原标为“1.7.0（开发中，未发布）”，实际是 1.6.3-local 评审批次的历史日志；下列条目均已并入 1.6.4–1.6.8，不是当前 1.7.0 的计划。当前 1.7.0 见顶部；历史命令面 31 条，当前为 32 条（新增 `sync.scrub` 等发布入口，不新增业务动词）。**

- **`sync -i` 交互菜单回补**（B6，commit `cf30b95`）：新增 `features/sync_interactive.py`，列项目 → 选项目 → 列动作 → 确认 → 复用 registry 既有 handler 执行；菜单层不碰 COM、不写盘，默认预览口径，写盘动词不带 `--apply`；非 TTY 仍由入口壳 `ensure_interactive` 人话退出 1。
- **`EXCEL_CLOSED` 优先于超时码**（B4，commit `0e5d047`）：看门狗等待期间按 PID 三态存活二次分类，目标实例已退出报 `EXCEL_CLOSED`（`stage=assert`），优先于 `WATCHDOG_TIMEOUT`；探测未知不把权限问题误报成崩溃。
- **`VBE_COMPILE_ERROR` 可达**（B2，commit `0e5d047`）：弹窗识别族回补后，超时 + 存活 + 编译错误框（`#32770` + 编译错误 / `#If`/`#Else`/`#End If` 关键字，按目标 PID 过滤）报 `VBE_COMPILE_ERROR`（`stage=vbe`），不再混成超时码。
- **其余模态框 `WATCHDOG_TIMEOUT` 且 hint 点名窗口**（B2，commit `0e5d047`）：通用 VBE 模态框、Excel MsgBox / InputBox 等非编译错误模态框归 `WATCHDOG_TIMEOUT`（`stage=assert`），`hint` 点名弹窗标题与正文关键片段；旧 `VBE_DIALOG` / `MODAL_DIALOG` 细分码不在当前枚举中。
- **复核口径确认（已接受设计差异，非行为修复）**：缺宏 / 禁用宏仍统一收口 `RUN_FAILED`，区分信息只落 `hint`（B3，commit `cf30b95`）；`agent assert` 业务失败退出 1，不恢复旧 exit 2 壳层合成（B5，commit `cf30b95`）。
- **内部重构 · 四处 features 拆分**：`audit_engine` 1349 → 142 门面，拆出 `audit_rules` / `audit_scan` / `audit_structure` / `audit_indicators`（`fe170ef`）；`sync_push` 1232 → 771，拆出 `sync_preflight` / `sync_shells` / `sync_init`（`d04fd48`）；`run` 992 → 312，拆出 `run_macro` / `run_watch`；`agent_io` 888 → 692，拆出 `agent_probe` / `agent_logs`（`8c7d79f`）。
- **内部重构 · 公共原语回补**：新增 `toolkit/watchdog.py`（线程看门狗原语、PID 存活探测，B1/B4）与 `toolkit/dialogs.py`（弹窗识别族，B2）；移除死函数 `macro_error_hint`。
- **内部重构 · 基线台账**：新增 `tests/baseline/COVERAGE.md` 与 `tests/baseline/1.5.6/coverage.json`（31 动词 / 124 条案例证据，三态定义），`test_release_diff` 基线不可用分支 `skipUnless` → `skipTest`（`c219a35`、`0b02fec`）。
- 测试基线：**494 例 / 490 通过 / skipped 4 / 0 failure / 0 error**；`compileall` 通过。skipped 4 = 1 条 B5 已接受设计差异硬跳过 + 3 条环境条件用例（缺 `mapping.json` 的 overview 单行 JSON 1、基线标签不可用 2）。

## 1.6.6 — 2026-11-23（suggested 定位纠正与簿白名单 FullName 优先）

**主题：收掉 1.6.5 报告 §5 的两处必修——`EXCEL_BUSY` / 重试耗尽两处 suggested 与隔离预检真相相反，`keep_workbooks` 只按 `Name` 匹配会误保留同名多簿；命令面与 1.6.5 等价，P1-4 与 P2-6 未触碰。**

- **两处 suggested 定位纠正**（commit `ba67e37`）：`EXCEL_BUSY` 与重试耗尽的 `COM_CONTEXT` 由「仍不行再用 --isolate 或批量模式」改为「先处理忙/弹窗/受保护视图；该表仍不可用时，换一张能打开的表再走 isolate」——错误码、stage、放行白名单、退出码、retryable 一律未动，只改话术。
- **簿白名单 `FullName` 优先**（commit `ba67e37`）：`keep_workbooks` 的名单项先按 `FullName` 归一比较（大小写不敏感 + 路径 / 斜杠归一），未命中再回退按 `Name` 大小写不敏感比较，命中任一即保留；残留簿问题清单改为「Name（FullName）」，便于同名多簿定位。
- **安装清单最小刷新**：按本轮用户口径不跑安装器打包，只在安装副本本地重算 73 条交付件 sha256 并把 `.installed.json` 的 `version` 由 1.6.4 刷到 1.6.6，`doctor --verify-package` 恢复 `ok:true`（checked 73 / mismatch 空 / version_match true）。
- **测试**（commit `ba67e37`）：新增 5 例——同名多簿按 `FullName` 只保留名单内那本、`Name` 旧写法仍兼容、路径大小写 / 斜杠归一命中、两处 suggested 不含误导措辞。
- 测试基线：**547 例 / 543 通过 / skipped 4 / 0 failure / 0 error**；`compileall` 通过（exit 0）。
- **明写不做（遗留）**：N5 警告通道口径未裁决、归属快照以 `id(xl)` 为键只适用单线程短生命周期、真实 Excel 会话端到端未复跑，详见 `tests/reports/release-acceptance-1.6.6.md` §3。

## 1.6.5 — 2026-11-23（自有实例收尾纵深收口与 style 警告口径过渡）

**主题：落地 1.6.4 评审登记的 N1–N7 非阻断项：归属快照复核、簿白名单、其余回收路径统一 `quit_owned`、ROT 显式收口、还原未生效入错误、style 双缓冲与警告前缀归一、isolate 阻断提示对齐；命令面与 1.6.4 等价，不新增业务动词。P1-4（exit 2 仅人话）与 P2-6（audit 全仓 `ProjectMode.NONE`）冻结契约本轮未触碰。**

- **自有实例收尾统一走 `quit_owned`**（N3，commit `6606880`）：`common_subset.close_owned`（被 `agent.push` / `sync.push` 复用）、`agent_io` 的 `book open` 三处失败回收与 `_quit_if_empty`、`sync_push` 单项目与批量两处 `_close_owned`、`run_macro` 隔离开表失败分支，由裸 `Quit()` + `except: pass` 改为 `toolkit.excel.quit_owned`；回收问题清单进 warnings / stats（`isolate_cleanup=ok` / `isolate_cleanup_problems`），不再静默吞异常。结论：全仓仅剩 2 处 `Quit`——`quit_owned` 内部的收尾调用、`scripts/make_lab_workbook.py` 开发辅助脚本（已判定保留：不引入 toolkit 依赖且无诊断通道）。
- **归属快照复核**（N1，commit `6606880`）：`new_excel` 起实例登记 pid/hwnd 快照；`quit_owned` 收尾前比对，字段不一致则不碰簿、不 Quit，问题清单记「归属校验不通过」；无快照的旧调用路径（attach 借来的实例、桩对象）保持旧行为，不因缺快照报错。
- **簿白名单**（N2，commit `6606880`）：`quit_owned` 新增 `keep_workbooks`，名单内簿保留不关；关完后名单外仍有簿则不 Quit、上报剩余簿名——宁可留实例，不吞他人数据。隔离路径已传入白名单（目标表之外只保留宏侧自开参考簿）；`keep_workbooks=None` 为 1.6.4 旧行为（关全部簿后退出）。
- **ROT 不可达记录显式收口**（N7，commit `6606880`）：新增 `reset_rot_unreachable()`，供测试隔离 / 重入主动清场；枚举开头清空与 1.6.4 的 `COM_CONTEXT` 分流语义不变。
- **还原未生效记入错误**（N4，commit `6606880`）：`AppGuard.restore` 回读仍为开启态（赋值被 Excel 静默忽略）时，向 `errors` 追加「还原 <attr> 未生效：回读仍为 …」，按整条文案去重；`describe()` 末尾附上未生效条目；`_is_off` 归一口径与两属性独立迭代不动。`agent_io` 与 `sync_push` 收尾处已把 `AppGuard.errors`（截 3 条）接进 warnings，有改动附 `describe()` 行，既有人读输出未动；跑宏路径既有 warnings 通道带出 `guard.errors`。
- **style 报告双缓冲与警告前缀归一**（N5，commit `6606880`；过渡口径）：1.6.4 的单缓冲改为 `out_buf` / `err_buf` 双缓冲；人话模式 stderr 警告行并入报告时保留 / 补齐 / 去重为单个 `[警告]` 前缀，`[错误]` 行保持错误标记、不误标成警告；JSON 模式只收报告行进 `stats.lines`，警告走 `stats.warnings`，错误 / traceback 杂项不进信封。**产品口径未裁决：本条只做到前缀可辨识，未重新分流警告通道。**
- **`run --isolate` 阻断提示对齐**（N6，commit `6606880`）：隔离预检放行白名单仍仅 `NO_EXCEL` / `TARGET_NOT_OPEN`；`COM_CONTEXT` 阻断隔离副本时，hint 说明三类成因（弹窗 / 受保护视图 / 异权会话）并明确「隔离模式也不会启动」，suggested 为可复制的 `run --project … --macro … --isolate --apply`；放行白名单与错误码归类未动。attach 侧三处网关默认 suggested（rot / ready / books）也一并改为「先关卡住实例重试附着；该表仍不可用时换一张能打开的表再走 isolate」，`tests/test_sync_push.py` 的契约字面量随口径同步（断言强度未放宽）。
- **测试**（commit `6606880`）：新增 `tests/test_owned_lifecycle_1_6_5.py`、`test_app_guard_1_6_5.py`、`test_agent_patch_1_6_5.py`、`test_run_1_6_5.py`，覆盖 1.6.4 评审 §7.2 缺测：attach 会话 close 不 Quit、`quit_owned` 异常 / 幂等 / 归属不符 / 白名单残留、ROT 落表与下轮清空 / 显式 reset、style 非 JSON 含警告与 JSON 失败 stderr 为空、app_guard `"FALSE"` 归类与还原未生效去重。
- 测试基线：**542 例 / 538 通过 / skipped 4 / 0 failure / 0 error**；`compileall` 通过（exit 0）；本轮 +42 例全部来自新增 1.6.5 用例。
- **文档同步**：`references/toolkit.md` 原语清单补 `reset_rot_unreachable`，`quit_owned` 签名与场景表按归属复核 / 白名单两行口径更新；`references/sync-commands.md`、`references/pitfalls.md` 补 `run --isolate` 的 `COM_CONTEXT` 阻断口径；`VERSION` 与 `SKILL.md` metadata.version 均为 1.6.5。

## 1.6.4 — 2026-11-23（隔离实例收口与 attach 失败分流定向修复）

**主题：避免误关用户正在使用的 Excel：工具自起实例先关簿再退出并上报收尾诊断，ROT 可见但绑定失败按 `COM_CONTEXT` 分流，未开表提示中性化；收口 agent_io / run_watch 附着文案、style 报告 stderr 捕获、状态栏假值误报；命令面与 1.6.3 等价，不新增业务动词。**

- **隔离实例只回收自身拥有的实例**（commit `b52f6d1`）：`toolkit/excel.py` 新增 `quit_owned(xl, save_changes=False)`：先置 `DisplayAlerts=False`、逐簿 `Close(SaveChanges=False)`，再 `Quit`，避免 dirty 工作簿触发隐藏保存提示、进程滞留与 `~$` 锁文件；返回未清干净的诊断列表，不再像旧实现那样把退出异常静默吞掉。`ExcelSession.close()` 与 `run --isolate` 收尾改走该原语：收尾正常记 `stats.isolate_cleanup=ok`，有问题写 `stats.isolate_cleanup_problems` 并进 warnings。
- **ROT 可见但绑定失败按 `COM_CONTEXT` 分流**（commit `b52f6d1`）：`attach_running_excel` 枚举 ROT 时记录「看得到 moniker、`BindToObject` 失败」的工作簿；目标落在该清单时，后续 `GetActiveObject` 失败或附着后找不到表均报 `COM_CONTEXT/stage=attach`（hint 指向弹窗 / 受保护视图 / 权限不同的会话，suggested 指向关卡住实例或 `--isolate`），不再降级成「没打开」的 `TARGET_NOT_OPEN`。
- **未开表兜底文案中性化**（commit `b52f6d1`）：`TARGET_NOT_OPEN` 的 suggested 不再带「push 默认自动开隐藏实例」口径，改为「先打开目标工作簿，或用 `book --action open --project <项目> --apply`」；`run` 附着失败的 hint / suggested 同步改为指向 `run --isolate`（带项目与宏名），不再指向 `list`。
- **agent_io / run_watch 附着文案对齐入口**（commit `b52f6d1`）：`agent.py book --action close` 附着失败提示改为「表没打开：无需关闭；要打开目标表用 `book --action open`」；`agent_io._attach_or_open` 与 `run_watch` 的附着失败按各自入口给 hint 与可执行 open 命令，不再共用中性兜底串。
- **style 报告同时捕获 stderr**（commit `b52f6d1`）：`agent_patch` 的 `style check/fix` 原先只 `redirect_stdout`，JSON 模式下 `msg` 固定写 stderr，报告会漏到 stderr、`stats.lines` 为空；现 stdout / stderr 一并截入同一缓冲，报告统一进 human 与 `stats.lines`，JSON 出口保持干净。
- **状态栏 / CutCopyMode 假值不再误报还原**（commit `b52f6d1`）：`app_guard` 对 `StatusBar` / `CutCopyMode` 的「已关闭」做归一（False / 0 / None / 空串 / `"false"` / `"0"` / `"none"`，含 COM 读回的字符串 `"FALSE"`）；归一后视为关闭、不记改动；设置后回读确认确实落成关闭才记 `changed`，避免把一次无效设置记成「已还原」。
- **测试与回归**（commit `b52f6d1`）：新增 `tests/test_regression_1_6_4.py` 6 例（`quit_owned` 关簿顺序 / 退出失败上报、ROT 绑定失败 `COM_CONTEXT`、未开表兜底无 push 口径、字符串 FALSE 不算改动、真实状态栏仍还原）；`test_run.py` / `test_agent_io.py` / `test_agent_patch.py` 断言同步（attach 桩接受关键字参数、`_acquire` 传宏名、style 报告非空且 stderr 为空）。
- 测试基线：**500 例 / 496 通过 / skipped 4 / 0 failure / 0 error**；`compileall` 通过。
- 实测复核：真实项目（示例项目，工作簿 `示例表 V1.xlsm`）复跑 P0-1：`isolate_cleanup=ok`，无新增 Excel 实例，无 `~$` 锁文件。
- **文档同步**：`references/toolkit.md` 原语清单补 `quit_owned`，场景表补「自起实例回收」与「ROT 绑定失败」两行；`SKILL.md` metadata 版本号随 1.6.4 升级。

## 1.6.3 — 2026-11-23（命令与数据流统一化定版）

**主题：公共库子集判定与写回范围收单源、三入口失败壳统一、audit 规则收敛、JSON 过程提示收口；命令面仍为 31 条，不新增业务动词。**

- **公共子集与写回范围单源**：`common_reference_reasons` / `common_version_stems` / `resolve_common_subset` / `find_project` / VBE 快照复原与关闭实例收进 `features/common_subset.py`；`core`、`sync_read`、`sync_push`、`sync_transform`、`run`、`agent_io` 只做参数适配，不再各自复制判定规则。
- **agent push 与 sync push 对齐**：`agent.push` 的预检、写回、计数三处按 `common_modules` 收窄，补齐 `--doc-modules`、`--no-backup`、`--no-save`、`--open-only`、`--rollback-on-fail` 五个开关；dry 预览显示将写回的公共模块范围与生效开关，未知模块告警在 `list/pull/push/run` 同文案、同进程只出现一次。
- **三入口失败壳统一**：`toolkit/cli.py` 新增 `run_dispatch()`，`sync.py` / `agent.py` / `audit.py` 的 `main` 不再各自维护 `try/except`；`add_version_arg()` 统一三入口 `--version`；`render` 增加“失败必有人话”兜底，`audit` 业务失败在 human 模式不再静默。
- **JSON 过程提示收口**：`toolkit/base.py` 新增 `note()`，JSON 模式下过程提示进入 `stats.warnings`，不再直写 stderr；`audit --out` 与 stdout 复用同一信封，`stats.warnings` 与 `exit_code` 保持一致。
- **audit 规则收敛**：18 对同名函数逐对裁决，13 对等价、1 对补错误元数据、4 对口径分叉按 `audit_engine` 收敛；`audit_a.py` 退化为参数适配与输出装配，只保留纯委托别名，并补跨视角一致性用例。
- **有意保留差异**：`attach_or_open`（错误码/排障提示按入口不同）、`source_preflight`（notes 软提醒、检查深度、子集过滤按入口不同）保留双份并在代码注释标明；失败回滚语义（agent 移除新增模块、sync/run 保留）按现有产品裁决保留。
- 测试基线：**475 例 / 467 通过 / skipped 8 / 0 failure / 0 error**；`compileall` 通过。
- 兼容性说明：`agent.push` 的写回范围从整目录收窄为 `common_modules` 子集，属行为变更；需要整目录写回时通过 `--all-project` 或配置口径显式放宽。

## 1.6.2 — 2026-11-22（修复合入至 `52368a5`）

**主题：收口建议命令绝对入口、open-only 失败分流、pull 已打开自动借会话、assert 能力边界、JSON 警告纯净、升级备份迁出 skills 根、doctor 显式包完整性校验；命令面与 1.6.1 等价，不新增业务动词。**

- **① 建议命令改绝对入口**：原先 `suggested_command` 用包内相对路径入口，换工作目录后不可直接复制执行。现新增 toolkit 助手 `entry_command(entry, verb, *args)`（`sys.executable` + 由包位置解析的入口绝对路径 + 参数，空格引用），覆盖 assert `ASSERT_NO_LOG`、push `--open-only` 两类提示、init 归属门禁拒绝提示三处。提交 `3a1630d`；`test_agent_patch` / `test_init` / `test_sync_push` 相关断言改为校验入口真实存在、非包内相对路径。
- **② push `--open-only` 未开表分流**：`push --open-only` 且表未打开原先冒泡 `NO_EXCEL/stage=attach`，与磁盘回退口径混淆。现仅 `auto_open=False` 且 code 为 `NO_EXCEL` 时改写 `TARGET_NOT_OPEN`（stage=attach、hint 指向去掉 `--open-only` 或先开表、suggested 为绝对重试命令）；`COM_CONTEXT` / `VBE_PROTECTED` 原样上抛；自动开表路径的 `NO_EXCEL` 仍回退隐藏实例磁盘。提交 `3a1630d`；`test_sync_push` 51 → 54。
- **③ pull 已打开自动借会话**：`pull --apply` 原先未区分表开没开，可能抓磁盘副本而漏掉未保存改动。现逐项目先附着已打开会话（含未保存改动、不另开副本、不关窗、按 `--no-backup` 留底），未打开走原隐藏实例磁盘；`--all-project` 逐项目判定、结果行带 `mode: open|disk`；仅 `NO_EXCEL` / `TARGET_NOT_OPEN` 允许退磁盘，`COM_CONTEXT` / `EXCEL_BUSY` / `EXCEL_CLOSED` / `VBE_PROTECTED` 原样上抛并给「本应借会话但失败、未退磁盘」hint；`plan --intent pull` 推荐命令统一为 `pull`；dry 只做 ROT 轻量探测。提交 `6d27801`；`test_sync_read` 22 → 26（净 +4，新增 6 移除 2）；evals id3 与 `references/sync.md` §3.1 口径同步（`0d2856e`）。
- **④ assert 能力边界写进文档**：assert 原先被当作通用宏执行器，只写工作表、不写日志的自测宏也被丢进去跑。现明确 assert 只接会写日志的断言宏（读「通过/失败」计数），零参 / 两参自测宏只写工作表时恒报 `ASSERT_NO_LOG`，应改用 `sync.py run`；边界写入 `SKILL.md`（常用命令表 + 边界节）与 `references/agent.md`。提交 `0a4eefb`。
- **⑤ assert 过期文档更正**：`references/agent.md` 原先把 `ASSERT_MACRO_ARGS` 标为「1.6 不可达、随 1.7 回补」，与实际行为不符。现删除该说法，按实际口径写：先零参调用、仅参数个数错误退 `(False, True)`、两次都错才报 `ASSERT_MACRO_ARGS`、第二次其它错误原样透传；顺带更正 assert 退出码口径（0 通过 / 1 业务失败 / 2 用法错）。提交 `0a4eefb`。
- **⑥ `--json` 警告收口**：JSON 模式下 `warn()` 原先仍写 stderr，可能打断单行 JSON 信封。现 JSON 模式下 `warn()` 不写 stderr 而缓存，`render.py` 不再向 stderr 打 warnings，两路合并进信封 `stats.warnings`（仅有警告时落键，不删既有键，不与 1.5.6 契约冲突）；三入口解析参数后即置 JSON 开关。提交 `eceb2d0`；`test_json_contract` 33 → 36（新增 3 例：JSON 收口 / 非 JSON 仍 stderr / legacy config 警告收口）。
- **⑦ 升级备份迁出 skills 根**：升级备份原先与安装位置耦合，可能落在 skills 发现根内被再次扫描。现备份根与安装位置解耦，固定用户级 `%LOCALAPPDATA%\vba-workbench-install-backups\vba-workbench\upgrade-<时间戳>\`（`LOCALAPPDATA` 缺失退 `%USERPROFILE%\.vba-workbench-install-backups\`），永不落在 skills 发现根内；旧备份（`vba-workbench.bak-*`、树内 `.backup\upgrade-*`、1.6.1 同级 `.install-backups`）只读探测并给人工清理建议，绝不自动删除、不新建；空目录修剪跳过 `bak-*`；rollback 顺序为新根 → 旧同级 → 旧树内。提交 `6082cd0`；`test_release_diff` 24 → 28（+4）。
- **⑧ doctor 显式包完整性校验**：doctor 原先没有按安装清单对账交付件的开关，本地正常改动也可能被误判。现新增显式开关 `sync.py doctor --verify-package`（manifest 登记；默认 doctor 行为不变，不因本地正常改动失败）；校验清单在位 / 字段齐全 / files 与磁盘对账 / 逐项 sha256 / VERSION 与清单 version 一致，结果进 `stats.package`（`present/checked/missing/mismatch/extra/missing_fields/version_match/manifest_version/disk_version`）；新增错误码 `PACKAGE_MANIFEST_MISSING`（无清单）、`PACKAGE_INTEGRITY`（缺字段 / 缺件 / 哈希不符 / 多余 / 版本不符，message 列具体相对路径，hint 说明本机正常改动可忽略、可重装或 `--apply` 刷新）；显式校验不通过退 1、通过退 0。提交 `163d3fe`；`test_core_sync_entry` 21 → 29（+8）。
- **⑨ 建议命令全部改绝对入口 + 白名单升级**：安装后 cwd 不在仓库根时，`suggested_command` / hint / 人话行 / 帮助 / 报错里的 `python vba-workbench/...` 相对串照抄找不到文件；`make_envelope` 白名单原只认 `python ` 前缀，会把已拼好的绝对命令降级进 hint。现全部改 `entry_command()` 绝对入口（解释器绝对路径 + 入口脚本绝对路径 + 参数），`toolkit/base.py` 新增 `looks_like_command()` 并在 `make_envelope` 改用它识别可执行命令，绝对命令原样保留在 `suggested_command` 独立字段、不降级进 hint；纯人话建议仍按旧口径降级。提交 `b86c720`（白名单助手）+ `b98bd05`（sync 系 15 处）+ `4e0877a`（agent 系 21 处）+ `98b9ec0`（audit/toolkit 23 处）；合并 `8823a85` / `4a3c73f` / `52368a5`；`scripts/**/*.py` 相对包路径命令串（`/` 与 `\` 两种写法）残留 0。
- 测试基线：450 例 / 32 模块 / skipped 9（0 failure / 0 error，通过 441）。

## 1.6.1 — 2026-11-22（修复合入至 `e662e80`；1.6 分层架构定向修复 6 项）

**主题：在 1.6 分层架构上收口失败边界、assert 回退、提示出口、公共库递归、安装清单与 init 归属门禁；命令面与 1.6.0 等价，不新增业务功能。**

- **P0 异常基类穿透**：features 跨 toolkit/adapters 边界原先只捕具体异常，基类 `BaseFailure` 会穿透到兜底出口。现跨边界统一捕 `BaseFailure`，进 `Outcome.errors` 的用 `to_contract_failure` 归一；确属纯契约层的保留 `ContractFailure`（`sync_push` 三处：`_find_project` / `_preflight_or_fail` / 批量 dry 预检）。提交 `99e69c0` + `23cb7ac`（回归证据）+ `7550182`（`sync_read` 三处补漏）；`97c0f02` 同时计入下条 P16-3。
- **P17 `agent assert` 回退**：原先先按两参签名调宏，参数个数不符即退；现先零参 `Run(macro)`，仅参数个数错误才退 `(False, True)`（含性能=False、静默=True），两次都个数不符报 `ASSERT_MACRO_ARGS`（stage=assert、retryable=False），第二次非 mismatch 原样透传。提交 `2b12353`；`test_agent_patch` 26 → 30。
- **P16-3 push 提示出口**：push 附着失败提示原先指向 `run --isolate`，与自身出口不符；现改自身口径，`attach_running_excel` 三处 `COM_CONTEXT` suggested 支持 `target_suggested` 覆盖（不传时默认文案逐字不变），push 传自身出口。提交 `ae67f94` + `97c0f02`；`test_sync_push` 45 → 51。
- **P18 doctor 公共库递归**：doctor 原先只从顶层 `iterdir` 判公共库，嵌套模块漏判；现复用 `toolkit.repo.files_under` 递归、改用 `SOURCE_EXTS` 并跳过占位壳 / 噪声，空目录仍 warn。提交 `1e76bd2`；`test_core_sync_entry` 19 → 21。
- **P19 安装清单**：apply 原先逐文件落盘、无清单可核对；现 apply 后原子写 `.installed.json`（version / source / installed_at / files 哈希），升级备份迁出安装树到 `<安装根>\..\.install-backups\vba-workbench\upgrade-<时间戳>`，旧位置 `.backup` 只读提示不删，新增 `verify_manifest` 与显式 rollback。提交 `a2e8940`；`test_release_diff` 20 → 24；实测 preview 58 → apply 落盘 59（58 交付 + 清单）、树内 0 个 `.backup`。
- **P20 init 配置归属门禁**：init 原先在非 workspace 目录也能写配置；现无显式 `--config` 且目标目录不在「workspace 目录 / 生效配置目录」内即拒写（USAGE / Stage.CONFIG / exit 2，提示含可直接复制的 `--config` 命令）；`--lab` 未显式给配置时写 lab 目录局部 `mapping.json`；env `VBA_WORKBENCH_CONFIG` 不算显式豁免。提交 `4c47f24` + `b6fbf8f`；`test_init` 5 → 9；端到端：env 指生产配置时生产 mapping sha256 前后逐字节一致。
- 测试基线：418 例 / 32 模块 / skipped 9（0 failure / 0 error，通过 409）。

## 1.6.0 — 2026-11-22（d6f9dc2；架构升级）

**主题：旧单体 `sync.py` 收口为三入口壳 + features 切片 + toolkit 原语 + adapters 网关的分层架构；命令面按主要功能等价对齐 1.5.6，不新增业务功能。**

- 三入口壳：`scripts/sync.py` 119 行、`scripts/agent.py` 97 行、`scripts/audit.py` 114 行；壳层只做参数解析、上下文装配、派发与信封渲染，不含动词特例。
- 分层：十个 features 处理器切片 + toolkit 公共原语 + adapters 网关；命令面由 manifest 单源登记，`references/commands.md` 为唯一命令表（由 `tests/tools/gen_commands_md.py` 从 manifest 生成，勿手改）。
- 命令面 31 动词不变（sync 16 / audit 6 / agent 9）。
- 测试基线：383 例 / 32 模块 / skipped 9（0 failure / 0 error）。

## 1.5.6 — 2026-11-22（55f2821；1.5.5 定向修复 11 条）

**主题：把“能跑但报错说不清”的失败路径逐个落到固定错误码与可执行出口；本轮只做定向验证，不跑全量。**

- P1 `agent assert` 不再假定宏必收两参：先 `Run(ref)` 调零参宏，遇参数个数错误再退到 `(含性能, 静默)`；两支都失败报 `ASSERT_MACRO_ARGS`，不落 `UNEXPECTED`。宏返回后日志只再等 5 秒，不再空耗剩余 `--timeout`。
- P2 `warn()` 一律走 stderr；`--json` 的 stdout 只留一行 JSON，不再被 `[警告] Traceback` 打断。
- P3 `list --json` 对未解析项目输出 `TARGET_NOT_FOUND/stage=resolve`，`hint` 带 `search_dir/prefix/期望文件`；`ok` 改为“全部项目都解析成功”，不再出现 `FAILED/命令失败/unknown`。
- P4 `classify_run_error` 补 `-2147352561` / `0x8002000f` / `parameter not optional`，参数个数不符统一 `ARG_MISMATCH`，message 不再拼裸 COM 元组。
- P5 `patch` 的格式错误、源码缺失、整行不唯一分别落 `PATCH_FORMAT` / `PATCH_FILE_NOT_FOUND` / `PATCH_BLOCK_AMBIGUOUS`；命中次数与行号写进 message，hint 说明“整行唯一、缩进一致”。
- P6 `agent logs` 无日志落 `LOGS_NOT_FOUND/stage=logs`，hint 指向 `run --apply` 或 `--log`。
- P7 `pull` 检测到目标表已打开时 `warn`：“抓的是磁盘副本，有未保存改动请改用 pull-open”；与 `plan --intent pull` 的口径一致。
- P8 `assert` 的 `NO_EXCEL/TARGET_NOT_OPEN` 改用断言专属出口：`agent.py book --action open`，不再指向它没有的 `run --isolate`。
- P9 `FILE_READONLY` 补 `suggested_command=attrib -R "<文件>"`，信封允许 `attrib ` 前缀保留。
- P10 文档的 `list` 字段改 `workbook`（`SKILL.md`、`QUICKSTART.md`），与 JSON 实际输出一致，不再双轨写 `target`。
- P11 `version --json` 增 `stats.configured`、`stats.common_dir`，未锁项目带 `reason`；无公共库时 `configured=false`，不再用 `mismatched=0` 冒充“校验通过”。

## 1.5.5 — 2026-11-22（5e8ee9c；规范整改 4 项：参数互斥、长文档目录、分册编号、失败类 evals）

**主题：把写回前的参数冲突挡在解析层；四份长 reference 可独立读；失败路径有固定用例。**

- 参数互斥：`--project` 与 `--all-project` 进入同一个 argparse 互斥组，`pull` / `push` / `gbk-export` 同用两参数时直接退出 2，不再让 `--all-project` 静默压掉 `--project`；`tests/test_cli_contract.py` 补 3 个动词的契约用例。
- 长文档目录：`references/agent.md`、`sync-commands.md`、`sync-config.md` 的加粗目录统一为 `## 目录`，`pitfalls.md` 补 16 条目录；四份均超过 100 行。
- 分册编号：`sync.md`（§1–§3）、`sync-commands.md`（§1–§4）、`sync-config.md`（§1–§2）、`sync-maintain.md`（§1–§5）各自从 §1 重编，消除跨册承接与 10/11 重复；`sync.md` 分册索引、`sync-config.md` 对 `sync-maintain.md §5`、`sync-maintain.md` 对命令册 §2.1/§2.4 的引用同步更新。
- 失败类 evals：`evals/evals.json` 12 → 18 条，新增 `NO_EXCEL`、`EXCEL_CLOSED`、`COM_CONTEXT`、`TARGET_NOT_OPEN`、`VBE_PROTECTED`/`VBE_DENIED`、`MACRO_DISABLED` 六条行为用例；`evals/README.md` 同步计数。
- 验证口径：只跑 `tests.test_cli_contract` 9 条（通过）、`sync.py --version`、`evals.json` 解析与 `git diff --check`；按本轮分工未跑全量测试。

## 1.5.4 — 2026-11-22（1703e03；1.5.3 全量测试 13 条：配置根、目标缺失分流、提示可执行、备份脱表）

**主题：让配置只落在用户指定位置；目标表找不到时先报 `TARGET_NOT_FOUND`；提示按动词给可执行出口，不再留 `<表名>` 占位。**

- 配置查找：设了 `VBA_WORKBENCH_CONFIG` 就只认它，文件不存在也把该路径原样报进 `CONFIG_NOT_FOUND`，不再回退到 `skills\mapping.json`。`init` 的 `--config/--dir/--src/--lab` 全部以配置文件所在目录为基准；`--lab` 且未给 `--dir` 时测试表、源码都建在 `<配置目录>/.lab`，不再写进共享的 skills 根。
- 目标表缺失：`connect_workbook` 在连 Excel 前先判 `search_dir/prefix`，找不到直接 `TARGET_NOT_FOUND/stage=resolve`，不再误报 `NO_EXCEL`；`push` / `pull-open` / `run` 都先过这一关。`plan` 在无目标时把已有的人话原因写进 `errors[0]`，不再是 `FAILED/命令失败`。
- 提示按动词分流：`NO_EXCEL` 复用 `TARGET_NOT_OPEN` 的 `hint/suggested_command`；`push --open-only` 明确说“不会自动开表”，`watch` 给 `book open`，`pull-open` 给 `pull`，`run` 给 `--isolate`；所有建议命令都用真实项目名，不再出现 `<表名>`。
- `FILE_READONLY` 补 `hint`（取消只读属性、排查网盘占用）；`run` 识别 `DISP_E_BADPARAMCOUNT (0x8002000E / -2147352562)` 为 `ARG_MISMATCH`，提示传参个数与过程签名，不再把裸 COM 元组当 `RUN_FAILED` 正文。
- `agent.py backup` 表关也能备份：拿不到 Excel 会话时直接复制配置解析到的 xlsm，不再要求先开表。
- `clean-backup` 的 `by_tag / kept_by_tag` 把无 tag 批次统一记为 `default`，不再用空字符串当 JSON 键；空备份目录也返回完整计数，不再 `KeyError`。
- `doctor` 修正档位：`ok` 项不再因为调用方默认 `warn` 被判警告；`项目可用数 1/1` 为 `ok`，`src/common` 缺失或无模块按 `warn`，与 `audit common` 的只读警告一致。
- H13 区分“缺文件”和“文件在但无编号行”：前者 `未跑/count=-`，后者 `警告/count=0` 并带样本；`agent.py probe` 的 `details[]` 键名统一为英文（`project/file/disk/open/vbe/modules/test_sheet/rows`），`pull-open` 的 `saved` 固定为 `false`，另给 `workbook_saved/wrote_src`。
- 文档同步：命令码表补 `TARGET_NOT_FOUND / ARG_MISMATCH`，`details[]` 键名统一说明，`pull-open` 的 `saved` 语义改成“本次不保存工作簿”。

## 1.5.3 — 2026-11-22（4c2d0e9；1.5.1 复测余项：keep=0、common 警告、assert 可配置、H05a 规则可见）

**主题：非模板项目也能用 assert；公共库缺失不再硬失败；备份保留数可逐类核对；H05a 的放行写法不再只藏在 audit.md。**

- `agent.py assert` 增加 `--macro 模块.过程` / `--sheet 表名`，并按“命令行 > 项目键 `assert_macro/assert_sheet` > `defaults` > 模板默认”解析；目标表没有配置的宏时先报 `ASSERT_MACRO_NOT_CONFIGURED`，给 `--macro` 可执行命令，不再落到只说“核对宏名”的 `MACRO_NOT_AVAILABLE`。`agent.py` 的 JSON 信封补回 `hint`，不再只透传 `suggested_command`。
- `audit.py common` 无公共库时不再 `AUDIT_FAILED`：按只读警告 `exit 3 / ok=true`，`stats.configured=false`，打印 `common_dir` 与配置键 `defaults.common_dir`；`overview` 口径不变。
- `clean-backup --keep 0` 钳到 1，并在 `stats.keep_requested / keep_note` 明确说明；新增 `stats.kept_by_tag / kept_names`，`kept=4` 能看出是 3 个 tag 各留 1 批再加全局 1 批，不再只给一个总数。
- H05a 的 `[加速例外]` 放行规则补进 `SKILL.md` 和 `references/pitfalls.md`；`audit.md` 仍是详细口径。
- 定向回归 13 条 OK（新增 assert 配置/common 警告/keep 口径）；全量单测仍由专人统一跑。

## 1.5.2 — 2026-11-22（ca746bb；1.5.0 复测余项：TARGET 提示分流、备份口径可见、book 空实例、common hint、run 计时）

**主题：把 1.5.0 复测里仍在的“提示串动词 / 空 Excel / 计时失真”收口；1/5/6/7 已在 1.5.1 修，需用 1.5.2 复测。**

- `TARGET_NOT_OPEN` 的 `hint` 与 `suggested_command` 按动词分流：`run → book open 或 --isolate`、`pull-open → sync.py pull`、`watch → book open`、`push → 直接 push（提示仍保留 --open-only）`，不再把 push 的 `--open-only` 塞给 watch/run。
- `clean-backup` 的 JSON 增加 `keep_scope=per-tag / keep_daily / min_keep / auto_clean / by_tag`；`--keep N` 明确是“每类保留 N 批”，加上全局 `min_keep` 后总保留会多于 N，`details[]` 继续逐条列出将删批次，预览可核对。
- H10 的每类阈值改为 `backup_keep + backup_min_keep`（默认 13），并保留总量 60 批 / 450 MB 兜底；`backup_keep_daily>0` 时不再按每类卡数，避免“每类 12 批”被误报。
- `agent.py book --action close` 关工作簿后若该实例已无其它工作簿，自动 `Quit` 回收；`book open` 中途失败也回收新实例，避免累积可见空窗口。
- `audit common` 无 `src/common` 时改为 `COMMON_NOT_CONFIGURED/stage=common`，补 `hint` 与 `overview` 可执行建议；audit 的 Failure 信封不再吞掉 `suggested_command`；人话入口同时打印 `说明 / 建议`。
- `run` 计时拆开：`cost_sec` 改为整条命令耗时，`run_sec` 只记 `Application.Run`；不再用一个 0.02 秒的数字代表整次隔离运行。
- 文档：`--project/--all-project` 从“通用开关”移到 `pull/push/gbk-export` 专用；`clean-backup --keep` 标注 per-tag；补 book 回收、common 错误码、run 计时口径。
- 定向回归新增 4 条（TARGET 提示、book 空实例、common hint、H10 阈值），本文件共 11 条；全量单测仍由专人统一跑。

## 1.5.1 — 2026-11-22（d81e4b4；1.4.9 环境 6 条问题：dry 预检、overview 门禁、H05a 误报、宏提示、watch 命令、H10 修复动作）

**主题：把“预览与执行两套口径”“overview 永久红”“安全执行式加速误报”收口。**

- P1 `push` / `push --all-project` 的 dry 预览现在同步跑源码预检：有致命项直接 `SOURCE_PREFLIGHT/stage=preflight` 失败，不再 dry 说没问题、`--apply` 才炸；单项目失败不拖垮整批。
- P2 `overview` 只把 `sheet_mismatch` / `binding_target` / `dup` / `shadow` 判为阻断；`manual_accel` 等口径提醒只写入 `stats.issues`、`stats.fatal`，默认 `ok=true`、退出 0，`--strict` 才按 1。`structure` 的 `exit 3` 语义不变。
- P3 H05a 只统计“关闭赋值”，成对还原（包括错误分支里的重复还原）不再计入违规；未还原的裸关闭仍报。
- P4 `sync.py run` 与 `agent assert` 的 `MACRO_NOT_AVAILABLE` 共用 `toolkit.base.macro_error_hint`，`AutomationSecurity=Low` 时两条路径都带“当前是 automation 实例…”句。
- P5 `plan --intent watch` 在表关闭时给出可直接执行的 `agent.py book --project X --action open --apply`，不再返回空 `command`。
- P6 H10 警告行新增 `suggested_command`（`sync.py clean-backup --keep 5 --apply`），人话与 Markdown 都打印“修复”动作。
- 定向回归新增 9 条：dry 预检失败、overview 仅 `manual_accel` 退 0、H05a 成对/未成对、宏提示一致、watch 命令、H10 字段；全量单测由专人统一跑。

## 1.5.0 — 2026-11-22（1ead5b5；自查 14 条整改：写回失败现场、门禁语义、静态盲区）

**主题：把 1.4.8 复查出的 14 条按影响收口；重点是写回失败后的内存现场、快捷键绑定静态检查、警告/失败退出码分离。**

- P0 写回失败：`sync_into` 新增 `rollback_on_fail`，`push` / `push --all-project` / `run --push` 加 `--rollback-on-fail`；开启后 postcheck 失败按写回前 CodeModule 正文复原本次写过的模块，并报告复原/失败数。默认仍不回滚（沿用历史约束），但 `plan` 明说失败会把 VBE 内存工程改脏。
- P0 备份清理：`prune_backups` 新增 `protect`，`defaults.backup_protect` 与 `clean-backup --keep-tag=关键字` 可保护人工回滚点；清理时列出被保护跳过的批次。
- P0 目标解析：`resolve_target` 对同版本候选中“同尺寸 + 同 mtime”的副本报警；内容相同提示重复副本，内容不同提示必须显式配置 `file`，不再静默取最后一份。
- P1 `plan` 对“表已打开且有未保存改动”给三条可执行处置：另存副本后 push、先手动保存再 push、关表不保存再 push；并说明 `--no-save` 不保护单元格改动、失败会改脏 VBE。
- P1 新增 `structure ⑭ binding_target`：静态解析 `注册快捷键(键, 名)` 与 `注册动作`，校验目标过程存在/Public/必填参数、窗体是否在 `窗体实例` 登记、表页文档模块方法是否存在；支持 `--ignore binding_target`。
- P1 退出码：只读体检（`structure` / `indicators` / `overview`）仅警告时退出 3，`--strict` 才升为 1；JSON `ok=true`、`stats.warned`，不再把警告伪装成业务失败。
- P1 风格检查：`Option Explicit` 判定下沉到 `toolkit.style_rules`，`agent style` 与 `push` 源码预检共用同一函数；`agent style --project P` 按“项目 src + 公共库”检查，与预检同口径；文档模块按 `module_name_of` 排除，不再两套规则各说各话。
- P1 缺过程注释：改为“过程头上紧邻注释块”（只跳过空行，连续 `'` / `Rem` 算），代码、输出文案、文档统一；不再用固定回看行数造成相邻包装 Sub 结论不一致。
- P1 `vbamod`：无备份时标 `status=未判定` 并提醒“模块集合未核对”，汇总不再落“无异常”。
- P1 人话输出：源码预检不再截断文件名/提醒条数；H10 明细补各类 tag 批数与阈值，通过时也打印。
- P1 `run`：无日志且宏为 Sub（未显式 `--log`）时输出 `log_status=not_applicable` 与“该宏不写日志属正常”；新增 `--log-optional`。
- P2 `indicators` 新增 H02b“同一模块内过程不重名”，`Property Get/Let/Set` 成对不算重名；总数 20 → 21。
- P2 回读白名单：只在模块最末忽略 `()`，命中时提示忽略模块数；正文中间的 `()` 仍参与比对。
- P2 文档：`sync-maintain.md` 统一 `vbamod --project` / `push --project` / `run --project --macro`；补 `--rollback-on-fail`、`--keep-tag`、`backup_protect`、退出码 3 口径。
- 单测 82 → 90：新增警告退出码、紧邻注释、binding_target、备份保护、同尺寸同 mtime 副本、style/预检共用 Option Explicit、正文中间 `()`、写回快照复原（含 `sync_into` 开关）。



**主题：把 1.4.5–1.4.8 复查发现和安装版实测的 6 个问题一次收口；错误码只留一份，文档清单不再各写各的。**

- `sync.py init --lab` 的模板脚本从 `tests/` 移到 `scripts/make_lab_workbook.py`，安装包不再依赖被排除的 `tests/`；缺脚本报 `MISSING_LAB_TEMPLATE/stage=init`，生成失败报 `LAB_TEMPLATE_FAILED/stage=init`；`FAILED/unknown` 兜底补 `stage=<动词>`。
- `source_preflight` 新增硬错误：`.cls` 缺 `VERSION 1.0 CLASS`、单语句续行超过 24 行（带文件+行号）；类模块 `Me.` 引用私有字段进软提醒。`pitfalls.md` 补 §13–15，`SKILL.md` 第二节加指向。
- `audit overview` 在公共库为空/未配置时仍出项目与结构总览（`stats.modules=0`、`common_configured=false`），只有 `audit common` 允许失败；`--ignore` 只注册在 `structure/indicators/overview`，支持 `H05a,H06` 逗号写法。
- `H05a` 对「同一过程 + `On Error` + 成对还原」的安全执行式加速放行；`H06` 放行清空式写入、测试/自测模块、自带 `结束` 的进度器类；项目实测两项均 0 警告。
- `Application.Run` 的“宏不可用”先读 `AutomationSecurity`：Low（自动化实例）判 `MACRO_NOT_AVAILABLE`，只有非 Low 才判 `MACRO_DISABLED`；`book open` 改为 `force_new=True`，不再改用户已有会话的 `DisplayAlerts/EnableEvents/AutomationSecurity`。
- COM 分类只留 `toolkit.excel.classify_com_failure` 一份：修正 0x80010001 才是 `RPC_E_CALL_REJECTED`，0x8000FFFF/0x8001010E 不再归忙/断开；`attach_running_excel` 不再把 `EXCEL_CLOSED/EXCEL_BUSY` 包成 `COM_CONTEXT`；`assert` 超时只枚举一次 VBE 弹窗，消除竞态。
- `find_project` 改抛 `Failure` 后清掉 10 处不可达 `if proj is None`；`run` journal 的非隔离模式统一 `mode=attach`，不再同场景出现 `open/attach`；找不到 `mapping.json` 补 `CONFIG_NOT_FOUND/stage=config`，不再落到 `FAILED/unknown`。
- 单测 73 → 82，新增 version 信封、init 缺模板、overview 无公共库、预检三条、AutomationSecurity、H05a/H06 回归。

## 1.4.8 — 2026-11-21（b7ce394；交互可见性与错误信封补齐）

**主题：`book open` 不再开在隐藏实例；`push` 项目解析、`run --isolate` journal、`audit` 误报都有明确出口。**

- `book --action open` 成功后强制 `Visible=True`，最小化时恢复 `xlNormal`，并输出 `visible / hwnd / pid`；不再借到隐藏实例就假装打开。
- `sync.py find_project` 未知项目补 `PROJECT_NOT_FOUND / stage=resolve`，不再返回 `FAILED/unknown`。
- `run --isolate` 因表已打开失败时，journal `mode` 记录请求模式 `isolate`；非隔离模式带 `--save` 会明确提示未生效，journal 增 `save_requested / save_ignored`。
- `audit structure/overview/indicators` 新增 `--ignore <分组>`，可忽略 `manual_accel` 等稳定误伤；`indicators` 的 H05a 同步放行。
- `com_call` 遇到 Excel 退出/RPC 断开报 `EXCEL_CLOSED` 且不再空转重试，Excel 忙报 `EXCEL_BUSY`，其余回 `COM_CONTEXT`；`assert` 运行阶段的这两类错误统一为 `stage=assert`。
- `agent pick` 的 xlsm 解析失败拆为 `TARGET_NOT_FOUND`；日志读不了报 `ASSERT_LOG_UNREADABLE`。
- 单测 66 → 73，新增可见性、断开/忙、日志不可读、sync 项目解析与 audit 忽略分组回归。

## 1.4.7 — 2026-11-21（86632b1；看门狗覆盖项目解析与宏内弹窗）

**主题：目标解析失败不再空码；宏弹 MsgBox/InputBox、宏运行时错误也能单独报出来。**

- `agent pick` 的项目解析失败补 `PROJECT_NOT_FOUND` / `PROJECT_NOT_SPECIFIED`、`stage=config`；项目名写错不再只返回一句人话。
- 超时诊断新增 Excel 模态框枚举：`MsgBox/InputBox` 报 `MODAL_DIALOG`、`stage=assert`，并带弹窗正文；VBE 里的“运行时错误”报 `RUN_FAILED`、`stage=assert`，不再混成 `VBE_DIALOG`。
- `.lab` 测试簿新增 `popup` / `runtime` 两种模式；实测 `MODAL_DIALOG` 与 `RUN_FAILED` 均按预期返回。
- 单测 64 → 66，新增项目解析错误码与 Excel 模态框识别。

## 1.4.6 — 2026-11-21（06947c3；看门狗常见错误收口）

**主题：把断言阶段能遇到的“切不到活动簿、宏被禁用、VBE 通用弹窗”都给出明确错误码，不再落到空码或 `WATCHDOG_TIMEOUT`。**

- `TARGET_NOT_ACTIVE` 的主线程分支补 `code/stage`；切不到目标活动簿时退出 1，不再出现空 `error_code`。
- `MACRO_DISABLED` 不再只靠字符串：Excel 把“宏不可用/所有宏被禁用”合在一句里时，先查 `测试表格工具.运行全部断言` 是否真实存在；存在则判 `MACRO_DISABLED`，不存在判 `MACRO_NOT_AVAILABLE`。
- 新增 `VBE_DIALOG`：超时后检测到本进程 VBE 模态框、但正文读不到编译关键词时，报 `VBE_DIALOG`/`stage=vbe`，不再混成 `WATCHDOG_TIMEOUT`；正文含“编译错误 / #Else / #End If”仍优先报 `VBE_COMPILE_ERROR`。
- 单测 62 → 64，新增 `TARGET_NOT_ACTIVE` 空码回归、通用 VBE 弹窗不误报、宏禁用合并提示分流。
- 单工作簿矩阵实测：目标未打开、正常通过、断言失败、无测试表、无新日志、缺宏、编译错误、无限循环超时均按预期分流。

## 1.4.5 — 2026-11-21（1828c19；看门狗错误分流补全）

**主题：`Application.Run` 因缺宏/禁用宏抛出的 `com_error` 不再落到 `UNEXPECTED`，VBE 编译框按正文识别。**

- `agent assert` 捕获 `Application.Run` 的缺宏异常：按消息分流为 `MACRO_NOT_AVAILABLE` / `MACRO_DISABLED`，`stage=assert`、退出 1，并给出核对模块名或信任位置的 hint。
- VBE 弹窗识别改为读对话框正文，只有出现「编译错误 / Compile error / #Else / #End If」才判 `VBE_COMPILE_ERROR`；标题只是 `Microsoft Visual Basic for Applications` 的其它模态框不再误报。
- 单测 61 → 62，新增缺宏/禁用宏分流的断言；VBE 弹窗断言改为正文命中。
- 实测 `.lab` 单工作簿矩阵：目标未打开、正常通过、断言失败、无测试表、无新日志、缺宏、编译错误、无限循环超时均按预期分流。

## 1.4.4 — 2026-11-21（5aadc86；写回回读口径统一）

**主题：push 的回读校验先做和写入端相同的 Attribute 剥离，并忽略 VBE 在末尾补的游离 `()` 占位。**

- `toolkit.vba_parse` 新增 `drop_attribute_lines` / `drop_trailing_placeholder`；`replace_module_code()` 改用前者，写回和回读共用一套口径。
- `postcheck_sync()` 比对源码与 VBE 前，源码侧剥过程级 `Attribute`，VBE 侧再抹掉末尾 `()` 空续行；正文其它差异仍照常报 `POSTCHECK_FAILED`。
- 修正 `表格专用函数.bas` 第 6 行残留的 `Attribute call产品预览...` 死属性。
- 单测 59 → 61，新增过程级 Attribute 忽略与尾部 `()` 占位忽略。

## 1.4.3 — 2026-11-21（2c1c4a1；VBE 弹窗归属校正）

**主题：VBE 编译错误只认本目标 Excel 进程的弹窗，避免把别的临时工作簿的报错算到当前断言头上。**

- `agent assert` 超时诊断新增进程过滤：先用目标 `Excel.Application.Hwnd` 取 PID，再枚举同 PID 的 `#32770` VBE 对话框；别的 Excel 进程开着编译错误框时不再误报 `VBE_COMPILE_ERROR`。
- 归属不明（拿不到 PID）时宁可回退 `WATCHDOG_TIMEOUT`，不做猜测。
- 单测保持 59 项，补同进程识别与跨进程忽略断言。

## 1.4.2 — 2026-11-21（66119b0；鲁棒性收口）

**主题：把未知异常、错误信封、probe 健康判定和文本入口的剩余缺口收干净，版本升 1.4.2。**

- 非 JSON 入口补未知异常兜底：`run_cli()` 统一捕获 `Exception`，traceback 进 stderr，再输出“未预期异常”和处置提示，退出 1；只对 `BrokenPipeError` 静默退出，不再吞掉其它 `OSError`。
- `audit` 派发层的 `Failure` / `Exception` 信封补 `exit_code=1`，并透传 `suggested` / `retryable`；`sync` 与 `agent` 的 `UNEXPECTED` 也补下一步提示。
- `agent probe` 对“Excel 开着但 COM 不接话”按 `COM_CONTEXT` 失败退出，项目行仍保留在 `details`；新增 `stats.excel`（`running / no_excel / com_context`），`--strict` 时“Excel 没开”也按失败退出，供健康门禁使用。
- `agent assert` 超时后只读枚举顶层窗口，识别 VBE 的 `#32770` 编译错误对话框，独立报 `VBE_COMPILE_ERROR`（不再和普通死循环混为 `WATCHDOG_TIMEOUT`）。
- `ASSERT_NO_TABLE` / `ASSERT_FAILED` 补 `hint` 处置建议；单测 57 → 59，新增 strict probe 与 VBE 弹窗识别断言。
- 本轮按用户要求**不生成/更新发布 zip**；`VERSION` 升到 1.4.2。

## 1.4.1 — 2026-11-21（a0bea3a；预检补全与验收收口）

**主题：把写回前的源码检查、回读校验明细和 JSON 失败契约补齐，并让预检真正发生在备份/开表之前。**

- `source_preflight` 从 3 项扩成完整清单。硬错误：源码目录不存在、没有可写模块、同目录模块重名、非占位壳 `.frm` 缺同名 `.frx`、`VB_Name` 与文件名不一致、非文档 `.bas/.cls` 正文为空、源码读不出；软提醒：BOM、CRLF、缺 `Option Explicit`、末行续行符/疑似截断、跨目录同名覆盖。
- 新增 `source_has_code()`：剥导出头后排除空行、注释、`Attribute VB_`、`Option` 语句，再判断有无真实代码；文档模块按 `ThisWorkbook / SheetN / ChartN / DialogSheetN / MacroSheetN` 豁免，`.frm` 不判空（只有设计没有代码的窗体不误伤）。
- `.frx` 检查在占位壳跳过之后执行，公共库窗体占位壳不会误报；同目录重名只报硬错误，不再额外报“跨目录同名”。
- 单表 `push`、`push --all-project`、`run --push` 都改成“先预检，再开表/备份”；预检失败不开 Excel、不产生新备份，`run` 还会写 `.run` journal（`mode=preflight`）。`sync_into()` 增加 `preflight` 参数，入口已预检时传 `False`，避免重复扫描。
- `postcheck` 失败明细补行数：`模块（源码 4 行 / VBE 2 行；首个差异第 3 行）`；新增 `_code_body()` 先剥导出头、去空行再比对；工程里找不到、回读失败各有明确原因。
- `Failure` 新增 `retryable`；`sync` / `audit` / `agent` 的错误对象统一透传，三个入口的成功和失败信封都带 `exit_code`（含 `list` / `doctor` 直出与 `main()` 兜底分支）。
- 单测 39 → 46：新增 `test_source_preflight.py`（4 项）与 `test_postcheck.py`（3 项），并在 JSON / 错误契约里补 `exit_code`、`retryable` 断言；最近一次 46/46 通过，约 17.1 秒。
- 真实 `.lab` 验收：`push --project 示例表 --apply` 预检 17 模块 / 致命 0，写回 9 个模块、回读一致、保存成功；临时造缺 `.frx` 窗体可被 `SOURCE_PREFLIGHT` 拦下且无新备份。

## 1.4.1 补丁 — 2026-11-21（工作区，未打包）

- `audit.py` 缺参路径补 `fail` 导入；`members` 缺 `--module`、`project` 缺 `--project` 现在退出 2，不再 `NameError`。
- `audit --json` 的人话与表格改走 stderr，stdout 严格只留一行信封；新增真实 `overview` 单行 JSON 回归测试。
- `agent assert` 增加运行期看门狗：宏在独立线程内重新 attach、`静默=True`、显式限定工作簿；超时报 `WATCHDOG_TIMEOUT`，不再无限等待；宏运行与日志等待共用同一个 `--timeout` 总预算。
- `测试表格工具` 的 `运行全部断言` 增加顶层 `On Error GoTo 失败` 与 `失败:` 收尾，异常时恢复 `ScreenUpdating/EnableEvents`、清状态栏、关日志。
- `agent assert` 对“工作簿没有 测试表”单独报 `ASSERT_NO_TABLE`（退出 2）；`.lab` 补上 `Microsoft ActiveX Data Objects 2.8` 引用后，空白表断言已实测能正常早退，不再被 VBE 编译错误弹窗拖成 `WATCHDOG_TIMEOUT`。
- 非 JSON 入口补未知异常兜底：`run_cli()` 统一捕获 `Exception`，traceback 进 stderr，再输出“未预期异常”和退出 1；`sync` / `audit` / `agent` 都走同一入口，不再裸崩。
- `audit` 派发层的 `Failure` / `Exception` 信封补 `exit_code=1`，并透传 `suggested` / `retryable`；`sync` 与 `agent` 的 `UNEXPECTED` 也补下一步提示。
- `agent probe` 对“Excel 开着但 COM 不接话”不再静默返回 0：按 `COM_CONTEXT` 失败退出，项目行仍保留在 `details`；`ASSERT_NO_TABLE` / `ASSERT_FAILED` 补 `hint` 处置建议。
- 单测 52 → 57：新增非 JSON 未知异常、audit `Failure`/`UNEXPECTED` 信封 `exit_code`、probe `COM_CONTEXT`、assert `ASSERT_FAILED` 提示五项断言。
- 本轮按用户要求**不生成/更新发布 zip**；`VERSION` 仍为 1.4.1。

## 1.4.0 — 2026-11-21（178521a；契约与写回安全）

**主题：把 JSON 输出契约、异常路径、写回前后校验和运行记录统一成可验收的标准。**

- `--json` 信封统一为 `{ok, action, details, errors, elapsed_ms, ts, stats[, exit_code]}`；`action` 固定 `<入口>.<动词>`（`sync.push` / `audit.overview` / `agent.probe`），命令统计全部进 `stats`，`ts` 带时区。
- `errors` 统一为对象数组：`code / stage / message / hint / suggested_command / retryable`；字符串错误由 `toolkit.base.make_envelope` 自动归一。
- 三个入口加异常兜底：连 `pythoncom` 都缺失时也输出标准信封并 `exit 1`，traceback 只进 stderr。
- 写回前新增 `source_preflight`：同目录模块重名、`VB_Name` 与文件名不一致、源码读不出时直接放弃写回（`SOURCE_PREFLIGHT`）。
- 写回后新增 `postcheck`：逐模块回读，按去导出头/空行/缩进/大小写的指纹比对源码；不一致不保存（`POSTCHECK_FAILED`），不自动回滚。
- `run --apply` 新增 `.run/<时间>-<项目>-<宏>.json` journal：记录 mode / pushed / saved / exit_code / result / error / log；不做 `--resume`。
- 新增 `references/pitfalls.md`（现象/原因/处理/禁止），`SKILL.md` 改为链接；3 份长 reference 补目录。
- `compatibility` 从 frontmatter 顶层移入 `metadata`，对齐官方校验器白名单。
- `sync-maintain.md` 章号从 §7 起连续重编号（原 §12/§13 → §10/§11），跨册引用同步修正。

## 1.3.0 — 2026-11-21（连接与执行模式）

**主题：把“表没开 / 表开着连不上”从用户操作问题变成 skill 自己的路由问题。**

- 单表 `push` 默认自动路由：表已打开 → 借已打开会话写回；`NO_EXCEL` / `TARGET_NOT_OPEN` → 自动开隐藏 Excel 打开磁盘文件、写回、保存、关闭。`--open-only` 恢复旧行为。
- `attach_running_excel(path)` 先枚举 ROT 里的工作簿 moniker，找到目标表所在 Excel 实例再连；修掉 `GetDisplayName()` 参数错误导致 ROT 枚举全跳过的问题。多 Excel 实例下不再误连“当前活动”实例。
- 表已打开但 COM 连不上时报 `COM_CONTEXT`（阶段 `attach`），不再误报“表没打开”，也不会另开副本覆盖未保存改动。
- 失败信封新增 `stage`、`error_code`、`suggested_command`：`NO_EXCEL`、`TARGET_NOT_OPEN`、`TARGET_NOT_FOUND`、`COM_CONTEXT`、`VBE_PROTECTED`、`VBE_DENIED`、`WORKBOOK_OPEN_FAILED`、`FILE_READONLY`、`SAVE_FAILED`、`TARGET_OPEN`。
- 新增只读 `plan` 动词（`--intent pull|push|run|watch`）：返回 `open / saved / vbe_state / mode / needs_open / will_save / command / risk`，不写盘、不启动 Excel。
- `run` 新增 `--isolate`：自己开隐藏 Excel 跑宏，要求目标表未打开，不触发 `Workbook_Open`；`--save` 才保存。
- 批量 `push` 改为逐张判断：已打开借会话、没打开自己开隐藏实例；不再要求先关 Excel。
- GBK 临时导入目录优先放工作簿同目录 `.vba-import-*`，失败退回系统 temp；避免 Excel 进程读不到沙箱 temp 导致 `VBComponents.Import` 报 `0x800A...`。
- `watch` 逐项目按路径连接目标实例，不再共用一个“当前活动”实例。
- `SKILL.md` 新增“先判模式”表与硬规则；`references/sync-commands.md`、`references/usage-examples.md` 同步 push 自动路由、`plan`、`run --isolate`、错误码分流。

## 许可证变更 — 2026-11-21

- `LICENSE` 由 MIT 改为 **Source-Available Non-Sale ShareAlike 1.0（SNSA-1.0）**：允许使用、修改和二开，禁止销售软件或其衍生作品；分发衍生作品必须使用同一许可证，并提供完整对应源码。
- `SKILL.md` 的 `license` 字段同步为 SNSA-1.0。
- 本次变更仅适用于新分发的副本；此前已按 MIT 获得的授权不因本次变更收回。
- `releases/vba-workbench-1.3.0.zip` 已按 SNSA-1.0 重建；`vba-workbench-1.2.0.zip` 旧包内仍是 MIT，不再作为对外分发入口。

## 重命名 — 2026-11-21

- skill 标识由 `vba-tools` 改为 `vba-workbench`；目录、脚本路径、文档、评测、CI 与 `--version` 输出同步。
- 配置仅使用 `VBA_WORKBENCH_CONFIG` 与 `%APPDATA%\VBA工程管家`；删除旧名兼容。
- 包版本仍为 `1.2.0`，行为不变。
- `description` 收紧为“操作 VBA 工程”而不是“提到 VBA 就触发”，排除纯语法问答、Excel 教学、WPS/Mac 与文件转换。

## 验收 — 2026-11-21

- 改名后 30 个命令动词在 `.lab\技能验收` 全量实跑：`sync.py` 15 个、`audit.py` 6 个视角、`agent.py` 9 个；`pull-open` / `push` / `run` / `assert` / `watch --seconds 5` 全退出 0。
- `audit indicators` 的 H16/H17 接入已打开 Excel 的 COM 实测，H18 改为“当前工作簿 vs 最近备份模块清单”的离线对账；测试表复跑 **20 通过 / 0 警告 / 0 未跑**。
- `agent.py book --action open` 在没有 Excel 实例时补建可见实例，不再报“操作无法使用”。
- 单测 **39/39**；发布包 `releases/vba-workbench-1.2.0.zip` 已生成，`latest.txt` 与 `SHA256SUMS.txt` 已更新。

## 1.2.0 — 2026-11-21（提交 `a937d21`）

**主题：把工具链收成“可交给任何人用的标准 skill”，并补齐 agent 友好的入口。**

- `agent.py probe --json` 改为统一信封（`action=probe`、`details` 为项目行、`projects` 计数），不再直接吐数组。
- `agent.py push` 改为调用 `sync.py push --apply --json` 并解析信封，不再抓中文输出、不再喂 `y`；子进程固定 UTF-8。
- 新增 `sync.py init`：生成 `mapping.json` 与 `src` 骨架；默认预览，`--apply` 落盘；`--config PATH` 可显式指定目标（测试与新机器用），`--lab` 生成空白测试表。
- 新增全局 `--version`（读包内 `VERSION`）；`sync.py` / `audit.py` 裸跑只打印帮助并退出 2，不再因 TTY 进交互菜单。
- 菜单只认显式 `-i`；`scripts/启动菜单.bat` 两条启动命令都带 `-i`，双击仍进菜单。
- `watch` 新增 `--seconds N` 有界运行；仍不支持 `--json`。
- `doctor` 的 `elapsed_ms` 改为真实计时；必需文件 9 项（配置样例移到 `assets/`）。
- `toolkit/cli.py` 的 stdout/stderr 统一 UTF-8（`errors=replace`）。
- 配置样例 `mapping.example.json` 从包根移到 `assets/`，与“包内只留样例、真实配置在仓库根”一致；所有提示与文档同步。
- `SKILL.md` 重写：description 收紧并补英文触发词，正文压到 3.4k 字符（80 行），新增 init / `--version` / `-i` / `--seconds` / `--config`；长示例移入 `references/usage-examples.md`。
- 新增 `evals/evals.json`（12 条行为用例，带 assertions）与 `evals/trigger-evals.json`（22 条触发/不触发），并加 `evals/README.md` 说明跑法。
- `references/toolkit.md` 加目录、修正目录树（子模块归位到 `toolkit/`、补 `style_rules.py`）。
- 单测 24 → 32（probe 信封、init 四条、入口契约四条）；全量 `Ran 32 tests ... OK`。
- 验收修复：`watch` 注册 `--no-backup`；`agent logs` 注册 `--tail`（0=不打）；`agent` 非 0 JSON 统一补 `errors`；单测 32 → 35，全量 35 项通过。
- 二次验收修复：批量 `push --all-project --json` 恢复每项目明细；`sync.py -i` / `audit.py -i` 非 TTY 不再 traceback；`watch --no-input` 冲突时改为跳过；`sync/audit --no-input` 修正漏 import 的 NameError；`gbk-export --apply` 先清空目标组再重建；`shells --json` 补 `dry`；`agent backup` 默认落 `.backup/<时间戳>-agent`；`metrics --out` 写报告要 `--apply`；`agent.py` 头部示例与真实参数对齐。单测 35 → 39。

## 1.1.0 — 2026-11-21（提交 `595ddee`）

**破坏性：三个入口统一成一套命令标准，删除全部旧别名与兼容写法。**

- 命令形状统一为 `<工具> <动词> [--project P] [--module M] [--macro 名] [--json] [--apply]`；`sync.py` / `audit.py` / `agent.py` 全部改 `argparse` 子命令，未知动词退出 2。
- **`--apply` 是唯一的写开关**：写盘/执行动词默认只预览（`dry:true`，不碰 Excel、不写盘）；只读动词直接出结果。
- 删除：`export` / `export-active` / `all` 别名、`sync.py <项目名>` 裸魔法、`--yes`、`--dry-run`、`--clean`、`--save`、`audit.py` 的 `--json PATH`、`agent.py` 的 `--dry/path/logdir/md`。替代：`--mode clean`、`--no-save`、`--out PATH`、`--file`、`--scope`、`--log`。
- 失败语义统一：`exit 0 ⇔ ok=true`；非 0 时 `errors` 非空（版本不一致、体检不合格、结构问题都算失败）；`exit 2` 为用法错（argparse 原生，不出 JSON）。
- `audit.py` 六视角（members/common/project/structure/indicators/overview）共用 `_emit_audit` 信封；`agent.py` 新增 `logs` 子命令，写类动词统一 `duration/dry` 门禁。
- 修 `clean-backup` 统计字段被局部变量顶掉；`push --all-project` 分支缩进；`structure` 元组序列化。
- 单测 23 → 24（JSON 契约、错误契约按新语法重写）；`doctor` 18 项全绿。`SKILL.md` 与全部 `references/` 按新标准重写。

## 1.0.5 — 2026-11-21

**行为变更**：`sync.py` 的 `--json` 从 5 个子命令扩到「除 `watch` 外全部」；`watch --json` 改为明确报错退出 2。
失败时 `errors` 一定有内容（此前 `promote` 找不到模块时 `errors` 为空）。

- `pull-open` / `version` / `shells` / `align` / `promote` / `gbk-export` / `clean-backup` / `vbamod` 全部支持 `--json`，
  字段见 `references/sync-commands.md §3.4`。
- `promote` 的「源码目录不存在 / 模块没找到 / 公共库已存在同名文件」由 `fail()` 改为抛 `Failure`，JSON 能带回原因。
- `clean-backup` 修内部变量与 `stats` 形参重名（统计字段原先丢失）。
- 单测 20 → 23（新增 JSON 覆盖、`clean-backup` 字段、`promote` 失败原因）。

## 1.0.4 — 2026-11-21

**行为变更 / 破坏性**：`push` 的 JSON 现在带逐模块 `details` 与 `backup_path`；`run` 的 JSON 多了 `result`
（Function 返回值为 `null` 表示 Sub）与失败 `error`。只读 `ok` 的调用方不受影响。

- **`run --json` 回传返回值**：`示例求和(2,3)` 现在在 `details[0].result` 里给出 5；宏报错时 `errors` 有原因。
- **`push` / `pull` 逐条明细**：`details[]` 每项目/每次写回一条（`modules / saved / reason`），`push` 单张另给 `backup_path`。
- **`agent.py` 全局 `--json`**：任意子命令都只吐一行 JSON（人话进 `lines`，`probe` 保持自己的行数组）。
- **`list --json` 真计时**：`elapsed_ms` 不再恒 0。
- **`doctor` 加包自检**：`VERSION` 与 `SKILL.md` 一致、必需文件齐全、Python ≥3.10（现 18 项）。
- **自测 18 → 20 项**：`mapping.example.json` 可解析、`agent.py probe --json` 单行。
- **交付治理**：新增 `EVAL.md`（评测报告）、`tests/MANUAL.md` + `tests/make_lab_workbook.py`（空白表 5 步验收）、`references/release.md`（发布检查清单 + SHA256/latest 指针）。
- **文档**：`sync-commands.md` 新增「退出码与 JSON 字段」总表；README 去掉「没有测试框架」的过时说法。

## 1.0.3 — 2026-11-21

- **结构体检退出码**：`_issues_total` 不再把已登记提醒（`shadow_known`）算成问题，`audit.py structure` 现在「无异常」就退出 0，可以直接当 CI/agent 门槛。
- **`audit.py structure --json -`**：`-` 当 stdout 处理并输出单行 JSON（原先被当成文件名，会在仓库根写一个叫 `-` 的文件）。
- **失败也有原因**：`find_project` 改为抛 `Failure`，`push 不存在项目 --json` 的 `errors` 现在带「没有名为…可用…」；Excel 未打开的报错补「不替你开表、先打开工作簿」的指引。
- **自测 16 → 18 项**：结构体检计数门槛、`--json` 失败原因契约。

## 1.0.2 — 2026-11-21

- **`--yes`**：写回确认自动通过（`sync.py push` / `agent.py push` 用），与 `--no-input` 互斥并报用法错误；修掉 `agent.py push` 在管道里被「非交互=不确认」挡住、静默不写回的回归。
- **`--json` 扩到写盘命令**：`pull` / `push`（含 `--all`）/ `run` 也输出单行 JSON（`projects/exported/failed`、`results/ok/total`、`macro/args/log`）；`msg/warn` 在 JSON 模式下全部转 stderr，stdout 严格一行。
- **doctor 加平台检查**：非 Windows 直接 FAIL，不再到 COM 调用才报错。
- **自测扩到 16 项**：JSON 单行契约（含覆盖层干扰）、`resolve_target` 版本择优与噪声排除、坏 `vbaProject.bin` 报错、SKILL 版本与 `VERSION` 一致。
- **CI**：`.github/workflows/ci.yml`（windows-latest，`compileall` + `unittest`）。
- **其他**：补 `LICENSE`（MIT）；`启动菜单.bat` 改为 `py -3` 优先、找不到 Python 给提示；`sync-config.md` §4.3 标题改「示例写法」；审计示例路径改 `示例模块`。

## 1.0.1 — 2026-11-21

- **文档拆分（T34 结清）**：原 840 行 `references/sync.md` 拆成四册——`sync.md`（总览：§0/§2/§6，290 行）、`sync-commands.md`（§1/§3/§10/§11，177 行）、`sync-config.md`（§4/§5，167 行）、`sync-maintain.md`（§7–§9/§12/§13，242 行）；章节号沿用，`sync.md` 顶部加分册索引。
- **配置覆盖层**：新增 `mapping.local.json`（跟随主配置、不入库）：`defaults` 按 key 覆盖，`projects` 按 name 合并/追加，`_drop` 临时剔除；`sync.py list` 会打印「已合并本机覆盖层」。
- **自测**：新增 `tests/test_local_overlay.py`（5 项），合计 10 项全过。
- **交付**：新增 `releases/vba-tools-1.0.1.zip`（不含 `__pycache__`，根目录名 = 技能名）；`SKILL.md` 补 8 条示例提示词。

## 1.0.0 — 2026-11-21

首个可对外分发的 skill 版本。

- **打包标准化**：目录改为 `SKILL.md` + `scripts/` + `references/` + `tests/`；`tools/` 全部 `git mv` 到 `vba-workbench/`，不保留映射与复制。
- **非交互**：三个入口统一支持 `--no-input`；没有 TTY 时裸跑打印用法并退出 2，不再卡在 `input()`。
- **结构化输出**：`sync.py list --json`、`sync.py doctor --json`、`audit.py indicators --json -` 输出单行 JSON（`ok / action / details / errors / elapsed_ms / ts`）。
- **doctor 升级**：从 6 项只显示，升级为 15 项带判定（pywin32 / Python / AccessVBOM / 配置来源 / 仓库可写 / 项目解析 / 源码目录 / 公共库），有不合格项退出 1。
- **配置查找**：`VBA_TOOLS_CONFIG` → 包根向上逐级 `mapping.json` → `%APPDATA%\VBA工具\mapping.json`；与工具同层的旧位置只作兼容并提示。
- **文档**：新增 `SKILL.md`、`QUICKSTART.md`、`requirements.txt`、`VERSION`、本文件；`references/` 四份沿用原 `sync/audit/toolkit/agent` 文档。
