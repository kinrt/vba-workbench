# 1.6.2 发布验收报告

- 对应提交：分支 `refactor/v1.6`；回归终态 `52368a5`（命令绝对化收口合入），文档口径随本次 `docs: 1.6.2 口径对齐 52368a5` 提交落库（报告不内嵌自身 hash）。`52368a5` 之后主干另有用户本机提交，只动业务 `src/` 窗体文件，与 `vba-workbench` 包无关，未纳入本报告、未 stage、未改 `src/`。
- 脱敏：项目名、表名、绝对路径一律以 `<project>`、`<module>`、`<path>`、`示例表A/B/C/D`、`C:/Users/<用户>/表格/示例表A`、`V100` 占位。
- 口径：本报告只登记 1.6.2 定向修复验收；1.6 基线命令面、错误码映射与已接受设计差异见 `release-acceptance-1.6.md`，1.6.1 六项修复见 `release-acceptance-1.6.1.md`，不重复改写。验收过程不启动真实 Excel、不碰业务表，预览与定向落盘均在临时目录 / 桩配置内完成。

## 1. 验收结论摘要

| 步骤 | 命令 / 口径 | 结果 | 退出码 | 结论 |
|---|---|---|---:|---|
| 编译 | `python -m compileall -q scripts tests` | 无语法告警 | 0 | 通过 |
| 全量 unittest | 32 个模块 / 450 例 | failures 0 / errors 0，`OK (skipped=9)`（通过 441） | 0 | 通过 |
| 命令表 | `gen_commands_md.py` | 覆盖 31 动词；生成物与 manifest 逐字节一致 | 0 | 功能等价 |
| 版本一致性 | `sync.py --version` | `vba-workbench 1.6.2`；`VERSION` 与 `SKILL.md metadata.version` 三处一致 | 0 | 通过 |
| 升级预览 | `upgrade_install.py --target <path>`（未加 `--apply`） | 计划交付 58 个文件；需先备份 0；待删孤儿 0；预览后目标目录 0 文件 | 0 | 未落盘 |
| 差分回放 | `diff_harness.py` | match 38 / approved 2 / pending 0 / failed 0 | 0 | 功能等价 |

> 规模口径（1.6.2 回归终态 HEAD `52368a5` 实测）：**32 模块 / 450 例 / skipped 9**（通过 441）。相对 1.6.1 基线 418 例累计净 +32：其中八项修复（终态 `93801e7`）净 +22——`test_core_sync_entry` +8、`test_json_contract` +3、`test_release_diff` +4、`test_sync_push` +3、`test_sync_read` +4（净）；命令绝对化收口（终态 `52368a5`）再净 +10——`test_agent_io` +3、`test_audit_engine` +2、`test_core_sync_entry` +3、`test_new_issues` +2，其余 28 模块 +0；收口段新增 10、移除 0。模块数与 skipped 不变。

## 2. 八项修复逐项验收

### 2.1 ① 建议命令改绝对入口

- **问题**：原先 `suggested_command` 用包内相对路径入口，换工作目录后不可直接复制执行；提示出口与真实入口位置脱钩。
- **改法**：新增 toolkit 助手 `entry_command(entry, verb, *args)`，拼 `sys.executable` + 由包位置解析的入口绝对路径 + 参数（空格引用）；覆盖 assert `ASSERT_NO_LOG`、push `--open-only` 两类提示、init 归属门禁拒绝提示三处。
- **验证证据**：提交 `3a1630d`。`test_agent_patch` / `test_init` / `test_sync_push` 相关断言改为校验入口真实存在、非包内相对路径；模块计数：`test_agent_patch` 30、`test_init` 9、`test_sync_push` 54。

### 2.2 ② push `--open-only` 未开表分流

- **问题**：`push --open-only` 且表未打开原先冒泡 `NO_EXCEL/stage=attach`，与自动开表回退磁盘的口径混淆，hint 未指明 open-only 专属出口。
- **改法**：仅 `auto_open=False` 且 code 为 `NO_EXCEL` 时改写 `TARGET_NOT_OPEN`（stage=attach、hint 指向去掉 `--open-only` 或先开表、suggested 为绝对重试命令）；`COM_CONTEXT` / `VBE_PROTECTED` 原样上抛；自动开表路径的 `NO_EXCEL` 仍回退隐藏实例磁盘。
- **验证证据**：提交 `3a1630d`。`test_sync_push` 51 → 54，新增 `test_open_only_no_excel_push_outcome_target_not_open`、`test_auto_open_no_excel_falls_back_to_disk`、`test_open_only_com_context_propagates_unchanged`。

### 2.3 ③ pull 已打开自动借会话

- **问题**：`pull --apply` 原先未区分表开没开，可能抓磁盘副本而漏掉未保存改动；批量场景无法区分本次走的是会话还是磁盘。
- **改法**：逐项目先附着已打开会话（含未保存改动、不另开副本、不关窗、按 `--no-backup` 留底），未打开走原隐藏实例磁盘；`--all-project` 逐项目判定、结果行带 `mode: open|disk`；仅 `NO_EXCEL` / `TARGET_NOT_OPEN` 允许退磁盘，`COM_CONTEXT` / `EXCEL_BUSY` / `EXCEL_CLOSED` / `VBE_PROTECTED` 原样上抛并给「本应借会话但失败、未退磁盘」hint；`plan --intent pull` 推荐命令统一为 `pull`；dry 只做 ROT 轻量探测。
- **验证证据**：提交 `6d27801`；evals id3 与 `references/sync.md` §3.1 口径同步（`0d2856e`）。`test_sync_read` 22 → 26（净 +4），新增 6 例：`test_plan_pull_open_command_uses_pull`、`test_pull_apply_closed_goes_through_disk_gateway`、`test_pull_apply_open_uses_session_not_disk`、`test_pull_apply_mixed_open_and_disk`、`test_pull_apply_open_session_failure_not_fallback_disk`、`test_pull_apply_disk_open_failure_keeps_contract_code`；移除 2 例：`test_pull_apply_success_goes_through_gateway`、`test_pull_apply_open_failure_goes_through_gateway`。

### 2.4 ④ assert 能力边界写进文档

- **问题**：assert 原先被当作通用宏执行器，只写工作表、不写日志的自测宏也被丢进去跑，失败口径与适用范围不清。
- **改法**：明确 assert 只接会写日志的断言宏（读「通过/失败」计数）；零参 / 两参自测宏只写工作表时恒报 `ASSERT_NO_LOG`，应改用 `sync.py run`；边界写入 `SKILL.md`（常用命令表 + 边界节）与 `references/agent.md`。
- **验证证据**：提交 `0a4eefb`。文档定向更新，无新增用例；`test_agent_patch` 30 例全绿。

### 2.5 ⑤ assert 过期文档更正

- **问题**：`references/agent.md` 原先把 `ASSERT_MACRO_ARGS` 标为「1.6 不可达、随 1.7 回补」，与实际行为不符；assert 退出码口径也有陈旧表述。
- **改法**：删除「1.6 不可达」说法，按实际口径写：先零参调用、仅参数个数错误退 `(False, True)`、两次都错才报 `ASSERT_MACRO_ARGS`、第二次其它错误原样透传；顺带更正 assert 退出码口径（0 通过 / 1 业务失败 / 2 用法错）。
- **验证证据**：提交 `0a4eefb`。文档定向更正，无新增用例；`test_agent_patch` 30 例全绿。

### 2.6 ⑥ `--json` 警告收口

- **问题**：JSON 模式下 `warn()` 原先仍写 stderr，可能打断单行 JSON 信封；警告出口与 JSON 契约不一致。
- **改法**：JSON 模式下 `warn()` 不写 stderr 而缓存，`render.py` 不再向 stderr 打 warnings，两路合并进信封 `stats.warnings`（仅有警告时落键，不删既有键，不与 1.5.6 契约冲突）；三入口解析参数后即置 JSON 开关。
- **验证证据**：提交 `eceb2d0`。`test_json_contract` 33 → 36，新增 `test_version_json_warnings_collected_into_stats`、`test_version_human_warnings_still_to_stderr`、`test_legacy_config_warning_collected_in_json`。

### 2.7 ⑦ 升级备份迁出 skills 根

- **问题**：升级备份原先与安装位置耦合，可能落在 skills 发现根内被再次扫描；旧批次备份缺少只读探测与回滚顺序口径。
- **改法**：备份根与安装位置解耦，固定用户级 `%LOCALAPPDATA%\vba-workbench-install-backups\vba-workbench\upgrade-<时间戳>\`（`LOCALAPPDATA` 缺失退 `%USERPROFILE%\.vba-workbench-install-backups\`），永不落在 skills 发现根内；旧备份（`vba-workbench.bak-*`、树内 `.backup\upgrade-*`、1.6.1 同级 `.install-backups`）只读探测并给人工清理建议，绝不自动删除、不新建；空目录修剪跳过 `bak-*`；rollback 顺序为新根 → 旧同级 → 旧树内。
- **验证证据**：提交 `6082cd0`。`test_release_diff` 24 → 28，新增 `test_backup_root_never_inside_skills_root`、`test_legacy_whole_backups_listed_readonly_not_deleted`、`test_rollback_restores_from_new_backup_root`、`test_rollback_falls_back_to_legacy_location`。

### 2.8 ⑧ doctor 显式包完整性校验

- **问题**：doctor 原先没有按安装清单对账交付件的开关，本地正常改动也可能被误判；缺少清单在位 / 哈希 / 版本一致性的显式校验路径。
- **改法**：新增显式开关 `sync.py doctor --verify-package`（manifest 登记；默认 doctor 行为不变，不因本地正常改动失败）；校验清单在位 / 字段齐全 / files 与磁盘对账 / 逐项 sha256 / VERSION 与清单 version 一致，结果进 `stats.package`；显式校验不通过退 1、通过退 0。
- **验证证据**：提交 `163d3fe`。`test_core_sync_entry` 21 → 29，新增 `test_doctor_default_skips_package_check`、`test_doctor_verify_package_missing_manifest_exit1`、`test_verify_package_pass`、`test_verify_package_tamper_reports_file_and_code`、`test_verify_package_missing_file`、`test_verify_package_extra_file`、`test_verify_package_version_mismatch`、`test_verify_package_manifest_missing`。

### 2.9 ⑨ 命令绝对化收口（白名单升级）

- **问题**：安装后 cwd 不在仓库根时，`suggested_command` / hint / 人话行 / 帮助 / 报错里的 `python vba-workbench/...` 相对命令串照抄找不到文件；另 `make_envelope` 白名单原只认 `python ` 前缀，会把已拼好的绝对命令降级进 hint。
- **改法**：全部提示出口改 `entry_command()` 绝对入口（`sys.executable` + 由包位置解析的入口脚本绝对路径 + 参数，空格引用）；`toolkit/base.py` 新增 `looks_like_command()` 并在 `make_envelope` 改用它识别可执行命令，绝对命令原样保留在 `suggested_command` 独立字段、不降级进 hint；纯人话建议仍按旧口径降级（`suggested_command` 置空、文本进 `hint`）。
- **验证证据**：切片提交 `b86c720`（白名单助手）、`b98bd05`（sync 系 15 处）、`4e0877a`（agent 系 21 处）、`98b9ec0`（audit/toolkit 23 处）；合并 `8823a85` / `4a3c73f` / `52368a5`。切片定向用例：sync 系 133、agent 系 85、audit/toolkit 系 94；全量 **450 例 / 441 通过 / skipped 9 / 0 failure / 0 error / exit 0**；`scripts/**/*.py` 中 `python vba-workbench`（`/` 与 `\` 两种写法）残留 **0**。
- **新增用例归属**（基线 440 → 实测 450，净 +10）：

| 模块 | 基线 | 实测 | 差值 | 新增用例（文件:行号） | 归属切片 |
|---|---:|---:|---:|---|---|
| `test_agent_io` | 22 | 25 | +3 | `test_entry_command_uses_absolute_interpreter_and_script`（:558）、`test_envelope_keeps_absolute_suggested_command`（:568）、`test_sources_have_no_relative_package_command`（:579） | agent 系 |
| `test_audit_engine` | 9 | 11 | +2 | `test_h10_suggested_carries_absolute_script`（:157）、`test_overview_actions_carry_absolute_script`（:169） | audit/toolkit 系 |
| `test_core_sync_entry` | 29 | 32 | +3 | `test_entry_command_uses_interpreter_and_absolute_script`（:485）、`test_make_envelope_keeps_absolute_suggested_command`（:495）、`test_source_files_have_no_relative_package_command`（:506） | sync 系 |
| `test_new_issues` | 13 | 15 | +2 | `test_list_command_carries_absolute_interpreter_and_script`（:215）、`test_slice_sources_hardcode_no_package_prefixed_command`（:220） | sync list 入口 + 全切片硬编码扫描 |
| 其余 28 模块 | — | — | 0 | 无增减 | — |

## 3. 全量回归结论

- 回归终态 `52368a5`；全量复跑：**32 模块 / 450 例 / 441 通过 / 0 failures / 0 errors / skipped 9 / exit 0**；`python -m compileall -q scripts tests` 退出 0，无语法告警。
- 累计净新增 32 例（八项修复 +22、命令绝对化收口 +10）全部通过，归因清晰（见 §1 规模口径与 §2 各条证据）；套件内联差分回放 `{'match': 38, 'approved': 2}`（38 行匹配，2 条已登记差异吸收），无未登记差异。
- skipped 9 沿用既有口径，与 1.6.1 一致，不计失败：
  - `test_cli_contract` 1：sync `-i` 菜单未迁移（1.7）。
  - `test_json_contract` 6：win32gui 弹窗识别 ×3、看门狗超时 ×1 归 1.7；宏分流码撤销、旧 exit 2 信封口径为已接受设计差异。
  - `test_watchdog` 2：看门狗原语未迁移（运行时 `skipUnless`）。

## 4. 升级预览结论

- 全新临时目录、未加 `--apply`：计划交付 **58** 份；需先备份 **0**；待删孤儿 **0**；受保护路径（`src/`、业务表 `.xlsm/.xlsb/.xlsx/.xlam/.xls/.frx`、`mapping.json`、`mapping.local.json`、`.backup/`、`.lab/`、`.run/`）保持不动。
- 预览后目标目录 0 文件，未落盘；`.installed.json` 仅在 `--apply` 路径由安装器原子落盘，预览模式不生成；清单不含 `docs/` 项（目录治理后口径保持）。
- 交付数仍为 58，未变成 59：`52368a5` 相对 1.6.1 交付根无 A/D 文件，全部为原地修改；`--verify-package` 开关落在既有入口与 features 切片内，未新增独立交付文件；references 12 份、scripts 41 份与 1.6.1 预览口径一致。
- 备份落点口径变化（不影响交付数）：外置备份迁到用户级目录 `%LOCALAPPDATA%\vba-workbench-install-backups\vba-workbench\upgrade-<时间戳>`，不进 skills 发现根；旧整包 / 旧批次只读探测、不自动删除。

## 5. 新增错误码与开关

### 5.1 开关

- `sync.py doctor --verify-package`：显式按 `.installed.json` 对账交付件哈希与 VERSION；**默认关闭**，不带该开关时 doctor 行为不变，本机正常改动不会让默认 doctor 失败。

### 5.2 错误码

- `PACKAGE_MANIFEST_MISSING`：无安装清单（源码树 / 未安装副本属正常，不崩）。
- `PACKAGE_INTEGRITY`：清单字段缺失、登记文件磁盘缺失、哈希不符、磁盘多出未登记文件、VERSION 与清单版本不一致；`message` 列出具体相对路径，`hint` 说明本机正常改动可忽略、可用安装器重装或 `--apply` 刷新清单；显式校验不通过退出 1，通过退出 0。

### 5.3 信封字段

- `stats.package`：`present / checked / missing / mismatch / extra / missing_fields / version_match / manifest_version / disk_version`。
- `stats.warnings`：JSON 模式下两路警告合并落键；仅有警告时落键，不删既有键。

## 6. 测试包（交付快照）

- 命名规则：`vba-workbench-1.6.2-<HEAD 短哈希>.zip`，沿用上一版结构；打包在本次提交之后进行，短哈希与 SHA256 随后在交付汇报给出，本报告不写死。
- 包内 **110 个文件**；不含 `SHA256SUMS.txt` / `latest.txt`；不含运行时目录（`.lab/`、`.run/`、`.backup/`、`__pycache__/` 等）。
