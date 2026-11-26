# vba-workbench 1.5.6 基线覆盖对照（COVERAGE.md）

> **命令源**：`scripts/toolkit/commands.py` 的 manifest（31 动词：sync 16 / agent 9 / audit 6），
> 不另立命令清单。**基线**：`baseline-1.5.6` @ `693d491`；当前 worktree HEAD `d0e9c2c`，VERSION `1.6.3`。
> **配套机读台账**：`tests/baseline/1.5.6/coverage.json`（字段与本表一一对应）。
> **本轮范围**：不启动 Excel、不跑真实写回，COM 全部打桩；三类 golden 未录，见第 4 节。

## 1. 状态口径

### 1.1 用例状态（case status，coverage.json `status`）

| 状态 | 含义 |
| --- | --- |
| covered | 有测试证据：`类.方法` + `文件:行`，行号经脚本逐条校验 |
| missing | 本轮无对应用例，记「缺」，不得视为已覆盖 |

### 1.2 迁移状态（migration_status，动词级）

| 状态 | 含义 |
| --- | --- |
| verified | 命令面与 1.5.6 基线差分等价（diff-report `match`），无未登记差异 |
| approved-diff | 差异已登记 `approved_deltas.json` 白名单（有意修复/规范变更/已裁决统一化） |
| pending | 存在待裁决真差异或缺口，尚未登记白名单，不得视为已迁移完成 |

### 1.3 四类用例（coverage_type）

| 类型 | 判定口径 |
| --- | --- |
| success | 预览/干跑或打桩成功链：退出 0、`ok=true`、stats/details 形状正确 |
| business_fail | 业务失败：退出 1（或约定的 3）、`ok=false`、errors 带已登记错误码 |
| usage_error | 用法错：解析层或 handler 用法校验退出 2，stdout 不出 JSON 信封 |
| json_purity | `--json` 时 stdout 仅一行信封；不支持 `--json` 的命令（watch）登记为拒绝 `--json` 用例 |

## 2. 总览

### 2.1 迁移状态计数（31 动词）

| 状态 | 数量 | 动词 |
| --- | --- | --- |
| verified | 25 | sync.version, sync.list, sync.vbamod, sync.plan, sync.pull, sync.pull-open, sync.init, sync.push, sync.shells, sync.align, sync.promote, sync.gbk-export, sync.clean-backup, sync.run, sync.watch, agent.probe, agent.book, agent.backup, agent.logs, audit.members, audit.common, audit.project, audit.structure, audit.indicators, audit.overview |
| approved-diff | 2 | sync.doctor, agent.push |
| pending | 4 | agent.assert, agent.patch, agent.metrics, agent.style |

### 2.2 四类用例覆盖率（有证据的动词数 / 31）

| 用例类型 | 覆盖 | 缺失动词 |
| --- | --- | --- |
| 成功（success） | 31/31 | — |
| 业务失败（business_fail） | 30/31 | sync.shells |
| 用法错（usage_error） | 31/31 | — |
| JSON 纯度（json_purity） | 31/31 | — |

### 2.3 证据缺失集中点

- **业务失败缺例**：`sync.shells`——仅有 create/clean 预览与非法 `--mode` 用法错，没有目标缺失/预检失败类业务失败用例。
- **待裁决缺口（pending）**：`agent.assert` / `agent.patch` / `agent.metrics` / `agent.style` 的 CLI `--json` 接受度为真差异，未登记白名单；
  进程内单行信封渲染有测试，但端到端 `--json` 行为需在 07 切片复核裁决后补录（见 `tests/test_release_diff.py` `KNOWN_REAL_GAPS`）。

## 3. 逐动词对照

### 3.1 sync 入口（16）

| 动词 | handler | 参数面 | 测试模块 | 成功 | 业务失败 | 用法错 | JSON 纯度 | 状态 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `version` | `features.core:version` | `--json`；`--no-input` | `tests/test_core_sync_entry.py` | VersionHandlerTest.test_version_success（tests/test_core_sync_entry.py:128） | VersionHandlerTest.test_version_mismatch_exit1（tests/test_core_sync_entry.py:140） | VersionHandlerTest.test_version_rejects_apply_exit2（tests/test_core_sync_entry.py:151） | VersionHandlerTest.test_version_json_single_line（tests/test_core_sync_entry.py:157） | verified |
| `list` | `features.core:list_projects` | `--json`；`--no-input` | `tests/test_core_sync_entry.py` | ListHandlerTest.test_list_success（tests/test_core_sync_entry.py:194） | ListHandlerTest.test_list_target_not_found_exit1（tests/test_core_sync_entry.py:207） | ListHandlerTest.test_list_rejects_apply_exit2（tests/test_core_sync_entry.py:218） | ListHandlerTest.test_list_json_single_line（tests/test_core_sync_entry.py:224） | verified |
| `doctor` | `features.core:doctor` | `--json`；`--no-input`；`--verify-package` | `tests/test_core_sync_entry.py` | DoctorHandlerTest.test_doctor_success（tests/test_core_sync_entry.py:253） | DoctorHandlerTest.test_doctor_failed_exit1（tests/test_core_sync_entry.py:262） | DoctorHandlerTest.test_doctor_rejects_apply_exit2（tests/test_core_sync_entry.py:315） | DoctorHandlerTest.test_doctor_json_single_line（tests/test_core_sync_entry.py:321） | approved-diff |
| `vbamod` | `features.sync_read:vbamod` | `--project`；`--json`；`--no-input` | `tests/test_sync_read.py` | TestVbamod.test_vbamod_success（tests/test_sync_read.py:96） | TestVbamod.test_vbamod_project_not_found（tests/test_sync_read.py:119） | TestVbamod.test_vbamod_usage_error（tests/test_sync_read.py:127） | TestVbamod.test_vbamod_json_purity（tests/test_sync_read.py:132） | verified |
| `plan` | `features.sync_read:plan` | `--project`（必填）；`--json`；`--no-input`；`--intent choice`={pull/push/run/watch}（默认 push） | `tests/test_sync_read.py` | TestPlan.test_plan_success_closed（tests/test_sync_read.py:154） | TestPlan.test_plan_target_missing（tests/test_sync_read.py:186） | TestPlan.test_plan_usage_error（tests/test_sync_read.py:198） | TestPlan.test_plan_json_purity（tests/test_sync_read.py:215） | verified |
| `pull` | `features.sync_read:pull` | `--project` / `--all-project`（互斥必填）；`--apply`；`--json`；`--no-input`；`--no-backup` | `tests/test_sync_read.py` | TestPull.test_pull_success_dry（tests/test_sync_read.py:236） | TestPull.test_pull_target_missing（tests/test_sync_read.py:250） | TestPull.test_pull_usage_error_neither_project（tests/test_sync_read.py:393） | TestPull.test_pull_json_purity（tests/test_sync_read.py:403） | verified |
| `pull-open` | `features.sync_read:pull_open` | `--project`（必填）；`--apply`；`--json`；`--no-input`；`--no-backup` | `tests/test_sync_read.py` | TestPullOpen.test_pull_open_success_dry（tests/test_sync_read.py:435） | TestPullOpen.test_pull_open_project_not_found（tests/test_sync_read.py:449） | TestPullOpen.test_pull_open_usage_error（tests/test_sync_read.py:456） | TestPullOpen.test_pull_open_json_purity（tests/test_sync_read.py:461） | verified |
| `init` | `features.sync_push:init` | `--project`（必填）；`--apply`；`--json`；`--no-input`；`--dir str`；`--prefix str`；`--src str`；`--config str`；`--lab` | `tests/test_sync_push.py` | InitTest.test_preview_does_not_write（tests/test_sync_push.py:241） | InitTest.test_invalid_json_is_business_failure（tests/test_sync_push.py:270） | InitTest.test_missing_project_is_usage_error_exit_2（tests/test_sync_push.py:282） | InitTest.test_json_envelope_single_line（tests/test_sync_push.py:299） | verified |
| `push` | `features.sync_push:push` | `--project` / `--all-project`（互斥必填）；`--apply`；`--json`；`--no-input`；`--doc-modules`；`--no-save`；`--no-backup`；`--open-only`；`--rollback-on-fail` | `tests/test_sync_push.py` | PushTest.test_preview_success_real_preflight（tests/test_sync_push.py:423） | PushTest.test_target_not_open_is_business_failure（tests/test_sync_push.py:492） | PushTest.test_missing_project_is_usage_error_exit_2（tests/test_sync_push.py:884） | PushTest.test_json_envelope_single_line（tests/test_sync_push.py:1127） | verified |
| `shells` | `features.sync_push:shells` | `--apply`；`--json`；`--no-input`；`--mode choice`={create/clean}（默认 create） | `tests/test_sync_push.py` | ShellsTest.test_create_preview_does_not_write（tests/test_sync_push.py:183） | 缺 | ShellsTest.test_bad_mode_is_usage_error_exit_2（tests/test_sync_push.py:213） | ShellsTest.test_json_envelope_single_line（tests/test_sync_push.py:218） | verified |
| `align` | `features.sync_transform:align` | `--project`；`--apply`；`--json`；`--no-input` | `tests/test_sync_transform.py` | AlignTest.test_success_preview_lists_same_name（tests/test_sync_transform.py:117） | AlignTest.test_business_failure_unknown_project（tests/test_sync_transform.py:128） | AlignTest.test_usage_error_bad_flag_exits_2（tests/test_sync_transform.py:136） | AlignTest.test_json_purity_single_envelope（tests/test_sync_transform.py:142） | verified |
| `promote` | `features.sync_transform:promote` | `--project`（必填）；`--apply`；`--json`；`--no-input`；`--module str`（必填） | `tests/test_sync_transform.py` | PromoteTest.test_success_preview（tests/test_sync_transform.py:164） | PromoteTest.test_business_failure_module_not_found（tests/test_sync_transform.py:193） | PromoteTest.test_usage_error_missing_args_exits_2（tests/test_sync_transform.py:211） | PromoteTest.test_json_purity_single_envelope（tests/test_sync_transform.py:220） | verified |
| `gbk-export` | `features.sync_transform:gbk_export` | `--project` / `--all-project`（互斥必填）；`--apply`；`--json`；`--no-input` | `tests/test_sync_transform.py` | GbkExportTest.test_success_preview（tests/test_sync_transform.py:243） | GbkExportTest.test_business_failure_unknown_project（tests/test_sync_transform.py:269） | GbkExportTest.test_usage_error_no_project_exits_2（tests/test_sync_transform.py:275） | GbkExportTest.test_json_purity_single_envelope（tests/test_sync_transform.py:281） | verified |
| `clean-backup` | `features.sync_transform:clean_backup` | `--apply`；`--json`；`--no-input`；`--keep int`；`--keep-tag append` | `tests/test_sync_transform.py` | CleanBackupTest.test_success_preview_defaults（tests/test_sync_transform.py:305） | CleanBackupTest.test_business_failure_bad_config（tests/test_sync_transform.py:333） | CleanBackupTest.test_usage_error_non_integer_keep_exits_2（tests/test_sync_transform.py:341） | CleanBackupTest.test_json_purity_single_envelope（tests/test_sync_transform.py:347） | verified |
| `run` | `features.run:run` | `--project`（必填）；`--apply`；`--json`；`--no-input`；`--macro str`（必填）；`--arg append`（默认 []）；`--push`；`--isolate`；`--save`；`--tail int`（默认 40）；`--log str`；`--log-optional`；`--rollback-on-fail` | `tests/test_run.py` | RunFeatureTest.test_dry_run_preview_touches_no_excel（tests/test_run.py:291） | RunFeatureTest.test_business_failure_exit1_with_code（tests/test_run.py:401） | RunFeatureTest.test_parser_usage_exit2（tests/test_run.py:434） | RunFeatureTest.test_run_json_is_single_line（tests/test_run.py:440） | verified |
| `watch` | `features.run:watch` | `--project`；`--apply`；`--json`（不支持，拒绝）；`--no-input`；`--interval float`（默认 1.0）；`--startup choice`={ask/src/vbe/each/skip}（默认 ask）；`--no-save`；`--no-backup`；`--doc-modules`；`--seconds int`（默认 0） | `tests/test_run.py` | WatchFeatureTest.test_target_resolution_and_probe_real_chain（tests/test_run.py:474） | WatchFeatureTest.test_no_target_exit1（tests/test_run.py:510） | WatchFeatureTest.test_parser_bad_choice_exit2（tests/test_run.py:629） | WatchFeatureTest.test_reject_json_exit2_no_json_on_stdout（tests/test_run.py:597） | verified |

### 3.2 agent 入口（9）

| 动词 | handler | 参数面 | 测试模块 | 成功 | 业务失败 | 用法错 | JSON 纯度 | 状态 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `probe` | `features.agent_io:probe` | `--project`；`--json`；`--no-input`；`--strict` | `tests/test_agent_io.py` | ProbeContract.test_probe_success_json_pure（tests/test_agent_io.py:114） | ProbeContract.test_probe_unknown_project_business_failure（tests/test_agent_io.py:125） | ProbeContract.test_probe_usage_error_exit_2（tests/test_agent_io.py:134） | ProbeContract.test_probe_success_json_pure（tests/test_agent_io.py:114） | verified |
| `book` | `features.agent_io:book` | `--project`；`--apply`；`--json`；`--no-input`；`--action choice`={open/close}（必填） | `tests/test_agent_io.py` | BookContract.test_book_open_preview_success_json_pure（tests/test_agent_io.py:173） | BookContract.test_book_unknown_project_business_failure（tests/test_agent_io.py:183） | BookContract.test_book_missing_action_usage_error_exit_2（tests/test_agent_io.py:192） | BookContract.test_book_open_preview_success_json_pure（tests/test_agent_io.py:173） | verified |
| `backup` | `features.agent_io:backup` | `--project`；`--apply`；`--json`；`--no-input`；`--tag str` | `tests/test_agent_io.py` | BackupContract.test_backup_preview_success_json_pure（tests/test_agent_io.py:206） | BackupContract.test_backup_unknown_project_business_failure（tests/test_agent_io.py:215） | BackupContract.test_backup_usage_error_exit_2（tests/test_agent_io.py:223） | BackupContract.test_backup_preview_success_json_pure（tests/test_agent_io.py:206） | verified |
| `logs` | `features.agent_io:logs` | `--project`；`--json`；`--no-input`；`--log str`；`--tail int`（默认 40） | `tests/test_agent_io.py` | LogsContract.test_logs_success_json_pure（tests/test_agent_io.py:237） | LogsContract.test_logs_not_found_business_failure（tests/test_agent_io.py:250） | LogsContract.test_logs_usage_error_exit_2（tests/test_agent_io.py:260） | LogsContract.test_logs_success_json_pure（tests/test_agent_io.py:237） | verified |
| `push` | `features.agent_io:push` | `--project`；`--apply`；`--json`；`--no-input`；`--doc-modules`；`--no-save`；`--no-backup`；`--open-only`；`--rollback-on-fail` | `tests/test_agent_io.py` | PushContract.test_push_preview_success_json_pure（tests/test_agent_io.py:276） | PushContract.test_push_unknown_project_business_failure（tests/test_agent_io.py:287） | PushContract.test_push_usage_error_exit_2（tests/test_agent_io.py:295） | PushContract.test_push_preview_success_json_pure（tests/test_agent_io.py:276） | approved-diff |
| `assert` | `features.agent_patch:handle_assert` | `--project`；`--apply`；`--json`；`--no-input`；`--macro str`；`--sheet str`；`--clear`；`--baseline int`（默认 35）；`--timeout float`（默认 120）；`--log str` | `tests/test_agent_patch.py` | AssertTest.test_assert_success（tests/test_agent_patch.py:438） | AssertTest.test_assert_failure_exit_1_with_code（tests/test_agent_patch.py:503） | AssertTest.test_assert_usage_error_exits_2（tests/test_agent_patch.py:591） | AssertTest.test_assert_json_is_single_line（tests/test_agent_patch.py:596） | pending |
| `patch` | `features.agent_patch:handle_patch` | `--project`；`--apply`；`--json`；`--no-input`；`--file str`（必填）；`--no-style` | `tests/test_agent_patch.py` | PatchTest.test_patch_dry_run_success（tests/test_agent_patch.py:169） | PatchTest.test_patch_format_error（tests/test_agent_patch.py:185） | PatchTest.test_patch_usage_error_exits_2（tests/test_agent_patch.py:217） | PatchTest.test_patch_json_is_single_line（tests/test_agent_patch.py:222） | pending |
| `metrics` | `features.agent_patch:handle_metrics` | `--apply`；`--json`；`--no-input`；`--out str` | `tests/test_agent_patch.py` | MetricsTest.test_metrics_success（tests/test_agent_patch.py:263） | MetricsFailureEntryTest.test_metrics_source_read_failure_is_unexpected_exit1（tests/test_agent_patch.py:652） | MetricsTest.test_metrics_usage_error_exits_2（tests/test_agent_patch.py:285） | MetricsTest.test_metrics_json_is_single_line（tests/test_agent_patch.py:290） | pending |
| `style` | `features.agent_patch:handle_style` | `--project`；`--apply`；`--json`；`--no-input`；`--action choice`={check/fix}（默认 check）；`--which choice`={all/eol/call/explicit}（默认 all）；`--scope str`（默认 src） | `tests/test_agent_patch.py` | StyleTest.test_style_success（tests/test_agent_patch.py:88） | StyleTest.test_style_business_failure_has_code（tests/test_agent_patch.py:98） | StyleTest.test_style_usage_error_exits_2（tests/test_agent_patch.py:108） | StyleTest.test_style_json_is_single_line（tests/test_agent_patch.py:113） | pending |

### 3.3 audit 入口（6）

| 动词 | handler | 参数面 | 测试模块 | 成功 | 业务失败 | 用法错 | JSON 纯度 | 状态 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `members` | `features.audit_a:handle_members` | `--json`；`--no-input`；`--module str`（必填）；`--exclude append`；`--all`；`--md path`；`--out path` | `tests/test_audit_a.py` | MembersTest.test_members_success（tests/test_audit_a.py:91） | MembersTest.test_members_business_failure_exit1_with_code（tests/test_audit_a.py:103） | MembersTest.test_members_usage_error_exit2_without_json（tests/test_audit_a.py:113） | MembersTest.test_members_json_output_is_single_line（tests/test_audit_a.py:118） | verified |
| `common` | `features.audit_a:handle_common` | `--json`；`--no-input`；`--project append`；`--exclude append`；`--all`；`--md path`；`--out path` | `tests/test_audit_a.py` | CommonTest.test_common_success（tests/test_audit_a.py:142） | CommonTest.test_common_business_failure_exit1_with_code（tests/test_audit_a.py:223） | CommonTest.test_common_usage_error_exit2_without_json（tests/test_audit_a.py:235） | CommonTest.test_common_json_output_is_single_line（tests/test_audit_a.py:240） | verified |
| `project` | `features.audit_a:handle_project` | `--project`（必填）；`--json`；`--no-input`；`--exclude append`；`--all`；`--md path`；`--out path` | `tests/test_audit_a.py` | ProjectTest.test_project_success（tests/test_audit_a.py:263） | ProjectTest.test_project_business_failure_exit1_with_code（tests/test_audit_a.py:276） | ProjectTest.test_project_usage_error_exit2_without_json（tests/test_audit_a.py:286） | ProjectTest.test_project_json_output_is_single_line（tests/test_audit_a.py:291） | verified |
| `structure` | `features.audit_b:structure` | `--json`；`--no-input`；`--exclude append`；`--ignore append`；`--strict`；`--all`；`--md str`；`--out str` | `tests/test_audit_b.py` | StructureTest.test_clean_exit_0（tests/test_audit_b.py:184） | StructureTest.test_strict_exit_1_with_code（tests/test_audit_b.py:205） | StructureTest.test_usage_error_exit_2（tests/test_audit_b.py:213） | StructureTest.test_json_single_line（tests/test_audit_b.py:220） | verified |
| `indicators` | `features.audit_b:indicators` | `--json`；`--no-input`；`--exclude append`；`--ignore append`；`--strict`；`--all`；`--md str`；`--out str` | `tests/test_audit_b.py` | IndicatorsTest.test_clean_exit_0（tests/test_audit_b.py:245） | IndicatorsTest.test_strict_exit_1_with_code（tests/test_audit_b.py:268） | IndicatorsTest.test_usage_error_exit_2（tests/test_audit_b.py:276） | IndicatorsTest.test_json_single_line（tests/test_audit_b.py:283） | verified |
| `overview` | `features.audit_b:overview` | `--json`；`--no-input`；`--project append`；`--exclude append`；`--ignore append`；`--strict`；`--all`；`--md str`；`--out str` | `tests/test_audit_b.py` | OverviewTest.test_clean_exit_0（tests/test_audit_b.py:386） | OverviewTest.test_fatal_exit_1_with_code（tests/test_audit_b.py:399） | OverviewTest.test_usage_error_exit_2（tests/test_audit_b.py:426） | OverviewTest.test_json_single_line（tests/test_audit_b.py:433） | verified |

## 4. 未录 golden（显式登记，不假装覆盖）

以下三类在本轮**未录制**：本轮约定不启动 Excel、不跑真实写回，COM 全部打桩。它们不属于上表四类用例，
将来在带 Excel 会话的本地集成环境录制后回填本节与 `coverage.json`。

| golden 类别 | 未录原因 | 将来录制入口 |
| --- | --- | --- |
| **Excel 状态 golden（CLOSED / OPEN_CLEAN / OPEN_DIRTY / OPEN_PROTECTED / OPEN_BUSY / READONLY / COM_UNAVAILABLE）** | 本轮约定不启动 Excel、不连真实 COM；测试内工作簿/会话一律替身打桩，未录制各状态机下的工作簿 probe（打开状态、保护状态、模块清单、dirty 状态）。 | 按架构方案 A.6.1 在 tests/baseline/1.5.6/commands/<entry>.<verb>/ 录制 pre_state + workbook_probe；带 @excel 标签的本地集成测试执行。 |
| **文件副作用 golden（文件树前后、备份批次、journal、.run/.backup/GBK 落盘与哈希）** | 本轮不跑真实写回；未录制 files_before/files_after（路径 + 大小 + sha256）、备份数量、journal 条目与 GBK/壳文件落点。 | 按 A.6.1 在 tests/baseline/1.5.6/side_effects/ 录制；record_1_5_6.py 扩展副作用采集段。 |
| **真实宏执行 golden（宏运行结果、断言重试与计数、保存/回滚链路）** | 本轮不启动 Excel；宏调用、参数个数重试、断言通过/失败计数、postcheck 回滚均由桩 runner 模拟，未录制真实宏执行的 stdout/日志/退出码与结果值。 | 在带 @excel 标签的本地集成环境，经 run/assert 对应命令录制宏执行 golden；入口可沿用 record_1_5_6.py 或新增宏场景段。 |

## 5. 证据规则与更新时机

- 每条证据必须能在 `tests/` 下定位到真实存在的 `类.方法` 与 `文件:行`；定位不到一律记 `missing`，不补造。
- 命令清单只认 `commands.py` manifest；新增/删除动词时先改 manifest，再同步本表与 `coverage.json`。
- **更新时机**：命令面、参数、出口码或测试用例发生变更；基线重录（`record_1_5_6.py`）后；pending 缺口裁决落定后。
- 机读字段见 `1.5.6/coverage.json` 的 `meta.migration_status_definitions` 与 `meta.evidence_rule`。
