# 1.6.1 发布验收报告

- 对应提交：`7010710`（分支 `refactor/v1.6`；代码终态 `e662e80`，文档升版 `7010710`）
- 脱敏：项目名、表名、绝对路径一律以 `<project>`、`<module>`、`<path>`、`示例表A/B/C/D`、`C:/Users/<用户>/表格/示例表A`、`V100` 占位。
- 口径：本报告只登记 1.6.1 定向修复验收；1.6 基线命令面、错误码映射与已接受设计差异见 `release-acceptance-1.6.md`，不重复改写。验收过程不启动真实 Excel、不碰业务表，预览与定向落盘均在临时目录 / 桩配置内完成。

## 1. 验收结论摘要

| 步骤 | 命令 / 口径 | 结果 | 退出码 | 结论 |
|---|---|---|---:|---|
| 编译 | `python -m compileall -q scripts tests` | 无语法告警 | 0 | 通过 |
| 全量 unittest | 32 个模块 / 418 例 | failures 0 / errors 0，`OK (skipped=9)`（通过 409） | 0 | 通过 |
| 命令表 | `gen_commands_md.py` | 覆盖 31 动词；生成物与 manifest 逐字节一致 | 0 | 功能等价 |
| 版本一致性 | `sync.py --version` | `vba-workbench 1.6.1`；`VERSION` 与 `SKILL.md metadata.version` 三处一致 | 0 | 通过 |
| 升级预览 | `upgrade_install.py --target <path>`（未加 `--apply`） | 计划交付 58 个文件；需先备份 0；待删孤儿 0；预览后目标目录 0 文件 | 0 | 未落盘 |
| 升级 apply（临时目录，非真实安装目录） | `upgrade_install.py --target <path> --apply` | 落盘 59（58 交付 + `.installed.json`）；清单 version=1.6.1、files_count=58；树内 `.backup` 0 | 0 | 可安装 |
| 差分回放 | `diff_harness.py` | match 38 / pending 0 / failed 0 | 0 | 功能等价 |

> 规模口径（1.6.1 终态 HEAD `e662e80` 实测）：**32 模块 / 418 例 / skipped 9**（通过 409）。较 1.6.0 基线 396 例新增 22 例：`test_agent_patch` +4、`test_core_sync_entry` +2、`test_init` +4、`test_release_diff` +4、`test_run` +2、`test_sync_push` +6；模块数与 skipped 不变。

## 2. 六项修复逐项验收

### 2.1 P0 异常基类穿透

- **问题**：features 跨 toolkit/adapters 边界只捕具体异常，基类 `BaseFailure` 会穿透到兜底出口；契约层失败与业务失败归类不一致。
- **改法**：跨边界统一捕 `BaseFailure`；进 `Outcome.errors` 的用 `to_contract_failure` 归一；确属纯契约层的保留 `ContractFailure`（`sync_push` 三处：`_find_project` / `_preflight_or_fail` / 批量 dry 预检）；`sync_read` 三处（vbamod / plan 探针 / pull-open）同口径收口。
- **验证证据**：提交 `99e69c0` + `23cb7ac`（回归证据）+ `7550182`（`sync_read` 三处补漏）+ `97c0f02`（见 §2.3）。用例：`test_sync_push` 新增 `test_attach_base_failure_falls_back_to_disk`、`test_attach_contract_com_context_not_fallback_to_disk`；`test_run` 新增 `test_isolate_falls_back_on_base_no_excel`、`test_isolate_non_fallback_contract_failure_propagates`。模块计数：`test_sync_push` 45 → 51、`test_run` 25 → 27。

### 2.2 P17 `agent assert` 回退

- **问题**：assert 原先先按两参签名调宏，参数个数不符即退；零参宏被误判参数错误，重试 / 报错口径不准。
- **改法**：先零参 `Run(macro)`，仅参数个数错误才退 `(False, True)`（含性能=False、静默=True）；两次都个数不符报 `ASSERT_MACRO_ARGS`（stage=assert、retryable=False）；第二次非 mismatch 原样透传。
- **验证证据**：提交 `2b12353`。`test_agent_patch` 26 → 30，新增 `test_assert_zero_arg_success_no_retry`、`test_assert_arg_mismatch_retries_two_args`、`test_assert_two_attempts_both_mismatch`、`test_assert_second_non_mismatch_propagates`。

### 2.3 P16-3 push 提示出口

- **问题**：push 附着失败提示指向 `run --isolate`，与 push 自身出口不符；建议命令不可直接执行。
- **改法**：push 附着失败提示改自身口径；`attach_running_excel` 三处 `COM_CONTEXT` suggested 支持 `target_suggested` 覆盖（不传时默认文案逐字不变），push 传自身出口。
- **验证证据**：提交 `ae67f94` + `97c0f02`。`test_sync_push` 45 → 51，新增 `test_com_context_suggested_passed_for_push_not_isolate`、`test_open_only_no_excel_hint_uses_push_not_isolate`、`test_default_suggested_retained_when_not_overridden`、`test_target_suggested_overrides_default`（另两条新增用例归属 §2.1）。

### 2.4 P18 doctor 公共库递归

- **问题**：doctor 判定公共库只从顶层 `iterdir`，嵌套目录下的模块漏判；占位壳 / 噪声文件混入计数。
- **改法**：复用 `toolkit.repo.files_under` 递归扫描，改用 `SOURCE_EXTS` 并跳过占位壳 / 噪声；空目录仍 warn。
- **验证证据**：提交 `1e76bd2`。`test_core_sync_entry` 19 → 21，新增 `test_doctor_common_dir_counts_nested_modules`、`test_doctor_common_dir_empty_only_dirs_still_warns`。

### 2.5 P19 安装清单

- **问题**：apply 逐文件落盘，无清单可核对文件版本 / 哈希；升级备份留在安装树内，旧位置 `.backup` 无处置口径。
- **改法**：apply 后原子写 `.installed.json`（version / source / installed_at / files 哈希）；升级备份迁出安装树到 `<安装根>\..\.install-backups\vba-workbench\upgrade-<时间戳>`；旧位置 `.backup` 只读提示不删；新增 `verify_manifest` 与显式 rollback。
- **验证证据**：提交 `a2e8940`。`test_release_diff` 20 → 24，新增 `test_manifest_written_and_updated_not_appended`、`test_manifest_detects_tampered_file`、`test_external_backup_restores_overwritten_file`、`test_legacy_backup_preserved_not_deleted`。定向实测：preview 58 → apply 落盘 59（58 交付 + 清单），树内 0 个 `.backup`。

### 2.6 P20 init 配置归属门禁

- **问题**：init 在非 workspace 目录也能写配置，容易把骨架落到业务目录；环境变量配置被误当作显式豁免。
- **改法**：无显式 `--config` 且目标目录不在「workspace 目录 / 生效配置目录」内 → 拒写（USAGE / Stage.CONFIG / exit 2，提示含可直接复制的 `--config` 命令）；`--lab` 未显式给配置时写 lab 目录局部 `mapping.json`；env `VBA_WORKBENCH_CONFIG` 不算显式豁免。
- **验证证据**：提交 `4c47f24` + `b6fbf8f`。`test_init` 5 → 9，新增 `test_lab_foreign_dir_refused_and_production_untouched`、`test_explicit_config_writes_lab_config_production_untouched`、`test_inside_workspace_apply_still_writes_active_config`、`test_refusal_message_carries_executable_config_command`。端到端：env 指生产配置时生产 `mapping.json` sha256 前后逐字节一致。

## 3. 全量回归结论

- 代码终态 `e662e80`；`7010710` 复跑同为 418 例 / 0 failures / 0 errors / skipped 9 / exit 0：**32 模块 / 418 例 / 409 通过 / 0 failures / 0 errors / skipped 9 / exit 0**；`python -m compileall -q scripts tests` 退出 0，无语法告警。
- 22 条新增用例全部通过，归因清晰（见 §1 规模口径与 §2 各条证据）；套件内联差分回放 `{'match': 38}`，无未登记差异。
- skipped 9 沿用既有口径，不计失败：
  - `test_cli_contract` 1：sync `-i` 菜单未迁移（1.7）。
  - `test_json_contract` 6：win32gui 弹窗识别 ×3、看门狗超时 ×1 归 1.7；宏分流码撤销、旧 exit 2 信封口径为已接受设计差异。
  - `test_watchdog` 2：看门狗原语未迁移（运行时 `skipUnless`）。

## 4. 升级预览结论

- 全新临时目录、未加 `--apply`：计划交付 **58** 份；需先备份 **0**；待删孤儿 **0**；受保护路径（`src/`、业务表 `.xlsm/.xlsb/.xlsx/.xlam/.xls/.frx`、`mapping.json`、`mapping.local.json`、`.backup/`、`.lab/`、`.run/`）保持不动。
- 预览后目标目录 0 文件，未落盘；`.installed.json` 仅在 `--apply` 路径由安装器原子落盘，预览模式不生成；清单不含 `docs/` 项（目录治理后口径保持）。
- P19 定向 apply 实测（临时安装树）：落盘 59 = 58 交付 + `.installed.json` 清单；树内 0 个 `.backup`；升级备份外置到 `<安装根>\..\.install-backups\vba-workbench\upgrade-<时间戳>`。
- `7010710` 实测预览 58、`sync.py --version` 为 `vba-workbench 1.6.1`、`.installed.json` version=1.6.1，三关（预览 / 落盘 / 清单校验）全过：apply exit 0，覆盖 58、备份 0、孤儿 0、自检通过；清单四件字段（version / source / installed_at / files）齐全、哈希复核通过；树内无 `docs/`、暂存残留 0；外置备份目录因全新目标无旧文件而未创建。
