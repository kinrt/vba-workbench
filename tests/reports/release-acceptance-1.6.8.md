# 1.6.8 验收报告（覆盖缺口闭合与遗留收口）

- 分支：`review/1.6.8-local`
- 代码提交：`630a911`（12 files / +411 −37）
- 文档提交：`9c6340e` 起（含 `653184f`、`45af66e`）
- 基线：`review/v1.6-local` 的 1.6.6（`9f5ea49`），续接 1.6.7（`400c420` / `bd6a113`）

## 1. 结论

| 项 | 结果 |
|---|---|
| `compileall scripts tests` | 通过（exit 0） |
| 全量 unittest | **566 例 / 564 通过 / 2 skipped / 0 failure / 0 error / 73.9 秒** |
| 8 个覆盖缺口 | **16 条补测全部并入正式分支** |
| 选项分支 | `logs --tail`、`run --tail/--log/--log-optional`、`assert --clear/--baseline` 已补 |
| 四个 agent 动词 `--json` | metrics / patch / assert 补进程内 CLI 证据，style 已有 |
| 遗留代码 | N5 警告独立通道、`agent_io` 单边还原、`make_lab_workbook` 裸 `Quit` 已收 |
| 真实项目 | 「示例项目」pull 与 isolate run 通过，用户 Excel 未被影响 |
| 安装 | 1.6.8 已装 73 件，`doctor --verify-package` 通过 |
| 判定 | **可交付** |

## 2. 覆盖缺口闭合（16 例）

| 缺口 | 新增用例 | 文件 |
|---|---|---|
| `watch` 源码变更写回 | `test_source_change_writes_back_once` | `tests/test_run.py` |
| `sync align --apply` | `test_success_apply_swaps_and_backs_up` | `tests/test_sync_transform.py` |
| `sync clean-backup --apply` | `test_apply_deletes_only_over_keep` | `tests/test_sync_transform.py` |
| `agent backup --apply` | `test_backup_apply_copies_closed_workbook` | `tests/test_agent_io.py` |
| `agent patch --apply` | `test_patch_apply_writes_file_and_skips_style` | `tests/test_agent_patch.py` |
| `agent metrics --out --apply` | `test_metrics_out_apply_writes_markdown` | `tests/test_agent_patch.py` |
| `style fix --apply` / `--which` | `test_style_fix_apply_with_which`、`test_style_which_filters_explicit_only`、`test_style_which_eol_only` | `tests/test_agent_patch.py` |
| `agent book --action close` 目标簿 | `test_book_close_target_workbook_and_quit_empty_instance` | `tests/test_agent_io.py` |

同时补的选项分支：`test_logs_tail_window_and_zero`、`test_run_log_target_and_tail_window`、`test_run_tail_zero_skips_log`、`test_run_log_optional_missing_is_normal`、`test_run_default_log_directory`、`test_assert_clear_to_baseline`。

## 3. 遗留收口

### 3.1 N5 独立警告通道

- 人话模式：stdout 只保留报告行；stderr 的 `[警告]` 行剥掉统一前缀后进 `Outcome.warnings`，render 再输出到 stderr。
- JSON 模式：仍单行信封，警告走 `stats.warnings`，不混进 `stats.lines`。
- 旧“警告并入 human”的过渡口径结束；`test_agent_patch_1_6_5` 的断言同步改为独立通道。

### 3.2 `agent_io` 单边还原

- 写回保存前只调用 `guard.restore()` 一次；移除模块级 `app_guard.restore` 二次调用。
- 顺序契约：`guard.restore → wb.Save`；`test_push_apply_real_writeback_chain` 已改断言。

### 3.3 开发脚本裸 `Quit`

- `scripts/make_lab_workbook.py` 改用 `toolkit.excel.new_excel` + `quit_owned`；全仓除 `quit_owned` 内部外不再有裸 `Quit`。

### 3.4 四个 agent 动词 `--json`

- `test_agent_json.py` 新增 `test_metrics_patch_assert_json_acceptance`：metrics 正常、patch/assert 失败信封均单行 JSON；style 由既有端到端用例覆盖。

## 4. 真实项目复跑（示例项目）

项目：`<测试工作区>\示例项目实战`；工作簿 `示例表 V1.xlsm`。

| 步骤 | 结果 |
|---|---|
| `sync.py pull --project 示例项目 --json` | ok，1 模块，disk 模式，113 ms |
| `sync.py run --macro 功能自测.零参自测 --isolate --apply --json` | ok，`mode=isolate`，`run_sec=0.54`，`isolate_cleanup=ok`，未保存用户簿 |
| 安装副本 1.6.8 再跑同一宏 | ok，`isolate_cleanup=ok`，1.73 秒 |
| 用户 Excel | PID 13908（`示例项目 V261121.xlsm`）保留；进程数 14 未增 |

## 5. 安装验收

- 目标：`<安装副本>`
- `VERSION` / `SKILL.md` metadata：**1.6.8 / 1.6.8**
- `.installed.json`：73 个交付件，version 1.6.8
- `doctor --verify-package`：`ok:true`，`checked 73`、`missing []`、`mismatch []`、`version_match true`
- 升级备份：`...\vba-workbench-install-backups\vba-workbench\upgrade-20261124-121755`

## 6. 仍按设计保留

- `agent style --which explicit` 仍是“只报不修”：本版只补 `--which` 过滤测试，不自动插 `Option Explicit`；需要自动补时再单独立项。
- `sync.shells` 无业务失败面：项目缺失/无公共模块均按空结果降级，`ok=true`；这是既有设计，不新增失败码。
- 2 条 skipped 为环境条件（基线标签/环境依赖），非产品失败。
- 未 push 远端。
