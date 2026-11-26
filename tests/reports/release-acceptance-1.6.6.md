# 1.6.6 发布验收报告（review/v1.6-local，简短版）

- 对应提交：**`ba67e37`**（`fix(1.6.6): suggested 定位纠正与簿白名单 FullName 优先匹配`，4 files / +121 −12）。
- 范围：只收 1.6.5 报告 §5 标记的两处必修；其余遗留项**本轮不做**，见 §3。
- 边界：未终止任何用户 Excel 进程（PID 15220 未被终止）、未打开/保存用户工作簿；`tests/.last-full-run.log` 未纳入提交。

## 1. 结论

| 步骤 | 结果 | 退出码 | 结论 |
|---|---|---:|---|
| `compileall -q scripts tests` | 无语法告警 | 0 | 通过 |
| 全量 unittest | **Ran 547，OK（skipped 4），0 failure / 0 error**，68.7 s | 0 | 通过 |
| 必修 1：两处 suggested | 只改话术，判定逻辑未动；5 例新增断言 | — | 通过 |
| 必修 2：白名单 `FullName` 优先 | 匹配规则升级，同名多簿用例覆盖 | — | 通过 |
| 必修 3：安装清单刷新 | 73 条哈希重算、version 1.6.4 → 1.6.6、`doctor --verify-package` `ok:true` | 0 | 通过 |

**总判定：可交付（可提交）。**

## 2. 逐项

### 2.1 两处 suggested 定位纠正（commit `ba67e37`）

- 问题：`EXCEL_BUSY` 与重试耗尽的 `COM_CONTEXT` 的 suggested 仍写「仍不行再用 --isolate 或批量模式」，而 1.6.5 已确认「目标实例 COM 不可达时隔离预检会阻断隔离副本」——文案把用户引向必然失败的路径。
- 改法：统一改为「先处理忙/弹窗/受保护视图；该表仍不可用时，换一张能打开的表再走 isolate」。错误码 / stage / 放行白名单 / 退出码 / retryable 均未动；`tests/test_sync_push.py` 无需同步（该文件只钉死 attach 三分支文案，未钉死这两处）。
- 证据：新增 2 条措辞断言（不含误导措辞、含「换一张能打开的表」）。

### 2.2 `keep_workbooks` 改 `FullName` 优先、`Name` 兼容（commit `ba67e37`）

- 问题：原先只按 `wb.Name` 匹配，同名多簿（不同目录的同名 xlsm）会被一并误保留。
- 改法：名单项先按 `FullName` 归一比较（大小写不敏感 + 路径 / 斜杠归一），未命中再回退按 `Name`；命中任一即保留。残留簿问题清单改为「Name（FullName）」。签名不变，`keep_workbooks=None` 仍是旧行为。
- 证据：3 条新增用例——同名多簿只保留 `FullName` 命中那本、`Name` 旧写法兼容、路径大小写 / 斜杠归一命中；另有 2 条 suggested 措辞用例，合计新增 5 例。

### 2.3 安装清单最小刷新

- 按本轮口径**不跑安装器打包**，只做本地最小刷新：按安装副本重算 73 条交付件 sha256，改写 `installed_at` 与 `version`。
- 结果：`VERSION` = **1.6.6**；manifest 73 条、`changed hashes 12`、`added 0` / `removed 0`；`doctor --verify-package` → `ok:true`，`stats.package`：`checked 73`、`missing []`、`mismatch []`、`extra []`、`version_match true`、项目 5/5，退出 0。
- 安装方式为直接复制交付件（`SKILL.md` / `VERSION` / `LICENSE` / `requirements.txt` + `assets` / `references` / `scripts`），交付件 73 件、`__pycache__` 0。

## 3. 明写不做（遗留，保持 1.6.5 口径）

- **N5 警告通道口径**：产品口径未裁决，1.6.6 未改；当前仍是「并入报告 + 单个 `[警告]` 前缀可辨识」的过渡态。
- **归属快照键**：仍为 `id(xl)`，只适用 CLI 单线程短生命周期；长驻 / 跨线程模型未验证。
- **真实 Excel 端到端**：写回、watch 增量、真实宏执行、VBE 弹窗时序未在带真实会话环境复跑；本轮仅桩对象测试。
- **`agent_io` 双还原接缝**、**`make_lab_workbook.py` 保留裸 `Quit`**：维持 1.6.5 的判定与理由，本轮未动。
- 未 push 远端。
