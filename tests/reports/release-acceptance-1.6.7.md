# 1.6.7 验收报告（归属收尾稳定键与真实 Excel 冒烟）

> 后续：1.6.8（`630a911`）已并入 16 条覆盖补测并收口 N5 / `agent_io` 双还原 / 裸 `Quit`；本报告第 5 节“仍未处理”已由 1.6.8 验收报告取代。

- 分支：`review/1.6.7-local`
- 代码提交：`c0c4499`（fix，5 files / +47 −8）
- 文档提交：`400c420`（CHANGELOG 回填）
- 基线：`review/v1.6-local` 的 1.6.6（`9f5ea49`）

## 1. 结论

| 项 | 结果 |
|---|---|
| `compileall` | 通过（exit 0） |
| 全量 unittest | **549 例 / 547 通过 / 2 skipped / 0 failure / 0 error**，63.5 秒 |
| 归属键修复 | `quit_owned` 改用 `(pid, hwnd)` 稳定键，读不到才退 `id(xl)`；新增 2 例 |
| 真实 Excel 冒烟 | 隔离 lab 工作簿 pull / push / run --isolate / watch / book close 全部通过 |
| 安装 | 1.6.7 已装到 `<安装副本>`；73 件交付，`doctor --verify-package` 通过 |
| 判定 | **可交付** |

## 2. 代码修复

### 2.1 归属收尾稳定键

- 问题：`_QUIT_DONE` 原以 `id(xl)` 为键，长驻 / 多实例场景下 Python 对象地址可能复用，导致新实例被误判“已收尾”而不 Quit。
- 改法：新增 `_quit_key()`，优先用 `(pid, hwnd)` 作幂等键；只有读不到 PID/Hwnd 时才退回 `id(xl)`。归属快照比对仍按当前代理对象读取，PID/Hwnd 变化仍拒绝退出。
- 测试：`test_owned_lifecycle_1_6_5` 新增 2 例——同 Hwnd 新代理不重复 Quit、不同 Hwnd 不互相挡回；该文件 31 例全绿。

## 3. 真实 Excel 冒烟（隔离 lab）

工作簿：`示例表 V1.xlsm`，项目 `示例项目`，位于 `.lab/e2e-1.6.6`。

| 步骤 | 结果 |
|---|---|
| `sync.py pull --apply` | 3 个模块导出，`mode=disk` |
| 源码加 `示例版本()` 后 `push --apply` | 3 个模块写回，`saved=true`，备份落盘，`isolate_cleanup=ok` |
| `run --isolate --macro 示例版本` | 返回 `1.6.6-e2e` |
| `watch --seconds 20 --startup src` + 改源码 | 写回 1 次；宏返回 `1.6.6-e2e-watch` |
| `agent book close --apply` | 目标簿关闭；用户 Excel PID 13908 保留，进程数 15→14 |

未触碰用户已打开工作簿；测试用表与源码均在隔离目录。

## 4. 安装验收

- 目标：`<安装副本>`
- `VERSION` / `SKILL.md` metadata：**1.6.7 / 1.6.7**
- `.installed.json`：73 个交付件，version 1.6.7
- `doctor --verify-package`：`ok:true`，`checked 73`、`missing []`、`mismatch []`、`version_match true`
- 升级备份：`<安装备份目录>\upgrade-20261124-115828`

## 5. 仍未处理

- 8 个覆盖缺口的 16 条补测尚未并入正式分支（`review/v1.6.7-local` 仍为 549 例）。
- N5 style 警告通道口径未裁决。
- `agent_io` 双还原接缝未单边收口。
- `make_lab_workbook.py` 裸 `Quit`、`agent logs --tail` / `run --log` / `assert --clear/--baseline` 分支测试等低风险遗留未动。
- 未 push 远端。
