# 1.5.6 基线录制（脱敏 Golden Baseline）

本目录是 1.6 重构的**对照基线**：旧版 1.5.6（标签 `baseline-1.5.6`，提交 `693d491`）
的命令面与解析层场景清单。重构后任何用户可见的命令名、参数、JSON 键、退出码
都必须与本基线对齐；差异只能登记在 `approved_deltas.json`。

## 目录

```
tests/baseline/
  record_1_5_6.py        # 录制脚本（静态 introspect + 解析层子进程场景）
  COVERAGE.md            # 31 动词 × 四类用例覆盖对照（人读，证据到 文件:行）
  1.5.6/
    ledger.json          # 命令面 + 场景录制结果（生成物，可重放）
    coverage.json       # 覆盖台账机读版（与 COVERAGE.md 一一对应）
    approved_deltas.json # 已知缺陷/有意变更白名单（初始为空）
  README.md
```

## 录制范围与边界

- **不启动 Excel、不碰业务表**；场景只在临时目录 + 桩 `mapping.json` 里跑。
- 静态面：直接走旧版 `build_parser()` 的 argparse 对象，不调用 handler；
  录动词、参数、取值范围、必填、默认值、是否 `--json` / `--apply` / `--project`。
- 动态场景：裸跑（退出 2）、未知动词（退出 2）、只读动词 `--json` 信封键集合。
- 归一化：临时目录 → `<sandbox>`，绝对路径 → `<path>`，时间戳 → `<ts>`；
  耗时 `elapsed_ms` 只作参考，不进等价判定。

## 重放

```powershell
# 默认从当前 git 顶层的 baseline-1.5.6 标签用 git archive 导出（不新建工作树）
python tests/baseline/record_1_5_6.py

# 或指定已导出的旧版副本（含 vba-workbench/scripts）
$env:VBA_WB_BASELINE_DIR = "<旧版副本根>"
python tests/baseline/record_1_5_6.py
```

## 覆盖台账（COVERAGE.md / coverage.json）

`COVERAGE.md` 与 `1.5.6/coverage.json` 是 31 个动词的用例覆盖对照，命令清单只认
`scripts/toolkit/commands.py` manifest，不另立清单。

- **四类用例**（`coverage_type`）：
  - `success`：成功链退出 0、`ok=true`、stats/details 形状正确；
  - `business_fail`：业务失败退出 1（或约定的 3）、`ok=false`、errors 带已登记错误码；
  - `usage_error`：用法错退出 2，stdout 不出 JSON 信封；
  - `json_purity`：`--json` 时 stdout 仅一行信封；不支持 `--json` 的常驻命令登记为拒绝 `--json` 用例。
- **用例状态**：`covered`（有 `类.方法` + `文件:行` 证据）/ `missing`（无对应用例，记「缺」，不补造）。
- **migration_status 三态**（动词级，与 `ledger.json` 的 legacy/ported/verified 字段互不替代）：

| 状态 | 含义 |
| --- | --- |
| `verified` | 命令面与 1.5.6 基线差分等价（diff-report `match`），无未登记差异 |
| `approved-diff` | 差异已登记 `approved_deltas.json` 白名单（有意修复 / 规范变更 / 已裁决统一化） |
| `pending` | 存在待裁决真差异或缺口，尚未登记白名单，不得视为已迁移完成 |

- **用途**：补全「31 动词 / 7 解析场景 / 无 COVERAGE」缺口，作为迁移对照与回归复核入口；
  只做文档与数据文件，不改产品代码、不启动 Excel。
- **更新时机**：命令面、参数、出口码或测试用例变更；基线重录（`record_1_5_6.py`）后；
  pending 缺口裁决落定后；三类 golden（Excel 状态 / 文件副作用 / 真实宏执行）在带 Excel 会话的
  集成环境录制后回填 `coverage.json` 的 `unrecorded_golden` 段。

## 退出码口径（对齐 1.5.6）

| 退出码 | 含义 |
|---|---|
| 0 | 成功 |
| 1 | 业务失败（信封 `errors` 非空） |
| 2 | 用法错：缺参、未知动词、参数互斥（argparse 解析层） |
| 3 | 仅警告待处理（audit 部分视角，非严格模式） |
| 130 | Ctrl+C 中断 |

## approved_deltas.json 白名单格式

只有两类差异允许登记（设计方案 A.6.3）：

1. 本次有意修复的旧缺陷；
2. 规范明确变更的字段/文案。

条目结构：

```json
{
  "approved_deltas": [
    {
      "entry": "sync",
      "verb": "version",
      "input": "version --json",
      "field": "exit_code",
      "old": 1,
      "new": 0,
      "reason": "桩配置下无项目，旧版误报 VERSION_MISMATCH；新版按空清单处理",
      "since": "1.6.0"
    }
  ]
}
```

白名单之外的任何命令集合 / 参数集合 / JSON 键集合 / 退出码 / 输出结构差异，
差分 harness 一律判失败。

## 当前已知缺口

- 1.5.6 共 31 个动词；当前 manifest 已全部登记，pending 0。
- 4 条已确认真差异（**待裁决，未登记白名单**）：`agent.assert/metrics/patch/style`
  的 `--json` 位。方向：旧 1.5.6 接受 `--json`，新 manifest 登记 `supports_json=False`。
  旧版 `agent` 入口在解析前先把 `--json` 从 argv 剥走（全局开关，不经子命令选项
  注册），九动词都落通用 JSON 信封分支；`sync`/`audit` 没有预剥离，未登记即拒绝。
  归 07 切片复核；harness 目前判 failed，裁决/修复前不要写进 `approved_deltas`。
- Excel 相关场景（打开/写回/运行宏/备份）不在本基线内，由带 `@excel` 标签的
  本地集成测试覆盖；本目录只锁命令面与解析层契约。
