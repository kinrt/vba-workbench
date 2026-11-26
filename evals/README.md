# evals：触发与行为验收

两类用例，改 `description` 或改命令契约后都要跑。

| 文件 | 用途 | 怎么用 |
|---|---|---|
| `trigger-evals.json` | 22 条 query（11 应触发 / 11 不应触发） | 逐条喂给 agent，只看“有没有选 vba-workbench”；应触发全中、不应触发全不中才算稳定 |
| `evals.json` | 18 条行为用例（含 6 条失败路径） | 真跑一遍，按 `assertions` 逐条核对；失败项对应到 `references/` 里缺的说明或脚本缺陷 |

## 判定标准

- **触发**：命中 description 的实体词（VBA、xlsm、vbaProject.bin、源码同步）或任务词（抓取/写回/查引用/体检）。
- **不触发**：纯 VBA 语法问答、Excel 操作教学、业务数据分析、WPS/Mac、文件格式转换。
- **行为**：命令选对、参数对、先预览后 `--apply`、只消费 JSON 信封、失败不编造成功。

## 回归节奏

1. 改 `SKILL.md` 或 `description`：先跑 `trigger-evals.json`。
2. 改脚本或 JSON 契约：跑单测 + `evals.json` 里相关条目。
3. 发布前：两组都跑，结果写进 `CHANGELOG.md`。
