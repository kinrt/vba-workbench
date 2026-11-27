# VBA 工具能力整合（最终收敛）

> 2026-11-27 决定：只保留 **T-03 / T-04 / T-05**，其余 9 项全部移出 1.7.x 队列。
> 执行状态以 `TODO.md` §10 为准；本文只保留这三项的范围、契约、验收与风险。

## T-03 `audit references` 引用体检

**要解决的问题**：Office 重装、换机器、加载项路径变化后，VBA 工程引用会变成 `MISSING` 或 `BROKEN`，现象通常是“编译报错但看不出是哪个引用”。

**命令**

```bat
python vba-workbench\scripts\audit.py references --project 示例表A --json
python vba-workbench\scripts\audit.py references --project 示例表A --md 其他\引用体检.md
```

**行为**

- 只读：不增删引用、不改工程、不保存工作簿；表没开时用隐藏实例只读打开，读完关闭自己开的那一份。
- 输出每条引用的 `name / description / guid / major / minor / full_path / built_in / broken / status / used_by`。
- `status` 三态：`ok`、`missing`（文件路径不存在）、`broken`（COM 报损坏）。
- `used_by` 是按已知类型前缀（`Scripting.` / `ADODB.` / `MSForms.` / `DAO.` / `SHDocVw.` / `MSComctlLib.` 等）扫源码得到的轻量提示，不是硬判定。
- 工程被锁 / 信任中心未放行：`readable=false`，退出 3，不伪报“没有引用”。

**退出码**：0 全部正常；3 有缺失/损坏或工程不可读；`--strict` 时 1；表路径/打开失败 1；用法错 2。

**验收**

- 引用清单字段齐全，`ok/missing/broken` 计数与 `details` 一致。
- 工程不可读时 `stats.readable=false`、`errors=[]`、退出 3。
- `--json` 仍为单行信封；`--md` 能落盘。
- 不启动任何写操作，不出现保存提示。

**风险**：`Reference.IsBroken` 在不同 Office 版本可能抛异常，单个属性失败只降级为空值，不拖垮整张清单。

## T-04 `audit lint` 四条静态规则

**要解决的问题**：`agent style` 关注命名/结构，但漏掉“声明了没用、重复声明、函数没返回值、Range/Cells 没限定工作表”这类容易在运行期才暴露的问题。

**命令**

```bat
python vba-workbench\scripts\audit.py lint --project 示例表A --json
python vba-workbench\scripts\audit.py lint --project 示例表A --strict --md 其他\lint.md
```

**只保留四条规则**

| 规则 | 说明 | 默认 |
|---|---|---|
| `L01_UNUSED_LOCAL` | 局部 `Dim` / `Static` 声明后未使用 | warning |
| `L02_DUPLICATE_DECL` | 同一作用域重复声明（模块级 / 过程内分别判） | warning |
| `L03_MISSING_RETURN` | `Function` / `Property Get` 没有给返回值赋值 | warning |
| `L04_UNQUALIFIED_RANGE_CELLS` | 裸 `Range(...)` / `Cells(...)`；`With` 块内放行 | warning |

**刻意不做**：全量 10 条、死存储、不可达分支、`ByVal/ByRef` 建议、隐式 Variant、自动修复。规则越多误报越多，维护成本高于收益。

**退出码**：0 无命中；3 有 warning；`--strict` 时 1；项目/源码不可读 1；用法错 2。

**验收**

- 四条规则各有正例与反例；限定 `Sheet1.Range(...)`、`With ws` 块内 `.Range/.Cells`、已使用变量、已赋值返回均不误报。
- `stats.by_rule` 四个键固定存在，计数与 `details` 一致。
- `--json` 单行；`--md` 落盘；不启动 Excel、不写源码。

**风险**：VBA 不是完整 AST，正则实现只做高置信判断；同一行多语句、极端续行可能漏报，不用来替代人工 review。

## T-05 `sync diff` 模块级差异报告

**要解决的问题**：`push` / `pull` 前只能凭记忆核对模块，批量写回容易漏看“新增/删除/过程变化”。

**命令**

```bat
python vba-workbench\scripts\sync.py diff --project 示例表A
python vba-workbench\scripts\sync.py diff --project 示例表A --all --json
python vba-workbench\scripts\sync.py diff --project 示例表A --md 其他\模块差异.md
```

**行为**

- 只读比较 **源码 vs 工作簿 VBE**，不写回、不保存、不增删模块。
- 状态以源码为基准：
  - `add`：源码有、工程无；
  - `update`：两边都有但正文不同；
  - `remove`：工程有、源码无；
  - `same`：忽略大小写/空白后相同。
- 过程级只列新增、删除、变化，不做字符级 diff。
- 公共库按项目 `common_modules` 子集比较；项目源码覆盖同名公共模块。
- `remove` 只报告，不删除；`push` 也不自动删工程模块，需人工确认。

**退出码**：0 正常出报告（有差异也是 0）；1 表路径/工程不可读；2 用法错。

**验收**

- 四个状态各有样例；`--all` 才列 `same`，默认只列差异。
- 过程级 `added/removed/changed` 与模块正文变化一致。
- `--json` 单行；`--md` 落盘；工程锁着时报 `VBE_PROTECTED`，不输出空差异。

**风险**：折叠空白/大小写后判定 `same` 可能掩盖纯格式变化；需要逐字对比时用 `vbamod`。

## 已移除项

T-06 项目清单互通、T-07 离线快照、T-08 跨项目搜索、T-09 发布干净工作簿、T-10 复杂度指标、T-11 MCP、T-12 Mermaid、T-13 格式化、T-14 宏安全预检。

原因：使用频率低，或与现有能力重叠，或回写/维护风险高于收益。以后确有真实场景再单独立项，不直接恢复旧编号。
