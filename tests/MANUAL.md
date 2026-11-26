# 手工验收（真实 Excel 往返）

自动化能覆盖的是纯函数（选版、配置、JSON 契约）；写回与运行宏必须碰 Excel，这里是固定 5 步。
每步都只针对空白测试工作簿，**不要拿正式工作簿练手**。

> 1.4.x–1.5.x 的历史验收记录保留了原文：其中的“表”“表页”分别指工作簿、工作表，“落盘”指写入文件，“口径”指判定规则。

## 0. 准备

```
python vba-workbench/scripts/make_lab_workbook.py <仓库根>/.lab/示例表 V1.xlsm
```

在仓库根放 `mapping.local.json`，新增一个项目（示例，路径换成你的）：

```json
{"projects": [{"name": "示例表", "search_dir": "<仓库根>/.lab", "prefix": "示例表 V",
  "require_version": true, "src": ".lab/src/projects/示例表"}]}
```

`sync.py doctor --json` 应显示「包文件齐全」「VERSION 与 SKILL 一致」全 OK。

## 1. 解析

```
python vba-workbench/scripts/sync.py list --json
```

预期：示例表 `[OK]`，原因里带「按版本号 V1 取最新，候选 1 个」。

## 2. 抓取（脚本自己开隐藏 Excel）

```
python vba-workbench/scripts/sync.py pull --project 示例表 --apply
```

预期：抓 3 个模块；`src/projects/示例表` 下出现 `模块/示例模块.bas`、`Microsoft Excel 对象/*.bas`。

## 3. 写回（工作簿开没开都行）

```
python vba-workbench/scripts/sync.py push --project 示例表 --apply --json
```

预期：先打印源码预检（示例表 17 个模块、致命 0），再按情况走已打开会话或自动开隐藏 Excel；单行 JSON，`ok=true`，`details[0].modules` 有数、`details[0].saved=true`，`backup_path` 指向 `.backup/<时间戳>/`。

工作簿未打开时会自动走磁盘模式（`mode=disk`），不再报“不会替你打开工作簿”。

## 4. 运行宏（Function 返回值必须回传）

```
python vba-workbench/scripts/sync.py run --project 示例表 --macro 示例求和 --arg 2 --arg 3 --isolate --apply --json
```

预期：`details[0].result == 5`、`exit_code=0`；每次真实 run 写 `.run/<时间>-示例表-示例求和.json`。把宏名改错时 `errors[0]` 带「宏没有跑完」。

## 4.1 负向验收：预检必须发生在备份 / 打开工作簿之前

在 `.lab/src/projects/示例表/窗体/` 放一个只有 `VERSION`/`Begin`/`Attribute` 的 `坏窗体.frm`，不放同名 `.frx`，然后跑：

```
python vba-workbench/scripts/sync.py push --project 示例表 --apply --json
```

预期：`errors[0].code=SOURCE_PREFLIGHT`、`stage=preflight`、退出 1；`.backup` 不出现新时间戳目录，说明没有备份也没有打开工作簿。测完删除临时 `.frm` 与空目录。

对 `run --push` 重复同一造法，预期同样退出 1，且 `.run/*.json` 的 `mode=preflight`、`exit_code=1`。

## 4.2 性能基线（发布前手工记录）

当前 `.lab/示例表 V1.xlsm` 是空白工作簿：1 个工作表、`A1`、约 382 KB、17 个源码模块。1.4.1 实测：

| 动作 | 固定开销 | 说明 |
|---|---|---|
| `push --apply` | 13–16 秒 | 含开隐藏 Excel、写回 9 个模块、回读校验；模块数增加时按比例增长 |
| `run --isolate --apply` | 约 11 秒 | 宏本身 `cost_sec≈0.02`，其余是启动/关闭 Excel |
| `list --json` | 约 7 秒 | 5 个项目逐个解析版本号与候选文件 |

这份基线只约束空白工作簿的启动 / COM / 写回固定开销，不代表大型工作簿的数据扫描性能；有固定的大型工作簿样本后再补行列数与宏耗时。

## 4.3 `agent assert` 空白表与编译引用

`.lab/示例表 V1.xlsm` 没有名为 `测试表` 的 ListObject，**不能拿它当断言通过样本**。先打开工作簿，再跑：

```
python vba-workbench/scripts/agent.py assert --project 示例表 --apply --no-input --timeout 30 --json
```

预期：退出 2，`errors[0].code=ASSERT_NO_TABLE`，日志里写“当前工作簿里没有找到名为 测试表 的表格对象，测试中止”，不会卡到 `WATCHDOG_TIMEOUT`。

`测试表格工具` 里的 `表格工具` 类用到 `ADODB.Connection / Command / Recordset / Field`。`.lab` 工作簿必须在 VBE「工具 → 引用」里勾上 `Microsoft ActiveX Data Objects 2.8 Library`（或 6.1）；缺引用时打开工程会弹“编译错误：用户定义类型未定义”，`agent assert` 只能看到宏不返回，最终报 `WATCHDOG_TIMEOUT`，不是真死循环。新建 `.lab` 后先补引用并编译一次，再跑上面命令。

把 `测试表格工具` 临时改名后再跑 `agent assert`，预期退出 1、`errors[0].code=RUN_FAILED`、`stage=assert`，hint 带本次宏名 / 模块名与编译问题线索，不再落到 `UNEXPECTED`；改回模块名即恢复。缺宏与禁用宏不再有独立分流码，区分信息只在 hint 文案。

同一工作簿里加一个含孤立 `#Else` 的模块再跑，预期超时后报 `VBE_COMPILE_ERROR`、`stage=vbe`，对话框正文应含“# Else 子句的前面必须为 # If”。只有通用标题、正文没有编译关键词的 VBE 弹窗不应被误判为编译错误；这类通用 VBE 模态框归 `WATCHDOG_TIMEOUT`（`stage=assert`），`hint` 点名弹窗标题与正文片段，旧 `VBE_DIALOG` 细分码已不在枚举中。

另有两个 Excel 工作簿同时打开、前一个弹窗挡住目标簿时，`agent assert` 连续切 5 次仍失败，预期报 `TARGET_NOT_ACTIVE`、`stage=attach`，不能出现空 `errors[0].code`。

同一工作簿把开关设成 `popup`（宏里 `MsgBox "测试弹窗", vbOKOnly`）再跑，预期超时后报 `WATCHDOG_TIMEOUT`、`stage=assert`，`hint` 点名弹窗标题与正文“测试弹窗”；设成 `runtime`（宏里 `Err.Raise 5`）再跑，预期报 `RUN_FAILED`、`stage=assert`，正文带“运行时错误”。旧 `MODAL_DIALOG` 细分码已不在枚举中；测完在弹窗上点“确定/结束”并关掉测试实例。

项目名写错时 `agent assert --project 不存在的项目` 预期退出 1、`PROJECT_NOT_FOUND`、`stage=config`，并在 message 里列出可选项目名；不给 `--project` 报 `PROJECT_NOT_SPECIFIED`。项目名对但 mapping 指不到 xlsm 时预期 `TARGET_NOT_FOUND`、`stage=config`。

`agent book --action open` 后窗口必须可见，输出应带 `visible / hwnd / pid`；最小化实例应恢复为普通窗口，不能出现“命令说已打开、屏幕上什么都没有”。

`sync.py push --project 不存在的项目 --json` 预期退出 1、`errors[0].code=PROJECT_NOT_FOUND`、`stage=resolve`，不再出现 `FAILED/unknown`。`run --isolate` 因表已打开失败时，journal 的 `mode` 应为 `isolate`，不能是 `unknown`；非隔离模式带 `--save` 时输出“仅对 --isolate 生效，本次未保存”，journal 带 `save_requested=true / save_ignored=true`。

`audit.py structure --ignore manual_accel --json` 预期不再把 `manual_accel` 计入问题数；`audit.py indicators --ignore manual_accel --json` 的 H05a 应为“通过 / 0”。需要同时豁免状态栏时可写 `--ignore H05a,H06`（逗号分隔也生效）。

`agent probe` 在 Excel 进程存在但 COM 不接话时，预期退出 1、`errors[0].code=COM_CONTEXT`，项目行仍保留在 `details`；不会再 `ok=true` 静默放过。若只是 Excel 没开，仍按 `NO_EXCEL` 的磁盘模式输出，不强制失败。

1.4.9 新增验收：安装副本里跑 `sync.py init --project 甲 --lab --apply --json`，不应再去找 `tests/make_lab_workbook.py`；缺脚本时 `errors[0].code=MISSING_LAB_TEMPLATE`、`stage=init`，Excel 生成失败时 `LAB_TEMPLATE_FAILED`。`sync.py version --json` 必须是单行 `sync.version` 信封。无 `src/common` 的项目跑 `audit.py overview --json` 应退出 0、`stats.modules=0`、`stats.common_configured=false`；`audit.py common --json` 才允许失败。
1.5.0 新增验收：`audit.py structure` 有警告时退 3、`--strict` 退 1；`indicators` 应有 H02b 行且 `Property Get/Let` 同名不算重名；`structure` 能查到 `注册快捷键` 目标不存在/有参/非 Public；`sync.py push --help` 有 `--rollback-on-fail`；`clean-backup --keep-tag=-覆盖前` 能保护回滚点；`agent.py style --project 示例表A` 的缺 `Option Explicit` 数与 `push` 预检一致；`doctor --json` 对同尺寸同 mtime 的双版本给 `warn` 与“重复副本”提示。
1.5.1 新增验收：`push --project X` 干跑遇到续行炸弹必须直接报 `SOURCE_PREFLIGHT/stage=preflight`，不能 dry 返回 0、加 `--apply` 才炸；`audit.py overview --json` 只含 `manual_accel` 时 `exit 0 / ok=true / stats.issues=6 / stats.fatal=0`，含 `sheet_mismatch` 才退 1；`安全执行` 式三项加速成对关闭、错误分支重复还原，H05a 应为 0；`sync.py run --macro 不存在` 与 `agent assert` 的宏不可用提示都必须带“当前是 automation 实例，AutomationSecurity=Low”；`sync.py plan --intent watch --json` 在表关闭时 `command` 为 `agent.py book --project X --action open --apply`；H10 警告行必须带 `suggested_command`。
1.5.2 新增验收：表关闭时分别跑 `run` / `pull-open` / `watch`，`TARGET_NOT_OPEN` 的 `suggested_command` 必须分别含 `--isolate` / `pull` / `book --action open`，且都不能出现 `--open-only`；`book open` 后 `book close`，输出带“空 Excel 实例已退出”，对应 pid 从 `Get-Process excel` 消失；`clean-backup --keep 1 --json` 的 `details` 应逐条列批次、`stats.keep_scope=per-tag`、`stats.keep_daily/min_keep/auto_clean/by_tag` 齐全；H10 每类阈值应为 `backup_keep + backup_min_keep`（默认 13），12 批单类不再误报；`run --isolate` 的 `stats.cost_sec` 接近外层 `elapsed_ms`，`stats.run_sec` 只反映宏本身。
1.5.3 新增验收：`agent assert --project P --macro 不存在.宏 --apply --json` 应在运行前报 `ASSERT_MACRO_NOT_CONFIGURED`，`hint` 提到 `assert_macro/--macro`，`suggested_command` 是一条完整 `python ... assert ...`；`audit common --json` 无公共库时 `exit 3 / ok=true / stats.configured=false / stats.config_key=defaults.common_dir`；`clean-backup --keep 0 --json` 的 `keep_recent=1`、`keep_note` 非空、`kept_by_tag` 能解释保留总数；`SKILL.md` 与 `pitfalls.md` 都能搜到 `[加速例外]`。

源码预检三条：把 `.cls` 首行 `VERSION 1.0 CLASS` 删掉、把一条语句拆成 25 个 `_` 续行，`push` 都应在 `SOURCE_PREFLIGHT/stage=preflight` 失败并带文件行号；类模块里 `Private mCount As Long` 后用 `Me.mCount`，只应出软提醒，不阻断写回。

宏禁用差异：把 xlsm 副本放到非信任 `%TEMP%` 用 `run --isolate` 跑，`AutomationSecurity=Low` 时会正常执行；若同一句“宏不可用”，统一收口 `RUN_FAILED`，不再按旧码分流为 `MACRO_NOT_AVAILABLE` / `MACRO_DISABLED`；缺宏线索（宏名 / 模块名 / 编译问题）与禁用线索（信任位置 / 宏安全设置）都写进 `hint`，歧义句两类线索并存，不能因为文件不在信任位置就只按禁用下结论。

H05a/H06 回归：`安全执行` 式“`On Error` 后成对还原三项加速”应不再报 `manual_accel`；自带 `结束` 并清空状态栏的进度器类、`运行自测`/测试模块不应报 H06；`audit.py indicators --ignore H05a,H06 --json` 两项都应为“通过 / 0”。

```
python vba-workbench/scripts/audit.py structure --json
python vba-workbench/scripts/audit.py indicators --json
```

预期：structure 无异常时退出 0；indicators 单行 JSON，H15 计算模式 4/4（含示例表）。

完成后删掉 `.lab/` 和 `mapping.local.json` 里那条项目即可，主配置不动。
