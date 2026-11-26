# 示例提示词与命令行清单

## 一、可直接复制的提示词

```
用 vba-workbench 看看环境能不能跑。
用 vba-workbench 列出所有项目，确认「示例表」配到哪个工作簿。
用 vba-workbench 把全部项目的 VBA 抓到 src，跑完告诉我哪些模块有变化。
用 vba-workbench 把「示例表」源码写回工作簿，然后运行宏 示例求和 验证。
用 vba-workbench 先 plan 一下「示例表」，告诉我现在该用已打开会话、磁盘写回还是隔离实例。
用 vba-workbench 把「示例表」源码写回；工作簿未打开也直接写，不要用 Start-Process 开 Excel。
用 vba-workbench 在隐藏实例里运行「示例表」的宏，不触发 Workbook_Open，完成后保存。
用 vba-workbench 查「表格工具」这个模块被哪些工作簿引用，只分析不修改。
用 vba-workbench 做一次公共库瘦身评估，列出可下放与可合并的模块。
用 vba-workbench 检查源码命名和格式，只报告问题。
用 vba-workbench 初始化一个新项目「示例表2」，表放在 D:\VBA\示例表2。
```

## 二、完整命令行

```bat
cd D:\VBA工具

python vba-workbench\scripts\sync.py --help
python vba-workbench\scripts\sync.py --version
python vba-workbench\scripts\sync.py doctor
python vba-workbench\scripts\sync.py doctor --json
python vba-workbench\scripts\sync.py list
python vba-workbench\scripts\sync.py list --json

REM 新机器初始化（默认只预览，加 --apply 才写）
python vba-workbench\scripts\sync.py init --project 示例表 --dir "D:\VBA\示例表"
python vba-workbench\scripts\sync.py init --project 示例表 --dir "D:\VBA\示例表" --apply
python vba-workbench\scripts\sync.py init --project 示例表 --lab --apply

REM 先判断状态：工作簿是否已打开、有没有未保存改动、该用哪种能力
python vba-workbench\scripts\sync.py plan --project 示例表 --intent pull --json
python vba-workbench\scripts\sync.py plan --project 示例表 --intent push --json
python vba-workbench\scripts\sync.py plan --project 示例表 --intent run --json
python vba-workbench\scripts\sync.py plan --project 示例表 --intent watch --json

REM 抓取：工作簿已打开且存在未保存改动
python vba-workbench\scripts\sync.py pull-open --project 示例表 --apply --json
REM 抓取：工作簿已关闭（脚本自己开隐藏实例）
python vba-workbench\scripts\sync.py pull --all-project --apply

REM 写回：工作簿已打开就借用会话，未打开自动开隐藏 Excel（单工作簿默认，不用额外开关）
python vba-workbench\scripts\sync.py push --project 示例表 --apply --json
python vba-workbench\scripts\sync.py push --all-project --apply
REM 只想写已打开的工作簿，工作簿未打开就失败（旧行为）
python vba-workbench\scripts\sync.py push --project 示例表 --apply --open-only

REM 写回后运行宏：工作簿已打开时使用已打开会话
python vba-workbench\scripts\sync.py run --project 示例表 --macro 示例求和 --apply --json
REM 工作簿未打开：隐藏实例运行，不触发 Workbook_Open；--save 才保存
python vba-workbench\scripts\sync.py run --project 示例表 --macro 示例求和 --isolate --apply --json
python vba-workbench\scripts\sync.py run --project 示例表 --macro 示例求和 --isolate --save --apply

REM 自动写回（常驻；agent 不用，除非加 --seconds）
python vba-workbench\scripts\sync.py watch --project 示例表 --apply --seconds 600
```

## 三、常用只读检查

```bat
python vba-workbench\scripts\audit.py members --module 表格工具 --json
python vba-workbench\scripts\audit.py common --json
python vba-workbench\scripts\audit.py project --project 示例表
python vba-workbench\scripts\audit.py structure --json
python vba-workbench\scripts\audit.py indicators --json
python vba-workbench\scripts\audit.py overview

python vba-workbench\scripts\agent.py probe --json
python vba-workbench\scripts\agent.py probe --json --strict   # 健康门禁：Excel 没开也按失败退出
python vba-workbench\scripts\agent.py book --action open --project 示例表 --json
python vba-workbench\scripts\agent.py style --action check
python vba-workbench\scripts\agent.py metrics --json
python vba-workbench\scripts\agent.py logs --tail 50
```

## 四、怎么读 JSON

```json
{"ok": true, "action": "sync.push", "details": [...], "errors": [], "elapsed_ms": 1234,
 "ts": "2026-11-21T19:40:27+08:00", "stats": {"project": "甲", "dry": false, "saved": true},
 "exit_code": 0}
```

- 先看 `ok`；失败看 `errors[0].code` / `errors[0].stage` / `errors[0].suggested_command`，不猜。
- `errors[0].code` 单一真源按组分列（逐码处置见 `pitfalls.md`）：
  - 项目/目标：`PROJECT_NOT_FOUND`（项目名写错；`sync` 为 `stage=resolve`，`agent` 为 `stage=config`）、`PROJECT_NOT_SPECIFIED`（没给项目名）、`TARGET_NOT_FOUND`（项目名对但 mapping 指不到 xlsm）、`NO_EXCEL`（没有 Excel；`probe --strict` 也用它做健康门禁）、`TARGET_NOT_OPEN`（工作簿未打开）、`TARGET_NOT_ACTIVE`（切不到目标活动簿，先关挡在前面的弹窗）、`TARGET_OPEN`（隔离模式但工作簿已打开）、`WORKBOOK_OPEN_FAILED`（密码/占用/受保护视图）。
  - Excel/COM：`COM_CONTEXT`（工作簿已打开但连不上，probe 也会直接失败）、`EXCEL_CLOSED`（Excel 已退出/崩溃，重开工作簿再运行）、`EXCEL_BUSY`（被弹窗/计算占用，先关弹窗）。
  - VBE/宏：`VBE_PROTECTED`（工程加密 / AccessVBOM 未开，读不了 `VBProject`）、`ARG_MISMATCH`（`DISP_E_BADPARAMCOUNT`，`--arg` 个数/顺序与过程签名不符）、`RUN_FAILED`（宏不存在 / 被禁用 / 运行时错误等业务失败的统一收口）、`VBE_COMPILE_ERROR`（超时枚举到 VBE 编译错误框，`stage=vbe`）、`WATCHDOG_TIMEOUT`（宏运行硬超时 / 弹窗卡住；通用 VBE / Excel 模态框仍归此码，`hint` 点名弹窗标题与正文关键片段）。超时后目标 PID 已退出报 `EXCEL_CLOSED`（`stage=assert`，优先于超时码）；缺宏、禁用宏的细分识别仍统一收口 `RUN_FAILED`（B3 已接受设计差异，从 `hint` 文案区分）；旧码 `VBE_DENIED` / `MACRO_NOT_AVAILABLE` / `MACRO_DISABLED` / `VBE_DIALOG` / `MODAL_DIALOG` 在当前枚举中已不存在。
  - 写回/断言：`SOURCE_PREFLIGHT`（源码预检没过；硬项见 `pitfalls.md` §11、§13–15）、`POSTCHECK_FAILED`（写回回读不一致；已忽略过程级 Attribute 与 VBE 尾部 `()` 占位）、`FILE_READONLY`/`SAVE_FAILED`（只读/保存失败）、`ASSERT_NO_LOG`（跑完没有新日志）、`ASSERT_NO_TABLE`（工作簿没有 测试表，断言未执行）、`ASSERT_LOG_UNREADABLE`（日志被占用/读不了）、`ASSERT_FAILED`（日志读不出通过/失败计数）。
  - 初始化/兜底：`MISSING_LAB_TEMPLATE`（安装包缺 `scripts/make_lab_workbook.py`）、`LAB_TEMPLATE_FAILED`（生成测试表失败）、`CONFIG_NOT_FOUND`（找不到 `mapping.json`，`stage=config`）、`UNEXPECTED`（未归类异常，看 stderr 最后 20 行）。
- 失败对象另带 `retryable`；`hint` 里有下一步处置，`suggested_command` 只有在是可直接执行的命令时才填。
- 写类动词的预览标记不统一：`push` / `pull` / `pull-open` / `run` 看 `stats.dry=true`，`clean-backup` 看 `stats.applied=false`，`promote` 看 `details[0].dry`；不要只认 `stats.dry`。
- 只消费 stdout 最后一行；不要解析前半段日志。`--json` 模式只输出一行信封，不输出可读文本；警告并入 `stats.warnings`，正常路径不另走 stderr；异常 traceback 只进 stderr。

更多字段：`sync-commands.md`。
