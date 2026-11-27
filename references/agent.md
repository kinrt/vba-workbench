# vba-workbench\scripts\agent.py —— Agent 专用统一入口

不面向人：没有交互菜单、不提问、不等人回车。**只给 Agent 用**。
人用的入口只有两个：`audit.py`（体检）、`sync.py`（抓取/写回）。

本册只讲 `agent.py` 的动词、JSON 信封与风格检查，前置概念与入口分工见 `SKILL.md`。

## 目录

- [`--json`（全局开关）](#json全局开关)
- [风格检查 / 修复（`style` 子命令）](#风格检查--修复style-子命令)
- [补丁文件格式](#补丁文件格式)
- [两天临时脚本的收编对照（73 个）](#两天临时脚本的收编对照73-个)
- [三条规矩](#三条规矩)

```
python vba-workbench\scripts\agent.py probe [--project P] [--strict]   四本状态一览（磁盘/已打开/工程/模块数/测试表行数）
python vba-workbench\scripts\agent.py book  --action open|close --project P [--apply]   开/关工作簿
python vba-workbench\scripts\agent.py backup --project P [--tag 名] [--apply]            整工作簿备份到 .backup/<时间戳>-agent
python vba-workbench\scripts\agent.py patch --file 清单.md [--project P] [--no-style] [--apply]  精确修改源码
python vba-workbench\scripts\agent.py assert --project P [--macro 模块.过程] [--sheet 工作表名] [--clear] [--baseline 行数] [--timeout N] [--log 目录] [--apply]
python vba-workbench\scripts\agent.py push  --project P [--apply]                        直接复用写回逻辑（不启动 sync.py 子进程）
python vba-workbench\scripts\agent.py style --action check|fix [--which all|eol|call|explicit] [--scope src|--project P] [--apply]
python vba-workbench\scripts\agent.py logs  --project P [--log 目录]                     最新测试日志摘要
python vba-workbench\scripts\agent.py metrics [--out 报告.md] [--apply]                    过度设计体检（写报告要 --apply）
```

`assert` 的宏和测试工作表不再写死：优先级 `--macro/--sheet` > 项目键 `assert_macro/assert_sheet` > `defaults` 同名键 > 模板默认 `测试表格工具.运行全部断言` / `测试表`。非模板工程在 `mapping.json` 配一次即可；目标工作簿里找不到配置的宏时报 `ASSERT_MACRO_NOT_CONFIGURED`，提示改 `assert_macro` 或 `--macro`，不再只说“核对宏名”。

宏的调用顺序（`_run_assert_macro` 现状）：先按**零参** `Run(macro)` 调一次；仅当这一次报参数个数错误（`ARG_MISMATCH`）才退到两参 `(含性能=False, 静默=True)` 再试。第一次若是其它错误原样上抛，不再试探两参；第二次仍是参数个数不符，才报 `ASSERT_MACRO_ARGS`（`stage=assert`、`retryable=False`）；第二次是其它错误则原样上抛，不把第一次的 `ARG_MISMATCH` 当最终结论。断言宏要么声明为零参，要么接受 `(含性能 As Boolean, 静默 As Boolean)`。

**能力边界**：`assert` 不是通用宏执行器，只接**会写日志**的断言宏——运行完要在新日志里读到「通过/失败」计数；零参/两参的自测宏只写工作表、不写日志时，恒报 `ASSERT_NO_LOG`。这类宏请改用 `sync.py run --project 甲 --macro 模块.过程 --apply`（工作簿未打开时加 `--isolate`）查看工作表结果。

一条规矩：**`agent.py <动词> [--project P] [--json] [--apply]`**。写操作/执行类动词（book/backup/patch/assert/push/style fix，以及带 `--out` 的 metrics）默认**只预览**，必须 `--apply`；只读动词（probe/logs/metrics 不带 `--out`/style check）直接出结果。`--json` 全动词可用。

结论约定：最后一行是结论；`assert` 退出码 `0=全通过 / 1=业务失败或未预期异常（含断言失败、没有新日志、没有测试工作表、读不出计数等） / 2=用法错`。目标工作簿里找不到 `assert_macro` 报 `ASSERT_MACRO_NOT_CONFIGURED`，宏格式不是「模块.过程」报 `ASSERT_MACRO_INVALID`。`--timeout` 是**运行期硬超时**：宏在独立线程里执行（看门狗基础函数 `toolkit.watchdog.run_with_watchdog`），超时不再等待；超时后先按 PID 存活二次分类，认不出原因才回退 `WATCHDOG_TIMEOUT`；不会强行结束用户已打开的 Excel。`book --action open` 用 `new_excel(hidden=False, force_new=True)` 自己起可见实例，**不借用户会话**（避免改掉用户的 `DisplayAlerts/EnableEvents/AutomationSecurity`），成功后输出 `visible/hwnd/pid`；`book --action close` 关闭工作簿后若该实例已无其它工作簿，会 `Quit` 回收空实例，不再累积可见空窗口。项目名写错/没给项目名分别报 `PROJECT_NOT_FOUND` / `PROJECT_NOT_SPECIFIED`（`stage=config`），mapping 指不到 xlsm 报 `TARGET_NOT_FOUND`。切目标簿前先按目标 PID 探测前台模态框：同进程有 Excel MsgBox/InputBox 挡住时直接报 `TARGET_NOT_ACTIVE`/`stage=attach`，编译错误框报 `VBE_COMPILE_ERROR`/`stage=vbe`；没有弹窗再连续 5 次 `Activate` 兜底。`Application.Run` 因模块/过程不存在、宏被禁用或业务报错抛 `com_error` 时，1.6 统一收口 `RUN_FAILED`（`stage=assert`），不再落到 `UNEXPECTED`；缺宏与禁用宏不恢复独立分流码（B3 复核维持已接受设计差异，见 `backlog-1.7.md`），区分信息只从 `hint` 文案区分宏名/编译问题与信任位置。`probe` 在“Excel 开着但 COM 不接话”时按 `COM_CONTEXT` 失败返回，不再只写一行备注后报成功。弹窗识别族（VBE 编译错误框、通用 VBE 模态框、Excel MsgBox/InputBox 模态框）已由 `toolkit.dialogs` 回补：超时后先判目标 PID 是否退出——已退出报 `EXCEL_CLOSED`（`stage=assert`，优先于超时码）；仍存活再按 PID 枚举弹窗，编译错误框（`#32770` + 编译错误 / `#If`/`#Else`/`#End If` 关键字）报 `VBE_COMPILE_ERROR`（`stage=vbe`），通用 VBE / Excel 模态框仍归 `WATCHDOG_TIMEOUT`（`stage=assert`），`hint` 会点名弹窗标题与正文关键片段；旧 `VBE_DIALOG` / `MODAL_DIALOG` 细分码在当前错误码枚举中不存在，归属不明回退 `WATCHDOG_TIMEOUT`。Excel 进程退出/RPC 断开报 `EXCEL_CLOSED`，Excel 忙拒答报 `EXCEL_BUSY`，运行阶段统一 `stage=assert`。断言日志相关：`ASSERT_NO_LOG`（运行完没有新日志；不写日志的自测宏也落到此码，改用 `sync.py run`）、`ASSERT_LOG_UNREADABLE`（日志被占用/读不出）、`ASSERT_FAILED`（日志里读不出通过/失败计数）；`ASSERT_NO_TABLE` 表示工作簿没有 `测试表`。别的 Excel 进程开着同类弹窗不会误报，归属不明时回退 `WATCHDOG_TIMEOUT`。普通 `probe` 在“Excel 没开”时仍退出 0（`stats.excel=no_excel`）；做健康门禁时加 `--strict`，没开也按 `NO_EXCEL` 失败。

## `--json`（全局开关）

任意子命令后面加 `--json`，stdout 只留**一行** JSON：

```json
{"ok": true, "action": "agent.probe", "details": [{"project": "…", "file": "…", "disk": true,
 "open": "否", "vbe": "可读写", "modules": "12", "test_sheet": "有", "rows": 34}],
 "errors": [], "elapsed_ms": 812, "ts": "2026-11-21T19:00:00+08:00",
 "stats": {"projects": 5, "excel": "running", "lines": []}, "exit_code": 0}
```

错误项是对象：`{"code": "...", "stage": "...", "message": "...", "hint": "...", "suggested_command": "...", "retryable": false}`；异常也会输出这个信封并 `exit 1`，traceback 只进 stderr。

- `probe` 也是统一信封：项目行在 `details[]`，行数在 `stats.projects`；不再直接吐数组。
- 失败时 `ok=false`、`errors` 给出原因，`exit_code` 与进程退出码一致。
- 其余可读文本不写进 JSON 信封；只有 `probe` / `patch` 等少数动词会把可读摘要放进 `stats.lines`，不要把它当成完整报告。
- `agent.py push` 直接调用 `features.agent_io:push`，不启动 `sync.py` 子进程；信封 `action` 是 `agent.push`，错误码与退出码由本命令自己收口。

## 风格检查 / 修复（`style` 子命令）

只动 `src\` 下 `.bas/.cls/.frm` 的文本格式，不碰工作簿、不调 Excel。
规则出处：README 硬规则 1（UTF-8 无 BOM + LF）、硬规则 3（调用括号风格）。
实现在 `toolkit\style_rules.py`，入口只有 `agent.py style`（原独立脚本 `style.py` 已并入，不再有第二个入口）。

### 检查项

| 项 | 命令里的名字 | 规则 | 修复动作 |
|---|---|---|---|
| 编码/行尾 | `eol` | UTF-8 **无 BOM** + **LF** | 去 BOM、CRLF/裸 CR 压成 LF |
| Option Explicit | `explicit` | 业务模块（`.bas/.cls`，文档模块除外）必须有 `Option Explicit` | 只报不修 |
| 无参空括号 | `call` | 无参调用保持 `Call Foo`，不许 `Call Foo()` | 去掉空括号 |
| 裸参数 | `call` | 带参调用一律 `Call Foo(a, b)`，不许裸参数 | 补 `Call` + 括号 |

裸参数修复的细节约定：

- `MsgBox "a", 1` / `Run "x"` / `obj.M a` / 行内 `If x Then Foo a` → 全部转 `Call ...(参数)`
- `Foo (x)` 外层双括号合并为 `Call Foo(x)`
- 命名参数保留：`Foo Field:=1` → `Call Foo(Field:=1)`（`:=` 不会被当成语句分隔符）
- 顶层冒号/`Else` 后的语句不吞进参数：`a.Add k, v: i = i + 1` → `Call a.Add(k, v): i = i + 1`
- 续行 `_`（空格+下划线结尾才算）：首行加 `Call Foo(`，末行补 `)`
- 行尾注释留在括号外：`Call 注册快捷键("^+C", "复制")   ' 注释`
- 声明行 `Public Sub Foo()`、`Attribute VB_*`、`.frm` 控件块、纯注释行不动

### 用法（唯一入口 `vba-workbench\scripts\agent.py style`，无菜单）

```
python vba-workbench\scripts\agent.py style --action check                     :: 只读全查（退出码非 0 = 有问题，可当提交前门禁）
python vba-workbench\scripts\agent.py style --action check --which eol         :: 只查编码/行尾
python vba-workbench\scripts\agent.py style --action check --which call        :: 只查调用括号
python vba-workbench\scripts\agent.py style --action fix                       :: 预览全部修复（不写入文件）
python vba-workbench\scripts\agent.py style --action fix --which call --apply  :: 写入文件：调用括号
python vba-workbench\scripts\agent.py style --action fix --which eol --apply   :: 只写入文件：行尾
python vba-workbench\scripts\agent.py style --action check --scope src\common   :: 限定范围（相对仓库根或绝对路径，默认 src）
python vba-workbench\scripts\agent.py style --action check --project 示例表A   :: 项目 src + 公共库，与 push 源码预检同规则
```

调用逻辑：**先 `check`，再 `fix` 看预览，确认无误才 `--apply`**。
`fix` 不带 `--apply` 永远不写入文件；`--apply` 写入文件后自动对同一批文件复检，有残留返回非 0。
`check --json` 的 `details[]` 逐条给 `rule / file / line / old / new / message`（`call` 类带行号与改法，`eol/explicit` 带文件级说明）；失败 `suggested` 保留本次 `--which`，带项目时保留 `--project`，自定义范围保留 `--scope`。
重复运行安全（幂等）：修完再 `check` 应全绿。

`agent.py patch --apply` 写入源码后会自动运行一遍 `style --action fix --apply` + `style --action check`，所以手工改源码不必再单独运行。

### 识别规则与边界（启发式）

- "裸参数"的判定依据 = **全库声明过的 Sub/Function 名**（`toolkit.vba_parse.declared_procs`）
  + 内置 `MsgBox / InputBox / Run` + `a.b.c` 点号方法形态。
- 语言语句不支持 Call 括号形式，**永不转换**：`Load / Unload / Debug.Print / ChDir /
  ChDrive / FileCopy / Kill / Name / MkDir / RmDir / SavePicture / SendKeys / Beep /
  Open / Close / Print # / Put / Get` 等。
- 赋值与表达式不动：首个"参数"是 `=` 或运算符（`Is / Like / Mod / And …`）即跳过，
  所以 `Then 单击待决 = mPending.Exists(slot)` 这类函数返回值赋值不会被误改。
- 已知不覆盖（刻意保守，遇到手工改）：
  - 冒号分隔的**后续**语句（`i = 1: Foo 1, 2` 里第二个调用不查，只查行首语句）；
  - `Else` 分支里的裸调用（`If x Then Call A(a) Else B 1` 里 `B 1` 不转）；
  - 与点号方法同名的变量赋值误伤已由运算符首参检查挡住，但极冷门形态仍建议看预览。

### 来源

由 2026-11 括号风格统一战役的临时脚本沉淀合并：
`check_eol.py` / `fix_lf.py`（行尾）、`verify_parens.py`（空括号残留）、
`call_style.py`（裸参数审计+转换，267 处实战验证）。
一次性纠错脚本（unify/fix_bare/revert 三个）无复用价值，未收录。

## 补丁文件格式

```
@@ src/common/模块/自定义通用函数.bas
-旧的一行
-旧的第二行
+新的一行
```

- `@@ ` 起一个新文件；同一文件可以有多组 `-`/`+`（连续 `-` = 旧块，连续 `+` = 新块）。
- 每组旧文本在全文件里**必须恰好命中 1 次**；命中 0 次或多次，整批一律不写入文件（不会半途替换）。
- **命中是"整行对齐"比对**，不是子串比对：缩进不同的两行不会被算成同一处（2026-11-20 之前用 `str.count`，24 空格的行会被 32 空格的行顺带命中，已改）。
- 源码行本身以 `-`/`+` 开头时，前面加一个反斜杠转义。
- 加 `--apply` 时：写完立即 `style --action fix --apply` + `style --action check`，不通过返回 1；不加只预览。

## 两天临时脚本的收编对照（73 个）

`probe_*.py` 72 个 + `t41.py`，全是一次性脚本，已全部删除；能力归并到上表。

| 类别 | 数量 | 代表脚本 | 现在怎么用 |
|---|---|---|---|
| A 项目/环境探测 | 16 | `probe_books`、`probe_state`、`probe_status`、`probe_mods`、`probe_vbp`、`probe_vp`、`probe_pw`、`probe_ver`、`probe_cfg`、`probe_datasrc`、`probe_yunying(2)`、`probe_zu`、`probe_other`、`probe_after_crash`、`probe_after_merge` | `agent.py probe`（四个工作簿汇总，含工程密码保护/模块数/测试工作表行数） |
| B 表结构/区段 | 9 | `probe_size`、`probe_rowcount`、`probe_range`、`probe_lo`、`probe_headers`、`probe_grid`、`probe_inittable`、`probe_merged`、`probe_cls_region` | `probe` 已覆盖行数/工作表对象；其余是当年的一次性结论，需要时用 `book` + `probe` |
| C 排序/切片器验证 | 19 | `probe_sort`、`probe_sort2`~`14`、`probe_slicer1`~`5` | 已定论并落进代码（`排序/筛选` 相关），不再需要脚本 |
| D 界面/对话框/快捷键 | 11 | `probe_dlg`、`probe_faceids`、`probe_findface`、`probe_guard`、`probe_ontime`、`probe_gettl`、`probe_stream`、`probe_decompress`、`probe_cross_origin`、`probe_ref`、`probe_set` | 已定论（FaceID/OnTime/GetTickCount 用法已进 `自定义通用函数`），不再需要脚本 |
| E 写回/修复/核对 | 16 | `probe_fix1`~`4`、`probe_rebuild1`、`probe_verify2`~`4`、`probe_verify_merge`、`probe_fieldcheck(2)`、`probe_check(2/3)`、`probe_cands`、`t41.py` | `push`（写回）+ `assert`（核对）+ `patch`（改源码） |
| F 运行断言 | 2 | `probe_run`、`probe_runtest` | `agent.py assert` |

原来手工写在别处的 `t45_edit / t45_test / t45_run / t46_edit / t46_run / probe_books` 也一并归到
`patch` / `assert` / `probe`，不再存在"每次现写一个 py"。

## 三条规矩

1. 查代码一律用 Grep 工具，**不许**写 py 去 grep。
2. 改源码一律 `agent.py patch --apply`，不许现场手写替换脚本。
3. 运行断言一律 `agent.py assert --apply`，写回一律 `agent.py push --apply`；任务收工时 workspace 里 `.py` 数量必须为 0。
