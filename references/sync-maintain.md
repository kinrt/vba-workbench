# sync.py 维护册：1.6 架构地图、硬规则、改法与取舍

> 版本：1.6.2（1.6 分层架构）；**本册为 1.6 重写**。
> 旧版基于单体 `sync.py` 的函数地图（`cmd_*` / `MENU_ITEMS` / `_dispatch` / `align_with_common` / `sync_into` 等）已随分层重构失效，不再保留；仍成立的运维知识收敛在 §6 硬规则与 §8 取舍。前置概念见 `SKILL.md`，抓取/写回原理见 `sync.md`，公共基础函数边界见 `toolkit.md`。

## 目录

- §1 分层地图（入口壳 / features 切片 / toolkit 基础函数 / adapters 网关）
- §2 命令登记与派发链路
- §3 如何加一个动词
- §4 如何改规则
- §5 测试怎么跑
- §6 硬规则（冻结区）
- §7 已知限制
- §8 关键取舍

## 1. 分层地图

### 1.1 入口壳：`scripts/sync.py`、`audit.py`、`agent.py`

- 入口文件不含任何动词特例；固定流程：`specs_for(entry)` → `build_parser` → `AppContext.build` → `dispatch` → `render`。
- 壳层职责：`--version`、`-i` 占位、`--no-input`、Failure / KeyboardInterrupt / 兜底异常翻译、JSON 信封渲染；三个入口互不 import。
- `-i` 交互菜单本体（`features/sync_interactive.py`）：`cli.ensure_interactive()` 先把关终端，非 TTY 直接给可读提示并退出 1；菜单层列项目 → 选项目 / 全局工具 / 新建项目 → 选动作；只读动作直接运行，写操作自动预览、回车执行，动作结束回菜单。执行交既有 handler，菜单层不碰 COM、不写入文件。

### 1.2 features 切片：`scripts/features/`

按命令族分切片，每个切片只声明自己的 `SPECS`（`CommandSpec`）与 handler；handler 签名 `fn(ctx: AppContext) -> Outcome`。

| 切片 | 动词 | 职责 |
|---|---|---|
| `core` | version / list / doctor / scrub | 只读：版本核对、项目解析、环境体检、交付件脱敏扫描，不碰 Excel、不写入文件 |
| `sync_read` | vbamod / plan / pull / pull-open | 只读与预览类：模块清单核对、模式探测、抓取源码 |
| `sync_push` | init / push / shells | 写操作路径：骨架初始化、写回工作簿、占位壳重建 |
| `sync_transform` | align / promote / gbk-export / clean-backup | 公共库换壳、模块提升、GBK 副本、备份清理 |
| `run` | run / watch | 宏运行编排 + 常驻监听；`MacroRunner` 的调用方 |
| `agent_io` | probe / book / logs / backup / push | agent 状态一览、开关工作簿、日志摘要、整工作簿备份、内部转调 push |
| `agent_patch` | patch / assert / style / metrics | 补丁写入文件、断言核对、风格检查、过度设计体检 |
| `audit_a` | members / common / project | 成员引用、公共库总览、项目依赖面 |
| `audit_b` | structure / indicators / overview | 结构问题、指标体检、全仓总览 |
| `audit_engine` | — | audit 三个视角共用的规则引擎（数据与判定） |
| `release` | — | 文档生成、基线、差分、安装与发布（`SPECS` 为空，无运行时命令） |

- 新动词优先落到现有切片；切片内私有辅助函数用 `_` 前缀；跨切片复用的逻辑下沉 toolkit，不在切片之间互相 import。

### 1.3 toolkit 公共基础函数：`scripts/toolkit/`

| 子模块 | 定位 |
|---|---|
| `contract` | 跨命令共享数据形状：`ErrorCode` / `Stage` / `MutationMode` / `Outcome` / `Failure`；新增码先登记再使用 |
| `registry` | manifest 形状（`CommandSpec` / `ArgSpec` / `ProjectMode`）与 argparse 生成器（`build_parser` / `dispatch` / `load_handler`） |
| `commands` | 按 `FEATURE_MODULES` 汇总切片 `SPECS`，提供 `specs_for(entry)`；重复登记直接报错 |
| `cli` | 入口骨架：`-i` / `--no-input` 注册、`ensure_interactive`、`run_cli` 异常翻译、`entry` 收口 |
| `context` | `AppContext`（配置、路径基准、JSON 模式、阶段计时）与 `Progress`；长任务计时 / 进度只写 stderr |
| `config` | `mapping.json` 加载、本机覆盖层合并、配置读写 |
| `repo` | 源码仓库基础函数：版本解析、模块索引、占位壳判定 / 生成、模块名与文件名换算、源码文件收集、噪声副本跳过 |
| `files` | 文本 / 备份基础函数：UTF-8 无 BOM + LF 写入、GBK 转码副本、备份目录与清理、日志尾行、临时目录清理 |
| `excel` | COM 会话：起隐藏 / 可见实例、附着已开 Excel、开关工作簿、组件查找 / 移除 / 导入 / 替换、VBE 权限探测、引用只读读取 |
| `session` / `state` | `WorkbookSession` 会话形状；`TargetState` / `Purpose` 与状态规则 |
| `macro` | `MacroRunner` 唯一宏执行 / 分类入口：活动簿核对、guard 快照 / 还原、`Application.Run`、错误分类、日志收尾；sync.run 与 agent.assert 共用 |
| `watch` | 常驻监听：快照、比对、基线刷新、监听日志 |
| `check` | 写回前预检与模块集合 / 文档模块对账 |
| `vba_parse` / `vbabin` | 源码过程声明抽取、注释 / 字符串剥离；`vbaProject.bin` 模块名读取 |
| `style_rules` | 风格检查规则：编码 / 行尾、Option Explicit、无参空括号、裸参数 |
| `ui` | 交互基础函数：`read_line` / `choose` / `ask_yes` / `pick_list` / `pick_list_default` / `pick_grouped` / `run_menu` 等；菜单本体由 `features/sync_interactive.py` 接入（B6 闭环），`ui` 提供骨架基础函数 |
| `render` | `Outcome` → 信封 / 可读文本渲染 |
| `base` / `app_guard` | `Failure` 与可读文本输出；Excel 状态（ScreenUpdating / Calculation / EnableEvents）快照与还原 |

### 1.4 adapters 网关：`scripts/adapters/excel_gateway.py`

- COM 失败分类（忙 / 断开 / 上下文）；features 与 macro 不另写第二套 `Application.Run` 分类器。

## 2. 命令登记与派发链路

1. 切片 `SPECS` 登记 `CommandSpec`：`verb` / `entry` / `handler`（`"模块:函数"`）/ `mutation` / `project` / `args` / `supports_json` / `exit_codes` / `summary` / `ref` / `tests`。
2. `commands.load_specs` 按 `FEATURE_MODULES` 汇总；`entry.verb` 重复登记直接报错。
3. `registry.build_parser` 从 manifest 生成 argparse；公共参数（`--json` / `--apply` / `--project` / `--all-project` / `--no-input`）由骨架统一附加。
4. `dispatch` → `load_handler` 按字符串导入函数 → 调 `handler(ctx)`，返回 `Outcome`。
5. `render` 按 `json_mode` 出一行 JSON 信封或可读文本。

- manifest 是命令面唯一登记源；`references/commands.md` 由 `tests/tools/gen_commands_md.py` 从 manifest 生成，不要手改。

## 3. 如何加一个动词

1. 按 §1.2 选切片，写 handler：`fn(ctx) -> Outcome`；参数从 `ctx.args` 取，失败抛 `Failure(code, stage, hint=...)`，成功返回 `Outcome`。
2. 在该切片 `SPECS` 加 `CommandSpec`：动词、入口、handler 字符串、变更模式（`MutationMode`）、项目模式（`ProjectMode`）、业务参数 `ArgSpec`、退出码、`summary` / `ref` / `tests`。
3. 公共参数由骨架统一附加，不要在 `ArgSpec` 重复登记。
4. 重新生成命令表：`python tests/tools/gen_commands_md.py`（输出 `references/commands.md`）。
5. 在对应命令册（`sync-commands.md` / `audit.md` / `agent.md`）补 `details[]` 字段与用法说明。
6. 补定向测试（`tests/test_*.py`）；未知动词由 argparse 退出 2，不会被当成项目名。

涉及工作簿的新命令不要自己写附着 / 备份：优先复用 `excel` / `session` / `macro` / `files` 基础函数；列源码文件走 `repo.source_files`（已跳过占位壳与噪声副本），不自己 `rglob` 判后缀。

## 4. 如何改规则

- **风格规则**（编码 / 行尾 / Option Explicit / 调用括号）：`toolkit/style_rules.py`；检查项与修复动作对照见 `agent.md` 风格检查一节。
- **体检指标**（members / common / project / structure / indicators / overview）：规则数据在 `features/audit_engine.py`，视角编排与报告在 `audit_a.py` / `audit_b.py`；数据层只产数据，排版层只出报告。
- **写回前预检**（硬项 / 软提醒、文档模块对账、模块集合核对）：`toolkit/check.py`。
- **占位壳判定与生成**：`toolkit/repo.py` 的 `is_shell_file` / `shell_text` / `SHELL_HINT`；生成与识别是一对，改提示文字必须同步改判据。
- **公共库 auto 子集判定**（模块名 / 过程名 / 闭包 / 占位壳四条依据）：`core` / `sync_read` / `run` 内的引用闭包辅助；证据先剥注释、再抹字符串字面量，标识符边界含中文。
- **引用设置检测**（只读、只警告）：`features/sync_push.py` / `sync_transform.py` 的 `_REF_RULES` / `_EARLY_BOUND_RE` / `_scan_required_refs`；脚本不改 VBE References。
- 放宽任何规则前先把反例写成单测；预检硬项只增不减。

## 5. 测试怎么跑

- 定向：在 `vba-workbench/` 下 `python -m unittest tests.test_xxx`；本批门禁为 `tests.test_gen_commands_md`、`tests.test_version_consistency`。
- toolkit / features 单测一律用假 session / 桩对象，不 new COM、不真启动 Excel；COM 相关只做离线规则断言。
- 语法自检：`python -m compileall -q scripts tests`。
- 长任务（批量读写 / 大循环）handler 内走 `ctx.progress.stage` 做阶段计时；进度写 stderr；可读文本在非 JSON 模式走 stdout，JSON 模式不输出可读文本，避免污染单行信封。
- 全量 unittest 不在本侧跑；写回 / 宏类行为走 `tests/MANUAL.md` 手工清单。

## 6. 硬规则（冻结区，改代码前必须知道）

1. **`.bas/.cls/.frm` 按 UTF-8（无 BOM）+ LF 存储**；GBK 只出现在 VBE 交换边界（pull 转 UTF-8、push 转临时 GBK、`gbk-export` 出 GBK 副本），别手工把 src 转成 GBK。
2. **`VBComponents.Export` 不补扩展名**，必须按 `comp.Type` 自己拼（`repo.EXT_FOR_TYPE`）。
3. **`.frx` 必须与 `.frm` 同目录同名**；移动 / 抓取一个窗体时两个一起处理。
4. **文档模块不能 `Remove`**，只能 `CodeModule.DeleteLines` + `AddFromString`，且要剥 `Attribute VB_` 行。
5. **公共库模块不能在项目目录里留第二份正文**——只留占位壳，否则必然分叉。
6. **占位壳的识别与生成要同时改**：`shell_text` 的文字和 `is_shell_file` 的判据是一对（`SHELL_HINT` 常量串起两边）。
7. **写回顺序是 common → 项目**，同名会被项目覆盖；公共库与项目同名是已知风险（尚无前缀方案）。
8. **脚本不改 VBE 引用**；引用检测只读、只警告。
9. `AutomationSecurity=3` 会禁用宏导致 `Run` 失败；探针与运行都用 `1`。
10. **换壳（`align`）绝不新增文件**：只把已有正文覆写成壳；写回也只刷新已存在的壳。壳的新建只允许在三处显式路径——promote 搬走后留壳、`shells` 命令、抓取时为公共库模块留占位。
11. **auto 判定只认证据、不认黑名单**：四条依据见 `sync.md`；匹配前必须剥注释，标识符边界必须用 `\w`（含中文），占位壳只当文件名证据、不读其正文。名单式排除属于把误判藏起来，不做。
12. **文件名不等于模块名**：文档模块源码文件名带工作表名。凡是拿文件名去比模块名的地方（过滤、查找组件、去重、清理旧文件、基线子集、宏名重名判定），一律 `repo.module_name_of(path)`，禁止 `f.stem` 直接比。
13. **写回绝不新建文档模块**：工作表被删除后它的模块也没了，残留 `SheetN.bas` 只 warn 跳过；想清掉就跑一次抓取。写回结束后再核对一次模块集合是否原样。
14. **预检硬项只增不减**：`.cls` 缺 `VERSION 1.0 CLASS`、单语句续行超过 24 行必须在预检里报 `SOURCE_PREFLIGHT`（带文件 + 行号）；类模块 `Me.` 引用私有字段只进软提醒，不进硬错误。
15. **规则放行要有边界**：`manual_accel` 只对「同一过程 + 有 `On Error` + 同一属性后面有还原」放行；放宽规则时先把反例写成单测。

## 7. 已知限制 / 没做的事

- 只支持 Windows 微软 Excel；WPS、Mac、网页版一律不行。
- 只处理 VBA 工程；工作表数据、条件格式、Power Query 不在同步范围（xlsm 本身不纳入版本库，只有源码进 Git）。
- 文档模块默认不写回（`import_document_modules: false`）。
- 工程受 Excel 密码保护时，脚本读不到 `VBProject`，只能人工先解除密码保护。
- 公共库与项目同名模块的冲突尚无自动方案（写回顺序决定谁覆盖谁）。
- `AutomationSecurity` 会把宏临时禁用，依赖 `Workbook_Open` 的流程在脚本里不会触发（这是故意的）。
- 自动写回（`watch`）只处理**模块级**增量：改名等于"删一个 + 建一个"，脚本只会 Import 新名字，不会自动删旧名字模块，需人工在 VBE 里删；`.frx` 缺失、源码含 GBK 编不出的字符时该模块写回失败（打印后继续监听，其余模块不受影响）。
- 自动写回不重算引用依赖，也不做写回前的公共模块子集确认面板（子集固定取 `auto` 结果）；一个进程盯所有当前已打开的已登记工作簿（`watch <项目名>` 才只监听一个）。
- 启动核对是**逐行文本比对**，不是语法比对：VBE 里"等价但写法不同"（例如 `Rem` 与 `'`）会被算成「内容不同」，需要人决定规则。
- 「工程里没有」在 `--startup src` 下会被 Import **新建模块**；在 `skip` 下保持缺失，之后第一次改这个文件才会建。
- 非交互运行（后台、管道、计划任务）遇到差异不会问，一律按"先不动"继续，差异清单只出现在输出与日志里。
- 「以工作簿内为准」会把公共库模块也抓回 `src/common`（这份公共代码被所有工作簿共享）；只想改本项目就只对项目模块用这条规则，公共库模块建议走 `pull` 后人工 diff。
- 启动核对只覆盖"这个工作簿用到的"文件：项目 `src` 全部 + 公共库 `auto` 子集；子集外的公共模块两边不一致不会报。
- 多工作簿模式的目标集合在**启动那一刻**确定：启动后才打开的工作簿不会自动纳入；中途关闭一个就剔出监听，全部关闭才退出。
- 每轮探活分三档，**"这一轮读不到"不等于"关了"**：工作簿对象名字读得到、或 `Application.Workbooks` 列得到 → 继续盯；列得到但没有这个名字 → 判真关闭；连 `Workbooks` 都读不到（重算 / 存盘 / 弹窗挡住、COM 那一刻拒答）→ 记一行"本轮跳过"，照样继续监听，不剔除目标。
- 多工作簿模式备份是**每个工作簿各一份**；`--no-backup` 可全关。
- 同一份 `src/common` 改动会写进所有已打开且用到它的工作簿：每个工作簿的独立性靠自己的 `auto` 子集，子集里没有的模块会打印"跳过"而不是硬写。

## 8. 关键取舍（为什么这么设计）

| 决策 | 原因 |
|---|---|
| 源码只放 `src/`，xlsm 不纳入版本库 | xlsm 是二进制，diff 无意义且体积大 |
| 公共模块在项目里只留一行注释壳 | 既能在文件树里"看得见"引用关系，又不会出现两份可编辑正文 |
| 壳放在原分类目录、同名同后缀 | 与 VBE 里看到的层级一致，打开就知道"要去公共库改" |
| 公共库按引用选择性写回（`auto`） | 不是每个工作簿都用公共工具类，全量写回会带进不必要的引用依赖（如 ADODB） |
| 引用设置不自动改 | 改 `References` 风险高、且与工程密码保护策略相关；脚本只警告 |
| 菜单统一 `0` 退出本层；写操作预览后回车执行 | 早期 `q`/`0`/回车三套语义混用、写回还有 `y/N` 二次确认，误操作与重复确认都多；现在只保留“预览 → 回车执行 / 0 返回” |
| 批量写回保留逐字「确认」 | 一次修改多个工作簿，多一道确认换取可回滚 |
| `promote` 只搬文件不写正文 / 不改引用 | 保持"移动 + 建壳"这一最小动作，复杂合并留给人在编辑器里做 |
