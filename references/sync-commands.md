# sync.py 命令册：环境、菜单、命令行、排错

> 本册只讲环境、菜单、命令行的用法和排错；前置概念与原理见 `SKILL.md`、`sync.md`。四册各自从 §1 编号，不要求按册顺序读；需要选册看 `sync.md` 的分册索引。

## 目录

- §1 运行环境
- §2 快速使用（2.1 交互菜单 / 2.2 交互约定 / 2.3 命令行 / 2.4 退出码与 JSON）
- §3 排错
- §4 改动后的自检

## 1. 运行环境

| 项 | 要求 |
|---|---|
| 系统 | Windows + 微软原版 Excel（WPS、Mac 版、网页版都不行） |
| Python | 3.10+（本机 3.13），需 `pywin32` |
| Excel 设置 | 必须勾选「信任对 VBA 工程对象模型的访问」（注册表 `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1`） |
| 工程 | 不能加密；工程被锁时脚本读不到 `VBProject`，会直接报错 |
| 路径 | 代码根 = `sync.py` 的上一级目录；`mapping.json` 放在**代码根**（与 `vba-workbench/` 平级），配置里的相对路径按 workspace 解析（workspace 缺省 = 配置文件所在目录） |

依赖检查：`python vba-workbench\scripts\sync.py doctor`。

---

---

## 2. 快速使用

### 2.1 交互菜单

> **现行事实**：`-i` 进入交互菜单：先列 `mapping.json` 里的项目，项目列表后面固定跟“全局工具”和
> “＋ 新建项目（拖入 .xlsm）”。选项目后列动作（既有登记子命令）；只读动作选中即运行，写操作选中后
> **自动先预览**，预览通过后再回车才给 handler 加 `--apply`，输入 `0` 返回上一层。每次动作结束都回到
> 当前菜单，可连续操作，不退出。菜单层只做选择与参数装配，**不碰 COM、不修改文件**。
> `watch` 属于常驻动作，会先问监听秒数（默认 600），再回车开始监听。
> 新建项目直接把 `.xlsm` 拖进窗口（或回车打开文件选择框）：文件名带 ` V<数字>` 时自动取项目名和前缀、
> 按目录取最高版本；没有版本号时按该文件精确登记；建完自动抓取一次源码并进入动作菜单。
> **非 TTY / 非交互环境**（管道、`--no-input`）仍按旧规则给可读提示并退出 1，不会卡在 `input()`。
> 需要额外必填参数的动作（`promote --module`、`run --macro`）不进菜单，请直接用命令行子命令。

双击 `vba-workbench\scripts\启动菜单.bat`（它已带 `-i`），或显式：

```bat
python vba-workbench\scripts\sync.py -i
```

直接运行 `sync.py` 只打印帮助并退出 2，**不再因为有没有终端而改变行为**。当前菜单动作映射到下列既有子命令
（项目级：`pull` / `pull-open` / `plan` / `push` / `align` / `gbk-export` / `vbamod` / `watch`；
项目列表里的“＋ 新建项目”对应 `init`；
“全局工具”：`list` / `version` / `doctor` / `shells` / `clean-backup`）；下表 1.5.6 菜单 1-13 为历史对照
（排序：**VBE→src 在前、src→VBE 在后；单个工作簿在前、批量居中、其他在后**）：

**方向**：**VBE→`src/`** 抓取（把工作簿里的代码抓下来，菜单 1、3）；**`src/`→VBE** 写回（把源码写进工作簿，菜单 **2、4、10、12**；
其中 10 = 自动写回（增量监听）、12 = 可选先写回再运行宏）。
用占位壳替换同名旧副本（菜单 8）与提升为公共模块（菜单 9）都**只改 `src/`**，不碰工作簿，不属于这两个方向；
菜单 5、6、7、11 是**只读**（清单 / 版本 / 体检 / 读 vbaProject.bin 核对模块），不改任何一边。

| 功能（原编号） | 方向 | 用在什么时候 |
|---|---|---|
| 抓取源码（1） | VBE→src | 清单**永远列全部已登记项目**（带 `[已打开]/[可用]/[不可用 原因]` + 路径），**回车 = 抓取全部已打开的工作簿**、填编号 = 抓取那一个（`_pick_projects(extras=True)` 顺带列出“打开但未登记”的工作簿，选中后按文件名抓取并可当场登记）。多个工作簿逐个抓取，单个失败不中断，最后打「成功 X / 失败 Y」 |
| 写回源码（2） | src→VBE | 同一个清单、同一条规则：**回车 = 写回全部已打开的工作簿（多个逐个写）**、填编号 = 只写那一个。写进工作簿时按**全路径**定位并切到前台，绝不写进别的工作簿（默认写回后保存；多个工作簿连续写不再逐个弹确认面板，结尾统一汇总；某个工程受密码保护会当场提示，解除密码保护后按回车继续） |
| 批量抓源码（3） | VBE→src | 多个工作簿·磁盘：工作簿未打开也行 → 脚本自己开一个隐藏 Excel 逐个抓取；同一个清单，**回车 = 全部项目**，填编号 = 只抓取一个 |
| 批量写回（4） | src→VBE | 多个工作簿：公共库改完 → **逐个预检 → 备份 → 写回 → 回读** → 汇总；预检失败的工作簿跳过，不备份、不打开工作簿 |
| 查看项目清单（5） | — | 确认每个工作簿解析到哪个 xlsm、源码目录在哪、公共模块选了哪些 |
| 校验公共库版本（6） | — | 见各工作簿锁定的 `common_version` 与公共库实际版本是否一致 |
| 环境体检（7） | — | pywin32 / AccessVBOM / 项目路径是否正常 |
| 对齐公共库（8） | src→src | 先选一个工作簿 → 列出该项目里与公共库**同名**的真实代码（两边行数/修改时间）→ 逐个 `y/N` 确认是否换成一行的占位壳；换壳前原文备份到 `.backup/<时间戳>-换壳/<项目名>/`。**只改已有文件，不新建任何文件** |
| 提升为公共模块（9） | src→src | 把某个工作簿的模块搬进 `src/common`，并在原位留占位壳 |
| 自动写回（10） | src→VBE | 多个工作簿·常驻：进入后先列**当前打开着的已登记工作簿**，**回车 = 全部监听**、填编号 = 只监听那一个（`watch_open_list`）；监听公共库 + 各项目 `src`，**改哪个模块就只把那个模块**写进它所属的工作簿（`import_folder(only={模块名})`）；公共库模块改动写进所有已打开且用到它的工作簿（各工作簿按自己的 `auto` 子集判断）。**启动先核对一遍**两边（项目 src 全部 + 公共库 `auto` 子集），有一致 / 仅大小写空白 / 内容不同 / 工程里没有四档结论，冲突就列清单让你决定处理规则：以源码为准（写回工作簿）/ 以工作簿内为准（抓回 src）/ 逐个决定 / 先不动。**整工作簿备份只在启动做一次**（`.backup/<时间戳>-自动写回/`，多个工作簿则各备一份），之后每次写回都记日志（`.backup/watch/<项目名>-<时间戳>.log`）。改动要**稳定一个轮询周期**才动手（防编辑器写一半）；源码被删只提示、不删工程里的模块；关闭一个工作簿就不再监听它，全部关闭才停止；**只改内存**（菜单不再问保存，Ctrl+S 才写入文件），命令行默认写回后保存，`--no-save` 才只改内存；`Ctrl+C` 停止并打印「写回 X / 跳过失败 Y / 运行 Z 秒」+ 日志路径 |
| 核对模块清单与正文（11） | 只读 | 读 `xlsm` 里的 `vbaProject.bin`（`toolkit.vbabin`），**不用打开工作簿、不用工程密码**：列出这个工作簿的全部模块名，并自动和 `.backup` 里**最近一份备份**比对模块集合与模块正文，打印「模块 M→N、编译缓存 a→b、消失/新增清单、正文变化模块」。写回 / 抓取之后自查“有没有模块被删、哪个模块代码变了”就按这个。选工作簿走统一选择器，**回车 = 全部项目逐个核对** |
| 运行宏（12） | src→VBE(可选)→运行 | 选一个**已打开**的工作簿 → 输宏名、参数 → 问“先写回源码再运行？” → 调 `Application.Run` → 打印返回值 + **这次新写入的日志尾部**。只运行宏**不需要解除 VBE 密码保护**（要“先写回”才需要）。宏名与模块重名时自动写成 `模块.过程`（否则 VBA 会把它解析成模块、报“该宏在此工作簿中不可用”）。默认只改内存、不保存；`run --isolate --save` 才写入文件；非隔离带 `--save` 会明确提示未生效，journal 记 `save_requested/save_ignored`。`--isolate` 因工作簿已打开失败时 journal 仍记请求模式 `isolate`。隔离预检只放行 `NO_EXCEL` / `TARGET_NOT_OPEN`；目标实例 `COM_CONTEXT`（弹窗 / 受保护视图 / 异权会话）时隔离模式也不启动，hint 给三类成因与下一步，suggested 为可复制的 `run --project … --macro … --isolate --apply`。 |
| 清理备份（13） | 删 `.backup` 旧批 | 先算一遍清单（批次数、体积、按来源分布、将要删哪几批、能释放多少）**只预览不动文件**，再问 `y` 才真删。清理规则与自动清理相同：每类留最近 `backup_keep` 批 + 每天最新 `backup_keep_daily` 份 + 全局最新 `backup_min_keep` 批；`.backup/GBK`、`.backup/watch` 不参与 |

**清单标签**：`[已打开]` = 这个工作簿正开在 Excel 里（抓取/写回都不用另开实例，能带未保存改动）；`[可用]` = 文件在、但未打开；`[不可用] 原因` = 配置解析不到文件（路径错/文件搬走）。`[当前打开·未登记]` 只出现在菜单 1 的清单里，表示“已打开但没登记进 `mapping.json`”。标签靠 `watch_open_list` 现读，Excel 读不到时一律退成 `[可用]`，不会误报已打开。菜单 4 里还打开着的工作簿标 `[已打开·要先关]` 并单独警告——那一步会自己打开文件写入并保存，打开着的文件会被占用挡掉。

**一个选择器管所有选工作簿（`_pick_projects`）**：菜单 1、2、3、8、9 全部走同一个函数，清单**永远列全部已登记项目**（不是只列打开的那几个——看得见哪些未打开、哪些路径失效），每行附标签与路径；差别只在“回车算什么”，用参数表达：

| 功能（1.5.6 历史对照） | 调用 | 回车 | 填编号 | 一个都没打开时 |
|---|---|---|---|---|
| 1 抓取源码 | `_pick_projects(cfg, …, extras=True)` | 抓取**全部已打开**的工作簿 | 那一个（或“未登记”那一个） | 提示“现在没有已打开的工作簿：请填编号”，重新等输入 |
| 2 写回源码 | `_pick_projects(cfg, …)` | 写回**全部已打开**的工作簿 | 那一个 | 同上 |
| 3 批量抓源码 | `_pick_projects(cfg, …, enter="all")` | 全部项目（磁盘逐个抓取） | 那一个 | 不适用（磁盘任务） |
| 8 对齐公共库 | `_pick_projects(cfg, …, multi=False)` | 打开着**恰好 1 个**时用那个 | 那一个 | 提示后等编号 |
| 9 提升为公共模块 | `_pick_projects(cfg, …, multi=False)` | 同 8 | 那一个 | 同 8 |

`multi=False`（8、9）碰上“打开着多个”时回车**不算数**、提示填编号——只改一个源码的功能不能瞎猜是哪个。菜单 10 自动写回仍用 `watch_open_list` 那一套（它的清单本来就只该是打开着的），回车 = 全部监听。

仍需先选工作簿的三项，原因不是懒：4 批量写回会逐个判断（已打开借会话、未打开自己开隐藏实例）；8 对齐公共库、9 提升为公共模块都是 **src→src**，跟“打开了哪些工作簿”无关，要选的是模块来源/归属（9 的清单标题已写明“从清单里选”）。5/6/7 本来就不用选工作簿。

`0` 在任何一级选择器里都是“返回菜单”，不会误触发默认动作。

**批量前提（菜单 3、4 共用）**：工作簿无打开密码 · VBA 工程未设置密码保护 · 写回时文件非只读 · 未被 Excel/网盘占用 · 不在受保护视图。单个失败不中断整批，最后给「成功 X / 失败 Y」，有失败时退出码 1。

**单个 `push` 前提（2026-11-21 起自动路由）**：工作簿已打开 → 借用已打开会话写回，保留未保存改动；工作簿未打开 → 自动开隐藏 Excel 打开磁盘文件、写回、保存、关闭。只有加 `--open-only` 才恢复旧行为（工作簿未打开就报「没找到正在运行的 Excel」）。工作簿已打开但 COM 连不上时不会另开副本，直接报 `COM_CONTEXT`，避免覆盖未保存内容。
**工作簿已打开且有未保存改动时**：`--no-save` 只是“VBA 不写 xlsm”，并不能保护单元格数据；要保数据就①另存副本并把 `mapping.file` 指向副本后 push，或②先手动保存（数据与 VBA 一起写入文件）再 push。直接关闭工作簿不保存会丢内存里的 VBA 改动。`pull --apply` 会逐项目探测：工作簿已打开就自动借当前已打开会话抓取（含未保存改动，不另开副本、不读陈旧磁盘），未打开则走隐藏实例读磁盘；结果行用 `mode: open|disk` 标出实际路径。`pull-open` 保留为只借用已打开会话的显式入口。工作簿已打开但 COM 连不上时不另开副本，直接报 `COM_CONTEXT`，不会静默退回磁盘读陈旧副本。

**批量 `push --all-project`**：逐个判断；已打开的工作簿借会话写回（不关闭），未打开的工作簿用本次隐藏实例写回后关闭。单个失败不中断整批。

**自动写回前提（菜单“常驻自动写回”、`watch`）**：工作簿已在 Excel 里打开 · VBA 工程已解除密码保护 · 信任中心勾了「信任对 VBA 工程对象模型的访问」（AccessVBOM）。菜单只监听**当前项目**；命令行给了 `--project` 就只监听那个项目，没给才遍历全部已登记项目。某个工作簿工程访问不了就跳过它并告警，一个都没有才报错退出。公共库文件只写各工作簿 `auto` 挑出来的那几个模块，其余改动打印「跳过」。

**自动写回的启动核对与备份/日志规则**（第三轮需求）：

- 目标集合：菜单按当前项目名传 `--project`，只监听这个项目；命令行给了项目名就 `connect_workbook` 监听那一个，没给就遍历 `cfg["projects"]`，用 `resolve_target` + `find_open_workbook` 挑出打开着的，`has_vbproject_access` 不过的跳过。
- 比对只针对“每个工作簿真正用的”文件：项目 `src` 全部 + 公共库 `auto` 子集（`watch_compare_set` 做这个过滤；子集外的公共模块不参与，否则会被报成“工程里没有”）。
- 比对前先剥掉导出头再逐行比：`strip_attribute_lines` 处理 `Attribute VB_*`（含续行）、`VERSION 1.0 CLASS`、`VERSION 5.00`（窗体）、`BEGIN…END` / `Begin {GUID} 名称 … End`（类与窗体）；再统一换行、去行尾空白、去首尾空行。
- 四档结论：`一致` / `仅大小写/空白`（再折叠成“单空格 + 小写”比较，`Rows`↔`rows` 这类 VBE 自动格式化不算冲突）/ `内容不同` / `工程里没有`。只有后两档算冲突，进清单等你决定；多个工作簿时每条清单带 `[项目名]` 前缀，**一次决定管所有工作簿**。
- 四条规则：`src` 以源码为准（把这批模块写回各自的工作簿）、`vbe` 以工作簿内为准（用 `export_component` 逐个抓回对应 `src` 目录）、`each` 逐个问（`[1] 源码 [2] 工作簿内 [3] 跳过`）、`skip` 先不动。菜单里是 `pick_list` 四选一 + 可选逐个；命令行是 `--startup`。
- **非交互运行（后台/管道，`sys.stdin.isatty()` 为假）一律按 `skip`**，只把清单打出来，不会挂在那里等键盘。
- 以工作簿内为准抓回后**必须立刻刷新所有工作簿的基线**（`watch_refresh_baseline`：新路径记 `(mtime_ns, 大小)`，旧路径 `pop`），否则下一轮会把它当成“新改动”又写回工作簿里，来回打乒乓。
- 备份：`backup_workbook(cfg, wb, tag="-自动写回")` 只在启动时对每个工作簿各执行一次（`--no-backup` 关掉）；写回本身不逐次备份整工作簿。备份完顺手按 `backup_auto_clean` 清理旧批（每类最近 10 批 + 每天 1 份 + 最新 3 批不动）。
- 日志：`WatchLog` 同时打控制台（`[时:分:秒]`）和 `.backup/watch/<项目名>-<YYYYMMDD-HHMMSS>.log`（多个工作簿时 `<项目名>` 是 `全部项目`；`YYYY-MM-DD HH:MM:SS`，UTF-8 + LF，逐行追加，首行是启动时间）；`note` 走正常输出、`problem` 走 `warn`（`[警告]` 前缀）、`quiet` 只进日志（“仅大小写/空白”的模块清单走这条，不刷屏）；启动与停止都会打印日志路径。
- 汇总规则：`stats.written` / `stats.skipped` 同时统计启动 `src`/`each` 写回与轮询写回；启动写回成功不再出现“日志有写回、汇总仍为 0”。

### 2.2 交互约定（所有层级统一）

- **只认数字编号**，`0` 恒等于**退出当前层**（主菜单 = 退出程序，子层 = 返回菜单）。
- **菜单选中即进入动作**：只读动作直接运行；写动作先自动预览，预览通过后 **回车 = 执行**，输入 `0` 返回动作菜单。
- **菜单不重复问确认**：没有单独的“预览 / 执行”模式，也没有 `y/N`；预览输出就是确认页。
- **输错编号**：提示「没有编号 X」后停在本层重输；预览确认页只接受回车或 `0`。
- **写回预览内容**：`auto` 挑出的公共模块连同**每个模块的依据**（模块名 / 过程名 X / 闭包←谁 / 占位壳 / 版本常量）和**文档模块开关**在同一屏打出来；判定不准时按 `0` 返回，不会写入任何文件。批量写回用 CLI（`push --all-project --apply`）。
- 已删除：`q` 取消/返回、`0. 手动输入模块名`、独立的「文档模块 y/N」和「按清单写回 y/N」两问；写回确认并入自动预览。

### 2.3 命令行（脚本化/试跑用）

一条规则：**`sync.py <动词> [--project] [--module] [--macro] [--json] [--apply]`**。
只读动词（list / doctor / scrub / version / vbamod）直接出结果；写操作命令默认**只预览**，加 `--apply` 才执行。

```bat
:: 版本与初始化
python vba-workbench\scripts\sync.py --version
python vba-workbench\scripts\sync.py init --project 示例表A --dir "D:\表格\示例表A" --json
python vba-workbench\scripts\sync.py init --project 示例表A --dir "D:\表格\示例表A" --apply
python vba-workbench\scripts\sync.py init --project 示例表B --file "D:\表格\示例表B.xlsm" --apply  :: 无版本号，按文件精确登记
python vba-workbench\scripts\sync.py init --project 外部表 --file "D:\外部\外部表 V1.xlsm" --allow-outside --apply  :: 菜单拖入外部文件
python vba-workbench\scripts\sync.py init --project 示例表A --lab --apply   :: 生成 .lab/示例表A V1.xlsm 空白测试表

:: 只读
python vba-workbench\scripts\sync.py list --json
python vba-workbench\scripts\sync.py doctor --json
python vba-workbench\scripts\sync.py doctor --verify-package --json   :: 显式对账 .installed.json 哈希与版本
python vba-workbench\scripts\sync.py scrub --json                    :: 按 scrub.deny 扫交付件，命中退出 1
python vba-workbench\scripts\sync.py scrub --path "C:\...\vba-workbench" --json  :: 复核已安装副本
python vba-workbench\scripts\sync.py version --json
python vba-workbench\scripts\sync.py vbamod --json
python vba-workbench\scripts\sync.py vbamod --project 示例表A --json
python vba-workbench\scripts\sync.py plan --project 示例表A --intent push --json
python vba-workbench\scripts\sync.py plan --project 示例表A --intent pull --json

:: 抓取（默认预览，--apply 才写 src）
python vba-workbench\scripts\sync.py pull --all-project --json
python vba-workbench\scripts\sync.py pull --project 示例表A --apply
python vba-workbench\scripts\sync.py pull-open --project 示例表A --apply

:: 写回（默认预览；--apply 才碰 Excel；单个工作簿未打开会自动开隐藏实例）
python vba-workbench\scripts\sync.py push --project 示例表B --json
python vba-workbench\scripts\sync.py push --project 示例表B --apply
python vba-workbench\scripts\sync.py push --project 示例表B --apply --open-only  :: 旧行为：必须已打开
python vba-workbench\scripts\sync.py push --all-project --apply

:: 运行宏
python vba-workbench\scripts\sync.py run --project 示例表A --macro 性能测试 --apply
python vba-workbench\scripts\sync.py run --project 示例表A --macro 测试表格工具 --arg False --arg True --tail 30 --apply
python vba-workbench\scripts\sync.py run --project 示例表B --macro 刷新全部 --push --apply
python vba-workbench\scripts\sync.py run --project 示例表B --macro 刷新全部 --isolate --save --apply

:: 维护
python vba-workbench\scripts\sync.py watch --project 示例表C --apply
python vba-workbench\scripts\sync.py watch --project 示例表C --apply --seconds 600
python vba-workbench\scripts\sync.py shells --json
python vba-workbench\scripts\sync.py shells --mode clean --apply
python vba-workbench\scripts\sync.py align --project 示例表B --apply
python vba-workbench\scripts\sync.py promote --project 示例表A --module 表格工具 --apply
python vba-workbench\scripts\sync.py gbk-export --all-project --apply
python vba-workbench\scripts\sync.py clean-backup --keep 3 --apply
```

公共参数由骨架统一附加：`--json`、`--no-input`、`--apply`、`--project`、`--all-project`、`--version`、`-i`。某个动词到底接受哪些参数，以 `references/commands.md` 的“参数”列为准；本节只列容易误用的专用开关，不再重复全量清单，避免文档漂移。

容易误用的专用开关：
- `pull` / `pull-open` / `push` / `watch`：`--no-backup`；`push` 另有 `--doc-modules`、`--no-save`、`--open-only`、`--rollback-on-fail`。
- `run`：`--macro`（必填）、`--arg`（可重复）、`--push`、`--isolate`、`--save`、`--tail`、`--log`、`--log-optional`、`--rollback-on-fail`。
- `plan`：`--intent pull|push|run|watch`（默认 `push`）。
- `watch`：`--interval`、`--seconds`、`--startup`；`--no-save`、`--no-backup`、`--doc-modules`。
- `shells`：`--mode create|clean`。
- `init`：`--dir`、`--file`、`--prefix`、`--src`、`--config`、`--allow-outside`、`--lab`。
- `clean-backup`：`--keep`（每类保留批数，最小 1；传 0 会钳到 1 并在 `keep_note` 说明）、`--keep-tag`（可重复；关键字以 `-` 开头时必须用等号写法）。
- `doctor`：`--verify-package`。
- `scrub`：`--path`（扫描包根；默认当前工具包根）。

命令名按方向：`pull`（VBE→src，自动判断：已打开借会话、未打开走磁盘）、`pull-open`（VBE→src，只借用已打开会话的显式入口）、`push`（src→VBE，单个工作簿默认自动判断是否已打开）、`run`（可选先写回再运行宏；`--isolate` 自己开隐藏实例）。`plan` 是只读推荐，不修改文件、不启动 Excel，只做一次探测。**旧名 `export` / `export-active` / `all`、以及“直接把项目名当动词”的写法已全部删除**；写错动词一律退出 2 并列出可用动词。

> `refs` 子命令已删除。引用检查只发生在**写回之后**：只针对本次实际写回的模块，缺哪个就警告一次，不修改你的引用。

> **`doctor --verify-package`**：只读显式校验安装清单 `.installed.json`。清单缺失给 `PACKAGE_MANIFEST_MISSING`（源码树/未安装副本属正常，不崩）；清单字段缺失、登记文件磁盘缺失、哈希不符、磁盘多出未登记文件、VERSION 与清单版本不一致，归 `PACKAGE_INTEGRITY`，message 列出具体相对路径，hint 提示本机正常改动可忽略、可用安装器重装或 `--apply` 刷新清单。校验结果进 `stats.package`；显式校验不通过退出 1，通过退出 0；不带该开关时 doctor 行为不变。

> **`audit.py` 同一个规矩**：`audit.py <视角> [--project] [--module] [--json] [--out PATH]`。六视角：`members --module M`、`common`、`project --project P`、`structure`、`indicators`、`overview`。它只读、不改任何源码；输出与 `sync.py` 同一信封（`details` 恒为数组）。判定规则见 `references\audit.md`。

### 2.4 退出码与 JSON 字段（agent 看这一段就够）

| 退出码 | 含义 |
|---|---|
| `0` | 成功；`list` 表示全部项目都能解析 |
| `1` | 执行失败或体检不合格（看 `errors` / 可读输出） |
| `2` | 用法错：未知动词、缺 `--project`/`--module`/`--macro`、给 `watch` 传 `--json`、`sync.py` / `audit.py` 直接运行（只打帮助） |
| `130` | 用户 Ctrl+C 中断 |

`--json` 统一信封：`ok / action / details / errors / elapsed_ms / ts / stats / exit_code`，三个入口一致。`action` 固定为 `<入口>.<动词>`（如 `sync.push`）；命令自己的计数字段全部进 `stats`，不再平铺顶层。**除 `watch` 外全部子命令都支持**（`watch` 是常驻命令，带 `--json` 直接退出 2）；脚本要机器读 watch 的前置条件，先运行 `plan --project <项目名> --intent watch --json`，看 `mode/command`。`ok=false` 时 `errors` 一定有内容（版本不一致、体检不合格、结构问题都算失败）；写操作命令预览时 `ok=true`，但预览标记不是统一字段：`push` / `pull` / `pull-open` / `run` 用 `stats.dry=true`，`clean-backup` 用 `stats.applied=false`，`promote` 用 `details[0].dry`；不要把 `stats.dry` 当成所有写命令的固定字段。`exit_code` 与进程退出码一致。

失败项是对象：`{"code", "stage", "message", "hint", "suggested_command", "retryable"}`。`stage` 取值 `config/parse/resolve/attach/open/vbe/pull/push/run/preflight/postcheck/save/patch/logs/assert/init/contract/unknown`；`backup` 和 `import` 不是有效 stage，不要按旧文档判断。`code` 可能是 `CONFIG_NOT_FOUND/TARGET_NOT_FOUND/NO_EXCEL/TARGET_NOT_OPEN/COM_CONTEXT/VBE_PROTECTED/WORKBOOK_OPEN_FAILED/FILE_READONLY/TARGET_OPEN/SAVE_FAILED/ARG_MISMATCH/RUN_FAILED/SOURCE_PREFLIGHT/POSTCHECK_FAILED/SCRUB_HIT/WATCHDOG_TIMEOUT/PACKAGE_INTEGRITY/PACKAGE_MANIFEST_MISSING/UNEXPECTED` 等。缺宏、禁用宏统一收口 `RUN_FAILED`；弹窗识别族已回补：编译错误框 `VBE_COMPILE_ERROR`，其余模态框 `WATCHDOG_TIMEOUT` 且 hint 点名窗口；旧细分码 `VBE_DENIED` / `MACRO_NOT_AVAILABLE` / `MACRO_DISABLED` / `VBE_DIALOG` / `MODAL_DIALOG` 已不在枚举中。agent 先看 `errors[0].code`，不要只凭 message 猜。`TARGET_NOT_OPEN` 与 `NO_EXCEL` 的 `hint/suggested_command` 都按动词分流：`push` 给自动打开工作簿或 `--open-only`，`run` 给 `--isolate`，`pull-open` 给 `pull`，`watch` 给 `agent.py book --action open --apply`；建议命令里直接写项目名，不留 `<项目名>`。`run --isolate` 预检遇 `COM_CONTEXT` 不退回磁盘副本：隔离模式也不启动，hint 列弹窗 / 受保护视图 / 异权会话三类成因，suggested_command 为可复制的 isolate 入口命令；JSON 信封里 hint / suggested_command 各就各位，不混进 stdout。

| 命令 | 结果行位置与字段 |
|---|---|
| `list` | `details[]` 每个项目一条：`name / ok / workbook / reason / src / common_modules`；未解析项同时进 `errors[]`（`code=TARGET_NOT_FOUND`、`stage=resolve`），`stats.total/resolved` 汇总 |
| `plan` | `details=[]`；单条结果在 `stats.results[0]`，字段 `name / intent / ok / workbook / target_reason / open / saved / vbe_state / probe / probe_detail / mode / needs_open / will_save / preserve_unsaved / command / risk`。同名摘要字段（`project/intent/mode/needs_open/will_save/command`）也在 `stats` 顶层。`mode` 取值 `open / disk / isolate / open_required / unavailable / unknown`；目标表缺失时 `errors[0].code=TARGET_NOT_FOUND` |
| `init` | `details[0]` 是登记项：`name / search_dir / prefix / require_version / src / note`；按文件精确登记时多一个 `file`。`stats.results[0]` 同形，另有 `stats.config / stats.dry / stats.created_src`。目录或文件不存在时 `errors` 非空、退出 2。`--lab` 生成空白表失败：缺脚本报 `MISSING_LAB_TEMPLATE/stage=init`，Excel/COM 失败报 `LAB_TEMPLATE_FAILED/stage=init` |
| `pull` | `details=[]`；每个项目的结果行在 `stats.results[]`：成功 `name / ok / modules / mode`，失败 `name / ok / reason / mode`。`mode=open / disk` 是该项目实际走的路径；预览标记只在 `stats.dry`，不在结果行 |
| `push` | 单张 `details[]`：`name / ok / target / modules`（`--apply` 后带 `mode / saved / backup_path`）；批量 `details[]`：`name / status` |
| `run` | `details[]`：`name / ok / macro / mode / args / result`（`--isolate` 时带 `mode / saved`）；`stats.cost_sec` 是整条命令耗时，`stats.run_sec` 是 `Application.Run` 本身耗时。journal 的 `mode` 统一为 `preflight / attach / isolate`，非隔离成功也是 `attach`，不再出现 `open`。失败时带 `error`；参数个数不符为 `ARG_MISMATCH` |
| `doctor` | `details[]` 每项一条：`name / level / detail`，`stats.failed` 是 FAIL 数；带 `--verify-package` 时另出 `stats.package`（`present / checked / missing / mismatch / extra / missing_fields / version_match`），显式校验不通过退出 1 |
| `scrub` | `details[]` 每命中一处一条：`file / line / word / text`；`stats.scope / files / deny_terms / hits / matched_terms / skipped_non_text`；命中 `SCRUB_HIT/stage=preflight` 且退出 1，零命中退出 0 |
| `pull-open` | `details=[]`；结果行在 `stats.results[]`：`name / ok / modules`。`stats.saved=false` 固定（本命令不保存工作簿），`stats.workbook_saved` 是抓取前状态，`stats.wrote_src` 表示源码已写入文件 |
| `version` | `details[]` 每个项目一条：`name / locked / match`（未锁定时带 `reason`），`stats.common_version / stats.configured / stats.common_dir / stats.mismatched`；`configured=false` 表示没找到 `src/common` 或缺少 `COMMON_VERSION`，不要把它当成“校验通过” |
| `shells` | `details[]` 每个项目一条：`name / ok / action(create/clean) / count`，`stats.created / stats.cleaned / stats.dry` |
| `align` | `details[]` 每个项目一条：`name / swapped / kept / already`，`stats.swapped / stats.kept / stats.dry` |
| `promote` | `stats.module / stats.project / stats.from_path / stats.to_path / stats.others / stats.shell_written / stats.refs`；`details[0]` 为 `{name, ok, to}`，预览带 `details[0].dry` |
| `gbk-export` | `stats.exported / stats.shells_skipped / stats.out_dir / stats.dry`；`details[]` 每组 `{name, exported, shells_skipped, out_dir, cleaned}`（缺源目录时带 `skipped`）。`--apply` 会先清空目标组再重建镜像，避免旧模块残留 |
| `clean-backup` | `stats.applied / stats.freed_bytes / stats.deleted_batches / stats.kept / stats.kept_bytes / stats.keep_recent / stats.keep_requested / stats.keep_note / stats.keep_daily / stats.min_keep / stats.auto_clean / stats.keep_scope(per-tag) / stats.by_tag / stats.kept_by_tag / stats.kept_names / stats.leftover / stats.protected`；`details[]` 为 `{batch, bytes}`，预览也逐条列出将删批次；`--keep 0` 会被钳到 1 并在 `keep_note` 说明 |
| `vbamod` | `details=[]`；每个项目的结果行在 `stats.results[]`：`name / file / modules / module_names / compared / same / count_before / count_after / lost / added`；`stats.examined / stats.changed` |

`details[]` 的键名统一用英文；`doctor` 固定 `name / level / detail`，`probe` 固定 `project / file / disk / open / vbe / modules / test_sheet / rows`；`plan / pull / pull-open / vbamod` 的结果行在 `stats.results[]`，不在 `details[]`。消费方按动词取字段，但字段名不再中英混用。

## 3. 排错

| 症状 | 原因 | 怎么办 |
|---|---|---|
| `无法访问 VBProject` | 没勾信任中心 / VBA 工程设置了密码保护 / 工程受保护 | 信任中心勾「信任对 VBA 工程对象模型的访问」；Excel 里取消工程保护 |
| 启动核对里**整个工作簿的模块全报「工程里没有」** | 工程受密码保护：`VBProject.Name` 照样读得到，一碰 `VBComponents` 才报「该工程已被保护」 | 用 `vbe_state()` 预检，现在会直接打「「<项目名>」的 VBA 工程现在读不了：…」；交互模式下**当场解除密码保护后按回车就继续**，不必退出重运行；后台/批量则跳过这个工作簿 |
| 写回时报文件被占用 | 工作簿已打开，或同步盘在同步 | 关闭工作簿窗口；暂停同步；再重试 |
| 写回后工作簿里中文变问号 / 乱码 | push 边界的 GBK 临时文件没生效 | 确认 `import_folder` 走 `gbk_import_copy` 转码；src 本身应是 UTF-8，别手工把 src 转成 GBK |
| 项目目录里出现空壳模块 | 某次写回把壳当成了模块 | 检查 `is_shell_file` 判据；从公共库重新写回 |
| `all` 里某个工作簿失败 | 单个异常被捕获进汇总 | 看该工作簿的 `技术细节`；常见是工程密码保护与文件占用 |
| 抓取后少了模块 | 该模块在某项目里是公共模块（不导出正文） | 正常，去 `src/common` 看真身 |
| 窗体打不开 | `.frx` 没跟 `.frm` 一起 | 检查两个文件是否同目录同名 |
| 预览正常、`--apply` 才失败 | 预览不碰 COM，实际执行要 Excel/VBE | 用 `doctor` 查 pywin32/AccessVBOM，再对单个工作簿 `push --project X --apply` 试 |

---

---

## 4. 改动后的自检

```bat
python -m py_compile vba-workbench\scripts\sync.py     :: 语法
python vba-workbench\scripts\sync.py doctor --json     :: 环境
python vba-workbench\scripts\sync.py list --json       :: 配置解析
python vba-workbench\scripts\sync.py push --all-project --json :: 写回演练（不动文件）
:: 交互菜单：终端里运行 sync.py -i 走选项目/选动作流程；管道里运行应给可读提示并退出 1（见 test_sync_interactive_non_tty_is_clean）
```

死代码自检（改完删函数后跑一次，输出应只剩"仅被引用 1 次"的正常函数）：

```powershell
python -c "import ast,pathlib,collections;t=ast.parse(pathlib.Path('vba-workbench/scripts/sync.py').read_text(encoding='utf-8'));d={n.name:n.lineno for n in t.body if isinstance(n,ast.FunctionDef)};u=collections.Counter(x.id for x in ast.walk(t) if isinstance(x,ast.Name) and isinstance(x.ctx,ast.Load));print([k for k in d if u[k]==0])"
```

---
