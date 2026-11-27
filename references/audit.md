# audit.py 使用与维护说明

`vba-workbench\scripts\audit.py` 是**只读**的代码盘点工具：不碰 Excel、不改源码，只回答
"这些代码还有没有人用、仓库结构有没有问题"。

它和 `sync.py` 一样只依赖 `vba-workbench\scripts\toolkit\` 包（仓库根定位、UTF-8 读源码（GBK 兜底）、`mapping.json` 加载、
剥注释、占位壳识别、控制台与菜单基础函数、CLI 入口骨架）。**两个入口脚本互不导入**。

> 文档模块的源码文件名带工作表名（`Sheet4(SKU统计).bas`），模块名取自 `toolkit.module_name_of()`；
> `members Sheet4` 用模块名即可。文件名规则见 `toolkit.md` 的「文档模块的文件名」一节。

## 目录

- §1 为什么需要它
- §2 用法（2.1 菜单 / 2.2 命令行）
- §3 八个视角
- §4 判定规则（4.1 成员引用 / 4.2 结构问题）
- §5 输出长什么样
- §6 已知局限
- §7 建议的瘦身流程
- §8 改这个工具时

---

## 1. 为什么需要它

- 源码已改存 UTF-8，中文成员名（`数据同步`、`简化字符串`、`解析字符串`）可以直接用 grep/ripgrep 搜到；
  但"搜到几次"不等于"还有没有人用"，人工盘点一个 4000 行的类仍然太慢。
- "有没有人用"不是数次数能回答的：`A` 只被 `B` 调用、`B` 又没人调用，那 `A` 和 `B` 是**一条死链**。
  所以要按调用图做**可达性**分析。
- 结构问题（旧副本没换壳、壳指向已删除的模块、工作簿在用的公共模块没有壳、两个模块声明同名过程）
  散在各处，靠肉眼发现太慢，而且会直接导致盘点结果出现"疑似同名"噪声。

---

## 2. 用法

### 2.1 交互菜单

`audit.py -i` 进入交互菜单；非 TTY / 管道环境直接给可读提示并退出 1。
菜单只读、不写文件；导出 Markdown 时由既有 handler 落盘。

```bat
python vba-workbench\scripts\audit.py -i
```

```
  1. 盘指定模块          先选范围（公共库 / 某个工作簿），再按分类选模块，看成员级明细
  2. 公共库总览          src/common 全部模块（不选对象时就是它）
  3. 项目依赖面          选一个工作簿，看它用了哪些公共模块
  4. 全仓总览            公共库 + 结构问题 + 整体瘦身意见
  5. 结构体检            旧副本 / 孤儿壳 / 缺壳 / 重名过程 / 文档模块对账 / 规则回退 / 绑定目标
  6. 指标体检            H01–H19 指标复跑
  7. 导出 Markdown 报告  选视角与范围，写到指定 .md
  8. 显示设置            排除自测 / 列出全部明细（会话内记住）
  0. 退出
```

交互约定与 `sync.py` 一致：**只认数字**、`0` 恒为退出当前层、回车视为无效输入、
动作结束直接回菜单（异常不会把菜单崩掉）。**选择分级**：菜单 1 和 7 先列范围
（公共库 / 各项目），选中后按分类（模块 / 类模块 / 窗体 / Excel 对象）分组列文件，
不一次性罗列几十个选项。**显示设置**集中在菜单 8：排除自测（默认开）/ 列出全部明细
（默认关），选编号即切换、会话内记住；每次执行前打印一行「判定规则：…」。

### 2.2 命令行

标准形状：**`audit.py <视角> [--project P] [--module M] [--json] [--out PATH]`**。
视角是白名单，未知视角退出 2；不再支持"裸模块名 = members"和 `--json -`。

```bat
python vba-workbench\scripts\audit.py -i                        :: 交互菜单（需要终端；管道里退出 1）
python vba-workbench\scripts\audit.py common --json             :: 公共库全部模块
python vba-workbench\scripts\audit.py members --module 表格工具 --json
python vba-workbench\scripts\audit.py project --project 示例表B --json
python vba-workbench\scripts\audit.py structure --json          :: 仅警告时退出码 3；加 --strict 才按 1
python vba-workbench\scripts\audit.py indicators --json
python vba-workbench\scripts\audit.py overview --json           :: 只有阻断组退出 1；规则提醒默认 0，--strict 才 1
python vba-workbench\scripts\audit.py overview --out 其他\全仓体检.md
python vba-workbench\scripts\audit.py lint --project 示例表A --json
python vba-workbench\scripts\audit.py references --project 示例表A --json
```

开关（各视角通用）：

| 开关 | 作用 |
|---|---|
| `--project 项目名`（可重复；`common` / `overview` 有效） | 统计范围只纳入这些项目；公共库内部调用仍计入 |
| `--module 模块名`（`members` 必填） | 指定模块/类/窗体；也接受 `src` 下相对路径 |
| `--exclude 关键字`（可重复） | 排除文件名含该关键字的源文件，如 `--exclude 测试` |
| `--ignore 分组/指标`（可重复，也接受逗号） | 仅 `structure` / `indicators` / `overview` 有效。结构分组如 `manual_accel`、`binding_target`；指标如 `H05a,H06`。`members/common/project` 不接受该参数 |
| `--strict` | `structure` / `indicators` / `overview` / `lint` / `references` 有效：有警告时按失败 1 处理。前两者与 `lint` / `references` 默认 3；`overview` 默认只对阻断组退 1，规则提醒默认 0 |
| `--all` | members 列全部成员；common 展开每个模块明细；project 展开本工作簿专属模块明细；`lint` 兼容但总是输出全部 |
| `--md PATH` | 另存 Markdown 报告（UTF-8 无 BOM + LF） |
| `--json` | stdout 只输出一行信封 JSON（`ok / action / details / errors / elapsed_ms / ts / stats / exit_code`）；**不打印可读文本和表格**，警告并入 `stats.warnings`。可读报告只在非 `--json` 模式输出；要留档用 `--out` |
| `--out PATH` | 把同一份 JSON 落到文件（不再用 `--json -`） |
| `--no-input` | 禁止提问，缺参数直接退出 2 |
| `-i` | 交互菜单：选范围 / 分类 / 模块，跑原六个视角，导出 Markdown；需要终端，管道里退出 1。`lint` / `references` 目前只走命令行 |

退出码：`0` 无异常；`1` 真异常 / 执行失败（`--strict` 下警告也归 1）；`2` 用法错；`3` 只读体检仅有警告（`structure` / `indicators` / `lint` / `references`）。`overview` 默认只对阻断组（`sheet_mismatch` / `binding_target` / `dup` / `shadow`）退 1，`manual_accel` 等规则提醒只计数、不退 3。

---

## 3. 八个视角

| 视角 | 对象 | 输出 |
|---|---|---|
| `members --module M` | 任意模块/类/窗体 | 成员表 + 四类判定 + 可删行数 + 死链清单 |
| `common` | `src/common` 全部模块 | 每模块一行摘要（成员/四类计数/可删行/引用方），零外部引用的模块，合计 |
| `project --project P` | 单个工作簿 | 依赖矩阵（引用次数、用到几个成员、涉及几个文件、主要引用方）、未被引用的公共模块、本工作簿专属模块 |
| `overview` | 公共库 + 全仓结构 | 上面两者的汇总 + **按影响排序的整体意见**，每条都带能直接跑的命令。公共库为空/未配置时仍出项目与结构总览（`stats.modules=0`、`stats.common_configured=false`），不整条失败 |
| `structure` | 全仓源码结构 | 结构问题清单（旧副本/孤儿壳/壳类型/缺壳/重名/文档模块/旧规则回退） |
| `indicators` | 指标体检 | 每项一条 `{id, title, status, detail, hint}`，`stats.passed / stats.warned` |
| `lint --project P` | 项目源码 | 四条高置信规则：未使用局部变量、重复声明、函数未赋值返回、裸 `Range/Cells`；只读，默认 warning 退 3 |
| `references --project P` | 工作簿 VBA 工程 | 每条引用的 `ok/missing/broken`、GUID/版本/路径、轻量 `used_by`；只读，默认 warning 退 3 |

`common` 要求有可盘模块；没有时按只读警告退出 3（`ok=true`、`stats.configured=false`），打印 `common_dir` 与配置键 `defaults.common_dir`，不伪装成代码失败；`overview` 不以“公共库未配置”失败，仍输出项目与结构总览。

`common` 会额外标出**版本常量模块**（按严格正则 `VERSION_RE` 判定：必须是 `COMMON_VERSION = "x.y.z"` 形式的常量声明，如 `CommonVersion.bas`；注释或字符串里提到不算）：
它零外部引用也不能退出公共库，因为 `sync.py version` 要读它。

### lint

```bat
python vba-workbench\scripts\audit.py lint --project 示例表A --json
python vba-workbench\scripts\audit.py lint --project 示例表A --strict --md 其他\lint.md
```

四条规则固定：`L01_UNUSED_LOCAL`、`L02_DUPLICATE_DECL`、`L03_MISSING_RETURN`、`L04_UNQUALIFIED_RANGE_CELLS`。规则实现在 `toolkit/vba_lint.py`，`features/audit_lint.py` 只做项目枚举与出口。命中默认 warning 退 3，`--strict` 才 1；不做自动修复。

### references

```bat
python vba-workbench\scripts\audit.py references --project 示例表A --json
python vba-workbench\scripts\audit.py references --project 示例表A --md 其他\引用体检.md
```

只读读取工作簿 `VBProject.References`；表没开时用隐藏实例只读打开自己的副本，读完关闭。状态：`ok` / `missing` / `broken`。工程被锁或信任中心未放行时 `stats.readable=false`、退出 3，不伪报业务失败。`used_by` 是按类型前缀扫源码的轻量提示，不是硬判定。

---

## 4. 判定规则（结果准不准全看这一节）

### 4.1 成员引用

| 环节 | 规则 |
|---|---|
| 读源码 | 走 `toolkit.read_source_text`（UTF-8-BOM → UTF-8 → GBK），**逐行剥掉注释**再匹配（`toolkit.strip_vba_lines`：字符串外的 `'` 截断、`Rem` 开头整行置空、全行去尾空白）；字符串里的单引号不当注释 |
| 统计范围 | `src/common/**` + `src/projects/**`，**跳过公共库占位壳**（壳只有一行注释） |
| 类 / 窗体（`.cls`、`.frm`） | 只认「确定接收者」：`实例.成员`、`With 实例 … End With` 里的 `.成员`、工厂链 `getTL(...).成员` |
| 实例名怎么来 | `Dim/Private/Public x As 表格工具`、`Set x = New 表格工具`、`For Each x As 表格工具`，以及**工厂函数赋值** `Set x = getTL(...)`——工厂 = 任何 `Function xxx(...) As 表格工具`。所以 `Private 表页1`（无类型 Variant）也能认出来 |
| 标准模块 / 文档模块（`.bas`） | 过程是**裸名调用**，判定改为：裸名 `xxx(...)` / `Call xxx`，或 `模块名.xxx`；并排除该文件自己声明的同名过程（那是在调它自己的） |
| 「疑似同名」列 | 接收者不是本类实例的 `xxx.成员`（如 `fso.Path`、别的工作簿的同名属性）。**只提示，不参与判定** |
| 隐式入口 | `Class_Initialize` / `Class_Terminate` / `Sub Main` / 事件过程（`*_Click`、`*_Change`、`*_Initialize` …）。它们是外部触发的，永远算可达 |
| 死代码 | 入口 = 有外部引用的成员 + 隐式入口；从入口沿调用图走不到的 → **不可达** |

四类结果的含义：

| 判定 | 含义 | 能干什么 |
|---|---|---|
| **不可达** | 外部零引用，且内部也没有可达路径走到它 | 整条死链一起删（含只被死成员调用的下游） |
| **仅自测** | 只有 `测试*` 模块在用 | 业务侧可删；删之前先瘦 `测试表格工具.bas`，否则它变编译错误 |
| **仅单文件** | 只被一个文件引用 | 多半是那个模块的私有工具，可考虑从公共库**下沉**回去（或标 `Private`） |
| **活跃** | 多处引用 | 保留 |

「行数」= 声明行到 `End Sub/Function/Property` 的行数（字段算 1 行），用来排优先级。

### 4.2 结构问题（`structure` / `overview`）

| 分类 | 判定 | 处理 |
|---|---|---|
| 旧副本 | 项目目录里的文件与公共库同名，且**不是占位壳** | `python vba-workbench\scripts\sync.py align --project <项目名> --apply` 逐个换壳 |
| 孤儿壳 | 项目里有壳，公共库已经没有这个模块 | `python vba-workbench\scripts\sync.py shells --mode clean --apply` |
| 壳类型不符 | 壳的扩展名与公共库同名文件不一致（`.bas` ↔ `.cls`） | 删壳后 `python vba-workbench\scripts\sync.py shells --apply` 重建 |
| 缺壳 | 工作簿在引用某公共模块，但项目目录里没有同名壳 | `python vba-workbench\scripts\sync.py shells --apply` |
| 公共库内重名过程 | 两个**标准模块**声明了同名 `Sub/Function`，裸名调用会有歧义 | 改名，或调用方统一用 `模块名.过程` |
| 项目与公共库重名过程 | 项目标准模块声明了公共库同名的过程 | 裸名调用会先撞项目里那份；删旧实现或改名 |
| 源码有、工程没有的文档模块 | 读 xlsm 里的 `vbaProject.bin` 取工作表模块名，与 `src` 里的 `SheetN` 文件对账（**不打开 Excel、不要工程密码**） | 工作表已删除 → `python vba-workbench\scripts\sync.py pull --all-project --apply` 清掉残留文件 |
| 工程有、源码没有的工作表模块 | 同上（反向） | 抓漏了 → 确认工作表存在后 `sync.py pull --all-project --apply` 补抓 |
| 裸 `New 表格工具` | 非占位文件里出现 `New 表格工具`（`自定义通用函数.bas` 的 `getTL` 本体、`测试表格工具.bas` 的性能被测对象除外） | 改成 `Set 工具 = getTL(表页, "表名")`：复用实例、顺带登记日志宿主 |
| 手写加速 | 只统计 `Application.ScreenUpdating/Calculation/EnableEvents = False/xlCalculationManual` 这类**关闭赋值**；同一过程带 `On Error` 且后续有还原时整组放行，错误分支里重复还原不再算违规（`自定义通用函数.bas`＝加速本体、`测试表格工具.bas`＝性能对照，这两个文件除外；上方 3 行内标 `' [加速例外]` 的块也跳过） | 改成 `Call 开启加速(...)` / `Call 关闭加速`；计数式嵌套才不会提前恢复 |
| 缺过程注释 | `Sub/Function` 头上没有紧邻注释块：只跳过空行，连续 `'` / `Rem` 行算；遇到代码/`End Sub` 即停（事件过程和 `Class_Initialize/Terminate/Main` 跳过，`Property` 不查） | 补 `' 函数名 / ' 功能 / ' 参数` 三行，格式照 `表格工具.cls` |

文档模块对账的判定与 `sync.py` 写回前的预检**共用一份实现**（`toolkit.check.check_doc_modules`），
所以两边结论不会漂。审计结束会打一行统一汇总，形如 `── 结构体检提醒：警告 2 · 预检 1 ──`，
分级为 `失败 / 警告 / 预检 / 校验`（与 `sync` 相同，见 `toolkit.md`「统一校验」）。

事件过程（`CommandButton1_Click` 之类）天然重名，这两项检查会跳过；类模块的成员也不会和
模块过程冲突（类成员只能通过实例访问），所以只比 `.bas`。

---

## 5. 输出长什么样

```
────────────────────────────────────────────────────────
公共库总览（src/common，6 个模块）
────────────────────────────────────────────────────────
  范围    ：78 个源文件（公共库 + 全部项目，已跳过占位壳）；排除关键字：测试

  模块            类型  成员  不可达  仅自测  仅单文件  活跃  可删行  引用方
  ──────────────────────────────────────────────────────────────────────────
  CommonVersion   模块     1       1       0         0     0       3       0
  HTTP请求与解析  模块    10      10       0         0     0     449       0
  jsonToDict      模块    26      13       0         1    12     574       1
  数据同步        模块    23       2       0         5    16       8      16
  自定义通用函数  模块    55      13       0        17    25     111      40
  表格工具        类     168      43       0        35    90     655      34

  合计：成员 283，不可达 82（约 1800 行），仅自测 0，仅单文件 58
```

`overview` 在这个表格下面给“整体意见”，例如：

```
  1. 「表格工具」删不可达 43 个成员，约 655 行：AutoFilterData、CPU序列号、FilterGreater…
  2. 「jsonToDict」删不可达 13 个成员，约 574 行：ApiExample、ParseIso…
  5. 「表格工具」35 个成员只被单个文件引用 —— 下沉回那个工作簿或改 Private
 10. 重名过程 3 个 —— 用 structure 看清单，改名或改用「模块名.过程」限定
```

成员表里 `外部` = 确定引用次数，`类内` = 有多少个过程引用它，`疑似` = 同名歧义次数，
`行数` = 删掉能省多少行，`⚠动态调用` = 成员名出现在 `Application.Run` / `Evaluate` 的字符串参数里。

---

## 6. 已知局限（删代码前必须人工过一遍）

1. **动态调用看不见**：`Application.Run "模块.过程"`、`Evaluate`、按钮/快捷键绑定的宏，
   静态分析扫不到。工具会标 `⚠动态调用`，但**标准模块里的 `Public Sub` 即使零引用也可能是宏入口**。
2. **`With` 嵌套**：按顺序压栈，`With` 目标写成跨行拼接的极端写法会漏。
3. **跨工作簿调用**：只统计本仓库 `src/` 里的引用；别的 xlsm 通过 `Application.Run "'表.xlsm!过程"`
   调它，这里看不见。
4. **重名成员**：不同类里的同名成员在「疑似」桶里会串，但判定只看「确定」桶，不影响结论。
5. **`--project` 的语义**：只决定纳入哪些**项目**文件；公共库内部的相互调用仍然计入
   （它们是真实引用，删成员前必须看见）。要查看“这个工作簿用了什么”，用 `project <项目名>`。

---

## 7. 建议的瘦身流程

1. `python vba-workbench\scripts\audit.py overview --exclude 测试 --json` → 先读整体意见，按影响排序。
2. `python vba-workbench\scripts\audit.py structure --json` → 先把结构问题清掉（旧副本/缺壳），否则盘点会带噪声。
3. `python vba-workbench\scripts\audit.py members --module 表格工具 --exclude 测试 --md 其他\表格工具-成员盘点.md`
   → 从「不可达」里先删**叶子**（`类内` 也为 0 的）。
4. 删完再跑一次：原先"只被死成员调用"的会浮出来，逐轮收敛（每轮都重算可达性）。
5. 「仅单文件」那批，逐个判断是该 `Private` 化，还是干脆搬回那个工作簿
   （搬回是 `sync.py promote` 的反向操作：手工挪文件 + `sync.py shells` 刷新壳）。
6. 「仅自测」那批：先决定要不要保留自测模块，再删业务侧无人用的成员。
7. 每次动完公共库，走 `sync.py push --all-project --apply` 写进各工作簿（写回前会先备份），再运行
   `python vba-workbench\scripts\sync.py version --json` 校验版本一致性。

> 提醒：删公共库成员属于"改唯一真源"，一定要先 `pull` 把各表最新代码抓回 `src/`，
> 否则你可能删掉的正是某个工作簿里还没抓回来的代码在用的东西。

---

## 8. 改这个工具时

- 共用逻辑（编码读写、壳判定、剥注释、菜单基础函数、显示宽度、**项目 / 模块选择**、CLI 骨架）一律改 `toolkit` 对应子模块，**不要**在 `audit.py` 里复制，
  也不要 `import sync`。主/子菜单用 `run_menu`、选范围用 `pick_list`、按分类选文件用 `pick_grouped`（配 `group_files` / `source_files` / `project_dirs`），
  入口脚本里不要自己写清单打印 + `choose` 那套；导出产物一律走 `write_text_utf8` / `dump_json`（UTF-8 无 BOM + LF）。
- 分层：`analyze/_analyze_target` 只产数据；`report*` 只排版；`write_*_md` 只出报告；
  `cmd_*` 负责串起来并管退出码。加新检查就加在数据层，别混进排版。
- 解析成员的三个正则（`FIELD_RE` / `PROP_RE` / `PROC_RE`）决定成员表；VBE 的
  `Attribute VB_*` 头行不属于成员，不用管（抓取产物里只有类模块有）。
- 一次命令内同一文件只读一次（`_LINES` 缓存，`cmd_*` 开头清空），加新视角时沿用 `lines_of()`。
- 新增视角：在 subparsers 与 `_dispatch` 里登记、写一个 `cmd_xxx`、在 `main()` 的 `_split_mode` 分支里接一行、
  菜单里加一项，并同步更新本文件和 README 工具索引里 `audit.py` 那一行。
