# sync.py 使用与维护说明

`sync.py` 是这套工具的主入口：**VBE → `src/` 抓取，`src/` → VBE 写回**，外加自动写回、运行宏、备份清理。

## 分册索引

2026-11-21 起，原 840 行拆成四册；每册独立编号，按需选读。

| 我想… | 读 |
|---|---|
| 知道它为什么必须经过 Excel、抓取/写回到底怎么发生、公共库怎么选、占位壳是什么 | 本册 §1、§2、§3 |
| 装环境、用菜单、用命令行、出错怎么查、改完怎么自检 | `sync-commands.md`（§1、§2、§3、§4） |
| 写 `mapping.json`、`defaults`/项目字段、换版本、目录与文件命名 | `sync-config.md`（§1、§2） |
| 改代码：代码地图、硬规则、加菜单/子命令、已知限制、设计取舍 | `sync-maintain.md`（§1–§3、§4、§5） |

本册只讲原理与数据流，前置概念见 `SKILL.md`；分册索引用于选读，不要求按册顺序读。本册目录：§1 这个脚本做什么 · §2 原理与底层逻辑（2.1–2.7）· §3 数据流（3.1–3.5）。

## 1. 这个脚本做什么

把 Excel 工作簿（`.xlsm`）里的 VBA 工程与仓库里的源码目录做双向同步：

```
xlsm 的 VBA 工程  ──抓取──▶  src/projects/<项目名>/   （与 src/common/）
       ▲                              │
       └──────────写回───────────────┘
```

- **抓取**：连 Excel 的 COM，把 `VBComponents` 逐个 `Export` 成 `.bas/.cls/.frm`，按 VBE 工程树分四个目录。
- **写回**：把源码目录下的文件 `Import` 回工程；文档模块（`Sheet1`、`ThisWorkbook`）用 `CodeModule` 写回。
- **公共库**：`src/common` 里的模块可以被多个工作簿共享，写回时按引用**选择性**写回，项目目录里只留一行注释的**占位壳**。
- **备份**：写回前把 xlsm 复制到 `.backup/<时间戳>/`，需要回滚时用写回前自动生成的整工作簿备份；但 `postcheck` 失败默认只放弃保存、不自动改 VBE 内存，要复原本次模块加 `--rollback-on-fail`。自动写回（`watch`）只在**启动时**备一次，之后每次写回不重复备整工作簿（CLI `watch` 默认写回后保存，菜单里的自动写回只改内存；`--no-save` 可统一关闭保存）。
- **备份自动清理**：每次备份成功后自动按规则删旧备份——每个 tag 保留最近 `backup_keep`（默认 10）批、每天再留最新 1 份、全局最新 3 批永不删；`.backup/GBK` 与 `.backup/watch` 不参与。想先看清单再删用 `clean-backup`（默认只预览，`--apply` 才动文件），“全局工具”里的“清理 `.backup` 旧批次”同。

脚本只做文件与工程的读写，**不改你的 VBE 引用设置**（`Tools → References` 永远手工勾）。

入口有两种用法：命令行子命令（见 `sync-commands.md` §2.3），或 `sync.py -i` 交互菜单
（选项目 → 选动作；只读动作选中即运行，写操作选中后自动预览，预览通过后再回车执行；0 返回上一层，
动作结束自动回菜单；非 TTY 环境给可读提示并退出 1）。
写操作执行时才给 handler 加 `--apply`；预览不写入文件。`watch` 会先问监听秒数（默认 600）。
新建项目在菜单里直接拖入 `.xlsm`：带 ` V<数字>` 时自动取最高版本，没有版本号时按文件精确登记。

---

---

## 2. 原理与底层逻辑

这一节回答"脚本凭什么能改动 Excel 里的代码"，也是理解后面所有行为（备份、只读打开、编码转码、权限报错）的基础。

### 2.1 为什么必须经过 Excel

VBA 代码存在 `.xlsm` 内部的 OLE 复合文档里（`vbaProject.bin`），这个二进制格式**没有公开规范**。脚本不解析它，而是让 Excel 自己读写：

```
Python  ──pywin32──▶  COM 接口  ──▶  Excel.exe  ──▶  VBE 的公开 API（VBComponents）
```

所以：**Excel 必须安装、必须是微软原版、必须能打开这个文件**。WPS／Mac 版／网页版都提供不了这套 COM 接口。

### 2.2 抓取：VBE 自己写文件

```python
xl   = win32com.client.DispatchEx("Excel.Application")   # 独立隐藏实例
wb   = xl.Workbooks.Open(路径, 0, True, True)             # ReadOnly
comp.Export(str(target))                                  # 文件由 Excel 写出
wb.Close(SaveChanges=False); xl.Quit()
```

- `comp.Export(...)` 与你在 VBE 里"右键 → 导出文件"是同一个动作，脚本不生成代码文本。
- **它不补扩展名**，所以必须按 `comp.Type` 自己拼 `.bas/.cls/.frm`（`EXT_FOR_TYPE`）。
- 抓取是**只读**打开（磁盘路径第三个参数 `ReadOnly=True`），运行完工作簿既没被改也没被保存；想留底就由已打开会话 / `pull-open` 用 `SaveCopyAs` 做 `.backup/<时间戳>-抓取前/`。

### 2.3 写回：三种情况，三套写法

```python
vbproj = wb.VBProject
vbproj.VBComponents.Remove(旧模块)       # 同名先删
vbproj.VBComponents.Import(str(文件))    # VBE 读文件、建模块
```

| 情况 | 怎么做 | 为什么 |
|---|---|---|
| 普通模块 / 类模块 / 窗体（Type 1/2/3） | `Remove` 旧的 → `Import` 新的 | 直接 Import 会变成 `模块1`、`模块1 1` 这种重复名 |
| 文档模块 `Sheet1`、`ThisWorkbook`（Type 100） | **不能 Remove**（它绑定工作表，删了工作簿就坏）→ `CodeModule.DeleteLines` 清空 + `AddFromString` 写回；写前 `strip_attribute_lines()` 把 VBE 导出头整块剥掉：`VERSION 1.0 CLASS`、`BEGIN … END`（含 `MultiUse = -1  'True`）、`Attribute VB_*` 及其 `_` 续行。源码只剩导出头时只清空、不 `AddFromString` | 文档模块就是绑定到工作表/工作簿的对象模块，与普通模块、类模块不同，不能 Remove，只能改正文；这些头行只在 `.cls/.frm` 文件里合法，进了模块正文就是语法错误（表现：VBE 顶上多出 VERSION / BEGIN / MultiUse / END 四行） |
| 公共库占位壳 | 直接跳过 | 壳只是一行注释，导进去会把公共库真身覆盖成空壳 |

`.frx` 不用单独处理：`Import` 一个 `.frm` 时，VBE 会自己去同目录找同名 `.frx`。

### 2.4 两个 Excel 实例，两条数据通道

| 通道 | 怎么拿 Excel | 能看到什么 | 用在 |
|---|---|---|---|
| 挂到已有实例 | `win32com.client.GetActiveObject("Excel.Application")` + `ActiveWorkbook` | 你**正在编辑、尚未保存**的内容 | 菜单里的“抓取源码”“写回源码” |
| 自己开隐藏实例 | `DispatchEx("Excel.Application")` + `Visible=False` | 磁盘上的文件 | `pull --all-project` / `push --all-project --apply`（批量走命令行，菜单只做单项目） |

这也是为什么“工作簿已打开”时脚本能抓到未保存的改动——它读的是内存里的工程，不是文件。

实例安全设置（`_new_excel`）：`EnableEvents=False`（不触发 `Workbook_Open`）、`DisplayAlerts=False`（不弹窗）、`AutomationSecurity=1`（**不能用 3**，3 会禁宏导致 `xl.Run` 失败）。

### 2.5 三道权限门槛

| 门槛 | 表现 | 依据 |
|---|---|---|
| 信任中心勾「信任对 VBA 工程对象模型的访问」 | 没勾 → 读 `wb.VBProject` 直接被拒 | 注册表 `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1` |
| 工程未设置密码保护 | 设了密码后连 `VBProject.Name` 都读不到 | VBE 的工程保护 |
| 文件未被占用 | `Workbooks.Open` 或 `Save` 失败 | Excel 自己开着、同步盘中都会占文件 |

脚本用 `has_vbproject_access()` 探测第一条，失败时给的就是"去勾信任中心"这句话。

### 2.6 编码架构：src 存 UTF-8，GBK 只存在于交换边界

VBE 的文本引擎按 **ANSI**（简体中文环境 = GBK）解码导入的文本文件，行尾按 CRLF，直接 `Import` UTF-8 文件会中文变问号、LF 会整段连行——这是 VBE 自身的行为，改不了。但把 GBK 当存储格式代价太大，这是 2026-11 转 UTF-8 的核心原因：

- ① 中文标识符（过程名/模块名/注释）在 grep/ripgrep/IDE 里**完全搜不到**，检索工具与 AI 助手对 `src/` 失灵，引用盘点只能自研 `audit.py`；
- ② 编辑 `.bas/.frm` 必须写脚本显式转码，普通编辑器直接打开会乱码或写坏字节；
- ③ Git diff 在 GBK 下难读，中文改动没法审查。

VBE 的 GBK 偏好只发生在 Import/Export 那一瞬间，这个约束完全可以由工具在边界吸收，所以人和 Git 全程 UTF-8 + LF（agent 与编辑器直接写，不用为行尾转码）。现在的架构把 GBK 压到边界上：

- **src 一律 UTF-8（无 BOM）+ LF**：Git、diff、grep、VSCode、agent 直接可用；
- **抓取（pull）**：`comp.Export` 产出 GBK 文件，`transcode_exported_file()` 立刻按 GBK 读、以 UTF-8 重写为源码文件；
- **写回（push）**：`gbk_import_copy()` 把 UTF-8 源码转成 GBK+CRLF 临时文件（`.frm` 连同同名 `.frx` 一起复制进临时目录），`Import` 用完即删；
- **手工导入**：不要把 `src` 里的文件直接拖进 VBE（LF 会让整段连成一行）；要么用 `python vba-workbench\scripts\sync.py gbk-export --project <项目名> --apply` 的产物（GBK+CRLF），要么等 pywin32 装好后走 `push`。
- **仓库外交付**：`gbk-export` 子命令把 src 镜像导出为 GBK 副本到 `.backup\GBK\<项目名>\<分类>\`（目录压平，不带 src/projects 前缀；`--all-project --apply` 时公共库单独放 `_公共库\`），供手工导入 VBE 或发给别的机器。

读文件走 `read_source_text()`，按 `utf-8-sig → utf-8 → gbk` 依次试——遗留 GBK 文件与手工放进来的 UTF-8 文件都能读。`.frx` 是二进制，任何路径都不转码。

### 2.7 与 VSCode 插件的关系：同一套 API，不同触发器

`excel-vba-sync` 插件本质也是走 COM 访问 `VBProject`，区别只在于它**常驻在编辑器进程里，你保存文件它就推送**；`sync.py` 默认是**你想同步时运行一次**，而 `watch`（菜单里的“常驻自动写回” / `python vba-workbench\scripts\sync.py watch --project <项目名> --apply`）就是脚本自带的常驻版：同样“保存文件即推送”，但只推**改动的那一个模块**，不做整工作簿写回。二者不互相依赖，脚本里没有任何一处读插件配置。

注意：**别让两者同时写同一个工作簿**。插件写的是内存工程，`sync.py` 写回前会 `Remove` 旧模块，先后的覆盖关系取决于谁后运行。

完整的往返只有四步：

```
VSCode 改文件 → 磁盘 src/*.bas →（菜单“写回源码”，默认写回并保存）→ xlsm
xlsm →（菜单“抓取源码”）→ 磁盘 src/*.bas → VSCode 看到
```

---

## 3. 数据流

### 3.1 抓取（xlsm → src）

1. `resolve_target()` 找到要导的 xlsm。
2. 逐项目判定会话路径（`pull --apply`；`--all-project` 每个项目分别判，不整批一刀切）：
   - **工作簿已打开**：`attach_running_excel()` 挂到当前已打开的 Excel 会话（ROT 枚举 / `GetActiveObject` 定位），`require_open_workbook()` 按全路径确认目标工作簿，`ensure_vbe_access()` 过 VBE 关卡；抓内存里的最新代码（含未保存改动），不另开副本、不关窗。`pull-open` 是只借用已打开会话的显式入口，走同一条会话路径。
   - **工作簿未打开**：隐藏实例 `open_disk_workbook(..., read_only=True)` 打开磁盘副本只读抓取，用完 `Close(SaveChanges=False)`。
   - 除 `NO_EXCEL` / `TARGET_NOT_OPEN` 外，附着或会话失败（`COM_CONTEXT` / `EXCEL_BUSY` / `EXCEL_CLOSED` / `VBE_PROTECTED` 等）原样上报，不退回磁盘读陈旧副本。
3. VBE 可读关卡：已打开会话用 `ensure_vbe_access()`（交互式可重试）；磁盘路径用 `has_vbproject_access()` 探测，读不到就给“勾信任中心”提示并跳过该工作簿。
4. 抓取前备份：已打开会话与 `pull-open` 用 `SaveCopyAs` 写到 `.backup/<时间戳>-抓取前/`（`--no-backup` 跳过；改名前历史目录仍叫 `-导出前`，内容一样，可照常回滚）；磁盘抓取走只读 `Workbook.Open`，不动文件。
5. `common_subset_stems()` 算出这个工作簿会用到哪些公共模块。
6. `export_workbook()` 遍历 `VBProject.VBComponents`：
   - 先用 `document_labels(wb)` 取 `{CodeName: 工作表名}`（遍历 `wb.Sheets`，不需要开 VBE）；
   - 公共库模块 → **不抓取正文**，项目里只留占位壳（防两份代码分叉）；
   - 其他模块按 `Type` 决定扩展名和子目录（`EXT_FOR_TYPE` / `TYPE_FOLDER`，因为 `Export` 不会自己补扩展名）；
   - 文档模块的文件名走 `doc_stem(模块名, 工作表名)`，落成 `Sheet4(SKU统计).bas`（见 `sync-config.md` §2.1）；
   - `comp.Export` 产出的是 GBK 文件，写入文件后 `transcode_exported_file()` 立刻按 GBK 读、以 UTF-8+LF 重写（`.frx` 二进制不转码）；
   - 清掉旧的同名文件（含旧位置、`.frx`、无扩展名残留；比较用模块名，所以老的 `Sheet4.bas` 也会被清）；
   - 工程里已删除的模块，对应文件一并删除（**占位壳不删**）。
7. `write_placeholder_shells()` 补齐占位壳。

抓取**不修改也不保存**工作簿。

### 3.2 写回（src → xlsm）

**打开前提（批量抓取/批量写回共用 `open_disk_workbook()`）**：隐藏实例自己打开工作簿，所以工作簿必须**没有打开密码**、**VBA 工程未设置密码保护**；写回还要求文件**不是只读属性**（要 `Save()`）。任一不满足都在“打开”这一步失败并给出可读提示；单个失败不中断整批，`cmd_all` 会记进汇总、`cmd_export` 打「成功 X / 失败 Y」并以 1 退出。

1. **源码预检（不碰 Excel）**：单个 `push` 在确认面板之后、连接/打开工作簿与备份之前先运行 `source_preflight_or_fail()`；`push --all-project` 在每个项目开 Excel 前运行；`run --push` 在打开运行工作簿前运行。失败直接 `SOURCE_PREFLIGHT` 放弃，不备份、不打开工作簿。硬错误：源码目录不存在、没有任何可写模块、同目录模块重名、非占位壳 `.frm` 缺同名 `.frx`、`VB_Name` 与文件名不一致、非文档 `.bas/.cls` 正文为空（`ThisWorkbook / SheetN / ChartN / DialogSheetN / MacroSheetN` 豁免；只有设计没有代码的 `.frm` 不算空）、源码读不出。软提醒：BOM、CRLF、缺 `Option Explicit`、末行以续行符 `_` 结尾/疑似截断、跨目录同名覆盖；提醒只汇总，不阻断。`sync_into()` 默认仍会预检；入口已经预检过时传 `preflight=False`，避免同一轮重复扫描。
2. 依次处理 `source_dirs()`：**先 `src/common`，后项目 `src`**（同名模块后写回的会覆盖前面的）。
3. `resolve_common_subset()` 决定公共库要写回哪些（见 §3.3）。
4. 写项目目录前先 `document_preflight()` 打一次**文档模块预检**（只读）：
   - 源码里有、工程里没有 → 列出名单，说明“工作表可能已删除、本次会逐个跳过”，并提示运行一次抓取清掉文件；
   - 工程里有、源码里没有 → 列出来，明确"写回不会新建"。
5. `import_folder()` 逐个文件：
   - **占位壳** → 跳过（否则会把公共库真身写成空壳）；
   - 同名 `.bas/.cls`（Type 1/2/3）→ **原地改正文**（`replace_module_code`），不 Remove、不 Import，避免 VBE 补序号；`.frm` / 新模块 / 类型不符才走 `import_component`，Import 前用 `gbk_import_copy()` 把 UTF-8 源码转成 GBK+CRLF 临时文件（`.frx` 与 `.frm` 同目录），整轮结束删除临时目录；
   - **文档模块**（Type 100，或名字像 `SheetN`）→ 不能 Remove，用 `CodeModule.DeleteLines` + `AddFromString`（写入前 `strip_attribute_lines()` 去掉 `Attribute VB_` 头）；默认跳过，除非 `--doc-modules`；
   - 文档模块在工程里**找不到**时只 warn 跳过，**绝不新建工作表/模块**（`overwrite_document_module` 返回 False）。
6. 每个目录写完，`report_references_for()` 做一次**静态引用比对**：扫描本次写回文件的 `As ADODB.xxx` 之类早期绑定，对照工程 `References`，缺就警告。工程读不到引用时不报（防误报）。
7. 全部写完后 `verify_component_set()` 做**模块清单校验**：把写回前/后的 `{模块名: 类型}` 快照对比，打印“一致 ✔”；一旦有新增或消失就告警——写回只该改正文，不该增删模块（预览模式会跳过这一步）。
8. 再跑 `postcheck_sync()`：对本次实际写回的模块逐个回读，先剥导出头与过程级 `Attribute`、去空行；VBE 侧再抹掉 `AddFromString` 在末尾补的游离 `()`，之后按 `fold_lines` 指纹比对源码；不一致就 `POSTCHECK_FAILED`，明细形如 `模块（源码 4 行 / VBE 2 行；首个差异第 3 行）`；**不保存、不自动回滚**。
9. 非演练时先算**缺壳**（`missing_shells()`：要写回但项目目录没有同名壳的公共模块）→ 逐条打印「模块名（依据：…）」→ 问一句要不要补壳，回车补 `create=True`；随后再刷新**已存在**的占位壳（`create=False`，写回本身绝不凭空多文件）。
10. 单个写回默认保存（菜单“写回源码”与 CLI 都是）；`--no-save` 只改内存，`cmd_all` 会自己 `wb.Save()`。
   选中哪个工作簿由**配置解析出的全路径**决定：`require_open_workbook()` 在已打开的 `Workbooks` 里按全路径匹配并 `Activate()` 到前台；**匹配不到直接 `Failure`**（附带“当前打开了哪些”），绝不退回 `ActiveWorkbook` 写进别的工作簿。`pull-open`（以及 pull 的已打开会话路径）用同一套定位，避免“把 A 工作簿的代码抓进 B 项目”。

### 3.3 公共库的选择性写回（`common_modules: "auto"`）

`common_reference_reasons()` 出判定依据，`referenced_common_modules()` 只是取它的键集合。四条依据：

| 依据 | 判定 | 说明 |
|---|---|---|
| `模块名` | 项目代码里出现该模块名（标识符级匹配） | 最硬的一条，例如 `Sheet14.bas` 里写了 `表格工具.initTable` |
| `过程名 X` | 项目代码里出现过程名 `X`，且 `X` **只在公共库这一个模块里声明**、**项目自己没有同名声明** | 项目自己声明过的同名过程属本地过程，不算引用公共库（`示例表B` 自己的 `test1` 就是这样被排除的） |
| `闭包←Y` | 已被选中的公共模块 `Y`，用它自己的代码按前两条再牵连出别的模块 | 例如 `表格工具.cls` 真的调用了 `自定义通用函数` 的过程 |
| `占位壳` | 项目目录里**已存在**该模块的同名壳文件（只取文件名，不读壳里的注释） | 换壳后正文没了，引用证据也一起没了——不加这条，该工作簿会被静默移出写回清单，工作簿里继续留着旧副本，公共库再改就分叉 |

四条硬约束（都是踩坑后加的）：

1. **先剥注释、再抹字符串字面量**（`strip_vba_comments` → `mask_strings`）：VBA 的 `' 之后`、`Rem` 行都不算引用；双引号里的名字也**不算引用**，因为那是软调用。
   - 反例 A（注释）：`表格工具.cls` 第 67 行注释写「改本类必须重跑 测试表格工具.bas」，早期版本把注释当代码，直接把 `测试表格工具` 拉进写回清单，还在项目目录里补了个壳。
   - 反例 B（字符串）：`自定义通用函数.getTL` 里 `Call Application.Run("初始化数据同步配置", getTL)` 是个**软依赖**（工作簿里没导入 `数据同步.bas` 就静默跳过）。早期把字符串当硬引用，闭包就把整个 `数据同步` 灌进根本不连库的 `示例表D`，于是每次新建工作表对象都弹「初始化表格【数据同步配置】出错了」。2026-11-18 起软调用不再算依据。
2. **传递闭包也剥注释 + 抹字符串**：公共库模块头部的说明性注释、以及互相之间的软调用，同样会滚雪球。
3. **标识符边界用 `\w`（含中文）而不是 `[A-Za-z0-9_]`**：否则 `解析` 会命中 `解析字符串`、`ID` 会命中 `ID列升序`，又变成误判。
   不使用任何硬编码"黑名单"——判的是"有没有证据"，不是"名字够不够长"。
4. **壳只当"名字证据"，绝不解析其正文**：壳里那行 `真实文件在：src/common/...` 本身带着模块名，若当文字证据读，会让"引用"永远成立、只增不减（写回清单再也不会收缩）。所以壳只贡献 `f.stem` 一个标识符。
   实测：`示例表A` 换壳 3 个模块后，加这条前从 6/7 掉到 5/7（`测试表格工具` 被静默移出）；加这条后恢复 6/7，`示例表B` 仍是 4/7。

带 `COMMON_VERSION` 常量的模块（`common_version_stems`）**始终写回**，理由记为"版本常量（强制写回）"。

三种模式：缺省 = 整库写回旧的；`"auto"` = 按引用挑；给列表 = 白名单（库里不存在的名字会警告并忽略）。白名单不是黑名单的变种：它是“这个工作簿就要这几个”，与判定逻辑无关。

**交互兜底**：菜单“写回源码”选中后自动预览，会把 auto 挑出的模块和**每个模块的依据**逐行打出来；确认无误再回车执行，输入 `0` 返回，一个字节都不写。批量写回用 CLI（`push --all-project --apply`），预览只打印清单不打断脚本；`--apply` 才真正写入。

**缺壳提示（`missing_shells()`）**：写回到一半会列出「本次要写回、但项目目录里还没有同名占位壳」的公共模块，每个都带判定依据，并问一句要不要补壳（回车 = 补，写进项目目录、记得 git 提交）。
壳是「这个工作簿声明要用它」的唯一人工可见依据，所以：**有壳才在换壳/对齐之后仍留得住；删壳等于这个工作簿不需要它**（软调用已不算依据，只有真引用才可能把模块拽回来）。

### 3.4 占位壳（项目目录里"看得见"公共模块）

- 位置：项目里**原分类目录**、**同名同后缀**，如 `类模块/表格工具.cls`；
- 内容**只有一行**：`' 此文件是公共库类模块，真实文件在：src/common/类模块/表格工具.cls`；
- 编码 UTF-8（无 BOM）+ LF，与真实源码一致；
- 识别（`is_shell_file`）：扩展名属于源码类 + 去空白后以 `'` 开头 + 含 `真实文件在：` + 整个文件只有一行。

三个消费者必须都认它：写回时跳过、抓取公共模块时不写正文、`auto` 判定引用时排除。

**生成规则**（`write_placeholder_shells`）：

| 情况 | 行为 |
|---|---|
| `auto` 判定这个工作簿引用了它 | 写壳 |
| **刚被 promote 搬走的模块** | 强制写壳（`always` 参数）——它已经不在项目源码里，`auto` 会判定"没引用"，不强制就会漏 |
| 项目目录里已经有它**已经存在的壳** | 保留并重写（不会被当成"多余"删掉，否则 promote 刚生成的壳下次刷新就没了） |
| 公共库里已经没有这个模块了 | 删掉这个壳（死壳） |
| 项目里还有同名**真实代码** | 只警告，绝不覆盖 |
| `shells --mode clean --apply` | 只删不建 |

早期版本按"`auto` 子集之外一律删除"处理，导致 promote 完立刻报"占位已生成"、实际却没文件（模块自己被搬走，谁都不再引用它）。现已改为：**壳一旦存在就保留，只有死壳或 `shells --mode clean --apply` 才清理**。

**把“同名旧副本”换成占位壳**（`align_with_common`，菜单“换公共库占位壳” / `align`；下文简称“换壳”）：

| 步骤 | 行为 |
|---|---|
| 1 | 选一个工作簿（`_pick_projects(..., multi=False)`，列全部已登记项目，带已打开/可用标签） |
| 2 | 扫 `src/projects/<项目名>` 下所有 `.bas/.cls/.frm`，`stem` 能在 `module_index()` 里找到的就是**同名** |
| 3 | 已是壳的直接计入"已对齐"，不重复问 |
| 4 | 真实代码逐个提示：项目行数/公共库行数、两边修改时间、公共库相对路径 |
| 5 | 答 `y` → 先把原文 `shutil.copy2` 到 `.backup/<时间戳>-换壳/<项目名>/<原相对路径>`，再把该文件覆写成一行壳；答 `N`（或回车）→ 保留不动 |
| 6 | 问完即结束——**不补壳、不新建任何文件** |

**硬规则（规则层面，不可破）**：换壳路径只做"已有正文 → 壳"的替换，**绝不新增任何文件**。

- 早期版本在问完之后调了 `write_placeholder_shells()` 补齐，结果按 `auto` 判定凭空多出一个 `测试表格工具.bas` 壳（`示例表B` 的工作簿里根本没这个模块）。现在 `cmd_align` / `menu_align` 都不再调它。
- 需要新建壳的场合只有三个，且都显式：`promote` 命令把模块搬进公共库后留壳、`shells` 命令、抓取时为公共库模块留占位。
- `write_placeholder_shells` 本身**不会**覆盖真实代码（安全保护），所以"换壳"只能由 `align` 这条显式、逐个确认的路径来做。
- 判断"哪边更新"没有可靠办法，提示里只给行数 + 修改时间，最终由人决定——**项目侧行数明显更多时，多半是项目里改过，先 promote / 对比，别急着换壳**。

### 3.5 统一提醒：一次操作里的"预检/校验/警告/失败"都要落到汇总

**实现位置**：收集器与判定在 `toolkit/check.py`（`Notes` / `check_module_set` / `check_doc_modules` / `check_export_set`），
`sync.py` 只留三件事——读 `VBProject`、把结果塞进收集器、打到控制台。`audit.py` / `agent.py` 的风格检查用的是同一份（见 `toolkit.md`「统一校验」）。

**为什么有这一层**：写回、抓取、运行宏、自动写回四条链路各自都会发现问题，散在流程中间容易被刷掉。现在统一成“边执行边收集、收尾集中提示”：`remind(level, text)` 只收集，命令结尾 `notes_summary()` 一次性打出来，同样的提示合并成 `×N`（自动写回一轮轮重复不会刷屏）。

```python
reset_notes()                       # 命令开头清空
...
remind("预检", "…")                  # 中间记一条（level ∈ 失败 / 警告 / 预检 / 校验）
...
notes_summary("写回提醒")            # 结尾统一打印
```

输出形如：

```
── 写回提醒：预检 2 · 校验 1 ──
  [预检]
     · 2 个文档模块在工程里不存在（工作表已删除）：Sheet22、Sheet23
  [校验]
     · 模块清单与写回前一致（19 个）✔
```

没有异常时明确打 `── 写回提醒：无异常 ──`，避免"没输出=没检查"的误判。

**各链路的关卡一览**（改代码时按下表对齐，别只加流程不加检查）：

| 动作 | 写前 | 写后 |
|---|---|---|
| `push` / `push --all-project --apply` / 菜单“写回源码” | `source_preflight`：重名/`VB_Name`/读不出；`document_preflight`：源码有/工程没有的文档模块、工程有/源码没有的工作表模块 | `verify_component_set`：模块集合必须原样；`postcheck_sync`：逐模块回读正文，不一致不保存 |
| 自动写回 `watch`（单模块） | 文档模块先确认在工程里存在，工作表已删除就直接说清并跳过；公共库子集外的模块跳过 | `verify_component_set`（每个模块写回后立即比一次）；`postcheck_sync` 回读比对 |
| 抓取 `pull` / `pull-open` / 菜单“抓取源码” | （读表，无写） | `verify_export_set`：工程里每个模块都要在源码目录落成文件（公共库占位不计） |
| 运行宏 `run` | 带 `--push` 时在打开运行工作簿前先跑源码预检，再备份/写回；预检失败写 `.run` journal 后退出 | 宏异常时记一条"失败"，结尾汇总；日志尾部照旧打印 |

**边界**：`verify_component_set` / `verify_export_set` 都依赖读得到 `VBProject`（信任中心允许访问 + 工程未设密码保护），读不到时静默跳过，不假装通过。

**应用状态还原**：`run` 运行完宏、`push` 写完工程，收尾各过一次 `toolkit.app_guard`——把事件响应 / 屏幕刷新 / 警告提示 / 状态栏打开、计算模式钉回**手动**、清掉残留 StatusBar 与剪切状态；真的改过才打印一行「已还原 Excel 状态：…」。
这是“界面点哪都不响应”“增删数据一直正在计算”这两类后遗症的兜底，宏内部仍要自己成对（`开启加速`/`关闭加速`）。

---
