# toolkit 使用与维护说明

`vba-workbench/scripts/toolkit/` 是 `sync.py`（同步工具）和 `audit.py`（代码盘点）**共用的通用层**。
一句话：**入口脚本只写业务，凡是"跟具体业务无关、两个脚本都可能用到"的能力，一律放进 toolkit，不许在入口脚本里再写一遍。**

## 目录

- [1. 结构](#1-结构) · [2. 每个模块管什么](#2-每个模块管什么)（含编码、备份清理、规则检查、COM 边界、统一校验）
- [3. 交互组件](#3-交互组件的正确用法) · [4. 下沉规则](#4-什么该下沉到-toolkit什么不该)
- [5. 新增能力步骤](#5-加新能力时的步骤) · [6. 入口脚本合规状态](#6-现有入口脚本的合规状态) · [7. 常见坑](#7-常见坑)

新看这个仓库时，记住三条就够：

1. 我在 `vba-workbench/scripts/sync.py` / `vba-workbench/scripts/audit.py` 里看到 `input()`、手写的编号循环、`open(..., encoding=...)`、`print()`、`sys.exit()` —— **基本就是写错了位置**，应该换成 toolkit 的基础函数。
2. 要加新能力，先问：它跟"VBA 源码同步/盘点"这个业务有关吗？无关 → 进 toolkit；有关 → 留在入口脚本，只调用 toolkit。
3. 改编码、换行、路径、目录结构这类"全局规则"，只在 toolkit 里改一次；入口脚本永远跟着 toolkit 走。

---

## 1. 结构

```
vba-workbench/
├─ SKILL.md              # skill 入口
├─ VERSION               # 版本号（doctor 校验）
├─ assets/
│  └─ mapping.example.json   # 配置样例，供 init 或复制
├─ scripts/
│  ├─ sync.py            # 业务：mapping.json、Excel/VBE 往返、watch 增量、菜单分支
│  ├─ audit.py           # 业务：成员引用、依赖矩阵、结构体检
│  ├─ agent.py           # 业务：Agent 统一入口（风格/复杂度/断言/写回）
│  ├─ 启动菜单.bat        # 双击入口（必须带 -i）
│  └─ toolkit/           # 通用层（本文件讲的就是它）
│     ├─ __init__.py     # 统一导出；规则摘要写在顶部 docstring
│     ├─ base.py         # 代码根、配置查找、workspace 基准、编码表、Failure、输出、路径
│     ├─ config.py       # mapping.json 读写与校验
│     ├─ files.py        # 写源码/转码/备份/找最近备份/移动/追加日志/删临时目录
│     ├─ repo.py         # src 目录、模块索引、占位壳、版本常量、文件名↔模块名
│     ├─ check.py        # 提醒收集器 + 模块集合/文档模块/抓取完整性对账
│     ├─ vba_parse.py    # 剥注释/字符串屏蔽/导出头剥离/成员表/代码规范化
│     ├─ style_rules.py  # 代码规则表（命名/注释/裸 New/加速）
│     ├─ excel.py        # Excel/VBE 的 COM 边界层
│     ├─ watch.py        # 文件快照 / 落点 / 连续两轮确认
│     ├─ app_guard.py    # Application 状态还原
│     ├─ watchdog.py     # 线程看门狗基础函数 / PID 三态存活探测
│     ├─ dialogs.py      # 弹窗 / 模态框识别族（枚举 + 纯函数分类）
│     ├─ ui.py           # 排版、读行、编号选择、菜单骨架
│     ├─ vbabin.py       # 只读解析 vbaProject.bin
│     └─ cli.py          # -i 开关、退出码、异常兜底
├─ references/           # sync 四册 + audit.md、agent.md、toolkit.md、usage-examples.md
```

单测与内部评测只在完整 `main` 的 `tests/` 与 `evals/`，不随安装包或公开快照。

运行期 `src/`（源码真源）与 `mapping.json` 在工作仓库根目录，不在本包内；包只带 `assets/mapping.example.json` 样例。

依赖方向严格单向，不能反向、不能循环：

```
base ─┬─ config
      ├─ repo ────────┬─ ui ── cli
      ├─ files        │
      ├─ vba_parse    │
      ├─ check        │
      ├─ vbabin       │
      ├─ watch        │
      └─ excel ───────┘   （excel 还用 files / repo / ui）
```

- `sync.py`、`audit.py`、`agent.py` **都只 `from toolkit import ...`**，彼此**不 import**。
- 子模块之间只能按上面的箭头依赖（`base` 不许用任何兄弟模块）。
- **`excel.py` 顶层只许标准库**：`win32com` / `pythoncom` 一律写在函数里。
  这样没装 pywin32、没有 Excel 的机器上 `import toolkit` 照样成功，`audit` / `style` 不受影响。
- `__init__.py` 是唯一对外门面：新增公开函数，记得同时加到 `from ... import` 和 `__all__`。

---

## 2. 每个模块管什么

| 模块 | 管的事 | 主要 API |
|---|---|---|
| `base` | 代码根 / 配置查找与 workspace 基准 / 源码扩展名 / 编码兜底顺序 / `Failure` / 输出 / 路径解析 / 读源码文本 | `CODE_ROOT` `ROOT`(workspace) `CONFIG_PATH` `CONFIG_ENV` `find_config` `SOURCE_EXTS` `Failure` `msg` `warn` `fail` `abs_path` `read_source_text` |
| `config` | `mapping.json` 加载（结构不对直接抛 `Failure`；旧 tools/ 位置会提示迁移）与写回 | `load_config` `load_raw_config` `save_config` `config_source` |
| `files` | 所有写入文件动作与编码落地 | `write_source` `transcode_exported_file` `gbk_import_copy` `write_text_utf8` `dump_json` `stamp_dir` `backup_copy` `backup_dirs` `backup_stats` `dir_size_bytes` `prune_backups` `latest_backup`（按备份目录名取最近一份，别按 mtime） `move_into` `append_line` `newest_file` `read_tail` `remove_tree` |
| `repo` | 只读：src 结构、模块索引、占位壳、版本常量、**文件名↔模块名换算**、**项目→xlsm 解析**、VBE 组件类型编号、**代码规则约定** | `project_dirs` `source_files` `source_files_at` `files_under` `rel_path` `module_index` `is_shell_file` `shell_text` `read_common_version` `find_source` `module_name_of` `doc_label_of` `doc_stem` `resolve_target` `parse_stem_version` `MODULE_GROUP_ORDER` `VBXT_STANDARD/CLASS/FORM/DOCUMENT` `IMPORTABLE` `EXT_FOR_TYPE` `TYPE_FOLDER` `RULE_ALLOW_NEW_FILES` `RULE_ALLOW_ACCEL_FILES` `RULE_NEW_MARK` `RULE_EXEMPT_MARK` `IMPLICIT_PROCS` |
| `check` | 通用校验：提醒收集与汇总、模块集合对账、文档模块对账、抓取完整性判定、**后缀模块判定**（`X1` 与 `X` 并存且源码里没有 `X1`）（**纯集合/纯文本，不碰 Excel**） | `Notes` `NOTE_LEVELS` `ModuleSetDiff`（`added` / `lost` / `suffix`） `check_module_set` `check_doc_modules` `check_export_set` `check_suffix_modules` |
| `vba_parse` | 纯文本处理，不懂业务；含**代码规则扫描**（裸 New／手写加速／缺过程注释） | `split_lines` `code_lines` `split_code_comment` `mask_strings` `strip_export_header` `parse_members` `is_event_proc` `norm_code_lines` `fold_lines` `strip_vba_lines` `declared_procs` `scan_rules` |
| `excel` | **COM 边界层**：起实例、连已开实例、按路径找工作簿、打开工作簿、识别工程状态（受密码保护/没授权/只在源码）、工程组件查找与增删替换、整工作簿备份 | `new_excel` `quit_owned` `ExcelSession` `attach_running_excel` `reset_rot_unreachable` `com_call` `find_open_workbook` `require_open_workbook` `open_disk_workbook` `backup_workbook` `vbe_state` `vbe_state_hint` `has_vbproject_access` `ensure_vbe_access` `find_component` `component_snapshot` `remove_module` `replace_module_code` |
| `watch` | **纯 stat** 的文件监听：快照、文件落在哪个源码根、连续两轮确认后才算"改完了" | `snap`→`watch_snapshot` `root_of`→`watch_root_of` `changed`→`watch_changed` |
| `vbabin` | 只读解析 xlsm 里的 `vbaProject.bin`：列模块、解压模块正文、比对两份的模块集合与正文（不碰 Excel、不需要工程密码） | `read_vba_bin` `vba_models` `vba_module_names` `vba_module_sources` `compare_vba_modules` |
| `app_guard` | **Application 状态还原**：运行宏/写回前后把事件响应、屏幕刷新、警告提示、状态栏、计算模式、剪切状态恢复为可用状态（默认钉手动计算） | `guard`（`with` 用） `AppGuard`（`changes` / `describe()`） `restore`（只还原） `calc_mode` |
| `ui` | 终端排版 + 所有键盘交互 | `SEP` `section` `grid` `read_line` `pause` `choose` `ask_yes` `run_action` `pick_list` `pick_list_default` `pick_grouped` `run_menu` |
| `watchdog` | **与 Excel/COM 无关**的并发与进程探测骨架：守护线程硬超时、三态存活探测、句柄反查 PID；不抛业务失败码，错误码分类由上层编排层完成 | `WatchdogTimeout` `run_with_watchdog` `pid_alive` `pid_from_hwnd` |
| `dialogs` | 弹窗 / 模态框识别族：窗口快照纯数据 + 纯函数分类（编译错误 / VBE 通用模态 / Excel 模态 / 未知）+ PID 过滤与优先级；真实 win32gui 枚举只在枚举函数内延迟导入 | `WindowInfo` `DialogMatch` `KIND_COMPILE_ERROR` `KIND_VBE_DIALOG` `KIND_MODAL_DIALOG` `KIND_UNKNOWN` `classify_window` `detect_dialog` `enumerate_windows` `detect_target_dialog` |
| `cli` | 命令行外壳 | `add_interactive_arg` `add_no_input_arg` `ensure_interactive` `run_cli` `entry` |

### 编码规则（只在 `files` 落地一次）

| 场景 | 规矩 | 用谁 |
|---|---|---|
| `src/` 里的 VBA 源码 | UTF-8 无 BOM + **LF** | `write_source` |
| VBE Export 出来的文件 | ANSI/GBK + CRLF → 就地转成上面的格式 | `transcode_exported_file` |
| 写回 VBE 之前的临时副本 | GBK + **CRLF**（VBE 按 ANSI 解码、按 CRLF 断行） | `gbk_import_copy` |
| 报告 / JSON / 日志 | UTF-8 无 BOM + LF | `write_text_utf8` `dump_json` `append_line` |

> 入口脚本里不要再出现 `encoding="gbk"`、`newline="\r\n"`、`open("a")` 这类字面量——要改规则只改一处。

### 备份批次的识别与清理（`files` 的四件套）

| 函数 | 作用 | 说明 |
|---|---|---|
| `backup_dirs(根)` | 扫 `<YYYYmmdd-HHMMSS><tag>` 备份目录，按时间升序返回 `(目录, 时间, tag)` | 只认时间戳开头的目录；`.backup/GBK`、`.backup/watch` 这类固定目录**不算备份批**，跳过 |
| `backup_stats(根)` | 体检用：`{count, bytes, by_tag}` | 批次数 / 总体积 / 每类来源的批次数 |
| `dir_size_bytes(目录)` | 目录内文件字节数和（不存在算 0） | 用于算"能释放多少" |
| `prune_backups(根, keep_recent=10, keep_daily=1, min_keep=3, dry_run=True)` | 备份自动清理 | 三条保留规则取并集：每 tag 最近 N 批 / 每 tag 每天最新 1 份 / 全局最新 3 批；`dry_run` 默认 True，真要删得显式传 False |

**边界**：toolkit 只负责"哪些批次该删、删掉能省多少"；删不删、怎么提示、要不要先预览，由调用端（`sync.py` 的 `cmd_clean_backup`）决定。

### 代码规则检查的分工（2026-11-18 起）

四条“改代码时最容易悄悄退回去”的规则，判定只有一份，按层分开放：

| 层 | 放什么 | 具体 |
|---|---|---|
| `repo` | 约定常量 | `RULE_ALLOW_NEW_FILES`（允许裸 `New 表格工具`）、`RULE_ALLOW_ACCEL_FILES`（允许手写开关）、`RULE_NEW_MARK`、`RULE_EXEMPT_MARK`（例外记号）、`IMPLICIT_PROCS`（`class_initialize` 等隐式入口） |
| `vba_parse` | 扫描 | `scan_rules(files)` → `{naked_new, manual_accel, no_comment}`，元素 `(文件, 行号, 片段)` |
| `check` | 判定 | `check_suffix_modules(comp_names, src_names)` → `[(后缀名, 原名)]`；`check_module_set` 顺手把它放进 `ModuleSetDiff.suffix` |
| `audit` / `sync` | 只在编排与打印 | 报告文案、`remind()`、Markdown 写入文件都留在入口脚本 |

**边界**：toolkit 只回答"算不算违规、违规名单是什么"；分组、措辞、要不要写报告属调用端。
新增检查时先往 `scan_rules` 里加一条规则，不要在 audit 或 sync 里另写正则。

### 文档模块的文件名（`repo` 的 `doc_*` 三件套）

`Microsoft Excel 对象/` 下的源码文件名带工作表名，和 VBE 左侧显示对齐：

```
Sheet1(汇总表).bas   Sheet4(SKU统计).bas   Sheet15(配置表).bas   ThisWorkbook.bas
```

**文件名带括号，模块名不带。** 三个函数各管一段，脚本里不许自己写正则：

| 函数 | 作用 | 例子 |
|---|---|---|
| `doc_stem(模块名, 工作表名)` | 生成文件主干（抓取时用） | `doc_stem("Sheet4", "SKU统计")` → `Sheet4(SKU统计)`；工作表名为空**或只是个编号**（`Sheet10`/`Chart3`）则原样返回；非法字符 `: * ?` 换成 `_` |
| `module_name_of(路径/文件名)` | 从文件名取模块名（**所有比较都用它**） | `Sheet4(SKU统计).bas` → `Sheet4`；`工具(旧).bas` → `工具(旧)`（普通名字不动） |
| `doc_label_of(路径/文件名)` | 取工作表名（只用于显示/判断“已带标签”） | `Sheet4(SKU统计).bas` → `SKU统计` |

规矩：

- 判断"这是哪个模块"→ `module_name_of(f)`，**不要** `f.stem`。`module_index()` 和 `find_source()` 内部已经换过来了。
- 只有"文档模块"才带标签：`ThisWorkbook` / `SheetN` / `ChartN` / `DialogSheetN` / `MacroSheetN` / `ModuleN`（`repo.DOC_MODULE_RE`）。普通模块、类模块、窗体一律原名。
- 老布局 `Sheet4.bas` 仍然能读；抓取时按新名写入文件并顺手删掉旧的同名文件，所以仓库里不会两种写法长期并存。
- `find_source("Sheet4")` 传模块名即可，传带标签的文件名也认。

### vbaProject.bin 只读核对（`vbabin`）

写回 / 抓取之后想确认"模块有没有被删、有没有多出来"，**不要打开 Excel 看，也不要依赖工程密码**：
xlsm 就是 zip，`vbaProject.bin` 是个 OLE 容器，每个 VBA 模块在里面是一条流，流名就是模块名（中文原样可读）。

```python
from toolkit import vba_module_names, compare_vba_modules

vba_module_names(r"<表目录>\示例表A V<数字>.xlsm")
# -> ["CommonVersion", "Sheet1", ... "表格工具"]

diff = compare_vba_modules(写回前的备份, 写回后的表)
diff["same"]          # 模块集合是否完全一致
diff["lost"]          # 消失的模块（这才是要命的）
diff["count_before"], diff["count_after"]
diff["cache_before"], diff["cache_after"]   # __SRP 编译缓存数，每次重编译都会变，不算异常
```

命令行（在 `vba-workbench/scripts/` 目录下）：

```bat
python -m toolkit.vbabin 表.xlsm              :: 列模块清单
python -m toolkit.vbabin 前.xlsm 后.xlsm      :: 比对模块集合
```

日常更省事的入口是菜单“核对模块清单与正文” / `python vba-workbench\scripts\sync.py vbamod --project 示例表A`：
它自动到 `.backup` 里找这个工作簿**最近一份备份**（按备份目录名的时戳排，不按文件 mtime——
`copy2` 会保留源文件时间，按 mtime 会选错），然后打印模块数、缓存数、消失/新增清单。

能力：模块集合 + 模块正文。它按 MS-OVBA 2.4.1.3 解压每个模块的 `CompressedSourceCode`，统一 CRLF / CR 并剔除 `Attribute` 元数据行后比较，所以能回答“哪个模块的代码变了”；解压失败时只报“正文未比对”，不会把未比对当成“一致”。

### Excel / VBE 的 COM 边界（`excel`）

凡是"碰 Excel 本身"的动作都在这里，业务脚本只决定**做不做、做什么**，不再自己写 `try/finally` 和重试：

| 场景 | 用谁 |
|---|---|
| 要一个看不见的 Excel 实例做批量活 | `with ExcelSession() as s:` → `s.app()`；退出自动 `quit_owned()` |
| 回收工具自己起的实例（工作簿可能是 dirty） | `quit_owned(xl, *, save_changes=False, keep_workbooks=None)`：先 `DisplayAlerts=False`、按白名单逐簿 `Close(SaveChanges=False)`、再 `Quit`；返回未清干净的诊断列表，空=干净。归属快照不一致或名单外仍有簿时不退出实例，只上报诊断（见下两行） |
| 归属快照对不上（PID/Hwnd 与起实例时不一致） | `quit_owned` 收尾前比对 `new_excel` 登记的 pid/hwnd 快照；不一致则不碰簿、不 Quit，问题清单记「归属校验不通过」。无快照的旧调用路径（attach 借来实例、桩对象）保持旧行为，不因缺快照报错 |
| 白名单模式下名单外仍有簿 | `keep_workbooks` 列出必须保留的簿名（宏在隔离实例里自开的参考簿）；关完后实例里仍有残留簿则 **不 Quit**，把剩余簿名写进问题清单——宁可留实例，也不吞掉别人数据。`keep_workbooks=None` 为旧行为：关全部簿后退出 |
| 连用户正在用的那个实例 | `attach_running_excel()`；ROT 里有目标但 `BindToObject` 失败时按 `COM_CONTEXT` 报，不降级成“没打开”；`reset_rot_unreachable()` 显式清空不可达记录（测试隔离 / 重入主动清场） |
| 按路径在已打开的工作簿里找它（找不到就报错并列出当前打开的工作簿） | `require_open_workbook(xl, path, 项目名, 用途)` |
| 从磁盘打开（可 `read_only=True`） | `open_disk_workbook(xl, path, read_only=...)` |
| 工程读不了，要分清“受密码保护 / 没授权 / 只在源码” | `vbe_state(wb)` → `(状态, 说明)`；`vbe_state_hint(状态)` → 用户的下一步动作 |
| 卡住“必须先解除密码保护才能继续”，允许用户当场去解除 | `ensure_vbe_access(wb, 标题)`（带回车重试） |
| 找模块 / 前后模块清单 | `find_component(vbproj, 模块名)` / `component_snapshot(vbproj)` |
| 删旧模块、覆盖文档模块正文 | `remove_module(vbproj, 名)` / `replace_module_code(comp, code)` |
| 整工作簿备份 | `backup_workbook(wb, 备份根目录, tag="-自动写回")` |

两条硬规矩：

- `component_snapshot` 只给"写前/写后比对用"，**不要**拿它当"工程里有哪些模块"的权威清单（它读的是打开的工程）。
- 文档模块只能 `DeleteLines` + `AddFromString`（`replace_module_code` 干的就是这个），**不能** `Remove` 再导入，否则会多出 `VERSION 1.0 CLASS / BEGIN MultiUse / End` 这种头。
- `replace_module_code` 会把正文里的 `Attribute` 行**一律丢掉**（`_drop_attribute_lines`）：过程级属性如
  `Attribute call_打印信息.VB_ProcData.VB_Invoke_Func = "I\n14"` 只属于导出文件，写进 `CodeModule` 就是语法错误
  （2026-11-18 实测整簿编译不过，VBE 卡在出错的模块上）。导出头另由 `vba_parse.strip_export_header` 处理。
  副作用：原地覆盖不改组件属性，Excel 内置的宏快捷键（`VB_Invoke_Func`）不会因此丢失；但若该模块被 Remove 后重新 Import，
  这类绑定不会随代码恢复，需要在「宏」对话框里重设或走项目自己的 注册快捷键。

### 运行完宏的应用状态还原（`app_guard`）

宏跑到一半出错，或自己忘了成对复位，会把 `Application` 留成半关闭状态——现象就是**界面点哪都不响应**、
**增删数据右下角一直"正在计算"**。脚本侧统一还原一次：

```python
from toolkit import app_guard

gd = app_guard.guard(app)          # 进：记快照
with gd:
    app.Run("'表.xlsm'!模块.过程")
if gd.changes:                     # 只打印真的改过的项
    print("已还原：", gd.describe())

app_guard.restore(app)             # 只还原、不进快照（出错收尾用）
app_guard.calc_mode(app)           # "手动" / "自动" / "读不到（…）"
```

| 项 | 还原成 | 为什么 |
|---|---|---|
| `EnableEvents` / `ScreenUpdating` / `DisplayAlerts` / `DisplayStatusBar` | `True` | 关着就等于界面哑、屏幕不刷、错误静默、进度条不显示 |
| `Calculation` | **手动(-4135)** 默认；传 `calculation="auto"` 才改自动 | 本库规则：工作簿长期手动，避免增删数据全表重算 |
| `StatusBar` / `CutCopyMode` | 清掉 | 残留提示与"复制中"虚线框会误导用户 |

规矩：**凡是脚本侧运行完 VBA 宏（`sync.py run` / `watch`）或写回工程（`push`），收尾必须过一次 `app_guard`**，
不要指望每个宏自己收尾；宏里当然也要成对（`开启加速`/`关闭加速`）。

### 文件监听（`watch`）

自动写回那种"源码一改就干活"的循环，三件事别自己写：

```python
from toolkit import watch_snapshot, watch_root_of, watch_changed

base = watch_snapshot(folders)                 # 启动时的基线
pending = {}                                   # 跨轮保留
while True:
    cur = watch_snapshot(folders)
    for f in watch_changed(cur, base, pending):   # 连续两轮一致才算"改完了"
        write_back(watch_root_of(f, folders), f)
        base[f] = cur[f]                          # 写完把基线抬到当前
```

- 快照用 `mtime_ns + 大小`，不是哈希：每轮只做一次 `stat`。
- 文件清单走 `repo.source_files`（递归、扩展名、跳过占位壳都在里面）。
- `watch_root_of` 决定“这个文件该写回哪个工作簿的那半边源码根”；公共库文件落在 `common_dir` 时要按项目子集再过滤一遍——**子集是业务规则，不进 toolkit**。

### 统一校验（`check`）

凡是"写之前核对 / 写完再验一遍"的动作，判定和提醒格式都只有这一份，别在业务脚本里另写。
它**不碰 Excel**：需要工程数据的调用方自己读成 `{模块名: 类型}` 或名字列表再传进来。

| API | 用途 |
|---|---|
| `Notes(level, text)` → `add/merge/extend` | 一次操作的提醒收集器；`render()` 出行列表（同文本合并 `×N`，空则"无异常"），`has("失败","警告")` 可直接决定退出码 |
| `ModuleSetDiff` + `check_module_set(before, after)` | 写回前后模块集合对比：`added` / `lost` / `ok` / `notes()` |
| `check_doc_modules(src_docs, comp_names)` | 文档模块对账，返回 `(孤儿, 未抓取)`：源码有工程没有（工作表已删除）、工程有源码没有（抓漏了） |
| `check_export_set(module_names, file_names, skip)` | 抓取完整性：工程模块是否都落成文件；`skip` 传公共库模块（项目里只是占位壳） |

分级固定为 `失败 / 警告 / 预检 / 校验`，汇总行形如 `── 写回提醒：预检 1 · 校验 1 ──`。

**三个入口怎么用**（规则必须一致，别各判一遍）：

| 入口 | 用法 |
|---|---|
| `sync.py` | 读 `wb.VBProject` → 交给 `check_*` 判定 → `Notes` 收尾；自动写回每个模块写前查存在性、写后验集合 |
| `audit.py` | 用 `vba_module_names(xlsm)` + `check_doc_modules` 离线对账（**不打开工作簿、不要密码**），结果进结构体检 ⑦⑧ 与汇总 |
| `agent.py style` | `check` 末尾 `Notes.render` 出汇总；`fix --apply` 写入文件后 `selfcheck()` 对同一批文件复检，有残留返回非 0 |

业务规则（公共库该写哪些、哪些模块算占位壳）**不进 check**——那是 `sync` 的事。

### features 本轮拆分切片（1.6.3-local）

下列切片在 `scripts/features/` 内按职责拆分，依赖方向仍单向：切片只调 toolkit 基础函数与 adapters 网关，彼此不互 import；入口壳只做参数适配与装配。

| 切片 | 管的事 | 边界 |
|---|---|---|
| `audit_rules` | audit 规则单一来源：阈值 / 白名单 / 分组登记 / 指标规则正则与纯辅助函数 | 只放无副作用的数据与纯函数；不枚举源码、不碰 COM、不写文件；引擎与其它切片不得复制阈值表 |
| `audit_scan` | 源码范围枚举、读取缓存与成员可达性扫描（structure / overview 共用） | 以 mixin 形式挂到 AuditEngine，只吃引擎实例上的 root/config/缓存；纯计算，不打印、不碰 COM、不写文件 |
| `audit_structure` | structure 视角：壳 / 重名 / 绑定 / 文档工作表等结构问题分组对账 | mixin 形式挂到 AuditEngine；规则取自 `audit_rules`，不复制阈值表 |
| `audit_indicators` | indicators 视角：H01–H19 离线体检指标计算 | 运行期探测（H16/H17）由 handler 经网关注入，引擎只判；纯计算，不碰 COM、不写文件 |
| `sync_preflight` | push 写回前源码预检：目录收集、类头 / 续行 / 私有字段 / 编码等静态检查 | 只读，不碰 Excel；返回 `(fatal, notes)`，fatal 非空必须放弃写回，notes 软提醒只进 warnings；与 run/agent 预检的差异轴保留双份，不抽公共骨架 |
| `sync_shells` | 公共库占位壳生成 / 清除：为项目用到的公共模块放同名注释壳，写回时占位壳跳过 | `--apply` 才真正写入 / 删除；默认预览只计数，不写入文件 |
| `sync_init` | init 动词：生成 / 更新 `mapping.json` 与 src 骨架 | 默认只预览，`--apply` 才写入文件；`--lab` 经接缝调空白工作簿生成脚本，测试替换接缝，不起 Excel |
| `run_macro` | run 宏执行装配：参数解析 → 预检 / 备份 / 写回 / 回读 → 运行宏 → 汇总结果与 journal | 不在本模块拼 JSON 信封、不回壳别的 CLI，输出交入口层 `render`；watch 常驻监听辅助不在本模块 |
| `run_watch` | watch 监听辅助：目标解析、探活、比对、单文件写回 / 抓回、基线刷新 | 会话附着 / 打开工作簿只走 adapters 网关；watch 主循环（睡眠钩子、交互闸门）留在入口层 `features.run`；单文件写回 / 回读委托 `run_macro`，不复制第二份 |
| `agent_probe` | probe 动词支撑基础函数：项目名查找、xlsm 解析、测试工作表名取值、适配器 Failure 契约归一 | 只走 adapters 网关与 toolkit 既有基础函数，不回壳别的 CLI、不 import 其它 features；写回链与日志汇总不在本模块 |
| `agent_logs` | logs 动词：日志目录定位、GBK 读取、通过 / 失败计数、尾部行截取 | 项目解析委托 `agent_probe` 既有基础函数，不复制第二份；不碰 COM、不写回工作簿 |
| `sync_interactive` | `sync -i` 菜单本体：列项目 / 全局工具 / 新建项目 → 选动作 → 只读直接运行、写操作先预览再回车执行 → 复用 registry 既有 handler | 菜单层不碰 COM、不写入文件，只做选择与参数装配；`reader` / `writer` / `runner` 可注入；非 TTY 把关由入口壳 `ensure_interactive` 负责；写操作执行时才给 handler 加 `--apply`；`watch` 先问监听秒数再确认开始；新建项目从拖入的 `.xlsm` 自动推导项目名、目录和版本规则 |

---

## 3. 交互组件的正确用法

### 3.1 普通清单（回车无效、只认数字）

```python
k = pick_list([f"{p['name']}   [可用] {path}" for p in projects], "选项目")
if k is None:
    return          # 用户按了 0
name = projects[k]["name"]
```

### 3.2 清单 + 「回车 = 默认项」

`pick_list_default` 是"最常用那几项"专用的：清单照常列全，回车直接取默认集合，省一次挑选。

```python
kind, idx = pick_list_default(labels,
                              prompt="把源码写进哪个工作簿",
                              default_hint="回车 = 已打开的 2 个",
                              exit_hint="返回菜单")
if kind == "exit":      # 按 0
    return
if kind == "default":   # 回车
    names = 已打开的那些
else:                   # 填了编号
    names = [labels[idx] 对应的那个]
```

- `default_hint=None` → 本层不接受回车（提示"请输入编号"），行为和 `pick_list` 一致。
- 返回值固定三段式 `("pick", 下标) / ("default", None) / ("exit", None)`，不要自己再解析字符串。
- **不许**在入口脚本里写 `print 清单 → while True: read_line → 校验数字 → 重问` 的循环（以前 sync.py 有两处这么干，已收进这个函数）。

### 3.3 分类分组清单

```python
flat, sel = pick_grouped([("模块", [...]), ("类模块", [...])], "选模块")
```
编号跨组连续，返回扁平列表下标；调用方自己维护同顺序的扁平列表。

### 3.4 菜单

```python
run_menu(title_lines, [(label, notes), ...], handler,
         prompt="请选择功能编号：", divider="-" * 62)
```
handler 收到 1 基编号（按 0 收到 0），返回 `None` 继续、返回 `int` 作为退出码。

### 3.5 执行 + 兜底

```python
run_action(cmd_single, cfg, name, False, True)   # 异常不外溢，自带 pause()
```
- `run_action` **自带 `pause()`**：菜单分支里不要再手工加。
- 业务函数自己抛 `Failure("可读错误信息")` → 由 `run_action` / `run_cli` 统一翻成 `[错误] …`，不要 `sys.exit`。

---

## 4. 什么该下沉到 toolkit，什么不该

| 该进 toolkit | 留在入口脚本 |
|---|---|
| 编码、换行、路径解析、备份/移动、找最近一份备份 | 具体项目的名字、`mapping.json` 字段含义 |
| 读一行、编号校验、清单排版、菜单骨架 | “写回哪些工作簿”“怎么挑默认项”这类业务选择规则 |
| 仓库结构查询（目录、模块名、占位壳、版本常量） | 业务判定（哪些算死代码、哪些算冲突、公共库该写哪几个） |
| 通用文本处理（剥注释、导出头、成员表、事件识别、规范化） | 某条日志写什么话、备份打什么 tag |
| Excel/COM 的**机械动作**（起实例、找工作簿、识别工程密码保护状态、找/删/替换模块、整工作簿备份） | **何时备份、要不要 `Save`、文档模块开关**、工程引用“哪些必须有” |
| **运行完宏的应用状态还原**（`app_guard`：事件/屏幕/提示/状态栏/计算模式/剪切状态） | 还原后要不要提示、提示什么话 |
| 文件监听三件套（快照 / 落点 / 两轮确认） | 公共库子集过滤、冲突时"以谁为准"的询问 |
| 输出格式（`msg`/`warn`/`fail`）与退出码 | 何时抛、抛什么话 |
| `Failure` 与异常兜底 | 业务规则的提示语 |

判断口诀：**"换一个工具也还会用到吗？"** 会 → toolkit；只有这个工具用 → 留业务。

---

## 5. 加新能力时的步骤

1. 想清楚归属：进 toolkit 哪个子模块（照 §2 的职责表）。
2. 在对应模块写函数，**纯函数优先、不打印**（提示语留给调用方；`ui` 里的清单/菜单例外，它就是负责排版）。
3. 在 `toolkit/__init__.py` 的 `from .x import ...` 与 `__all__` 各加一行。
4. 入口脚本改成调用它，删掉原来那段自己写的实现。
5. 跑一遍：`python -m py_compile vba-workbench\scripts\sync.py`，再逐个编译 toolkit 子模块（PowerShell 下别用 `vba-workbench\scripts\toolkit\*.py` 通配符，会报 `[Errno 22]`），然后真机跑一次菜单。
6. 文档：行为规则写到对应 md（`sync.md` / `audit.md` / 本文件）。

**不要**：在入口脚本里 `import shutil` / `open(...)` / `input()` / `sys.exit()`；不要复制一段 toolkit 代码"只改一点点"。

---

## 6. 现有入口脚本的合规状态

| 检查项 | sync.py | audit.py | agent.py style |
|---|---|---|---|
| `input()` 自带交互 | 无（全部 `read_line` / `choose` / `pick_list*`） | 无 | 无 |
| 手写编号循环 | 无（`_pick_projects`、菜单 10 已改走 `pick_list_default`） | 无 | 无 |
| 裸 `print()` | 无（2026-11 统一换成 `msg()`，29 处） | 无 | 无 |
| 自己开文件写 / 追加日志 | 无（`append_line`；临时目录 `remove_tree`） | 无（`write_text_utf8` / `dump_json`） | 无（`write_source`） |
| 自己拼编码 / 换行 | 无（`write_source` / `gbk_import_copy` / `transcode_exported_file`） | 无 | 无（`write_source`） |
| 自己 `sys.exit` | 无（`cli.entry` + `run_cli`） | 无 | 无 |
| 菜单 / 清单排版 | `run_menu` + `pick_list*` | 同 | 同 |
| 校验与提醒 | `document_preflight` / `verify_*` 判定走 `check` | 文档模块对账走 `check_doc_modules` | 汇总走 `Notes`，写入文件后 `selfcheck` |
| 运行宏/写回后的应用状态还原 | `cmd_run` 与 `cmd_push` 收尾走 `app_guard` | 不需要（不碰 Excel） | 不需要 |

> 2026-11 的一次自查里修掉的三处历史偏差：`_pick_projects` 手写"打印 + 重问"、菜单 10 手写"盯哪个"循环、29 处裸 `print`。现在这类行为只有 toolkit 一份实现。

---

## 7. 常见坑

| 坑 | 说明 |
|---|---|
| 管道/非 tty 下 `read_line` 返回空串 | EOF 不抛异常，返回 `""`。**不要把"空串"无脑当成回车确认**，除非该层明确设置了 `default_hint`；否则会变成无限重问（曾把测试卡死） |
| 全角数字 | `read_line` 已做 NFKC 归一，`１` → `1`。自己绕过它去 `int(raw)` 会崩 |
| `pause()` 与 `run_action` 叠加 | `run_action` 里已 `pause()`，外面再调一次会出现"要按两次回车" |
| `pick_list` 不支持回车 | 需要"回车=全选/默认"就用 `pick_list_default`，别在调用处再包一层循环 |
| 相对路径 | 一律 `abs_path()`（按仓库根解析）。用 `Path("src/...")` 会按当前工作目录走，换个目录启动就找不到 |
| 写源码 | 别用 `Path.write_text(...)`；用 `write_source`，否则 LF/UTF-8 规矩会破 |
| `Failure` 文案 | 抛 `Failure("可读错误信息")`，别在里面塞 traceback；`run_cli` 会打印到 stderr 并返回 1 |
| 宏名与模块重名 | `Application.Run("'书'!测试表格工具")` 会把名字解析成**模块**，报「该宏在此工作簿中不可用」。必须写 `'书'!模块.过程`（`sync.py cmd_run` 的 `_resolve_macro_ref` 会自动补） |
| 从宏里取结果 | VBE 立即窗口读不到，只有两条通道：`Application.Run` 的返回值（Function）、宏自己落的文件。取最新日志用 `files.newest_file(目录, "*.log", since=开始时间)`，打尾部用 `files.read_tail(path, N)`（GBK 自动回退） |
