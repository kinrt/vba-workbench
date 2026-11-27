# sync.py 配置册：mapping.json 与目录约定

> 本册只讲 `mapping.json` 与目录约定；前置概念见 `SKILL.md`，抓取/写回原理见 `sync.md`。四册各自从 §1 编号，不要求按册顺序读。

## 目录

- §1 配置 `mapping.json`（1.1 `defaults` / 1.2 项目字段 / 1.3 示例写法 / 1.4 本机覆盖层 / 1.5 换版）
- §2 目录约定（2.1 文档模块的文件名带工作表名）

## 1. 配置 `mapping.json`

文件在**代码根 `mapping.json`**（与 `vba-workbench/` 平级）；包内 `vba-workbench/assets/mapping.example.json` 是供人复制的占位样例（不含真实路径）。`sync.py init` 会按目标路径生成或合并配置。两种写法都支持：`{"defaults":…, "projects":[…]}`，或直接给一个项目数组。

**查找顺序**（`toolkit.base.find_config()`）：设了 `VBA_WORKBENCH_CONFIG` 就只认它，文件不存在也把该路径原样报进 `CONFIG_NOT_FOUND`，不再回退到其它位置；未设时才从代码根逐级向上找 `mapping.json`，最后看 `%APPDATA%\VBA工程管家\mapping.json`。都没有就报“找不到配置文件”并提示 `sync.py init` 或复制 `assets/mapping.example.json`。

**路径基准 workspace**：`abs_path()` 把所有相对路径（`src/...`、`.backup`、`common_dir`）按 workspace 解析。缺省 workspace = **配置文件所在目录**；配置里写一个 `workspace` 键可以改（相对配置文件目录或绝对路径），用于"配置放在别处、代码在仓库里"的场景。

`python vba-workbench\scripts\sync.py list --json` 的头部会打印实际用的**配置路径**和**基准目录**——"改了不生效"先看这两行。

**生成配置（新机器 / 新仓库）**：

```bat
python vba-workbench\scripts\sync.py init --project 示例表A --dir "D:\表格\示例表A" --json
python vba-workbench\scripts\sync.py init --project 示例表A --dir "D:\表格\示例表A" --apply
```

- 默认预览，`--apply` 才写 `mapping.json` 并建 `src/common`、`src/projects/<项目>`。
- 目标配置默认按上面的查找顺序定位；测试或多配置场景用 `--config PATH` 显式指定，避免落到真实根配置。`init` 的 `--dir` / `--src` 相对路径以**目标配置所在目录**为基准。
- 同名项目只更新字段，不追加重复项；目录不存在直接退出 2（JSON `errors` 有原因）。
- `--lab` 在配置目录下生成空白测试表（需要 Excel + pywin32）；没给 `--dir` 时默认 `<配置目录>/.lab`，源码也放同一目录，不再写进共享的 skills 根。

### 1.1 `defaults`

| 字段 | 默认 | 含义 |
|---|---|---|
| `common_dir` | `src/common` | 公共库源码目录（相对仓库根） |
| `backup_dir` | `.backup` | 备份根目录 |
| `import_document_modules` | `false` | 是否默认覆盖文档模块（等价于常开 `--doc-modules`） |
| `backup_auto_clean` | `true` | 每次备份成功后是否自动清理旧备份 |
| `backup_keep` | `10` | 每个 tag（`""` `-抓取前` `-自动写回` …）保留最近几批 |
| `backup_keep_daily` | `1` | 每个 tag 每天再留最新几份（0 = 只按批数留） |
| `backup_min_keep` | `3` | 全局最新几批永不删（`latest_backup` 核对模块清单要用最近一份） |
| `backup_protect` | `["-覆盖前"]` | 自动清理与 `clean-backup` 都跳过的 tag / 批次名关键字（人工回滚点）；`clean-backup --keep-tag` 可临时追加 |
| `assert_macro` | `测试表格工具.运行全部断言` | `agent.py assert` 的默认断言宏，写「模块.过程」；项目可同名键覆盖 |
| `assert_sheet` | `测试表` | `agent.py assert --clear` 与 `probe` 检查的测试工作表名；项目可同名键覆盖 |

### 1.2 项目字段

| 字段（必填） | 含义 |
|---|---|
| `name`（是） | 项目名，同时是命令行/菜单里选它的标识；不能重复 |
| `file`（二选一） | 固定路径的 xlsm（相对仓库根或绝对路径）；换版要手工改。菜单拖入无 ` V<数字>` 的工作簿时会自动按文件登记 |
| `search_dir`（二选一） | 目录里自动挑最新版本；配合 `prefix` 过滤文件名前缀 |
| `prefix`（否） | 文件名前缀，只写基名（`示例表A`、`示例表C`），不带 ` V` 和版本号 |
| `require_version`（否） | 默认 `true`：文件名必须以 `V<纯数字>` 结尾才算候选，否则报错；写 `false` 才退回按修改时间 |
| `version_pattern`（否） | 自定义版本号正则，第一个捕获组是数字；默认 `V<数字>` 结尾 |
| `src`（建议） | 该工作簿源码目录，如 `src/projects/示例表A` |
| `common_version`（否） | 该工作簿锁定的公共库版本，供“全局工具 → 核对公共库版本”比对 |
| `common_modules`（否） | `"auto"`＝按引用自动挑；列表＝白名单；缺省＝整库写回 |
| `assert_macro`（否） | 覆盖 `defaults.assert_macro`，写「模块.过程」 |
| `assert_sheet`（否） | 覆盖 `defaults.assert_sheet`，指定该项目的测试表 |
| `note`（否） | 备注，项目清单与 JSON 里显示 |

**版本识别**（`resolve_target` + `parse_stem_version`）：文件名必须**以 `V<纯数字>` 结尾**才算带版本（`示例表B` → 7、`示例表A V<数字>` → 261120）；`V` 前可以有空格，`V` 后必须是数字且名字到那里就结束——后面再跟任何文字都不算版本（`示例表A V<数字> (version 1)`、`示例表B - 副本`、`示例表C - 2022年` 全部不认）。按**整数**比较取最大，`V10` 比 `V9` 新；版本号相同的两份取修改时间最新的。

`require_version`（默认开）下，解析不到版本号的文件**直接不算候选**：宁可报「没有以 V<数字> 结尾的文件」也不猜，避免误选同族历史工作簿或 Excel 崩溃副本。只有项目显式写 `"require_version": false` 才退回“按修改时间取最新”。

> 以前这里是把数字当日期（`V240101` 就想当 2024-01-01）解读的，只能认 6 位/8 位，`V7` 这种递增序号认不出来。现已统一改成"数字即版本"，不再做日期换算。

### 1.3 示例写法

> 这里只挑几个项目**示范写法**，不要求列全 `mapping.json`；登记/停用工作簿都不用回来改这份清单。项目全集永远看 `python vba-workbench\scripts\sync.py list --json`。

```json
{
  "defaults": {"common_dir": "src/common", "backup_dir": ".backup", "import_document_modules": false,
               "require_version": true},
  "projects": [
    {"name": "示例表A",
     "search_dir": "C:/Users/<用户>/表格/示例表A",
     "prefix": "示例表A",
     "require_version": true,
     "src": "src/projects/示例表A",
     "common_version": "1.3.0", "common_modules": "auto"},
    {"name": "示例表B",
     "search_dir": "C:/Users/<用户>/表格/示例表B",
     "prefix": "示例表B",
     "require_version": true,
     "src": "src/projects/示例表B",
     "common_version": "1.3.0", "common_modules": "auto"}
  ]
}
```

两个项目都走 `search_dir` + `prefix`（版本自动滚，`prefix` 只写基名）：

| 项目 | 写法 | 选工作簿方式（`sync.py list --json` 实测） |
|---|---|---|
| `示例表A` | `search_dir` + `prefix` | 目录里 2 份带版本号，取版本号最大的那份 |
| `示例表B` | `search_dir` + `prefix` | 目录里混着 `V4`/`V5`/`V6`/`V6 - 副本`/`V6 - 历史` 等多份，只有合规的 `V<数字>` 算候选，取最大；以后出新版本不用改配置 |

`file` 固定路径仍可用（与 `search_dir` 二选一），但路径写死、**不会**自动跳到新版本，换版本要手工改配置，只留给确实没法按前缀区分的工作簿。

### 1.4 本机覆盖层 `mapping.local.json`（可选，不纳入版本库）

同一个目录下放一份 `mapping.local.json`（形状与主配置相同，已在 `.gitignore` 里），
`load_config()` 会先把它合并进主配置再校验，用来应对“换机器、临时切盘、某个工作簿先停用”：

| 写法 | 效果 |
|---|---|
| `{"defaults": {"require_version": false}}` | 只覆盖写到的键，其余沿用主配置 |
| `{"projects": [{"name": "示例表A", "search_dir": "D:/临时/示例表A"}]}` | 同名项目只覆盖写到的字段（`prefix`、`src` 等保持不变） |
| `{"projects": [{"name": "示例表C", ...}]}` | 主配置里没有的新项目，追加 |
| `{"_drop": ["示例表B"]}` | 临时剔除某几个项目，不动主配置 |

`sync.py list --json` 会先打印「已合并本机覆盖层：…」，看不到这行就是没生效（文件名/目录不对）。
登记新工作簿仍写主配置 `mapping.json`（`load_raw_config()` 不读覆盖层，避免把临时改动写进主配置）。

### 1.5 换版怎么走（配置永不改）

1. **改文件名**：`示例表A V<数字>.xlsm` → 下次直接存成 `示例表A V<更大的数字>.xlsm`，` V<数字>` 照写；目录、`src`、`mapping.json` 都不动。
2. **核对**：`python vba-workbench\scripts\sync.py list --json`，看该项目的「原因」里 `按版本号 V… 取最新，候选 N 个`，确认挑中的是新文件。
3. **对接**：工作簿已打开时，`python vba-workbench\scripts\sync.py pull --project <项目名> --apply` 会自动借已打开会话抓取含未保存改动的代码；`pull-open --project <项目名> --apply` 是只借用已打开会话的显式入口；或 `push --project <项目名> --apply`；全项目 `push --all-project --apply` 也行。

出问题时的两种典型报错：

- **`没有以 V<数字> 结尾的文件（缺版本号的 N 个：…）`** → 目录里有同族工作簿没写 ` V<数字>`（历史年度版本、`- 副本`、崩溃恢复的 `(version 1).xlsb`）。改文件名或挪走，别改回 `require_version: false`。
- **挑中了旧文件** → 先看候选数，再确认旧版本文件是否还在同目录；版本号按整数比，`V100` 比 `V99` 新，但 `V2026` 比 `V7` 大，跨命名体系混放会选错。

每个工作簿都靠 `search_dir` + `prefix` 解析：`sync.py list` 的「原因」会打「按版本号 V… 取最新，候选 N 个」，候选数就是目录里同族工作簿的份数；换版只改文件名，配置不动。

---

---

## 2. 目录约定

```
仓库根/                       ← mapping.json 所在目录（配置查找的锚点）
├─ mapping.json              本机配置（真实路径，不入公开包；路径基准 = 本目录）
├─ vba-workbench/                可独立分发的 skill 包（不含真实配置与 src）
│  ├─ SKILL.md               skill 入口（能力 / 触发条件 / 命令速查 / 安全边界）
│  ├─ assets/
│  │  └─ mapping.example.json 配置占位样例（对外交付用，不含真实路径）
│  ├─ requirements.txt / VERSION / CHANGELOG.md / LICENSE
│  ├─ scripts/
│  │  ├─ sync.py             同步主脚本（只写业务逻辑）
│  │  ├─ audit.py            成员引用盘点（只读，与 sync.py 互不导入）
│  │  ├─ agent.py            Agent 统一入口
│  │  ├─ 启动菜单.bat        双击进交互菜单（带 -i）
│  │  └─ toolkit/            公共层包：base（代码根/配置查找/workspace 基准/UTF-8 读源码（GBK 兜底）/Failure/消息）、
│  │                         config（mapping.json 读写）、vba_parse（剥注释/过程声明）、
│  │                         repo（项目目录/壳识别/版本正则/find_source）、
│  │                         files（write_source=UTF-8+LF、边界转码、备份、移动）、
│  │                         check / style_rules / excel / watch / app_guard / vbabin /
│  │                         ui（run_menu/pick_list/pick_grouped）、cli（-i/退出码）
│  └─ references/            说明书：sync 四册 + audit.md / agent.md / toolkit.md / usage-examples.md / release.md
├─ src/
│  ├─ common/                公共库（所有工作簿共享）
│  │  ├─ 模块/               .bas
│  │  ├─ 类模块/             .cls
│  │  └─ 窗体/               .frm + .frx
│  └─ projects/<项目名>/       每个工作簿一份，目录结构同上
│     ├─ Microsoft Excel 对象/   Sheet1、ThisWorkbook 等文档模块（.bas）
│     ├─ 窗体/
│     ├─ 模块/
│     └─ 类模块/
├─ .backup/<时间戳>/          写回前备份的 xlsm（watch 的 tag 是「-自动写回」，启动只备一次）
├─ .backup/<时间戳>-抓取前/    pull-open / 抓取前用 SaveCopyAs 留的底（`-换壳` / `-删表页前` 是历史 tag 名，同理，tag 区分来源）
├─ .backup/GBK/              仓库外 VBE 导入镜像：<公共库|项目>/…（每次 gbk-export --apply 先清空目标组再重建，不是历史备份）
├─ .backup/watch/             自动写回日志：<项目>-<YYYYMMDD-HHMMSS>.log（UTF-8 + LF，可长期留着对账）
├─ 其他/                      暂存区：暂时不用的文件、旧资料
└─ .vscode/settings.json     按 UTF-8 + LF 读源码（GBK 只在交换边界，由脚本转码）
```

分类目录名与 VBE 工程资源管理器一致：`Microsoft Excel 对象` / `窗体` / `模块` / `类模块`。脚本也兼容"全部平铺在项目根"的旧布局（`iter_source_files` 用 `rglob`）。

> `EVAL.md`、`evals/`、`tests/` 只存在于完整 `main` 开发仓库，不随安装包或公开快照；安装包只含 82 个交付文件。

### 2.1 文档模块的文件名带工作表名

`Microsoft Excel 对象/` 下的文件，文件名 = **模块名 + 括号里的工作表名**，和 VBE 左侧显示的一一对上：

```
Sheet1(汇总表).bas      ← VBE 里是  Sheet1 (汇总表)
Sheet4(SKU统计).bas
Sheet15(配置表).bas
ThisWorkbook.bas         ← 没有工作表名，照旧
```

- **只有文件名带括号，模块名不带**：写回时 `module_name_of()` 把括号整段丢掉，工程里的模块仍然是 `Sheet1`；读源码时按模块名比较，任何地方都不要拿文件名主干去当模块名。
- 规则只住在 `toolkit/repo.py` 一处：`doc_stem(模块名, 工作表名)` 生成文件名、`module_name_of(路径)` 取模块名、`doc_label_of(路径)` 取工作表名。`sync.py` / `audit.py` 都调这三个函数，别在脚本里另写正则。
- 老布局 `Sheet4.bas` 照样认；下次抓取时按新名写入文件，旧的同名文件被顺手清掉（`export_component` 里按模块名匹配清理）。
- 工作表名里的非法字符（`:` `*` `?` 等）换成 `_`。
- **工作表名本身就只是个编号时不加标签**（`Sheet10`、`Chart3`…）：那种名字既认不出是哪个工作簿，还会和 CodeName 搅在一起。判断在 `doc_stem` 里（`_GENERIC_TAB_RE`），所以像 `示例表D V<数字>` 里那 9 个被改名为 `Sheet15` 之类、又处于 `veryHidden` 的工作表，源码名仍是 `Sheet22.bas`。
- 没有标签的 `SheetN.bas` 有两种情况：`ThisWorkbook`（本来就没有工作表名）、或者工作表已从工作簿里删掉（这类通常是历史残留，抓取会按“工程里没有就删文件”清掉）。

好处是 `src` 里一眼就能看出是哪个工作簿，`git diff` 也不再只有一堆 `Sheet7.bas`。

**`其他/` 是暂存区**：放暂时不用的文件（停用项目的 xlsm、一次性脚本）和以后可能需要追溯、但不便放在源码目录的旧资料（如已停用的设计稿）。

- 它**不参与同步**：不登记在 `mapping.json`、`module_index` / `common_subset_stems` 只扫 `src/common`、`auto` 判定只扫 `src/projects/<项目名>`，所以放在这里不会被任何工作簿的写回/抓取动作碰到。
- 归类原则：能删就别放；放着是为了“以后还想翻回来”，不是“懒得分类”。
- 文本文件（`.md`）纳入版本库便于追溯；xlsm 等二进制按既定策略不纳入版本库（见 `sync-maintain.md` §5 取舍）。
- 从 `其他/` 复活一个工作簿：把文件挪回工作目录 → 在 `mapping.json` 登记（或菜单“新建项目”拖入）→ 用菜单“抓取源码”抓回 `src/`。

---
