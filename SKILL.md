---
name: vba-workbench
description: VBA 工程管家（vba-workbench）：在 Windows + 微软 Excel 上操作 .xlsm 里的 VBA 工程——把工程与 src/ 源码双向同步（pull 抓取 / push 写回 / watch 自动写回 / run 运行宏），并做只读代码盘点、风格检查与环境体检。用于提到 VBA 工程、vbaProject.bin、导出/导入模块、源码同步、提交前查格式、删公共模块前查引用、Excel COM/VBE 权限排查的用户任务；纯 VBA 语法问答、Excel 操作教学、WPS/Mac、文件转换不用。
license: "Source-Available Non-Sale ShareAlike 1.0 (SNSA-1.0)"
metadata:
  version: "1.6.13"
  compatibility: Windows 10/11 + 微软 Excel 2010+；Python 3.10+；pywin32。命令按 Windows 路径写，cmd 与 PowerShell 等价执行。
---

# VBA 工程管家（vba-workbench）

把只存在于 `.xlsm` 二进制里的 VBA 同步成可 diff 的 `src/` 源码，再原样写回工作簿。
定位：**可独立分发的 skill**；拿走 `vba-workbench/` + 自己的 `mapping.json` 就能用。

## 术语约定

| 术语 | 含义 | 不再混用 |
|---|---|---|
| 工作簿 | 一个 `.xlsm` 文件 | 用“表”泛指整个文件 |
| 工作表 | Excel 里的 Sheet | 表页 |
| 抓取 | VBE → `src/` 源码 | 导出、抓源码 |
| 写回 | `src/` 源码 → VBE | 写盘、落盘 |
| 运行宏 | 通过 `run` 执行 VBA 宏 | 跑宏、执行宏 |
| 自动写回 | `watch` 监听源码变化并写回 | 常驻、监听（单独使用时） |
| 解除密码保护 | 取消 VBA 工程的密码保护 | 解密、解锁 |
| 判定规则 | 检查项的判定标准 | 口径 |
| 可读提示 | 面向人的文本输出 | 人话 |

真实菜单标签、模块名、工作表名里的“表”保留原文；`整表备份`、`表格工具`、`测试表` 是固定名称，不做替换。讨论 ListObject（Excel 表格对象）时仍用“表格”。

## 一、命令标准

```
<工具> <动词> [--project P] [--module M] [--macro 名] [--json] [--apply]
```

| 规则 | 内容 |
|---|---|
| 动词 | 白名单；未知动词退出 2，不会把错词当项目名 |
| 参数 | `--project` / `--module` / `--macro` / `--all-project`；不用位置参数 |
| 写操作 | 只读动词直接出结果；写操作默认只预览，`--apply` 才执行 |
| JSON | `--json` 是布尔开关，stdout 恒一行 `{ok, action, details, errors, elapsed_ms, ts, stats, exit_code}`；`action` 为 `<入口>.<动词>`（如 `sync.push`），`errors` 是对象数组，`ts` 带时区；`audit.py` 另有 `--out PATH` |
| 失败 | `exit 0 或 3 ⇒ ok=true`；`exit 1 ⇒ ok=false 且 errors 非空`（每项含 `code/stage/message/hint/suggested_command/retryable`）；`exit 3` 只读体检“仅警告”（`ok=true`、`stats.warned>0` 或 `stats.configured=false`；`structure/indicators --strict` 时升为 1，`common` 无公共库也按 3 但不接受 `--strict`）；`overview` 例外：规则提醒默认 0，只有阻断组退 1；`exit 2` 用法错（只有可读提示，不出 JSON） |
| 已删除 | `--yes`、`--dry-run`、`--clean`、全局 `--save`（仅 `run --isolate` 保留）、`export/all` 别名、项目名当动词 |

`--json` 写在动词后；`watch` 不支持 `--json`（常驻命令，传了退出 2）。

## 二、什么时候用

### 2.1 先判断工作簿状态：是否已打开、有没有未保存改动

| 场景 | 命令 | 需要先打开工作簿 | 保存 |
|---|---|---|---|
| 工作簿已打开，有未保存改动，要抓取源码 | `pull`（自动借用已打开会话；`pull-open` 为只借用会话的显式入口） | 是 | 不保存 |
| 工作簿已打开，要写回 | `push` | 是（默认自动借用已打开会话） | 是 |
| 工作簿已打开，要运行宏 | `run` | 是 | 默认不保存 |
| 工作簿未打开，要抓取源码 | `pull` | 否，自己开隐藏 Excel 只读 | 不保存 |
| 工作簿未打开，单个工作簿要写回 | `push` | 否，默认自动开隐藏 Excel 写回 | 是 |
| 工作簿未打开，多个工作簿一起写回 | `push --all-project` | 否 | 是 |
| 工作簿未打开，要运行宏 | `run --isolate` | 否，隐藏实例；不触发 `Workbook_Open` | 默认不保存，`--save` 才保存 |
| 只读盘点模块 | `vbamod` / `audit` | 否 | 不改 |
| 需要看 VBE / 手工解除密码保护 | `agent.py book --action open` | 打开可见 Excel | 不保存 |
| 不确定该用哪个 | `sync.py plan --project 甲 --intent push --json` | 只探测，不写 | 不写 |

### 2.2 硬规则

1. **工作簿未打开时使用磁盘模式，不要要求用户先打开工作簿。** `push` 默认在 `NO_EXCEL/TARGET_NOT_OPEN` 时自动转隐藏实例写回；`--open-only` 才保留旧行为。
2. **工作簿已打开且有未保存改动时，禁止磁盘写回。** 只能走已打开会话；连不上就报 `COM_CONTEXT`，不要另开副本。
3. **agent 不许用 `Start-Process` 直接打开 Excel。** 要打开可见工作簿只能走 `agent.py book --action open`。
4. **`watch` 只能监听已打开的工作簿**；工作簿未打开时先 `book open`，或改用 `push --all-project` 一次性写回。
5. **`run --isolate` 只适合不依赖界面/事件的宏**；隐藏实例不触发 `Workbook_Open`，需要界面或打开事件的宏必须用已打开会话。
6. **写回前先看预检**：`.cls` 缺 `VERSION 1.0 CLASS`、单语句续行超过 24 行是硬错误；类模块 `Me.` 引用私有字段只软提醒。现象、处理与禁止事项见 `references/pitfalls.md` §11、§13–15。

常用命令：

| 场景 | 命令 |
|---|---|
| 工作簿里改完，提交到 Git | `sync.py pull --project 甲 --apply`（自动判断：已打开借会话、未打开走磁盘）；`pull-open` 只借用已打开会话 |
| 源码改完，写回工作簿后才能运行 | `sync.py push --project 甲 --apply`（自动判断工作簿是否已打开） |
| 多个工作簿一起写回 | `sync.py push --all-project --apply` |
| 边改边自动写回 | `sync.py watch --project 甲 --apply`（工作簿必须已打开） |
| 写回后直接运行宏 | `sync.py run --project 甲 --macro 性能测试 --apply`（工作簿已打开）；工作簿未打开时加 `--isolate` |
| 运行断言宏，核对通过/失败计数 | `agent.py assert --project 甲 --apply`（只接会写日志的断言宏；只写工作表的自测宏改用 `sync.py run`） |
| 换新版 xlsm，确认配置到哪个工作簿 | `sync.py list --json` |
| 新机器 / 新仓库初始化 | `sync.py init --project 甲 --dir <工作簿目录> --apply` |
| 删除代码前查引用 | `audit.py members --module 表格工具 --json` |
| 公共库瘦身 / 结构体检 | `audit.py common` / `audit.py structure --json` |
| 项目里确实要直接关闭加速 | 优先用 `表格工具.安全执行` 成对关闭并还原；非关不可时在关闭块**上方三行内**标 `[加速例外]`，H05a 才放行（规则详见 `references/audit.md`） |
| 提交前查格式 | `agent.py style --action check` |
| 过度设计体检 | `agent.py metrics` |
| 环境是否可运行 | `sync.py doctor --json` |

**不适用**：非 Windows、WPS、Mac Excel、非 VBA 源码同步、需要联网的协作。

## 三、核心流程

1. **初始化**：`sync.py init ... --apply` 生成 `mapping.json` 与 `src` 骨架；或复制 `assets/mapping.example.json`。
2. **自检**：`sync.py doctor --json` → `failed=0`。
3. **核对**：`sync.py list --json` → 每个项目 `ok=true` 且 `workbook` 指向真实 xlsm；未解析项看 `errors[].code=TARGET_NOT_FOUND`。
4. **判断工作簿状态**：`sync.py plan --project 甲 --intent pull|push|run|watch --json`，按返回的 `mode/needs_open/command` 执行；工作簿已打开且有未保存改动时尤其要先看。
5. **抓取**：`sync.py pull --all-project --apply`（自动判断：已打开借会话、未打开走磁盘）；`pull-open --project 甲 --apply` 为只借用已打开会话的显式入口。
6. **改源码**：只改 `src/`，改完运行 `agent.py style --action check`。
7. **写回**：`sync.py push --project 甲 --apply`（自动判断工作簿是否已打开；顺序是**源码预检 → 备份 → 写回 → 回读校验**，预检失败连 Excel 都不打开）。
8. **验证**：`sync.py run --project 甲 --macro 宏名 --apply`（工作簿已打开）或加 `--isolate`（工作簿未打开）；每次实际运行会写 `.run/<时间>-<项目>-<宏>.json` 运行记录（只记录，不支持 `--resume`）；再运行 `audit.py structure --json` / `indicators --json`。
9. **提交**：`git add src` / `git commit`；`xlsm` 和 `.backup` 不纳入版本库。

## 四、工具与资料

- `scripts/sync.py`：17 个动词（init / list / doctor / scrub / version / vbamod / plan / pull / pull-open / push / run / watch / shells / align / promote / gbk-export / clean-backup）。
- `scripts/audit.py`：只读 6 视角（members / common / project / structure / indicators / overview），同一信封。
- `scripts/agent.py`：probe / book / backup / patch / assert / push / style / logs / metrics；写类动词也要 `--apply`。
- `scripts/toolkit/`：公共基础函数；业务编排在 `scripts/features/`。三个入口只做参数解析、上下文组装与渲染，入口之间不互相 import；`features` 之间也不互相 import。
- `src/`：源码唯一真源（`src/common` + `src/projects/<项目>`）；`assets/`：配置模板；`evals/`：触发与行为用例。
- 资料按需读：`references/commands.md`（32 动词完整命令表，manifest 生成）、`references/sync.md`（总览）、`references/sync-commands.md`（动词/JSON 字段）、`references/sync-config.md`（mapping）、`references/sync-maintain.md`（改内部逻辑）、`references/audit.md`（判定规则）、`references/toolkit.md`（公共基础函数）、`references/agent.md`、`references/pitfalls.md`（常见问题与处理）、`references/release.md`；1.7 待办与已关闭项见 `references/backlog-1.7.md`。日常调用只需本页。

## 五、给 agent 的约定

- **先判断工作簿状态再执行**：不确定就 `sync.py plan --project X --intent ... --json`；不要先让用户打开工作簿，也不要 `Start-Process`。
- **先预览再执行**：不带 `--apply` 看 `details`；`plan` / `pull` / `pull-open` / `vbamod` 的结果行在 `stats.results`，不要只看 `details`。确认后加 `--apply`；预览不碰 Excel、不修改文件。
- **非交互**：`--no-input` 关闭所有提问；缺参数直接失败。
- **串行**：同一 Excel 实例不要并行两支写回；`watch` 与 `push` 不要同时写同一个工作簿。
- **硬性限制**：agent/CI 不运行 `watch`（无限循环；必须运行就加 `--seconds N`）；不操作未登记工作簿；不改业务数据；工程有密码保护时先解除保护，不要重试。
- **退出码**：0 成功；1 业务失败（看 `errors[0].code` / `stage`）；2 用法错；3 只读体检仅有警告（`structure` / `indicators` 加 `--strict` 升为 1；`common` 无公共库也退 3，但不接受 `--strict`）；`overview` 默认只对 `sheet_mismatch/binding_target/dup/shadow` 退 1，规则提醒只计数，`--strict` 才按 1；130 用户中断。
- **幂等**：`pull` / `push` 可重复；`push` 每次先备份。
- **错误分流**：先看 `errors[0].code` 与 `stage`，不要凭报错文本猜。完整码表只维护在 `references/usage-examples.md`（`errors[0].code` 段），逐码现象、原因、处理与禁止事项见 `references/pitfalls.md`。高频处置：`TARGET_NOT_FOUND` 查 `search_dir/prefix`；`EXCEL_CLOSED` 重开工作簿；`EXCEL_BUSY` 先关闭弹窗或等待计算；`NO_EXCEL/TARGET_NOT_OPEN` 走该动词给出的出口；`COM_CONTEXT` 查权限/会话；`CONFIG_NOT_FOUND` 按报出的完整路径建 `mapping.json`；`SOURCE_PREFLIGHT` 先修源码（不打开工作簿、不备份）；`ARG_MISMATCH` 核对 `--arg` 个数/顺序；缺宏 / 禁用宏 / 业务报错 1.6 统一归 `RUN_FAILED`，宏卡住 / 弹窗阻塞归 `WATCHDOG_TIMEOUT`；旧细分码已撤销，区分信息只看 `hint`（宏名/编译问题 vs 信任位置）。

## 六、边界

- 源码保存为 UTF-8 无 BOM + LF；只有写回 VBE 的 GBK 临时副本与 `.backup/GBK` 用 CRLF。临时副本默认放在工作簿同目录的 `.vba-import-*`，结束时删除；这样隐藏实例/隔离实例也能读到，不会因沙箱 temp 不可见报 `0x800A9D00`。`.backup/GBK` 是当前 `src` 的镜像，不是历史备份；每次 `gbk-export --apply` 先清空目标组再重建。
- 写回后默认保存；`--no-save` 只改内存；备份在 `.backup/<时间戳>/`。`postcheck` 失败默认只“不保存”，VBE 内存可能已被写脏；要按写回前正文复原本次模块，加 `--rollback-on-fail`。
- 备份自动清理按 `backup_keep` / `backup_keep_daily` / `backup_min_keep` 三档执行；人工回滚点写进 `defaults.backup_protect`（tag 或批次名关键字）或 `clean-backup --keep-tag`，清理时跳过并列出。
- `run --isolate` 不触发 `Workbook_Open`，也不适合依赖界面/事件的宏；`--save` 才把运行结果写入磁盘。
- `agent.py assert` 只接**会写日志**的断言宏：靠运行完新日志里的「通过/失败」计数出结论；只写工作表、不写日志的自测宏恒报 `ASSERT_NO_LOG`。它不是通用宏执行器，测自测宏或只看工作表结果请改用 `sync.py run`（工作簿未打开时加 `--isolate`）。
- 不主动删 `.backup` 历史批次；自动清理和显式 `clean-backup --apply` 都只按保留规则删旧批，`backup_protect` / `--keep-tag` 命中的批次会被跳过并列出，不会被这两条路径删除；要删先解除保护。
- 文档与示例必须脱敏：真实用户名、机器名、业务工作簿名、客户名、经营数字都不进入发布包。
- 发布按 `references/release.md`：版本三处一致 → doctor → 脱敏扫描（`sync.py scrub --json`） → 打包 → `SHA256SUMS.txt` + `latest.txt`。

示例提示词与完整命令行清单见 `references/usage-examples.md`。
