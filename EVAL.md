# 评测报告（1.6.3）

> 现行判定标准以 `SKILL.md` 与《SKILL管理规范.md》v1.5 为准（1.6.13 发布基线：32 个动词 / 74 个交付件 / 591 例）；本页是 1.6.3 评审历史快照，下方的 31 动词、494 例、`1.6.3` 版本号只用于追溯，不作为现行验收基线。

> 历史术语说明：下文保留 1.6.3 原文，其中“口径”“原语”“落盘”分别指判定标准 / 判定规则、公共基础函数、写入文件；历史数字与结论不改。

> 历史范围：1.6 分层架构（入口壳 → features 切片 → toolkit 基础函数 → adapters 网关）。自动化基线为 `review/v1.6-local` 评审终态全量复测实测结果（HEAD `6ee66d0`）：**494 例 / 33 模块 / skipped 4**（0 failure / 0 error，通过 490）；命令面 **31 个动词**（sync 16 / audit 6 / agent 9）。沿革：1.6.2 终态 `52368a5` 为 450 例 / 32 模块 / skipped 9（通过 441），仅作对照，非当前基线。

## 范围

- 做什么：Windows + 微软 Excel 上，`.xlsm` 的 VBA 源码与 Git 源码双向同步（pull / push / watch / run），外加只读盘点、风格检查、环境体检。
- 不做什么：非 Windows、WPS、Mac Excel、非 VBA 语言、联网协作。
- 平台规则：命令在 `cmd` / PowerShell 里等价执行，脚本内部不依赖 Unix-only 命令。

## 自动化评测（CI 可跑，不碰 Excel）

### 规模与文件统计

- 测试模块：33 个；用例 494 例；skipped 4（0 failure / 0 error，通过 490）。规模沿革：复测前基线 31 模块 / 355 例 → 批次 3 口径 32 模块 / 383 例（P1/P6/P8 补用例 + P8 新增分层护栏 `tests/test_import_layers.py`）→ P14 口径 396 例（P12/P13 补用例 +13，模块仍 32、skipped 仍 9）→ 1.6.1 口径 418 例（六项修复补用例 +22，模块仍 32、skipped 仍 9）→ 1.6.2 八项修复口径 440 例 → 1.6.2 命令绝对化收口口径 450 例（模块仍 32、skipped 仍 9）→ 1.6.3 定版口径 475 例（skipped 8）→ 1.6.3 评审终态当前口径 494 例（模块 33、skipped 4：看门狗 / 弹窗 / PID 原语与交互菜单回补恢复 8 条旧跳过，B5 1 条维持硬跳过；本评审环境另 3 条环境条件用例跳过，见下条）。355 / 383 / 396 / 418 / 440 / 450 / 475 仅作沿革对照，非当前口径。
- skipped 构成（本评审环境实测 4 条）：1 条硬 `@unittest.skip`——`test_json_contract.test_agent_assert_failed_has_actionable_hint`（B5 已复核为接受设计差异，旧 exit 2 壳层合成不恢复）；1 条缺仓库 `mapping.json` 的 audit overview 单行 JSON 用例（`test_json_contract.test_audit_overview_stdout_is_single_json_line`，`skipUnless`）；2 条 `test_release_diff.RealBaselineTest` 基线仓库 / `baseline-1.5.6` 标签不可用（`skipTest`）。后 3 条为环境条件，在带配置与基线标签的环境实跑通过、不计产品缺陷；CI 环境缺配置或缺标签，skip 数会相应增加。
- 命令面：31 动词（sync 16 / audit 6 / agent 9），manifest 是唯一登记源；1.6.10 已增至 32（见文首现行口径说明）。
- 入口行数：`scripts/sync.py` 74 行、`scripts/agent.py` 59 行、`scripts/audit.py` 77 行；`SKILL.md` 122 行。
- 版本三处一致：`VERSION`、`SKILL.md` metadata.version、`--version` 输出均为 `1.6.3`。
- 交付与复测：升级预览交付 **58** 个文件为 1.6.2 历史口径（沿革：58 历史 → 60（P1 后）→ 61（P6 补 LICENSE）→ 62（P7 新增 `references/commands.md`）→ **58（P11 目录治理：删除 `docs/commands.md` 短指针，三份验收/差分/对齐文档移入 `tests/` 不交付）**）；P19 起 apply 路径另落 `.installed.json` 清单（落盘 59 = 58 交付 + 清单）为 1.6.2 历史口径，本轮未跑升级预览、交付数未复测。差分 `match 38 / pending 0 / failed 0` 同为 1.6.2 历史口径；命令表 31 动词生成物逐字节一致；`python -m compileall -q scripts tests` 退出 0。

### 场景分组与判据

| # | 场景 | 判据 | 结果 |
|---|---|---|---|
| 1–2 | 版本号择优 | `V99/V100/V100 - 副本/历史` 中只选 `V100`；无版本号返回 None | 通过 |
| 3–4 | 配置查找 | 环境变量 > 向上逐级 > `%APPDATA%`；找不到给明确报错 | 通过 |
| 5–6 | 本机覆盖层 | defaults 按键覆盖、projects 按名合并/追加、`_drop` 生效 | 通过 |
| 7–8 | `--json` 契约 | `list` stdout 一行；覆盖层提示只在 stderr | 通过 |
| 9–10 | 失败契约 | 项目名写错时 `errors` 有原因；退出码 1 | 通过 |
| 11–12 | 结构体检门槛 | 已登记提醒不计入失败；真问题计数为 1 | 通过 |
| 13–14 | 坏工程 | 坏 zip / 有 zip 无 bin 两种都抛 Failure | 通过 |
| 15–16 | 版本一致 | `VERSION` == `SKILL.md metadata.version` == `--version` == 1.6.3 | 通过 |
| 17–18 | 示例配置 | `assets/mapping.example.json` 能被真正解析，项目名正确 | 通过 |
| 19–20 | agent probe 信封 | 单行 JSON，`action=agent.probe`、`details` 为项目行、`stats.projects` 计数 | 通过 |
| 21 | JSON 全覆盖 | 除 `watch` 外 15 个 sync 动词都有 `--json` 分支；`watch --json` 报错退出 2 | 通过 |
| 22 | `clean-backup` 字段 | `applied / freed_bytes / kept / kept_bytes / deleted_batches` 齐全 | 通过 |
| 23 | 失败原因 | `promote` 模块名不存在时 `errors` 非空、退出 1 | 通过 |
| 24 | `init` 预览 | 不带 `--apply` 不写配置、不建 `src`，`dry=true` | 通过 |
| 25 | `init --apply` | 写入 `mapping.json` 与 `src/common`、项目目录 | 通过 |
| 26 | `init` 合并 | 同名项目更新字段，不追加重复项 | 通过 |
| 27 | `init` 目录不存在 | 退出 2，JSON `errors` 非空 | 通过 |
| 28 | 未知动词 | 退出 2，不把错词当项目名 | 通过 |
| 29 | `--version` | 打印 `vba-workbench 1.6.3`，退出 0 | 通过 |
| 30 | 裸跑 `sync.py` | 只打印帮助、退出 2，不出 JSON | 通过 |
| 31 | 裸跑 `audit.py` | 只打印帮助、退出 2 | 通过 |
| 32 | 菜单门禁 | 没有 `-i` 一律不进菜单（`--no-input` 裸跑也退出 2） | 通过 |
| 33 | agent style 失败 JSON | 非 0 且 `ok=false` 时 `errors` 非空 | 通过 |
| 34 | `watch` 参数注册 | `watch --help` 含 `--no-backup` | 通过 |
| 35 | `agent logs` 参数注册 | `logs --help` 含 `--tail` 与默认 40 | 通过 |
| 36 | 批量 push JSON | `push --all-project --json` 的 `details` 每项目一条 | 通过 |
| 37 | `shells` 预览字段 | `shells --json` 带 `stats.dry=true` | 通过 |
| 38 | 非 TTY 交互入口 | `sync.py -i` / `audit.py -i` 无 traceback、退出 1 | 通过 |
| 39 | `--no-input` 入口 | `sync list --no-input` 与 `audit members --no-input` 不 NameError | 通过 |
| 40 | 源码预检·空模块 | 非文档 `.bas/.cls` 空正文判致命；`SheetN` 文档模块豁免 | 通过 |
| 41 | 源码预检·缺 `.frx`/重名 | 非占位壳 `.frm` 缺同名 `.frx`、同目录模块重名都判致命 | 通过 |
| 42 | 源码预检·`VB_Name`/空目录 | 文件头与文件名不一致、源码目录不存在/无可写模块都判致命 | 通过 |
| 43 | 源码预检·软提醒 | BOM / CRLF / 缺 `Option Explicit` / 续行截断只提醒、不阻断 | 通过 |
| 44 | 回读校验·一致 | 指纹一致时 `postcheck` 通过 | 通过 |
| 45 | 回读校验·截断 | 明细给出源码行数 / VBE 行数 / 首个差异行 | 通过 |
| 46 | 回读校验·模块缺失 | 工程里找不到时给出明确原因，不误判为正文差异 | 通过 |

定向门禁：`python -m unittest tests.test_gen_commands_md tests.test_version_consistency`；语法自检 `python -m compileall -q scripts tests`。

## 手工评测（真实 Excel，空白表）

按 `tests/MANUAL.md` 在 `.lab/示例表A V100.xlsm` 上跑（默认预览 + `--apply`）：

| 步骤 | 判据 | 结果 |
|---|---|---|
| `list --json` | 项目解析，示例表命中 V100 | 通过 |
| `version --json` | common 版本一致、`mismatched=0` | 通过 |
| `shells --json` | 预览与创建数一致 | 通过 |
| `pull --all-project --json` | 预览全部项目、不碰 Excel | 通过 |
| `push --project 示例表A --json` | `target=.lab\示例表A V100.xlsm`、`modules=3`、`dry=true` | 通过 |
| `push --project 示例表A --apply` | 退出 0、`saved/backup` 均 true | 通过 |
| `run --project 示例表A --macro 示例求和 --apply --json` | `details[0].result == 5` | 通过 |
| run 宏名错误 | `errors` 有原因、退出 1 | 通过 |
| `init` 新项目预览/落盘 | 先预览、再 `--apply`，骨架正确 | 通过 |
| `audit.py indicators --json` | 通过数 = 检查数、0 警告、退出 0 | 通过 |
| `audit.py structure --json` | `issues=0`、退出 0 | 通过 |
| `agent.py probe --json` | 信封 + 项目行明细 | 通过 |
| `agent.py backup --project 示例表A` | 无 `--apply` 打印 `[dry] …`、退出 0 | 通过 |
| `doctor --json` | `failed=0`、`elapsed_ms>0`、必需文件齐全 | 通过 |

## 已知限制

- CI 跑不了写回 / 宏：GitHub 的 Windows runner 没有 Office，这两条只能手工验收（`tests/MANUAL.md`）。
- `watch` 不提供 JSON 事件流，只用于人本机常驻；agent/CI 用 `--seconds` 有界跑。
- 目标表必须有版本号（`V<纯数字>`）且 `require_version` 默认开启；无版本号按设计拒绝猜。
- 同一 Excel 会话不要并行两支写回。
- 缺宏 / 禁用宏统一收口 `RUN_FAILED`，区分信息只落 `hint`（B3 复核维持已接受设计差异，不恢复旧分流码）；看门狗公共原语（`toolkit.watchdog`）与弹窗识别族（`toolkit.dialogs`）已回补：超时后先按 PID 存活二次分类——已退出报 `EXCEL_CLOSED`（优先于超时码），存活再认弹窗，编译错误框报 `VBE_COMPILE_ERROR`（`stage=vbe`），其余模态框归 `WATCHDOG_TIMEOUT` 且 `hint` 点名窗口标题与正文片段（B1/B2/B4 闭环，见 `references/backlog-1.7.md`）。

## 结论

- 满足：包结构（`SKILL.md / scripts / references / assets / evals / tests`）、渐进披露（`SKILL.md` 122 行，细节按需读 `references/`）、单行 JSON 信封、版本三处一致、脱敏（`scrub.deny` 扫描交付件范围零命中；`CHANGELOG.md` / `tests/reports/` 的本机历史路径不入安装副本与发布包）、尺寸远低于 16 MB / 1024 文件上限。
- 待观察：CI 只有静态与纯函数测试，写回回归依赖手工清单；`evals/` 需在真实 agent 上定期跑。
