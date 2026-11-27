# 命令表（由 manifest 生成，请勿手改）

> 本文件由 `tests/tools/gen_commands_md.py` 从 `scripts/toolkit/commands.py` 的
> manifest 生成；动词、参数、退出码以登记为准。重新生成：

```
python tests/tools/gen_commands_md.py
```

公共参数（各动词由骨架统一附加，不在下表重复）：

- `--no-input`：非交互，缺参直接失败；
- `--apply`：写操作才有的写入开关，不加只预览；
- `--project` / `--all-project`：项目选择，取值见各行；
- `--json`：输出一行 JSON 信封；标注「否」的动词显式拒绝该开关。

已登记动词合计：35 个。

### sync.py（sync）

| 动词 | 参数 | 变更模式 | 退出码 | --json | 对应测试 | 参考 |
|---|---|---|---|---|---|---|
| `align`<br>把项目里与公共库同名的旧真身换成占位壳 | `--project` 可选 | 默认预览，--apply 才写入文件 | 0/1/2 | 是 | transform.align.success, transform.align.not_found, transform.align.usage, transform.align.json | references/sync-commands.md#align |
| `clean-backup`<br>清理 .backup 旧批次（默认只预览） | `--keep`（整数）；`--keep-tag`（可重复） | 默认预览，--apply 才写入文件 | 0/1/2 | 是 | transform.clean_backup.success, transform.clean_backup.bad_config, transform.clean_backup.usage, transform.clean_backup.json | references/sync-commands.md#clean-backup |
| `diff`<br>只读比较源码与工程模块（新增/修改/仅工程/一致） | `--all`；`--md`（路径）；`--out`（路径）；`--project` 必填 | 只读（不改代码） | 0/1/2 | 是 | tests/test_sync_diff.py | references/sync.md#diff |
| `doctor`<br>环境体检（不合格退出 1） | `--verify-package` | 只读（不改代码） | 0/1/2 | 是 | contract.doctor.success, contract.doctor.failed, contract.doctor.verify-package | references/sync-commands.md#doctor |
| `gbk-export`<br>导出 src 的 GBK 副本到 .backup/GBK（不改 src） | `--project`/`--all-project` 二选一 | 默认预览，--apply 才写入文件 | 0/1/2 | 是 | transform.gbk_export.success, transform.gbk_export.not_found, transform.gbk_export.usage, transform.gbk_export.json | references/sync-commands.md#gbk-export |
| `init`<br>生成/更新 mapping.json 与 src 骨架（默认只预览） | `--dir`；`--file`（路径）；`--prefix`；`--src`；`--config`；`--allow-outside`；`--lab`；`--project` 必填 | 默认预览，--apply 才写入文件 | 0/1/2 | 是 | contract.init.preview, contract.init.apply | references/sync-commands.md#init |
| `list`<br>每个项目解析到哪个 xlsm | — | 只读（不改代码） | 0/1/2 | 是 | contract.list.success, contract.list.target-not-found | references/sync-commands.md#list |
| `plan`<br>只读推荐：按当前状态选 open/disk/isolate/open_required/unavailable 模式 | `--intent`（pull / push / run / watch）；`--project` 必填 | 只读（不改代码） | 0/1/2 | 是 | contract.plan.success, contract.plan.target_missing, contract.plan.usage, contract.plan.json | references/sync-commands.md#plan |
| `promote`<br>把一个模块从项目源码搬进公共库 | `--module`；`--project` 必填 | 默认预览，--apply 才写入文件 | 0/1/2 | 是 | transform.promote.success, transform.promote.not_found, transform.promote.usage, transform.promote.json | references/sync-commands.md#promote |
| `pull`<br>VBE→src：工作簿已打开时借用会话抓取，未打开时从磁盘抓取 | `--no-backup`；`--project`/`--all-project` 二选一 | 默认预览，--apply 才写入文件 | 0/1/2 | 是 | contract.pull.success, contract.pull.target_missing, contract.pull.usage, contract.pull.json | references/sync-commands.md#pull |
| `pull-open`<br>VBE→src：从已打开的 Excel 抓（含未保存改动） | `--no-backup`；`--project` 必填 | 默认预览，--apply 才写入文件 | 0/1/2 | 是 | contract.pull-open.success, contract.pull-open.target_missing, contract.pull-open.usage, contract.pull-open.json | references/sync-commands.md#pull-open |
| `push`<br>src→VBE：写回工作簿（工作簿未打开时自动开隐藏实例写回） | `--doc-modules`；`--no-save`；`--no-backup`；`--open-only`；`--rollback-on-fail`；`--project`/`--all-project` 二选一 | 写操作，必须 --apply | 0/1/2 | 是 | contract.push.success, contract.push.target_not_open | references/sync-commands.md#push |
| `run`<br>写回（可选）→运行宏→取返回值与日志尾部 | `--macro`；`--arg`（可重复）；`--push`；`--isolate`；`--save`；`--tail`（整数）；`--log`；`--log-optional`；`--rollback-on-fail`；`--project` 必填 | 主操作只读，附带开关写入文件 | 0/1/2 | 是 | contract.run.success, contract.run.failed, contract.run.usage, contract.run.json | references/sync-commands.md#run |
| `scrub`<br>按 scrub.deny 扫交付件，命中退出 1 | `--path`（路径） | 只读（不改代码） | 0/1/2 | 是 | contract.scrub.success, contract.scrub.hit, contract.scrub.config-invalid | references/sync-commands.md#scrub |
| `shells`<br>公共库占位壳（同名一行注释，写回时跳过） | `--mode`（create / clean） | 默认预览，--apply 才写入文件 | 0/1/2 | 是 | contract.shells.create, contract.shells.clean | references/sync-commands.md#shells |
| `vbamod`<br>核对各工作簿的模块清单与正文（和最近一次备份比） | `--project` 可选 | 只读（不改代码） | 0/1/2 | 是 | contract.vbamod.success, contract.vbamod.not_found, contract.vbamod.usage, contract.vbamod.json | references/sync-commands.md#vbamod |
| `version`<br>核对公共库版本与各项目锁定版本 | — | 只读（不改代码） | 0/1/2 | 是 | contract.version.success, contract.version.mismatch | references/sync-commands.md#version |
| `watch`<br>常驻监听：源码一变就写回已打开的工程 | `--interval`（小数）；`--startup`（ask / src / vbe / each / skip）；`--no-save`；`--no-backup`；`--doc-modules`；`--seconds`（整数）；`--project` 可选 | 常驻运行（不支持 --json） | 0/1/2 | 否 | contract.watch.success, contract.watch.failed, contract.watch.usage, contract.watch.reject_json | references/sync-commands.md#watch |

### audit.py（audit）

| 动词 | 参数 | 变更模式 | 退出码 | --json | 对应测试 | 参考 |
|---|---|---|---|---|---|---|
| `common`<br>公共库全部模块的引用数/体积/建议 | `--project`（可重复）；`--exclude`（可重复）；`--all`；`--md`（路径）；`--out`（路径） | 只读（不改代码） | 0/1/2/3 | 是 | audit_a.common.success, audit_a.common.not_configured_exit3, audit_a.common.usage, audit_a.common.json_pure | references/audit.md#common |
| `indicators`<br>体检指标 H01–H19 复跑 | `--exclude`（可重复）；`--ignore`（可重复）；`--strict`；`--all`；`--md`；`--out` | 只读（不改代码） | 0/1/2/3 | 是 | tests/test_audit_b.py | references/audit.md |
| `lint`<br>四条高置信静态 lint（只读，默认 warning 退出 3） | `--exclude`（可重复）；`--strict`；`--all`；`--md`（路径）；`--out`（路径）；`--project` 必填 | 只读（不改代码） | 0/1/2/3 | 是 | tests/test_audit_lint.py | references/audit.md#lint |
| `members`<br>某模块被谁调用（默认展开 public 之外的成员） | `--module`；`--exclude`（可重复）；`--all`；`--md`（路径）；`--out`（路径） | 只读（不改代码） | 0/1/2 | 是 | audit_a.members.success, audit_a.members.not_found, audit_a.members.usage, audit_a.members.json_pure | references/audit.md#members |
| `overview`<br>公共库 + 结构的汇总意见 | `--project`（可重复）；`--exclude`（可重复）；`--ignore`（可重复）；`--strict`；`--all`；`--md`；`--out` | 只读（不改代码） | 0/1/2/3 | 是 | tests/test_audit_b.py | references/audit.md |
| `project`<br>某个工作簿的依赖矩阵 | `--exclude`（可重复）；`--all`；`--md`（路径）；`--out`（路径）；`--project` 必填 | 只读（不改代码） | 0/1/2 | 是 | audit_a.project.success, audit_a.project.not_found, audit_a.project.usage, audit_a.project.json_pure | references/audit.md#project |
| `references`<br>只读检查 VBA 工程引用（缺失/损坏/路径；默认警告退出 3） | `--strict`；`--md`（路径）；`--out`（路径）；`--project` 必填 | 只读（不改代码） | 0/1/2/3 | 是 | tests/test_audit_references.py | references/audit.md#references |
| `structure`<br>旧副本 / 孤儿壳 / 缺壳 / 重名过程 / 文档模块对账 / 旧规则回退 | `--exclude`（可重复）；`--ignore`（可重复）；`--strict`；`--all`；`--md`；`--out` | 只读（不改代码） | 0/1/2/3 | 是 | tests/test_audit_b.py | references/audit.md |

### agent.py（agent）

| 动词 | 参数 | 变更模式 | 退出码 | --json | 对应测试 | 参考 |
|---|---|---|---|---|---|---|
| `assert`<br>运行断言宏并核对通过/失败计数（看门狗硬超时） | `--macro`；`--sheet`；`--clear`；`--baseline`（整数）；`--timeout`（小数）；`--log`；`--project` 可选 | 写操作，必须 --apply | 0/1/2 | 是 | contract.assert.success, contract.assert.no_table, contract.assert.no_log, contract.assert.timeout, contract.assert.usage, contract.assert.json | references/agent.md |
| `backup`<br>整工作簿备份到 .backup/<时间戳>-agent | `--tag`；`--project` 可选 | 写操作，必须 --apply | 0/1/2 | 是 | contract.backup.success, contract.backup.target-not-found | references/agent.md |
| `book`<br>开/关工作簿 | `--action`（open / close）；`--project` 可选 | 写操作，必须 --apply | 0/1/2 | 是 | contract.book.open, contract.book.close-not-open | references/agent.md |
| `logs`<br>最新测试日志摘要 | `--log`；`--tail`（整数）；`--project` 可选 | 只读（不改代码） | 0/1/2 | 是 | contract.logs.success, contract.logs.not-found | references/agent.md |
| `metrics`<br>源码体检基线（文件/过程/长度/薄包装/模块变量/布尔/On Error 规则） | `--out` | 主操作只读，附带开关写入文件 | 0/1/2 | 是 | contract.metrics.success, contract.metrics.usage, contract.metrics.json | references/agent.md#源码体检基线metrics |
| `patch`<br>按整行块清单批量替换源码（dry-run 只核对命中） | `--file`；`--no-style`；`--project` 可选 | 写操作，必须 --apply | 0/1/2 | 是 | contract.patch.success, contract.patch.format, contract.patch.ambiguous, contract.patch.json | references/agent.md#按清单打补丁patch |
| `probe`<br>项目状态一览：磁盘/已打开/工程/模块数/测试工作表行数 | `--strict`；`--project` 可选 | 只读（不改代码） | 0/1/2 | 是 | contract.probe.success, contract.probe.target-not-found | references/agent.md |
| `push`<br>src→VBE 写回工作簿 | `--doc-modules`；`--no-save`；`--no-backup`；`--open-only`；`--rollback-on-fail`；`--project` 可选 | 写操作，必须 --apply | 0/1/2 | 是 | contract.push.preview, contract.push.project-not-found | references/agent.md |
| `style`<br>风格检查/修复（check 只读，fix 要 --apply） | `--action`（check / fix）；`--which`（all / eol / call / explicit）；`--scope`；`--project` 可选 | 默认预览，--apply 才写入文件 | 0/1/2 | 是 | contract.style.success, contract.style.failure, contract.style.usage, contract.style.json | references/agent.md#风格检查--修复style-子命令 |
