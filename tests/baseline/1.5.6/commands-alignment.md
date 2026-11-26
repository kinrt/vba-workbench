# 命令对齐清单（1.5.6 argparse vs 1.6 manifest）

- 旧版动词：31；新版已登记：31；待迁移：0
- 口径：参数集合已剔除骨架公共位（--json/--no-input/--apply/--project/--all-project）

## sync.py（sync）

| 动词 | 旧参数 | 新参数 | 旧 --json | 新 --json | 新退出码 | 结论 |
|---|---|---|---|---|---|---|
| `align` | — | — | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `clean-backup` | keep, keep-tag | keep, keep-tag | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `doctor` | — | verify-package | 是 | 是 | 0/1/2 | 存在差异（见差分报告） |
| `gbk-export` | — | — | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `init` | config, dir, lab, prefix, src | config, dir, lab, prefix, src | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `list` | — | — | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `plan` | intent | intent | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `promote` | module | module | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `pull` | no-backup | no-backup | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `pull-open` | no-backup | no-backup | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `push` | doc-modules, no-backup, no-save, open-only, rollback-on-fail | doc-modules, no-backup, no-save, open-only, rollback-on-fail | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `run` | arg, isolate, log, log-optional, macro, push, rollback-on-fail, save, tail | arg, isolate, log, log-optional, macro, push, rollback-on-fail, save, tail | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `shells` | mode | mode | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `vbamod` | — | — | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `version` | — | — | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `watch` | doc-modules, interval, no-backup, no-save, seconds, startup | doc-modules, interval, no-backup, no-save, seconds, startup | 否 | 否 | 0/1/2 | 对齐（无差异） |

## audit.py（audit）

| 动词 | 旧参数 | 新参数 | 旧 --json | 新 --json | 新退出码 | 结论 |
|---|---|---|---|---|---|---|
| `common` | all, exclude, md, out | all, exclude, md, out | 是 | 是 | 0/1/2/3 | 对齐（无差异） |
| `indicators` | all, exclude, ignore, md, out, strict | all, exclude, ignore, md, out, strict | 是 | 是 | 0/1/2/3 | 对齐（无差异） |
| `members` | all, exclude, md, module, out | all, exclude, md, module, out | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `overview` | all, exclude, ignore, md, out, strict | all, exclude, ignore, md, out, strict | 是 | 是 | 0/1/2/3 | 对齐（无差异） |
| `project` | all, exclude, md, out | all, exclude, md, out | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `structure` | all, exclude, ignore, md, out, strict | all, exclude, ignore, md, out, strict | 是 | 是 | 0/1/2/3 | 对齐（无差异） |

## agent.py（agent）

| 动词 | 旧参数 | 新参数 | 旧 --json | 新 --json | 新退出码 | 结论 |
|---|---|---|---|---|---|---|
| `assert` | baseline, clear, log, macro, sheet, timeout | baseline, clear, log, macro, sheet, timeout | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `backup` | tag | tag | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `book` | action | action | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `logs` | log, tail | log, tail | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `metrics` | out | out | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `patch` | file, no-style | file, no-style | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `probe` | strict | strict | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `push` | — | — | 是 | 是 | 0/1/2 | 对齐（无差异） |
| `style` | action, scope, which | action, scope, which | 是 | 是 | 0/1/2 | 对齐（无差异） |
