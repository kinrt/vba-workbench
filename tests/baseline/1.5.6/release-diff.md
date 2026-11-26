# 旧新差分报告（1.5.6 vs 当前工作树）

统计：match 38 / approved 2 / pending 0 / failed 0

| 命令 | 输入 | 旧输出 | 新输出 | 差异结论 |
|---|---|---|---|---|
| sync.align | — | {"args": [], "json": true, "project": true, "all_project": false} | {"args": [], "json": true, "project": true, "all_project": false} | 等价｜surface |
| sync.clean-backup | — | {"args": ["keep", "keep-tag"], "json": true, "project": false, "all_project": false} | {"args": ["keep", "keep-tag"], "json": true, "project": false, "all_project": false} | 等价｜surface |
| sync.doctor | — | [] | ["verify-package"] | 白名单内有意变更｜args |
| sync.doctor | — | null | {"names": ["verify-package"], "kind": "flag", "required": false, "choices": [], "default": false} | 白名单内有意变更｜arg-added |
| sync.doctor | — | {"args": [], "json": true, "project": false, "all_project": false} | {"args": ["verify-package"], "json": true, "project": false, "all_project": false} | 等价｜surface |
| sync.gbk-export | — | {"args": [], "json": true, "project": true, "all_project": true} | {"args": [], "json": true, "project": true, "all_project": true} | 等价｜surface |
| sync.init | — | {"args": ["config", "dir", "lab", "prefix", "src"], "json": true, "project": true, "all_project": false} | {"args": ["config", "dir", "lab", "prefix", "src"], "json": true, "project": true, "all_project": false} | 等价｜surface |
| sync.list | — | {"args": [], "json": true, "project": false, "all_project": false} | {"args": [], "json": true, "project": false, "all_project": false} | 等价｜surface |
| sync.plan | — | {"args": ["intent"], "json": true, "project": true, "all_project": false} | {"args": ["intent"], "json": true, "project": true, "all_project": false} | 等价｜surface |
| sync.promote | — | {"args": ["module"], "json": true, "project": true, "all_project": false} | {"args": ["module"], "json": true, "project": true, "all_project": false} | 等价｜surface |
| sync.pull | — | {"args": ["no-backup"], "json": true, "project": true, "all_project": true} | {"args": ["no-backup"], "json": true, "project": true, "all_project": true} | 等价｜surface |
| sync.pull-open | — | {"args": ["no-backup"], "json": true, "project": true, "all_project": false} | {"args": ["no-backup"], "json": true, "project": true, "all_project": false} | 等价｜surface |
| sync.push | — | {"args": ["doc-modules", "no-backup", "no-save", "open-only", "rollback-on-fail"], "json": true, "project": true, "all_project": true} | {"args": ["doc-modules", "no-backup", "no-save", "open-only", "rollback-on-fail"], "json": true, "project": true, "all_project": true} | 等价｜surface |
| sync.run | — | {"args": ["arg", "isolate", "log", "log-optional", "macro", "push", "rollback-on-fail", "save", "tail"], "json": true, "project": true, "all_project": false} | {"args": ["arg", "isolate", "log", "log-optional", "macro", "push", "rollback-on-fail", "save", "tail"], "json": true, "project": true, "all_project": false} | 等价｜surface |
| sync.shells | — | {"args": ["mode"], "json": true, "project": false, "all_project": false} | {"args": ["mode"], "json": true, "project": false, "all_project": false} | 等价｜surface |
| sync.vbamod | — | {"args": [], "json": true, "project": true, "all_project": false} | {"args": [], "json": true, "project": true, "all_project": false} | 等价｜surface |
| sync.version | — | {"args": [], "json": true, "project": false, "all_project": false} | {"args": [], "json": true, "project": false, "all_project": false} | 等价｜surface |
| sync.watch | — | {"args": ["doc-modules", "interval", "no-backup", "no-save", "seconds", "startup"], "json": false, "project": true, "all_project": false} | {"args": ["doc-modules", "interval", "no-backup", "no-save", "seconds", "startup"], "json": false, "project": true, "all_project": false} | 等价｜surface |
| audit.common | — | {"args": ["all", "exclude", "md", "out"], "json": true, "project": true, "all_project": false} | {"args": ["all", "exclude", "md", "out"], "json": true, "project": true, "all_project": false} | 等价｜surface |
| audit.indicators | — | {"args": ["all", "exclude", "ignore", "md", "out", "strict"], "json": true, "project": false, "all_project": false} | {"args": ["all", "exclude", "ignore", "md", "out", "strict"], "json": true, "project": false, "all_project": false} | 等价｜surface |
| audit.members | — | {"args": ["all", "exclude", "md", "module", "out"], "json": true, "project": false, "all_project": false} | {"args": ["all", "exclude", "md", "module", "out"], "json": true, "project": false, "all_project": false} | 等价｜surface |
| audit.overview | — | {"args": ["all", "exclude", "ignore", "md", "out", "strict"], "json": true, "project": true, "all_project": false} | {"args": ["all", "exclude", "ignore", "md", "out", "strict"], "json": true, "project": true, "all_project": false} | 等价｜surface |
| audit.project | — | {"args": ["all", "exclude", "md", "out"], "json": true, "project": true, "all_project": false} | {"args": ["all", "exclude", "md", "out"], "json": true, "project": true, "all_project": false} | 等价｜surface |
| audit.structure | — | {"args": ["all", "exclude", "ignore", "md", "out", "strict"], "json": true, "project": false, "all_project": false} | {"args": ["all", "exclude", "ignore", "md", "out", "strict"], "json": true, "project": false, "all_project": false} | 等价｜surface |
| agent.assert | — | {"args": ["baseline", "clear", "log", "macro", "sheet", "timeout"], "json": true, "project": true, "all_project": false} | {"args": ["baseline", "clear", "log", "macro", "sheet", "timeout"], "json": true, "project": true, "all_project": false} | 等价｜surface |
| agent.backup | — | {"args": ["tag"], "json": true, "project": true, "all_project": false} | {"args": ["tag"], "json": true, "project": true, "all_project": false} | 等价｜surface |
| agent.book | — | {"args": ["action"], "json": true, "project": true, "all_project": false} | {"args": ["action"], "json": true, "project": true, "all_project": false} | 等价｜surface |
| agent.logs | — | {"args": ["log", "tail"], "json": true, "project": true, "all_project": false} | {"args": ["log", "tail"], "json": true, "project": true, "all_project": false} | 等价｜surface |
| agent.metrics | — | {"args": ["out"], "json": true, "project": false, "all_project": false} | {"args": ["out"], "json": true, "project": false, "all_project": false} | 等价｜surface |
| agent.patch | — | {"args": ["file", "no-style"], "json": true, "project": true, "all_project": false} | {"args": ["file", "no-style"], "json": true, "project": true, "all_project": false} | 等价｜surface |
| agent.probe | — | {"args": ["strict"], "json": true, "project": true, "all_project": false} | {"args": ["strict"], "json": true, "project": true, "all_project": false} | 等价｜surface |
| agent.push | — | {"args": [], "json": true, "project": true, "all_project": false} | {"args": [], "json": true, "project": true, "all_project": false} | 等价｜surface |
| agent.style | — | {"args": ["action", "scope", "which"], "json": true, "project": true, "all_project": false} | {"args": ["action", "scope", "which"], "json": true, "project": true, "all_project": false} | 等价｜surface |
| sync.— | — | {"exit_code": 2, "envelope_keys": []} | {"exit_code": 2, "envelope_keys": []} | 等价｜scenario |
| sync.— | __baseline_no_such_verb__ | {"exit_code": 2, "envelope_keys": []} | {"exit_code": 2, "envelope_keys": []} | 等价｜scenario |
| audit.— | — | {"exit_code": 2, "envelope_keys": []} | {"exit_code": 2, "envelope_keys": []} | 等价｜scenario |
| audit.— | __baseline_no_such_verb__ | {"exit_code": 2, "envelope_keys": []} | {"exit_code": 2, "envelope_keys": []} | 等价｜scenario |
| agent.— | — | {"exit_code": 2, "envelope_keys": []} | {"exit_code": 2, "envelope_keys": []} | 等价｜scenario |
| agent.— | __baseline_no_such_verb__ | {"exit_code": 2, "envelope_keys": []} | {"exit_code": 2, "envelope_keys": []} | 等价｜scenario |
| sync.version | version --json | {"exit_code": 1, "envelope_keys": ["action", "details", "elapsed_ms", "errors", "exit_code", "ok", "stats", "ts"]} | {"exit_code": 1, "envelope_keys": ["action", "details", "elapsed_ms", "errors", "exit_code", "ok", "stats", "ts"]} | 等价｜scenario |
