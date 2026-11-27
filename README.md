# vba-workbench

把 Excel 里的 VBA 抓成 `src/` 源码，改完再写回工作簿。写前备份，写完回读；Git 能看到每个模块改了什么，跑断言时有超时保护，不会一直等。

Windows + 微软 Excel · Python 3.10+ · v1.7.1 · SNSA-1.0

[3 分钟上手](#quickstart) · [日常怎么用](#workflow) · [它能做什么](#capabilities) · [交给 Agent](#agent) · [继续看文档](#docs)

## 这个工具解决什么事

VBA 代码藏在 `.xlsm` 的二进制工程里。Git 只能看到整个工作簿变没变，看不到某个模块改了什么；改代码要打开 VBE，改完还得点按钮验证；几个工作簿共用同一段模块时，只能靠复制，时间一长就各改各的。

vba-workbench 把 VBA 模块抓成 `src/` 下的文本文件。你用 VS Code、IDE 或 Agent 改这些文件，Git 正常 diff、审查和回滚；改完再由工具写回工作簿。写回前会备份，写完会回读比对，失败可以按备份回滚。

它不改你的 VBA 写法，也不碰工作簿里的业务数据。写回只更新 VBA 模块。

适合已经在用 `.xlsm`、又想用 Git 和编辑器维护 VBA 的人。WPS、Mac Excel 和非 Windows 环境用不了。

<a id="quickstart"></a>
## 3 分钟上手

### 1. 拿到完整目录

下载 [main 源码压缩包](https://github.com/kinrt/vba-workbench/archive/refs/heads/main.zip)，或克隆仓库：

```bash
git clone https://github.com/kinrt/vba-workbench.git
cd vba-workbench
```

只复制 `SKILL.md` 不够，`scripts/` 和 `references/` 要一起带着。

### 2. 安装 Python 依赖

```bash
python -m pip install -r requirements.txt
```

### 3. 打开 Excel 的 VBA 工程访问权限

在 Excel「信任中心 → 宏设置」勾选：

**信任对 VBA 工程对象模型的访问**

目标工作簿的 VBA 工程如果设了密码，先解除保护。

### 4. 跑一次环境体检

```bash
python scripts/sync.py doctor
```

最后看到“体检通过”就可以继续。第一次用建议只跑只读检查，不要急着写回。

### 5. 打开菜单，新建项目

```text
scripts\启动菜单.bat
```

在项目列表里选 **“＋ 新建项目（拖入 .xlsm）”**，把工作簿拖进窗口后回车；直接回车会打开文件选择框。

建好后会自动抓取一次源码，并进入这个项目的动作菜单。

<a id="workflow"></a>
## 日常怎么用

主入口还是 `scripts\启动菜单.bat`。菜单会列出 `mapping.json` 里的项目，选项目后再选动作。

一次写回大概是这样：

```text
【项目：示例表A · 选择动作】
1. 抓取源码 VBE→src
4. 写回源码 src→VBE
8. 常驻自动写回

选动作编号：4

—— 预览，未修改任何文件 ——
源码预检：17 个模块，致命 0，提醒 2
将写回：公共库 3 个 + 项目 14 个
目标工作簿：示例表A V3.xlsm
—— 预览结束，未修改任何文件 ——

回车 = 执行
0    = 返回
```

写操作会先出预览，确认后再回车。只读动作不改工作簿。动作结束后回到当前菜单，可以接着操作。

改完 `src/` 里的代码不会自动生效，要执行一次 “写回源码”，或者开 “常驻自动写回”。公共模块只在 `src/common` 维护，项目目录里的同名文件只是占位壳。

菜单里还有抓取未保存改动、换公共库占位壳、导出 GBK 副本、核对模块清单、清理旧备份这些动作，启动后都能看到。

<a id="capabilities"></a>
## 它能做什么

**日常同步**

- 抓取：把工作簿里的 VBA 模块抓进 `src/`，工作簿开着也能抓未保存的改动。
- 写回：把 `src/` 写回工作簿，先预览，再备份、写回、回读。
- 常驻写回：监听已打开的工作簿，改哪个模块就写回哪个模块；需要先设监听时长，默认 600 秒。

**检查和修改**

- 只读盘点：成员引用、工程引用、公共库体积、结构问题、模块差异、体检指标。
- 静态检查：格式检查、四条高置信 lint、过度设计体检。
- 批量修改：按清单打补丁、统一格式、补 / 清占位壳、把模块晋升到公共库。

**运行和交付**

- 运行宏：在已打开的工作簿里执行，或开隐藏实例执行不依赖界面的宏。
- 运行断言：读日志里的通过 / 失败计数；宏卡死或弹窗卡住会超时，不会一直等。
- 交付前：整工作簿备份、导出 GBK 副本、脱敏扫描、版本核对。

完整 34 项能力见 [references/capabilities.md](references/capabilities.md)。

<a id="agent"></a>
## 交给 Agent

把整个 `vba-workbench/` 目录放到 Agent 能读取的位置，`scripts/` 和 `references/` 也要在。只给一个 `SKILL.md` 跑不起来。

支持 Skill 的平台放进 `.skills/` 并刷新列表；不支持的平台直接把目录或 `SKILL.md` 路径交给它。

先拿一句只读任务试试有没有加载成功：

> 用 vba-workbench 先做环境体检和项目清单，只读，不要修改工作簿。

下任务时说清三件事：操作哪个工作簿、要完成什么、能不能写回。比如：

> 给示例表 A 新增批量导入功能：先只读梳理现有入口和模块，给我改动计划；确认后改 `src/` 里的代码，补一个自测宏，先预览写回，再运行断言。

Agent 的完整调用规则在 [SKILL.md](SKILL.md)，README 只讲怎么开始用。

<a id="faq"></a>
## 常见问题

**它和普通的 VBA 导出工具有什么区别？**

普通导出只把代码拿出来。这里还负责写回、备份、回读、运行宏和断言，改完能直接回到工作簿里验证。

**会改坏工作簿吗？**

写回前有预检和备份，写后有回读比对，失败可以按备份恢复。只读动作不修改工作簿。第一次用仍建议先预览，或先备份原文件。

**必须用 AI 吗？**

不必须。人可以只用启动菜单，IDE 可以只改 `src/`；Agent 只是批量修改和验证时更省事。

**支持 WPS 或 Mac Excel 吗？**

不支持。当前只支持 Windows + 微软 Excel。

<a id="docs"></a>
## 继续看文档

- 第一次用、想看更多例子：[`references/usage-examples.md`](references/usage-examples.md)
- 完整命令、参数、JSON 字段：[`references/commands.md`](references/commands.md)
- 抓取 / 写回的原理和数据流：[`references/sync.md`](references/sync.md)
- `mapping.json`、换版本、目录命名：[`references/sync-config.md`](references/sync-config.md)
- 修改内部代码、增加命令：[`references/sync-maintain.md`](references/sync-maintain.md)
- 变更分级、测试门禁、评审清单：[`references/change-standards.md`](references/change-standards.md)
- 调用链、分层、每个 Python 文件做什么：[`references/architecture.md`](references/architecture.md)
- 审计规则和结构检查：[`references/audit.md`](references/audit.md)
- 常见错误和处理：[`references/pitfalls.md`](references/pitfalls.md)
- 发布打包和验收：[`references/release.md`](references/release.md)

## 更新

替换整个 `vba-workbench/` 目录，保留工作区的 `mapping.json` 和 `src/`；刷新 Skill 列表后先做只读体检。打包、校验和脱敏按 [`references/release.md`](references/release.md)。

## 限制

- 只支持 Windows + 微软 Excel；WPS、Mac Excel 和非 Windows 环境不适用。
- 没有 Web 服务，也没有数据库，不负责联网协作。
- `*.xlsm` 不纳入版本库；`src/` 统一 UTF-8（无 BOM）+ LF。
- 不操作未登记工作簿；同一 Excel 实例不要并行写回。
- 发布前要脱敏：真实用户名、机器名、工作簿名、工作表名、客户名和经营数字都不进发布包。

## 许可证

`SNSA-1.0`：允许公开托管与分发，但不是 OSI 开源，禁止销售。完整条款见 [LICENSE](LICENSE)。
