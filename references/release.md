# 发布检查清单

发布 = 打一个 zip 包放到 `releases/`，并在仓库里留一条 Git 记录。顺序不能颠倒。

## 1. 版本三处一致

```
python vba-workbench/scripts/sync.py doctor
```

- `VERSION`、`SKILL.md metadata.version`、`CHANGELOG.md` 顶部段落，三处同一个号。
- `python -m unittest discover -s vba-workbench/tests` 全过（含版本一致自测）。Python 3.13 要求起点可导入；实际用：

  ```
  cd vba-workbench/tests
  python -m unittest discover -s . -p "test_*.py"
  ```

  当前基线 626 例（1 例按环境跳过）。
- 安装副本必须带 `scripts/make_lab_workbook.py`（不能依赖 `tests/`）：在安装目录跑一次
  `python scripts/sync.py version --json`（必须是单行 `sync.version` 信封）和
  `python scripts/sync.py init --project 临时 --lab --apply --json`（能生成空白工作簿；缺脚本才是
  `MISSING_LAB_TEMPLATE/stage=init`）。临时项目和 `.lab` 验证完删掉。

## 2. 脱敏

```
python vba-workbench\scripts\sync.py scrub --json
```

- 读仓库根 `mapping.json` 的 `scrub.deny`，按安装器交付件白名单扫 `vba-workbench/**` 里的文本（`.bat` 由 `read_source_text` 回退 GBK；图片等非文本跳过）。
- 命中即阻断：真实路径、真实工作簿名、设备名、带 `V2611xx` 的版本号都不许出现；`errors[0].code=SCRUB_HIT`，退出 1。
- 默认扫源码包；要复核已安装副本，用 `--path <安装目录>`。`mapping.local.json`、`.lab/`、`tests/`、`evals/`、`releases/` 不入扫描范围。

## 3. 打包

一条命令按安装器交付白名单收集 82 件、校验上限并生成 zip；默认只预览，带 `--apply` 才写：

```
python tests/tools/make_release.py
python tests/tools/make_release.py --apply
```

- 根目录名 = skill 名（`vba-workbench/`），不要多套一层。- 排除 `__pycache__`、`*.pyc`、`*.xlsm`。
- 上限：单文件 ≤2 MB、整包 ≤16 MB、文件 ≤1024 个。
- 打包完成后用 `.installed.json` 对账：条目数、缺件、多余件、逐件 sha256 必须全一致。

## 4. 校验和 + 指针

`make_release.py --apply` 会算出 zip 的 SHA256，追加到 `releases/SHA256SUMS.txt`（同版本旧行先删），并把 `latest.txt` 写成当前文件名。要手工复核：

```
certutil -hashfile releases\vba-workbench-1.7.1.zip SHA256
```

交付时先发 `latest.txt` 指的那份，三份旧包并存容易拿错。

## 5. 记录

- `git commit` 消息写清「版本 + 主要改动 + 行为变更」。
- `src/common/待办清单.md` 追加一条 T 记录（做了什么、实测结论）。
