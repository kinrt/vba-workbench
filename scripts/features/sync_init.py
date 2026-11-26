# -*- coding: utf-8 -*-
"""sync_push 切片内聚段：init 动词——生成/更新 mapping.json 与 src 骨架。

默认只预览；``--apply`` 才写盘。``--lab`` 通过 ``_run_lab_generator`` 接缝调用工具自带的
空白表生成脚本，测试替换此接缝，不起 Excel。
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from pathlib import Path

from toolkit.base import (PKG_ROOT, Failure as BaseFailure, entry_command)
from toolkit.contract import (ErrorCode, Failure, Outcome, Stage,
                               to_contract_failure)

_VERSION_RE = re.compile(r"[Vv](\d+)$")


def _stem_base(stem: str) -> tuple[str, bool]:
    """文件名带 ` V<数字>` 时返回去掉版本号的基名；否则原样返回。"""
    m = _VERSION_RE.search(stem)
    if not m:
        return stem, False
    return (stem[:m.start()].rstrip() or stem), True


def _run_lab_generator(script: Path, target: Path) -> None:
    """--lab 调用工具自带的空白表生成脚本；测试替换此接缝，不起 Excel。"""
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}
    result = subprocess.run(
        [sys.executable, str(script), str(target)],
        capture_output=True, text=True, encoding="utf-8",
        errors="replace", env=env,
    )
    if result.returncode:
        raise Failure(
            f"生成测试表失败：{(result.stderr or result.stdout)[-800:]}",
            code=ErrorCode.LAB_TEMPLATE_FAILED, stage="init",
            suggested="看上面的 Excel/COM 报错；确认本机装了 Excel 与 pywin32 后重跑 init --apply",
        )


def _is_within(roots: list[Path], target: Path) -> bool:
    """target 是否落在任一允许根之内（纯路径包含判定，不碰磁盘）。"""
    for base in roots:
        try:
            target.relative_to(base)
            return True
        except ValueError:
            continue
    return False


def init(ctx) -> Outcome:
    """生成/更新仓库根的 mapping.json，并建好 src 骨架；--apply 才写盘。"""
    args = ctx.args
    name = (getattr(args, "project", "") or "").strip()
    if not name:
        outcome = Outcome(action="sync.init", ok=False, exit_code=2)
        outcome.add_error(Failure("init 需要 --project <项目名>",
                                  code=ErrorCode.USAGE, stage=Stage.CONFIG))
        return outcome

    raw_prefix = (getattr(args, "prefix", "") or "").strip()
    lab = bool(getattr(args, "lab", False))
    explicit_config = bool((getattr(args, "config", "") or "").strip())
    allow_outside = bool(getattr(args, "allow_outside", False))
    default_cfg = _default_config_path()
    target_cfg = (Path(args.config).expanduser().resolve()
                  if explicit_config else default_cfg)
    base_dir = target_cfg.parent
    raw_file = (getattr(args, "file", "") or "").strip()
    raw_dir = (getattr(args, "dir", "") or "").strip()

    file_path: Path | None = None
    if raw_file:
        file_path = Path(raw_file).expanduser()
        if not file_path.is_absolute():
            file_path = base_dir / file_path
        file_path = file_path.resolve()
        if not file_path.is_file():
            outcome = Outcome(action="sync.init", ok=False, exit_code=2)
            outcome.add_error(Failure(
                f"文件不存在：{file_path}",
                code=ErrorCode.USAGE, stage=Stage.CONFIG))
            return outcome
        if file_path.suffix.lower() not in (".xlsm", ".xlsb"):
            outcome = Outcome(action="sync.init", ok=False, exit_code=2)
            outcome.add_error(Failure(
                f"只支持 .xlsm / .xlsb：{file_path.name}",
                code=ErrorCode.USAGE, stage=Stage.CONFIG))
            return outcome
        raw_dir = str(file_path.parent)

    dir_given = bool(raw_file or raw_dir)
    if not raw_dir:
        if not lab:
            outcome = Outcome(action="sync.init", ok=False, exit_code=2)
            outcome.add_error(Failure(
                "init 需要 --dir <表所在目录> 或 --file <工作簿>（--lab 时可省略，"
                "默认 <配置目录>/.lab）",
                code=ErrorCode.USAGE, stage=Stage.CONFIG))
            return outcome
        raw_dir = str(base_dir / ".lab")
    search = Path(raw_dir)
    if not search.is_absolute():
        search = base_dir / search
    search = search.resolve()
    if not lab and not search.is_dir():
        outcome = Outcome(action="sync.init", ok=False, exit_code=2)
        outcome.add_error(Failure(
            f"目录不存在：{search}（先建好表目录，或加 --lab 生成测试表）",
            code=ErrorCode.USAGE, stage=Stage.CONFIG))
        return outcome

    if file_path is not None:
        stem_base, versioned = _stem_base(file_path.stem)
        prefix = raw_prefix or stem_base
    else:
        versioned = False
        prefix = raw_prefix or name

    # 归属门禁：只有命令行显式 --config 或 --allow-outside 才豁免；环境变量
    # VBA_WORKBENCH_CONFIG 与 walk-up 发现的都是“当前生效配置”，不属于显式另选，
    # 仍受门禁保护。未豁免时，游离目录不得追加进当前生效配置；允许根 = 配置
    # workspace 目录 + 生效配置所在目录；其它目录一律拒绝，避免临时/演示项目写进
    # 生产 mapping.json。菜单拖入外部工作簿时显式传 --allow-outside。
    allowed_roots = [ctx.root, target_cfg.parent]
    if not explicit_config and not allow_outside and not _is_within(allowed_roots, search):
        local_cfg = search / "mapping.json"
        init_args = ["--config", local_cfg.as_posix(), "--project", name,
                      "--dir", search.as_posix()]
        if lab:
            init_args.append("--lab")
        if getattr(args, "apply", False):
            init_args.append("--apply")
        suggested = entry_command("sync.py", "init", *init_args)
        outcome = Outcome(action="sync.init", ok=False, exit_code=2)
        outcome.add_error(Failure(
            f"目标目录 {search} 不在当前配置工作区（{ctx.root}）或配置目录"
            f"（{target_cfg.parent}）之内，为避免污染生效配置 {target_cfg}，"
            f"已拒绝写入；临时/演示项目请改用独立配置：{suggested}",
            code=ErrorCode.USAGE, stage=Stage.CONFIG,
            hint="未指定 --config 时只允许登记工作区/配置目录内的项目；"
                 "游离目录必须显式指定独立 mapping.json",
            suggested=suggested))
        return outcome

    # --lab 未显式指定配置时，写 lab 目录自己的局部配置，不落到 env 指向的配置里。
    if not explicit_config and lab:
        target_cfg = search / "mapping.json"
        base_dir = target_cfg.parent

    src_rel = (getattr(args, "src", "") or "").strip()
    if not src_rel:
        src_rel = (f".lab/src/projects/{name}" if lab and not dir_given
                   else f"src/projects/{name}")
    proj = {
        "name": name,
        "search_dir": search.as_posix(),
        "prefix": prefix,
        "require_version": True,
        "src": src_rel,
        "note": f"文件名必须是「{prefix} V<数字>.xlsm」；目录与项目名不带版本号",
    }
    if file_path is not None and not versioned:
        proj["file"] = file_path.as_posix()
        proj["require_version"] = False
        proj["note"] = f"按文件精确登记：{file_path.name}；不再按版本号自动选表"

    try:
        config_existed = target_cfg.exists()
        if config_existed:
            try:
                data = json.loads(target_cfg.read_text(encoding="utf-8"))
            except json.JSONDecodeError as e:
                raise Failure(f"{target_cfg} 不是合法 JSON：{e}",
                              code=ErrorCode.CONFIG_INVALID, stage=Stage.CONFIG)
        else:
            data = {"defaults": {"common_dir": "src/common", "backup_dir": ".backup",
                                  "import_document_modules": False, "backup_auto_clean": True,
                                  "backup_keep": 10, "backup_keep_daily": 0,
                                  "backup_min_keep": 3, "backup_protect": ["-覆盖前"],
                                  "require_version": True},
                    "projects": []}
        if isinstance(data, list):
            data = {"projects": data}
        data.setdefault("projects", [])
        idx = next((i for i, p in enumerate(data["projects"]) if p.get("name") == name), None)
        if idx is None:
            data["projects"].append(proj)
        else:
            merged = {**data["projects"][idx], **proj}
            if versioned:
                merged.pop("file", None)
                merged["require_version"] = True
            data["projects"][idx] = merged

        dry = not bool(getattr(args, "apply", False))
        human: list[str] = []
        if dry:
            human.append(f"[dry] 将写入 {target_cfg}")
            if file_path is not None and not versioned:
                human.append(f"       项目 {name}：按文件精确登记 {file_path} → {src_rel}")
            else:
                human.append(f"       项目 {name}：{proj['search_dir']} / {prefix} V<数字>.xlsm → {src_rel}")
            human.append("       加 --apply 才写入文件")
        else:
            target_cfg.parent.mkdir(parents=True, exist_ok=True)
            target_cfg.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n",
                                   encoding="utf-8", newline="\n")
            human.append(f"已写入 {target_cfg}")
            for d in (base_dir / "src" / "common", base_dir / src_rel):
                d.mkdir(parents=True, exist_ok=True)
            if lab:
                search.mkdir(parents=True, exist_ok=True)
                target = search / f"{prefix} V1.xlsm"
                if target.exists():
                    human.append(f"测试表已存在：{target}")
                else:
                    script = PKG_ROOT / "scripts" / "make_lab_workbook.py"
                    if not script.is_file():
                        raise Failure(
                            f"生成测试表的脚本缺失：{script}",
                            code=ErrorCode.MISSING_LAB_TEMPLATE, stage="init",
                            suggested="安装包应包含 scripts/make_lab_workbook.py；"
                                      "重装 vba-workbench，或手工建一张空白表后再 init --apply",
                        )
                    _run_lab_generator(script, target)
                    human.append(f"已生成测试表：{target}")
    except BaseFailure as e:
        return Outcome(action="sync.init", ok=False, details=[],
                       errors=[to_contract_failure(e)],
                       exit_code=1, elapsed_ms=ctx.elapsed_ms())

    return Outcome(
        action="sync.init",
        ok=True,
        details=[proj],
        stats={
            "results": [proj],
            "config": str(target_cfg),
            "dry": dry,
            "created_src": [] if dry else ["src/common", src_rel],
        },
        human=human,
        exit_code=0,
        elapsed_ms=ctx.elapsed_ms(),
    )


def _default_config_path() -> Path:
    """配置缺省位置；延迟读取 toolkit.base.CONFIG_PATH，避免导入期绑定。"""
    from toolkit.base import CONFIG_PATH
    return CONFIG_PATH
