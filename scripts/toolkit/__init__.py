# -*- coding: utf-8 -*-
"""
toolkit —— sync.py 与 audit.py 共用的通用层
============================================================
规矩（改这里之前先读）：

  1. 功能文件（sync.py / audit.py）只写业务；通用逻辑一律进本子包对应模块，
     两个入口脚本都只 `from toolkit import ...`，**不互相 import**。
  2. 子模块依赖方向：base ← config / repo / files / vba_parse ← ui ← cli；
     本文件统一再导出公共 API，入口脚本不 import 子模块细节。
  3. 编码约定：本包 .py 自身 UTF-8 无 BOM + LF；VBA 源码一律 UTF-8（无 BOM）+ LF
     （写入统一走 files.write_source）。只有写回 VBE 前的 GBK 临时副本与 .backup/GBK
     镜像用 CRLF——GBK/CRLF 只存在于 VBE 交换边界，走 files.transcode_exported_file /
     files.gbk_import_copy。
  4. 菜单一律 ui.run_menu；清单一律 ui.pick_list / pick_list_default（回车=默认项）/
     pick_grouped；输入一律 ui.read_line / choose / ask_yes；输出一律 msg / warn / fail；
     文件读写 / 备份 / 移动 / 编码转换一律 files 底层原语。
     **入口脚本不许自己写 for + input + 重问的编号循环**——那正是本包存在的理由。
     每条规矩的完整版与"什么该下沉到这里"见同目录上级的 references/toolkit.md。
  5. 文件名 ≠ 模块名：文档模块的源码叫 `Sheet4(SKU统计).bas`（表页名放进括号，和 VBE 左侧
     显示对齐）。要模块名一律 repo.module_name_of(path)，生成文件名用 repo.doc_stem(...)，
     显示表页名用 repo.doc_label_of(...)；入口脚本里禁止拿 f.stem 去比模块名。
"""

from .base import (APP_CONFIG, CODE_ROOT, CONFIG_ENV, CONFIG_PATH,
                   PKG_ROOT, ROOT, SHELL_HINT, SOURCE_ALL_EXTS, SOURCE_EXTS, TEXT_ENCODINGS,
                   Failure, abs_path, cached_warnings, entry_command, make_envelope, fail,
                   find_config, interactive_ok, msg,
                   read_source_text, reset_warnings, set_non_interactive,
                   set_tool_name, warn,
                   assume_yes, json_mode, set_assume_yes, set_json_mode)
from .check import (NOTE_LEVELS, ModuleSetDiff, Notes, check_doc_modules,
                    check_export_set, check_module_set, check_sheet_modules,
                    check_suffix_modules)
from .cli import (add_interactive_arg, add_no_input_arg, apply_no_input, ensure_interactive,
                  entry, run_cli)
from .config import config_source, load_config, load_raw_config, save_config
from .files import (append_line, backup_copy, backup_dirs, backup_stats, dir_size_bytes,
                    dump_json, gbk_import_copy,
                    latest_backup, move_into, newest_file, prune_backups, read_tail,
                    remove_tree, stamp_dir, transcode_exported_file, write_source,
                    write_text_utf8)
from .watch import changed as watch_changed
from .watch import root_of as watch_root_of
from .watch import snap as watch_snapshot
# watchdog（线程看门狗 + PID 存活探测；win32 调用在函数内延迟导入）
from .watchdog import WatchdogTimeout, pid_alive, pid_from_hwnd, run_with_watchdog
# dialogs（弹窗/模态识别：纯函数分类 + win32gui 延迟枚举）
from .dialogs import (DialogMatch, KIND_COMPILE_ERROR, KIND_MODAL_DIALOG,
                      KIND_UNKNOWN, KIND_VBE_DIALOG, WindowInfo, detect_dialog,
                      detect_target_dialog, enumerate_windows)
# excel 是 COM 边界层：本文件顶层不 import win32com（全在函数里延迟导入），
# 所以没装 pywin32 / 没有 Excel 的机器上 import toolkit 照样成功，audit/style 不受影响。
from .excel import (ExcelSession, attach_running_excel, backup_workbook, classify_com_failure, com_call,
                    component_snapshot, ensure_vbe_access, find_component,
                    find_excel_with_workbook, find_open_workbook, has_vbproject_access, new_excel,
                    open_disk_workbook, remove_module, replace_module_code,
                    import_component,
                    require_open_workbook, vbe_state, vbe_state_hint)
from .repo import (DEFAULT_VERSION_PATTERNS, DOC_MODULE_RE, EXT_FOR_TYPE, TYPE_FOR_EXT,
                   IMPLICIT_PROCS, IMPORTABLE, MODULE_GROUP_ORDER,
                   RULE_ALLOW_ACCEL_FILES, RULE_ALLOW_NEW_FILES, RULE_EXEMPT_MARK,
                   RULE_NEW_MARK, SHELL_TYPE_NAME, TYPE_FOLDER,
                   VBXT_CLASS, VBXT_DOCUMENT, VBXT_FORM, VBXT_STANDARD, VERSION_RE,
                   doc_label_of, doc_stem, files_under, find_source,
                   is_noise_file, is_shell_file, module_index, module_name_of,
                   parse_stem_version, project_dirs, read_common_version,
                   rel_path, resolve_target, shell_text, source_files,
                   source_files_at)
from .ui import (SEP, ask_yes, choose, display_width, grid, group_files,
                 pad_display, pause, pick_grouped, pick_list, pick_list_default,
                 rjust, run_action, run_menu, read_line, section)
from .vba_parse import (EVENT_SUFFIX, MEMBER_END_RE, MEMBER_FIELD_RE,
                        MEMBER_PROC_RE, MEMBER_PROP_RE, VBA_PROC_RE, code_lines,
                        declared_procs, drop_attribute_lines, drop_trailing_placeholder,
                        fold_lines, is_event_proc, mask_strings,
                        norm_code_lines, parse_members, scan_rules, split_code_comment,
                        split_lines, strip_export_header, strip_vba_comments,
                        strip_vba_lines)
from .vbabin import (compare_vba_modules, read_vba_bin, sheet_pages,
                     vba_documents, vba_module_names, vba_models)
# style_rules 是"只动 src 文本"的风格规则层（行尾/BOM、调用括号），入口在 vba-workbench\scripts\agent.py style
from .style_rules import (check_eol, check_option_explicit, has_option_explicit,
                          run_checks, run_fix, scan_bare, scan_noarg, scope_files, selfcheck)
# app_guard 是 Application 状态还原层（也只碰 COM 属性，不 import pywin32）
from .app_guard import (AppGuard, calc_mode, guard, restore)

__all__ = [
    # base
    "ROOT", "CODE_ROOT", "CONFIG_PATH", "APP_CONFIG", "CONFIG_ENV",
    "find_config", "PKG_ROOT", "interactive_ok", "set_non_interactive",
    "make_envelope", "set_tool_name",
    "assume_yes", "set_assume_yes", "json_mode", "set_json_mode",
    "cached_warnings", "reset_warnings",
    "SOURCE_EXTS", "SOURCE_ALL_EXTS", "TEXT_ENCODINGS",
    "SHELL_HINT", "Failure", "msg", "warn", "fail", "abs_path", "read_source_text",
    "entry_command",
    # config
    "load_config", "load_raw_config", "save_config", "config_source",
    # vba_parse
    "strip_vba_lines", "strip_vba_comments", "VBA_PROC_RE", "declared_procs",
    "split_lines", "code_lines", "split_code_comment", "mask_strings",
    "strip_export_header", "parse_members", "is_event_proc", "EVENT_SUFFIX",
    "norm_code_lines", "fold_lines", "drop_attribute_lines", "drop_trailing_placeholder",
    "MEMBER_FIELD_RE", "MEMBER_PROP_RE", "MEMBER_PROC_RE", "MEMBER_END_RE",
    # 代码口径扫描（裸 New / 手写加速 / 缺过程注释），与本包 repo 的 RULE_* 配套
    "scan_rules",
    # check（通用校验：提醒收集 + 模块集合对账，纯逻辑不含 COM）
    "NOTE_LEVELS", "Notes", "ModuleSetDiff", "check_module_set",
    "check_doc_modules", "check_export_set", "check_suffix_modules",
    "check_sheet_modules",
    # vbabin（xlsm 里的 vbaProject.bin：只读模块清单，不要 Excel/密码）
    "read_vba_bin", "vba_models", "vba_module_names", "compare_vba_modules",
    "sheet_pages", "vba_documents",
    # style_rules（只动 src 文本：行尾/BOM、无参空括号、裸参数调用括号）
    "scope_files", "check_eol", "check_option_explicit", "has_option_explicit",
    "scan_noarg", "scan_bare",
    "run_checks", "run_fix", "selfcheck",
    # app_guard（Application 进出还原：事件/屏幕/计算/状态栏/剪切状态）
    "AppGuard", "guard", "restore", "calc_mode",
    # repo
    "MODULE_GROUP_ORDER", "project_dirs", "source_files", "source_files_at",
    "files_under", "rel_path", "module_index",
    "is_noise_file", "is_shell_file", "SHELL_TYPE_NAME", "shell_text", "VERSION_RE",
    "read_common_version", "find_source",
    # 代码口径约定（白名单 / 例外记号 / 隐式入口），检查脚本一律从这里取
    "RULE_ALLOW_NEW_FILES", "RULE_ALLOW_ACCEL_FILES", "RULE_NEW_MARK",
    "RULE_EXEMPT_MARK", "IMPLICIT_PROCS",
    # VBE 模块类型与目录映射（写回 / 抓取 / 核对共用一套编号）
    "VBXT_STANDARD", "VBXT_CLASS", "VBXT_FORM", "VBXT_DOCUMENT", "IMPORTABLE",
    "EXT_FOR_TYPE", "TYPE_FOR_EXT", "TYPE_FOLDER",
    # 文档模块文件名：Sheet4(SKU统计).bas ↔ 模块名 Sheet4
    "DOC_MODULE_RE", "module_name_of", "doc_label_of", "doc_stem",
    # 项目 → xlsm 解析（sync 写回、audit 离线核对共用）
    "resolve_target", "parse_stem_version", "DEFAULT_VERSION_PATTERNS",
    # files
    "write_source", "transcode_exported_file", "gbk_import_copy",
    "write_text_utf8", "dump_json", "stamp_dir", "backup_copy",
    "move_into", "append_line", "remove_tree", "newest_file", "read_tail",
    "latest_backup",
    # watch（纯 stat 的文件监听三件套：快照 / 落点 / 两轮确认）
    "watch_snapshot", "watch_root_of", "watch_changed",
    # watchdog（线程看门狗原语 + PID 存活探测三态）
    "WatchdogTimeout", "run_with_watchdog", "pid_alive", "pid_from_hwnd",
    # dialogs（弹窗/模态识别族：纯函数分类 + 窗口枚举）
    "WindowInfo", "DialogMatch", "detect_dialog", "detect_target_dialog",
    "enumerate_windows", "KIND_COMPILE_ERROR", "KIND_VBE_DIALOG",
    "KIND_MODAL_DIALOG", "KIND_UNKNOWN",
    # excel（COM 边界层；win32com 在函数里延迟导入，顶层只 stdlib）
    "new_excel", "ExcelSession", "attach_running_excel",
    "com_call", "find_excel_with_workbook", "find_open_workbook",
    "require_open_workbook", "open_disk_workbook",
    "vbe_state", "vbe_state_hint", "has_vbproject_access", "ensure_vbe_access",
    "find_component", "component_snapshot", "remove_module", "replace_module_code",
    "import_component",
    "backup_workbook",
    # ui
    "SEP", "display_width", "pad_display", "rjust", "grid", "section",
    "read_line", "pause", "choose", "ask_yes", "run_action",
    "group_files", "pick_list", "pick_list_default", "pick_grouped", "run_menu",
    # cli
    "add_interactive_arg", "add_no_input_arg", "apply_no_input", "ensure_interactive",
    "run_cli", "entry",
]
