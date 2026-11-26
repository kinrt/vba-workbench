# -*- coding: utf-8 -*-
"""生成空白测试表：一张表页 + 一个标准模块，计算模式手动。

    在工具包根目录运行 `python scripts/make_lab_workbook.py 输出/示例表 V1.xlsm`

只用来做回归验证（pull / push / run 的真实往返），不碰任何正式工作簿。
需要 Windows + 微软 Excel + pywin32；跑完 Excel 会自己退出。
"""

import pathlib
import sys

from toolkit import excel

CODE = ('Option Explicit\n\n'
        "' 测试用最小过程：两数相加\n"
        'Public Function 示例求和(a As Long, b As Long) As Long\n'
        '    示例求和 = a + b\n'
        'End Function\n')


def main() -> int:
    if len(sys.argv) < 2:
        print("用法：python make_lab_workbook.py <输出.xlsm>")
        return 2
    target = pathlib.Path(sys.argv[1]).resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    # 走统一 gateway：登记归属快照，收尾用 quit_owned；不再裸 Quit。
    app = excel.new_excel(hidden=True, force_new=True)
    problems: list[str] = []
    try:
        wb = app.Workbooks.Add()
        wb.Worksheets(1).Name = "表页1"
        comp = wb.VBProject.VBComponents.Add(1)   # 标准模块
        comp.Name = "示例模块"
        comp.CodeModule.AddFromString(CODE)
        app.Calculation = -4135                    # xlManual
        wb.SaveAs(str(target), FileFormat=52)           # xlOpenXMLWorkbookMacroEnabled
        wb.Close(SaveChanges=False)
    finally:
        problems = excel.quit_owned(app)
    if problems:
        print("实例收尾未完成：" + "；".join(problems), file=sys.stderr)
        return 1
    print(f"已生成：{target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
