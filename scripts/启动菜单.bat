@echo off
rem 双击本文件：工作目录切到包外一层（仓库根），再优先用 py 启动器进菜单
cd /d "%~dp0..\.."

where py >nul 2>nul
if not errorlevel 1 (
    py -3 "%~dp0sync.py" -i
    goto :end
)
where python >nul 2>nul
if not errorlevel 1 (
    python "%~dp0sync.py" -i
    goto :end
)
echo 没找到 Python：请安装 Python 3.10+ 并加入 PATH，或安装 py 启动器。
:end
echo.
pause
