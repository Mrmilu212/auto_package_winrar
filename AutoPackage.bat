@echo off
cd /d "%~dp0"

where pythonw >nul 2>&1 && start "" pythonw main.py && exit /b 0
where python >nul 2>&1 && start "" python main.py && exit /b 0

echo 未找到 pythonw 或 python，请将 Python 加入系统 PATH。
echo 安装依赖: python -m pip install -r requirements.txt
pause
exit /b 1
