@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  py -3 -m venv .venv
  if errorlevel 1 goto failed
)
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -m accounting_report ui
if errorlevel 1 goto failed
exit /b 0
:failed
echo.
echo 界面启动失败，请检查 Python 是否安装，并查看上方提示。
pause
exit /b 2
