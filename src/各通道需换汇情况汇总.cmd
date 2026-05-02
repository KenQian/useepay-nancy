@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PYTHON_CMD=python"

where %PYTHON_CMD% >nul 2>nul
if errorlevel 1 (
  echo Python 未安装或未加入 PATH。
  echo 请先安装 Python，然后重新双击此文件。
  pause
  exit /b 1
)

%PYTHON_CMD% "%SCRIPT_DIR%fx_summary_workflow\fx_summary_workflow_app.py"
if errorlevel 1 (
  echo.
  echo 程序未成功完成，请查看界面提示或 result 文件夹中的日志。
  pause
)

endlocal
