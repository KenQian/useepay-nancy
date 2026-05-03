@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PYTHON_CMD=python"
set "LOG_FILE=%SCRIPT_DIR%merchant_analyzer_startup.log"

cd /d "%SCRIPT_DIR%"

echo [%DATE% %TIME%] Starting merchant anomaly detector>"%LOG_FILE%"
echo Launcher: %~f0>>"%LOG_FILE%"
echo Working directory: %CD%>>"%LOG_FILE%"
echo PATH: %PATH%>>"%LOG_FILE%"
echo.>>"%LOG_FILE%"
echo Python lookup:>>"%LOG_FILE%"
where %PYTHON_CMD% >>"%LOG_FILE%" 2>&1
echo.>>"%LOG_FILE%"

"%PYTHON_CMD%" --version >>"%LOG_FILE%" 2>&1
if errorlevel 1 (
  echo Python is not installed or is not available in PATH.
  echo.
  echo Please install Python and select "Add python.exe to PATH" during installation.
  echo After installation, double-click this file again.
  echo.
  echo Details were written to:
  echo "%LOG_FILE%"
  pause
  exit /b 1
)

"%PYTHON_CMD%" "%SCRIPT_DIR%merchant_analyzer\merchant_anomaly_detector_windows.py" >>"%LOG_FILE%" 2>&1
if errorlevel 1 (
  echo.
  echo The program did not finish successfully. Check the app message or logs in the result folder.
  echo.
  echo Startup details were written to:
  echo "%LOG_FILE%"
  pause
)

endlocal
