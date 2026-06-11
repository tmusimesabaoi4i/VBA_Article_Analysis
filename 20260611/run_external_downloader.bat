@echo off
setlocal

chcp 65001 >nul

set "JOB_PATH=%~1"
set "ROOT_DIR=%~2"
set "PROXY_RAW=%~3"
set "WORKERS=%~4"

if "%JOB_PATH%"=="" (
  echo [ERROR] JOB_PATH is empty.
  pause
  exit /b 1
)

if "%ROOT_DIR%"=="" (
  echo [ERROR] ROOT_DIR is empty.
  pause
  exit /b 1
)

if "%WORKERS%"=="" (
  set "WORKERS=3"
)

set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%external_downloader.ps1"

if not exist "%PS1%" (
  echo [ERROR] external_downloader.ps1 was not found.
  echo Expected: "%PS1%"
  pause
  exit /b 1
)

echo ============================================================
echo External Downloader Started
echo ============================================================
echo JOB_PATH  : "%JOB_PATH%"
echo ROOT_DIR  : "%ROOT_DIR%"
echo PROXY_RAW : "%PROXY_RAW%"
echo WORKERS   : %WORKERS%
echo PS1       : "%PS1%"
echo ============================================================

powershell.exe -NoProfile -ExecutionPolicy Bypass ^
  -File "%PS1%" ^
  -Mode Master ^
  -JobPath "%JOB_PATH%" ^
  -RootDir "%ROOT_DIR%" ^
  -ProxyRaw "%PROXY_RAW%" ^
  -MaxWorkers %WORKERS%

echo.
echo ============================================================
echo External Downloader Finished
echo ============================================================
echo Excel側で ImportLatestExternalResults を実行すると結果を取り込めます。
echo.
pause
endlocal