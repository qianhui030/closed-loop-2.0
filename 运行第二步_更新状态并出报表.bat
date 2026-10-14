@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PYTHONPATH=%~dp0src

echo ============================================================
echo  A-Scan 闭环  第二步：更新状态并出报表
echo ============================================================
echo.
echo 请确认已把系统查询结果放入 Status_Check\Input
echo.
echo 1 = 仅检查字段和覆盖率（不更新、不归档、不改 Pool）
echo 2 = 正式更新（需收货日期已确认，且字段核对通过）
set "MODE="
set /p MODE=请选择 (1/2):

set "RUNID="
set /p RUNID=第一步批次 run_id (只有一个待处理批次时可直接回车):

set EXTRA=
if not "%RUNID%"=="" set EXTRA=--run-id %RUNID%
if "%MODE%"=="1" set EXTRA=%EXTRA% --inspect-only

echo.
if "%MODE%"=="1" (
  echo 运行覆盖率检查...
) else (
  echo 运行正式更新... 生成原生 Excel Pivot 期间请不要操作 Excel。
)
python "src\stage2.py" %EXTRA%
set RC=%ERRORLEVEL%

echo.
if %RC%==0 (
  if "%MODE%"=="1" (
    echo [完成] 仅检查。详见 Working\{run_id}\status_check_inspect.json
  ) else (
    echo [完成] 报表在 Final_Output\{kpi_date}\{run_id}\Ascan_KPI.xlsx
  )
) else (
  echo [失败] 退出码 %RC%。详见 Logs 目录。Pool 与输入均未改动，可修复后重跑。
)
pause
endlocal
