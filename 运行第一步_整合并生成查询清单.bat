@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PYTHONPATH=%~dp0src

echo ============================================================
echo  A-Scan 闭环  第一步：整合并生成查询清单
echo ============================================================
echo.

set "KPIDATE="
set /p KPIDATE=请输入本次 KPI 日期 (YYYY-MM-DD, 必填):
if "%KPIDATE%"=="" (
  echo [中止] 必须显式指定 KPI 日期，程序不会用文件名或 cohort_date 众数代替。
  pause
  exit /b 2
)

echo.
echo 三源合并方式:
echo   1 = 重新运行 Data Combination 三源合并 ^(推荐^)
echo   2 = 使用本次已经生成的 Unified Master 宽表
set "MERGEOPT="
set /p MERGEOPT=请选择 (1/2):
if "%MERGEOPT%"=="2" (set MERGEMODE=reuse) else (set MERGEMODE=run)

set "RECVDATE="
echo.
echo 收货文件内没有日期列时需要显式给出收货日期。
set /p RECVDATE=收货日期 (YYYY-MM-DD, 没有收货文件可直接回车跳过):

set EXTRA=
if not "%RECVDATE%"=="" set EXTRA=--receiving-date %RECVDATE%

echo.
echo 运行中...
python "src\stage1.py" --kpi-date %KPIDATE% --merge %MERGEMODE% %EXTRA%
set RC=%ERRORLEVEL%

echo.
if %RC%==0 (
  echo [完成] 请用 Step1_Output\{run_id}\Status_Check_List.xlsx 去做系统查询
  echo        （只有 Tracking Number 一列）。把下载结果放入 Status_Check\Input，然后运行第二步。
) else (
  echo [失败] 退出码 %RC%。详见 Logs 目录。输入与中间数据已保留。
)
pause
endlocal
