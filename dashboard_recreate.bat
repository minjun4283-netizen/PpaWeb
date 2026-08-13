@echo off
setlocal EnableExtensions EnableDelayedExpansion

cd /d "%~dp0" || goto :FAIL

set "LOG=%~dp0dashboard_recreate.log"
set "JSON=%~1"
if not defined JSON set "JSON=%TEMP%\ppa_export.json"

set "PY=%~dp0..\python-embed\python.exe"
set "SCRIPT=%~dp0build_dashboard.py"
set "OUT=%~dp0PPA현황.html"

> "!LOG!" echo ================================================
>> "!LOG!" echo   PPA 대시보드 생성
>> "!LOG!" echo ================================================
>> "!LOG!" echo.

echo ================================================
echo   PPA 대시보드 생성
echo ================================================
echo.

echo 입력 JSON:
echo(!JSON!
>> "!LOG!" echo JSON=!JSON!
echo.

if not exist "!JSON!" (
    echo [오류] JSON 파일을 찾을 수 없습니다.
    >> "!LOG!" echo [오류] JSON 파일 없음: !JSON!
    goto :FAIL
)

if exist "!PY!" (
    echo 동봉된 python-embed를 사용합니다.
    >> "!LOG!" echo PY=!PY!
) else (
    where python >nul 2>&1
    if errorlevel 1 (
        echo [오류] python을 찾을 수 없습니다.
        >> "!LOG!" echo [오류] python 없음
        goto :FAIL
    )
    set "PY=python"
    echo 시스템 python을 사용합니다.
    >> "!LOG!" echo PY=python
)

if not exist "!SCRIPT!" (
    echo [오류] build_dashboard.py 를 찾을 수 없습니다.
    echo(!SCRIPT!
    >> "!LOG!" echo [오류] SCRIPT 없음: !SCRIPT!
    goto :FAIL
)

>> "!LOG!" echo SCRIPT=!SCRIPT!
>> "!LOG!" echo OUT=!OUT!

echo 대시보드 생성 중...
echo.
>> "!LOG!" echo 대시보드 생성 시작

"!PY!" --version >> "!LOG!" 2>&1
"!PY!" "!SCRIPT!" --json="!JSON!" --out="!OUT!" >> "!LOG!" 2>&1
if errorlevel 1 goto :FAIL

if not exist "!OUT!" (
    echo [오류] HTML 파일이 생성되지 않았습니다.
    >> "!LOG!" echo [오류] HTML 생성 안됨
    goto :FAIL
)

echo ================================================
echo   완료
echo ================================================
start "" "!OUT!"
exit /b 0

:FAIL
echo.
echo ================================================
echo   생성 실패 - 로그 파일을 확인해주세요
echo ================================================
echo 로그 파일:
echo(!LOG!
start "" notepad "!LOG!"
pause
exit /b 1