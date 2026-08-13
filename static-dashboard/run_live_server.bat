@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ================================================
echo   PPA 실시간 입력 서버 시작
echo ================================================
echo.

rem --- 1. 대상 xlsm 경로 찾기 -----------------------------------------
rem     인자로 넘어오면 그걸 쓰고(엑셀 매크로가 넘겨주는 경우),
rem     없으면 이 파일 바로 위(작업 폴더)에서 xlsm 파일을 자동으로 찾습니다.
set "XLSM=%~1"
if not "%XLSM%"=="" goto :GOTXLSM

set "FOUND="
set "XCOUNT=0"
for %%F in ("%~dp0..\*.xlsm") do (
  set "FOUND=%%~fF"
  set /a XCOUNT+=1
)
if "%XCOUNT%"=="0" (
  echo [오류] 상위 폴더에서 xlsm 파일을 찾지 못했습니다.
  echo   - 이 배치파일이 static-dashboard 폴더 안에 있고, 그 바로 위 폴더에
  echo     PPA 엑셀 파일이 있는지 확인해주세요.
  echo   - 또는 직접 경로를 지정하세요: run_live_server.bat "엑셀파일경로.xlsm"
  goto :FAIL
)
if not "%XCOUNT%"=="1" (
  echo [오류] 상위 폴더에 xlsm 파일이 여러 개 있어 자동으로 고를 수 없습니다.
  echo   직접 경로를 지정해주세요: run_live_server.bat "엑셀파일경로.xlsm"
  goto :FAIL
)
set "XLSM=%FOUND%"

:GOTXLSM
echo 대상 엑셀 파일:
echo   "%XLSM%"
echo.

rem --- 2. python 위치 - PATH 우선, 없으면 동봉된 python-embed ----------
set "PY=python"
where python >nul 2>&1
if errorlevel 1 (
  if exist "%~dp0..\python-embed\python.exe" (
    set "PY=%~dp0..\python-embed\python.exe"
    echo PATH에 python이 없어 동봉된 python-embed를 사용합니다.
  ) else (
    echo [오류] python을 찾을 수 없습니다.
    echo   PATH에 python이 있거나, 이 폴더 바로 위에 python-embed 폴더가
    echo   있어야 합니다. static-dashboard\README.md 의 Python 설치 안내를
    echo   참고해주세요.
    goto :FAIL
  )
)

rem --- 3. pywin32 설치 여부 확인 - 없으면 여기서 바로 안내하고 종료 -----
"%PY%" -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
  echo [오류] pywin32 가 설치되어 있지 않습니다.
  echo   이 기능(웹 화면 입력을 실제 엑셀에 반영)은 pywin32(Windows COM
  echo   자동화)가 반드시 있어야 동작합니다. static-dashboard\README.md 의
  echo   'pywin32 설치' 절차를 먼저 진행해주세요.
  goto :FAIL
)

echo 서버를 시작합니다. 이 창을 닫으면 서버도 함께 종료됩니다.
echo 브라우저가 몇 초 안에 자동으로 열립니다(안 열리면 아래 주소를 직접
echo 열어주세요: http://127.0.0.1:8842).
echo.

"%PY%" ppa_liveserver.py --xlsm "%XLSM%"
goto :END

:FAIL
echo.
echo ================================================
echo   시작 실패 - 위 메시지를 확인해주세요
echo ================================================
pause
exit /b 1

:END
echo.
pause
exit /b 0
