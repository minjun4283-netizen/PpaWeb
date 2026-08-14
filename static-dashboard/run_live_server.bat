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
echo "%XLSM%"| findstr /i "OneDrive" >nul
if not errorlevel 1 (
  echo [참고] OneDrive 동기화 폴더 안의 파일로 보입니다. 완전히 동기화되지
  echo   않았거나 다른 프로그램이 열어둔 상태면 서버가 파일을 찾지 못할 수
  echo   있습니다.
)
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

rem --- pywin32 설치 여부는 ppa_liveserver.py 자신이 시작하자마자 바로
rem     확인해서(별도로 python을 한 번 더 띄워 미리 검사하지 않음 - 그만큼
rem     시작이 빨라집니다) 없으면 친절한 안내와 함께 즉시 종료합니다.

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
