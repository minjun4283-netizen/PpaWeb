@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ================================================
echo   PPA 대시보드 생성
echo ================================================
echo.

rem --- 1. 대상 xlsm 경로 결정 -----------------------------------------
rem     인자로 넘어오면 그걸 쓰고(엑셀 매크로가 넘겨주는 경우),
rem     없으면 이 폴더 바로 위(작업 폴더)에서 xlsm 파일을 자동으로 찾습니다.
set "XLSM=%~1"
if not "%XLSM%"=="" goto :GOTXLSM

set "FOUND="
set "XCOUNT=0"
for %%F in ("%~dp0..\..\*.xlsm") do (
  set "FOUND=%%~fF"
  set /a XCOUNT+=1
)
if "%XCOUNT%"=="0" (
  echo [오류] 상위 폴더에서 xlsm 파일을 찾지 못했습니다.
  echo   - 이 배치파일이 (작업폴더)\_program\static-dashboard\ 안에 있고,
  echo     그 작업폴더에 PPA 엑셀 파일이 있는 구조인지 확인해주세요.
  echo   - 또는 직접 경로를 지정: dashboard_recreate.bat "엑셀파일경로.xlsm"
  goto :FAIL
)
if not "%XCOUNT%"=="1" (
  echo [오류] 상위 폴더에 xlsm 파일이 여러 개 있어 자동으로 고를 수 없습니다.
  echo   직접 경로를 지정해주세요: dashboard_recreate.bat "엑셀파일경로.xlsm"
  goto :FAIL
)
set "XLSM=%FOUND%"

:GOTXLSM
echo 대상 엑셀 파일:
echo   "%XLSM%"
echo "%XLSM%"| findstr /i "OneDrive" >nul
if not errorlevel 1 (
  echo [참고] OneDrive 동기화 폴더 안의 파일로 보입니다. 완전히 동기화되지
  echo   않았거나 다른 프로그램이 열어둔 상태면 오류가 날 수 있습니다.
)
echo.

rem --- 2. python 결정 - PATH 우선, 없으면 동봉된 python-embed ----------
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

rem --- 3. 항상 같은 출력 파일명 -----------------------------------------
rem     매번 파일명이 같아야 비교 대상 스냅샷도 매번 같은 파일을 가리켜서
rem     지난번 실행 대비 변경사항이 계속 누적 비교됩니다. 절대 바꾸지 마세요.
rem     최상위(작업폴더)에 생성되도록 이 폴더(_program\static-dashboard)
rem     기준으로 두 단계 위 경로를 씁니다.
set "OUT=..\..\PPA현황.html"
set "SNAP=..\..\PPA현황_snapshot.json"

if exist "%SNAP%" (
  echo 비교 대상 스냅샷: %SNAP% - 있음, 지난번 대비 변경사항이 표시됩니다.
) else (
  echo 비교 대상 스냅샷: %SNAP% - 없음, 이번이 첫 생성이거나 스냅샷이 사라졌습니다.
)
echo.
echo %OUT% 생성 중...
echo.

"%PY%" build_dashboard.py --xlsm="%XLSM%" --out="%OUT%"
if errorlevel 1 goto :FAIL

echo.
echo ================================================
echo   완료: "%CD%\..\..\PPA현황.html"
echo ================================================
start "" "%OUT%"
goto :END

:FAIL
echo.
echo ================================================
echo   생성 실패 - 위 메시지를 확인해주세요
echo ================================================
pause
exit /b 1

:END
echo.
pause
exit /b 0
