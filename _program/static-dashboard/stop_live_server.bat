@echo off
setlocal enabledelayedexpansion

echo ================================================
echo   PPA 실시간 입력 서버 중지 (수동)
echo ================================================
echo.
echo 참고: 이제 엑셀에서 파일을 닫거나 대시보드 탭을 닫으면 서버가 자동으로
echo 함께 종료됩니다 - 평소에는 이 파일을 실행할 필요가 없습니다.
echo 자동 종료가 안 됐을 때(예: 브라우저를 통째로 강제 종료한 경우 등)만
echo 이 파일로 직접 정리하세요.
echo.

set "PORT=8842"
set "FOUND_PID="

for /f "tokens=5" %%P in ('netstat -ano ^| findstr /r /c:":%PORT% .*LISTENING"') do (
  set "FOUND_PID=%%P"
)

if "%FOUND_PID%"=="" (
  echo %PORT% 포트에서 실행 중인 서버를 찾지 못했습니다 - 이미 꺼져 있는 것
  echo 같습니다.
  goto :END
)

echo PID %FOUND_PID% 로 실행 중인 서버를 찾았습니다.
echo 먼저 정상 종료를 요청합니다(엑셀 COM 연결을 정리하며 꺼짐)...
curl -s -m 3 -X POST http://127.0.0.1:%PORT%/api/shutdown >nul 2>&1

timeout /t 3 /nobreak >nul

netstat -ano | findstr /r /c:":%PORT% .*LISTENING" >nul
if errorlevel 1 (
  echo 서버가 정상 종료됐습니다.
  goto :END
)

echo 정상 종료 요청이 통하지 않아 프로세스를 직접 종료합니다...
taskkill /PID %FOUND_PID% >nul 2>&1
timeout /t 2 /nobreak >nul

netstat -ano | findstr /r /c:":%PORT% .*LISTENING" >nul
if not errorlevel 1 (
  echo 그래도 안 꺼져 강제로 종료합니다...
  taskkill /PID %FOUND_PID% /F >nul 2>&1
)

echo.
echo 서버를 중지했습니다.
echo.
echo 참고: 그래도 엑셀 파일이 "읽기 전용"으로 계속 뜨면, 작업 관리자 -
echo 자세히 탭에서 화면에 안 보이는 EXCEL.EXE 프로세스가 남아있지 않은지
echo 확인해 종료하고, 같은 폴더의 "~$파일명.xlsm" 잠금 파일을 지워보세요
echo (진짜 아무도 안 쓰고 있을 때만 지우세요).

:END
echo.
pause
