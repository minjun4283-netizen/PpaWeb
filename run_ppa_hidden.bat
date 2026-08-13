@echo off
setlocal EnableExtensions DisableDelayedExpansion

set "BASE=%~dp0"
if "%BASE:~-1%"=="\" set "BASE=%BASE:~0,-1%"

set "SERVERPY=C:\Users\A1432\AppData\Local\Python\pythoncore-3.14-64\python.exe"
set "BUILDPY=%BASE%\python-embed\python.exe"
set "XLSM=C:\Users\A1432\OneDrive - SKI E&S\(부서문서함) E&S 재생E사업운영팀 - General\DB\Access DB_v0.9.1.xlsm"

cd /d "%BASE%\static-dashboard"

"%SERVERPY%" "%BASE%\static-dashboard\ppa_server.py" --xlsm "%XLSM%" --html "%BASE%\static-dashboard\dashboard.html" --build-script "%BASE%\static-dashboard\build_dashboard.py" --python "%BUILDPY%" --form-js "%BASE%\static-dashboard\dashboard_form.js"

pause