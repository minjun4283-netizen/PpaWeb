@echo off
rem 이 파일은 run_live_server_hidden.vbs 가 내부적으로만 호출합니다 - 직접
rem 더블클릭하지 마세요(창 없이 실행되도록 vbs가 숨김 모드로 띄웁니다).
rem
rem 인자: %1=python 실행 파일  %2=ppa_liveserver.py 경로  %3=xlsm 경로
rem       %4=표준출력/오류를 저장할 로그 파일 경로
rem
rem cmd 안에 따옴표를 겹겹이 넣는 대신, 이미 검증된 배치파일의 "%~1" 인자
rem 전달 방식을 그대로 씁니다 - 경로에 공백이나 & 같은 특수문자가 있어도
rem 안전합니다.
"%~1" "%~2" --xlsm "%~3" > "%~4" 2>&1
