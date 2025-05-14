@echo off
echo 초과근무 분석기를 실행합니다...

REM 가상 환경 활성화
call venv\Scripts\activate.bat

REM 프로그램 실행
python overtime_analyzer.py
