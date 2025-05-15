@echo off
echo 초과근무 분석 도구 초기 설정을 시작합니다...

REM Python이 설치되어 있는지 확인
where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo Python이 설치되어 있지 않습니다. Python을 설치한 후 다시 시도해주세요.
    pause
    exit /b
)

REM 가상 환경 생성
echo 가상 환경을 생성합니다...
python -m venv venv

REM 가상 환경 활성화
echo 가상 환경을 활성화합니다...
call venv\Scripts\activate.bat

REM 필요한 패키지 설치
echo 필요한 패키지를 설치합니다...
python -m pip install --upgrade pip
pip install -r requirements.txt

REM 프로그램 실행
echo.
echo 설정이 완료되었습니다!
echo 프로그램을 지금 실행하시겠습니까?
echo 1. 초과근무 분석기
echo 2. 나중에 실행
set /p run_choice=선택 (1/2): 

if "%run_choice%"=="1" (
    python app.py
) else (
    echo.
    echo 나중에 실행하려면:
    echo - 'run.bat' 파일 실행
)

pause
