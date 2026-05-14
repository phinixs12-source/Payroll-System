@echo off
chcp 65001 > nul
title 인건비 관리 시스템

echo.
echo  ========================================
echo   인건비 관리 시스템 시작 중...
echo  ========================================
echo.

REM Python 설치 확인
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo  [오류] Python이 설치되어 있지 않습니다.
    echo  https://www.python.org 에서 Python을 설치해주세요.
    pause
    exit /b
)

REM 패키지 설치 확인
python -c "import flask" > nul 2>&1
if %errorlevel% neq 0 (
    echo  필요한 패키지를 설치합니다...
    pip install -r requirements.txt
    echo.
)

echo  브라우저에서 http://localhost:5000 으로 접속하세요.
echo  종료하려면 이 창을 닫으세요.
echo.

REM 브라우저 자동 열기 (3초 후)
start /b cmd /c "timeout /t 2 > nul && start http://localhost:5000"

python app.py

pause
