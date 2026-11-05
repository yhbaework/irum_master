@echo off
chcp 65001 >nul
setlocal

rem 대학공시 마스터 작업용 가상환경(irum_master) 생성 스크립트
rem 이 파일이 있는 폴더(작업환경 최상위) 기준으로 동작한다.
cd /d "%~dp0"

set "VENV_DIR=irum_master"
set "PY_VER=3.14"

echo [INFO] Python %PY_VER% 확인 중...
py -%PY_VER% --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python %PY_VER% 를 찾을 수 없습니다. python.org 에서 설치 후 다시 실행하세요.
    echo         설치된 버전 목록:
    py -0p
    goto :fail
)
py -%PY_VER% --version

if exist "%VENV_DIR%\Scripts\python.exe" (
    echo [INFO] 기존 가상환경 %VENV_DIR% 을 재사용합니다.
) else (
    echo [INFO] 가상환경 %VENV_DIR% 생성 중...
    py -%PY_VER% -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo [ERROR] 가상환경 생성 실패
        goto :fail
    )
)

set "VENV_PY=%VENV_DIR%\Scripts\python.exe"

echo [INFO] pip 업그레이드 중...
"%VENV_PY%" -m pip install --upgrade pip

rem 외부망이 차단된 PC 대비: 폴더에 .whl 이 있으면 오프라인 설치한다.
dir /b *.whl >nul 2>&1
if not errorlevel 1 (
    echo [INFO] 로컬 .whl 발견 - 오프라인 설치를 시도합니다.
    "%VENV_PY%" -m pip install --no-index --find-links=. -r requirements.txt
) else (
    echo [INFO] requirements.txt 설치 중...
    "%VENV_PY%" -m pip install -r requirements.txt
)
if errorlevel 1 (
    echo [ERROR] 패키지 설치 실패
    echo         외부망이 차단된 PC라면 인터넷이 되는 PC에서 아래 명령으로 받은 .whl 을
    echo         이 폴더에 넣고 다시 실행하세요.
    echo         pip download -r requirements.txt -d .
    goto :fail
)

echo.
echo [INFO] 설치 완료. 설치된 패키지:
"%VENV_PY%" -m pip list

echo.
echo [INFO] 실행 방법 (연도 폴더 안에서):
echo        cd 2027
echo        ..\%VENV_DIR%\Scripts\python.exe "DB build 마스터.py"
goto :end

:fail
echo.
echo [ERROR] setup 중단
pause
exit /b 1

:end
pause
exit /b 0
