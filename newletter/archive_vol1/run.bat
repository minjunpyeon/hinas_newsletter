@echo off
chcp 949 >nul
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo ============================================
echo   HiNAS Newsletter - Outlook 메일 생성
echo ============================================
echo.

rem ===== 1. 필수 파일 확인 =====
if not exist "create_outlook_email.py" (
    echo [오류] create_outlook_email.py 를 찾을 수 없습니다.
    echo        이 run.bat 은 archive_vol1 폴더 안에서 실행해야 합니다.
    echo.
    pause
    exit /b 1
)
if not exist "newsletter_email.html" (
    echo [오류] newsletter_email.html 을 찾을 수 없습니다.
    echo.
    pause
    exit /b 1
)

rem ===== 2. 이미지(assets) 확인 - 상위 폴더에 있음 =====
set "ASSETS=%~dp0..\assets"
if not exist "!ASSETS!" (
    echo [경고] 이미지 폴더를 찾을 수 없습니다: !ASSETS!
    echo        이미지가 본문에 안 보일 수 있습니다. ^(메일 생성은 계속 진행^)
    echo.
) else (
    set "MISSING="
    for %%F in (feat_hd.png avikus_wordmark.png feat_control1.png feat_control2.png) do (
        if not exist "!ASSETS!\%%F" set "MISSING=!MISSING! %%F"
    )
    if defined MISSING (
        echo [경고] 누락된 이미지:!MISSING!
        echo        해당 이미지는 본문에 안 보일 수 있습니다.
        echo.
    )
)

rem ===== 3. Python 실행기 찾기 (py 런처 우선) =====
set "PYEXE="
where py >nul 2>nul && set "PYEXE=py"
if not defined PYEXE (
    where python >nul 2>nul && set "PYEXE=python"
)

rem ----- Microsoft Store 가짜 python 스텁 걸러내기 -----
if defined PYEXE (
    %PYEXE% -c "import sys" >nul 2>nul
    if errorlevel 1 set "PYEXE="
)

if not defined PYEXE (
    echo [오류] Python 이 설치되어 있지 않습니다.
    echo.
    echo  1^) https://www.python.org/downloads/  에서 Python 3 설치
    echo  2^) 설치 첫 화면에서 [Add python.exe to PATH] 반드시 체크
    echo  3^) 설치 후 이 run.bat 을 다시 실행
    echo.
    echo  ※ "python" 입력 시 Microsoft Store 가 열린다면:
    echo     설정 ^> 앱 ^> 고급 앱 설정 ^> 앱 실행 별칭 에서
    echo     python.exe / python3.exe 항목을 끄세요.
    echo.
    pause
    exit /b 1
)

echo [확인] Python 발견: %PYEXE%
%PYEXE% --version
echo.

rem ===== 4. pywin32 설치 확인 후 없으면 자동 설치 =====
%PYEXE% -c "import win32com.client" >nul 2>nul
if errorlevel 1 (
    echo [설치] 필요한 패키지 pywin32 를 설치합니다...
    %PYEXE% -m pip install --upgrade pip
    %PYEXE% -m pip install pywin32
    %PYEXE% -c "import win32com.client" >nul 2>nul
    if errorlevel 1 (
        echo.
        echo [오류] pywin32 설치 실패. 인터넷 연결을 확인하세요.
        echo        수동 설치: %PYEXE% -m pip install pywin32
        echo.
        pause
        exit /b 1
    )
    echo [확인] pywin32 설치 완료.
    echo.
)

rem ===== 5. 메일 생성 실행 =====
echo [실행] Outlook 메일 생성 중...
echo.
%PYEXE% create_outlook_email.py
set "RC=%errorlevel%"

echo.
if not "%RC%"=="0" (
    echo ============================================
    echo   [오류] 메일 생성 실패 ^(코드 %RC%^)
    echo   - Outlook 데스크톱 앱이 실행/로그인 상태인지 확인하세요.
    echo ============================================
) else (
    echo ============================================
    echo   완료. Outlook 새 메일 창을 확인하세요.
    echo   받는 사람 입력 후 [보내기]를 누르면 됩니다.
    echo ============================================
)
echo.
pause
endlocal
