@echo off
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo ============================================
echo   Avikus Newsletter Issue 01 - Outlook
echo ============================================
echo.

rem ===== 1. required files =====
if not exist "create_outlook_email.py" (
    echo [ERROR] create_outlook_email.py not found.
    echo         run.bat must stay inside the issue01_avikus folder.
    echo.
    pause
    exit /b 1
)
if not exist "newsletter_email.html" (
    echo [ERROR] newsletter_email.html not found.
    echo.
    pause
    exit /b 1
)

rem ===== 2. check assets =====
set "ASSETS=%~dp0assets"
if not exist "!ASSETS!" (
    echo [WARN] assets folder not found: !ASSETS!
    echo        Images may not appear in the mail body.
    echo.
) else (
    set "MISSING="
    for %%F in (logo_avikus_2.png feat_nav2.png feat_svm1.png feat_control1.png feat_control2.png feat_cloud.png) do (
        if not exist "!ASSETS!\%%F" set "MISSING=!MISSING! %%F"
    )
    if defined MISSING (
        echo [WARN] missing images:!MISSING!
        echo        Those images may not appear in the mail body.
        echo.
    )
)

rem ===== 3. find Python (prefer py launcher) =====
set "PYEXE="
where py >nul 2>nul && set "PYEXE=py"
if not defined PYEXE (
    where python >nul 2>nul && set "PYEXE=python"
)

rem ----- guard against Microsoft Store fake python -----
if defined PYEXE (
    %PYEXE% -c "import sys" >nul 2>nul
    if errorlevel 1 set "PYEXE="
)

if not defined PYEXE (
    echo [ERROR] Python is not installed.
    echo.
    echo  1^) Install Python 3 from https://www.python.org/downloads/
    echo  2^) On the first screen, check [Add python.exe to PATH]
    echo  3^) Re-run run.bat after installation
    echo.
    pause
    exit /b 1
)

echo [OK] Python found: %PYEXE%
%PYEXE% --version
echo.

rem ===== 4. ensure pywin32 =====
%PYEXE% -c "import win32com.client" >nul 2>nul
if errorlevel 1 (
    echo [SETUP] installing required package pywin32 ...
    %PYEXE% -m pip install --upgrade pip
    %PYEXE% -m pip install pywin32
    %PYEXE% -c "import win32com.client" >nul 2>nul
    if errorlevel 1 (
        echo.
        echo [ERROR] pywin32 install failed. Check your internet connection.
        echo         Manual: %PYEXE% -m pip install pywin32
        echo.
        pause
        exit /b 1
    )
    echo [OK] pywin32 installed.
    echo.
)

rem ===== 5. run =====
echo [RUN] creating Outlook draft ...
echo.
%PYEXE% create_outlook_email.py
set "RC=%errorlevel%"

echo.
if not "%RC%"=="0" (
    echo ============================================
    echo   [ERROR] failed to create draft ^(code %RC%^)
    echo   - Make sure the Outlook desktop app is running and signed in.
    echo ============================================
) else (
    echo ============================================
    echo   Done. Check the new Outlook draft window.
    echo   Enter recipients, then click Send.
    echo ============================================
)
echo.
pause
endlocal
