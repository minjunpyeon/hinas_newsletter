@echo off
chcp 949 >nul
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo ============================================
echo   HiNAS Newsletter - Outlook ���� ����
echo ============================================
echo.

rem ===== 1. �ʼ� ���� Ȯ�� =====
if not exist "create_outlook_email.py" (
    echo [����] create_outlook_email.py �� ã�� �� �����ϴ�.
    echo        �� run.bat �� archive_vol1 ���� �ȿ��� �����ؾ� �մϴ�.
    echo.
    pause
    exit /b 1
)
if not exist "newsletter_email.html" (
    echo [����] newsletter_email.html �� ã�� �� �����ϴ�.
    echo.
    pause
    exit /b 1
)

rem ===== 2. �̹���(assets) Ȯ�� - ���� ������ ���� =====
set "ASSETS=%~dp0assets"
if not exist "!ASSETS!" set "ASSETS=%~dp0..\assets"
if not exist "!ASSETS!" (
    echo [���] �̹��� ������ ã�� �� �����ϴ�: !ASSETS!
    echo        �̹����� ������ �� ���� �� �ֽ��ϴ�. ^(���� ������ ��� ����^)
    echo.
) else (
    set "MISSING="
    for %%F in (feat_hd.png avikus_wordmark.png feat_control1.png feat_control2.png) do (
        if not exist "!ASSETS!\%%F" set "MISSING=!MISSING! %%F"
    )
    if defined MISSING (
        echo [���] ������ �̹���:!MISSING!
        echo        �ش� �̹����� ������ �� ���� �� �ֽ��ϴ�.
        echo.
    )
)

rem ===== 3. Python ����� ã�� (py ��ó �켱) =====
set "PYEXE="
where py >nul 2>nul && set "PYEXE=py"
if not defined PYEXE (
    where python >nul 2>nul && set "PYEXE=python"
)

rem ----- Microsoft Store ��¥ python ���� �ɷ����� -----
if defined PYEXE (
    %PYEXE% -c "import sys" >nul 2>nul
    if errorlevel 1 set "PYEXE="
)

if not defined PYEXE (
    echo [����] Python �� ��ġ�Ǿ� ���� �ʽ��ϴ�.
    echo.
    echo  1^) https://www.python.org/downloads/  ���� Python 3 ��ġ
    echo  2^) ��ġ ù ȭ�鿡�� [Add python.exe to PATH] �ݵ�� üũ
    echo  3^) ��ġ �� �� run.bat �� �ٽ� ����
    echo.
    echo  �� "python" �Է� �� Microsoft Store �� �����ٸ�:
    echo     ���� ^> �� ^> ���� �� ���� ^> �� ���� ��Ī ����
    echo     python.exe / python3.exe �׸��� ������.
    echo.
    pause
    exit /b 1
)

echo [Ȯ��] Python �߰�: %PYEXE%
%PYEXE% --version
echo.

rem ===== 4. pywin32 ��ġ Ȯ�� �� ������ �ڵ� ��ġ =====
%PYEXE% -c "import win32com.client" >nul 2>nul
if errorlevel 1 (
    echo [��ġ] �ʿ��� ��Ű�� pywin32 �� ��ġ�մϴ�...
    %PYEXE% -m pip install --upgrade pip
    %PYEXE% -m pip install pywin32
    %PYEXE% -c "import win32com.client" >nul 2>nul
    if errorlevel 1 (
        echo.
        echo [����] pywin32 ��ġ ����. ���ͳ� ������ Ȯ���ϼ���.
        echo        ���� ��ġ: %PYEXE% -m pip install pywin32
        echo.
        pause
        exit /b 1
    )
    echo [Ȯ��] pywin32 ��ġ �Ϸ�.
    echo.
)

rem ===== 5. ���� ���� ���� =====
echo [����] Outlook ���� ���� ��...
echo.
%PYEXE% create_outlook_email.py
set "RC=%errorlevel%"

echo.
if not "%RC%"=="0" (
    echo ============================================
    echo   [����] ���� ���� ���� ^(�ڵ� %RC%^)
    echo   - Outlook ����ũ�� ���� ����/�α��� �������� Ȯ���ϼ���.
    echo ============================================
) else (
    echo ============================================
    echo   �Ϸ�. Outlook �� ���� â�� Ȯ���ϼ���.
    echo   �޴� ��� �Է� �� [������]�� ������ �˴ϴ�.
    echo ============================================
)
echo.
pause
endlocal
