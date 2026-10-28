
@echo off
setlocal
cd /d "%~dp0"

set "SCRIPT=MERGE_TOOL.py"
set "APP_NAME=Merge Jupiter Tool"

echo ========================================
echo   Building %APP_NAME%
echo ========================================
echo.

if not exist "%SCRIPT%" (
    echo ERROR: %SCRIPT% not found in %~dp0
    pause
    exit /b 1
)

REM Only install packages when imports are missing.
python -c "import fastexcel, polars, xlsxwriter, tkinterdnd2, PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo Installing missing Python packages...
    python -m pip install fastexcel polars xlsxwriter tkinterdnd2 pyinstaller
    if errorlevel 1 (
        echo ERROR: pip install failed.
        pause
        exit /b 1
    )
)

REM Prefer ICO when both icons exist.
set "ICON_ARGS="
if exist "icon.png" set ICON_ARGS=--add-data "icon.png;."
if exist "icon.ico" set "ICON_ARGS=--icon icon.ico --add-data icon.ico;."
if exist "icon.png" if not exist "icon.ico" set "ICON_ARGS=--add-data icon.png;."

echo Starting PyInstaller...
echo.

python -m PyInstaller --noconfirm --windowed --onedir ^
  --distpath "." ^
  --name "%APP_NAME%" ^
  %ICON_ARGS% ^
  --collect-all fastexcel ^
  --collect-all polars ^
  "%SCRIPT%"

if errorlevel 1 (
    echo.
    echo ERROR: Build failed. See messages above.
    pause
    exit /b 1
)

REM Remove temporary build files.
if exist build rmdir /s /q build
if exist "%APP_NAME%.spec" del "%APP_NAME%.spec"

echo.
echo Build completed.
echo Output: "%~dp0%APP_NAME%\%APP_NAME%.exe"
pause
