@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo Installing runtime + PyInstaller (if needed)...
python -m pip install -r requirements.txt -r requirements-build.txt
if errorlevel 1 (
    echo pip failed. Ensure Python 3.10+ is installed and on PATH.
    pause
    exit /b 1
)

echo.
echo Building standalone exe (see packaging.spec)...
python -m PyInstaller --noconfirm --clean packaging.spec
if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Done. Executable: dist\简易工程计算器.exe
explorer dist
pause
