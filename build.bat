@echo off
echo ================================================
echo   Visiting Card Scanner v2 — EXE Builder
echo ================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found.
    pause
    exit /b 1
)

echo [1/3] Checking PyInstaller...
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    python -m pip install pyinstaller
)

echo [2/3] Cleaning old build...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

echo [3/3] Building EXE...
python -m PyInstaller build.spec

echo.
if exist "dist\VisitingCardScanner\VisitingCardScanner.exe" (
    echo ================================================
    echo   BUILD SUCCESSFUL
    echo   Location: dist\VisitingCardScanner\
    echo ================================================
) else (
    echo [ERROR] Build failed. Check output above.
)

pause
