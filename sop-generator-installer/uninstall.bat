@echo off
setlocal

echo ============================================
echo   SOP Generator - Uninstaller
echo ============================================
echo.
echo This will remove:
echo   - Backend:  %LOCALAPPDATA%\SOP_Generator\backend\
echo   - Plugin:   %APPDATA%\camunda-modeler\resources\plugins\sop-generator\
echo.
echo Generated documents and history data will NOT be deleted.
echo.

set /p CONFIRM=Are you sure you want to uninstall? (Y/N):
if /i not "%CONFIRM%"=="Y" (
    echo.
    echo Uninstall cancelled.
    pause
    exit /b 0
)

echo.

REM --- Remove plugin ---
if exist "%APPDATA%\camunda-modeler\resources\plugins\sop-generator" (
    rmdir /S /Q "%APPDATA%\camunda-modeler\resources\plugins\sop-generator"
    echo [OK] Plugin removed.
) else (
    echo [SKIP] Plugin directory not found.
)

REM --- Remove backend ---
if exist "%LOCALAPPDATA%\SOP_Generator\backend" (
    rmdir /S /Q "%LOCALAPPDATA%\SOP_Generator\backend"
    echo [OK] Backend removed.
) else (
    echo [SKIP] Backend directory not found.
)

echo.
echo ============================================
echo   Uninstall Complete
echo ============================================
echo.
echo Python packages (Flask, python-docx, etc.) were NOT removed.
echo To remove them manually:
echo   pip uninstall flask python-docx lxml waitress docxtpl
echo.
echo Restart Camunda Modeler to complete the uninstall.
echo.
pause
