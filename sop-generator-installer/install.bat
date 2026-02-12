@echo off
setlocal enabledelayedexpansion

echo ============================================
echo   SOP Generator - Camunda Modeler Plugin
echo   Installer
echo ============================================
echo.

REM --- Step 1: Check Python is available ---
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    echo.
    echo Please install Python 3.8+ from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo [OK] %%v found.

REM --- Step 2: Check pip is available ---
python -m pip --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] pip is not available. Please reinstall Python with pip enabled.
    pause
    exit /b 1
)
echo [OK] pip found.
echo.

REM --- Step 3: Create backend directory ---
set "BACKEND_DIR=%LOCALAPPDATA%\SOP_Generator\backend"
echo Installing backend to: %BACKEND_DIR%

if not exist "%BACKEND_DIR%\templates" mkdir "%BACKEND_DIR%\templates"

REM --- Step 4: Copy backend files ---
echo Copying backend files...
xcopy /Y /Q "%~dp0backend\*.*" "%BACKEND_DIR%\" >nul
xcopy /Y /Q "%~dp0backend\templates\*.*" "%BACKEND_DIR%\templates\" >nul
if errorlevel 1 (
    echo [ERROR] Failed to copy backend files.
    pause
    exit /b 1
)
echo [OK] Backend files installed.

REM --- Step 5: Install Python dependencies ---
echo Installing Python dependencies (this may take a minute)...
python -m pip install -r "%BACKEND_DIR%\requirements.txt" --quiet 2>nul
if errorlevel 1 (
    echo [WARNING] Some dependencies may have failed to install.
    echo Try running manually: python -m pip install -r "%BACKEND_DIR%\requirements.txt"
) else (
    echo [OK] Python dependencies installed.
)

REM --- Step 6: Verify Word template exists ---
if not exist "%BACKEND_DIR%\final_master_template_2.docx" (
    echo Creating Word template...
    pushd "%BACKEND_DIR%"
    python create_template.py >nul 2>&1
    popd
    if exist "%BACKEND_DIR%\final_master_template_2.docx" (
        echo [OK] Word template created.
    ) else (
        echo [WARNING] Could not create Word template. SOP generation may fail.
    )
) else (
    echo [OK] Word template present.
)

REM --- Step 7: Install Camunda Modeler plugin ---
set "PLUGIN_DIR=%APPDATA%\camunda-modeler\resources\plugins\sop-generator"
echo.
echo Installing plugin to: %PLUGIN_DIR%

if not exist "%APPDATA%\camunda-modeler\resources\plugins" (
    echo [NOTE] Camunda Modeler plugins directory not found.
    echo        Creating it now - make sure Camunda Modeler is installed.
    mkdir "%APPDATA%\camunda-modeler\resources\plugins" >nul 2>&1
)
if not exist "%PLUGIN_DIR%" mkdir "%PLUGIN_DIR%"

xcopy /Y /Q "%~dp0plugin\*.*" "%PLUGIN_DIR%\" >nul
if errorlevel 1 (
    echo [ERROR] Failed to copy plugin files.
    pause
    exit /b 1
)
echo [OK] Plugin installed.

REM --- Step 8: Done ---
echo.
echo ============================================
echo   Installation Complete!
echo ============================================
echo.
echo   Backend installed to:
echo     %BACKEND_DIR%
echo.
echo   Plugin installed to:
echo     %PLUGIN_DIR%
echo.
echo   Next steps:
echo     1. Close and reopen Camunda Modeler
echo     2. Open a BPMN diagram
echo     3. Press Ctrl+Shift+G (or Plugins menu ^> Generate SOP Document)
echo.
pause
