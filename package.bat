@echo off
setlocal enabledelayedexpansion

:: STEP 1 - Run helper script (it creates version.ini)
powershell -ExecutionPolicy Bypass -File .\package_version_helper.ps1

:: STEP 2 - Build executable
python -m PyInstaller --onefile --name EmailListFusion --icon=fusion.ico app.py --version-file version.ini

echo Packaging done
echo.

:: STEP 3 - Extract full version string (e.g., 1.2.2.0)
for /f "delims=" %%V in ('powershell -NoProfile -Command ^
    "(Get-Content version.ini | Select-String 'FileVersion').ToString().Split(\"'\", [System.StringSplitOptions]::RemoveEmptyEntries)[3]"') do (
    set "fullver=%%V"
)

:: STEP 4 - Trim to major.minor.patch
if defined fullver (
    for /f "tokens=1-3 delims=." %%A in ("!fullver!") do (
        set "version=%%A.%%B.%%C"
    )
)

:: STEP 5 - Final check and copy
if not defined version (
    echo [ERROR] Version not found.
    exit /b 1
)

echo [DEBUG] fullver: !fullver!
echo [DEBUG] version: !version!
copy /Y "dist\EmailListFusion.exe" "EmailListFusion_v!version!.exe"

echo Done.
pause
