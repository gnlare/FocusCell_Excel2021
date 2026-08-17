@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_release.ps1"
set "EXITCODE=%ERRORLEVEL%"
echo.
if not "%EXITCODE%"=="0" (
    echo Build failed. Check BUILD_LOG.txt.
) else (
    echo Build completed. See INSTALL\KR and INSTALL\EN.
)
pause
exit /b %EXITCODE%
