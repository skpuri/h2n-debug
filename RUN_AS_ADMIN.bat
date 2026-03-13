@echo off
echo ============================================
echo  Hand2Note 4 - Diagnostic Tool
echo  Please wait, this will take ~45 seconds
echo ============================================
echo.

:: Self-elevate to Administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting Administrator privileges...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: Run the diagnostic script
powershell -ExecutionPolicy Bypass -File "%~dp0H2N_Diagnose.ps1"

echo.
echo Done! Check your Desktop for H2N_DiagnosticReport.html
pause
