@echo off
setlocal

:: Set base path to this script’s directory
set "BASEDIR=%~dp0"
set "LOGDIR=%BASEDIR%log"

:: Ensure log folder exists
if not exist "%LOGDIR%" (
    mkdir "%LOGDIR%"
)

:: Log that the bat was triggered
echo [%date% %time%] Triggered with slide: %1 >> "%LOGDIR%\trigger_log.txt"

:: Call the PowerShell script with the slide path
powershell -ExecutionPolicy Bypass -File "%BASEDIR%run_inference.ps1" -SlidePath "%~1"

endlocal

echo.
echo Done. Press any key to exit...
pause > nul
