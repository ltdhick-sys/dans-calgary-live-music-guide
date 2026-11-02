@echo off
rem Double-click to refresh the Calgary concert list and open the result.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Update-CalgaryConcerts.ps1" -Open
echo.
pause
