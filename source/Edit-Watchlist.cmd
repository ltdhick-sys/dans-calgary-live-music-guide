@echo off
setlocal
title Edit watchlist - Dan's Calgary Live Music Guide
rem Opens the watchlist in Notepad, then redraws the guide (no download) so the "Watching" marks update right away.
if not exist "%~dp0data\watchlist.txt" (
  > "%~dp0data\watchlist.txt" echo # One artist or venue per line. Lines starting with # are ignored.
)
start "" notepad.exe "%~dp0data\watchlist.txt"
echo.
echo   Your watchlist is open in Notepad: one artist or venue per line.
echo   Save it (Ctrl+S) and close Notepad, then press any key here to apply it.
echo.
pause >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Update-CalgaryConcerts.ps1" -RenderOnly -Open
if errorlevel 1 (
  echo.
  echo   The guide could not be redrawn. The message above says why.
  pause
)
