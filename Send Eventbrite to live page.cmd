@echo off
title Send Eventbrite to the live page
rem Reads Eventbrite's Calgary listings from this PC, sends them to the GitHub project and refreshes the live page.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0source\Send-EventbriteFeed.ps1"
echo.
pause
