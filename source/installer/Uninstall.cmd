<# :
@echo off
rem Removes Dan's Calgary Live Music Guide: weekly task, shortcuts, Settings > Apps entry and this folder.
setlocal
set "UNINST_DIR=%~dp0"
set "UNINST_SELF=%~f0"
rem work from the temp folder, so this window isn't holding the folder it removes
cd /d "%TEMP%"
if not defined SETUP_QUIET (
  choice /C YN /M "Remove Dan's Calgary Live Music Guide, its weekly task and its saved show list"
  if errorlevel 2 exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression (Get-Content -LiteralPath $env:UNINST_SELF -Raw)"
exit /b
#>
$dir = $env:UNINST_DIR.TrimEnd('\')
if (-not (Test-Path (Join-Path $dir '.installed'))) {
  Write-Host "This folder is not an installed copy, so nothing was removed: $dir" -ForegroundColor Yellow
  exit 1
}
$programs = [Environment]::GetFolderPath('Programs'); $desktop = [Environment]::GetFolderPath('Desktop')
# current names, then the names used by the first version
foreach ($task in "Dan's Calgary Live Music Guide update", 'Calgary Concerts Update') {
  try { Unregister-ScheduledTask -TaskName $task -Confirm:$false -ErrorAction Stop; Write-Host 'Removed the weekly task.' } catch { }
}
foreach ($name in "Dan's Calgary Live Music Guide", 'Calgary Live Music Guide') {
  Remove-Item (Join-Path $programs $name) -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item (Join-Path $desktop "$name.lnk") -Force -ErrorAction SilentlyContinue
}
Remove-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\CalgaryLiveMusicGuide' -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Dan's Calgary Live Music Guide was removed." -ForegroundColor Green
if (-not $env:SETUP_QUIET) { [void](Read-Host 'Press Enter to close') }
# This file is still open until the window closes, so the files go a moment later. Only the program's own
# files and folders are deleted, then the folder itself if nothing else is left in it.
$own = 'Update-CalgaryConcerts.ps1', 'template.html', 'Install-WeeklyTask.ps1', 'Run-Update.cmd', 'Edit-Watchlist.cmd', 'README.md',
       'config.json', 'config.previous.json', '.installed', 'Uninstall.cmd'
$steps = @('timeout /t 2 /nobreak >nul')
$steps += 'data', 'output' | ForEach-Object { 'rd /s /q "{0}\{1}"' -f $dir, $_ }
$steps += $own | ForEach-Object { 'del /f /q "{0}\{1}"' -f $dir, $_ }
$steps += 'rd "{0}"' -f $dir
Start-Process cmd.exe -ArgumentList ('/c ' + ($steps -join ' & ')) -WindowStyle Hidden -WorkingDirectory $env:TEMP
