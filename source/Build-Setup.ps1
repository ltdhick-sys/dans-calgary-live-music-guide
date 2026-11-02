<#
.SYNOPSIS
  Packs the updater, its settings and the current show list into one file,
  "Dan's Calgary Live Music Guide.cmd", next to this source folder.

.DESCRIPTION
  Copy that one file to any Windows 10/11 PC and double-click it. The first time, it installs to
  %LOCALAPPDATA%\Programs\CalgaryConcerts (no admin rights needed), schedules the weekly refresh,
  adds Start menu and desktop shortcuts and a Settings > Apps entry, and opens the guide.
  After that, double-clicking the same file just opens the guide; a newer build upgrades it.
  Re-run this build after changing the script, template, config or ignore list.

.PARAMETER StripApiKey
  Leave the Ticketmaster key out of the installer, for a copy you give to someone else (or the GitHub release).
  By default the key from config.json or data\ticketmaster.key is included.

.PARAMETER OutDir
  Folder for the .cmd and .zip. Defaults to the folder above this one; use another folder for a shared copy so
  your own keyed installer isn't overwritten.

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File .\Build-Setup.ps1
#>
param(
  [switch]$StripApiKey,
  [string]$OutDir   # where the .cmd and .zip go; default is the folder above this one
)

$ErrorActionPreference = 'Stop'
$root    = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$version = (Get-Date).ToString('yyyy.MM.dd.HHmm', [Globalization.CultureInfo]::InvariantCulture)   # sortable, so a newer build always upgrades an older install
$utf8    = New-Object Text.UTF8Encoding $false
function To-Crlf([string]$s) { return ($s -replace "`r`n", "`n" -replace "`n", "`r`n") }

$stage = Join-Path $env:TEMP ('calgary-build-' + [guid]::NewGuid().ToString('N'))
$files = Join-Path $stage 'files'
New-Item -ItemType Directory -Force (Join-Path $files 'data'), (Join-Path $files 'output') | Out-Null
try {
  # the PowerShell files must be plain ASCII: Windows PowerShell 5.1 reads them in the PC's own code page
  foreach ($f in 'Update-CalgaryConcerts.ps1', 'Install-WeeklyTask.ps1', 'installer\setup-body.ps1') {
    if ([IO.File]::ReadAllText((Join-Path $root $f), [Text.Encoding]::UTF8) -match '[^\x00-\x7F]') { throw "$f must be plain ASCII (Windows PowerShell 5.1 misreads other characters)." }
  }
  $ship = 'Update-CalgaryConcerts.ps1', 'template.html', 'Install-WeeklyTask.ps1', 'Run-Update.cmd', 'Edit-Watchlist.cmd', 'README.md', 'config.json',
          'data\shows.csv', 'data\ignore.txt', 'data\manual.tsv'
  foreach ($f in $ship) { Copy-Item (Join-Path $root $f) (Join-Path $files $f) }
  # made by the first refresh, so they may not exist yet
  foreach ($f in 'data\venues.csv', 'data\runs.csv', 'data\source-runs.csv') { if (Test-Path (Join-Path $root $f)) { Copy-Item (Join-Path $root $f) (Join-Path $files $f) } }
  Copy-Item (Join-Path $root 'installer\Uninstall.cmd') (Join-Path $files 'Uninstall.cmd')
  # a blank starter watchlist, never yours (data\config.local.json, your own settings, isn't shipped either)
  Copy-Item (Join-Path $root 'installer\watchlist.default.txt') (Join-Path $files 'data\watchlist.txt')
  # draw the shipped page fresh from the staged copy, so it has the current template and no personal watchlist
  $render = Start-Process powershell.exe -Wait -PassThru -WindowStyle Hidden -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
              ('"{0}"' -f (Join-Path $files 'Update-CalgaryConcerts.ps1')), '-RenderOnly')
  if ($render.ExitCode -ne 0 -or -not (Test-Path (Join-Path $files 'output\calgary-live-music.html'))) { throw 'Could not draw the guide page for the installer.' }

  # batch files need Windows line endings
  foreach ($f in 'Run-Update.cmd', 'Edit-Watchlist.cmd', 'Uninstall.cmd') {
    $p = Join-Path $files $f
    [IO.File]::WriteAllText($p, (To-Crlf ([IO.File]::ReadAllText($p))), [Text.Encoding]::ASCII)
  }
  # The key comes from config.json or data\ticketmaster.key (kept out of the GitHub copy). It goes into the
  # installer's config.json unless -StripApiKey.
  $p = Join-Path $files 'config.json'
  $tmKey = [string](@(([IO.File]::ReadAllText($p, [Text.Encoding]::UTF8) | ConvertFrom-Json).sources | Where-Object { $_.type -eq 'ticketmaster' })[0].api_key)
  $keyFile = Join-Path $root 'data\ticketmaster.key'
  if (-not $tmKey -and (Test-Path $keyFile)) { $tmKey = ([IO.File]::ReadAllText($keyFile)).Trim() }
  $shipKey = if ($StripApiKey) { '' } else { $tmKey }
  $cfg = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8) -replace '("type":\s*"ticketmaster",\s*"api_key":\s*)"[^"]*"', ('${1}"' + $shipKey + '"')
  [IO.File]::WriteAllText($p, $cfg, $utf8)

  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip = Join-Path $stage 'payload.zip'
  [IO.Compression.ZipFile]::CreateFromDirectory($files, $zip, [IO.Compression.CompressionLevel]::Optimal, $false)
  $b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($zip))
} finally {
  Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
}
$lines = for ($i = 0; $i -lt $b64.Length; $i += 120) { $b64.Substring($i, [math]::Min(120, $b64.Length - $i)) }

# cmd runs the batch lines up to "exit /b"; PowerShell sees them as a <# #> comment and runs the rest.
# Get-Content + Invoke-Expression (rather than .NET calls) still works on PCs where IT restricts
# PowerShell, so the body can explain the problem instead of crashing.
$header = @'
<# :
@echo off
setlocal
title Dan's Calgary Live Music Guide
set "SETUP_SELF=%~f0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression (Get-Content -LiteralPath $env:SETUP_SELF -Raw)"
set "rc=%errorlevel%"
if %rc% neq 0 if not defined SETUP_QUIET (
  echo.
  echo   Setup did not finish. The message above says why.
  pause
)
exit /b %rc%
#>
'@
$body = [IO.File]::ReadAllText((Join-Path $root 'installer\setup-body.ps1'), [Text.Encoding]::UTF8)
if ($body -match '[^\x00-\x7F]') { throw 'installer\setup-body.ps1 must be plain ASCII because it runs inside a .cmd file.' }

$text = $header + "`n" +
        "`$Version = '$version'`n" +
        "`$Payload = @'`n" + ($lines -join "`n") + "`n'@`n" +
        $body
# the one file to copy to another PC sits next to the source folder
if (-not $OutDir) { $OutDir = Split-Path $root -Parent }
New-Item -ItemType Directory -Force $OutDir | Out-Null
$out = Join-Path $OutDir "Dan's Calgary Live Music Guide.cmd"
[IO.File]::WriteAllText($out, (To-Crlf $text), [Text.Encoding]::ASCII)

# the same file zipped with short instructions, for sending by email
$howTo = @'
Dan's Calgary Live Music Guide - how to install

1. Right-click this zip file and choose "Extract All...", then click Extract.
2. In the folder that opens, double-click "Dan's Calgary Live Music Guide.cmd".
   If Windows says "Windows protected your PC", click "More info", then "Run anyway".
3. After a few seconds the guide opens in its own window. There are no questions.

To open it again later, use the "Dan's Calgary Live Music Guide" shortcut on the desktop
or in the Start menu, or double-click the same .cmd file again.

If something goes wrong, the window stays open and says why. A log is saved to
%TEMP%\DansCalgaryLiveMusicGuide-setup.log
'@
$zipOut = [IO.Path]::ChangeExtension($out, '.zip')
$zipStage = Join-Path $env:TEMP ('calgary-zip-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $zipStage | Out-Null
try {
  Copy-Item -LiteralPath $out -Destination $zipStage
  [IO.File]::WriteAllText((Join-Path $zipStage 'How to install.txt'), (To-Crlf $howTo), [Text.Encoding]::ASCII)
  if (Test-Path -LiteralPath $zipOut) { Remove-Item -LiteralPath $zipOut -Force }
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  [IO.Compression.ZipFile]::CreateFromDirectory($zipStage, $zipOut, [IO.Compression.CompressionLevel]::Optimal, $false)
} finally {
  Remove-Item $zipStage -Recurse -Force -ErrorAction SilentlyContinue
}

$shows = @(Import-Csv (Join-Path $root 'data\shows.csv') -Encoding UTF8).Count
$keyNote = if (-not $tmKey) { 'no Ticketmaster key in config.json yet' } elseif ($StripApiKey) { 'Ticketmaster key left out' } else { 'Ticketmaster key included' }
Write-Host ("Built {0}  ({1:N0} KB, version {2}, {3} shows seeded, {4})" -f $out, ((Get-Item -LiteralPath $out).Length / 1KB), $version, $shows, $keyNote)
Write-Host ("Zipped for email: {0}  ({1:N0} KB)" -f $zipOut, ((Get-Item -LiteralPath $zipOut).Length / 1KB))
