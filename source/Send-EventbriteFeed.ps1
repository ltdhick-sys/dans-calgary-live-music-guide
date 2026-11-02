<#
.SYNOPSIS
  Sends this PC's Eventbrite listings to the live web page on GitHub and refreshes the page.

.DESCRIPTION
  Eventbrite turns away GitHub's servers, so the live page can't read it itself. This reads Eventbrite's Calgary
  concert listings from this PC (the same way the guide always does), uploads them to the GitHub project as
  feeds/eventbrite.csv, and starts a refresh of the live page. The weekly refresh keeps using that list for up to
  8 days, so running this once a week (before Monday morning) keeps Eventbrite shows on the page.

  Needs the GitHub command-line tool (gh), signed in to the account that owns the project.

.PARAMETER Repo
  owner/name of the GitHub project. Defaults to the project this folder was cloned from or pushed to.

.PARAMETER Schedule
  Sets up a weekly Windows task that runs this script (default Mondays 06:30, before the live page's Monday
  refresh). Use -Day and -Time to choose another slot. -Unschedule removes the task.

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File .\Send-EventbriteFeed.ps1
  powershell -NoProfile -ExecutionPolicy Bypass -File .\Send-EventbriteFeed.ps1 -Schedule
#>
param(
  [string]$Repo,
  [switch]$Schedule,
  [switch]$Unschedule,
  [ValidateSet('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')]
  [string]$Day = 'Monday',
  [string]$Time = '06:30'
)

$ErrorActionPreference = 'Stop'
$root     = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$project  = Split-Path $root -Parent
$TaskName = "Dan's Calgary Live Music Guide - send Eventbrite"
# keep a record of the last run, since the weekly task runs without a window
$logFile = Join-Path $env:TEMP 'DansCalgaryLiveMusicGuide-send-eventbrite.log'
if (-not $Schedule -and -not $Unschedule) { try { Start-Transcript -Path $logFile -Force | Out-Null } catch { } }
trap { Write-Host "Stopped: $($_.Exception.Message)" -ForegroundColor Red; try { Stop-Transcript | Out-Null } catch { }; exit 1 }

if ($Unschedule) {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
  Write-Host "Removed the weekly task ""$TaskName""."
  return
}

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { throw 'The GitHub command-line tool (gh) is not installed on this PC.' }
$authText = (cmd /c 'gh auth status 2>&1') -join "`n"
if ($LASTEXITCODE -ne 0) {
  Write-Host $authText
  Write-Host ("gh: {0}; APPDATA: {1}; sign-in file found: {2}" -f (Get-Command gh).Source, $env:APPDATA, (Test-Path (Join-Path $env:APPDATA 'GitHub CLI\hosts.yml')))
  $ghDir = Join-Path $env:APPDATA 'GitHub CLI'
  Write-Host ("files in {0}: {1}" -f $ghDir, $(if (Test-Path $ghDir) { (Get-ChildItem $ghDir -Force | ForEach-Object Name) -join ', ' } else { '(folder missing)' }))
  Write-Host ("saved GitHub credentials (names only): {0}" -f ((cmd /c 'cmdkey /list' | Select-String 'gh:|github' | ForEach-Object { $_.Line.Trim() }) -join '; '))
  throw 'GitHub is not signed in on this PC. Run: gh auth login --web'
}
if (-not $Repo) {
  $url = (cmd /c "git -C ""$project"" remote get-url origin 2>nul")
  $Repo = ([string]$url).Trim() -replace '^https://github\.com/', '' -replace '\.git$', ''
  if ($Repo -notmatch '^[\w.-]+/[\w.-]+$') { throw 'Could not tell which GitHub project to send to. Run again with -Repo owner/name.' }
}

if ($Schedule) {
  if ($Time -notmatch '^([01]?\d|2[0-3]):([0-5]\d)$') { throw "Time must be like 06:30, not '$Time'." }
  # On a mapped network drive, the task uses the full network path, which works even when the letter isn't mapped
  $self = $PSCommandPath
  $drive = Get-PSDrive -Name $self.Substring(0, 1) -ErrorAction SilentlyContinue
  if ($drive -and $drive.DisplayRoot -like '\\*') { $self = $drive.DisplayRoot.TrimEnd('\') + $self.Substring(2) }
  $action  = New-ScheduledTaskAction -Execute 'powershell.exe' `
               -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -Repo {1}' -f $self, $Repo)
  $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $Day -At ([datetime]::Today.AddHours([int]$Matches[1]).AddMinutes([int]$Matches[2]))
  # local clock time with no UTC offset, so it stays at the same hour after daylight saving changes
  $inv = [Globalization.CultureInfo]::InvariantCulture
  $trigger.StartBoundary = [DateTimeOffset]::Parse([string]$trigger.StartBoundary, $inv).LocalDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss", $inv)
  # if the PC is off or asleep then, it runs at the next chance; the upload refreshes the page itself
  $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RunOnlyIfNetworkAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                -ExecutionTimeLimit (New-TimeSpan -Minutes 15)
  Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Force `
    -Description "Sends this PC's Eventbrite listings to the live page of $Repo" | Out-Null
  Write-Host "Scheduled ""$TaskName"" every $Day at $Time." -ForegroundColor Green
  Write-Host "Each run's output is saved to $logFile"
  return
}

# 1. read Eventbrite from this PC
$tmp = Join-Path $env:TEMP ('calgary-feed-' + [guid]::NewGuid().ToString('N'))
try {
  Write-Host "Reading Eventbrite listings from this PC (about a minute)..."
  & (Join-Path $root 'Update-CalgaryConcerts.ps1') -ExportFeed $tmp

  # 2. upload each list to its place in the project, under the account's private no-reply address
  $me = ((cmd /c 'gh api user') -join "`n") | ConvertFrom-Json
  $who = @{ name = $me.login; email = ('{0}+{1}@users.noreply.github.com' -f $me.id, $me.login) }
  $cfg = Get-Content (Join-Path $root 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
  foreach ($src in @($cfg.sources | Where-Object { $_.enabled -and $_.feed })) {
    $file = Join-Path $tmp (Split-Path $src.feed -Leaf)
    if (-not (Test-Path $file)) { continue }
    $sha = cmd /c "gh api repos/$Repo/contents/$($src.feed) --jq .sha 2>nul"
    $body = [ordered]@{
      message   = "Update $($src.name) list for the live page"
      content   = [Convert]::ToBase64String([IO.File]::ReadAllBytes($file))
      committer = $who
      author    = $who
    }
    if ($LASTEXITCODE -eq 0 -and $sha) { $body.sha = ([string]$sha).Trim() }
    $bodyFile = Join-Path $tmp 'body.json'
    [IO.File]::WriteAllText($bodyFile, ($body | ConvertTo-Json -Compress), (New-Object Text.UTF8Encoding $false))
    $res = cmd /c "gh api -X PUT repos/$Repo/contents/$($src.feed) --input ""$bodyFile"" 2>&1"
    if ($LASTEXITCODE -ne 0) { throw "Upload to GitHub failed: $res" }
    Write-Host "Sent $(@(Import-Csv $file).Count) $($src.name) shows to $Repo/$($src.feed)." -ForegroundColor Green
  }

  # 3. refresh the live page now
  cmd /c "gh workflow run live-page.yml --repo $Repo >nul 2>&1"
  if ($LASTEXITCODE -eq 0) {
    $site = cmd /c "gh api repos/$Repo/pages --jq .html_url 2>nul"
    Write-Host "The live page is refreshing; it updates in about 3 minutes: $site" -ForegroundColor Green
  } else {
    Write-Host 'The list was sent; the live page picks it up at its next weekly refresh.' -ForegroundColor Yellow
  }
} finally {
  Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
  try { Stop-Transcript | Out-Null } catch { }
}
