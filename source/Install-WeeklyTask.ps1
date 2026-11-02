<#
.SYNOPSIS
  Schedules Update-CalgaryConcerts.ps1 to run weekly with Windows Task Scheduler.

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File .\Install-WeeklyTask.ps1
  powershell -NoProfile -ExecutionPolicy Bypass -File .\Install-WeeklyTask.ps1 -Day Sunday -Time 18:00
  powershell -NoProfile -ExecutionPolicy Bypass -File .\Install-WeeklyTask.ps1 -Remove
#>
param(
  [ValidateSet('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')]
  [string]$Day = 'Monday',
  [string]$Time = '07:00',
  [string]$TaskName = "Dan's Calgary Live Music Guide update",
  [switch]$Remove
)

$ErrorActionPreference = 'Stop'
$root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

if ($Remove) {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
  Write-Host "Removed scheduled task ""$TaskName""."
  return
}

$script   = Join-Path $root 'Update-CalgaryConcerts.ps1'
$action   = New-ScheduledTaskAction -Execute 'powershell.exe' -WorkingDirectory $root `
              -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}"' -f $script)
if ($Time -notmatch '^([01]?\d|2[0-3]):([0-5]\d)$') { throw "Time must be like 07:00 or 18:30, not '$Time'." }
$trigger  = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $Day -At ([datetime]::Today.AddHours([int]$Matches[1]).AddMinutes([int]$Matches[2]))
# Windows stores the start as UTC ("...13:00:00Z"); rewrite it as local clock time with no offset, so the run
# stays at the same hour after daylight saving changes
$inv = [Globalization.CultureInfo]::InvariantCulture
$trigger.StartBoundary = [DateTimeOffset]::Parse([string]$trigger.StartBoundary, $inv).LocalDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss", $inv)
# StartWhenAvailable: if the PC was off or asleep at the scheduled time, run at the next chance.
# Laptops: run on battery too, and don't stop halfway if the charger is unplugged.
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RunOnlyIfNetworkAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
              -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

# Runs as you, only while you're signed in, so mapped drives like V: are available
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Force `
  -Description "Refreshes Dan's Calgary Live Music Guide: output\calgary-live-music.html and .csv" | Out-Null

Write-Host "Scheduled ""$TaskName"" every $Day at $Time."
Write-Host "Run it now:  Start-ScheduledTask -TaskName ""$TaskName"""
Write-Host "Remove it:   .\Install-WeeklyTask.ps1 -Remove"
