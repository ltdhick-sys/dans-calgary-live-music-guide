# Dan's Calgary Live Music Guide: installer and launcher in one file.
# Build-Setup.ps1 puts this after a small batch header, a $Version line and the base64 $Payload (a zip
# of the program files and the current show list), producing "Dan's Calgary Live Music Guide.cmd".
# The first double-click installs the program (or upgrades an older copy) and opens the guide.
# After that, double-clicking the same file just opens the guide.
# Keep this file plain ASCII: it runs from inside a .cmd file.
#
# Optional environment variables (for testing or unattended installs):
#   SETUP_INSTALLDIR  install somewhere other than %LOCALAPPDATA%\Programs\CalgaryConcerts
#   SETUP_QUIET       messages only: don't open the guide or start a background refresh
#   SETUP_NOSYSTEM    skip the scheduled task, shortcuts and Settings > Apps entry
#   SETUP_FORCE       reinstall even when this version is already installed
#   SETUP_WHEN        weekly update day and time, e.g. "Sunday 18:00" (default Monday 07:00)

$ErrorActionPreference = 'Stop'
function Say([string]$m, [string]$c = 'Gray') { Write-Host $m -ForegroundColor $c }
$log = Join-Path $env:TEMP 'DansCalgaryLiveMusicGuide-setup.log'
# any error that isn't handled below stops setup with a readable message and a non-zero exit,
# so the window stays open instead of closing before you can read it
trap {
  Say ''
  Say "  Setup stopped: $($_.Exception.Message)" 'Red'
  if (Test-Path $log) { Say "  Details were saved to $log" 'Yellow' }
  try { Stop-Transcript | Out-Null } catch { }
  exit 1
}
$AppName  = "Dan's Calgary Live Music Guide"
$OldName  = 'Calgary Live Music Guide'   # name used by the first version
$TaskName = "Dan's Calgary Live Music Guide update"
$RegKey   = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\CalgaryLiveMusicGuide'
$quiet    = [bool]$env:SETUP_QUIET
$noSystem = [bool]$env:SETUP_NOSYSTEM
$dir      = if ($env:SETUP_INSTALLDIR) { $env:SETUP_INSTALLDIR } else { Join-Path $env:LOCALAPPDATA 'Programs\CalgaryConcerts' }
$guide    = Join-Path $dir 'output\calgary-live-music.html'

# Locked-down company PCs run PowerShell in a restricted mode that can't install anything.
# Say so plainly instead of failing with a wall of red text.
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
  Say ''
  Say "  This PC's IT policy restricts PowerShell ($($ExecutionContext.SessionState.LanguageMode)), so $AppName can't be installed here." 'Yellow'
  Say '  Ask IT to allow PowerShell scripts for your account, or use the online copy of the guide instead.' 'Yellow'
  exit 2
}

function Find-Edge {
  foreach ($p in "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") {
    if ($p -and (Test-Path $p)) { return $p }
  }
  return $null
}
# The guide opens in its own Edge window, like a program. Without Edge, the default browser opens it.
function Open-Guide {
  if ($quiet) { Say "  (quiet mode: would open $guide)"; return }
  $edge = Find-Edge
  if ($edge) { Start-Process $edge -ArgumentList ('--app="{0}"' -f ([uri]$guide).AbsoluteUri) }
  else { Start-Process $guide }
}

# ---------------------------------------------------------------- already installed: just open it

$marker = Join-Path $dir '.installed'
$installed = if (Test-Path $marker) { ((Get-Content $marker -Raw).Trim() -split ' ')[-1] } else { '' }
if ($installed -and $installed -ge $Version -and (Test-Path $guide) -and -not $env:SETUP_FORCE) {
  Say "  Opening $AppName..."
  Open-Guide
  exit 0
}

# ---------------------------------------------------------------- install or upgrade

Say ''
Say "  $AppName  -  setting up version $Version" 'Cyan'
Say "  Installing to $dir"
Say ''
try { Start-Transcript -Path $log -Force | Out-Null } catch { }

try {
  # 1. unpack
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $tmp = Join-Path $env:TEMP ('calgary-setup-' + [guid]::NewGuid().ToString('N'))
  New-Item -ItemType Directory -Force $tmp | Out-Null
  try {
    $zip = Join-Path $tmp 'payload.zip'
    [IO.File]::WriteAllBytes($zip, [Convert]::FromBase64String(($Payload -replace '\s', '')))
    $src = Join-Path $tmp 'files'
    [IO.Compression.ZipFile]::ExtractToDirectory($zip, $src)

    $upgrade = Test-Path (Join-Path $dir 'Update-CalgaryConcerts.ps1')
    New-Item -ItemType Directory -Force $dir, (Join-Path $dir 'data'), (Join-Path $dir 'output') | Out-Null

    # program files are always replaced
    foreach ($f in 'Update-CalgaryConcerts.ps1', 'template.html', 'Install-WeeklyTask.ps1', 'Run-Update.cmd', 'Edit-Watchlist.cmd', 'Uninstall.cmd', 'README.md') {
      Copy-Item (Join-Path $src $f) (Join-Path $dir $f) -Force
    }

    # Settings: take the new defaults, but keep a Ticketmaster key already entered on this PC (whatever the new
    # build carries). Your own changes belong in data\config.local.json, which setup never touches.
    $cfgPath = Join-Path $dir 'config.json'
    $newCfg = [IO.File]::ReadAllText((Join-Path $src 'config.json'), [Text.Encoding]::UTF8)
    if (Test-Path $cfgPath) {
      try {
        $old = [IO.File]::ReadAllText($cfgPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
        $oldKey = [string](@($old.sources | Where-Object { $_.type -eq 'ticketmaster' })[0].api_key)
        if ($oldKey) { $newCfg = $newCfg -replace '("type":\s*"ticketmaster",\s*"api_key":\s*)"[^"]*"', ('${1}"' + $oldKey + '"') }
      } catch { }
      Copy-Item $cfgPath (Join-Path $dir 'config.previous.json') -Force
    }
    [IO.File]::WriteAllText($cfgPath, $newCfg, (New-Object Text.UTF8Encoding $false))

    # the show list, watchlist, ignore list, manual rows, venue locations and last output are only seeded when missing.
    # The refresh history goes with the seeded show list, so it is only copied on a first install.
    foreach ($f in 'data\shows.csv', 'data\watchlist.txt', 'data\ignore.txt', 'data\manual.tsv', 'data\venues.csv', 'data\runs.csv', 'data\source-runs.csv',
                   'output\calgary-live-music.html', 'output\calgary-live-music.csv') {
      $to = Join-Path $dir $f
      if ($upgrade -and $f -in 'data\runs.csv', 'data\source-runs.csv') { continue }
      if (-not (Test-Path $to) -and (Test-Path (Join-Path $src $f))) { Copy-Item (Join-Path $src $f) $to }
    }

    # on an upgrade, redraw the existing page with the new program so page fixes apply right away
    if ($upgrade -and (Test-Path $guide)) {
      $redraw = Start-Process powershell.exe -WindowStyle Hidden -Wait -PassThru -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
                  ('"{0}"' -f (Join-Path $dir 'Update-CalgaryConcerts.ps1')), '-RenderOnly')
      if ($redraw.ExitCode -ne 0) {
        # fall back to pouring the old page's data into the new template
        $oldPage = [IO.File]::ReadAllText($guide, [Text.Encoding]::UTF8)
        $d = [regex]::Match($oldPage, '<script id="data" type="application/json">(.*?)</script>', 'Singleline')
        $i = [regex]::Match($oldPage, '<script id="info" type="application/json">(.*?)</script>', 'Singleline')
        if ($d.Success -and $i.Success) {
          $tpl = [IO.File]::ReadAllText((Join-Path $dir 'template.html'), [Text.Encoding]::UTF8)
          [IO.File]::WriteAllText($guide, $tpl.Replace('__DATA__', $d.Groups[1].Value).Replace('__INFO__', $i.Groups[1].Value), (New-Object Text.UTF8Encoding $false))
        }
      }
    }
    [IO.File]::WriteAllText($marker, "$AppName $Version")
  } finally {
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
  }
  if ($upgrade) { Say '  Program updated. Your show list, watchlist, ignore list, manual rows and own settings were kept.' 'Green' }
  else { Say '  Program files installed.' 'Green' }

  # 2. weekly task, shortcuts, Settings > Apps entry
  if (-not $noSystem) {
    # clear out the task and shortcuts from the first version, so an upgrade doesn't leave two weekly runs
    try { Unregister-ScheduledTask -TaskName 'Calgary Concerts Update' -Confirm:$false -ErrorAction Stop } catch { }
    Remove-Item (Join-Path ([Environment]::GetFolderPath('Programs')) $OldName) -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item (Join-Path ([Environment]::GetFolderPath('Desktop')) "$OldName.lnk") -Force -ErrorAction SilentlyContinue

    $days = @{ mon = 'Monday'; tue = 'Tuesday'; wed = 'Wednesday'; thu = 'Thursday'; fri = 'Friday'; sat = 'Saturday'; sun = 'Sunday' }
    $day = 'Monday'; $time = '07:00'
    $m = [regex]::Match([string]$env:SETUP_WHEN, '^(?i)(mon|tue|wed|thu|fri|sat|sun)[a-z]*\s+([01]?\d|2[0-3]):([0-5]\d)$')
    if ($m.Success) { $day = $days[$m.Groups[1].Value.ToLower()]; $time = '{0:00}:{1}' -f [int]$m.Groups[2].Value, $m.Groups[3].Value }
    try {
      $action   = New-ScheduledTaskAction -Execute 'powershell.exe' -WorkingDirectory $dir `
                    -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}"' -f (Join-Path $dir 'Update-CalgaryConcerts.ps1'))
      # an upgrade keeps the day and time already chosen on this PC, unless SETUP_WHEN asks for another
      $existing = $null
      try { $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop } catch { }
      if ($existing -and -not $m.Success -and @($existing.Triggers).Count) {
        $trigger = @($existing.Triggers)
        $whenText = 'on the schedule already set on this PC'
      } else {
        $hm = $time.Split(':')
        $trigger = @(New-ScheduledTaskTrigger -Weekly -DaysOfWeek $day -At ([datetime]::Today.AddHours([int]$hm[0]).AddMinutes([int]$hm[1])))
        $whenText = "every $day at $time"
      }
      # Windows stores the start as UTC ("...13:00:00Z"); rewrite it as local clock time with no offset, so the run
      # stays at the same hour after daylight saving changes
      foreach ($t in $trigger) {
        $inv = [Globalization.CultureInfo]::InvariantCulture
        $t.StartBoundary = [DateTimeOffset]::Parse([string]$t.StartBoundary, $inv).LocalDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss", $inv)
      }
      # laptops: run on battery too, and don't stop halfway if the charger is unplugged
      $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RunOnlyIfNetworkAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
      Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Force `
        -Description "Refreshes $AppName" | Out-Null
      Say "  The list refreshes itself $whenText." 'Green'
    } catch {
      Say "  Could not set up the weekly refresh: $($_.Exception.Message)" 'Yellow'
      Say '  The guide still works. Use "Update concert list now" in the Start menu to refresh it.' 'Yellow'
    }
    $policy = @('MachinePolicy', 'UserPolicy' | ForEach-Object { [string](Get-ExecutionPolicy -Scope $_) }) | Where-Object { $_ -in 'AllSigned', 'Restricted' }
    if ($policy) { Say "  Note: IT policy sets PowerShell to $($policy[0]), which may stop the weekly refresh. The guide still opens with the list it has." 'Yellow' }

    $edge = Find-Edge
    $guideTarget = if ($edge) { $edge } else { $guide }
    $guideArgs   = if ($edge) { '--app="{0}"' -f ([uri]$guide).AbsoluteUri } else { '' }
    $ws = New-Object -ComObject WScript.Shell
    function New-Link([string]$path, [string]$target, [string]$arguments, [string]$desc, [string]$workDir = $dir) {
      $s = $ws.CreateShortcut($path)
      $s.TargetPath = $target; $s.Arguments = $arguments; $s.WorkingDirectory = $workDir; $s.Description = $desc
      $s.Save()
    }
    $menu = Join-Path ([Environment]::GetFolderPath('Programs')) $AppName
    New-Item -ItemType Directory -Force $menu | Out-Null
    New-Link (Join-Path $menu "$AppName.lnk") $guideTarget $guideArgs 'Open the concert list'
    New-Link (Join-Path $menu 'Update concert list now.lnk') (Join-Path $dir 'Run-Update.cmd') '' 'Refresh the list from the listing sites'
    New-Link (Join-Path $menu 'Edit watchlist.lnk') (Join-Path $dir 'Edit-Watchlist.cmd') '' 'Artists and venues to be notified about'
    New-Link (Join-Path $menu 'Settings and sources (README).lnk') (Join-Path $env:SystemRoot 'notepad.exe') ('"' + (Join-Path $dir 'README.md') + '"') 'How it works and how to change it'
    # started from the temp folder, so the uninstaller isn't sitting in the folder it removes
    New-Link (Join-Path $menu "Uninstall $AppName.lnk") (Join-Path $dir 'Uninstall.cmd') '' 'Remove the program and its weekly task' $env:TEMP
    New-Link (Join-Path ([Environment]::GetFolderPath('Desktop')) "$AppName.lnk") $guideTarget $guideArgs 'Open the concert list'

    New-Item -Path $RegKey -Force | Out-Null
    $props = [ordered]@{
      DisplayName = $AppName; DisplayVersion = $Version; Publisher = $AppName
      InstallLocation = $dir; UninstallString = ('"' + (Join-Path $dir 'Uninstall.cmd') + '"')
      NoModify = 1; NoRepair = 1
    }
    foreach ($k in $props.Keys) {
      $type = if ($props[$k] -is [int]) { 'DWord' } else { 'String' }
      New-ItemProperty -Path $RegKey -Name $k -Value $props[$k] -PropertyType $type -Force | Out-Null
    }
    Say '  Added to the Start menu, the desktop and Settings > Apps.' 'Green'
  }
} catch {
  Say ''
  Say "  Setup stopped: $($_.Exception.Message)" 'Red'
  Say "  A log was saved to $log" 'Yellow'
  try { Stop-Transcript | Out-Null } catch { }
  exit 1
}
try { Stop-Transcript | Out-Null } catch { }

# ---------------------------------------------------------------- open it, refresh in the background

Open-Guide
if (-not $quiet) {
  # the guide opens right away with the list it shipped with; this fetches the latest and says when it's done
  Start-Process powershell.exe -WindowStyle Hidden -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f (Join-Path $dir 'Update-CalgaryConcerts.ps1')), '-NotifyDone')
  Say '  The list is being refreshed in the background; a notification says when it is done (about 3 minutes).' 'Cyan'
}
Say ''
Say "  Done. Next time, open it from the desktop or Start menu, or double-click this same file again." 'Cyan'
if (-not $quiet) { Start-Sleep -Seconds 5 }
exit 0
