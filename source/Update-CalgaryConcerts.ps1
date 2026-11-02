<#
.SYNOPSIS
  Refreshes the Calgary live-music list and rebuilds the HTML guide and CSV.

.DESCRIPTION
  Reads every enabled source in config.json (Songkick, Eventbrite, Showpass, Live Nation venue pages,
  and the Ticketmaster API when a key is set), merges what it finds into data\shows.csv, drops shows
  whose date has passed, and writes:
    output\calgary-live-music.html   the filterable guide (open in any browser)
    output\calgary-live-music.csv    the same list for Excel
    output\changes\<date>.txt        what was added or cancelled this run, plus per-source counts

  Shows stay in data\shows.csv until their date passes, even if a source stops listing them.
  Rows in data\manual.tsv are merged on every run, for shows you want to add by hand.

.PARAMETER Open
  Opens the HTML guide when the run finishes.

.PARAMETER DryRun
  Fetches and reports what would change without writing anything.

.PARAMETER RenderOnly
  Skips the listing sites and only redraws the guide from the saved list, for example after
  editing data\watchlist.txt. Nothing in data\ is changed.

.PARAMETER NotifyDone
  Shows a Windows notification when the refresh finishes (used for the first refresh after install).

  Settings: config.json holds the defaults. Put your own changes in data\config.local.json (same layout,
  only the settings you change); it overrides config.json and is never touched by an upgrade.

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File .\Update-CalgaryConcerts.ps1 -Open
#>
[CmdletBinding()]
param(
  [string]$ConfigPath,
  [switch]$Open,
  [switch]$DryRun,
  [switch]$RenderOnly,
  [switch]$NotifyDone
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'   # the progress bar slows every download in Windows PowerShell 5.1
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Dates, times and numbers are written the same way on every PC, whatever its regional settings (a Thai or
# Arabic calendar would otherwise make every show look past). Page labels use English (Canada) names.
$listSep = (Get-Culture).TextInfo.ListSeparator   # Excel's column separator on this PC, for the CSV copy
$inv = [Globalization.CultureInfo]::InvariantCulture
$en  = [Globalization.CultureInfo]::GetCultureInfo('en-CA')
[Threading.Thread]::CurrentThread.CurrentCulture = $inv
function D([datetime]$d, [string]$fmt) { return $d.ToString($fmt, $inv) }
function Label([datetime]$d, [string]$fmt) { return $d.ToString($fmt, $en) }

$log = New-Object System.Collections.Generic.List[string]
function Log([string]$m) { $line = '[{0}] {1}' -f (D (Get-Date) 'HH:mm:ss'), $m; Write-Host $line; $log.Add($line) }

$root     = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $ConfigPath) { $ConfigPath = Join-Path $root 'config.json' }
$dataDir  = Join-Path $root 'data'
$outDir   = Join-Path $root 'output'
$storeCsv = Join-Path $dataDir 'shows.csv'
$htmlPath = Join-Path $outDir 'calgary-live-music.html'
$cfg      = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
# your own settings override config.json and survive upgrades
$localCfg = Join-Path $dataDir 'config.local.json'
if (Test-Path $localCfg) {
  try {
    $over = Get-Content $localCfg -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($p in $over.PSObject.Properties) {
      if ($cfg.PSObject.Properties[$p.Name]) { $cfg.($p.Name) = $p.Value } else { $cfg | Add-Member NoteProperty $p.Name $p.Value }
    }
    Log "Using your settings from data\config.local.json ($(@($over.PSObject.Properties).Count) changed)."
  } catch { Log "data\config.local.json could not be read, so it was ignored: $($_.Exception.Message)" }
}

# The Ticketmaster key can live in data\ticketmaster.key instead of config.json, which keeps it out of shared copies.
$tmKeyFile = Join-Path $dataDir 'ticketmaster.key'
if (Test-Path $tmKeyFile) {
  $fileKey = ([IO.File]::ReadAllText($tmKeyFile)).Trim()
  foreach ($s in @($cfg.sources | Where-Object { $_.type -eq 'ticketmaster' })) { if ($fileKey -and -not $s.api_key) { $s.api_key = $fileKey } }
}

function Open-Guide {
  # open in its own Edge window, like a program; fall back to the default browser
  $edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") |
    Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
  if ($edge) { Start-Process $edge -ArgumentList ('--app="{0}"' -f ([uri]$htmlPath).AbsoluteUri) } elseif (Test-Path $htmlPath) { Start-Process $htmlPath }
}

# One update at a time: the weekly task, "Update concert list now", Edit watchlist and the first refresh after
# install can overlap. A redraw waits up to 2 minutes for a running update; a second full update just stops.
if (-not $DryRun) {
  $lock = New-Object Threading.Mutex($false, 'Local\DansCalgaryLiveMusicGuide')
  $haveLock = $false
  try { $haveLock = $lock.WaitOne($(if ($RenderOnly) { 120000 } else { 0 })) } catch [Threading.AbandonedMutexException] { $haveLock = $true }
  if (-not $haveLock) {
    Log 'Another update is already running, so this one stopped. The guide is redrawn when that one finishes.'
    if ($Open) { Open-Guide }
    exit 0
  }
}

$today    = (Get-Date).Date
$todayS   = D $today 'yyyy-MM-dd'
$horizonS = D $today.AddDays([int]$cfg.window_days) 'yyyy-MM-dd'
$newSince = D $today.AddDays(-[int]$cfg.new_badge_days) 'yyyy-MM-dd'
# a run that hangs on slow sites stops reading new pages after 20 minutes (the weekly task is stopped at 30)
$deadline = (Get-Date).AddMinutes(20)
$partial = @{}   # sources that stopped partway, so their results don't count toward "may be cancelled"

# ---------------------------------------------------------------- helpers

function Clean([string]$s) {
  if (-not $s) { return '' }
  $s = [System.Net.WebUtility]::HtmlDecode($s)
  return ($s -replace '[\t\r\n]+', ' ' -replace '\s{2,}', ' ').Trim()
}
function Norm([string]$s) {
  if (-not $s) { return '' }
  # strip accents so "Angelique" matches the accented spelling, and treat "&" as "and"
  $d = $s.Normalize([Text.NormalizationForm]::FormD)
  $plain = -join ($d.ToCharArray() | Where-Object { [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne 'NonSpacingMark' })
  return ($plain.ToLowerInvariant() -replace '&', 'and' -replace '^the\s+', '' -replace '[^a-z0-9]', '')
}
# like Norm but keeps one space between words, for whole-word matching
function Words([string]$s) {
  if (-not $s) { return '' }
  $d = $s.Normalize([Text.NormalizationForm]::FormD)
  $plain = -join ($d.ToCharArray() | Where-Object { [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne 'NonSpacingMark' })
  return ($plain.ToLowerInvariant() -replace '&', ' and ' -replace '[^a-z0-9]+', ' ').Trim()
}
# a real yyyy-MM-dd calendar date ("2026-11-31" and "05/11/2026" are not)
function IsDate([string]$s) { $d = [datetime]::MinValue; return [datetime]::TryParseExact([string]$s, 'yyyy-MM-dd', $inv, 'None', [ref]$d) }
function DayName([string]$s) { $d = [datetime]::MinValue; if ([datetime]::TryParseExact($s, 'yyyy-MM-dd', $inv, 'None', [ref]$d)) { return Label $d 'ddd' }; return '' }

function Get-Text([string]$url, [hashtable]$headers = @{}) {
  if ((Get-Date) -gt $deadline) { throw 'Time limit for this run reached; the rest is left for next time.' }
  $wait = [int]$cfg.request_delay_ms
  for ($try = 1; $try -le 2; $try++) {
    try {
      Start-Sleep -Milliseconds $wait
      $r = Invoke-WebRequest -Uri $url -UseBasicParsing -UserAgent $cfg.user_agent -TimeoutSec 40 -Headers $headers
      return [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray())
    } catch {
      $code = try { [int]$_.Exception.Response.StatusCode } catch { 0 }
      if ($try -eq 2 -or $code -in 401, 403, 404) { throw }
      $wait = 4000
    }
  }
}

function VenueAlias([string]$v) {
  $v = Clean $v
  $v = ($v -replace '\s*\([^)]*\d[^)]*\)', '').Trim()           # drop street-address parentheticals
  $v = ($v -replace ',\s*(calgary|ab|alberta|canada)\b.*$', '').Trim()
  foreach ($a in $cfg.venue_aliases) { if ($v -match $a.pattern) { return $a.name } }
  return $v
}

function GenreBucket([string]$g) {
  switch -Regex ($g) {
    'festival' { return 'Festival' }
    'tribute' { return 'Tribute' }
    'metal' { return 'Metal' }
    'punk|hardcore|\bemo\b' { return 'Punk' }
    'edm|electro|techno|\bhouse\b|dubstep|drum|\bdj\b|bass music|dance' { return 'EDM / DJ' }
    'hip.?hop|r&b|\brap\b|afrobeat|soul' { return 'Hip-hop, R&B & soul' }
    'country|western' { return 'Country' }
    'jazz|big band|standards|swing' { return 'Jazz' }
    'blues' { return 'Blues' }
    'classical|choral|choir|cappella|orchestra|symphon|philharmonic|film score|piano|carnatic|opera|gospel|christian|candlelight' { return 'Classical, choral & gospel' }
    'folk|roots|celtic|irish|singer|acoustic|americana|bluegrass|fiddle' { return 'Folk & roots' }
    'latin|reggae|world|balkan|flamenco|caribbean' { return 'World & Latin' }
    'indie|shoegaze' { return 'Indie' }
    'rock|psych' { return 'Rock' }
    'pop' { return 'Pop' }
  }
  return 'Other'
}

# "2026-11-13T20:00:00-07:00" or "20:00:00" -> "20:00"; midnight usually means "no time given"
function TimeOf([string]$s) {
  $m = [regex]::Match($s, '(?:T|^)([01]\d|2[0-3]):([0-5]\d)')
  if (-not $m.Success -or ($m.Groups[1].Value -eq '00' -and $m.Groups[2].Value -eq '00')) { return '' }
  return '{0}:{1}' -f $m.Groups[1].Value, $m.Groups[2].Value
}
function Num($x) { $d = 0.0; if ([double]::TryParse([string]$x, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }; return $null }
function PriceStr($d) { if ($null -eq $d) { return '' }; return ([double]$d).ToString('0.##', [Globalization.CultureInfo]::InvariantCulture) }

# start time written in the notes, e.g. "8 PM", "7:30pm", "happy hour set 5-7pm" (takes the first time of a range)
function NotesTime([string]$n) {
  if ($n -match '(?i)\bnoon\b') { return '12:00' }
  $m = [regex]::Match($n, '(?i)\b(1[0-2]|0?[1-9])(?::([0-5]\d))?\s*(?:-|to|\u2013)\s*(?:1[0-2]|0?[1-9])(?::[0-5]\d)?\s*([ap])\.?m\b')
  if (-not $m.Success) { $m = [regex]::Match($n, '(?i)\b(1[0-2]|0?[1-9])(?::([0-5]\d))?\s*([ap])\.?m\b') }
  if (-not $m.Success) { return '' }
  $h = [int]$m.Groups[1].Value % 12
  if ($m.Groups[3].Value -match '(?i)p') { $h += 12 }
  return '{0:00}:{1}' -f $h, $(if ($m.Groups[2].Success) { $m.Groups[2].Value } else { '00' })
}
# prices written in the notes, e.g. "$20 adv / $25 door"; "free" or "no cover" counts as 0
function NotesPrice([string]$n) {
  $vals = @([regex]::Matches($n, '\$\s?(\d{1,4}(?:\.\d{1,2})?)') | ForEach-Object { Num $_.Groups[1].Value })
  if ($vals.Count) { return ,@(($vals | Measure-Object -Minimum).Minimum, ($vals | Measure-Object -Maximum).Maximum) }
  if ($n -match '(?i)(^|[;,.]\s*)free\b(?!\s+(parking|for|with|shuttle))|\bfree (admission|entry|show|concert|event)\b|\bno cover\b') { return ,@(0, 0) }
  return ,@($null, $null)
}

# Calgary area for a venue: a config override, else the nearest named neighbourhood, else the city quadrant
function AreaOf([string]$venue) {
  foreach ($o in $cfg.venue_areas) { if ($venue -match $o.pattern) { return $o.area } }
  $vi = $venueInfo[(Norm $venue)]
  $lat = if ($vi) { Num $vi.lat } else { $null }; $lon = if ($vi) { Num $vi.lon } else { $null }
  if ($null -ne $lat -and $null -ne $lon -and $lat -ne 0) {
    $best = $null; $bestKm = 1e9
    foreach ($a in $cfg.areas) {
      $dy = ($lat - $a.lat) * 111.2; $dx = ($lon - $a.lon) * 111.2 * [math]::Cos($lat * [math]::PI / 180)
      $km = [math]::Sqrt($dx * $dx + $dy * $dy)
      if ($km -le $a.radius_km -and $km -lt $bestKm) { $best = $a.name; $bestKm = $km }
    }
    if ($best) { return $best }
    $c = $cfg.city
    if ($lat -lt $c.south -or $lat -gt $c.north -or $lon -lt $c.west -or $lon -gt $c.east) { return 'Outside Calgary' }
    return ('{0}{1} Calgary' -f $(if ($lat -ge $c.ns_split_lat) { 'N' } else { 'S' }), $(if ($lon -ge $c.ew_split_lon) { 'E' } else { 'W' }))
  }
  if ($vi -and $vi.city -and $vi.city -notmatch '(?i)calgary') { return 'Outside Calgary' }
  if ($vi -and $vi.address -match '(?i)\b(NE|NW|SE|SW)\b') { return ($Matches[1].ToUpperInvariant() + ' Calgary') }
  if ($venue -cmatch '\b(NE|NW|SE|SW)\b') { return ($Matches[1] + ' Calgary') }   # venue listed as a street address
  return 'Area not known'
}

# Windows notification; clicking it opens the guide. Returns $false if Windows won't show it.
function Show-Toast([string]$title, [string[]]$lines, [string]$openPath) {
  try {
    $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
    $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
    $esc = [Security.SecurityElement]
    $texts = (@($title) + @($lines | Select-Object -First 2) | ForEach-Object { '<text>' + $esc::Escape($_) + '</text>' }) -join ''
    $xml = '<toast activationType="protocol" launch="{0}"><visual><binding template="ToastGeneric">{1}</binding></visual></toast>' -f $esc::Escape(([uri]$openPath).AbsoluteUri), $texts
    $doc = New-Object Windows.Data.Xml.Dom.XmlDocument
    $doc.LoadXml($xml)
    # PowerShell's own app id, which Windows already allows to show notifications
    $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show((New-Object Windows.UI.Notifications.ToastNotification $doc))
    return $true
  } catch {
    Log "Could not show a Windows notification: $($_.Exception.Message)"
    return $false
  }
}

function TierOf([string]$venue, [string]$gb) {
  $v = $venue.ToLowerInvariant()
  if ($gb -eq 'Festival') { return 'fest' }
  if ($v -match $cfg.tiers.arena) { return 'arena' }
  if ($v -match $cfg.tiers.club) { return 'club' }
  if ($v -match $cfg.tiers.hall) { return 'hall' }
  return 'club'
}

# ---------------------------------------------------------------- store

# time, price_min/max, tracked and sources were added later; older lists get the empty columns on load.
# sources = the listing sites that have reported the show (separated by |). A show is only marked "may be
# cancelled" when those same sites have read fine twice since anyone last saw it.
$cols = 'date', 'artist', 'support', 'venue', 'genre', 'notes', 'source', 'tier', 'time', 'price_min', 'price_max', 'first_seen', 'last_seen', 'origin', 'tracked', 'sources'
$store = New-Object System.Collections.Generic.List[object]
if (Test-Path $storeCsv) {
  Import-Csv $storeCsv -Encoding UTF8 | ForEach-Object {
    $row = $_
    foreach ($c in $cols) { if (-not $row.PSObject.Properties[$c]) { $row | Add-Member NoteProperty $c '' } }
    if (-not $row.tracked -and $row.origin -and $row.origin -ne 'baseline') { $row.tracked = '1' }
    if (-not $row.sources -and $row.origin -and $row.origin -ne 'baseline') { $row.sources = $row.origin }
    if (IsDate $row.date) { $store.Add($row) } else { Log "Skipped a saved show with an unreadable date '$($row.date)': $($row.artist)" }
  }
}

# venue locations reported by the sites (for neighbourhoods and map links)
$venueCsv = Join-Path $dataDir 'venues.csv'
$venueInfo = @{}
if (Test-Path $venueCsv) { Import-Csv $venueCsv -Encoding UTF8 | ForEach-Object { $venueInfo[(Norm $_.venue)] = $_ } }
function Note-Venue($c) {
  if (-not $c.venue) { return }
  $k = Norm $c.venue
  $vi = $venueInfo[$k]
  if (-not $vi) { $vi = [pscustomobject]@{ venue = $c.venue; lat = ''; lon = ''; address = ''; city = '' }; $venueInfo[$k] = $vi }
  if (-not $vi.lat -and $c.lat) { $vi.lat = $c.lat; $vi.lon = $c.lon }
  if (-not $vi.address -and $c.address) { $vi.address = $c.address }
  if (-not $vi.city -and $c.city) { $vi.city = $c.city }
}

# Refresh history. runs.csv has one line per refresh; source-runs.csv one line per source per refresh
# (ok = read completely and found something). Shows are only flagged "may be cancelled" from sources that
# worked, so a source that is switched off, has no key, or is down doesn't flag (or block) anything.
$runsCsv = Join-Path $dataDir 'runs.csv'
$srcRunsCsv = Join-Path $dataDir 'source-runs.csv'
$srcOkDates = @{}
if (Test-Path $srcRunsCsv) {
  Import-Csv $srcRunsCsv | Where-Object { $_.ok -eq '1' } | ForEach-Object {
    if (-not $srcOkDates.ContainsKey($_.source)) { $srcOkDates[$_.source] = New-Object System.Collections.Generic.List[string] }
    if (-not $srcOkDates[$_.source].Contains($_.date)) { $srcOkDates[$_.source].Add($_.date) }
  }
}

# artists and venues to watch for; new matches raise a Windows notification. Whole words only, so "Low"
# doesn't match "Mellow Yellow".
$watchFile = Join-Path $dataDir 'watchlist.txt'
$watch = @(if (Test-Path $watchFile) {
  Get-Content $watchFile -Encoding UTF8 | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') -and (Norm $_).Length -ge 3 }
})
function Watched($r) {
  $hay = ' ' + (Words "$($r.artist) | $($r.support) | $($r.venue)") + ' '
  foreach ($w in $watch) { if ($hay.Contains(' ' + (Words $w) + ' ')) { return $w } }
  return ''
}
$startCount = $store.Count
$past = @($store | Where-Object { $_.date -lt $todayS })
foreach ($p in $past) { [void]$store.Remove($p) }
Log "Loaded $startCount shows; $($past.Count) have passed and will be archived."

$byDate = @{}
foreach ($r in $store) { if (-not $byDate.ContainsKey($r.date)) { $byDate[$r.date] = New-Object System.Collections.Generic.List[object] }; $byDate[$r.date].Add($r) }
$knownVenues = @{}; foreach ($r in $store) { if (-not $knownVenues.ContainsKey((Norm $r.venue))) { $knownVenues[(Norm $r.venue)] = $r.venue } }
$ignoreFile = Join-Path $dataDir 'ignore.txt'
$ignore = @(if (Test-Path $ignoreFile) {
  Get-Content $ignoreFile -Encoding UTF8 | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') } |
    ForEach-Object { try { [void][regex]::new($_); $_ } catch { [regex]::Escape($_) } } |   # plain names with brackets still work
    Where-Object {
      # a line like ".*" or "x|" would match every show and wipe the list
      if ('' -match $_) { Log "data\ignore.txt: '$_' would match every show, so it was skipped."; $false } else { $true }
    }
})
# names added to ignore.txt also come off the existing list
$dropped = @($store | Where-Object { $a = $_.artist; @($ignore | Where-Object { $a -match $_ }).Count -gt 0 })
foreach ($d in $dropped) { [void]$store.Remove($d) }
if ($dropped.Count) { Log "Removed $($dropped.Count) listed shows that match data\ignore.txt." }
$artistGenre = @{}; foreach ($r in $store) { if ($r.genre -and -not $artistGenre.ContainsKey((Norm $r.artist))) { $artistGenre[(Norm $r.artist)] = $r.genre } }

$added = New-Object System.Collections.Generic.List[object]
$cancelled = New-Object System.Collections.Generic.List[object]
$stats = [ordered]@{}

# Same billing? $min is the shortest name allowed to match as a prefix/substring of the other.
function Same-Act([string]$a, [string]$b, [int]$min = 6) {
  $na = Norm $a; $nb = Norm $b
  if (-not $na -or -not $nb) { return $false }
  if ($na -eq $nb) { return $true }
  if ([math]::Min($na.Length, $nb.Length) -lt $min) { return $false }
  return ($na.Contains($nb) -or $nb.Contains($na))
}

# Merge one candidate show into the store. Returns 'new', 'matched' or 'skipped'.
# Same date + same venue matches loosely ("P.O.D." vs "P.O.D. & Buckcherry");
# same date at a differently named venue needs a longer, more distinctive name to match.
function Merge-Show($c, [string]$origin) {
  $existing = $null
  if ($byDate.ContainsKey($c.date)) {
    foreach ($r in $byDate[$c.date]) {
      $min = if ((Norm $r.venue) -eq (Norm $c.venue)) { 3 } else { 8 }
      if ((Same-Act $r.artist $c.artist $min) -or (Same-Act $r.artist $c.title $min)) { $existing = $r; break }
    }
  }
  Note-Venue $c
  if ($existing) {
    $existing.last_seen = $todayS
    $existing.tracked = '1'
    if (@($existing.sources -split '\|') -notcontains $origin) { $existing.sources = (@(@($existing.sources -split '\|') + $origin | Where-Object { $_ })) -join '|' }
    if (-not $existing.time -and $c.time) { $existing.time = $c.time }
    if ($c.price_min -ne '') { $existing.price_min = $c.price_min; $existing.price_max = $c.price_max }   # prices change, keep the latest
    if (-not $existing.support -and $c.support) { $existing.support = $c.support }
    if (-not $existing.genre -and $c.genre) { $existing.genre = $c.genre }
    if (-not $existing.source -and $c.url) { $existing.source = $c.url }
    if ($c.cancelled -and $existing.notes -notmatch '^CANCELLED') {
      $existing.notes = ('CANCELLED. ' + $existing.notes).Trim(); $cancelled.Add($existing)
    }
    return 'matched'
  }
  if ($c.cancelled) { return 'skipped' }
  $genre = $c.genre
  if (-not $genre -and $artistGenre.ContainsKey((Norm $c.artist))) { $genre = $artistGenre[(Norm $c.artist)] }
  $gb = GenreBucket ("$genre " + $(if (-not $genre) { "$($c.artist) $($c.title)" }))
  if (-not $genre -and $gb -ne 'Other') { $genre = $gb.ToLowerInvariant() }
  $row = [pscustomobject]@{
    date = $c.date; artist = $c.artist; support = $c.support; venue = $c.venue; genre = $genre
    notes = $c.notes; source = $c.url; tier = (TierOf $c.venue (GenreBucket $genre))
    time = $c.time; price_min = $c.price_min; price_max = $c.price_max
    first_seen = $todayS; last_seen = $todayS; origin = $origin; tracked = '1'; sources = $origin
  }
  $store.Add($row); $added.Add($row)
  if (-not $byDate.ContainsKey($row.date)) { $byDate[$row.date] = New-Object System.Collections.Generic.List[object] }
  $byDate[$row.date].Add($row)
  return 'new'
}

# Common filtering for every source. Returns $null when the show should be ignored.
# $more can carry time, pmin, pmax (numbers), lat, lon and address when the source has them.
function New-Candidate($src, [string]$date, [string]$title, [string[]]$performers, [string]$venue, [string]$city, [string]$url, [string]$genre, [string]$notes, [bool]$isCancelled, [hashtable]$more = @{}) {
  if (-not $date -or $date.Length -lt 10) { return $null }
  $time = if ($more.time) { TimeOf $more.time } else { TimeOf $date }
  $date = $date.Substring(0, 10)
  if (-not (IsDate $date)) { return $null }
  if ($date -lt $todayS -or $date -gt $horizonS) { return $null }
  if ($city -and $city -notmatch $cfg.allowed_localities) { return $null }
  $title = Clean $title
  $venue = VenueAlias $(if ($venue) { $venue } elseif ($src.venue) { $src.venue } else { '' })
  if (-not $venue) { return $null }
  if ($cfg.exclude_venue_pattern -and $venue -match $cfg.exclude_venue_pattern) { return $null }
  if ($knownVenues.ContainsKey((Norm $venue))) { $venue = $knownVenues[(Norm $venue)] }   # reuse the spelling already in the list
  $performers = @($performers | Where-Object { $_ } | ForEach-Object { Clean $_ })
  if ($performers.Count) {
    $artist = $performers[0]
    $support = ($performers | Select-Object -Skip 1 -First 8) -join ', '
  } else {
    $artist = $title -replace '\s+(@|at)\s+.*$', '' -replace '\s+tickets?$', '' -replace '\s*[-|]\s*calgary.*$', ''
    $support = ''
  }
  $artist = Clean $artist
  if (-not $artist) { return $null }
  # the non-music filter looks at the event itself, not at support acts, and never at rows you added by hand
  $hay = "$title $artist $support $genre"
  if (-not $src.manual -and "$title $artist $genre" -match $cfg.exclude_pattern) { return $null }
  foreach ($x in $ignore) { if ($artist -match $x -or $title -match $x) { return $null } }
  if ($src.music_only -and $hay -notmatch $cfg.music_pattern -and -not $knownVenues.ContainsKey((Norm $venue)) -and -not $performers.Count) { return $null }
  # keep the tour name or event title as a note when it says more than the artist name
  $extra = ''
  if ($title -and (Norm $title) -ne (Norm $artist)) {
    $extra = ($title -replace ('^' + [regex]::Escape($artist) + '\s*[:\-\u2013|]?\s*'), '' -replace '\s+@\s+.*$', '').Trim()
    if ((Norm $extra) -eq (Norm $venue) -or $extra.Length -gt 120) { $extra = '' }
  }
  $n = (@($extra, (Clean $notes)) | Where-Object { $_ }) -join '; '
  return [pscustomobject]@{
    date = $date; artist = $artist; title = $title; support = $support; venue = $venue
    genre = (Clean $genre).ToLowerInvariant(); notes = $n; url = ($url -replace '[?&]utm_[^#]*$', ''); cancelled = $isCancelled
    time = $time; price_min = (PriceStr $more.pmin); price_max = (PriceStr $(if ($null -ne $more.pmax) { $more.pmax } else { $more.pmin }))
    lat = [string]$more.lat; lon = [string]$more.lon; address = (Clean ([string]$more.address)); city = (Clean $city)
  }
}

# ---------------------------------------------------------------- source readers

# A page after the first one failed: keep what was read, and don't count this source toward "may be cancelled".
function Stopped($src, [int]$page, $err, [int]$kept) {
  $partial[$src.name] = $true
  Log ('{0}: stopped at page {1} ({2}); kept the {3} shows read before that.' -f $src.name, $page, ($err.Exception.Message -replace '\s+', ' '), $kept)
}

function Find-LdEvents($node, $acc) {
  if ($null -eq $node) { return }
  if ($node -is [System.Array]) { foreach ($n in $node) { Find-LdEvents $n $acc }; return }
  if ($node -isnot [System.Management.Automation.PSCustomObject]) { return }
  $t = $node.PSObject.Properties['@type']
  if ($t -and ((@($t.Value) -join ' ') -match '\b(MusicEvent|Event|Festival)\b')) { $acc.Add($node); return }
  foreach ($p in 'itemListElement', 'item', '@graph', 'subEvent') {
    $pp = $node.PSObject.Properties[$p]; if ($pp) { Find-LdEvents $pp.Value $acc }
  }
}

function Read-JsonLd($src) {
  $out = New-Object System.Collections.Generic.List[object]
  $seen = @{}
  $pages = if ($src.url -match '\{page\}') { [int]$src.max_pages } else { 1 }
  for ($p = 1; $p -le $pages; $p++) {
    $url = $src.url -replace '\{page\}', $p
    try { $html = Get-Text $url } catch { if ($p -eq 1) { throw }; Stopped $src $p $_ $out.Count; break }
    $events = New-Object System.Collections.Generic.List[object]
    foreach ($m in [regex]::Matches($html, '<script[^>]*application/ld\+json[^>]*>(.*?)</script>', 'Singleline')) {
      try { Find-LdEvents ($m.Groups[1].Value | ConvertFrom-Json) $events } catch { }
    }
    $fresh = 0
    foreach ($e in $events) {
      $key = "$($e.url)|$($e.startDate)|$($e.name)"
      if ($seen.ContainsKey($key)) { continue }
      $seen[$key] = $true; $fresh++
      $loc = @($e.location)[0]
      $city = ''
      if ($loc -and $loc.address) { $city = if ($loc.address -is [string]) { $loc.address } else { [string]$loc.address.addressLocality } }
      $perf = @(@($e.performer) + @($e.performers) | Where-Object { $_ } | ForEach-Object { if ($_ -is [string]) { $_ } else { $_.name } })
      $status = [string]$e.eventStatus
      $note = if ($status -match 'Postponed|Rescheduled') { 'Postponed or rescheduled' } else { '' }
      $more = @{ lat = [string]$loc.geo.latitude; lon = [string]$loc.geo.longitude
                 address = $(if ($loc.address -is [string]) { $loc.address } elseif ($loc.address) { [string]$loc.address.streetAddress } else { '' }) }
      $prices = @(@($e.offers) | Where-Object { $_ } | ForEach-Object { $_.price; $_.lowPrice; $_.highPrice } | ForEach-Object { Num $_ } | Where-Object { $_ -gt 0 })
      if ($prices.Count) { $more.pmin = ($prices | Measure-Object -Minimum).Minimum; $more.pmax = ($prices | Measure-Object -Maximum).Maximum }
      $c = New-Candidate $src ([string]$e.startDate) ([string]$e.name) $perf ([string]$loc.name) $city ([string]$e.url) '' $note ($status -match 'Cancelled') $more
      if ($c) { $out.Add($c) }
    }
    if ($fresh -eq 0) { break }
  }
  return $out
}

function Read-Showpass($src) {
  $out = New-Object System.Collections.Generic.List[object]
  $q = [uri]::EscapeDataString($src.query)
  for ($p = 1; $p -le [int]$src.max_pages; $p++) {
    $url = "https://www.showpass.com/api/public/events/?q=$q&starts_on__gte=$todayS&page_size=$($src.page_size)&page=$p"
    try { $j = (Get-Text $url) | ConvertFrom-Json } catch { if ($p -eq 1) { throw }; Stopped $src $p $_ $out.Count; break }
    foreach ($e in $j.results) {
      if ($e.is_published -eq $false) { continue }
      $date = if ($e.local_starts_on) { [string]$e.local_starts_on } else {
        $tz = [TimeZoneInfo]::FindSystemTimeZoneById('Mountain Standard Time')
        D ([TimeZoneInfo]::ConvertTimeFromUtc(([datetime]$e.starts_on).ToUniversalTime(), $tz)) 'yyyy-MM-dd'
      }
      $perf = @($e.artists | Where-Object { $_ } | ForEach-Object { if ($_ -is [string]) { $_ } else { $_.name } })
      $note = if ($e.sold_out -or $e.inventory_sold_out) { 'SOLD OUT' } else { '' }
      $more = @{ address = [string]$e.venue.street_name }
      if ($e.time_to_be_determined -or $e.date_time_to_be_determined) { $more.time = 'TBD' }
      $prices = @(@($e.ticket_types) | Where-Object { $_ } | ForEach-Object { Num $_.price } | Where-Object { $null -ne $_ })
      if ($prices.Count) {
        $paid = @($prices | Where-Object { $_ -gt 0 })
        if ($paid.Count) { $more.pmin = ($paid | Measure-Object -Minimum).Minimum; $more.pmax = ($paid | Measure-Object -Maximum).Maximum }
        else { $more.pmin = 0; $more.pmax = 0 }   # every ticket type is $0: a free (RSVP) show
      }
      $c = New-Candidate $src $date ([string]$e.name) $perf ([string]$e.venue.name) ([string]$e.venue.city) "https://www.showpass.com/$($e.slug)/" '' $note ([string]$e.status -match 'cancel') $more
      if ($c) { $out.Add($c) }
    }
    if (-not $j.next) { break }
  }
  return $out
}

function Read-Ticketmaster($src) {
  $out = New-Object System.Collections.Generic.List[object]
  if (-not $src.api_key) { Log 'Ticketmaster API skipped: add a free key from developer.ticketmaster.com to config.json to turn it on.'; return $out }
  # month-sized windows keep each query under the API's 1,000-result paging limit
  $from = $today
  $calls = 0
  while ((D $from 'yyyy-MM-dd') -le $horizonS) {
    $to = $from.AddMonths(1)
    $page = 0
    do {
      $url = 'https://app.ticketmaster.com/discovery/v2/events.json?apikey={0}&latlong={1}&radius={2}&unit=km&classificationName=music&size=200&sort=date,asc&page={3}&startDateTime={4}T00:00:00Z&endDateTime={5}T00:00:00Z' -f `
        $src.api_key, $src.latlong, $src.radius_km, $page, (D $from 'yyyy-MM-dd'), (D $to 'yyyy-MM-dd')
      $calls++
      try { $j = (Get-Text $url) | ConvertFrom-Json } catch { if ($calls -eq 1) { throw }; Stopped $src $calls $_ $out.Count; return $out }
      foreach ($e in @($j._embedded.events)) {
        if (-not $e) { continue }
        $v = @($e._embedded.venues)[0]
        $perf = @($e._embedded.attractions | ForEach-Object { $_.name })
        $cls = @($e.classifications)[0]
        $genre = @($cls.genre.name, $cls.subGenre.name) | Where-Object { $_ -and $_ -ne 'Undefined' -and $_ -ne 'Other' } | Select-Object -First 1
        $code = [string]$e.dates.status.code
        $note = if ($code -match 'postponed|rescheduled') { 'Postponed or rescheduled' } else { '' }
        $more = @{ lat = [string]$v.location.latitude; lon = [string]$v.location.longitude; address = [string]$v.address.line1 }
        $more.time = if ($e.dates.start.timeTBA -or $e.dates.start.noSpecificTime) { 'TBD' } else { [string]$e.dates.start.localTime }
        $pr = @($e.priceRanges)[0]
        if ($pr -and (Num $pr.min) -gt 0) { $more.pmin = Num $pr.min; $more.pmax = Num $pr.max }
        $c = New-Candidate $src ([string]$e.dates.start.localDate) ([string]$e.name) $perf ([string]$v.name) ([string]$v.city.name) ([string]$e.url) ([string]$genre) $note ($code -eq 'cancelled') $more
        if ($c) { $out.Add($c) }
      }
      $page++
    } while ($j.page -and $page -lt [int]$j.page.totalPages -and $page -lt 5)
    $from = $to
  }
  return $out
}

function Read-Manual {
  $out = New-Object System.Collections.Generic.List[object]
  $path = Join-Path $dataDir 'manual.tsv'
  if (-not (Test-Path $path)) { return $out }
  $src = [pscustomobject]@{ music_only = $false; venue = ''; manual = $true }
  $lines = Get-Content $path -Encoding UTF8 | Where-Object { $_ -and -not $_.StartsWith('#') }
  foreach ($r in ($lines | ConvertFrom-Csv -Delimiter "`t")) {
    if (-not (IsDate ([string]$r.date))) { Log "data\manual.tsv: skipped '$($r.artist)' because its date '$($r.date)' isn't a real date written as yyyy-MM-dd."; continue }
    $perf = @($r.artist) + @(($r.support -split ',\s*') | Where-Object { $_ })
    $c = New-Candidate $src $r.date $r.artist $perf $r.venue '' $r.source $r.genre $r.notes $false
    if ($c) { $c.support = Clean $r.support; $c.notes = Clean $r.notes; $out.Add($c) }
  }
  return $out
}

# ---------------------------------------------------------------- run sources

if ($RenderOnly) {
  Log 'Redrawing the guide from the saved list; the listing sites were not checked.'
} else {
  $sources = @($cfg.sources | Where-Object { $_.enabled })
  $sources += [pscustomobject]@{ name = 'Manual additions (data\manual.tsv)'; type = 'manual' }
  foreach ($src in $sources) {
    $st = [ordered]@{ found = 0; new = 0; matched = 0; error = '' }
    try {
      $cands = switch ($src.type) {
        'jsonld'       { Read-JsonLd $src }
        'showpass'     { Read-Showpass $src }
        'ticketmaster' { Read-Ticketmaster $src }
        'manual'       { Read-Manual }
        default        { throw "Unknown source type '$($src.type)'" }
      }
      foreach ($c in @($cands)) {
        if (-not $c) { continue }
        $st.found++
        switch (Merge-Show $c $src.name) { 'new' { $st.new++ } 'matched' { $st.matched++ } }
      }
      Log ('{0,-42} found {1,4}  new {2,4}  already listed {3,4}' -f $src.name, $st.found, $st.new, $st.matched)
    } catch {
      $st.error = ($_.Exception.Message -replace '\s+', ' ')
      Log ('{0,-42} ERROR: {1}' -f $src.name, $st.error)
    }
    $stats[$src.name] = $st
  }
}

# ---------------------------------------------------------------- write outputs

$final = @($store | Sort-Object date, artist)
if (-not $RenderOnly) { Log "Result: $($final.Count) upcoming shows ($($added.Count) new, $($cancelled.Count) newly cancelled, $($past.Count) past shows archived)." }

if ($DryRun) {
  Log 'Dry run: nothing written.'
  if ($added.Count) { Log 'Would add:'; $added | Sort-Object date | ForEach-Object { Log ("  {0}  {1}  @ {2}" -f $_.date, $_.artist, $_.venue) } }
  return
}

# A file that can't be written (open in Excel, say) is logged and skipped, and the rest of the run carries on.
function TryWrite([string]$what, [scriptblock]$do) {
  try { & $do; return $true }
  catch { Log ("Could not write {0}: {1}. Close it if it is open in another program; it is written again next run." -f $what, ($_.Exception.Message -replace '\s+', ' ')); return $false }
}

if (-not $RenderOnly) {
  # back up the store, keep the 10 most recent backups
  $backupDir = Join-Path $dataDir 'backup'
  New-Item -ItemType Directory -Force $backupDir | Out-Null
  if (Test-Path $storeCsv) { Copy-Item $storeCsv (Join-Path $backupDir ('shows-{0}.csv' -f (D (Get-Date) 'yyyyMMdd-HHmmss'))) }
  Get-ChildItem $backupDir -Filter 'shows-*.csv' | Sort-Object Name -Descending | Select-Object -Skip 10 | Remove-Item

  # past shows go to the archive first; if the archive can't be written they stay in the list until it can
  if ($past.Count) {
    $archive = Join-Path $dataDir 'past-shows.csv'
    $archived = TryWrite 'data\past-shows.csv' {
      # an archive from an older version gets the newer columns first, so the rows line up
      if ((Test-Path $archive) -and ((Get-Content $archive -TotalCount 1) -notmatch '"sources"')) {
        $old = @(Import-Csv $archive -Encoding UTF8)
        $old | Select-Object $cols | Export-Csv $archive -NoTypeInformation -Encoding UTF8
      }
      $past | Select-Object $cols | Export-Csv $archive -NoTypeInformation -Encoding UTF8 -Append
    }
    if (-not $archived) { $final = @(@($final) + $past | Sort-Object date, artist) }
  }
  # written to a temporary file and swapped in, so an interrupted run never leaves a half-written list
  $tmp = "$storeCsv.tmp"
  $final | Select-Object $cols | Export-Csv $tmp -NoTypeInformation -Encoding UTF8
  Move-Item -LiteralPath $tmp -Destination $storeCsv -Force
  [void](TryWrite 'data\venues.csv' { @($venueInfo.Values) | Sort-Object venue | Select-Object venue, lat, lon, address, city | Export-Csv $venueCsv -NoTypeInformation -Encoding UTF8 })

  # refresh history: a source counts as read when it had no error, didn't stop partway and found something
  foreach ($name in @($stats.Keys)) {
    $st = $stats[$name]
    $st.ok = if (-not $st.error -and -not $partial[$name] -and $st.found -gt 0) { '1' } else { '0' }
    if ($st.ok -eq '1') {
      if (-not $srcOkDates.ContainsKey($name)) { $srcOkDates[$name] = New-Object System.Collections.Generic.List[string] }
      if (-not $srcOkDates[$name].Contains($todayS)) { $srcOkDates[$name].Add($todayS) }
    }
  }
  $siteNames = @($stats.Keys | Where-Object { $_ -notlike 'Manual*' })
  $sitesRead = @($siteNames | Where-Object { $stats[$_].ok -eq '1' }).Count
  [void](TryWrite 'data\source-runs.csv' {
    @($stats.Keys) | ForEach-Object { [pscustomobject]@{ date = $todayS; source = $_; ok = $stats[$_].ok; found = $stats[$_].found } } |
      Export-Csv $srcRunsCsv -NoTypeInformation -Encoding UTF8 -Append
  })
  [void](TryWrite 'data\runs.csv' {
    [pscustomobject]@{ date = $todayS; time = (D (Get-Date) 'HH:mm'); ok = $(if ($sitesRead) { '1' } else { '0' }); added = $added.Count } |
      Export-Csv $runsCsv -NoTypeInformation -Encoding UTF8 -Append
  })
}

# The date the list was last checked: the latest refresh in which at least one listing site was read. A run
# where every site failed (no internet, blocked by a proxy) keeps the older date, so the page can warn.
$goodDates = @($srcOkDates.Keys | Where-Object { $_ -notlike 'Manual*' } | ForEach-Object { $srcOkDates[$_] })
if (-not $goodDates.Count -and (Test-Path $runsCsv)) { $goodDates = @(Import-Csv $runsCsv | Where-Object { $_.ok -eq '1' } | ForEach-Object { $_.date }) }
if (-not $goodDates.Count) { $goodDates = @($store | ForEach-Object { $_.last_seen } | Where-Object { IsDate $_ }) }
$runS = @($goodDates | Sort-Object)[-1]
if (-not (IsDate $runS)) { $runS = $todayS }
$runDate = [datetime]::ParseExact($runS, 'yyyy-MM-dd', $inv)

New-Item -ItemType Directory -Force $outDir, (Join-Path $outDir 'changes') | Out-Null

function PriceText($lo, $hi) {
  if ($null -eq $lo) { return '' }
  if ($lo -eq 0 -and (-not $hi)) { return 'Free' }
  if ($null -eq $hi -or $hi -eq $lo) { return '$' + (PriceStr $lo) }
  return '$' + (PriceStr $lo) + '-$' + (PriceStr $hi)
}

# "May be cancelled": a show in the next 90 days that every site which listed it has now read fine on two
# refresh days since anyone last saw it. Far-off shows aren't checked, since they can drop off a site's
# first pages without being cancelled; shows only from sources that are off or failing are never flagged.
$goneBy = D $today.AddDays(90) 'yyyy-MM-dd'
function IsGone($r) {
  if ($r.date -gt $goneBy -or $r.notes -match '^CANCELLED' -or -not $r.last_seen) { return $false }
  $srcs = @($r.sources -split '\|' | Where-Object { $_ })
  if (-not $srcs.Count) { return $false }
  foreach ($s in $srcs) {
    if (-not $srcOkDates.ContainsKey($s)) { return $false }
    if (@($srcOkDates[$s] | Where-Object { $_ -gt $r.last_seen }).Count -lt 2) { return $false }
  }
  return $true
}

$pageRows = @(foreach ($r in $final) {
  $np = NotesPrice $r.notes
  $pmin = if ($r.price_min -ne '') { Num $r.price_min } else { $np[0] }
  $pmax = if ($r.price_max -ne '') { Num $r.price_max } else { $np[1] }
  [pscustomobject]@{
    date = $r.date; artist = $r.artist; support = $r.support; venue = $r.venue; genre = $r.genre
    gb = (GenreBucket $r.genre); notes = $r.notes; source = $r.source; tier = $r.tier
    time = $(if ($r.time) { $r.time } else { NotesTime $r.notes }); pmin = $pmin; pmax = $pmax
    isNew = ($r.origin -ne 'baseline' -and $r.first_seen -ge $newSince)
    added = $(if ($r.origin -ne 'baseline') { $r.first_seen } else { '' })   # the page finds the latest refresh's additions from this
    gone  = (IsGone $r)
  }
})
$venueMap = [ordered]@{}
foreach ($v in @($final | ForEach-Object { $_.venue } | Sort-Object -Unique)) {
  $vi = $venueInfo[(Norm $v)]
  $venueMap[$v] = [ordered]@{ area = (AreaOf $v); addr = $(if ($vi) { [string]$vi.address } else { '' }) }
}
$srcLine = if ($RenderOnly) { '' } else { 'Sources this run: ' + (($stats.Keys | Where-Object { $stats[$_].found -or $stats[$_].error } | ForEach-Object {
  if ($stats[$_].error) { "$_ (not read)" } else { "$_ $($stats[$_].found)" } }) -join ', ') + '.' }
$info = [ordered]@{
  rangeLabel = (Label $today 'MMM d, yyyy') + ' to ' + (Label $today.AddDays([int]$cfg.window_days) 'MMM d, yyyy')
  checked    = Label $runDate 'MMM d, yyyy'
  run        = $runS
  stale_days = [int]$cfg.stale_after_days
  lede       = $cfg.page.lede
  caveats    = @(@($cfg.page.caveats) + $srcLine | Where-Object { $_ })
  watch      = @($watch)
  venues     = $venueMap
}
$json = ConvertTo-Json -InputObject $pageRows -Compress -Depth 3
$infoJson = ConvertTo-Json -InputObject $info -Compress -Depth 5
$tpl = [IO.File]::ReadAllText((Join-Path $root 'template.html'), [Text.Encoding]::UTF8)
$html = $tpl.Replace('__DATA__', $json).Replace('__INFO__', $infoJson)
$wroteHtml = TryWrite 'the guide page' { [IO.File]::WriteAllText($htmlPath, $html, (New-Object Text.UTF8Encoding $false)) }
if ($wroteHtml) { Log "Wrote $htmlPath" }

if (-not $RenderOnly) {
  # shows from the watchlist that this run found for the first time
  $watchHits = @($added | Where-Object { Watched $_ } | Sort-Object date)
  $hitLines = @($watchHits | ForEach-Object { '{0}: {1} at {2}' -f (Label ([datetime]::ParseExact($_.date, 'yyyy-MM-dd', $inv)) 'ddd MMM d'), $_.artist, $_.venue })

  # change report
  $rep = New-Object System.Collections.Generic.List[string]
  $rep.Add("Calgary concerts update, $todayS")
  $rep.Add("$($final.Count) upcoming shows. $($added.Count) new, $($cancelled.Count) newly cancelled, $($past.Count) past shows archived.")
  $rep.Add('')
  $rep.Add('FROM YOUR WATCHLIST')
  if ($hitLines.Count) { $hitLines | ForEach-Object { $rep.Add("  $_") } } else { $rep.Add('  none') }
  $rep.Add('')
  $rep.Add('NEW SHOWS')
  if ($added.Count) { $added | Sort-Object date | ForEach-Object { $rep.Add(('  {0}  {1}{2}  @ {3}' -f $_.date, $_.artist, $(if ($_.support) { " (with $($_.support))" }), $_.venue)) } } else { $rep.Add('  none') }
  $rep.Add('')
  $rep.Add('CANCELLED')
  if ($cancelled.Count) { $cancelled | ForEach-Object { $rep.Add(('  {0}  {1}  @ {2}' -f $_.date, $_.artist, $_.venue)) } } else { $rep.Add('  none') }
  $rep.Add('')
  $rep.Add('LOG')
  $log | ForEach-Object { $rep.Add("  $_") }
  $repPath = Join-Path $outDir ('changes\{0}.txt' -f (D (Get-Date) 'yyyy-MM-dd_HHmm'))
  if (TryWrite 'the change report' { [IO.File]::WriteAllLines($repPath, $rep, (New-Object Text.UTF8Encoding $false)) }) { Log "Change report: $repPath" }

  # notifications
  if ($siteNames.Count -and -not $sitesRead) {
    [void](Show-Toast 'Concert list not updated' @('None of the listing sites could be read this time.', 'It tries again next week; the guide keeps the list it has.') $htmlPath)
  } else {
    if ($hitLines.Count) {
      $title = if ($hitLines.Count -eq 1) { 'New show from your watchlist' } else { "$($hitLines.Count) new shows from your watchlist" }
      $body = @($hitLines[0])
      if ($hitLines.Count -eq 2) { $body += $hitLines[1] } elseif ($hitLines.Count -gt 2) { $body += "and $($hitLines.Count - 1) more. Open the guide to see them." }
      if (Show-Toast $title $body $htmlPath) { Log "Watchlist notification shown ($($hitLines.Count) shows)." }
    }
    if ($NotifyDone) {
      [void](Show-Toast 'Concert list updated' @("$($final.Count) upcoming shows; $($added.Count) found since this copy was built.", 'Click to open the guide.') $htmlPath)
    }
  }
}

# the list for Excel, separated the way Excel expects on this PC (a comma, or a semicolon in many countries)
$csvRows = for ($k = 0; $k -lt $final.Count; $k++) {
  $r = $final[$k]; $p = $pageRows[$k]
  [pscustomobject][ordered]@{
    date = $r.date; day = (DayName $r.date); time = $p.time; artist = $r.artist; support = $r.support
    venue = $r.venue; area = $venueMap[$r.venue].area; genre = $r.genre; price = (PriceText $p.pmin $p.pmax); notes = $r.notes
    source = $r.source; tier = $r.tier; first_seen = $r.first_seen; may_be_cancelled = $(if ($p.gone) { 'yes' } else { '' })
  }
}
$delim = if ($listSep -and $listSep.Length -eq 1) { [char]$listSep } else { ',' }
[void](TryWrite 'output\calgary-live-music.csv' { $csvRows | Export-Csv (Join-Path $outDir 'calgary-live-music.csv') -NoTypeInformation -Encoding UTF8 -Delimiter $delim })

if ($Open) { Open-Guide }
if ($lock) { try { $lock.ReleaseMutex() } catch { } }
