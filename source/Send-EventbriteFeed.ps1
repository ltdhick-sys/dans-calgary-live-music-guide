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

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File .\Send-EventbriteFeed.ps1
#>
param([string]$Repo)

$ErrorActionPreference = 'Stop'
$root    = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$project = Split-Path $root -Parent

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { throw 'The GitHub command-line tool (gh) is not installed on this PC.' }
cmd /c 'gh auth status >nul 2>nul'
if ($LASTEXITCODE -ne 0) { throw 'GitHub is not signed in on this PC. Run: gh auth login --web' }
if (-not $Repo) {
  $url = (cmd /c "git -C ""$project"" remote get-url origin 2>nul")
  $Repo = ([string]$url).Trim() -replace '^https://github\.com/', '' -replace '\.git$', ''
  if ($Repo -notmatch '^[\w.-]+/[\w.-]+$') { throw 'Could not tell which GitHub project to send to. Run again with -Repo owner/name.' }
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
}
