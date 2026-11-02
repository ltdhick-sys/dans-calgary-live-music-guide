# Dan's Calgary Live Music Guide

Rebuilds the Calgary concert list (next 12 months, every venue size) from public listing sites. It runs on plain Windows PowerShell 5.1 with nothing to install.

## Install on another PC

The only file you need is **`Dan's Calgary Live Music Guide.cmd`**, in the folder above this one. Everything else lives in this `source` folder, which the other PC doesn't need.

1. Copy that one file to the other PC by USB, OneDrive or email.
2. Double-click it.
   - If Windows says "Windows protected your PC", click **More info**, then **Run anyway**.
   - Or right-click the file, choose **Properties**, tick **Unblock**, then run it.
3. There are no questions. Setup takes a few seconds, the guide opens in its own window, and the list refreshes in the background.

After that, double-clicking the same file just opens the guide. So do the desktop and Start menu shortcuts. A newer copy of the file upgrades the program automatically.

What it sets up:
- The program in `%LOCALAPPDATA%\Programs\CalgaryConcerts`. No admin rights are needed.
- A weekly scheduled refresh, Mondays at 7:00. To pick another time, run `Install-WeeklyTask.ps1 -Day Sunday -Time 18:00` in the program folder.
- A **Dan's Calgary Live Music Guide** folder in the Start menu with these shortcuts:
  - the guide
  - Update concert list now
  - Edit watchlist
  - the README
  - Uninstall
- A desktop shortcut to the guide.
- An entry under Settings > Apps that uninstalls everything.

The installer already holds the show list as of the day it was built, so the guide works straight away.

**Upgrading:** after changing anything here, run `Build-Setup.ps1` and then double-click the new file on the other PC. It replaces the program and keeps that PC's show list, ignore list, manual rows and Ticketmaster key. It also removes the task and shortcuts left by the first version, which was called "Calgary Live Music Guide".
```
powershell -NoProfile -ExecutionPolicy Bypass -File .\Build-Setup.ps1
```
The build includes the Ticketmaster key from `config.json`. Add `-StripApiKey` to build a copy without it before giving the file to someone else.

**If it doesn't open:** if setup fails, the window stays open with the reason, and a log is saved to `%TEMP%\DansCalgaryLiveMusicGuide-setup.log`. On company PCs where IT restricts PowerShell, setup says so plainly. IT would need to allow it.

**Unattended install:** set these environment variables before running the file.
- `SETUP_INSTALLDIR`: install to a different folder.
- `SETUP_QUIET=1`: show messages only. Don't open the guide or start a refresh.
- `SETUP_NOSYSTEM=1`: skip the scheduled task, shortcuts and Apps entry.
- `SETUP_FORCE=1`: reinstall even if this version is already installed.
- `SETUP_WHEN=Sunday 18:00`: pick the weekly refresh time.

## Run it

- **Double-click `Run-Update.cmd`.** It takes about 3 minutes and opens the result in its own Edge window when it's done.
- Or from PowerShell:

  ```
  powershell -NoProfile -ExecutionPolicy Bypass -File .\Update-CalgaryConcerts.ps1 -Open
  ```

  Add `-DryRun` to see what would change without writing anything, or `-RenderOnly` to redraw the guide from
  the saved list without checking the sites (a couple of seconds; used after editing the watchlist).

## Using the guide

- **Star (☆) a show** to keep it under **My shows**. Stars are saved in this PC's browser, not shared between computers.
- **Calendar button** on a show opens a small menu:
  - **Add to My calendar** puts the show on the guide's own calendar. The **My calendar** view (next to List and Calendar) shows a month grid of your plans with the shows listed underneath, warns if two land on the same day, and flags any that are cancelled or drop off the list. It's kept on this PC.
  - **Google Calendar** opens a filled-in event.
  - **Outlook** saves a calendar file; open it to add the show. "Add all to Outlook" in My calendar (and "Add these to Outlook" under My shows) saves them all in one file. On the online copy at claude.ai, which can't save files, Outlook opens on the web instead.
- Once you scroll past the big title, a slim title bar stays at the top with the venue-size colour key; click a colour to show only that size of venue.
- **Spotify and YouTube buttons** search for the act so you can hear them first. The pin next to a venue opens Google Maps.
- **Not interested (eye button)** hides every show by that act on this PC. The "hidden · Show" button brings them back.
- **When** picks Tonight, This weekend, Next 7 or 30 days, or your own dates. **Areas** filters by neighbourhood or city quadrant.
- **Free** lists shows marked free or no cover. Prices and times appear where the listing gives them.
- **Sort** by date, venue, newest additions or price. **Calendar** shows a month grid; click a day to list its shows.
- Search ignores accents and forgives small typos ("calahan" finds Callahan).
- Banners at the top warn when the list hasn't been refreshed for 10 days, when a refresh found shows from your watchlist, and when a starred show is cancelled or may be.

## Watchlist

`data\watchlist.txt` holds artists and venues you care about, one per line. When a refresh finds a new show that
mentions one, Windows shows a notification (click it to open the guide) and the show is marked **Watching**.
Use **Edit watchlist** in the Start menu: it opens the file in Notepad and redraws the guide when you're done.

## Neighbourhoods and "may be cancelled"

- Venue locations come from the listing sites and are kept in `data\venues.csv`. Each venue is put in the nearest
  neighbourhood from `areas` in `config.json`, otherwise in its city quadrant (NW, NE, SW, SE) or "Outside Calgary".
  Venues the sites don't locate can be placed by hand with `venue_areas` in `config.json`.
- `data\runs.csv` records each refresh and `data\source-runs.csv` how each listing site did. Each show remembers which
  sites listed it. A show in the next 90 days is marked **May be cancelled** once every site that listed it has been
  read successfully on two later refresh days without it. A site that is switched off, has no key, fails or returns
  nothing doesn't count, so it can't cause false flags.

## What you get

| File | What it is |
|---|---|
| `output\calgary-live-music.html` | The filterable guide. It opens in any browser and works offline, apart from the fonts. |
| `output\calgary-live-music.csv` | The same list for Excel, with time, area and price columns. |
| `output\changes\<date_time>.txt` | Each run's new shows, cancellations, per-source counts and any errors. |

## Run it weekly

```
powershell -NoProfile -ExecutionPolicy Bypass -File .\Install-WeeklyTask.ps1
```

This sets up a Windows scheduled task for Mondays at 7:00.
- To pick another day or time, add `-Day Sunday -Time 18:00`.
- To remove the task, add `-Remove`.

The task runs as you, only while you're signed in, so the `V:` drive is available. If the PC is off at the scheduled time, it runs at the next chance.

## Where the shows come from

| Source | How it's read | Notes |
|---|---|---|
| Songkick (Calgary metro) | Structured event data on the listing pages | Music only. This is the backbone source. |
| Eventbrite (Calgary concerts) | Structured event data | Small rooms, churches and house concerts. Results vary a little from run to run. |
| Showpass | Public API | Filtered to music. |
| Live Nation venue pages | Structured event data | About 18 Calgary rooms: Saddledome, Jubilee, Grey Eagle, Jack Singer, MacEwan, Dickens, the Arrowhead and others. |
| Ticketmaster Discovery API | Official API | **Off until you add a free key** (see below). It's the best source for big venues and the only one that separates music from comedy. |
| `data\manual.tsv` | Your hand-entered rows | For shows no source carries. |

### Add the Ticketmaster key (recommended, about 2 minutes)
1. Sign up at <https://developer.ticketmaster.com>. It's free, with 5,000 calls a day; a run uses about 15.
2. Copy your **Consumer Key**.
3. Save it as the only line of `data\ticketmaster.key` (or paste it into `"api_key": ""` on the Ticketmaster line of
   `config.json`). The separate file keeps the key out of copies you share, such as the GitHub project; the
   installer you build with `Build-Setup.ps1` still includes it unless you add `-StripApiKey`.

## How the list is kept

- `data\shows.csv` is the master list. It started with the 577 shows researched on Oct 28, 2026.
- Each run:
  1. Adds shows it hasn't seen before.
  2. Fills in missing details on shows it has.
  3. Marks shows a source reports as cancelled.
  4. Moves shows whose date has passed to `data\past-shows.csv`.
- Shows stay on the list until their date passes, even if no source mentions them again. That's how shows from venues the script can't read stay on the list.
- The page hides any show dated before today, so past shows disappear the next day even between weekly updates. If the page is left open past midnight, it refreshes itself.
- Duplicates are matched on date plus artist. Small naming differences are handled: "&" vs "and", accents, and "P.O.D." vs "P.O.D. & Buckcherry".
- Shows added in the last 7 days get a **New** badge and a "New this week" filter on the page.
- "Added in last refresh" shows only what the most recent refresh found. If that refresh found nothing new,
  the button names the date of the last refresh that did (for example "Added Oct 28").
- A backup of `shows.csv` is saved in `data\backup\` before every run, and the last 10 are kept.

## Your own settings

`config.json` holds the defaults and is replaced when you install a newer build. Put your own changes in
**`data\config.local.json`** instead: same layout, only the settings you change. Each top-level setting there
replaces the one in `config.json` (lists such as `venue_aliases` or `sources` are replaced whole, so copy the full
list and edit it). For example, to change how long a list can go without a refresh before the page warns:

```
{ "stale_after_days": 14 }
```

An upgrade keeps this file, your watchlist, ignore list, manual rows, show list, the Ticketmaster key already on the
PC and the weekly schedule you chose.

## Keeping it tidy

Check the change report every so often.

- **A non-music show got in** (usually a comedian on a Live Nation page): add the name to `data\ignore.txt`. It's removed on the next run and never added again.
- **A show is missing** (a Palace, Blues Can or Calgary Phil date, or a festival lineup): add a tab-separated row to `data\manual.tsv`.
- **A venue shows up under two names:** add a pattern to `venue_aliases` (in `data\config.local.json`, see above).
- **A kind of event keeps slipping through:** add a word to `exclude_pattern` (titles) or `exclude_venue_pattern` (venues).
  The title filter doesn't look at support acts, and never applies to rows you add in `manual.tsv`.
- Lines in `ignore.txt` are patterns: a plain name works, but a line that would match every show (like `.*`) is skipped.
- If a run can't write a file because it's open in Excel, it says so in the change report and carries on.

### Adding another source
Any page with schema.org event data (most ticketing and venue pages) can be added to `sources` in `config.json` as `"type": "jsonld"`. Use `{page}` in the URL if it's paginated.

For a Live Nation venue:
1. Find the venue on <https://www.livenation.com>.
2. Copy its `/venue/<id>/<name>-events` URL.
3. Add a line like the existing ones.

## Limits

- These venue sites load their calendars by script, so the script doesn't read them directly: the Palace, Grey Eagle's own site, Werklund Centre, the Blues Can, the Calgary Phil, MRU tickets. Their shows still arrive when Songkick, Live Nation or Eventbrite carry them.
- Without the Ticketmaster key, Live Nation pages can't tell comedy from music, so `ignore.txt` does that job.
- The page on claude.ai doesn't update itself. To refresh it, ask Claude to republish `output\calgary-live-music.html`, or share the HTML file directly (OneDrive, SharePoint or email).
