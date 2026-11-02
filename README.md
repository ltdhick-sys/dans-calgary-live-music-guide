# Dan's Calgary Live Music Guide

Every announced music show in Calgary for the next 12 months, from the Saddledome and the Jubilee down to the
Palomino, the King Eddy and the Blues Can, in one page you can filter, star and add to your calendar. A small
Windows program refreshes the list every week from public listing sites (Songkick, Eventbrite, Showpass,
Live Nation venue pages and, with a free key, Ticketmaster).

## Install (Windows 10 or 11)

1. Download **Dan's Calgary Live Music Guide.zip** from the latest release on this page (under **Releases**).
2. Right-click it, choose **Extract All**, then double-click `Dan's Calgary Live Music Guide.cmd`.
   If Windows says "Windows protected your PC", click **More info**, then **Run anyway**.
3. There are no questions. The guide opens in its own window and the list refreshes in the background.

It installs for your user only (no admin rights), adds Start menu and desktop shortcuts and a weekly refresh,
and can be removed from Settings > Apps. The release build has no Ticketmaster key; to add your own free key,
see `source/README.md`.

## What the guide does

- Filter by date (tonight, this weekend, your own dates), venue, area, month, genre, venue size and price.
- Star shows, keep plans on **My calendar**, and send any show to Google Calendar or Outlook.
- A watchlist of artists and venues raises a Windows notification when a new show is announced.
- Listen first on Spotify or YouTube, open the venue in Google Maps, hide acts you're not interested in.
- Copy any list into Excel.

## For developers

Everything is in [`source`](source): plain Windows PowerShell 5.1 and one HTML template, nothing to install.
[`source/README.md`](source/README.md) covers how the list is built, the settings, and how to make the
single-file installer (`source/Build-Setup.ps1`; add `-StripApiKey` for a copy to share).
