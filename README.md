# NinjaTrader Market Replay bulk downloader

NinjaTrader's Historical Data window lets you request **one** Market Replay date at a time.
Downloading a few months for a few symbols means hundreds of pointing, clicking and waiting
cycles.

This script does it for you. Give it a date range and a symbol, and it downloads every day in
turn, skipping anything already on disk, and tells you exactly which days the server does not
have.

> [!IMPORTANT]
> **NinjaTrader must be open, running and connected to your data feed while this runs.**
>
> Replay data is delivered over NinjaTrader's own connection to its servers. It is not an HTTP
> download and there is no API key or URL to use instead, so the script cannot fetch anything
> with NinjaTrader closed — it will exit with a message telling you to start it.
>
> You do **not** have to open the Historical Data window yourself; the script does that. It is
> NinjaTrader itself that has to be running.

```
.\Invoke-ReplayDownload.ps1 -AutoContract -Roots NQ,ES -From 2026-09-01 -To 2026-11-20 -WeekdaysOnly
```

```
BATCH MARKET REPLAY DOWNLOAD
  instruments : NQ, ES
  dates       : 2026-09-01 .. 2026-11-20  (59 days)
  requests    : 118   already on disk: 0
  window      : Historical Data ready (pid 33720)
[1/118] NQ 09-26 2026-09-01
        OK  154.4 MB  (2.9s)
[2/118] NQ 09-26 2026-09-02
        OK  131.1 MB  (2.4s)
...
```

About **3 seconds per day-file** on a normal connection, so a three-month, two-symbol backfill
takes a few minutes rather than an afternoon.

---

## Requirements

| | |
|---|---|
| **Windows** | The script drives the NinjaTrader UI through Windows UI Automation. |
| **NinjaTrader 8** | Running **and connected**. Replay data comes over NinjaTrader's own connection, so it cannot be fetched with NinjaTrader closed. |
| **PowerShell** | 5.1 (built into Windows) or PowerShell 7. |

You do **not** need to open the Historical Data window. The script opens `Tools → Historical Data`
itself and leaves it as it found it.

You do **not** need to sign in through the script, configure an API key, or install anything.
It uses the NinjaTrader session you already have.

---

## Using your computer while it runs

The mouse is free. The **keyboard is not reliably yours** while a run is in progress.

Each item touches NinjaTrader's controls, and that moves **keyboard focus** to the window — even
though the script never asks for focus and does not activate NinjaTrader. Measured during a live
run, sampling the focused window:

```
before run        -> your window
on click          -> NinjaTrader
after restore     -> your window
```

So focus is taken and then handed back, once per item. If you are typing at the moment it happens,
those keystrokes can go to NinjaTrader instead of where you meant.

The script restores the window you had focused before each item, and this usually works, but
**treat it as best effort rather than a guarantee**. Over repeated runs it sometimes does not stick,
leaving focus in NinjaTrader until you click back. Windows restricts background processes from
taking the foreground, and the window's activation is asynchronous, so the restore races it.

Practical guidance: for a long backfill, expect that typing may be interrupted and click back to
your window if it is. For a short run, just wait until it finishes.

If you would rather the script never touched focus at all, pass `-NoFocusRestore`. Focus will then
stay in NinjaTrader between items, which is predictable rather than intermittent.

**Do not edit the Historical Data window while a run is in progress** — don't change the instrument
or date boxes, click Download, or close the window. The script re-reads those fields immediately
before clicking, so if they change underneath it, it refuses to click rather than requesting the
wrong day:

```
REFUSING TO CLICK: inst='RTY 12-26' date='11/19/2026'
```

That date is then reported `FAIL`. Nothing is downloaded wrongly, but the run does stop on that
item. Re-run the same command and it continues from where it stopped.

Interacting with NinjaTrader's charts or DOMs is harmless; it is only that one window that matters.

---

## Install

Put the script somewhere permanent — not inside your NinjaTrader data folder, since almost
everything in there is disposable.

```powershell
git clone https://github.com/lecrackzor/nt-replay-downloader.git
cd nt-replay-downloader
```

Or download the folder and extract it anywhere, for example `C:\Tools\nt-replay-downloader`.

There is nothing to install, no module to import and no dependency to restore. `Invoke-ReplayDownload.ps1`
is a single self-contained script.

## Quick start

**Start NinjaTrader and connect to your data feed first.** Everything below needs it running.

```powershell
# 1. Find out what you would get, without downloading anything.
.\Invoke-ReplayDownload.ps1 -Plan -AutoContract -Roots NQ -From 2026-11-01 -To 2026-11-20 -WeekdaysOnly

# 2. Download it.
.\Invoke-ReplayDownload.ps1 -AutoContract -Roots NQ -From 2026-11-01 -To 2026-11-20 -WeekdaysOnly
```

Run `-Plan` first whenever you are about to fetch a large range. It never touches NinjaTrader,
takes well under a second, and prints coverage per contract plus the dates you are missing.
`-Plan` is the one thing that works with NinjaTrader closed, because it only reads your disk.

### Windows may block the script the first time

```powershell
Unblock-File .\Invoke-ReplayDownload.ps1
```

If your execution policy is restrictive, run the script in a session started with:

```powershell
powershell -ExecutionPolicy Bypass -File .\Invoke-ReplayDownload.ps1
```

---

## Two ways to name what you want

### Automatic contract resolution (recommended)

Use `-AutoContract` with a symbol **root**. The script maps each date to the quarterly contract
that was front-month on that date, because that is the only contract NinjaTrader will serve for
that date.

```powershell
.\Invoke-ReplayDownload.ps1 -AutoContract -Roots NQ,ES,RTY -From 2026-09-01 -To 2026-11-20 -WeekdaysOnly
```

This is built for CME quarterly index futures — NQ, ES, RTY, YM and their micros. Their contracts
expire on the third Friday of March, June, September and December and roll about eight days before
expiry. Other products may use a different expiry or roll convention, so use a fixed `-Instrument`
for those.

### A fixed contract

Use `-Instrument` with the name exactly as NinjaTrader displays it.

```powershell
.\Invoke-ReplayDownload.ps1 -Instrument 'NQ 12-26' -Dates 2026-11-12,2026-11-13
```

Not sure of the spelling? Open `Tools → Historical Data` in NinjaTrader and read the instrument
box, or run `.\Invoke-ReplayDownload.ps1 -DumpControls` to print what the script currently sees.

> **Getting this wrong looks like missing data.** If you ask for a date that falls outside the
> life of the contract you named, NinjaTrader answers with the same *"no market replay data"*
> message it uses for a genuinely absent day. It is easy to conclude the server has no archive
> when the real problem is the contract name. Prefer `-AutoContract`.

---

## Common recipes

```powershell
# Three months, three symbols, weekdays only
.\Invoke-ReplayDownload.ps1 -AutoContract -Roots NQ,ES,RTY -From 2026-09-01 -To 2026-11-20 -WeekdaysOnly

# What do I already have? (fast, no NinjaTrader needed)
.\Invoke-ReplayDownload.ps1 -Plan -AutoContract -Roots NQ,ES,RTY -From 2026-08-25 -To 2026-11-23 -WeekdaysOnly

# One day across several fixed contracts
.\Invoke-ReplayDownload.ps1 -Instrument 'NQ 12-26','ES 12-26','RTY 12-26' -Dates 2026-11-20

# Re-fetch only what is missing: just re-run the same command, present files are skipped
.\Invoke-ReplayDownload.ps1 -AutoContract -Roots NQ -From 2026-09-01 -To 2026-11-20 -WeekdaysOnly

# Show timing for each step (useful when reporting a slow run)
.\Invoke-ReplayDownload.ps1 -AutoContract -Roots NQ -Dates 2026-11-20 -Trace

# Close the Historical Data window when finished
.\Invoke-ReplayDownload.ps1 -AutoContract -Roots NQ -Dates 2026-11-20 -CloseWindow
```

Re-running a range is cheap and safe. Files already on disk are skipped, so the script can be
interrupted at any point and started again to continue where it left off.

---

## What you end up with

Files land where NinjaTrader puts them:

```
%USERPROFILE%\Documents\NinjaTrader 8\db\replay\<CONTRACT>\<YYYYMMDD>.nrd
```

Use `-NinjaTraderHome` if your NinjaTrader data folder is somewhere else.

Run logs go to `logs\` next to the script. **Keep them.** They record the finished size of every
file, and the script uses them to tell a complete file from one that was cut short — see
[Truncated files](#truncated-files) below.

---

## How much history is available

Roughly **90 days**, as a rolling window. Measured on 2026-11-23, the oldest retrievable session
was exactly 90 calendar days back, and the cutoff was identical for every symbol tested. Asking
for one day earlier returns *"no market replay data"*.

Because it is a rolling window, the cutoff moves every day. Practically: fetch what you need
before it ages out.

---

## Outcomes

Each date ends in one of four states:

| Status | Meaning |
|---|---|
| `OK` | Downloaded. The size is printed. |
| `SKIP` | Already on disk and verified complete. |
| `NO-DATA` | The server does not hold that date. Collected into one list at the end. |
| `FAIL` | The script could not complete it, with a reason. |

`NO-DATA` is normal, not an error. Exchange holidays, weekends you did not filter out, and dates
older than the retention window all produce it. A run of purely `NO-DATA` results exits `0`.

Exit codes:

| Code | Meaning |
|---|---|
| `0` | Every requested date is `OK`, `SKIP` or `NO-DATA`. |
| `1` | At least one date `FAIL`ed. The reason is printed. |
| `2` | Bad arguments, for example no dates or no instrument. |
| `3` | NinjaTrader is not running, or the Historical Data window could not be opened. |

---

## Troubleshooting

### `NinjaTrader is not running`

Replay data is delivered over NinjaTrader's own connection. Start NinjaTrader, connect to your
data feed, and re-run. There is no way around this — it is not an HTTP download.

### `Could not open Tools > Historical Data`

Open the window manually (`Tools → Historical Data`), then run:

```powershell
.\Invoke-ReplayDownload.ps1 -DumpControls
```

If that reports any control as `MISSING`, NinjaTrader has renamed a UI element, most likely after
a version upgrade. Please open an issue with your NinjaTrader version.

### Someone used the keyboard and a date came back `FAIL`

Look at the reason printed against that date. If it says `REFUSING TO CLICK: inst='...'
date='...'`, the window's fields were changed underneath the script — usually because someone typed
in the instrument or date box while it was running.

That is the guard working correctly: the script compares what it wrote against what the window
actually holds, and declines to click rather than requesting a day you did not ask for. Re-run the
same command to pick up where it stopped.

Ordinary typing elsewhere does not corrupt a run. It may land in NinjaTrader instead of your
document, because focus moves there briefly for each item — see
[Using your computer while it runs](#using-your-computer-while-it-runs).

### `click did not register -- restoring the window and retrying once`

The script clicked Download and nothing happened. It restores the window and retries once, then
moves on. This is expected occasionally.

If it happens repeatedly, **close the Historical Data window**. The next run reopens it, which
clears the stuck state. NinjaTrader does not need to be restarted.

### Truncated files

NinjaTrader writes a replay file gradually, so a run killed mid-transfer leaves a short file that
looks finished to a naive check. Silently keeping one would leave a corrupt day in your dataset.

The script compares each file against the size recorded in its log when that download completed,
and re-downloads anything that does not match:

```
TRUNCATED (30.0 MB, log says 99.6 MB) -- re-downloading
```

This is why the `logs\` folder matters. If you delete the logs, the size reference is lost and a
short file can only be spotted by an absolute minimum size — far weaker. Re-download the affected
range to rebuild both the files and the log entries.

### Everything is slow

Run with `-Trace` to see where the time goes. A typical day-file takes 2–4 seconds. If it takes
much longer, NinjaTrader is likely busy (charts rebuilding, a connection retry) rather than the
script being inefficient.

### The script opened or moved my windows

It should not. The window is opened only if it is absent, and it is never minimised, maximised or
moved. If you see it change size, that is a bug worth reporting.

---

## How it works

The short version: it fills in the instrument and date boxes, clicks Download, and waits for the
status line at the bottom of the window to go blank before moving on. Ground truth is always the
file on disk.

There is a fair amount of measured detail behind that — why the date box has to be written through
one element and not the other, why the refusal dialog is invisible to a normal window scan, why the
click has to be verified rather than trusted. Those notes are in
[docs/HOW-IT-WORKS.md](docs/HOW-IT-WORKS.md), mainly so the reasoning survives the next time
someone touches this.

---

## Scope and caveats

- **This automates a UI.** It is not an API client. A NinjaTrader update can rename a control and
  break it; `-DumpControls` exists to diagnose that quickly.
- **It does not bypass anything.** It uses your NinjaTrader session and requests data the same way
  the window does, at the same rate a person would click, plus the few seconds a download takes.
- **Replay files are NinjaTrader's format.** This downloads them; it does not parse them.

## License

MIT. See [LICENSE](LICENSE).
