# How it works, and why it is built this way

This file records the measurements behind the design. Most of them are not obvious, and each one
silently broke an earlier version. If you are changing the script, read this first — several
apparently reasonable "simplifications" reintroduce a bug.

---

## Why UI automation instead of the network

Replay data arrives over NinjaTrader's HDS connection: a TLS stream to `hds-us-nt-*.ninjatrader.com`
on port 31655 with proprietary framing. There is no URL to fetch and no public endpoint to call.

Reverse-engineering that protocol would mean defeating TLS, would break on any server change, and
would be a terms-of-service problem. Driving the vendor's own dialog instead uses the vendor
client, the vendor's authentication and the vendor's own rate. That is the supported surface.

**NinjaTrader must therefore be running and connected.** Only the GUI is automated away.

---

## Why the loop exists at all

NinjaTrader's Market Replay section has an instrument selector, a **single** date picker, and a
Download button. There is no date-range control for replay. The Export section does have Start and
End dates, but Export writes a local file — it does not fetch replay data.

NinjaTrader genuinely has no bulk mode. The script supplies the loop.

---

## The date control has two elements, and only one matters

The replay date is an `Edit` with AutomationId `HistoricalDataWindowMarketReplayDateSelector`.
Inside it sits a `PART_InputTextBox`.

Writing to the **inner** box changes the visible text but not the value NinjaTrader binds, so
Download acts on the previous date. Observed directly: the inner box displayed `11/02/2026` while
the outer element still read `11/18/2026`.

Set the **outer** element, in `M/d/yyyy` format. Not ISO.

---

## Element references go stale

After a dialog has been shown and dismissed, NinjaTrader rebuilds the control tree, so any element
reference captured earlier is dead. A stale reference does not throw — it reports values from
before. An early version read back `Y` when `NQ 06-26` had been requested.

References are therefore re-acquired at the start of every item, and both fields are **read back
and verified** before the click. If either does not match what was requested, the click is refused
rather than sent with the wrong values.

---

## The refusal dialog is not a top-level window

A refused request raises a modal reading *"There is no market replay data available for the
instrument/date."*

It is a nested `Window` with `AutomationId='NTMessageBox'` living **inside the Historical Data
window** — not a new top-level desktop window. Enumerating top-level windows finds nothing; while
one was on screen, 17 unrelated top-level windows were visible instead.

Two consequences:

1. **Find it by AutomationId, scoped to the Historical Data window.** Not by scanning every
   NinjaTrader window for text.
2. **Do not filter by window title.** An earlier attempt skipped windows whose title matched
   "Control Center", on the guess that the dialog was hosted there. That guess was wrong.

Without detection the loop blocks forever with no progress and no error.

### The cost of getting this wrong

The original implementation scanned the descendant tree of every NinjaTrader window for matching
text, on every 200 ms poll:

| operation | cost |
|---|---|
| `FindFirst(Descendants, AutomationId='NTMessageBox')` on the window | 2 ms |
| reading its text children | 13 ms |
| whole-tree text scan across all 8 open windows | 1260 ms |

With five charts and DOMs open, one scan cost about 1.26 s. Since the completion check required
several consecutive stable polls, the script spent roughly **7 seconds per file after the transfer
had already finished**, on a download that only takes 2–3 seconds. That is the entire "it is slow
to start the next download" symptom.

Two fixes: scope the lookup by AutomationId, and only probe at all while the target file is still
absent — a refusal never creates a file, so the success path needs no probe.

---

## The status line is the completion signal

`txtMessage` at the bottom of the window reports download state:

| moment | `txtMessage` |
|---|---|
| idle, before the click | `''` |
| on click | `Loading data ... NQ 12-26 Globex` |
| transfer finished | `''` (blank again) |

It is a `TextBlock`, so `GetCurrentPattern(ValuePattern)` throws. The text is in the element's
`Name`. A **cached** handle answers that in about 0.1 ms; a fresh `FindFirst` costs about 115 ms,
so the handle is kept for the whole item.

Completion is:

```
message blank  AND  size >= 1 MB  AND  size unchanged for 400 ms
```

The size condition is a backstop in case the message blanks a moment before the write lands.

Two facts force the extra conditions:

1. NinjaTrader writes the file **incrementally**. "File exists" alone fires on a partial file —
   89 MB was observed present at 1.2 s while the transfer was still running.
2. NinjaTrader stamps the finished file with the **data's own date**, never the download time.
   `20261112.nrd` carries an mtime of 2026-11-13, and so on. Any test of the form
   "mtime newer than the click" can never be true.

---

## Concurrent use of the computer, and what actually interferes

UI Automation addresses controls directly rather than injecting input at the cursor position, so the
*download* does not need the window focused. The *keyboard* is a different matter.

Measured 2026-11-23 during live downloads:

| interference | result |
|---|---|
| 156 cursor movements across the screen during one download | `OK 99.6 MB (3.1s)`, unaffected |
| Focus pulled away from the Historical Data window 11 times | `OK 99.6 MB (2.2s)`, unaffected |
| Date field rewritten 12 times mid-run, simulating a user editing it | `REFUSING TO CLICK: inst='RTY 12-26' date='11/19/2026'` → `FAIL`, no file created |

### Keyboard focus moves to NinjaTrader, once per item

Writing an Infragistics `XamMaskedEditor` (the date box), invoking the Download button, and
expanding the replay section each move **keyboard focus** to that control. Sampling the focused
window during a run shows it plainly:

```
before run     -> your window
on click       -> NinjaTrader
after restore  -> your window
```

So anything typed at that moment can go into NinjaTrader instead of where it was intended.

`Restore-Foreground` puts the previously focused window back, and it is called after the click and
again at the end of each item. **It is best effort, not a guarantee.** Over repeated runs the
foreground sometimes does not come back, because Windows restricts background processes from taking
the foreground and the target window's activation is asynchronous — the restore races it. One
attempt with a `Start-Sleep` before it helped in some runs and not others.

This is why the README tells users to expect interrupted typing rather than promising otherwise, and
why `-NoFocusRestore` exists for a predictable alternative. Do not claim this is solved without
re-measuring.

### Closing the window needs WM_CLOSE

`WindowPattern.Close()` on the Historical Data window returns **without error and does not close
it** — verified by re-enumerating the window list afterwards. `PostMessage(hwnd, WM_CLOSE)` does
work. `-CloseWindow` therefore uses WM_CLOSE and then polls until the window is actually gone before
reporting success; an earlier version printed "closed" while the window was still open.

The read-back check that produces the `REFUSING TO CLICK` guard is not incidental. The script writes
the instrument and date, then re-reads both and compares them against what was requested before
clicking. Without that check the run would have silently downloaded **2026-11-19** instead of
2026-11-20, which is far worse than a `FAIL`, because it would look like success.

Re-running the same command after such a `FAIL` is the correct recovery: files already present are
skipped, so it resumes rather than repeating work.

---

## The click must be verified, not trusted

`InvokePattern` against a **minimised** WPF window can be a silent no-op: no exception, the button
still reports enabled, no download starts, the status line stays blank, and the run then sits until
timeout.

This was originally believed to work, because one test fetched 214.5 MB in 4.8 s with the window
minimised. That test was invalid — the window had already been painted visible before being
minimised. Two later runs hung outright with the instrument and date correctly armed.

A silent no-op is indistinguishable from a slow server, so the click is now proven: the script
waits `-ClickGraceSec` (default 6) for `txtMessage` to read `Loading data` **or** for bytes to
arrive. If neither happens it restores the window and retries once.

The window is never minimised or maximised by the script. It is opened if absent and otherwise
left exactly as found.

### When NinjaTrader itself gets stuck

After heavy use the replay subsystem can wedge: window normal, fields reading back correctly,
button enabled, no modal, NinjaTrader responsive — and every click mechanism (InvokePattern,
focus+invoke, foreground+invoke) does nothing, with no error.

**Closing the Historical Data window clears it.** The next run reopens it. NinjaTrader itself does
not need restarting. Observed on 2026-11-23: three dates refused every click for several minutes,
then all three downloaded in 2.1–3.5 s immediately after the window was closed and reopened.

---

## Contracts: why `-AutoContract` exists

NinjaTrader will not serve a date outside a contract's life. Tested in both directions:

| instrument | date | result |
|---|---|---|
| `NQ 12-26` | 2026-06-01 | *"no market replay data"* |
| `NQ 06-26` | 2026-11-18 | *"no market replay data"* |
| `NQ 12-26` | 2026-11-18 | downloads |

A batch run against one contract name therefore produces "no data" for every date outside that
contract's life — and the message is identical to a genuinely absent day, which makes it easy to
misdiagnose as a retention limit.

CME quarterly index futures expire on the third Friday of March, June, September and December, and
open interest rolls roughly eight days before expiry. The front month for a date is the earliest
quarterly expiry at least `ROLL_DAYS` (8) after it.

This convention is specific to those products. It is not a general futures rule.

---

## Truncated files, and why the log is the reference

Because writes are incremental, a run killed mid-transfer leaves a short file that "size greater
than zero" accepts forever. A silently corrupt day is worse than a missing one.

A replay file has no magic header or trailer — checked, the bytes at both ends are raw records — so
there is nothing in the file itself to validate. Integrity is established against an **external**
fact instead: the size this script recorded in its log when the download completed.

Real files match their logged size to within about 50 KB, because the log records size to 0.1 MB.
The tolerance is 160 KB: one rounding step plus slack.

Logs are merged oldest-first so that, where a date appears more than once, the most recent run
wins. Without that ordering an older entry could overwrite a newer one and a correctly
re-downloaded file would be reported as truncated forever.

Deleting `logs\` therefore weakens the guard to an absolute minimum-size test. Re-download the
range to rebuild both.

---

## Opening the window

Control Center has no standard `MenuBar` control, so the menu cannot be driven conventionally. Its
items are `MenuItem` controls inside a `Menu` named `PART_Menu`, addressable by AutomationId —
`ControlCenterMenuItemTools` for Tools. The popup is hosted in its own window, so every NinjaTrader
window is searched for the `Historical Data` item, matched on exact name.

The replay section starts **collapsed**, and while it is collapsed the instrument, date and
Download controls do not exist in the visual tree at all. Every lookup returns null. The expander
must be expanded before anything else works — the script does this at preflight and again before
each item.

---

## Testing changes

Before trusting a change, exercise all of these:

1. `-Plan` on a known range — must run with NinjaTrader **closed**.
2. A real download of one small symbol-day, then confirm the reported size matches the file.
3. `-Plan` again — the same date must now report `Complete True`.
4. Truncate a completed file to ~30 MB, then `-Plan`. It must report `Complete False`, and a real
   run must print `TRUNCATED` and re-download it.
5. A date older than the retention window — must report `NO-DATA` and dismiss the dialog.
6. `-DumpControls` — all four controls `FOUND`.
7. Argument errors — no dates, no instrument, and `-AutoContract` without `-Roots` must each exit
   `2` with a useful message.
