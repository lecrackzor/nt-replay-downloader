<#
.SYNOPSIS
    Batch-download NinjaTrader 8 Market Replay data.

.DESCRIPTION
    NinjaTrader's Historical Data window can request Market Replay data for one date at a time.
    This script drives that window for you: give it a date range and one or more instruments and
    it downloads each day in turn, skipping what is already on disk.

    It works by automating NinjaTrader's own dialog with Windows UI Automation. It uses your
    existing NinjaTrader session and the vendor's own delivery path, so there is nothing to
    configure and no credentials to supply.

.PARAMETER Instrument
    Instrument name exactly as NinjaTrader displays it, for example 'NQ 12-26'.
    If you are unsure of the spelling, open Tools > Historical Data and read the instrument
    box, or run -DumpControls to print it.

.PARAMETER Roots
    Symbol roots for use with -AutoContract, for example 'NQ','ES','RTY'. Each date is mapped
    to the quarterly contract that was front-month on that date, which is the only pairing
    NinjaTrader will serve.

.PARAMETER Dates
    Explicit dates as yyyy-MM-dd, comma separated.

.PARAMETER From
    First date of an inclusive range, yyyy-MM-dd.

.PARAMETER To
    Last date of an inclusive range, yyyy-MM-dd.

.PARAMETER WeekdaysOnly
    With -From/-To, skip Saturdays and Sundays. Exchange holidays are still requested; they
    simply come back as NO-DATA, which is harmless.

.PARAMETER AutoContract
    Resolve the contract per date from -Roots instead of using a fixed -Instrument.
    Built for CME quarterly index futures (NQ, ES, RTY, YM and their micros). Other products
    may use a different expiry and roll convention, so check before relying on it.

.PARAMETER Plan
    Print what would be downloaded and exit. Never touches NinjaTrader, so it works even when
    NinjaTrader is closed. Use this first for a large range.

.PARAMETER NinjaTraderHome
    NinjaTrader user data folder. Defaults to %USERPROFILE%\Documents\NinjaTrader 8.

.PARAMETER LogDir
    Where run logs are written. Defaults to a logs folder beside this script. These logs are
    also the reference used to detect half-finished files, so keep them.

.PARAMETER TimeoutSec
    Seconds to wait for one download before giving up on that date. A replay day is typically
    20-250 MB and completes in a few seconds on a normal connection.

.PARAMETER ClickGraceSec
    Seconds to wait for a click to visibly register before treating it as a missed click and
    retrying once. See docs/HOW-IT-WORKS.md.

.PARAMETER Trace
    Print per-step timings. Useful when reporting a performance problem.

.PARAMETER DumpControls
    Print the state of the Historical Data controls and exit. Run this after a NinjaTrader
    upgrade if the script stops finding the window.

.PARAMETER CloseWindow
    Close the Historical Data window when the run finishes. By default the window is left
    exactly as it was found.

.EXAMPLE
    .\Invoke-ReplayDownload.ps1 -Plan -AutoContract -Roots NQ,ES -From 2026-11-01 -To 2026-11-20 -WeekdaysOnly

    Show what a two-week, two-symbol run would fetch, without downloading anything.

.EXAMPLE
    .\Invoke-ReplayDownload.ps1 -AutoContract -Roots NQ,ES -From 2026-11-01 -To 2026-11-20 -WeekdaysOnly

    Download those fourteen weekdays for both symbols.

.EXAMPLE
    .\Invoke-ReplayDownload.ps1 -Instrument 'NQ 12-26' -Dates 2026-11-12,2026-11-13

    Download two specific days of one specific contract.

.NOTES
    Windows only. Requires NinjaTrader 8 to be running and connected.
    Exit codes: 0 success, 1 some dates failed, 2 bad arguments, 3 NinjaTrader unavailable.
#>

[CmdletBinding()]
param(
    [string[]]$Instrument,
    [string[]]$Dates,
    [datetime]$From,
    [datetime]$To,
    [switch]$WeekdaysOnly,
    [switch]$Plan,
    [switch]$DumpControls,
    [switch]$AutoContract,
    [string[]]$Roots,
    [string]$NinjaTraderHome = (Join-Path $env:USERPROFILE 'Documents\NinjaTrader 8'),
    [string]$LogDir,
    [int]$TimeoutSec = 420,
    [int]$ClickGraceSec = 6,
    [int]$PollMs = 200,
    [int]$SettleSec = 0,
    [switch]$Trace,
    [switch]$CloseWindow
)

$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

if (-not $LogDir) { $LogDir = Join-Path $PSScriptRoot 'logs' }
$Replay  = Join-Path $NinjaTraderHome 'db\replay'
$REFUSAL = 'no market replay|not available|no data|unavailable'

# Completion and integrity thresholds.
$MIN_BYTES  = 1MB      # every real replay day is far larger than this; anything smaller is a stub
$STABLE_MS  = 400      # size must hold still this long once the status line has gone blank
$SIZE_SLACK = 160KB    # log records size to 0.1 MB, so allow one rounding step plus slack

# ---------------------------------------------------------------------- output
# Two kinds of line matter for the integrity guard: the "[n/total] instrument date" header and the
# "OK <size> MB" result. They are written to a log file automatically, which is why the guard does
# not need the caller to redirect output.
$script:LogFile = $null

function Write-RunLog([string]$text) {
    if ($script:LogFile) {
        try { Add-Content -LiteralPath $script:LogFile -Value $text -Encoding utf8 } catch { }
    }
}

function Say([string]$text, [string]$color = 'Gray') {
    Write-Host $text -ForegroundColor $color
    Write-RunLog $text
}

# ---------------------------------------------------------------------- UI helpers
function Get-NTProcess {
    return (Get-Process NinjaTrader -ErrorAction SilentlyContinue | Select-Object -First 1)
}

function Get-NTWindows {
    $p = Get-NTProcess
    if (-not $p) { return @() }
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $pc = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $p.Id)
    return @($root.FindAll([System.Windows.Automation.TreeScope]::Children, $pc))
}

function Get-HDWindow {
    foreach ($w in Get-NTWindows) {
        if ($w.Current.AutomationId -eq 'HistoricalDataWindow') { return $w }
    }
    return $null
}

function Get-ById($parent, [string]$id) {
    if (-not $parent) { return $null }
    $c = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::AutomationIdProperty, $id)
    try { return $parent.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $c) }
    catch { return $null }
}

function Get-Val($e) {
    if (-not $e) { return '' }
    try { return $e.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value }
    catch { return '' }
}

function Set-Val($e, [string]$v) {
    if (-not $e) { throw "control not found while setting '$v'" }
    $e.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue($v)
}

function Invoke-El($e) {
    if (-not $e) { throw 'control not found to invoke' }
    $e.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
}

# Open Tools > Historical Data if it is not already up.
#
# Control Center has no standard menu bar, so the menu cannot be driven the usual way. Its items
# are MenuItem controls inside a Menu named 'PART_Menu' and are addressable by AutomationId, which
# is what this uses. The popup lives in its own window, so every window is searched for the item.
function Open-HistoricalData([int]$waitSec = 20) {
    $hd = Get-HDWindow
    if ($hd) { return $hd }

    $cc = $null
    foreach ($w in Get-NTWindows) {
        if ($w.Current.AutomationId -eq 'ControlCenter') { $cc = $w }
    }
    if (-not $cc) { return $null }

    $tools = Get-ById $cc 'ControlCenterMenuItemTools'
    if (-not $tools) { return $null }
    try {
        $ep = $tools.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern)
        if ($ep.Current.ExpandCollapseState -ne 'Expanded') {
            $ep.Expand(); Start-Sleep -Milliseconds 600
        }
    } catch { return $null }

    $mi = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::MenuItem)
    $chosen = $null
    foreach ($w in Get-NTWindows) {
        foreach ($m in $w.FindAll([System.Windows.Automation.TreeScope]::Descendants, $mi)) {
            $n = ''
            try { $n = $m.Current.Name } catch { continue }
            if ($n -eq 'Historical Data') { $chosen = $m; break }
        }
        if ($chosen) { break }
    }
    if (-not $chosen) { return $null }
    try { Invoke-El $chosen } catch { return $null }
    try { $ep.Collapse() } catch { }

    $deadline = (Get-Date).AddSeconds($waitSec)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 250
        $hd = Get-HDWindow
        if ($hd) {
            # The replay section starts collapsed. Until it is expanded the instrument, date and
            # Download controls are not in the visual tree at all, so every lookup returns null.
            $exp = Get-ById $hd 'HistoricalDataWindowMarketReplayExpander'
            if ($exp) {
                try {
                    $x = $exp.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern)
                    if ($x.Current.ExpandCollapseState -ne 'Expanded') {
                        $x.Expand(); Start-Sleep -Milliseconds 400
                    }
                } catch { }
            }
            return $hd
        }
    }
    return $null
}

# Re-acquire every control. Called before each item because dismissing a dialog rebuilds the tree,
# which makes any element reference captured earlier stale.
function Get-Controls {
    $win = Get-HDWindow
    if (-not $win) { return $null }
    $exp = Get-ById $win 'HistoricalDataWindowMarketReplayExpander'
    if ($exp) {
        try {
            $p = $exp.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern)
            if ($p.Current.ExpandCollapseState -ne 'Expanded') { $p.Expand() }
        } catch { }
    }
    return [pscustomobject]@{
        Win  = $win
        Inst = Get-ById $win 'textBox'
        Date = Get-ById $win 'HistoricalDataWindowMarketReplayDateSelector'
        Dl   = Get-ById $win 'HistoricalDataWindowMarketReplayDownloadButton'
        Msg  = Get-ById $win 'txtMessage'
    }
}

# The refusal dialog is a nested Window (AutomationId 'NTMessageBox') living inside the Historical
# Data window, not a top-level window, so a desktop enumeration does not see it. Looking it up by
# AutomationId inside the window is about 2 ms. Scanning every NinjaTrader window for text instead
# costs roughly 1.2 s per call with a few charts open, which is slow enough to be the dominant cost
# of a run.
function Get-MessageBox($win) {
    if (-not $win) { return $null }
    $c = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::AutomationIdProperty, 'NTMessageBox')
    try { return $win.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $c) }
    catch { return $null }
}

function Get-RefusalText($win, [string]$pattern = '') {
    $mb = Get-MessageBox $win
    if (-not $mb) { return $null }
    $tn = [System.Windows.Automation.ControlType]::Text
    $tc = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty, $tn)
    try {
        foreach ($t in $mb.FindAll([System.Windows.Automation.TreeScope]::Descendants, $tc)) {
            $n = $t.Current.Name
            if (-not $n) { continue }
            if (-not $pattern) { if ($n.Length -gt 12) { return $n } }
            elseif ($n -match $pattern) { return $n }
        }
    } catch { }
    return $null
}

function Clear-Dialog($win) {
    $mb = Get-MessageBox $win
    if (-not $mb) { return $false }
    $bc = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Button)
    try {
        foreach ($b in $mb.FindAll([System.Windows.Automation.TreeScope]::Descendants, $bc)) {
            if ($b.Current.Name -eq 'OK' -and $b.Current.IsEnabled) {
                Invoke-El $b
                Start-Sleep -Milliseconds 250
                return $true
            }
        }
    } catch { }
    return $false
}

function Wait-Enabled($c, [int]$sec = 25) {
    $deadline = (Get-Date).AddSeconds($sec)
    while ((Get-Date) -lt $deadline) {
        try {
            if ($c.Inst.Current.IsEnabled -and $c.Dl.Current.IsEnabled) { return $true }
        } catch { }
        Start-Sleep -Milliseconds 300
    }
    return $false
}

function File-State([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    $f = Get-Item -LiteralPath $path
    return [pscustomobject]@{ Size = $f.Length; MTime = $f.LastWriteTimeUtc }
}

# ---------------------------------------------------------------------- partial-file guard
# NinjaTrader writes a replay file incrementally, so a run that is killed mid-transfer leaves a
# short file that looks finished to a "size greater than zero" test. Accepting one would leave a
# silently truncated day in the dataset, which is worse than a missing day.
#
# A replay file has no magic header or trailer to check, so integrity is established against an
# external fact instead: the size this script recorded in its log when the download completed.
function Get-RecordedSizes([string]$logDir) {
    $map = @{}
    if (-not (Test-Path -LiteralPath $logDir)) { return $map }
    # Merge every log oldest first, so where a date appears more than once the most recent run wins.
    # Without the sort an older entry could overwrite a newer one and a re-downloaded file would be
    # reported as truncated forever.
    $logs = Get-ChildItem -LiteralPath $logDir -Filter *.log -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime
    foreach ($lg in $logs) {
        $cur = $null
        foreach ($line in (Get-Content -LiteralPath $lg.FullName -ErrorAction SilentlyContinue)) {
            if ($line -match '^\[(\d+)/(\d+)\]\s+(.+?)\s+(\d{4})-(\d{2})-(\d{2})\s*$') {
                $cur = '{0}|{1}{2}{3}' -f $Matches[3].Trim(), $Matches[4], $Matches[5], $Matches[6]
            } elseif ($line -match '^\s+OK\s+([\d\.]+) MB' -and $cur) {
                $map[$cur] = [int64]([double]$Matches[1] * 1MB)
            }
        }
    }
    return $map
}

function Test-Complete([string]$path, [hashtable]$recorded, [string]$key) {
    $st = File-State $path
    if (-not $st -or $st.Size -lt $MIN_BYTES) { return $false }
    if ($recorded.ContainsKey($key)) {
        if ([math]::Abs($st.Size - $recorded[$key]) -gt $SIZE_SLACK) { return $false }
    }
    return $true
}

function Step($name, $sw) {
    if ($Trace) {
        Write-Host ("        [trace] {0,-14} {1,6} ms" -f $name, $sw.ElapsedMilliseconds) -ForegroundColor DarkMagenta
    }
    $sw.Restart()
}

# ---------------------------------------------------------------------- contract resolution
# NinjaTrader will not serve a date outside a contract's life. Requesting one produces the same
# "no market replay data" refusal as a genuinely missing day, which is easy to mistake for a
# retention limit. Mapping each date to the contract that was front-month then avoids that.
#
# CME quarterly index futures expire on the third Friday of March, June, September and December,
# and open interest rolls a week or so before expiry. The front month for a date is therefore the
# earliest quarterly expiry at least ROLL_DAYS after it.
$ROLL_DAYS = 8
$MONTH_CODE = @{ 3 = '03'; 6 = '06'; 9 = '09'; 12 = '12' }

function Get-ThirdFriday([int]$year, [int]$month) {
    $d = [datetime]::new($year, $month, 1)
    $firstFri = $d.AddDays((5 - [int]$d.DayOfWeek + 7) % 7)
    return $firstFri.AddDays(14)
}

function Resolve-Contract([string]$root, [datetime]$date) {
    $cands = @()
    foreach ($y in @($date.Year, $date.Year + 1)) {
        foreach ($m in 3, 6, 9, 12) {
            $cands += [pscustomobject]@{
                Exp  = Get-ThirdFriday $y $m
                Name = ('{0} {1}-{2}' -f $root, $MONTH_CODE[$m], ($y % 100).ToString('00'))
            }
        }
    }
    foreach ($c in ($cands | Sort-Object Exp)) {
        if ($c.Exp -ge $date.AddDays($ROLL_DAYS)) { return $c.Name }
    }
    return $null
}

# ---------------------------------------------------------------------- plan
if (-not $Dates -and $From -and $To) {
    $d = $From.Date; $Dates = @()
    while ($d -le $To.Date) {
        if (-not $WeekdaysOnly -or $d.DayOfWeek -notin 'Saturday', 'Sunday') { $Dates += $d }
        $d = $d.AddDays(1)
    }
}

if ($DumpControls) {
    $c = Get-Controls
    if (-not $c) {
        Write-Host "The 'Historical Data' window is not open. Open it with Tools > Historical Data, then re-run." -ForegroundColor Red
        exit 3
    }
    $ok = Wait-Enabled $c 10
    Write-Host "controls enabled: $ok"
    foreach ($n in 'Inst', 'Date', 'Dl', 'Msg') {
        Write-Host ('  {0,-5} {1}' -f $n, $(if ($c.$n) { 'FOUND' } else { 'MISSING' }))
    }
    Write-Host "  instrument value: '$((Get-Val $c.Inst))'"
    Write-Host "  date value      : '$((Get-Val $c.Date))'"
    exit 0
}

if (-not $Dates) {
    Write-Host 'No dates given. Use -Dates yyyy-MM-dd[,...] or -From/-To (optionally -WeekdaysOnly).' -ForegroundColor Red
    Write-Host 'Example: .\Invoke-ReplayDownload.ps1 -AutoContract -Roots NQ -From 2026-11-01 -To 2026-11-20 -WeekdaysOnly'
    exit 2
}
$Dates = @($Dates | ForEach-Object { ([datetime]$_).Date })

if (-not $AutoContract -and -not $Instrument) {
    Write-Host 'Name an instrument with -Instrument, or use -AutoContract with -Roots.' -ForegroundColor Red
    Write-Host "  -Instrument 'NQ 12-26'                  a fixed contract"
    Write-Host "  -AutoContract -Roots NQ,ES              contract resolved per date"
    Write-Host 'Open Tools > Historical Data in NinjaTrader to see the exact instrument spelling.'
    exit 2
}

# Build the work queue. With -AutoContract the instrument is derived per date, because that is the
# only pairing NinjaTrader will serve.
if ($AutoContract) {
    # A comma separated string arrives as multiple arguments, so 'NQ,ES' becomes two values and a
    # single quoted 'NQ,ES,RTY' becomes one. Flattening embedded commas handles both call styles.
    $roots = @($Roots | ForEach-Object { $_ -split ',' } |
               ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if (-not $roots) {
        Write-Host '-AutoContract needs -Roots, for example -Roots NQ,ES' -ForegroundColor Red
        exit 2
    }
    $queue = foreach ($root in $roots) {
        foreach ($d in $Dates) {
            $dd = [datetime]$d
            $contract = Resolve-Contract $root $dd
            if (-not $contract) { continue }
            $tgt = Join-Path (Join-Path $Replay $contract) ($dd.ToString('yyyyMMdd') + '.nrd')
            [pscustomobject]@{
                Instrument = $contract
                Date       = $dd
                DateStr    = $dd.ToString('yyyy-MM-dd')
                Target     = $tgt
                Exists     = (Test-Path -LiteralPath $tgt)
            }
        }
    }
    $shownInstruments = $roots -join ', '
} else {
    $queue = foreach ($inst in $Instrument) {
        foreach ($d in $Dates) {
            $dd = [datetime]$d
            $tgt = Join-Path (Join-Path $Replay $inst) ($dd.ToString('yyyyMMdd') + '.nrd')
            [pscustomobject]@{
                Instrument = $inst
                Date       = $dd
                DateStr    = $dd.ToString('yyyy-MM-dd')
                Target     = $tgt
                Exists     = (Test-Path -LiteralPath $tgt)
            }
        }
    }
    $shownInstruments = $Instrument -join ', '
}

Write-Host ''
Write-Host 'BATCH MARKET REPLAY DOWNLOAD' -ForegroundColor Cyan
Write-Host ('  instruments : {0}' -f $shownInstruments)
Write-Host ('  dates       : {0} .. {1}  ({2} days)' -f
             ([datetime]$Dates[0]).ToString('yyyy-MM-dd'),
             ([datetime]$Dates[-1]).ToString('yyyy-MM-dd'), $Dates.Count)
Write-Host ('  requests    : {0}   already on disk: {1}' -f
             $queue.Count, (@($queue | Where-Object Exists).Count))
Write-Host ''

$recorded = Get-RecordedSizes $LogDir

# Squash a list of dates into readable spans, so 40 missing days print as "09-22..10-23" rather
# than as forty comma separated stamps. A gap of three days or less stays inside one span, which
# keeps weekends from breaking a continuous stretch into pieces.
function Compress-DateRanges([datetime[]]$dates) {
    if (-not $dates) { return '' }
    $out = @(); $start = $null; $prev = $null
    foreach ($dt in ($dates | Sort-Object)) {
        if ($null -eq $start) { $start = $dt; $prev = $dt; continue }
        if (($dt - $prev).Days -le 3) { $prev = $dt; continue }
        $out += $(if ($start -eq $prev) { $start.ToString('MM-dd') }
                  else { '{0}..{1}' -f $start.ToString('MM-dd'), $prev.ToString('MM-dd') })
        $start = $dt; $prev = $dt
    }
    $out += $(if ($start -eq $prev) { $start.ToString('MM-dd') }
              else { '{0}..{1}' -f $start.ToString('MM-dd'), $prev.ToString('MM-dd') })
    return ($out -join '  ')
}

# -Plan never touches NinjaTrader: it only compares disk against the log, so it works whether or
# not NinjaTrader is running, and it is not blocked by the preflight below. This is also the fastest
# way to answer "what do I already have", so it prints coverage per instrument as well as the work
# list.
if ($Plan) {
    # Same completeness test as the run, so a half-finished file is reported as needing work rather
    # than counted as already present.
    $checked = foreach ($item in $queue) {
        $key = '{0}|{1}' -f $item.Instrument, $item.Date.ToString('yyyyMMdd')
        [pscustomobject]@{
            Instrument = $item.Instrument
            Date       = $item.Date
            DateStr    = $item.DateStr
            Complete   = Test-Complete $item.Target $recorded $key
            Target     = $item.Target
        }
    }

    $summary = $checked | Group-Object Instrument | Sort-Object Name | ForEach-Object {
        $done = @($_.Group | Where-Object Complete)
        $miss = @($_.Group | Where-Object { -not $_.Complete })
        [pscustomobject]@{
            Instrument = $_.Name
            Have       = $done.Count
            Missing    = $miss.Count
            Coverage   = '{0:P0}' -f ($done.Count / [math]::Max(1, $_.Count))
            Missing_   = Compress-DateRanges ($miss.Date)
        }
    }
    $summary | Format-Table Instrument, Have, Missing, Coverage, Missing_ -AutoSize

    if ($checked.Count -le 40) {
        $checked | Select-Object Instrument, DateStr, Complete, Target | Format-Table -AutoSize
    } else {
        Write-Host ('  {0} requests in total; add them to a run to download.' -f $checked.Count) -ForegroundColor DarkGray
    }
    Write-Host 'PLAN ONLY -- nothing downloaded.' -ForegroundColor Yellow
    exit 0
}

# ---------------------------------------------------------------------- preflight
# Establish the window before the first item, so a missing window is reported once and clearly
# instead of surfacing as a confusing per-item failure.
$ntProc = Get-NTProcess
if (-not $ntProc) {
    Write-Host ''
    Write-Host 'NinjaTrader is not running.' -ForegroundColor Red
    Write-Host '  Replay data is delivered over NinjaTrader''s own connection, so NinjaTrader must be'
    Write-Host '  running and connected. Start it, then re-run this script.'
    exit 3
}
$hd0 = Open-HistoricalData
if (-not $hd0) {
    Write-Host ''
    Write-Host 'Could not open Tools > Historical Data in NinjaTrader.' -ForegroundColor Red
    Write-Host '  Open it manually, then run -DumpControls to check that the script can see the controls.'
    exit 3
}
Write-Host ('  window      : Historical Data ready (pid {0})' -f $ntProc.Id) -ForegroundColor DarkGray

# The window is deliberately never minimised or maximised. InvokePattern against a minimised WPF
# window can silently do nothing -- no exception, button still reports enabled, no download starts
# -- so a run leaves the window exactly as it found it.

# ---------------------------------------------------------------------- run
New-Item -ItemType Directory -Force $LogDir | Out-Null
$script:LogFile = Join-Path $LogDir ('replay-{0}.log' -f (Get-Date).ToString('yyyyMMdd-HHmmss'))

$results = @()
$n = 0
if ($Trace) { Write-Host ('  [trace] recorded sizes from log: {0}' -f $recorded.Count) -ForegroundColor DarkMagenta }
$swStep = [Diagnostics.Stopwatch]::StartNew()
$fileTimes = @()

foreach ($item in $queue) {
    $n++
    $label = '{0} {1}' -f $item.Instrument, $item.DateStr
    $key = '{0}|{1}' -f $item.Instrument, $item.Date.ToString('yyyyMMdd')
    $itemSw = [Diagnostics.Stopwatch]::StartNew()
    Say ('[{0}/{1}] {2}' -f $n, $queue.Count, $label) 'Cyan'

    $before = File-State $item.Target
    if ($before -and $before.Size -gt 0) {
        if (Test-Complete $item.Target $recorded $key) {
            Write-Host ('        present already ({0:N1} MB), skip' -f ($before.Size / 1MB)) -ForegroundColor DarkGray
            $results += [pscustomobject]@{ Item = $label; Status = 'SKIP'; Detail = 'present' }
            continue
        }
        # Short file: a previous run was interrupted mid-transfer. Re-fetch it rather than keep a
        # truncated day.
        $expect = if ($recorded.ContainsKey($key)) { $recorded[$key] / 1MB } else { 0 }
        Write-Host ('        TRUNCATED ({0:N1} MB, log says {1:N1} MB) -- re-downloading' -f ($before.Size / 1MB), $expect) -ForegroundColor Yellow
        Remove-Item -LiteralPath $item.Target -Force -ErrorAction SilentlyContinue
    }

    # Re-acquire: dismissing a dialog from the previous item rebuilds the control tree.
    $c = Get-Controls
    if (-not $c) {
        Write-Host '        Historical Data window not up -- opening it' -ForegroundColor DarkGray
        Open-HistoricalData | Out-Null
        $c = Get-Controls
    }
    Step 'get-controls' $swStep
    if (-not $c) {
        Write-Host '        cannot open Historical Data -- is NinjaTrader running?' -ForegroundColor Red
        $results += [pscustomobject]@{ Item = $label; Status = 'FAIL'; Detail = 'no HD window' }
        break
    }

    if (-not (Wait-Enabled $c 25)) {
        $txt = Get-RefusalText $c.Win $REFUSAL
        if ($txt) { Write-Host ('        blocking dialog: {0}' -f $txt) -ForegroundColor DarkGray }
        Clear-Dialog $c.Win | Out-Null
        $c = Get-Controls
        if (-not $c -or -not (Wait-Enabled $c 15)) {
            $results += [pscustomobject]@{ Item = $label; Status = 'FAIL'; Detail = 'controls disabled' }
            continue
        }
    }

    # Set both fields, then verify them, and refuse to click if either does not read back.
    #
    # Only write a field that actually differs. NinjaTrader keeps the instrument between downloads
    # and it only changes at a contract roll, so rewriting it every item spent most of a second per
    # file re-entering a value that was already there. Reading it back first usually removes both
    # writes.
    $dtText = $item.Date.ToString('M/d/yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
    try {
        $curInst = (Get-Val $c.Inst).Trim()
        if ($curInst -ne $item.Instrument) {
            # Clear first. The instrument box is a search field, so writing into a box that already
            # holds a value merges into it rather than replacing it, which yields a stray leading
            # character. The verify below catches it if this still happens.
            Set-Val $c.Inst ''
            Start-Sleep -Milliseconds 120
            Set-Val $c.Inst $item.Instrument
            Start-Sleep -Milliseconds 220
        }
        $curDate = Get-Val $c.Date
        $curParsed = [datetime]::MinValue
        $dateSame = [datetime]::TryParse($curDate, [System.Globalization.CultureInfo]::InvariantCulture,
                                         'None', [ref]$curParsed) -and ($curParsed.Date -eq $item.Date)
        if (-not $dateSame) {
            Set-Val $c.Date ''
            Start-Sleep -Milliseconds 60
            Set-Val $c.Date $dtText
            Start-Sleep -Milliseconds 120
        }
    } catch {
        Write-Host ('        could not set fields: {0}' -f $_.Exception.Message) -ForegroundColor Red
        $results += [pscustomobject]@{ Item = $label; Status = 'FAIL'; Detail = 'set failed' }
        continue
    }
    Step 'arm-fields' $swStep

    $vInst = (Get-Val $c.Inst).Trim()
    $vDate = Get-Val $c.Date
    $parsed = [datetime]::MinValue
    $dateOk = [datetime]::TryParse($vDate, [System.Globalization.CultureInfo]::InvariantCulture,
                                   'None', [ref]$parsed) -and ($parsed.Date -eq $item.Date)
    if ($vInst -ne $item.Instrument -or -not $dateOk) {
        Write-Host ('        REFUSING TO CLICK: inst=''{0}'' date=''{1}''' -f $vInst, $vDate) -ForegroundColor Red
        $results += [pscustomobject]@{ Item = $label; Status = 'FAIL'; Detail = "verify '$vInst'/'$vDate'" }
        continue
    }

    # Click, then prove the click landed. See docs/HOW-IT-WORKS.md for why this is not optional.
    $clickAt = Get-Date
    try { Invoke-El $c.Dl } catch {
        Write-Host ('        click failed: {0}' -f $_.Exception.Message) -ForegroundColor Red
        $results += [pscustomobject]@{ Item = $label; Status = 'FAIL'; Detail = 'click failed' }
        continue
    }
    $clickConfirmed = $false
    $retriedClick = $false

    # --- detection loop
    #
    # The status line at the bottom of the window reports download state: it reads
    # 'Loading data ... <instrument>' once the click lands and goes blank when the transfer ends.
    # It is a TextBlock, so its text is in the element Name and ValuePattern throws; read from a
    # cached handle it costs about 0.1 ms, against roughly 115 ms for a fresh lookup.
    #
    # Completion is "message blank AND file at least $MIN_BYTES AND size unchanged for $STABLE_MS".
    # The size condition is a backstop in case the message blanks a moment before the write lands.
    #
    # A refused request never creates a file, so the dialog is only probed while the file is still
    # absent. That keeps the probe off the success path entirely.
    $outcome = 'PENDING'; $detail = ''
    $hardDeadline = (Get-Date).AddSeconds($TimeoutSec)
    $lastSize = -1
    $stableSince = $null
    $msg = $c.Msg
    if (-not $msg) { $msg = Get-ById $c.Win 'txtMessage' }

    while ((Get-Date) -lt $hardDeadline) {
        Start-Sleep -Milliseconds $PollMs

        $state = ''
        try { $state = $msg.Current.Name } catch { $state = $null }
        if ($state -match 'Loading data') { $clickConfirmed = $true }

        $st = File-State $item.Target
        $size = if ($st) { $st.Size } else { 0 }
        if ($size -gt 0) { $clickConfirmed = $true }

        if (-not $clickConfirmed -and -not $retriedClick -and
            ((Get-Date) - $clickAt).TotalSeconds -ge $ClickGraceSec) {
            $retriedClick = $true
            Write-Host '        click did not register -- restoring the window and retrying once' -ForegroundColor Yellow
            try {
                $wpr = $c.Win.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
                $wpr.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Normal)
                Start-Sleep -Milliseconds 600
            } catch { }
            $c2 = Get-Controls
            if ($c2 -and $c2.Dl) {
                try { Invoke-El $c2.Dl; $clickAt = Get-Date } catch {
                    Write-Host ('        retry click failed: {0}' -f $_.Exception.Message) -ForegroundColor Red
                }
            }
        }

        if ($size -gt 0) {
            $blank = ($state -eq '') -or ($null -eq $state)
            if ($size -eq $lastSize) {
                if ($null -eq $stableSince) { $stableSince = Get-Date }
                if ($blank -and $size -ge $MIN_BYTES -and
                    ((Get-Date) - $stableSince).TotalMilliseconds -ge $STABLE_MS) {
                    $outcome = 'OK'
                    $detail = '{0:N1} MB' -f ($size / 1MB)
                    break
                }
            } else {
                $stableSince = $null
                $lastSize = $size
                Write-Host ('        ...{0:N1} MB' -f ($size / 1MB)) -ForegroundColor DarkGray
            }
        } else {
            $lastSize = -1
            $stableSince = $null
            $txt = Get-RefusalText $c.Win $REFUSAL
            if ($txt) {
                $outcome = 'NO-DATA'; $detail = $txt
                Clear-Dialog $c.Win | Out-Null
                break
            }
        }
    }

    if ($outcome -eq 'PENDING') {
        # The file may have completed right at the deadline.
        $st = File-State $item.Target
        if ($st -and $st.Size -ge $MIN_BYTES) {
            $outcome = 'OK'; $detail = '{0:N1} MB' -f ($st.Size / 1MB)
        } else {
            $outcome = 'FAIL'; $detail = "timeout ${TimeoutSec}s, no file"
        }
    }

    if ($outcome -eq 'OK') {
        if ($SettleSec -gt 0) { Start-Sleep -Seconds $SettleSec }
        $st = File-State $item.Target
        if ($st) { $detail = '{0:N1} MB' -f ($st.Size / 1MB) }
        Step 'detect-ok' $swStep
        Say ('        OK  {0}  ({1:N1}s)' -f $detail, $itemSw.Elapsed.TotalSeconds) 'Green'
        $results += [pscustomobject]@{ Item = $label; Status = 'OK'; Detail = $detail }
        $fileTimes += $itemSw.Elapsed.TotalSeconds
    } else {
        Step 'detect-other' $swStep
        Say ('        {0}  {1}  ({2:N1}s)' -f $outcome, $detail, $itemSw.Elapsed.TotalSeconds) 'Yellow'
        $results += [pscustomobject]@{ Item = $label; Status = $outcome; Detail = $detail }
        $fileTimes += $itemSw.Elapsed.TotalSeconds
    }
}

# ---------------------------------------------------------------------- report
Write-Host ''
Write-Host 'SUMMARY' -ForegroundColor Cyan
$ok   = @($results | Where-Object Status -eq 'OK')
$skip = @($results | Where-Object Status -eq 'SKIP')
$nod  = @($results | Where-Object Status -eq 'NO-DATA')
$fail = @($results | Where-Object Status -eq 'FAIL')
Write-Host ('  downloaded {0}   already present {1}   NO DATA {2}   FAILED {3}' -f
             $ok.Count, $skip.Count, $nod.Count, $fail.Count)
if ($fileTimes.Count) {
    Write-Host ('  avg {0:N1}s/file over {1} items (min {2:N1}s, max {3:N1}s)' -f
        (($fileTimes | Measure-Object -Average).Average), $fileTimes.Count,
        ($fileTimes | Measure-Object -Minimum).Minimum, ($fileTimes | Measure-Object -Maximum).Maximum)
}
if ($nod.Count) {
    Write-Host ''
    Write-Host '  no data (the server does not hold these):' -ForegroundColor Yellow
    $nod | ForEach-Object { Write-Host ('    {0}' -f $_.Item) }
}
if ($fail.Count) {
    Write-Host ''
    Write-Host '  failures:' -ForegroundColor Red
    $fail | ForEach-Object { Write-Host ('    {0}  ({1})' -f $_.Item, $_.Detail) }
}
Write-Host ''
Write-Host ('  log         : {0}' -f $script:LogFile) -ForegroundColor DarkGray

if ($CloseWindow) {
    try {
        $wpc = $hd0.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
        $wpc.Close()
        Write-Host '  window      : closed' -ForegroundColor DarkGray
    } catch { Write-Host '  window      : could not close' -ForegroundColor DarkGray }
}

if ($fail.Count) { exit 1 }
exit 0
