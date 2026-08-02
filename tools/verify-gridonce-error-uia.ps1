# verify-gridonce-error-uia.ps1 — regression check for the grid rtd-once ERROR
# path (xll-gen v0.8.37).
#
# The bug it guards: a one-shot GRID handler that returned an error had nowhere
# to put it. RtdOnceGridRegistry stored results only, so the failure was dropped
# on the floor and the wrapper fell through to re-issue xlfRtd against a topic
# that was already connected — which wedges the cell at the loading placeholder
# FOREVER (AGENTS.md §19.3). The scalar rtd-once path had always surfaced its
# errors; the grid twin had not.
#
# The fix stores the message as a TRANSIENT entry: it paints once (the cell gets
# the error text, because a grid return is Q/LPXLOPER12 so a string fits), and
# it is reclaimed at calc end when no live topic holds the key — so it is NOT
# subject to the function's memoize_ttl and the next recalc genuinely retries.
# Those are two separate claims and this script tests them separately.
#
# Method: YDH(ticker, 0) fails on the ARGUMENT (`days must be >= 1`) before any
# network call, so the error is deterministic and this test does not depend on
# Yahoo being reachable or on a ticker staying invalid. YDH declares
# memoize_ttl: 10m, which is exactly what makes check 2 meaningful — if errors
# were memoized like results, the second identical call would be served from the
# cache and the handler would run only once.
#
# PASS requires all four:
#   1. the error TEXT reaches the cell (not the loading placeholder, not blank)
#   2. an identical call issued later re-invokes the handler (error not memoized)
#   3. a VALID call still spills a real grid (the error entry poisons nothing)
#   4. re-entering the failing formula errors again rather than wedging

. "$PSScriptRoot\uia-common.ps1"
$ErrorActionPreference = 'Continue'

$P = Resolve-ShowcasePaths
$xll = $P.Xll; $goLog = $P.GoLog; $nativeLog = $P.NativeLog
if (-not (Test-Path $xll)) { Write-Output "MISSING XLL: $xll"; exit 2 }
# Environment before product: an unmet Trust Center lever produces this
# script's exact failure symptoms. See Assert-ExcelTrustPreconditions.
Assert-ExcelTrustPreconditions -XllPath $xll -RequireRtd

foreach ($l in @($goLog, $nativeLog)) {
    try { if (Test-Path $l) { Clear-Content $l -ErrorAction Stop } } catch { Write-Output "(could not clear ${l}: $_)" }
}

Stop-ShowcaseProcesses -IncludeGoServer
Start-Sleep 1

$fails = New-Object System.Collections.Generic.List[string]
function Fail([string]$m) { $script:fails.Add($m); Write-Output "  FAIL: $m" }
function Pass([string]$m) { Write-Output "  pass: $m" }

# The loading placeholder(s) a cell must NOT be sitting on once settled.
$pending = @('#GETTING_DATA', '#N/A')

# Excel COM surfaces cell errors through Value2 as negative Int32 codes.
# NEVER decide "settled" from .Text: .Text is what the cell DISPLAYS, so a
# column too narrow to fit '#GETTING_DATA' renders it as '#######', which
# passes a naive string check and makes the driver tear Excel down while the
# handler is still running. That mistake produced a fake "the cell is wedged"
# reading here once already — the giveaway was xlAutoClose landing 112 ms after
# the RTD connect. Value2 is display-independent.
$xlErrGettingData = -2146826245   # 0x800A07FA
$xlErrNA          = -2146826246   # 0x800A07F9

function Test-Pending($ws, [string]$addr) {
    $v = $null
    try { $v = $ws.Range($addr).Value2 } catch { return $true }
    if ($null -eq $v) { return $true }                       # not painted yet
    if ($v -is [int] -and ($v -eq $xlErrGettingData -or $v -eq $xlErrNA)) { return $true }
    return $false
}

# Wait-Settled polls a cell until it stops showing a loading placeholder.
# Returns the final .Text (for reporting only). RTD pushes arrive on Excel's own
# schedule, so this polls rather than sleeping a fixed amount.
function Wait-Settled($ws, [string]$addr, [int]$timeoutSec = 45) {
    $deadline = (Get-Date).AddSeconds($timeoutSec)
    while ((Get-Date) -lt $deadline) {
        if (-not (Test-Pending $ws $addr)) { break }
        Start-Sleep -Milliseconds 400
    }
    try { return [string]$ws.Range($addr).Text } catch { return '' }
}

# Set-Formula and Test-DynamicArrays are shared in uia-common.ps1 (both this
# script and verify-ydp-stranding-uia.ps1 need the spill path, and having two
# copies is how one of them drifted onto .Formula unnoticed).

# Count the native-log WARN the grid-once error branch emits. This is what
# distinguishes "handler ran again" from "cache served the stored error".
function Count-HandlerFailures {
    if (-not (Test-Path $nativeLog)) { return 0 }
    return @(Select-String -Path $nativeLog -Pattern 'rtd-once YDH: one-shot handler failed' -SimpleMatch -ErrorAction SilentlyContinue).Count
}

$app = $null
try {
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $true
    $app.DisplayAlerts = $false
    Write-Output "RegisterXLL: $xll"
    $app.RegisterXLL($xll) | Out-Null
    Start-Sleep 3   # xlAutoOpen launches the Go server and registers the RTD server

    $wb = $app.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)

    Test-DynamicArrays -Worksheet $ws

    # --- 1. the error reaches the cell -------------------------------------
    Write-Output "[1] grid rtd-once handler error must PAINT, not wedge"
    Set-Formula -Worksheet $ws -Address "A1" -Formula '=YDH("MSFT",0)'
    $t1 = Wait-Settled $ws "A1"
    Write-Output "    A1 settled to: '$t1'"
    if (Test-Pending $ws "A1") {
        Fail "A1 never settled (still '$t1') — the grid-once error was dropped and the cell is wedged"
    } elseif ($t1 -notmatch 'days must be') {
        Fail "A1 settled to '$t1' but does not carry the handler's message"
    } else {
        Pass "A1 shows the handler error text"
    }
    $afterFirst = Count-HandlerFailures

    # --- 2. an identical later call must RE-INVOKE the handler --------------
    # YDH is memoize_ttl:10m. If the error were stored like a result, this
    # second call would be a cache hit and the handler would NOT run again.
    Write-Output "[2] identical args later must re-run the handler (error is transient, not memoized)"
    Start-Sleep 2
    Set-Formula -Worksheet $ws -Address "A10" -Formula '=YDH("MSFT",0)'
    $t2 = Wait-Settled $ws "A10"
    Write-Output "    A10 settled to: '$t2'"
    $afterSecond = Count-HandlerFailures
    Write-Output "    handler-failure log lines: $afterFirst -> $afterSecond"
    if ($t2 -notmatch 'days must be') {
        Fail "A10 settled to '$t2'; the repeat call did not surface the error"
    } elseif ($afterSecond -le $afterFirst) {
        Fail "handler-failure count did not increase ($afterFirst -> $afterSecond) — the error was MEMOIZED under memoize_ttl instead of being transient"
    } else {
        Pass "handler re-ran for the repeat call ($afterFirst -> $afterSecond)"
    }

    # --- 3. a valid call still works ----------------------------------------
    Write-Output "[3] a VALID grid-once call must still spill"
    if ($hasDA) {
        # .Formula2, NOT .Formula — see the DA probe above. Via .Formula this
        # spills nothing and the check below would pass vacuously.
        $ws.Range("C1").Formula2 = '=YDH("MSFT",30)'
    } else {
        $ws.Range("C1:H12").FormulaArray = '=YDH("MSFT",30)'
    }
    $t3 = Wait-Settled $ws "C1" 90
    $v3 = $null; try { $v3 = $ws.Range("C1").Value2 } catch {}
    Write-Output "    C1 settled to: Text='$t3' Value2='$v3'"
    if (Test-Pending $ws "C1") {
        Fail "C1 never settled — a VALID grid-once call is wedged at the loading placeholder"
    } elseif ("$v3" -match 'days must be') {
        Fail "C1 got the error from the OTHER args — the error entry is keyed too broadly"
    } elseif ("$v3".Trim() -ne 'Date') {
        # A live-endpoint failure is not this test's business. It still SETTLES
        # the cell (that is the fix under test), so report it as inconclusive
        # for check 3 only, not as a failure of the grid-once path.
        Write-Output "    INCONCLUSIVE: settled, but to '$v3' rather than the 'Date' header (live-endpoint failure, not the grid-once path)"
    } else {
        Pass "valid grid still spills (C1='Date')"
        $spill = $null; try { $spill = $ws.Range("D1").Value2 } catch {}
        if ("$spill".Trim() -ne 'Open') { Fail "grid did not spill across (D1='$spill', expected 'Open')" }
        else { Pass "grid spilled (D1='Open')" }
    }

    # --- 4. re-entering the failing formula must error again, not wedge -----
    Write-Output "[4] re-entering the failing formula must error again (self-heal, no wedge)"
    $before4 = Count-HandlerFailures
    $ws.Range("A1").ClearContents()
    Start-Sleep 2
    Set-Formula -Worksheet $ws -Address "A1" -Formula '=YDH("MSFT",0)'
    $t4 = Wait-Settled $ws "A1"
    $after4 = Count-HandlerFailures
    Write-Output "    A1 settled to: '$t4' (handler failures $before4 -> $after4)"
    if ($t4 -notmatch 'days must be') {
        Fail "A1 re-entry settled to '$t4' — the cell did not recover"
    } elseif ($after4 -le $before4) {
        Fail "A1 re-entry did not re-run the handler ($before4 -> $after4)"
    } else {
        Pass "re-entry errors again and re-runs the handler"
    }

    $wb.Close($false)
} catch {
    Fail "DRIVER ERROR: $_"
} finally {
    # Stage 1: graceful.
    try { if ($app) { $app.DisplayAlerts = $false; $app.Quit() } } catch {}
    try { if ($app) { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep 2
    # Stage 2: force-kill survivors (Excel + the Go server the XLL spawned).
    $left = Get-Process EXCEL, xll_showcase -EA SilentlyContinue
    if ($left) {
        Write-Output "force-killing residual: $(($left | ForEach-Object { "$($_.ProcessName)#$($_.Id)" }) -join ', ')"
        $left | Stop-Process -Force -EA SilentlyContinue
    }
}

Write-Output ""
if ($fails.Count -eq 0) {
    Write-Output "RESULT: PASS"
    exit 0
} else {
    Write-Output "RESULT: FAIL ($($fails.Count))"
    $fails | ForEach-Object { Write-Output " - $_" }
    exit 1
}
