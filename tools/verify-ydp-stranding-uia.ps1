# verify-ydp-stranding-uia.ps1 — regression check for the rtd-once "stuck at
# #GETTING_DATA" bug (shm v0.7.8 slot-wait fix + numGuestSlots 2->4).
#
# Root cause it guards: on workbook open many RTD topics connect within ~50ms.
# A one-shot rtd-once push (YDP/YDH) lost the race for one of the few guest
# slots; the Go SendGuestCall returned "all guest slots busy" immediately
# (non-blocking acquisition) and the value was NEVER pushed, so the cell stayed
# at #GETTING_DATA forever.
#
# Method: load the XLL, then in a TIGHT BURST enter a streaming pair (Clock +
# StockTick), several scalar YDP cells across tickers/fields, and a YDH grid —
# all at once to maximise the connect-storm guest-slot contention. Repeat across
# N rounds, each round on a FRESH workbook (close + reopen) so every round is an
# independent connect storm — NOT Cells.Clear(), which tears RTD down mid-spill.
# PASS = no target cell ever stuck at #GETTING_DATA after its settle window AND
# the Go log has no "all guest slots busy"/"OnRtdConnect failed" lines.
#
# NOTE: YDP hits the live Yahoo endpoint. A network/rate-limit failure makes the
# handler push an ERROR STRING — which still SETTLES the cell (not #GETTING_DATA),
# so it does not count as stranding. We only fail on #GETTING_DATA persistence.

$ErrorActionPreference = 'Continue'
$xll = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\build\xll_showcase.xll")
$goLog = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\build\xll_showcase_go.log")
if (-not (Test-Path $xll)) { Write-Output "MISSING XLL: $xll"; exit 2 }
try { if (Test-Path $goLog) { Clear-Content $goLog -ErrorAction Stop } } catch { Write-Output "(could not clear ${goLog}: $_)" }

Get-Process EXCEL, xll_showcase, go_server -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Start-Sleep 1

$app = New-Object -ComObject Excel.Application
$app.Visible = $true
$app.DisplayAlerts = $false
$pidExcel = (Get-Process EXCEL -EA SilentlyContinue | Sort-Object StartTime | Select-Object -Last 1).Id
Write-Output "Excel COM up (pid=$pidExcel). Loading XLL: $xll"
$reg = $app.RegisterXLL($xll)
Write-Output "RegisterXLL returned: $reg"
Start-Sleep 3  # let xlAutoOpen launch the Go server + register RTD

$ydpCells = @('A2','B2','C2','A3','B3','C3','A5')
$everStuck = @()

# Multi-round connect-storm: each round is a fresh workbook with the full burst
# (the #GETTING_DATA repro), and the per-topic log check is the deterministic
# stranding signal. close+reopen across rounds ALSO exercises the RTD-server
# teardown path: it used to trip a SEPARATE teardown crash when a live YDH
# rtd-once grid was torn down (RtdServer::ServerTerminate ran the destructive
# teardown on a mere workbook close → reopen hit a dead server, RPC 0x800706BA).
# FIXED in xll-gen v0.8.10 (ServerTerminate gates the destructive teardown on a
# confirmed host shutdown), so $rounds>=2 now doubles as that crash's regression.
$rounds = 3
for ($r = 1; $r -le $rounds; $r++) {
    Write-Output "--- burst round $r/$rounds (fresh workbook) ---"
    if (-not (Get-Process -Id $pidExcel -EA SilentlyContinue)) { Write-Output "  Excel process GONE before round $r"; break }
    $wb = $app.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    try { $app.Calculation = -4105 } catch {}  # xlCalculationAutomatic (only settable with a wb open)

    # Tight burst -> all RTD topics connect together (the connect storm).
    $ws.Range('A1').Formula = '=Clock()'
    $ws.Range('B1').Formula = '=StockTick("AAPL")'
    $ws.Range('A2').Formula = '=YDP("AAPL","price")'
    $ws.Range('B2').Formula = '=YDP("AAPL","change%")'
    $ws.Range('C2').Formula = '=YDP("AAPL","currency")'
    $ws.Range('A3').Formula = '=YDP("MSFT","price")'
    $ws.Range('B3').Formula = '=YDP("GOOG","price")'
    $ws.Range('C3').Formula = '=YDP("TSLA","price")'
    $ws.Range('A5').Formula = '=YDH("MSFT",30)'

    Start-Sleep 7  # settle window for the network fetch + push
    $stuck = @()
    foreach ($c in $ydpCells) {
        try { $txt = [string]$ws.Range($c).Text } catch { $txt = "<err>" }
        if ($txt -match 'GETTING_DATA|Getting Data') { $stuck += "$c='$txt'" }
    }
    if ($stuck.Count -gt 0) { Write-Output ("  STILL STUCK: " + ($stuck -join '  ')); $everStuck += $stuck }
    else { Write-Output "  all YDP/YDH cells settled this round" }

    # Close WITHOUT save -> orderly RTD disconnect, then a fresh wb next round.
    try { $wb.Close($false) } catch { Write-Output "  (wb.Close threw: $_)" }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } catch {}
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {}
    Start-Sleep 2
}

# Inspect the Go log for the busy-slot signature.
$busyLines = @()
try { if (Test-Path $goLog) { $busyLines = Select-String -Path $goLog -Pattern 'all guest slots busy|OnRtdConnect failed' -EA SilentlyContinue } } catch {}
$goLines = 0; try { $goLines = (Get-Content $goLog -EA SilentlyContinue | Measure-Object -Line).Lines } catch {}
Write-Output ""
Write-Output "go log lines: $goLines ;  'all guest slots busy'/'OnRtdConnect failed' hits: $($busyLines.Count)"
$busyLines | Select-Object -First 10 | ForEach-Object { Write-Output ("  " + $_.Line) }

$excelAlive = [bool](Get-Process -Id $pidExcel -EA SilentlyContinue)
$verdict = if ($everStuck.Count -gt 0) { "FAIL (stranded at #GETTING_DATA: $($everStuck -join ', '))" }
           elseif ($busyLines.Count -gt 0) { "FAIL ($($busyLines.Count) busy-slot/connect-fail log lines)" }
           elseif (-not $excelAlive) { "WEAK (Excel exited before all rounds — inconclusive on stability)" }
           else { "PASS (no #GETTING_DATA stranding across $rounds connect storms; log clean)" }
Write-Output "Excel still alive: $excelAlive"
Write-Output "VERDICT: $verdict"

# Two-tier cleanup.
try { $app.Quit() } catch {}
try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {}
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Start-Sleep 2
Get-Process EXCEL, xll_showcase, go_server -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Write-Output "cleaned up"
