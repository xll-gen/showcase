# ghost-check.ps1 -- focused S1' ghost-Excel repro + evidence capture.
#
# Launches Excel with the showcase XLL, waits robustly for the ready window,
# invokes the Build ribbon command (heavy calc -> CalculationEnded -> arms the
# xlcOnTime deferred runner), then faithfully closes via WindowPattern.Close()
# (a user clicking X -- NOT COM Application.Quit, which would mask the bug).
# After close it observes for 30s WITHOUT force-killing and reports:
#   * whether EXCEL.EXE actually EXITS (the S1' ghost question)
#   * whether the Go server is reaped
#   * whether _go.log / _native.log become FREE
#   * whether a window reopens (S1)
#
# -KillExcelOnClose : simulate the user force-killing the ghost right at close
#                     (races graceful teardown's CloseHandle(hJob) -> S2 orphan).
#
# This is a more robust sibling of diagnose-close-uia.ps1: it settles before the
# first window poll (the COM-bounce variant makes Excel slow to show its window)
# and tracks the window-owning pid (EXCEL.EXE can relaunch off the Start-Process pid).
#
# NOTE on shared scaffold: the heavy lifting (UIA bootstrap, lock probe, window
# enumeration, faithful-close + Alt+F4 escalation, two-tier teardown) lives in
# uia-common.ps1. This file keeps only the ghost-specific sequencing + verdict.
param([switch]$KillExcelOnClose, [switch]$FastClose)
. "$PSScriptRoot\uia-common.ps1"
$ErrorActionPreference = 'Continue'
$P = Resolve-ShowcasePaths
$xll = $P.Xll; $goLog = $P.GoLog; $nativeLog = $P.NativeLog

Initialize-Uia

Write-Output "=== ghost-check : XLL=$xll ==="
# Environment before product: an unmet Trust Center lever produces this
# script's exact failure symptoms. See Assert-ExcelTrustPreconditions.
Assert-ExcelTrustPreconditions -XllPath $xll -RequireRtd
# Drop post-crash resiliency residue BEFORE launching. DisabledItems and
# StartupItems are pure harness debris -- a previous crashed run can leave the
# add-in on Excel's disabled list, and then this script measures an add-in that
# was never loaded and blames the product. DocumentRecovery is deliberately NOT
# touched (it can hold a real user's unsaved workbooks); the gate above WARNs
# about it and leaves the choice to a human.
Clear-ExcelResiliency
Stop-ShowcaseProcesses
Start-Sleep 1

$tmp = New-ShowcaseWorkbook -Name "ghost_check.xlsx"

# $launchedAt gates the native-log rung of the load ladder: this script does NOT
# clear the logs, so a log left by an EARLIER run would otherwise read as proof
# that the add-in loaded in THIS one.
$launchedAt = Get-Date
$p = Start-ShowcaseExcel -Workbook $tmp -Xll $xll
Write-Output "launched EXCEL pid=$($p.Id)"

# Robust wait: settle first (the COM-bounce makes Excel slow to materialize, and
# polling UIA too eagerly during init can transiently throw on Current.Name),
# then poll up to ~30s.
$excelPid = $p.Id
$win = Wait-ShowcaseWindow -LaunchedPid $p.Id -ResolvedPid ([ref]$excelPid) -SettleFirst 6 -Tries 38
if (-not $win) {
    # "No ready window" has two causes with one symptom: the add-in never loaded
    # (environment) or it loaded and wedged Excel (product). Say which.
    Write-Output "FAIL: no ready window"
    Write-XllLoadDiagnosis -XllPath $xll -NativeLog $nativeLog -Since $launchedAt
    Stop-ShowcaseProcesses; return
}
Write-Output "ready Excel window pid=$excelPid (launched pid=$($p.Id))"
Start-Sleep 2

$srvBefore = Get-Process xll_showcase -EA SilentlyContinue
Write-Output ("server pids after load: " + (($srvBefore.Id) -join ','))
# No server process after the window is up is the signature of an add-in that was
# never loaded at all -- and the ghost/orphan verdicts below would be about nothing.
if (-not $srvBefore) {
    Write-Output "WARN: no server process after load -- the add-in should be loaded by now"
    Write-XllLoadDiagnosis -XllPath $xll -NativeLog $nativeLog -Since $launchedAt
}

# select custom ribbon tab + invoke Build (heavy calc -> CalculationEnded -> arms xlcOnTime)
if (Invoke-RibbonButton -Window $win -NamePattern 'Build') { Write-Output "invoked Build" }
else {
    # The showcase ribbon tab IS the add-in. No Build button = the add-in did not
    # load its ribbon, so the close-time verdicts below have nothing to observe.
    Write-Output "WARN: Build button not found/invoked"
    Write-XllLoadDiagnosis -XllPath $xll -NativeLog $nativeLog -Since $launchedAt
}
# -FastClose: close almost immediately after Build to catch the xlcOnTime
# deferred runner while it is still ARMED (it arms at CalculationEnded and Excel
# dispatches the OnTime macro at the next idle, which a normal 5s wait lets drain).
# This is the HIGH #2 verification: does CancelDeferredRunner de-queue an armed runner?
if ($FastClose) { Start-Sleep -Milliseconds 250; Write-Output "FAST CLOSE (250ms after Build)" } else { Start-Sleep 5 }

# === FAITHFUL CLOSE via WindowPattern.Close() + Don't-Save, escalating to a
# faithful Alt+F4 if the app frame lingers (so teardown actually fires). =======
Close-ExcelWindowFaithful -Window $win -Escalate:(-not $KillExcelOnClose)

if ($KillExcelOnClose) {
    Write-Output "--- KILL: force-terminating EXCEL immediately (simulating user/crash) ---"
    Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
}

Write-Output "--- observing natural fate for 30s (no kill) ---"
$windowGoneAt=$null; $reopenAt=$null; $reopenPid=$null; $excelExitAt=$null; $sawWindowGone=$false
for ($t=0; $t -lt 30; $t++) {
    Start-Sleep 1
    $wins = Get-ExcelWindows
    $exPid = (Get-Process EXCEL -EA SilentlyContinue).Id
    $svPid = (Get-Process xll_showcase -EA SilentlyContinue).Id
    if (-not $sawWindowGone -and $wins.Count -eq 0) { $windowGoneAt=$t; $sawWindowGone=$true }
    if ($sawWindowGone -and $wins.Count -gt 0 -and -not $reopenAt) { $reopenAt=$t; try { $reopenPid=$wins[0].Current.ProcessId } catch {} }
    if (-not $exPid -and -not $excelExitAt) { $excelExitAt=$t }
    Write-Output ("  t={0,2}s  win={1}  EXCEL=[{2}]  server=[{3}]" -f $t, $wins.Count, ($exPid -join ','), ($svPid -join ','))
    if ($wins.Count -eq 0 -and -not $exPid -and -not $svPid) { Write-Output "  (fully settled)"; break }
}

Write-Output "--- final state ---"
$excelNow=(Get-Process EXCEL -EA SilentlyContinue).Id
$srvNow=(Get-Process xll_showcase -EA SilentlyContinue).Id
$orphan=($srvNow -and -not $excelNow)
$ghost=($excelNow -and (Get-ExcelWindows).Count -eq 0)
Write-Output ("EXCEL now: [" + ($excelNow -join ',') + "]   server now: [" + ($srvNow -join ',') + "]")
Write-Output ("window gone at: ${windowGoneAt}s   reopened at: ${reopenAt}s (pid=$reopenPid)   EXCEL exited at: ${excelExitAt}s")
Write-Output ("LOCK _go.log     : " + (Probe-Lock $goLog))
Write-Output ("LOCK _native.log : " + (Probe-Lock $nativeLog))
Write-Output "=== VERDICT ==="
Write-Output ("  S1 window reopened          : " + $(if($reopenAt -ne $null){"YES at ${reopenAt}s (pid=$reopenPid)"}else{"no"}))
Write-Output ("  S1' ghost EXCEL (no window) : " + $(if($ghost){"YES (pid=$excelNow lingering windowless)"}else{"no"}))
Write-Output ("  S2 orphan server (no EXCEL) : " + $(if($orphan){"YES (server=$srvNow)"}else{"no"}))
Stop-ShowcaseProcesses
Write-Output "cleaned up"
