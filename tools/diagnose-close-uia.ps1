# diagnose-close-uia.ps1 — UIA repro + evidence capture for the two close-time bugs:
#   (S1) "창을 다 닫았는데 창이 하나 다시 열린다"  -> uncancelled xlcOnTime reopen
#   (S2) "native/go log를 못 지운다, Excel 프로세스는 없다" -> orphaned Go server holding inherited log handle
#
# Faithful close: drives the REAL window close via UIA WindowPattern.Close() (a user
# clicking X), handling the Korean/English save dialog — NOT a COM Application.Quit
# (which would mask the bug). After close it watches for:
#   * REOPEN  : any Excel top-level window / new EXCEL pid reappearing post-close
#   * ORPHAN  : xll_showcase server process still alive after EXCEL is gone
#   * LOCK    : _go.log / _native.log still locked, with the holding process named
#               via the Restart Manager API (rstrtmgr.dll) — no Sysinternals needed.
#
# Usage:  pwsh -File tools\diagnose-close-uia.ps1
# Requires the built XLL at ..\build\xll_showcase.xll and an interactive desktop session.

param([switch]$KillExcelOnClose)   # simulate user force-killing the ghost Excel (Task Manager / crash)
. "$PSScriptRoot\uia-common.ps1"
$ErrorActionPreference = 'Continue'
$P = Resolve-ShowcasePaths
$xll = $P.Xll; $goLog = $P.GoLog; $nativeLog = $P.NativeLog

Initialize-Uia

Write-Output "=== diagnose-close-uia : XLL=$xll ==="
# Environment before product: an unmet Trust Center lever produces this
# script's exact failure symptoms. See Assert-ExcelTrustPreconditions.
Assert-ExcelTrustPreconditions -XllPath $xll -RequireRtd
Stop-ShowcaseProcesses
Start-Sleep 1

# blank workbook (XLL alone shows Start screen)
$tmp = New-ShowcaseWorkbook -Name "uia_close_diag.xlsx"

# launch Excel normally (no held COM client)
$p = Start-ShowcaseExcel -Workbook $tmp -Xll $xll
$origPid = $p.Id
Write-Output "launched EXCEL pid=$origPid"

$win = Wait-ShowcaseWindow -LaunchedPid $origPid -Tries 30
if (-not $win) { Write-Output "FAIL: no ready window"; Stop-ShowcaseProcesses; return }
Start-Sleep 2

$srvBefore = Get-Process xll_showcase -EA SilentlyContinue
Write-Output ("server pids after load: " + (($srvBefore.Id) -join ',' ))

# select custom ribbon tab + invoke Build (heavy calc -> CalculationEnded -> arms xlcOnTime)
if (Invoke-RibbonButton -Window $win -NamePattern 'Build') { Write-Output "invoked Build" }
Start-Sleep 5   # let the build/calc run and arm the deferred runner

# === FAITHFUL CLOSE via WindowPattern.Close() (user clicking X) ===============
Close-ExcelWindowFaithful -Window $win

# simulate the user killing the ghost Excel right after close (Task Manager / crash):
# this races the graceful teardown's CloseHandle(hJob) and exposes server orphaning.
if ($KillExcelOnClose) {
    Write-Output "--- KILL: force-terminating ghost EXCEL immediately (simulating user/crash) ---"
    Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
}

# === OBSERVE NATURAL FATE for 40s — NO force-kill until the very end =========
# Sample every 1s: EXCEL pids, visible Excel windows, server pids. Record the
# moment the window disappears, the moment it REAPPEARS (S1), whether EXCEL ever
# fully exits, and whether the server outlives EXCEL (S2 orphan).
Write-Output "--- observing natural fate for 40s (no kill) ---"
$windowGoneAt = $null; $reopenAt = $null; $reopenPid = $null
$excelExitAt  = $null; $sawWindowGone = $false
for ($t=0; $t -lt 40; $t++) {
    Start-Sleep 1
    $wins  = Get-ExcelWindows
    $exPid = (Get-Process EXCEL -EA SilentlyContinue).Id
    $svPid = (Get-Process xll_showcase -EA SilentlyContinue).Id
    if (-not $sawWindowGone -and $wins.Count -eq 0) { $windowGoneAt = $t; $sawWindowGone = $true }
    if ($sawWindowGone -and $wins.Count -gt 0 -and -not $reopenAt) {
        $reopenAt = $t; try { $reopenPid = $wins[0].Current.ProcessId } catch {}
    }
    if (-not $exPid -and -not $excelExitAt) { $excelExitAt = $t }
    Write-Output ("  t={0,2}s  win={1}  EXCEL=[{2}]  server=[{3}]" -f $t, $wins.Count, ($exPid -join ','), ($svPid -join ','))
    # stop early once fully settled: no window, no excel, no server
    if ($wins.Count -eq 0 -and -not $exPid -and -not $svPid) { Write-Output "  (fully settled)"; break }
}

Write-Output "--- final state ---"
$excelNow = (Get-Process EXCEL -EA SilentlyContinue).Id
$srvNow   = (Get-Process xll_showcase -EA SilentlyContinue).Id
$orphan   = ($srvNow -and -not $excelNow)
$ghost    = ($excelNow -and (Get-ExcelWindows).Count -eq 0)   # alive process, no window
Write-Output ("EXCEL now: [" + ($excelNow -join ',') + "]   server now: [" + ($srvNow -join ',') + "]")
Write-Output ("window gone at: " + $windowGoneAt + "s   reopened at: " + $reopenAt + "s (pid=$reopenPid, orig=$origPid)   EXCEL exited at: " + $excelExitAt + "s")
Write-Output ("LOCK _go.log     : " + (Probe-Lock $goLog))
Write-Output ("LOCK _native.log : " + (Probe-Lock $nativeLog))

Write-Output "=== VERDICT ==="
Write-Output ("  S1 window reopened          : " + $(if($reopenAt -ne $null){"YES at ${reopenAt}s (pid=$reopenPid)"}else{"no"}))
Write-Output ("  S1' ghost EXCEL (no window) : " + $(if($ghost){"YES (pid=$excelNow lingering windowless)"}else{"no"}))
Write-Output ("  S2 orphan server (no EXCEL) : " + $(if($orphan){"YES (server=$srvNow)"}else{"no"}))

# cleanup (only now)
Stop-ShowcaseProcesses
Write-Output "cleaned up"
