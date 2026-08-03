param(
    [int]$SampleSeconds = 20
)

# measure-idle-cpu.ps1 -- how much CPU the add-in burns while doing NOTHING.
#
# WHY THIS EXISTS. xll::WorkerLoop was a pure spin loop, so the XLL's worker thread
# consumed a full core for as long as the add-in was loaded. shm v0.8.21 +
# xll-gen v0.8.51 made it park. That is a claim about a RUNNING PROCESS, so no unit
# test can settle it -- it needs measuring, and the measurement belongs in the repo
# rather than in a transcript.
#
# WHAT IS MEASURED. Total processor time consumed while a workbook with the XLL
# loaded sits idle: no formulas, no RTD topics, no interaction. The XLL worker thread
# lives inside EXCEL.EXE, so Excel's own CPU time is the observable; the Go server is
# reported separately for context.
#
# HOW TO READ IT. Watch "% of one core". A spinning worker pins one core, so a
# pre-fix build reads ~100%. A parked worker should be a small fraction of that.
# Excel's own idle housekeeping is NOT zero, so do not expect 0% -- the comparison
# that means anything is this script against a build before and after the change.
#
# ENCODING: ASCII-only, same rule as uia-common.ps1.

. "$PSScriptRoot\uia-common.ps1"
$ErrorActionPreference = 'Continue'

$P = Resolve-ShowcasePaths
$xll = $P.Xll
if (-not (Test-Path $xll)) { Write-Output "MISSING XLL: $xll"; exit 2 }
Assert-ExcelTrustPreconditions -XllPath $xll -RequireRtd
Clear-ExcelResiliency

Stop-ShowcaseProcesses -IncludeGoServer
Start-Sleep 1

$app = New-Object -ComObject Excel.Application
$app.Visible = $true
$app.DisplayAlerts = $false
$pidExcel = (Get-Process EXCEL -EA SilentlyContinue | Sort-Object StartTime | Select-Object -Last 1).Id
Write-Output "Excel pid=$pidExcel. Loading XLL..."
$app.RegisterXLL($xll) | Out-Null
Start-Sleep 3
$wb = $app.Workbooks.Add()
Write-Output "Workbook open, add-in loaded, nothing else. Settling 5s before sampling..."
Start-Sleep 5

function Get-CpuSeconds([int]$procId) {
    $p = Get-Process -Id $procId -EA SilentlyContinue
    if (-not $p) { return $null }
    return $p.TotalProcessorTime.TotalSeconds
}

$srv = Get-Process xll_showcase -EA SilentlyContinue | Select-Object -First 1
$pidSrv = if ($srv) { $srv.Id } else { 0 }

$e0 = Get-CpuSeconds $pidExcel
$s0 = if ($pidSrv) { Get-CpuSeconds $pidSrv } else { 0 }
$t0 = Get-Date

Write-Output ("Sampling {0}s of IDLE time..." -f $SampleSeconds)
Start-Sleep $SampleSeconds

$e1 = Get-CpuSeconds $pidExcel
$s1 = if ($pidSrv) { Get-CpuSeconds $pidSrv } else { 0 }
$wall = ((Get-Date) - $t0).TotalSeconds

$eCpu = $e1 - $e0
$sCpu = if ($pidSrv) { $s1 - $s0 } else { 0 }

Write-Output ""
Write-Output ("=== IDLE CPU over {0:N1}s wall ===" -f $wall)
Write-Output ("  EXCEL.EXE (hosts the XLL worker) : {0,7:N2} core-seconds = {1,6:N1}% of one core" -f $eCpu, (100 * $eCpu / $wall))
if ($pidSrv) {
    Write-Output ("  xll_showcase.exe (Go server)     : {0,7:N2} core-seconds = {1,6:N1}% of one core" -f $sCpu, (100 * $sCpu / $wall))
} else {
    Write-Output "  xll_showcase.exe                 : not running"
}
Write-Output ""
Write-Output "~100%/core for Excel means the park is NOT in effect (a spinning worker pins one"
Write-Output "core). A parked worker is a small fraction; Excel's own idle work keeps it > 0."

Stop-ShowcaseCom -App $app -Workbook $wb -IncludeGoServer
Write-Output "cleaned up"
