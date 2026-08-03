param(
    [int]$Rounds = 3,
    [int]$WatchSeconds = 120
)

# measure-command-rtd-connect.ps1 -- how long after the ribbon Build command do
# the RTD cells it wrote actually CONNECT?
#
# WHY THIS EXISTS. The backlog carried: "RTD cells created by the command path
# connect late or not at all -- 6 topics all connected at once about 74 s after
# the Build click, with a 67 s gap in the native log, and a different run in the
# same session showed zero connects inside the sampling window. Typing the same
# formula into a cell connects within a second." That is a claim about wall-clock
# behavior of a running Excel, so no unit test can settle it, and a number
# measured once in a transcript is a number nobody can re-measure.
#
# WHAT IT MEASURES. Per round: click Build via UIA, then poll the native log and
# record the timestamp of every "RTD Connect" line, reporting the delay from the
# click to the FIRST and LAST connect. A round where the add-in never loaded, or
# where Build was never found, is reported as VOID rather than as "0 connects" --
# those two are different findings and conflating them is how an environment
# problem gets logged as a product defect.
#
# NOT a pass/fail gate: it prints delays. Judging them needs the direct-entry
# baseline, which verify-rtd-stream-uia.ps1 already covers (sub-second).

. "$PSScriptRoot\uia-common.ps1"
$ErrorActionPreference = 'Continue'

$P = Resolve-ShowcasePaths
$xll = $P.Xll
$nativeLog = $P.NativeLog
$goLog = $P.GoLog

Write-Output "=== measure-command-rtd-connect : XLL=$xll ==="
Assert-ExcelTrustPreconditions -XllPath $xll -RequireRtd
Clear-ExcelResiliency
Initialize-Uia

$rows = @()
for ($i = 1; $i -le $Rounds; $i++) {
    # STOP FIRST, THEN DELETE, AND CHECK THE DELETE. Both logs, every round: the
    # Go log is what this script reads and it is append-only across runs, so a
    # surviving one credits this round with the previous round's connects.
    #
    # The order is load-bearing and the check is not paranoia -- both were
    # measured. Deleting before Stop-ShowcaseProcesses leaves the previous
    # round's Go server holding the file, Remove-Item fails, -SilentlyContinue
    # swallows it, and every round then reports the FIRST round's numbers. That
    # produced three rounds identical to 0.1 s, which is the only reason it was
    # noticed. A silent cleanup failure must be a hard stop, not a warning.
    Stop-ShowcaseProcesses
    Start-Sleep -Seconds 2
    foreach ($f in @($nativeLog, $goLog)) {
        if (Test-Path $f) { Remove-Item $f -Force -ErrorAction SilentlyContinue }
        if (Test-Path $f) {
            throw ("round {0}: could not delete {1} (a process still holds it). Refusing to measure against a stale log." -f $i, $f)
        }
    }

    # Launch the way every other driver does: a real workbook PLUS the XLL.
    # Passing only the .xll loads the add-in but opens no workbook, and
    # Wait-ShowcaseWindow requires a workbook window -- so the round times out
    # and reports VOID while the add-in is in fact loaded and fine. (Measured:
    # 3/3 VOID with the log written, which is what sent me to read this helper.)
    $tmp = New-ShowcaseWorkbook -Name "cmd_rtd_connect.xlsx"
    $launchedAt = Get-Date
    $proc = Start-ShowcaseExcel -Workbook $tmp -Xll $xll
    $win = Wait-ShowcaseWindow -LaunchedPid $proc.Id -SettleFirst 6 -Tries 38
    if (-not $win) {
        Write-Output ("round {0}: VOID -- no Excel window" -f $i)
        Write-XllLoadDiagnosis -XllPath $xll -NativeLog $nativeLog -Since $launchedAt
        Stop-ShowcaseProcesses
        continue
    }

    $clickedAt = Get-Date
    if (-not (Invoke-RibbonButton -Window $win -NamePattern 'Build')) {
        # No Build button means the add-in's ribbon never appeared. That is an
        # add-in-load question, not an RTD-timing one.
        Write-Output ("round {0}: VOID -- Build button not found (the ribbon tab IS the add-in)" -f $i)
        Write-XllLoadDiagnosis -XllPath $xll -NativeLog $nativeLog -Since $clickedAt
        Stop-ShowcaseProcesses
        continue
    }

    $first = $null; $last = $null; $count = 0
    $deadline = $clickedAt.AddSeconds($WatchSeconds)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 2
        if (-not (Test-Path $goLog)) { continue }
        # READ THE GO LOG, NOT THE NATIVE ONE. The native side logs the connect
        # with LogDebug ("RTD ConnectData: TopicID="), and a release build is
        # compiled with XLL_DEBUG=OFF, so that line CANNOT appear -- measured:
        # 0 [DEBUG] lines in a release native log. A probe for it reports
        # "0 connects" in every round and looks like a strong reproduction of
        # the very defect being investigated. The Go side logs the same event at
        # INFO ("RTD Connect request received"), which survives the release
        # build, and it is the other end of the same handshake.
        $hits = Select-String -Path $goLog -Pattern 'RTD Connect request received' -ErrorAction SilentlyContinue
        if ($hits) {
            $count = @($hits).Count
            $stamps = foreach ($h in $hits) {
                if ($h.Line -match 'time=(\S+?)\s') { [datetime]::Parse($Matches[1]) }
            }
            $stamps = @($stamps | Sort-Object)
            if ($stamps.Count -gt 0) { $first = $stamps[0]; $last = $stamps[-1] }
            # Six topics is the full showcase sheet; stop once they are all in.
            if ($count -ge 6) { break }
        }
    }

    $dFirst = if ($first) { [math]::Round(($first - $clickedAt).TotalSeconds, 1) } else { $null }
    $dLast = if ($last) { [math]::Round(($last - $clickedAt).TotalSeconds, 1) } else { $null }
    $rows += [pscustomobject]@{ Round = $i; Connects = $count; FirstSec = $dFirst; LastSec = $dLast }
    Write-Output ("round {0}: connects={1}  first={2}s  last={3}s  (watched {4}s)" -f `
        $i, $count, $(if ($null -ne $dFirst) { $dFirst } else { 'none' }), $(if ($null -ne $dLast) { $dLast } else { 'none' }), $WatchSeconds)

    Stop-ShowcaseProcesses
    Start-Sleep -Seconds 3
}

Write-Output '=== SUMMARY ==='
Write-Output ("valid rounds        : {0} of {1}" -f $rows.Count, $Rounds)
if ($rows.Count -gt 0) {
    Write-Output ("rounds with 0 connects : {0}" -f @($rows | Where-Object { $_.Connects -eq 0 }).Count)
    $withAny = @($rows | Where-Object { $null -ne $_.FirstSec })
    if ($withAny.Count -gt 0) {
        Write-Output ("first-connect delay : min {0}s  max {1}s" -f `
            (($withAny | Measure-Object FirstSec -Minimum).Minimum), (($withAny | Measure-Object FirstSec -Maximum).Maximum))
        Write-Output ("last-connect delay  : min {0}s  max {1}s" -f `
            (($withAny | Measure-Object LastSec -Minimum).Minimum), (($withAny | Measure-Object LastSec -Maximum).Maximum))
    }
}
Write-Output ("residual processes  : {0}" -f @(Get-Process EXCEL, xll_showcase -ErrorAction SilentlyContinue).Count)

# Clean up THIS SCRIPT'S OWN recovery litter and nothing else.
#
# Force-killing Excel each round makes it register the scratch workbook under
# Resiliency\DocumentRecovery, and those entries pop a recovery pane on the NEXT
# run that hijacks the window the driver waits for. Clear-ExcelResiliency
# deliberately refuses to touch DocumentRecovery because it can hold a user's
# unsaved work -- that stays true. What is removed here is only the subkeys whose
# payload names the workbook this script itself created, so a real document is
# never in scope. Measured: 10 entries accumulated over 7 rounds, all ours.
#
# The path is DERIVED from $ExcelSecurityKey rather than written out. A
# hand-written copy of it was silently wrong once already: a stray control
# character landed in the version segment, Test-Path went false, and this whole
# block was skipped WITHOUT A WORD while the entries piled up. Deriving it also
# tracks whichever Office version the rest of the harness already resolved.
$dr = Join-Path ($ExcelSecurityKey -replace '\\Security$', '\Resiliency') 'DocumentRecovery'
if (-not (Test-Path $dr)) {
    Write-Output 'recovery cleanup: no DocumentRecovery key, nothing to do'
} else {
    $removed = 0
    foreach ($sk in Get-ChildItem $dr) {
        $blob = ''
        foreach ($n in $sk.GetValueNames()) {
            $v = $sk.GetValue($n)
            $blob += if ($v -is [byte[]]) { [Text.Encoding]::Unicode.GetString($v) } else { [string]$v }
        }
        if (($blob -replace '[^ -~]', ' ') -match 'cmd_rtd_connect') {
            Remove-Item $sk.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            $removed++
        }
    }
    Write-Output ("recovery entries cleared (ours only): {0}" -f $removed)
}
