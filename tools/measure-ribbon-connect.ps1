param(
    [int]$Rounds = 20,
    [int]$SettleSeconds = 16,
    [switch]$DeleteAddinKeyEachRound
)

# measure-ribbon-connect.ps1 -- how often does the ribbon COM add-in FAIL to connect?
#
# WHY THIS EXISTS. The backlog carried "ribbon COM add-in connect fails sporadically
# (~3 of 20 cold starts)" for weeks, and v0.8.50 answered it only with Stage A: it
# started RECORDING a named fault class, three HRESULTs and the Resiliency state,
# and then the item sat waiting on "a field round of ~20 starts". That round is a
# property of many RUNNING Excel sessions, so no unit test can settle it -- and a
# rate measured once in a transcript is a rate nobody can re-measure. It belongs in
# the repo, like measure-idle-cpu.ps1.
#
# WHAT IT MEASURES. Per round: does the XLL load at all (native log created), does
# the ribbon connect, is the 60-attempt budget exhausted, and how many
# "connect step FAILED" lines appear with which dominant fault. Cold start every
# round -- Excel is fully killed and the native log deleted between rounds, because
# the reported failure is a STARTUP race and a warm Excel would not exercise it.
#
# WHY -DeleteAddinKeyEachRound EXISTS. v0.8.50 stopped deleting
# HKCU\...\Excel\Addins\<ProgID> at teardown, so the key -- and therefore our row in
# Excel's COMAddIns collection -- now survives into the NEXT startup. The leading
# hypothesis for why the failure stopped reproducing is exactly that: the collection
# used to be populated only after a mid-session registry write that Excel had to
# notice, which is where kProgIdNotInCollection would come from. This switch deletes
# the key before each round to put the old condition back. It is the A/B for that
# hypothesis, and it is OFF by default because it changes user-visible state.
#
# NOT a pass/fail gate. It prints a rate. Judging the rate needs a comparison
# baseline, and single-session numbers are the only ones worth comparing (the
# BENCHMARK/EXPERIMENTS discipline in shm applies here too).

. "$PSScriptRoot\uia-common.ps1"
$ErrorActionPreference = 'Continue'

$P = Resolve-ShowcasePaths
$xll = $P.Xll
$nativeLog = $P.NativeLog

Write-Output "=== measure-ribbon-connect : XLL=$xll ==="
# Environment before product: an unmet Trust Center lever produces "the XLL never
# loaded", which this script would otherwise report as a connect failure.
Assert-ExcelTrustPreconditions -XllPath $xll -RequireRtd
Clear-ExcelResiliency

$addinKey = 'HKCU:\Software\Microsoft\Office\Excel\Addins\XllShowcase.Ribbon'
$results = @()

for ($i = 1; $i -le $Rounds; $i++) {
    if (Test-Path $nativeLog) { Remove-Item $nativeLog -Force -ErrorAction SilentlyContinue }
    if ($DeleteAddinKeyEachRound -and (Test-Path $addinKey)) {
        Remove-Item $addinKey -Recurse -Force -ErrorAction SilentlyContinue
    }
    $started = Get-Date

    $p = Start-Process excel.exe -ArgumentList "`"$xll`"" -PassThru
    Start-Sleep -Seconds $SettleSeconds

    $txt = ''
    if (Test-Path $nativeLog) { $txt = Get-Content $nativeLog -Raw }

    $loaded    = [bool]($txt.Length -gt 0)
    $connected = [bool]($txt -match 'COM add-in connected')
    $failed60  = [bool]($txt -match 'connect failed after 60 attempts')
    $stepFails = ([regex]::Matches($txt, 'connect step FAILED')).Count
    $fault = ''
    if ($txt -match 'connect step FAILED[^\r\n]{0,3}([^\r\n]{0,60})') { $fault = $Matches[1].Trim() }

    # "The XLL never loaded" is NOT a connect failure, and conflating the two is how
    # an environment problem gets reported as a product defect. Say which it was.
    if (-not $loaded) {
        Write-Output ("run {0,3}: XLL DID NOT LOAD -- diagnosing before counting it as a connect failure" -f $i)
        Write-XllLoadDiagnosis -XllPath $xll -NativeLog $nativeLog -Since $started
    }

    $results += [pscustomobject]@{
        Round = $i; Loaded = $loaded; Connected = $connected
        Failed60 = $failed60; StepFails = $stepFails; Fault = $fault
    }
    Write-Output ("run {0,3}: loaded={1,-5} connected={2,-5} failed60={3,-5} stepFails={4} {5}" -f `
        $i, $loaded, $connected, $failed60, $stepFails, $fault)

    try { $p.CloseMainWindow() | Out-Null } catch {}
    Start-Sleep -Seconds 4
    # Two-tier: graceful close above, then force. A round that hangs must not stall
    # the remaining rounds -- the rate is the deliverable.
    Get-Process EXCEL, xll_showcase -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

$loadedN    = @($results | Where-Object Loaded).Count
$connectedN = @($results | Where-Object Connected).Count
$failed60N  = @($results | Where-Object Failed60).Count
$stepN      = @($results | Where-Object { $_.StepFails -gt 0 }).Count

Write-Output '=== SUMMARY ==='
Write-Output ("rounds            : {0}  (settle {1}s, addin key deleted each round: {2})" -f $Rounds, $SettleSeconds, [bool]$DeleteAddinKeyEachRound)
Write-Output ("xll loaded        : {0}" -f $loadedN)
Write-Output ("ribbon connected  : {0}" -f $connectedN)
Write-Output ("failed after 60   : {0}" -f $failed60N)
Write-Output ("any step FAILED   : {0}" -f $stepN)
$results | Where-Object { $_.StepFails -gt 0 } | Group-Object Fault | ForEach-Object {
    Write-Output ("  fault: '{0}' x{1}" -f $_.Name, $_.Count)
}
$residual = @(Get-Process EXCEL, xll_showcase -ErrorAction SilentlyContinue).Count
Write-Output ("residual processes: {0}" -f $residual)

if ($loadedN -lt $Rounds) {
    Write-Output 'NOTE: at least one round never loaded the XLL. Those rounds say nothing about the connect race -- read the diagnosis lines above before quoting a rate.'
}
